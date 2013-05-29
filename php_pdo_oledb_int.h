/*
  +----------------------------------------------------------------------+
  | PHP Version 5                                                        |
  +----------------------------------------------------------------------+
  | Copyright (c) 1997-2007 The PHP Group                                |
  +----------------------------------------------------------------------+
  | This source file is subject to version 3.0 of the PHP license,       |
  | that is bundled with this package in the file LICENSE, and is        |
  | available through the world-wide-web at the following url:           |
  | http://www.php.net/license/3_0.txt.                                  |
  | If you did not receive a copy of the PHP license and are unable to   |
  | obtain it through the world-wide-web, please send a note to          |
  | license@php.net so we can mail you a copy immediately.               |
  +----------------------------------------------------------------------+
  | Author: Chung Leong <chernyshevsky@hotmail.com>                      |
  +----------------------------------------------------------------------+
*/

#include <msdaguid.h>
#include <ole2.h>
#include <oledb.h>
#include <oledberr.h>
#include <sqloledb.h>
#include <msdasc.h>
#include <mlang.h>

#ifdef PHP_WIN32
# define PDO_OLEDB_TYPE	"Win32"
#endif

enum {
	CONVERT_FROM_INPUT_TO_UNICODE = 0,
	CONVERT_FROM_UNICODE_TO_OUTPUT,
	CONVERT_FROM_INPUT_TO_QUERY,
	CONVERT_FROM_INPUT_TO_VARCHAR,
	CONVERT_FROM_VARCHAR_TO_OUTPUT,
	CONVERTER_COUNT,
};

typedef struct {
	int persistent;
	int refcount;
	char *charset;
	char *queryCharset;
	char *varcharCharset;
	IMLangConvertCharset *pIMLangConvertCharsets[CONVERTER_COUNT];
} pdo_oledb_conversion;

typedef struct {
	const char *file;
	int line;

	HRESULT errcode;
	char *errmsg;
} pdo_oledb_error_info;

typedef struct {
	DWORD flags;
	long timeout;
	char *appname;

	IDBCreateCommand *pIDBCreateCommand;
	IDBProperties *pIDBProperties;
	ITransactionLocal *pITransactionLocal;
	IMultiLanguage *pIMultiLanguage;
	
	pdo_oledb_conversion *conv;

	pdo_oledb_error_info einfo;
} pdo_oledb_db_handle;

typedef ULONG DBLENGTH;

typedef struct {
	char *columnName;
	int columnNameLen;
	char *tableName;
	int tableNameLen;
	char *catalogName;
	int catalogNameLen;
	zval *defaultValue;

	BOOL isDynamic;
	BOOL isUnique;
	BOOL isKey;
	BOOL isAutoIncrement;
} pdo_oledb_column_meta_data;

typedef struct {
	DWORD flags;
	char *name;
	int nameLen;
	int precision;
	unsigned long maxLen;
	DBCOLUMNFLAGS columnFlags;
	DBORDINAL ordinal;
	DBTYPE columnType;
	DBTYPE retrievalType;
	int pdoType;
	int columnLength;
	DBLENGTH byteOffset;
	unsigned int byteCount;
	HACCESSOR hAccessorColumn;
	pdo_oledb_column_meta_data *metadata;
	pdo_oledb_conversion *conv;
} pdo_oledb_column;

typedef struct {
	DWORD flags;
	pdo_oledb_db_handle *H;
	pdo_oledb_column *columns;

	ICommand *pICommand;
	ICommandWithParameters *pICommandWithParameters;
	IAccessor *pIAccessorCommand;
	HACCESSOR hAccessorCommand;

	IMultipleResults *pIMultipleResults;
	IRowset *pIRowset;
	IAccessor *pIAccessorRowset;
	HACCESSOR hAccessorRowset;
	DBCOUNTITEM rowsAffected;
	DBCOUNTITEM rowIndex;

	HROW hRow;

	DB_UPARAMS paramCount;
	DBPARAMINFO	*paramInfo;
	OLECHAR *paramNamesBuffer;
	DBORDINAL paramOrdinal;

	void *inputBuffer;
	DBBYTEOFFSET nextInputOffset;

	void *outputBuffer;
	DBBYTEOFFSET nextOutputOffset;

	pdo_oledb_conversion *conv;
	pdo_oledb_error_info einfo;
} pdo_oledb_stmt;

typedef struct {
	DWORD flags;

	LPCWSTR dataType;
	unsigned int dataTypeWidth;
	DBORDINAL ordinal;
	DBLENGTH dataLength;
	DBTYPE retrievalType;
	DWORD ioFlags;
	unsigned int byteCount;
	void *dataPointer;
	DBLENGTH byteOffset;

	int intValue;
	double doubleValue;
	char *varcharValue;
	BSTR unicodeValue;
    IUnknown *stream;

	pdo_oledb_conversion *conv;
} pdo_oledb_param;
	
extern pdo_driver_t pdo_oledb_driver;
extern pdo_driver_t pdo_mssql_driver;

extern struct pdo_stmt_methods oledb_stmt_methods;

void pdo_oledb_init_error_table(void);
void pdo_oledb_fini_error_table(void);

enum {
	PDO_OLEDB_ATTR_USE_ENCRYPTION = PDO_ATTR_DRIVER_SPECIFIC,
	PDO_OLEDB_ATTR_USE_INTEGRATED_AUTHENTICATION,
	PDO_OLEDB_ATTR_USE_CONNECTION_POOLING,
	PDO_OLEDB_ATTR_GET_EXTENDED_METADATA,
	PDO_OLEDB_ATTR_CONVERT_DATATIME,
	PDO_OLEDB_ATTR_APPLICATION_NAME,
	PDO_OLEDB_ATTR_UNICODE,	
	PDO_OLEDB_ATTR_ENCODING,
	PDO_OLEDB_ATTR_QUERY_ENCODING,
	PDO_OLEDB_ATTR_CHAR_ENCODING,
	PDO_OLEDB_ATTR_AUTOTRANSLATE,
	PDO_OLEDB_ATTR_TRUNCATE_STRING,
};

#define PDO_OLEDB_CURSOR_SERVER_SIDE	0x80000000

typedef PDO_API int (*php_pdo_register_driver_proc)(pdo_driver_t *driver);
typedef PDO_API void (*php_pdo_unregister_driver_proc)(pdo_driver_t *driver);
typedef PDO_API int (*php_pdo_parse_data_source_proc)(const char *data_source, unsigned long data_source_len, struct pdo_data_src_parser *parsed, int nparams);
typedef PDO_API zend_class_entry * (*php_pdo_get_dbh_ce_proc)(void);
typedef PDO_API zend_class_entry * (*php_pdo_get_exception_proc)(void);
typedef PDO_API int (*pdo_parse_params_proc)(pdo_stmt_t *stmt, char *inquery, int inquery_len, char **outquery, int *outquery_len TSRMLS_DC);
typedef PDO_API void (*pdo_raise_impl_error_proc)(pdo_dbh_t *dbh, pdo_stmt_t *stmt, const char *sqlstate, const char *supp TSRMLS_DC);
typedef PDO_API void (*php_pdo_dbh_addref_proc)(pdo_dbh_t *dbh TSRMLS_DC);
typedef PDO_API void (*php_pdo_dbh_delref_proc)(pdo_dbh_t *dbh TSRMLS_DC);
typedef PDO_API void (*php_pdo_stmt_addref_proc)(pdo_stmt_t *stmt TSRMLS_DC);
typedef PDO_API void (*php_pdo_stmt_delref_proc)(pdo_stmt_t *stmt TSRMLS_DC);

extern php_pdo_register_driver_proc _php_pdo_register_driver;
extern php_pdo_unregister_driver_proc _php_pdo_unregister_driver;
extern php_pdo_parse_data_source_proc _php_pdo_parse_data_source;
extern php_pdo_get_dbh_ce_proc _php_pdo_get_dbh_ce;
extern php_pdo_get_exception_proc _php_pdo_get_exception;
extern pdo_parse_params_proc _pdo_parse_params;
extern pdo_raise_impl_error_proc _pdo_raise_impl_error;
extern php_pdo_dbh_addref_proc _php_pdo_dbh_addref;
extern php_pdo_dbh_delref_proc _php_pdo_dbh_delref;
extern php_pdo_stmt_addref_proc _php_pdo_stmt_addref;
extern php_pdo_stmt_delref_proc _php_pdo_stmt_delref;

extern IDataInitialize *pIDataInitialize;
extern IMultiLanguage *pIMultiLanguage;

#define SAFE_STRING(s) ((s)?(s):"")

/* VC6 doesn't support variadic macro :-( */

#define CALL(m, p, ...)					p->lpVtbl->m(p, __VA_ARGS__)
#define ADDREF(p)						CALL(AddRef, (p))
#define RELEASE(p)						CALL(Release, (p))
#define SAFE_ADDREF(p)					((p) ? ADDREF(p) : 0)
#define SAFE_RELEASE(p)					((p) ? RELEASE(p) : 0)
#define QUERY_INTERFACE(p, i, q)		CALL(QueryInterface, (p), &i, (void **) &q)

#define SAFE_EFREE(p)				if(p) efree(p);

void oledb_create_conversion_options(pdo_oledb_conversion **pConv, int persistent);
void oledb_copy_conversion_options(pdo_oledb_conversion **pConv, pdo_oledb_conversion *src);
void oledb_release_conversion_options(pdo_oledb_conversion *conv);
void oledb_split_conversion_options(pdo_oledb_conversion **pConv, int mainintainPersistence);
HRESULT oledb_set_conversion_option(pdo_oledb_conversion **pConv, long attr, zval *val, int maintainPersistence TSRMLS_DC);
HRESULT oledb_get_conversion_option(pdo_oledb_conversion *conv, long attr, zval *val TSRMLS_DC);

HRESULT oledb_create_bstr(pdo_oledb_conversion *conv, LPCSTR s, UINT len, BSTR *pWs, UINT *pLenW, int conversion_type);
HRESULT oledb_convert_bstr(pdo_oledb_conversion *conv, BSTR ws, UINT lenW, LPSTR *pS, UINT *pLen, int conversion_type);
HRESULT oledb_convert_string(pdo_oledb_conversion *conv, LPCSTR src, UINT lenSrc, LPSTR *pDest, UINT *pLenDest, int conversion_type);

UINT oledb_get_proper_truncated_length(LPCSTR s, UINT len, const char *charset);

HRESULT oledb_create_lob_stream(pdo_oledb_conversion *conv, IUnknown *pUnk, DBLENGTH length, int conversion, pdo_stmt_t *stmt, php_stream **pStream TSRMLS_DC);
HRESULT oledb_create_zval_stream(pdo_oledb_conversion *conv, zval *value, int unicode, IUnknown **pUnk, DBLENGTH *pLength TSRMLS_DC);

extern void _pdo_oledb_error(pdo_dbh_t *dbh, pdo_stmt_t *stmt, HRESULT result, const char *file, int line TSRMLS_DC);
#define pdo_oledb_error(h, hr) _pdo_oledb_error(h, NULL, hr, __FILE__, __LINE__ TSRMLS_CC)
#define pdo_oledb_error_stmt(s, hr) _pdo_oledb_error(s->dbh, s, hr, __FILE__, __LINE__ TSRMLS_CC)

void oledb_set_automation_error(LPCWSTR msg, LPCWSTR sqlcode);
HRESULT oledb_set_internal_flag(long attr, zval *val, DWORD mask, DWORD *pFlags);
HRESULT oledb_get_internal_flag(long attr, DWORD flags, zval *val);

#define STRING_AS_UNICODE	(1 << 0)
#define STRING_AS_LOB		(1 << 1)
#define STRING_AS_BYTES		(1 << 2)
#define VARIABLE_LENGTH		(1 << 3)
#define ALIESED_COLUMN		(1 << 4)
#define TRUNCATE_STRING		(1 << 4)

#define MULTIPLE_RESULTS	(1 << 7)

#define SECURE_CONNECTION	(1 << 8)
#define CONNECTION_POOLING	(1 << 9)
#define ENCRYPTION			(1 << 10)
#define AUTOTRANSLATE		(1 << 11)

#define UNIQUE_ROWS			(1 << 16)
#define ADD_TABLE_NAME		(1 << 17)
#define ADD_CATALOG_NAME	(1 << 18)
#define CONVERT_DATE_TIME	(1 << 19)
#define SCROLLABLE_CURSOR	(1 << 20)
#define SERVER_SIDE_CURSOR	(1 << 21)

