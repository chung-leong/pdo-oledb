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

/* $Id: pdo_oledb.c,v 1.14.2.7.2.2 2007/01/01 09:36:05 sebastian Exp $ */

#ifdef HAVE_CONFIG_H
#include "config.h"
#endif

#include "php.h"
#include "php_ini.h"
#include "ext/standard/info.h"
#include "pdo/php_pdo.h"
#include "pdo/php_pdo_driver.h"
#include "php_pdo_oledb.h"
#include "php_pdo_oledb_int.h"

/* {{{ pdo_oledb_functions[] */
const zend_function_entry pdo_oledb_functions[] = {
	{NULL, NULL, NULL}
};
/* }}} */

/* {{{ pdo_oledb_deps[] */
#if ZEND_MODULE_API_NO >= 20050922
static zend_module_dep pdo_oledb_deps[] = {
	ZEND_MOD_REQUIRED("pdo")
	{NULL, NULL, NULL}
};
#endif
/* }}} */

/* {{{ pdo_oledb_module_entry */
zend_module_entry pdo_oledb_module_entry = {
#if ZEND_MODULE_API_NO >= 20050922
	STANDARD_MODULE_HEADER_EX, NULL,
	pdo_oledb_deps,
#else
	STANDARD_MODULE_HEADER,
#endif
	"PDO_OLEDB",
	pdo_oledb_functions,
	PHP_MINIT(pdo_oledb),
	PHP_MSHUTDOWN(pdo_oledb),
	PHP_RINIT(pdo_oledb), 
	PHP_RSHUTDOWN(pdo_oledb), 
	PHP_MINFO(pdo_oledb),
	"1.0.1",
	STANDARD_MODULE_PROPERTIES
};
/* }}} */

ZEND_GET_MODULE(pdo_oledb)


int link_pdo(void) {
	HMODULE hLib = GetModuleHandle("php_pdo.dll");
	if(!hLib) {
		return FALSE;
	}

	_php_pdo_register_driver = (php_pdo_register_driver_proc) GetProcAddress(hLib, "php_pdo_register_driver");
	_php_pdo_unregister_driver = (php_pdo_unregister_driver_proc) GetProcAddress(hLib, "php_pdo_unregister_driver");
	_php_pdo_parse_data_source = (php_pdo_parse_data_source_proc) GetProcAddress(hLib, "php_pdo_parse_data_source");
	_php_pdo_get_dbh_ce = (php_pdo_get_dbh_ce_proc) GetProcAddress(hLib, "php_pdo_get_dbh_ce");
	_php_pdo_get_exception = (php_pdo_get_exception_proc) GetProcAddress(hLib, "php_pdo_get_exception");
	_pdo_parse_params = (pdo_parse_params_proc) GetProcAddress(hLib, "pdo_parse_params");
	_pdo_raise_impl_error = (pdo_raise_impl_error_proc) GetProcAddress(hLib, "pdo_raise_impl_error");
	_php_pdo_dbh_addref = (php_pdo_dbh_addref_proc) GetProcAddress(hLib, "php_pdo_dbh_addref");
	_php_pdo_dbh_delref = (php_pdo_dbh_delref_proc) GetProcAddress(hLib, "php_pdo_dbh_delref");
	_php_pdo_stmt_addref = (php_pdo_stmt_addref_proc) GetProcAddress(hLib, "php_pdo_stmt_addref");
	_php_pdo_stmt_delref = (php_pdo_stmt_delref_proc) GetProcAddress(hLib, "php_pdo_stmt_delref");

	return _php_pdo_register_driver 
		&& _php_pdo_unregister_driver
		&& _php_pdo_parse_data_source
		&& _php_pdo_get_dbh_ce
		&& _php_pdo_get_exception
		&& _pdo_parse_params
		&& _pdo_raise_impl_error
		&& _php_pdo_dbh_addref
		&& _php_pdo_dbh_delref
		&& _php_pdo_stmt_addref
		&& _php_pdo_stmt_delref;
}

/* {{{ PHP_MINIT_FUNCTION */
PHP_MINIT_FUNCTION(pdo_oledb)
{
	HRESULT hr;

	if(!link_pdo()) {
		return FAILURE;
	}

	if (FAILURE == _php_pdo_register_driver(&pdo_oledb_driver)) {
		return FAILURE;
	}

	if (FAILURE == _php_pdo_register_driver(&pdo_mssql_driver)) {
		return FAILURE;
	}

#undef  REGISTER_PDO_CLASS_CONST_LONG
#define REGISTER_PDO_CLASS_CONST_LONG(const_name, value) \
	zend_declare_class_constant_long(_php_pdo_get_dbh_ce(), const_name, sizeof(const_name)-1, (long)value TSRMLS_CC);

	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_USE_ENCRYPTION", (long)PDO_OLEDB_ATTR_USE_ENCRYPTION);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_USE_AUTO_TRANSLATION", (long)PDO_OLEDB_ATTR_AUTOTRANSLATE);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_USE_INTEGRATED_AUTHENTICATION", (long)PDO_OLEDB_ATTR_USE_INTEGRATED_AUTHENTICATION);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_USE_CONNECTION_POOLING", (long)PDO_OLEDB_ATTR_USE_CONNECTION_POOLING);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_GET_EXTENDED_METADATA", (long)PDO_OLEDB_ATTR_GET_EXTENDED_METADATA);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_CONVERT_DATATIME", (long)PDO_OLEDB_ATTR_CONVERT_DATATIME);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_UNICODE", (long)PDO_OLEDB_ATTR_UNICODE);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_APPLICATION_NAME", (long)PDO_OLEDB_ATTR_APPLICATION_NAME);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_ENCODING", (long)PDO_OLEDB_ATTR_ENCODING);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_QUERY_ENCODING", (long)PDO_OLEDB_ATTR_QUERY_ENCODING);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_CHAR_ENCODING", (long)PDO_OLEDB_ATTR_CHAR_ENCODING);
	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_ATTR_TRUNCATE_STRING", (long)PDO_OLEDB_ATTR_TRUNCATE_STRING);

	REGISTER_PDO_CLASS_CONST_LONG("OLEDB_CURSOR_SERVER_SIDE", (long)PDO_OLEDB_CURSOR_SERVER_SIDE);

	hr = CoInitialize(NULL);
	/* try to initialize MDAC */
	hr = CoCreateInstance(&CLSID_MSDAINITIALIZE, NULL, CLSCTX_INPROC_SERVER, &IID_IDataInitialize,(void**) &pIDataInitialize);

	/* start MLang */
	hr = CoCreateInstance(&CLSID_CMultiLanguage, NULL, CLSCTX_INPROC_SERVER, &IID_IMultiLanguage, (void **) &pIMultiLanguage);

	return SUCCEEDED(hr) ? SUCCESS : FAILURE;
}
/* }}} */

/* {{{ PHP_MSHUTDOWN_FUNCTION
 */
PHP_MSHUTDOWN_FUNCTION(pdo_oledb)
{
	if(!_php_pdo_unregister_driver) {
		return FAILURE;
	}

	/* Not sure why a deadlock occurs here sometimes */
	/*SAFE_RELEASE(pIDataInitialize);*/
	SAFE_RELEASE(pIMultiLanguage);
	CoUninitialize();

	_php_pdo_unregister_driver(&pdo_mssql_driver);
	_php_pdo_unregister_driver(&pdo_oledb_driver);
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_RINIT_FUNCTION */
PHP_RINIT_FUNCTION(pdo_oledb)
{
	CoInitialize(NULL);
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_MSHUTDOWN_FUNCTION
 */
PHP_RSHUTDOWN_FUNCTION(pdo_oledb)
{
	CoUninitialize();
	return SUCCESS;
}
/* }}} */

/* {{{ PHP_MINFO_FUNCTION
 */
PHP_MINFO_FUNCTION(pdo_oledb)
{
	HRESULT hr;
	ISourcesRowset *pISourcesRowset = NULL;
	IMultiLanguage *pIMultiLanguage = NULL;

	php_info_print_table_start();

	hr = CoCreateInstance(&CLSID_OLEDB_ENUMERATOR, NULL, CLSCTX_INPROC_SERVER, &IID_ISourcesRowset,(void**) &pISourcesRowset);
	if (pISourcesRowset) {
		IRowset *pIRowset = NULL;
		hr = CALL(GetSourcesRowset, pISourcesRowset, NULL, &IID_IRowset, 0, NULL, (IUnknown **) &pIRowset);
		if (pIRowset) {
			IAccessor *pIAccessor = NULL;
			hr = QUERY_INTERFACE(pIRowset, IID_IAccessor, pIAccessor);
			if (pIAccessor) {
				DBBINDING bindings[3];
				DBBINDSTATUS bind_statuses[3];
				HACCESSOR hAccessor = 0;

				struct string256 {
					DWORD status;
					DBLENGTH length;
					WCHAR value[256];
				};
				struct {
					struct string256 name;
					struct string256 description;
					unsigned int source_type;
				} buffer;

				ZeroMemory(bindings, sizeof(bindings));
				bindings[0].iOrdinal = 1;
				bindings[0].wType = DBTYPE_WSTR;
				bindings[0].obStatus = 0;
				bindings[0].obLength = bindings[0].obStatus + sizeof(buffer.name.status);
				bindings[0].obValue = bindings[0].obLength + sizeof(buffer.name.length);
				bindings[0].cbMaxLen = sizeof(buffer.name) - bindings[0].obValue;
				bindings[0].dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;

				bindings[1].iOrdinal = 3;
				bindings[1].wType = DBTYPE_WSTR;
				bindings[1].obStatus = sizeof(buffer.name);
				bindings[1].obLength = bindings[1].obStatus + sizeof(buffer.description.status);
				bindings[1].obValue = bindings[1].obLength + sizeof(buffer.description.length);
				bindings[1].cbMaxLen = sizeof(buffer.description) - bindings[1].obValue;
				bindings[1].dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;

				bindings[2].iOrdinal = 4;
				bindings[2].wType = DBTYPE_UI4;
				bindings[2].obValue = sizeof(buffer.name) + sizeof(buffer.description);
				bindings[2].dwPart = DBPART_VALUE;

				/* create accessor to retrieve the columns */
				hr = CALL(CreateAccessor, pIAccessor, DBACCESSOR_ROWDATA, 3, bindings, 0, &hAccessor, bind_statuses);

				if (hAccessor) {
					HROW hRow = 0, *hRows = &hRow;
					DBCOUNTITEM rows_obtained;

					php_info_print_table_header(2, "OLE-DB Providers" , "Description");
					while(CALL(GetNextRows, pIRowset, DB_NULL_HCHAPTER, 0, 1, &rows_obtained, &hRows), rows_obtained > 0) {
						hr = CALL(GetData, pIRowset, hRow, hAccessor, &buffer);
						CALL(ReleaseRows, pIRowset, rows_obtained, hRows, NULL, NULL, NULL);
						if (SUCCEEDED(hr)) {
							if (buffer.source_type == DBSOURCETYPE_DATASOURCE_TDP) {
								char name[267] = "";
								char description[267] = "";
								if (buffer.name.status == DBSTATUS_S_OK || buffer.name.status == DBSTATUS_S_TRUNCATED) {
									int len = WideCharToMultiByte(CP_ACP, 0, buffer.name.value, buffer.name.length, name, sizeof(name), NULL, NULL);
									name[len] = '\0';
								}
								if (buffer.description.status == DBSTATUS_S_OK || buffer.description.status == DBSTATUS_S_TRUNCATED) {
									int len = WideCharToMultiByte(CP_ACP, 0, buffer.description.value, buffer.description.length, description, sizeof(description), NULL, NULL);
									description[len] = '\0';
								}
								php_info_print_table_row(2, name, description);
							}
						}
					}
					CALL(ReleaseAccessor, pIAccessor, hAccessor, NULL);
				}

				RELEASE(pIAccessor);
			}
			RELEASE(pIRowset);
		}
		RELEASE(pISourcesRowset);
	}

	hr = CoCreateInstance(&CLSID_CMultiLanguage, NULL, CLSCTX_INPROC_SERVER, &IID_IMultiLanguage, (void **) &pIMultiLanguage);
	if (pIMultiLanguage) {
		IEnumCodePage *pEnumCodePage = NULL;
		hr = CALL(EnumCodePages, pIMultiLanguage, MIMECONTF_IMPORT | MIMECONTF_EXPORT,  &pEnumCodePage);
		if (pEnumCodePage) {
			MIMECPINFO info;
			ULONG fetched;

			php_info_print_table_header(2, "Character-sets" , "Description");
			while (hr = CALL(Next, pEnumCodePage, 1, &info, &fetched), hr == S_OK) {
				char charset[MAX_MIMECP_NAME];
				char description[MAX_MIMECP_NAME];
				WideCharToMultiByte(CP_ACP, 0, info.wszWebCharset, -1, charset, sizeof(charset), NULL, NULL);
				WideCharToMultiByte(CP_ACP, 0, info.wszDescription, -1, description, sizeof(description), NULL, NULL);
				php_info_print_table_row(2, charset, description);
			}
			RELEASE(pEnumCodePage);
		}
		RELEASE(pIMultiLanguage);
	}

	php_info_print_table_end();
}
/* }}} */

php_pdo_register_driver_proc _php_pdo_register_driver;
php_pdo_unregister_driver_proc _php_pdo_unregister_driver;
php_pdo_parse_data_source_proc _php_pdo_parse_data_source;
php_pdo_get_dbh_ce_proc _php_pdo_get_dbh_ce;
php_pdo_get_exception_proc _php_pdo_get_exception;
pdo_parse_params_proc _pdo_parse_params;
pdo_raise_impl_error_proc _pdo_raise_impl_error;
php_pdo_dbh_addref_proc _php_pdo_dbh_addref;
php_pdo_dbh_delref_proc _php_pdo_dbh_delref;
php_pdo_stmt_addref_proc _php_pdo_stmt_addref;
php_pdo_stmt_delref_proc _php_pdo_stmt_delref;

IDataInitialize *pIDataInitialize = NULL;
IMultiLanguage *pIMultiLanguage = NULL;