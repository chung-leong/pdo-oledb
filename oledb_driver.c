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
#include "zend_exceptions.h"

#define CP_UTF16 1200

static HRESULT oledb_get_codepage(const char *name, int *pCodePage) {
	HRESULT hr = E_FAIL;

	if (name) {
		if (pIMultiLanguage) {
			if (_stricmp(name, "utf8") == 0 || _stricmp(name, "utf-8") == 0) {
				*pCodePage = CP_UTF8;
				hr = S_OK;
			} else if (_stricmp(name, "utf16") == 0 || stricmp(name, "utf-16") == 0) {
				*pCodePage = CP_UTF16;
				hr = S_OK;
			} else {
				MIMECSETINFO cp_info;
				LONG len_w = MultiByteToWideChar(CP_ACP, 0, name, -1, NULL, 0);
				BSTR name_w = SysAllocStringLen(NULL, len_w);
				MultiByteToWideChar(CP_ACP, 0, name, -1, name_w, len_w);
				hr = CALL(GetCharsetInfo, pIMultiLanguage, name_w, &cp_info);
				if (SUCCEEDED(hr)) {
					*pCodePage = cp_info.uiCodePage;
				} else {
					WCHAR buffer[256];
					if (wcslen(name_w) > 64) name_w[64] = '\0';
					wcscpy(buffer, L"MLang does not recognize ");
					wcscat(buffer, name_w);
					wcscat(buffer, L" as a valid encoding.");
					oledb_set_automation_error(buffer, L"58004");
				}
				SysFreeString(name_w);
			}
		}
	} else {
		*pCodePage = -1;
		hr = S_FALSE;
	}
	return hr;
}

static HRESULT oledb_create_charset_converter(pdo_oledb_conversion *conv, int type, int fromCodepage, int toCodepage) {
	HRESULT hr = E_FAIL;
	if (pIMultiLanguage) {
		SAFE_RELEASE(conv->pIMLangConvertCharsets[type]);
		if (fromCodepage != toCodepage && fromCodepage >= 0 && toCodepage >= 0) {
			hr = CALL(CreateConvertCharset, pIMultiLanguage, fromCodepage, toCodepage, 0, &conv->pIMLangConvertCharsets[type]);
			if (!SUCCEEDED(hr)) {
				oledb_set_automation_error(L"MLang cannot create converter for encoding.", L"58004");
			}
		} else {
			conv->pIMLangConvertCharsets[type] = NULL;
			hr = S_FALSE;
		}
	}
	return hr;
}

static HRESULT oledb_share_charset_converter(pdo_oledb_conversion *conv, int type, int other) {
	SAFE_RELEASE(conv->pIMLangConvertCharsets[type]);
	if (conv->pIMLangConvertCharsets[other]) {
		conv->pIMLangConvertCharsets[type] = conv->pIMLangConvertCharsets[other];
		ADDREF(conv->pIMLangConvertCharsets[type]);
	} else {
		conv->pIMLangConvertCharsets[type] = NULL;
	}
	return S_OK;
}

HRESULT oledb_create_bstr(pdo_oledb_conversion *conv, LPCSTR s, UINT len, BSTR *pWs, UINT *pLenW, int conversion_type) {
	HRESULT hr = E_UNEXPECTED;
	LPWSTR ws = NULL;
	UINT len_w = 0;
	IMLangConvertCharset *converter = conv->pIMLangConvertCharsets[conversion_type];

	if(s) {
		UINT converted = 0;

		if (len == -1) {
			len = strlen(s);
		}

		if (converter) {
			do {
				len_w += (len - converted);
				SysFreeString(ws);
				ws = SysAllocStringLen(NULL, len_w);
				converted = len;
				hr = CALL(DoConversionToUnicode, converter, (BYTE *) s, &converted, ws, &len_w);
			} while (SUCCEEDED(hr) && converted < len);
			ws[len_w] = '\0';
		} else {
			/* just copy the bytes */
			len_w = len / 2;
			ws = SysAllocStringLen(NULL, len_w);
			memcpy(ws, s, len_w * sizeof(WCHAR));
			ws[len_w] = '\0';
		}
	}
	if(pWs) {
		*pWs = ws;
	} else if(ws) {
		SysFreeString(ws);
	}
	if(pLenW) {
		*pLenW = len_w;
	}
	return hr;
}

HRESULT oledb_convert_bstr(pdo_oledb_conversion *conv, BSTR ws, UINT lenW, LPSTR *pS, UINT *pLen, int conversion_type) {
	HRESULT hr = E_UNEXPECTED;
	LPSTR s = NULL;
	UINT len = 0;
	IMLangConvertCharset *converter = conv->pIMLangConvertCharsets[conversion_type];

	if (ws) {
		if (lenW == -1) {
			lenW = wcslen(ws);
		}

		if (converter) {
			UINT converted = 0, buffer_size = 0, input, output;
			do {
				buffer_size += lenW;
				s = erealloc(s, buffer_size + 1);
				input = lenW - converted; 
				output = buffer_size - len;
				hr = CALL(DoConversionFromUnicode, converter, ws, &input, s, &output);
				if (SUCCEEDED(hr)) {
					converted += input;
					len += output;
				} else {
					break;
				}
			} while (SUCCEEDED(hr) && converted < lenW);
			s[len] = '\0';
		} else {
			/* just copy the bytes */
			len = lenW * sizeof(WCHAR);
			s = emalloc(len + sizeof(WCHAR));
			memcpy(s, ws, len + sizeof(WCHAR));
		}
	}
	if(pS) {
		*pS = s;
	} else if(s) {
		efree(s);
	}
	if (pLen) {
		*pLen = len;
	}
	return hr;
}

HRESULT oledb_convert_string(pdo_oledb_conversion *conv, LPCSTR src, UINT lenSrc, LPSTR *pDest, UINT *pLenDest, int conversion_type) {
	HRESULT hr = E_UNEXPECTED;
	LPSTR dest = NULL;
	UINT len_dest = 0;
	IMLangConvertCharset *converter = conv->pIMLangConvertCharsets[conversion_type];

	if (!converter) {
		*pDest = (LPSTR) src;
		*pLenDest = lenSrc;
		return S_FALSE;
	}

	if(src) {
		UINT converted_src = 0, total_converted_src = 0;
		UINT converted_dest = 0, total_converted_dest = 0;
		do {
			len_dest += (lenSrc - total_converted_src);
			dest = erealloc(dest, len_dest + 1);
			converted_src = lenSrc - total_converted_src;
			converted_dest = len_dest - total_converted_dest;
			hr = CALL(DoConversion, converter, (BYTE *) src + total_converted_src, &converted_src, (BYTE *) dest + total_converted_dest, &converted_dest);
			total_converted_src += converted_src;
			total_converted_dest += converted_dest;
		} while (hr == S_OK && total_converted_src < lenSrc);
		if (hr == S_FALSE) {
			hr = E_FAIL;
		}
		len_dest = total_converted_dest;
		dest[len_dest] = '\0';
	}
	if(pDest) {
		*pDest = dest;
	} else if(dest) {
		efree(dest);
	}
	if(pLenDest) {
		*pLenDest = len_dest;
	}
	return hr;
}

UINT oledb_get_proper_truncated_length(LPCSTR s, UINT len, const char *charset) 
{
	if (pIMultiLanguage) {
		int codepage;
		if (oledb_get_codepage(charset, &codepage) == S_OK) {
			DWORD context = 0;
			UINT len_w = 8000, len_a = len;
			WCHAR buffer_w[8000];

			/* see how many char can be converted */
			if (CALL(ConvertStringToUnicode, pIMultiLanguage, &context, codepage, (BYTE *) s, &len_a, buffer_w, &len_w) == S_OK) {
				len = len_a;
			}
		}
	}
	return len;
}

DWORD constant_to_internal_flag(int constant) 
{
	switch(constant) {
		case PDO_ATTR_FETCH_TABLE_NAMES: return ADD_TABLE_NAME;
		case PDO_ATTR_FETCH_CATALOG_NAMES: return ADD_CATALOG_NAME;
		case PDO_OLEDB_ATTR_UNICODE: return STRING_AS_UNICODE;
		case PDO_OLEDB_ATTR_GET_EXTENDED_METADATA: return UNIQUE_ROWS;
		case PDO_OLEDB_ATTR_CONVERT_DATATIME: return CONVERT_DATE_TIME;
		case PDO_OLEDB_ATTR_USE_INTEGRATED_AUTHENTICATION: return SECURE_CONNECTION;
		case PDO_OLEDB_ATTR_USE_CONNECTION_POOLING: return CONNECTION_POOLING;
		case PDO_OLEDB_ATTR_USE_ENCRYPTION: return ENCRYPTION;
		case PDO_OLEDB_ATTR_AUTOTRANSLATE: return AUTOTRANSLATE;
		case PDO_OLEDB_ATTR_TRUNCATE_STRING: return TRUNCATE_STRING;
	}
	return 0;
}

static int oledb_handle_fetch_error_func(pdo_dbh_t *dbh, pdo_stmt_t *stmt, zval *info TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	pdo_oledb_error_info *einfo = &H->einfo;

	if (stmt) {
		pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
		einfo = &S->einfo;
	} else {
		einfo = &H->einfo;
	}

	if (!SUCCEEDED(einfo->errcode)) {
		add_next_index_long(info, einfo->errcode);
		add_next_index_string(info, (einfo->errmsg) ? einfo->errmsg : "", 1);
	}

	return 1;
}

static int oledb_handle_closer(pdo_dbh_t *dbh TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	
	if (H) {
		SAFE_RELEASE(H->pITransactionLocal);
		SAFE_RELEASE(H->pIDBCreateCommand);
		SAFE_RELEASE(H->pIDBProperties);

		if (H->einfo.errmsg) {
			pefree(H->einfo.errmsg, dbh->is_persistent);
		}
		if (H->appname) {
			pefree(H->appname, dbh->is_persistent);
		}
		oledb_release_conversion_options(H->conv);
		pefree(H, dbh->is_persistent);
		dbh->driver_data = NULL;
	}
	return 0;
}

HRESULT oledb_stmt_set_driver_options(pdo_stmt_t *stmt, zval *driver_options TSRMLS_DC);

static int oledb_handle_preparer(pdo_dbh_t *dbh, const char *sql, long sql_len, pdo_stmt_t *stmt, zval *driver_options TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	pdo_oledb_stmt *S = ecalloc(1, sizeof(*S));
	int ret = 0;

	HRESULT hr;
	ICommandText *pICommandText = NULL;
	ICommandPrepare *pICommandPrepare = NULL;
	BSTR sql_w = NULL;
	char *nsql = NULL;
	int nsql_len = 0;

	S->H = H;
	S->flags = H->flags;
	oledb_copy_conversion_options(&S->conv, H->conv);

	stmt->driver_data = S;
	stmt->methods = &oledb_stmt_methods;

	hr = oledb_stmt_set_driver_options(stmt, driver_options TSRMLS_CC);
	if (!SUCCEEDED(hr)) goto cleanup;

	hr = CALL(CreateCommand, H->pIDBCreateCommand, NULL, &IID_ICommand, (IUnknown **) &S->pICommand);

	/* perform prepare only if provider supports placeholders */
	QUERY_INTERFACE(S->pICommand, IID_ICommandWithParameters, S->pICommandWithParameters);
	if (S->pICommandWithParameters) {
		QUERY_INTERFACE(S->pICommandWithParameters, IID_IAccessor, S->pIAccessorCommand);
		stmt->supports_placeholders = PDO_PLACEHOLDER_POSITIONAL;

		ret = _pdo_parse_params(stmt, (char*)sql, sql_len, &nsql, &nsql_len TSRMLS_CC);

		if (ret == 1) {
			/* query was rewritten */
			sql = nsql;
			sql_len = nsql_len;
		} else if (ret == -1) {
			/* failed to parse */
			strcpy(dbh->error_code, stmt->error_code);
			ret = 0;
			goto cleanup;
		}

		hr = QUERY_INTERFACE(S->pICommand, IID_ICommandText, pICommandText);
		if (!pICommandText) goto cleanup;

		oledb_create_bstr(S->conv, sql, -1, &sql_w, NULL, CONVERT_FROM_INPUT_TO_QUERY);

		hr = CALL(SetCommandText, pICommandText, &DBGUID_DEFAULT, sql_w);
		if (!SUCCEEDED(hr)) goto cleanup;

		/* don't do a round-trip to server to prepare statement for PDO::query() */
		if (!stmt->active_query_string) {
			hr = QUERY_INTERFACE(S->pICommand, IID_ICommandPrepare, pICommandPrepare);
			if (pICommandPrepare) {
				hr = CALL(Prepare, pICommandPrepare, 0);
				if (!SUCCEEDED(hr)) goto cleanup;

				if (S->pICommandWithParameters) {
					/* see if the provider can derive parameter information */
					hr = CALL(GetParameterInfo, S->pICommandWithParameters, &S->paramCount, &S->paramInfo, &S->paramNamesBuffer);
				}
			} else {
				hr = S_OK;
			}
		}
		ret = 1;
	} else {
		stmt->supports_placeholders = PDO_PLACEHOLDER_NONE;
		ret = 1;
	}

cleanup:
	SAFE_RELEASE(pICommandText);
	SAFE_RELEASE(pICommandPrepare);
	SAFE_EFREE(nsql);
	SysFreeString(sql_w);
	pdo_oledb_error(dbh, hr);
	return ret;
}

static long oledb_handle_doer(pdo_dbh_t *dbh, const char *sql, long sql_len TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	int ret = -1;

	HRESULT hr;
	ICommandText *pICommandText = NULL;
	ICommand *pICommand = NULL;
	DBPARAMS params = { 0, 0, 0 };
	DBCOUNTITEM rows_affected;
	BSTR sql_w = NULL;

	hr = CALL(CreateCommand, H->pIDBCreateCommand, NULL, &IID_ICommandText, (IUnknown **) &pICommandText);
	if (!pICommandText) goto cleanup;

	oledb_create_bstr(H->conv, sql, -1, &sql_w, NULL, CONVERT_FROM_INPUT_TO_QUERY);

	hr = CALL(SetCommandText, pICommandText, &DBGUID_DEFAULT, sql_w);
	if (!SUCCEEDED(hr)) goto cleanup;

	hr = QUERY_INTERFACE(pICommandText, IID_ICommand, pICommand);
	if (!pICommand) goto cleanup;

	/* execute command without opening a result set */
	hr = CALL(Execute, pICommand, NULL, &IID_NULL, &params, &rows_affected, NULL);
	if (!SUCCEEDED(hr)) goto cleanup;

	ret = 0;

cleanup:
	SAFE_RELEASE(pICommandText);
	SAFE_RELEASE(pICommand);
	SysFreeString(sql_w);
	pdo_oledb_error(dbh, hr);
	return ret;
}

static int oledb_handle_quoter(pdo_dbh_t *dbh, const char *unquoted, int unquotedlen, char **quoted, int *quotedlen, enum pdo_param_type param_type  TSRMLS_DC)
{
	return 0;
}

static int oledb_handle_begin(pdo_dbh_t *dbh TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	int ret = 0;

	HRESULT hr = S_FALSE;
	if (H->pITransactionLocal) {
		hr = CALL(StartTransaction, H->pITransactionLocal, ISOLATIONLEVEL_ISOLATED, 0, NULL, NULL);
		if (!SUCCEEDED(hr)) goto cleanup;

		ret = 1;
	}

cleanup:
	pdo_oledb_error(dbh, hr);
	return ret;
}

static int oledb_handle_commit(pdo_dbh_t *dbh TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	int ret = 0;

	HRESULT hr = S_FALSE;
	if (H->pITransactionLocal) {
		hr = CALL(Commit, H->pITransactionLocal, FALSE, XACTTC_SYNC, 0);
		if (!SUCCEEDED(hr)) goto cleanup;

		ret = 1;
	}

cleanup:
	pdo_oledb_error(dbh, hr);
	return ret;
}

static int oledb_handle_rollback(pdo_dbh_t *dbh TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	int ret = 0;

	HRESULT hr = S_FALSE;
	if (H->pITransactionLocal) {
		hr = CALL(Abort, H->pITransactionLocal, NULL, FALSE, FALSE);
		if (!SUCCEEDED(hr)) goto cleanup;

		ret = 1;
	}

cleanup:
	pdo_oledb_error(dbh, hr);
	return ret;
}

static char *oledb_handle_last_insert_id(pdo_dbh_t *dbh, const char *name, unsigned int *len TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	char *ret = NULL;

	HRESULT hr;
	ICommand *pICommand = NULL;
	ICommandText *pICommandText = NULL;
	IRowset *pIRowset = NULL;
	IAccessor *pIAccessor = NULL;
	DBCOUNTITEM rows_obtained;
	DBBINDING id_binding;
	DBBINDSTATUS id_bind_status;
	HACCESSOR hAccessor = 0;
	HROW hRow = 0, *hRows;

	struct {
		DWORD status;
		DBLENGTH length;
		char value[128];
	} buffer;

	hr = CALL(CreateCommand, H->pIDBCreateCommand, NULL, &IID_ICommandText, (IUnknown **) &pICommandText);
	if (!pICommandText) goto cleanup;

	hr = CALL(SetCommandText, pICommandText, &DBGUID_DEFAULT, L"SELECT @@IDENTITY");
	if (!SUCCEEDED(hr)) goto cleanup;

	hr = QUERY_INTERFACE(pICommandText, IID_ICommand, pICommand);
	if (!pICommand) goto cleanup;

	hr = CALL(Execute, pICommand, NULL, &IID_IRowset, NULL, NULL, (IUnknown **) &pIRowset);
	if (!pIRowset) goto cleanup;

	hr = QUERY_INTERFACE(pIRowset, IID_IAccessor, pIAccessor);
	if (!pIAccessor) goto cleanup;

	ZeroMemory(&id_binding, sizeof(id_binding));
	id_binding.iOrdinal = 1;
	id_binding.wType = DBTYPE_STR;
	id_binding.obStatus = 0;
	id_binding.obLength = id_binding.obStatus + sizeof(buffer.status);
	id_binding.obValue = id_binding.obLength + sizeof(buffer.length);
	id_binding.cbMaxLen = sizeof(buffer) - id_binding.obValue;
	id_binding.dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;

	/* create accessor to retrieve the column */
	hr = CALL(CreateAccessor, pIAccessor, DBACCESSOR_ROWDATA, 1, &id_binding, 0, &hAccessor, &id_bind_status);
	if (!SUCCEEDED(hr)) goto cleanup;

	hRows = &hRow;
	CALL(GetNextRows, pIRowset, DB_NULL_HCHAPTER, 0, 1, &rows_obtained, &hRows);
	if (!rows_obtained) goto cleanup;

    hr = CALL(GetData, pIRowset, hRow, hAccessor, &buffer);
	CALL(ReleaseRows, pIRowset, rows_obtained, hRows, NULL, NULL, NULL);

	if (!SUCCEEDED(hr)) goto cleanup;

	if (buffer.status == DBSTATUS_S_OK) {
		ret = estrndup(buffer.value, buffer.length);
		*len = buffer.length;
	}

cleanup:
	SAFE_RELEASE(pICommandText);
	SAFE_RELEASE(pICommand);
	SAFE_RELEASE(pIRowset);
	if (pIAccessor) {
		if (hAccessor) {
			CALL(ReleaseAccessor, pIAccessor, hAccessor, NULL);
		}
		RELEASE(pIAccessor);
	}
	pdo_oledb_error(dbh, hr);
	return ret;
}

HRESULT oledb_set_internal_flag(long attr, zval *val, DWORD mask, DWORD *pFlags)
{
	HRESULT hr = S_FALSE;
	DWORD flag = constant_to_internal_flag(attr);
	if (flag) {
		if (!(flag & mask)) {
			oledb_set_automation_error(L"Illegal operation", L"58004");
			hr = E_FAIL;
		} else {
			convert_to_boolean(val);
			if (Z_LVAL_P(val)) {
				*pFlags |= flag;
			} else {
				*pFlags &= ~flag;
			}
			hr = S_OK;
		}
	} else if (attr == PDO_ATTR_CURSOR) {
		convert_to_boolean(val);
		if (mask & (SERVER_SIDE_CURSOR | SERVER_SIDE_CURSOR)) {
			if (Z_LVAL_P(val) & PDO_OLEDB_CURSOR_SERVER_SIDE) {
				*pFlags |= SERVER_SIDE_CURSOR;
			} else {
				*pFlags &= ~SERVER_SIDE_CURSOR;
			}
			if ((Z_LVAL_P(val) & 0x0FFFFFFF) == PDO_CURSOR_SCROLL) {
				*pFlags |= SCROLLABLE_CURSOR;
			} else {
				*pFlags &= ~SCROLLABLE_CURSOR;
			}
			hr = S_OK;
		} else {
			oledb_set_automation_error(L"Illegal operation", L"58004");
			hr = E_FAIL;
		}
	}
	return hr;
}

HRESULT oledb_get_internal_flag(long attr, DWORD flags, zval *val) 
{
	DWORD flag = constant_to_internal_flag(attr);
	if (flag) {
		ZVAL_BOOL(val, (flags & flag) ? TRUE : FALSE);
		return 1;
	} else if (attr == PDO_ATTR_CURSOR) {
		int cursor_type = 0;
		if (flags & SERVER_SIDE_CURSOR) {
			cursor_type |= PDO_OLEDB_CURSOR_SERVER_SIDE;
		}
		if (flags & SCROLLABLE_CURSOR) {
			cursor_type |= PDO_CURSOR_SCROLL;
		}
		ZVAL_LONG(val, cursor_type);
		return 1;
	}
	return 0;
}

void oledb_create_conversion_options(pdo_oledb_conversion **pConv, int persistent)
{
	*pConv = pecalloc(1, sizeof(**pConv), persistent);
	(*pConv)->refcount = 1;
	(*pConv)->persistent = persistent;
}

void oledb_copy_conversion_options(pdo_oledb_conversion **pConv, pdo_oledb_conversion *src)
{
	*pConv = src;
	(*pConv)->refcount++;
}

void oledb_release_conversion_options(pdo_oledb_conversion *conv)
{
	conv->refcount--;
	if (conv->refcount == 0) {
		int i;
		SAFE_EFREE(conv->charset);
		SAFE_EFREE(conv->queryCharset);
		SAFE_EFREE(conv->varcharCharset);
		for(i = 0; i < CONVERTER_COUNT; i++) {
			SAFE_RELEASE(conv->pIMLangConvertCharsets[i]);
		}
		efree(conv);
	}
}

void oledb_split_conversion_options(pdo_oledb_conversion **pConv, int mainintainPersistence)
{
	if ((*pConv)->refcount > 1) {
		/* need to split it */
		pdo_oledb_conversion *org, *conv;
		int persistent = (*pConv)->persistent && mainintainPersistence;
		int i;
		org = *pConv;
		conv = pemalloc(sizeof(*conv), persistent);
		conv->refcount = 1;
		conv->persistent = persistent;
		if(org->charset) {
			conv->charset = pestrdup(org->charset, persistent);
		}
		if(org->queryCharset) {
			conv->queryCharset = pestrdup(org->queryCharset, persistent);
		}
		if(org->varcharCharset) {
			conv->varcharCharset = pestrdup(org->varcharCharset, persistent);
		}
		for(i = 0; i < CONVERTER_COUNT; i++) {
			conv->pIMLangConvertCharsets[i] = org->pIMLangConvertCharsets[i];
			SAFE_ADDREF(conv->pIMLangConvertCharsets[i]);
		}

		*pConv = conv;
	}
}

HRESULT oledb_set_conversion_option(pdo_oledb_conversion **pConv, long attr, zval *val, int maintainPersistence TSRMLS_DC)
{
	HRESULT hr = S_FALSE;
	pdo_oledb_conversion *conv;
	int codepage, queryCodepage, varcharCodepage;
	
	oledb_split_conversion_options(pConv, maintainPersistence);
	conv = *pConv;

	switch (attr) {
		case PDO_OLEDB_ATTR_ENCODING:
			if (conv->charset) {
				pefree(conv->charset, conv->persistent);
				conv->charset = NULL;
			}
			if (Z_TYPE_P(val) != IS_NULL) {
				convert_to_string(val);
				if (Z_STRLEN_P(val) > 0) {
					conv->charset = pestrdup(Z_STRVAL_P(val), conv->persistent);
				}
			} 
			hr = oledb_get_codepage(conv->charset, &codepage);
			if (!SUCCEEDED(hr)) goto cleanup;
			hr = oledb_get_codepage(conv->varcharCharset, &varcharCodepage);
			if (!SUCCEEDED(hr)) goto cleanup;

			/* converter for Unicode text */
			hr = oledb_create_charset_converter(conv, CONVERT_FROM_INPUT_TO_UNICODE, codepage, CP_UTF16);
			hr = oledb_create_charset_converter(conv, CONVERT_FROM_UNICODE_TO_OUTPUT, CP_UTF16 , codepage);

			/* converter for codepage text (if encoding is different from output/input)  */
			hr = oledb_create_charset_converter(conv, CONVERT_FROM_INPUT_TO_VARCHAR, codepage, varcharCodepage);
			hr = oledb_create_charset_converter(conv, CONVERT_FROM_VARCHAR_TO_OUTPUT, varcharCodepage , codepage);

			/* change how queries are decoded as well */
			if (!conv->queryCharset) {
				hr = oledb_share_charset_converter(conv, CONVERT_FROM_INPUT_TO_QUERY, CONVERT_FROM_INPUT_TO_UNICODE);
			}
			if (SUCCEEDED(hr)) {
				hr = S_OK;
			}
			break;
		case PDO_OLEDB_ATTR_QUERY_ENCODING:
			if (conv->queryCharset) {
				pefree(conv->queryCharset, conv->persistent);
				conv->queryCharset = NULL;
			}
			if (Z_TYPE_P(val) != IS_NULL) {
				convert_to_string(val);
				if (Z_STRLEN_P(val) > 0) {
					conv->queryCharset = pestrdup(Z_STRVAL_P(val), conv->persistent);
				}
			} 
			if (conv->queryCharset) {
				hr = oledb_get_codepage(conv->queryCharset, &queryCodepage);
				if (!SUCCEEDED(hr)) goto cleanup;
				hr = oledb_create_charset_converter(conv, CONVERT_FROM_INPUT_TO_QUERY, queryCodepage, CP_UTF16);
			} else {
				hr = oledb_share_charset_converter(conv, CONVERT_FROM_INPUT_TO_QUERY, CONVERT_FROM_INPUT_TO_UNICODE);
			}
			if (SUCCEEDED(hr)) {
				hr = S_OK;
			}
			break;
		case PDO_OLEDB_ATTR_CHAR_ENCODING:
			if (conv->varcharCharset) {
				pefree(conv->varcharCharset, conv->persistent);
				conv->varcharCharset = NULL;
			}
			if (Z_TYPE_P(val) != IS_NULL) {
				convert_to_string(val);
				if (Z_STRLEN_P(val) > 0) {
					conv->varcharCharset = pestrdup(Z_STRVAL_P(val), conv->persistent);
				}
			}
			hr = oledb_get_codepage(conv->charset, &codepage);
			if (!SUCCEEDED(hr)) goto cleanup;
			hr = oledb_get_codepage(conv->varcharCharset, &varcharCodepage);
			if (!SUCCEEDED(hr)) goto cleanup;

			hr = oledb_create_charset_converter(conv, CONVERT_FROM_INPUT_TO_VARCHAR, codepage, varcharCodepage);
			hr = oledb_create_charset_converter(conv, CONVERT_FROM_VARCHAR_TO_OUTPUT, varcharCodepage , codepage);
			if (SUCCEEDED(hr)) {
				hr = S_OK;
			}
			break;
	}
cleanup:
	return hr;
}

HRESULT oledb_get_conversion_option(pdo_oledb_conversion *conv, long attr, zval *val TSRMLS_DC)
{
	HRESULT hr = S_FALSE;
	switch (attr) {
		case PDO_OLEDB_ATTR_ENCODING:
			ZVAL_STRING(val, (conv->charset ? conv->charset : ""), TRUE);
			break;
		case PDO_OLEDB_ATTR_QUERY_ENCODING:
			if (conv->queryCharset) {
				ZVAL_STRING(val, conv->queryCharset, TRUE);
			} else {
				ZVAL_STRING(val, (conv->charset ? conv->charset : ""), TRUE);
			}
			break;
		case PDO_OLEDB_ATTR_CHAR_ENCODING:
			if (conv->varcharCharset) {
				ZVAL_STRING(val, conv->varcharCharset, TRUE);
			} else {
				ZVAL_STRING(val, (conv->charset ? conv->charset : ""), TRUE);
			}
			break;
	}
	return hr;
}

static HRESULT oledb_set_driver_option(pdo_dbh_t *dbh, long attr, zval *val TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	HRESULT hr;
	
	switch (attr) {
		case PDO_OLEDB_ATTR_APPLICATION_NAME:
			convert_to_string(val);
			if (H->appname) {
				pefree(H->appname, dbh->is_persistent);
			}
			H->appname = pestrdup(Z_STRVAL_P(val), dbh->is_persistent);
			hr = S_OK;
			break;
		default:
			hr = oledb_set_conversion_option(&H->conv, attr, val, TRUE TSRMLS_CC);
			if (hr == S_FALSE) {
				DWORD mask = SECURE_CONNECTION | CONNECTION_POOLING | ENCRYPTION | AUTOTRANSLATE | STRING_AS_UNICODE | STRING_AS_LOB | TRUNCATE_STRING | UNIQUE_ROWS | ADD_TABLE_NAME | ADD_CATALOG_NAME | CONVERT_DATE_TIME | SCROLLABLE_CURSOR | SERVER_SIDE_CURSOR;
				hr = oledb_set_internal_flag(attr, val, mask, &H->flags);
			}
	}
	return hr;
}

HRESULT oledb_set_driver_options(pdo_dbh_t *dbh, zval *driver_options TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	HRESULT hr = S_FALSE;

	if (driver_options) {
		HashTable *ht = Z_ARRVAL_P(driver_options);
		zval **v;
		char *str_key;
		int key;

		zend_hash_internal_pointer_reset(ht);
		while (SUCCESS == zend_hash_get_current_data(ht, (void**)&v)) {
			if (zend_hash_get_current_key(ht, &str_key, &key, FALSE) == HASH_KEY_IS_LONG) {
				SEPARATE_ZVAL_IF_NOT_REF(v);
				hr = oledb_set_driver_option(dbh, key, *v TSRMLS_CC);
				if (!SUCCEEDED(hr)) goto cleanup;
			}
			zend_hash_move_forward(ht);
		}
	}
	if (!driver_options || !zend_hash_index_exists(Z_ARRVAL_P(driver_options), PDO_OLEDB_ATTR_ENCODING)) {
		/* use default codepage if none specified */
		zval charset;
		ZVAL_STRING(&charset, "windows-1252", 1);
		hr = oledb_set_driver_option(dbh, PDO_OLEDB_ATTR_ENCODING, &charset TSRMLS_CC);
		zval_dtor(&charset);
	}

cleanup:
	return hr;
}

static int oledb_handle_set_attr(pdo_dbh_t *dbh, long attr, zval *val TSRMLS_DC)
{
	HRESULT hr = oledb_set_driver_option(dbh, attr, val TSRMLS_CC);
	pdo_oledb_error(dbh, hr);
	return SUCCEEDED(hr);
}

static int oledb_get_property(pdo_oledb_db_handle *H, const GUID *prop_set_id, DBPROPID prop_id, VARIANT *pVar) 
{
	HRESULT hr;
	DBPROPIDSET prop_id_set;  
	DBPROPSET *prop_set = NULL;
	ULONG prop_set_count;
	BOOL success = FALSE;

	prop_id_set.cPropertyIDs = 1;
	prop_id_set.guidPropertySet = *prop_set_id;
	prop_id_set.rgPropertyIDs = &prop_id;

	hr = CALL(GetProperties, H->pIDBProperties, 1, &prop_id_set, &prop_set_count, &prop_set);

	if (SUCCEEDED(hr) && SUCCEEDED(prop_set->rgProperties->dwStatus)) {
		if (pVar) {
			*pVar = prop_set->rgProperties->vValue;
		} else {
			VariantClear(&prop_set->rgProperties->vValue);
		}
		success = TRUE;
	}
	CoTaskMemFree(prop_set->rgProperties);
	CoTaskMemFree(prop_set);
	return success;
}

static int oledb_get_data_source_property(pdo_oledb_db_handle *H, DBPROPID prop_id, VARIANT *pVar) 
{
	return oledb_get_property(H, &DBPROPSET_DATASOURCEINFO, prop_id, pVar);
}

static char *oledb_get_version_info(pdo_oledb_db_handle *H, DBPROPID name_id, DBPROPID ver_id)
{
	VARIANT nameV, versionV;
	char *name = NULL, *version = NULL, *result = NULL;
	unsigned int name_len, version_len;

	VariantInit(&nameV);
	VariantInit(&versionV);

	oledb_get_data_source_property(H, name_id, &nameV);
	oledb_get_data_source_property(H, ver_id, &versionV);

	if (V_VT(&nameV) == VT_BSTR) {
		oledb_convert_bstr(H->conv, V_BSTR(&nameV), -1, &name, &name_len, CONVERT_FROM_UNICODE_TO_OUTPUT);
	}
	if (V_VT(&versionV) == VT_BSTR) {
		oledb_convert_bstr(H->conv, V_BSTR(&versionV), -1, &version, &version_len, CONVERT_FROM_UNICODE_TO_OUTPUT);
	}
	VariantClear(&nameV);
	VariantClear(&versionV);

	if (name && version) {
		result = emalloc(name_len + 1 + version_len + 1);
		memcpy(result, name, name_len);
		result[name_len] = ' ';
		memcpy(result + name_len + 1, version, version_len + 1);
		efree(name);
		efree(version);
	} else if (name) {
		result = name;
	} else if (version) {
		result = version;
	}
	return result;
}

static int oledb_handle_get_attr(pdo_dbh_t *dbh, long attr, zval *val TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	char *str;
	HRESULT hr;

	switch (attr) {
		case PDO_ATTR_SERVER_VERSION:	  
			if (str = oledb_get_version_info(H, DBPROP_DBMSNAME, DBPROP_DBMSVER)) {
				ZVAL_STRING(val, str, 0);
				hr = S_OK;
			} 
			break;
		case PDO_ATTR_CLIENT_VERSION:
			if (str = oledb_get_version_info(H, DBPROP_PROVIDERFRIENDLYNAME, DBPROP_PROVIDERVER)) {
				ZVAL_STRING(val, str, 0);
				hr = S_OK;
			} 
			break;
		case PDO_ATTR_TIMEOUT:
			ZVAL_LONG(val, H->timeout);
			hr = S_OK;
		case PDO_OLEDB_ATTR_APPLICATION_NAME:
			ZVAL_STRING(val, H->appname, TRUE);
			hr = S_OK;
		default:
			hr = oledb_get_conversion_option(H->conv, attr, val TSRMLS_CC);
			if (hr == S_FALSE) {
				hr = oledb_get_internal_flag(attr, H->flags, val);
			}
	}
	return SUCCEEDED(hr);
}

static struct pdo_dbh_methods oledb_methods = {
	oledb_handle_closer,
	oledb_handle_preparer,
	oledb_handle_doer,
	oledb_handle_quoter,
	oledb_handle_begin,
	oledb_handle_commit,
	oledb_handle_rollback,
	oledb_handle_set_attr,
	oledb_handle_last_insert_id,
	oledb_handle_fetch_error_func,
	oledb_handle_get_attr,
	NULL,	/* check_liveness */
};

static void oledb_add_prop_int(DBPROPSET *prop_set, DBPROPID prop_id, int n, int required)
{
	DBPROP *p = &prop_set->rgProperties[prop_set->cProperties++];
	p->dwOptions = required ? DBPROPOPTIONS_REQUIRED : DBPROPOPTIONS_OPTIONAL;
	VariantInit(&p->vValue);
	p->dwPropertyID = prop_id;
	V_VT(&p->vValue) = VT_I4;
	V_I4(&p->vValue) = n;
}


static void oledb_add_prop_bool(DBPROPSET *prop_set, DBPROPID prop_id, VARIANT_BOOL b, int required)
{
	DBPROP *p = &prop_set->rgProperties[prop_set->cProperties++];
	p->dwOptions = required ? DBPROPOPTIONS_REQUIRED : DBPROPOPTIONS_OPTIONAL;
	VariantInit(&p->vValue);
	p->dwPropertyID = prop_id;
	V_VT(&p->vValue) = VT_BOOL;
	V_BOOL(&p->vValue) = b;
}

static void oledb_add_prop_string(DBPROPSET *prop_set, DBPROPID prop_id, BSTR ws, int required)
{
	DBPROP *p = &prop_set->rgProperties[prop_set->cProperties++];
	p->dwOptions = required ? DBPROPOPTIONS_REQUIRED : DBPROPOPTIONS_OPTIONAL;
	VariantInit(&p->vValue);
	p->dwPropertyID = prop_id;
	V_VT(&p->vValue) = VT_BSTR;
	V_BSTR(&p->vValue) = ws;
}

static void oledb_convert_and_add_prop_string(pdo_oledb_conversion *conv, DBPROPSET *prop_set, DBPROPID prop_id, const char *s, int required)
{
	WCHAR *ws;
	oledb_create_bstr(conv, s, -1, &ws, NULL, CONVERT_FROM_INPUT_TO_UNICODE);
	oledb_add_prop_string(prop_set, prop_id, ws, required);
}

static void oledb_free_prop_strings(DBPROPSET *prop_set)
{
	unsigned int i = 0;
	for(i = 0; i < prop_set->cProperties; i++) {
		VariantClear(&prop_set->rgProperties[i].vValue);
	}
}

/* }}} */

static void oledb_check_provider_capability(pdo_dbh_t *dbh TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	VARIANT var;

	/* See if data source supports returning multiple result-sets */
	if (oledb_get_property(H, &DBPROPSET_DATASOURCEINFO, DBPROP_MULTIPLERESULTS, &var)) {
		if (V_I4(&var) == DBPROPVAL_MR_SUPPORTED) {
			H->flags = MULTIPLE_RESULTS;
		}
	}
}

static HRESULT oledb_set_initialization_properties(pdo_dbh_t *dbh, const char *host, const char *dbname TSRMLS_DC) /* {{{ */
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	int i, j;

	HRESULT hr = S_OK;
	DBPROP props0[16], props1[16];
	DBPROPSET prop_sets[2];
	int prop_set_count = 0;

	/* Set the initialization properties */
	prop_sets[prop_set_count].rgProperties = props0;
	prop_sets[prop_set_count].cProperties = 0;
	prop_sets[prop_set_count].guidPropertySet = DBPROPSET_DBINIT;
	if (host) {
		oledb_convert_and_add_prop_string(H->conv, &prop_sets[prop_set_count], DBPROP_INIT_DATASOURCE, host, 1);
	}
	if (dbname) {
		oledb_convert_and_add_prop_string(H->conv, &prop_sets[prop_set_count], DBPROP_INIT_CATALOG, dbname, 1);
	}
	if (H->timeout) {
		oledb_add_prop_int(&prop_sets[prop_set_count], DBPROP_INIT_TIMEOUT, H->timeout, 0);
	}
	if (H->flags & SECURE_CONNECTION) {
		if (host) {
			oledb_add_prop_string(&prop_sets[prop_set_count], DBPROP_AUTH_INTEGRATED, SysAllocString(L"SSPI"), 1);
		}
	} else {
		if (dbh->username) {
			oledb_convert_and_add_prop_string(H->conv, &prop_sets[prop_set_count], DBPROP_AUTH_USERID, dbh->username, 1);
		}
		if (dbh->password) {
			oledb_convert_and_add_prop_string(H->conv, &prop_sets[prop_set_count], DBPROP_AUTH_PASSWORD, dbh->password, 1);
			oledb_add_prop_bool(&prop_sets[prop_set_count], DBPROP_AUTH_ENCRYPT_PASSWORD, VARIANT_TRUE, 0);
			oledb_add_prop_bool(&prop_sets[prop_set_count], DBPROP_AUTH_CACHE_AUTHINFO, VARIANT_TRUE, 0);		 
			oledb_add_prop_bool(&prop_sets[prop_set_count], DBPROP_AUTH_PERSIST_SENSITIVE_AUTHINFO, VARIANT_TRUE, 0);
			oledb_add_prop_bool(&prop_sets[prop_set_count], DBPROP_AUTH_PERSIST_ENCRYPTED, VARIANT_TRUE, 0);
		}
	}
	if (H->flags & CONNECTION_POOLING) {
		DWORD services = DBPROPVAL_OS_RESOURCEPOOLING;
		oledb_add_prop_int(&prop_sets[prop_set_count], DBPROP_INIT_OLEDBSERVICES, services, 0);
	}
	if (prop_sets[prop_set_count].cProperties > 0) {
		prop_set_count++;
	}

	/* SQL Server specific */
	if (oledb_get_property(H, &DBPROPSET_SQLSERVERDBINIT, SSPROP_INIT_APPNAME, NULL)) {
		prop_sets[prop_set_count].rgProperties = props1;
		prop_sets[prop_set_count].cProperties = 0;
		prop_sets[prop_set_count].guidPropertySet = DBPROPSET_SQLSERVERDBINIT;
		if (H->appname) {
			oledb_convert_and_add_prop_string(H->conv, &prop_sets[prop_set_count], SSPROP_INIT_APPNAME, H->appname, 0);
		}
		if (H->flags & ENCRYPTION) {
			oledb_add_prop_bool(&prop_sets[prop_set_count], SSPROP_INIT_ENCRYPT, VARIANT_TRUE, 0);
		}
		if (!(H->flags & AUTOTRANSLATE)) {
			oledb_add_prop_bool(&prop_sets[prop_set_count], SSPROP_INIT_AUTOTRANSLATE, VARIANT_FALSE, 0);
		}
		if (prop_sets[prop_set_count].cProperties > 0) {
			prop_set_count++;
		}
	}

	if (prop_set_count > 0) {
		hr = CALL(SetProperties, H->pIDBProperties, prop_set_count, prop_sets);
		for (i = 0; i < prop_set_count; i++) {
			oledb_free_prop_strings(&prop_sets[i]);
		}
	}

	/* error out if a required property isn't set */
	if (hr != S_OK) {
		for (i = 0; i < prop_set_count; i++) {
			for (j = 0; j < (int) prop_sets[i].cProperties; j++) {
				if (prop_sets[i].rgProperties[j].dwStatus != DBPROPSTATUS_OK 
				&& 	prop_sets[i].rgProperties[j].dwOptions == DBPROPOPTIONS_REQUIRED) {
					goto cleanup;
				}
			}
		}
	}
	hr = S_OK;

cleanup:
	return hr;
}

static int oledb_handle_factory_mssql(pdo_dbh_t *dbh, zval *driver_options TSRMLS_DC) /* {{{ */
{
	pdo_oledb_db_handle *H;

	const char *charset = NULL;
	const char *dbname = NULL;
	const char *host = NULL;

	struct pdo_data_src_parser vars[] = {
		{ "host",		"localhost",	0 },
		{ "dbname",		NULL,			0 },
	};

	int i;

	HRESULT hr;
	IDBInitialize *pIDBInitialize = NULL;
	IDBCreateSession *pIDBCreateSession = NULL;

	H = pecalloc(1, sizeof(*H), dbh->is_persistent);
	dbh->driver_data = H;
	H->flags = CONVERT_DATE_TIME;
	H->timeout = 30;
	oledb_create_conversion_options(&H->conv, dbh->is_persistent);

	hr = oledb_set_driver_options(dbh, driver_options TSRMLS_CC);
	if (!SUCCEEDED(hr)) goto cleanup;

	_php_pdo_parse_data_source(dbh->data_source, dbh->data_source_len, vars, sizeof(vars) / sizeof(vars[0]));
	host = vars[0].optval;
	dbname = vars[1].optval;

	/* Create the data source object. */
	if (pIDataInitialize) {
		/* use MDAC */
		hr = CALL(CreateDBInstance, pIDataInitialize, &CLSID_SQLOLEDB, NULL, CLSCTX_INPROC_SERVER, NULL,
								&IID_IDBInitialize, (IUnknown **) &pIDBInitialize);
	} else {
		/* otherwise create the object directly */
		hr = CoCreateInstance(&CLSID_SQLOLEDB, NULL, CLSCTX_INPROC_SERVER,
								&IID_IDBInitialize, (void**) &pIDBInitialize);
	}
	if (!pIDBInitialize) goto cleanup;

	/* Get an IDBProperties pointer */
	hr = QUERY_INTERFACE(pIDBInitialize, IID_IDBProperties, H->pIDBProperties);
	if (!H->pIDBProperties) goto cleanup;

	/* Set initialization properties */
	oledb_set_initialization_properties(dbh, host, dbname TSRMLS_CC);

	/* Connect to the database */
	hr = CALL(Initialize, pIDBInitialize);
	if (!SUCCEEDED(hr)) goto cleanup;

	/* Create a session */
	hr = QUERY_INTERFACE(pIDBInitialize, IID_IDBCreateSession, pIDBCreateSession);
	if (!pIDBCreateSession) goto cleanup;

	hr = CALL(CreateSession, pIDBCreateSession, NULL, &IID_IDBCreateCommand, (IUnknown **) &H->pIDBCreateCommand);
	if (!H->pIDBCreateCommand) goto cleanup;

	oledb_check_provider_capability(dbh TSRMLS_CC);

	/* interface for handling transaction */
	QUERY_INTERFACE(H->pIDBCreateCommand, IID_ITransactionLocal, H->pITransactionLocal);

	dbh->alloc_own_columns = 1;
	dbh->max_escaped_char_length = 2;
	dbh->methods = &oledb_methods;

cleanup:
	for (i = 0; i < (sizeof(vars) / sizeof(vars[0])); i++) {
		if (vars[i].freeme) {
			efree(vars[i].optval);
		}
	}
	SAFE_RELEASE(pIDBInitialize);
	SAFE_RELEASE(pIDBCreateSession);
	pdo_oledb_error(dbh, hr);
	return SUCCEEDED(hr);
}
/* }}} */

static HRESULT oledb_merge_settings(pdo_dbh_t *dbh TSRMLS_DC) 
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;

	HRESULT hr;
	DBPROPID prop_ids0[] = { DBPROPVAL_OS_RESOURCEPOOLING, DBPROP_INIT_TIMEOUT, DBPROP_AUTH_INTEGRATED, DBPROP_AUTH_PASSWORD, DBPROP_AUTH_USERID };
	DBPROPID prop_ids1[] = { SSPROP_INIT_AUTOTRANSLATE, SSPROP_INIT_ENCRYPT, SSPROP_INIT_APPNAME };
	DBPROPIDSET prop_id_sets[2];  
	ULONG prop_set_count = 2, i, j;
	DBPROPSET *prop_sets = NULL;

	prop_id_sets[0].cPropertyIDs = sizeof(prop_ids0) /sizeof(prop_ids0[0]);
	prop_id_sets[0].guidPropertySet = DBPROPSET_DBINIT;
	prop_id_sets[0].rgPropertyIDs = prop_ids0;
	prop_id_sets[1].cPropertyIDs = sizeof(prop_ids0) /sizeof(prop_ids0[0]);
	prop_id_sets[1].guidPropertySet = DBPROPSET_SQLSERVERDBINIT;
	prop_id_sets[1].rgPropertyIDs = prop_ids1;

	hr = CALL(GetProperties, H->pIDBProperties, prop_set_count, prop_id_sets, &prop_set_count, &prop_sets);
	if (!SUCCEEDED(hr)) goto cleanup;

	for (i = 0; i < prop_set_count; i++) {
		for (j = 0; j < prop_sets[i].cProperties; j++) {
			DBPROPID id = prop_sets[i].rgProperties[j].dwPropertyID;
			DWORD status = prop_sets[i].rgProperties[j].dwStatus;
			VARIANT *value = &prop_sets[i].rgProperties[j].vValue;

			if (IsEqualIID(&prop_sets[i].guidPropertySet, &DBPROPSET_DBINIT)) {
				switch(id) {
					case DBPROPVAL_OS_RESOURCEPOOLING:
						if (V_VT(value) == VT_I4 && (V_I4(value) & DBPROPVAL_OS_RESOURCEPOOLING)) {
							H->flags |= CONNECTION_POOLING;
						}
						break;
					case DBPROP_INIT_TIMEOUT:
						if (V_VT(value) == VT_I4) {
							H->timeout = V_I4(value);
						}
						break;
					case DBPROP_AUTH_INTEGRATED:
						if (V_VT(value) == VT_BSTR && _wcsicmp(V_BSTR(value), L"SSPI") == 0) {
							H->flags |= SECURE_CONNECTION;
						}
						break;
				}
			} else if (IsEqualIID(&prop_sets[i].guidPropertySet, &DBPROPSET_SQLSERVERDBINIT)) {
				switch(id) {
					case SSPROP_INIT_AUTOTRANSLATE:
						if (V_VT(value) == VT_BOOL && V_BOOL(value) == VARIANT_TRUE) {
							H->flags |= AUTOTRANSLATE;
						}
						break;
					case SSPROP_INIT_ENCRYPT:
						if (V_VT(value) == VT_I4 && V_BOOL(value) == VARIANT_TRUE) {
							H->flags |= ENCRYPTION;
						}
						break;
					case SSPROP_INIT_APPNAME:
						if (V_VT(value) == VT_BSTR) {
							char *appname;
							oledb_convert_bstr(H->conv, V_BSTR(value), -1, &appname, NULL, CONVERT_FROM_UNICODE_TO_OUTPUT);
							if (H->appname) {
								pefree(H->appname, dbh->is_persistent);
							}
							if (dbh->is_persistent) {
								H->appname = pestrdup(appname, 1);
								efree(appname);
							} else {
								H->appname = appname;
							}
						}
						break;
				}
			}
		}
	}
	for (i = 0; i < prop_set_count; i++) {
		oledb_free_prop_strings(&prop_sets[i]);
	}

cleanup:
	CoTaskMemFree(prop_sets);
	return hr;
}

static int oledb_handle_factory_oledb(pdo_dbh_t *dbh, zval *driver_options TSRMLS_DC) /* {{{ */
{
	pdo_oledb_db_handle *H;

	HRESULT hr;
	IDBCreateSession *pIDBCreateSession = NULL;
	IDBInitialize *pIDBInitialize = NULL;

	const char *init_str = dbh->data_source;
	BSTR init_str_w = NULL;

	H = pecalloc(1, sizeof(*H), dbh->is_persistent);
	dbh->driver_data = H;
	H->flags = CONVERT_DATE_TIME;
	oledb_create_conversion_options(&H->conv, dbh->is_persistent);

	hr = oledb_set_driver_options(dbh, driver_options TSRMLS_CC);
	if (!SUCCEEDED(hr)) goto cleanup;

	oledb_create_bstr(H->conv, init_str, -1, &init_str_w, NULL, CONVERT_FROM_INPUT_TO_UNICODE);

	if (!pIDataInitialize) goto cleanup;

	/* Get data source through MDAC */
	hr = CALL(GetDataSource, pIDataInitialize, NULL, CLSCTX_INPROC_SERVER, init_str_w,
							  &IID_IDBInitialize, (IUnknown **) &pIDBInitialize);

	/* Get an IDBProperties pointer */
	hr = QUERY_INTERFACE(pIDBInitialize, IID_IDBProperties, H->pIDBProperties);
	if (!H->pIDBProperties) goto cleanup;

	/* Set initialization properties */
	hr = oledb_set_initialization_properties(dbh, NULL, NULL TSRMLS_CC);
	if (!SUCCEEDED(hr)) goto cleanup;

	/* merge settings from init-string with driver_options array */
	oledb_merge_settings(dbh TSRMLS_CC);

	/* Connect to the database */
	hr = CALL(Initialize, pIDBInitialize);
	if (!SUCCEEDED(hr)) goto cleanup;

	/* Create a session */
	hr = QUERY_INTERFACE(pIDBInitialize, IID_IDBCreateSession, pIDBCreateSession);
	if (!pIDBCreateSession) goto cleanup;

	hr = CALL(CreateSession, pIDBCreateSession, NULL, &IID_IDBCreateCommand, (IUnknown **) &H->pIDBCreateCommand);
	if (!H->pIDBCreateCommand) goto cleanup;

	oledb_check_provider_capability(dbh TSRMLS_CC);

	/* interface for handling transaction */
	QUERY_INTERFACE(H->pIDBCreateCommand, IID_ITransactionLocal, H->pITransactionLocal);

	dbh->alloc_own_columns = 1;
	dbh->max_escaped_char_length = 2;
	dbh->methods = &oledb_methods;

cleanup:
	SysFreeString(init_str_w);
	SAFE_RELEASE(pIDBInitialize);
	SAFE_RELEASE(pIDBCreateSession);
	pdo_oledb_error(dbh, hr);
	return SUCCEEDED(hr);
}
/* }}} */

pdo_driver_t pdo_mssql_driver = {
	PDO_DRIVER_HEADER(mssql),
	oledb_handle_factory_mssql
};

pdo_driver_t pdo_oledb_driver = {
	PDO_DRIVER_HEADER(oledb),
	oledb_handle_factory_oledb
};
