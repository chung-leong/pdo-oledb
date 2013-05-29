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
  | Author: Chung Leong <cleong@cal.berkeley.edu>                        |
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

const char *oledb_hresult_text(HRESULT hr);

void _pdo_oledb_error(pdo_dbh_t *dbh, pdo_stmt_t *stmt, HRESULT result, const char *file, int line TSRMLS_DC)
{
	pdo_oledb_db_handle *H = (pdo_oledb_db_handle *)dbh->driver_data;
	pdo_error_type *pdo_err;
	pdo_oledb_error_info *einfo;
	pdo_oledb_stmt *S;
	int persistent = 0;

	if (!H) return;

	if (stmt) {
		S = (pdo_oledb_stmt*)stmt->driver_data;
		pdo_err = &stmt->error_code;
		einfo   = &S->einfo;
	} else {
		pdo_err = &dbh->error_code;
		einfo   = &H->einfo;
		persistent = dbh->is_persistent;
	}

	einfo->file = file;
	einfo->line = line;
	einfo->errcode = result;

	if (einfo->errmsg) {
		pefree(einfo->errmsg, persistent);
		einfo->errmsg = NULL;
	}

	if (!SUCCEEDED(result)) {
		HRESULT hr;
		IErrorInfo *pIErrorInfo = NULL;
		IErrorRecords *pIErrorRecords = NULL;
		ISQLErrorInfo *pISQLErrorInfo = NULL;

		/* get the error object */
		hr = GetErrorInfo(0, &pIErrorInfo);
		if (pIErrorInfo) {
			BSTR descr_w = NULL;
			BSTR sql_state_w = NULL;
			LONG native_error = 0;
			char *descr = NULL;

			/* use IErrorRecords if possible */
			hr = QUERY_INTERFACE(pIErrorInfo, IID_IErrorRecords, pIErrorRecords);
			if (pIErrorRecords) {
				BSTR *record_descrs;
				LONG record_count = 0, i;
				
				hr = CALL(GetRecordCount, pIErrorRecords, &record_count);
				record_descrs = (BSTR *) ecalloc(record_count, sizeof(*record_descrs));
				for (i = 0; i < record_count; i++) {
					IErrorInfo *pIErrorRecordInfo = NULL;
					hr = CALL(GetErrorInfo, pIErrorRecords, i, 0x0409, &pIErrorRecordInfo);
					if (pIErrorRecordInfo) {
						hr = CALL(GetDescription, pIErrorRecordInfo, &record_descrs[i]);
						RELEASE(pIErrorRecordInfo);
					}
				}

				if (record_descrs) {
					/* join the descriptions together */
					int total_len = 0, offset = 0;
					for (i = 0; i < record_count; i++) {
						if (record_descrs[i]) {
							total_len += wcslen(record_descrs[i]) + 1;
						}
					}
					descr_w = SysAllocStringLen(NULL, total_len);
					for (i = 0; i < record_count; i++) {
						/* don't append description if it's the same as the previous one */
						if (record_descrs[i] && (i == 0 || wcscmp(record_descrs[i], record_descrs[i - 1]) != 0)) {
							if(offset != 0) {
								descr_w[offset - 1] = ' ';
							}
							wcscpy(descr_w + offset, record_descrs[i]);
							offset += wcslen(record_descrs[i]) + 1;
							SysFreeString(record_descrs[i]);
						}
					}
					efree(record_descrs);
				}

				for (i = 0; i < record_count; i++) {
					/* get a ISQLErrorInfo interface to retrieve the SQL state */
					hr = CALL(GetCustomErrorObject, pIErrorRecords, i, &IID_ISQLErrorInfo, (IUnknown **) &pISQLErrorInfo);
					if (pISQLErrorInfo) {
						hr = CALL(GetSQLInfo, pISQLErrorInfo, &sql_state_w, &native_error);
						if (SUCCEEDED(hr)) {
							/* use the first one obtained */
							break;
						}
					}
				}
			} else {
				hr = CALL(GetDescription, pIErrorInfo, &descr_w);
			}

			if (sql_state_w) {
				char *sql_state;
				int sql_state_len;
				oledb_convert_bstr(H->conv, sql_state_w, -1, &sql_state, &sql_state_len, CONVERT_FROM_UNICODE_TO_OUTPUT);
				if(sql_state_len <= 5) {
					strcpy(*pdo_err, sql_state);
				}
				efree(sql_state);
			} else {
				/* don't know what happened */
				strcpy(*pdo_err, "58004");
			}
			if (descr_w) {
				oledb_convert_bstr(H->conv, descr_w, -1, &descr, NULL, CONVERT_FROM_UNICODE_TO_OUTPUT);
				if (persistent) {
					einfo->errmsg = pestrdup(descr, persistent);
					efree(descr);
				} else {
					einfo->errmsg = descr;
				}
			}

			SysFreeString(descr_w);
			SysFreeString(sql_state_w);

			SAFE_RELEASE(pISQLErrorInfo);
			SAFE_RELEASE(pIErrorInfo);
			SAFE_RELEASE(pIErrorRecords);
		} else {
			/* no IErrorInfo, see if it's a pre-defined OLE-DB error  */
			const char *msg = oledb_hresult_text(result);
			if (msg) {
				einfo->errmsg = pestrdup(msg, persistent);
			} else {
				/* maybe it's a generic Windows error */
				LPVOID lpMsgBuf;
				FormatMessage( 
					FORMAT_MESSAGE_ALLOCATE_BUFFER | 
					FORMAT_MESSAGE_FROM_SYSTEM | 
					FORMAT_MESSAGE_IGNORE_INSERTS,
					NULL,
					result,
					MAKELANGID(LANG_ENGLISH, SUBLANG_DEFAULT),
					(LPTSTR) &lpMsgBuf,
					0,
					NULL 
				);
				if (lpMsgBuf) {
					einfo->errmsg = pestrdup((char *) lpMsgBuf, persistent);
					LocalFree(lpMsgBuf);
				}
			}
		}

		if (!dbh->methods) {
			zend_throw_exception_ex(_php_pdo_get_exception(), 0 TSRMLS_CC, "SQLSTATE[%s] [%d] %s",
					*pdo_err, einfo->errcode, einfo->errmsg);
		}
	} else { /* no error */
		strcpy(*pdo_err, PDO_ERR_NONE);
	}
}

typedef struct {
	IErrorInfoVtbl *lpIErrorInfoVtbl;
	IErrorRecordsVtbl *lpIErrorRecordsVtbl;
	ISQLErrorInfoVtbl *lpISQLErrorInfoVtbl;
	BSTR description;
	BSTR sqlcode;
	ULONG refcount;
} oledb_error_info;

#define DEFINE_THIS(vtbl, offset)		oledb_error_info *this = ((oledb_error_info *) &((IUnknown *) vtbl)[- offset])

static HRESULT STDMETHODCALLTYPE oledb_error_info_QueryInterface(oledb_error_info *this,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject)
{
	if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_IErrorInfo)) {
		*ppvObject = &this->lpIErrorInfoVtbl;
	} else if (IsEqualIID(riid, &IID_IErrorRecords)) {
		*ppvObject = &this->lpIErrorRecordsVtbl;
	} else if (IsEqualIID(riid, &IID_ISQLErrorInfo)) {
		*ppvObject = &this->lpISQLErrorInfoVtbl;
	} else {
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}
	ADDREF((IUnknown *) *ppvObject);
	return S_OK;
}

static ULONG STDMETHODCALLTYPE oledb_error_info_AddRef(oledb_error_info *this)
{
	return ++(this->refcount);
}

static ULONG STDMETHODCALLTYPE oledb_error_info_Release(oledb_error_info *this)
{
	if(--(this->refcount) == 0) {
		SysFreeString(this->description);
		SysFreeString(this->sqlcode);
		CoTaskMemFree(this);
		return 0;
	}
	return this->refcount;
}

static HRESULT STDMETHODCALLTYPE IErrorInfo_QueryInterface(IErrorInfo *ptr,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject)
{
	DEFINE_THIS(ptr, 0);
	return oledb_error_info_QueryInterface(this, riid, ppvObject);
}

static ULONG STDMETHODCALLTYPE IErrorInfo_AddRef(IErrorInfo *ptr)
{
	DEFINE_THIS(ptr, 0);
	return oledb_error_info_AddRef(this);
}

static ULONG STDMETHODCALLTYPE IErrorInfo_Release(IErrorInfo *ptr)
{
	DEFINE_THIS(ptr, 0);
	return oledb_error_info_Release(this);;
}

static HRESULT STDMETHODCALLTYPE IErrorInfo_GetGUID(IErrorInfo *ptr,
			GUID *pGUID)
{
	DEFINE_THIS(ptr, 0);
	ZeroMemory(pGUID, sizeof(*pGUID));
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE IErrorInfo_GetSource(IErrorInfo *ptr,
			BSTR *pBstrSource)
{
	DEFINE_THIS(ptr, 0);
	*pBstrSource = SysAllocString(L"PDO-OLEDB");
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE IErrorInfo_GetDescription(IErrorInfo *ptr,
			BSTR *pBstrDescription)
{
	DEFINE_THIS(ptr, 0);
	*pBstrDescription = SysAllocString(((oledb_error_info *) this)->description);
	return S_OK;
}

static HRESULT STDMETHODCALLTYPE IErrorInfo_GetHelpFile(IErrorInfo *ptr,
			BSTR *pBstrHelpFile)
{
	DEFINE_THIS(ptr, 0);
	*pBstrHelpFile = NULL;
	return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE IErrorInfo_GetHelpContext(IErrorInfo *ptr,
			DWORD *pdwHelpContext)
{
	DEFINE_THIS(ptr, 0);
	return E_NOTIMPL;
}

IErrorInfoVtbl IErrorInfo_Vtbl = {
	IErrorInfo_QueryInterface,
	IErrorInfo_AddRef,
	IErrorInfo_Release,
	IErrorInfo_GetGUID,
	IErrorInfo_GetSource,
	IErrorInfo_GetDescription,
	IErrorInfo_GetHelpFile,
	IErrorInfo_GetHelpContext
};

static HRESULT STDMETHODCALLTYPE IErrorRecords_QueryInterface(IErrorRecords *ptr,
			REFIID riid,
			void **ppvObject)
{
	DEFINE_THIS(ptr, 1);
	return oledb_error_info_QueryInterface(this, riid, ppvObject);
}

static ULONG STDMETHODCALLTYPE IErrorRecords_AddRef(IErrorRecords *ptr)
{
	DEFINE_THIS(ptr, 1);
	return oledb_error_info_AddRef(this);
}

static ULONG STDMETHODCALLTYPE IErrorRecords_Release(IErrorRecords *ptr)
{
	DEFINE_THIS(ptr, 1);
	return oledb_error_info_Release(this);;
}

        
static HRESULT STDMETHODCALLTYPE IErrorRecords_AddErrorRecord(IErrorRecords *ptr,
			ERRORINFO *pErrorInfo,
			DWORD dwLookupID,
			DISPPARAMS *pdispparams,
			IUnknown *punkCustomError,
			DWORD dwDynamicErrorID)
{
	DEFINE_THIS(ptr, 1);
	return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE IErrorRecords_GetBasicErrorInfo(IErrorRecords *ptr,
			ULONG ulRecordNum,
			ERRORINFO *pErrorInfo)
{
	DEFINE_THIS(ptr, 1);
	return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE IErrorRecords_GetCustomErrorObject(IErrorRecords *ptr,
			ULONG ulRecordNum,
			REFIID riid,
			IUnknown **ppObject)
{
	return QUERY_INTERFACE(ptr, IID_ISQLErrorInfo, *ppObject);
}
        
static HRESULT STDMETHODCALLTYPE IErrorRecords_GetErrorInfo(IErrorRecords *ptr,
			ULONG ulRecordNum,
			LCID lcid,
			IErrorInfo **ppErrorInfo)
{
	return QUERY_INTERFACE(ptr, IID_IErrorInfo, *ppErrorInfo);
}

static HRESULT STDMETHODCALLTYPE IErrorRecords_GetErrorParameters(IErrorRecords *ptr,
			ULONG ulRecordNum,
			DISPPARAMS *pdispparams)
{
	return E_NOTIMPL;
}

static HRESULT STDMETHODCALLTYPE IErrorRecords_GetRecordCount(IErrorRecords *ptr,
			ULONG *pcRecords)
{
	*pcRecords = 1;
	return S_OK;
}

IErrorRecordsVtbl IErrorRecords_Vtbl = {
	IErrorRecords_QueryInterface,
	IErrorRecords_AddRef,
	IErrorRecords_Release,
	IErrorRecords_AddErrorRecord,
	IErrorRecords_GetBasicErrorInfo,
	IErrorRecords_GetCustomErrorObject,
	IErrorRecords_GetErrorInfo,
	IErrorRecords_GetErrorParameters,
	IErrorRecords_GetRecordCount,
};

static HRESULT STDMETHODCALLTYPE ISQLErrorInfo_QueryInterface(ISQLErrorInfo *ptr,
			REFIID riid,
			void **ppvObject)
{
	DEFINE_THIS(ptr, 2);
	return oledb_error_info_QueryInterface(this, riid, ppvObject);
}

static ULONG STDMETHODCALLTYPE ISQLErrorInfo_AddRef(ISQLErrorInfo *ptr)
{
	DEFINE_THIS(ptr, 2);
	return oledb_error_info_AddRef(this);
}

static ULONG STDMETHODCALLTYPE ISQLErrorInfo_Release(ISQLErrorInfo *ptr)
{
	DEFINE_THIS(ptr, 2);
	return oledb_error_info_Release(this);;
}

static HRESULT STDMETHODCALLTYPE ISQLErrorInfo_GetSQLInfo(ISQLErrorInfo *ptr,
			BSTR *pbstrSQLState,
			LONG *plNativeError)
{
	DEFINE_THIS(ptr, 2);
	*pbstrSQLState = SysAllocString(this->sqlcode);
	*plNativeError = -1;
	return S_OK;
}

ISQLErrorInfoVtbl ISQLErrorInfo_Vtbl = {
	ISQLErrorInfo_QueryInterface,
	ISQLErrorInfo_AddRef,
	ISQLErrorInfo_Release,
	ISQLErrorInfo_GetSQLInfo,
};

void oledb_set_automation_error(LPCWSTR msg, LPCWSTR sqlcode) {
	oledb_error_info *obj = (oledb_error_info *) CoTaskMemAlloc(sizeof(*obj));
	IUnknown *pIErrorInfo = NULL;
	obj->lpIErrorInfoVtbl = &IErrorInfo_Vtbl;
	obj->lpIErrorRecordsVtbl = &IErrorRecords_Vtbl;
	obj->lpISQLErrorInfoVtbl = &ISQLErrorInfo_Vtbl;
	obj->description = SysAllocString(msg);
	obj->sqlcode = SysAllocString(sqlcode);
	obj->refcount = 0;
	
	QUERY_INTERFACE((IUnknown *) obj, IID_IErrorInfo, pIErrorInfo);
	SetErrorInfo(0, (IErrorInfo *) obj);
	SAFE_RELEASE(pIErrorInfo);
}