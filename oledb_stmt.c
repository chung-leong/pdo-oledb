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

static void oledb_stmt_clear_rowset(pdo_stmt_t *stmt TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	int i;

	if (S->hRow) {
		HROW *hRows = &S->hRow;
		CALL(ReleaseRows, S->pIRowset, 1, hRows, NULL, NULL, NULL);
	}
	if (S->pIAccessorRowset) {
		if(S->hAccessorRowset) {
			CALL(ReleaseAccessor, S->pIAccessorRowset, S->hAccessorRowset, NULL);
		}
		for(i = 0; i < stmt->column_count; i++) {
			if(S->columns[i].hAccessorColumn) {
				CALL(ReleaseAccessor, S->pIAccessorRowset, S->columns[i].hAccessorColumn, NULL);
			}
		}
		RELEASE(S->pIAccessorRowset);
	}
	for(i = 0; i < stmt->column_count; i++) {
		pdo_oledb_column *C = &S->columns[i];
		SAFE_EFREE(C->name);
		if (C->metadata) {
			pdo_oledb_column_meta_data *m = C->metadata;
			SAFE_EFREE(m->catalogName);
			SAFE_EFREE(m->tableName);
			SAFE_EFREE(m->columnName);
			if (m->defaultValue) {
				zval_ptr_dtor(&m->defaultValue);
			}
		}
		oledb_release_conversion_options(C->conv);
	}
	SAFE_EFREE(S->columns);
	SAFE_RELEASE(S->pIRowset);
	S->columns = NULL;
	stmt->column_count = 0;

	S->pIRowset = NULL;
	S->pIAccessorRowset = NULL;
	S->hAccessorRowset = 0;
	S->hRow = 0;
	S->rowIndex = 0;
}

static int oledb_stmt_dtor(pdo_stmt_t *stmt TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	oledb_stmt_clear_rowset(stmt TSRMLS_CC);
	SAFE_RELEASE(S->pIMultipleResults);
	SAFE_RELEASE(S->pICommand);
	SAFE_RELEASE(S->pICommandWithParameters);
	CoTaskMemFree(S->paramInfo);
	CoTaskMemFree(S->paramNamesBuffer);
	if (S->pIAccessorCommand) {
		if(S->hAccessorCommand) {
			CALL(ReleaseAccessor, S->pIAccessorCommand, S->hAccessorCommand, NULL);
		}
		RELEASE(S->pIAccessorCommand);
	}
	SAFE_EFREE(S->inputBuffer);
	SAFE_EFREE(S->outputBuffer);
	if (S->einfo.errmsg) {
		efree(S->einfo.errmsg);
	}
	oledb_release_conversion_options(S->conv);
	efree(S);
	return 1;
}

/* Advance a byte offset, maintaining 32-bit alignment */
#define ADVANCE_OFFSET(i, a)	{ i += a; if (i & 0x0003) i = (i+3) & ~0x0003; }

static HRESULT oledb_stmt_bind_column(pdo_stmt_t *stmt, pdo_oledb_column *C TSRMLS_DC)
{
	HRESULT hr = S_OK;
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;

	if (C->columnFlags & DBCOLUMNFLAGS_ISLONG) {
		/* each blob column needs its own accessor */
		DBBINDING blob_binding;
		DBBINDSTATUS blob_bind_status;
		DBOBJECT dbo;

		ZeroMemory(&blob_binding, sizeof(blob_binding));
		blob_binding.iOrdinal = C->ordinal;
		blob_binding.wType = DBTYPE_IUNKNOWN;
		blob_binding.obStatus = 0;
		blob_binding.obLength = sizeof(DWORD);
		blob_binding.obValue = sizeof(DWORD) + sizeof(DBLENGTH);
		blob_binding.dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;
		blob_binding.pObject = &dbo;
		dbo.dwFlags = STGM_READ;
		dbo.iid = IID_ISequentialStream;

		hr = CALL(CreateAccessor, S->pIAccessorRowset, DBACCESSOR_ROWDATA, 1, &blob_binding, 0, &C->hAccessorColumn, &blob_bind_status);
		if (!SUCCEEDED(hr)) goto cleanup;

		switch (C->columnType & ~DBTYPE_BYREF) {
			case DBTYPE_WSTR:
			case DBTYPE_STR:
			case DBTYPE_BSTR:
				C->pdoType = PDO_PARAM_STR;
				break;
			case DBTYPE_BYTES:
			default:
				C->pdoType = PDO_PARAM_LOB;
				break;
		}

		C->retrievalType = DBTYPE_IUNKNOWN;
		C->flags |= VARIABLE_LENGTH;
	} else {
		/* non-blob columns are retrieved from a single buffer */
		C->retrievalType = C->columnType;
		if (C->columnType & DBCOLUMNFLAGS_ISFIXEDLENGTH) {
			C->flags &= ~VARIABLE_LENGTH;
		} else {
			C->flags |= VARIABLE_LENGTH;
		}

		switch (C->columnType & ~DBTYPE_BYREF) {
			case DBTYPE_WSTR:	/* nvarchar */
				C->pdoType = PDO_PARAM_STR;
				C->byteCount = (C->maxLen + 1) * 2;
				C->flags |= VARIABLE_LENGTH;
			break;
			case DBTYPE_BYTES:	/* varbinary */
				C->pdoType = PDO_PARAM_STR;
				C->byteCount = C->maxLen + 1;
				C->flags |= VARIABLE_LENGTH;
			break;
			case DBTYPE_GUID:	/* uniqueidentifier */
			case DBTYPE_STR:	/* varchar */
				C->pdoType = PDO_PARAM_STR;
				C->byteCount = C->maxLen + 1;
				C->flags |= VARIABLE_LENGTH;
			break;
			case DBTYPE_BOOL:			/* bit */
				C->pdoType = PDO_PARAM_BOOL;
				C->retrievalType = DBTYPE_UI1;
				C->byteCount = sizeof(unsigned char);
			break;
			case DBTYPE_UI1:			/* tinyint */
			case DBTYPE_I1:				/* not used by SQL Server */
			case DBTYPE_UI2:			/* not used by SQL Server */
			case DBTYPE_I2:				/* smallint */
			case DBTYPE_I4:				/* int */
				C->pdoType = PDO_PARAM_INT;
				C->retrievalType = DBTYPE_I4;
				C->byteCount = sizeof(DWORD);
			break;
			case DBTYPE_R4:				/* real */
			case DBTYPE_R8:				/* float */
			case DBTYPE_CY:				/* money, smallmoney */
			case DBTYPE_NUMERIC:		/* decimal */
				/* fetch these types as a double */
				C->pdoType = PDO_PARAM_STR;
				C->retrievalType = DBTYPE_R8;
				C->byteCount = sizeof(double);
			break;
			case DBTYPE_UI4:			/* not used by SQL Server */
				C->pdoType = PDO_PARAM_STR;
				C->byteCount = sizeof(DWORD);
			break;
			case DBTYPE_DATE:			/* not used by SQL Server */
				C->retrievalType = DBTYPE_DBTIMESTAMP;
			case DBTYPE_DBTIMESTAMP:	/* datetime, smalldatetime */
				C->pdoType = PDO_PARAM_STR;
				if (!S->flags & CONVERT_DATE_TIME) {
					C->byteCount = sizeof(DBTIMESTAMP);
				} else {
					/* let SQL Server do the conversion to string */
					C->retrievalType = DBTYPE_STR;
					C->byteCount = C->precision + 1;
				}
			break;
			case DBTYPE_VARIANT:
				C->pdoType = PDO_PARAM_STR;
				C->byteCount = sizeof(VARIANT);
				break;
			default:
				/* retrieve everything else as strings */
				C->pdoType = PDO_PARAM_STR;
				C->retrievalType = DBTYPE_STR;
				C->byteCount = max(C->precision, (int) C->maxLen) + 1;
				C->flags |= VARIABLE_LENGTH;
			break;
		}

		/* column is actually returned by ref */
		if (C->columnFlags & DBTYPE_BYREF) {
			C->retrievalType = C->columnType;
			C->byteCount = sizeof(void *);
		}
	}
cleanup: 
	return hr;
}

static HRESULT oledb_stmt_get_column_meta_data(pdo_stmt_t *stmt TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;

	HRESULT hr;
	IColumnsRowset *pIColumnsRowset = NULL;
	IRowset *pMetaIRowset = NULL;
	IAccessor *pMetaIAccessor = NULL;
	IColumnsInfo *pMetaIColumnsInfo = NULL;

	DBCOUNTITEM rows_obtained;
	HACCESSOR hMetaAccessor = 0;
	HROW hRow = 0, *hRows;

	DBORDINAL meta_column_count, i;
	DBCOLUMNINFO *meta_column_info = NULL;
	OLECHAR *meta_column_info_buffer = NULL;

	DBBINDING *meta_bindings = NULL;
	DBBINDSTATUS *meta_bind_status = NULL;
	DBLENGTH next_offset = 0;

	int row_num = 0;
	char *buffer = NULL;

	/* get an IRowset to the metadata table */
	hr = QUERY_INTERFACE(S->pIRowset, IID_IColumnsRowset, pIColumnsRowset);
	if (!pIColumnsRowset) goto cleanup;

	hr = CALL(GetColumnsRowset, pIColumnsRowset, NULL, 0, NULL, &IID_IRowset, 0, NULL, (IUnknown **) &pMetaIRowset);
	if (!pMetaIRowset) goto cleanup;

	hr = QUERY_INTERFACE(pMetaIRowset, IID_IColumnsInfo, pMetaIColumnsInfo);
	if (!pMetaIColumnsInfo) goto cleanup;

	/* see what columns are there */
	hr = CALL(GetColumnInfo, pMetaIColumnsInfo, &meta_column_count, &meta_column_info, &meta_column_info_buffer);
	if (!SUCCEEDED(hr)) goto cleanup;

	/* create an accessor to the meta data */
	hr = QUERY_INTERFACE(pMetaIRowset, IID_IAccessor, pMetaIAccessor);
	if (!pMetaIAccessor) goto cleanup;

	meta_bindings = ecalloc(meta_column_count, sizeof(*meta_bindings));
	meta_bind_status = ecalloc(meta_column_count, sizeof(*meta_bind_status));

	for (i = 0; i < meta_column_count; i++) {
		DBCOLUMNINFO *m = &meta_column_info[i];
		DBBINDING *b = &meta_bindings[i];

		switch(m->wType & ~DBTYPE_BYREF) {
			case DBTYPE_WSTR:
				b->cbMaxLen = (m->ulColumnSize + 1) * sizeof(WCHAR);
				break;
			case DBTYPE_UI2:
			case DBTYPE_BOOL:
			case DBTYPE_I4:
			case DBTYPE_UI4:
				b->cbMaxLen = sizeof(DWORD);
				break;
			case DBTYPE_I8:
			case DBTYPE_UI8:
				b->cbMaxLen = sizeof(DWORD) * 2;
				break;
			case DBTYPE_VARIANT:
				b->cbMaxLen = sizeof(VARIANT);
				break;
			case DBTYPE_GUID:
				b->cbMaxLen = sizeof(GUID);
				break;
			case DBTYPE_IUNKNOWN:
				b->cbMaxLen = sizeof(IUnknown *);
				break;

		}
		if (m->wType & DBTYPE_BYREF) {
			b->cbMaxLen = sizeof(void *);
		}

		b->wType = m->wType;
		b->iOrdinal = m->iOrdinal;
		b->dwPart = DBPART_VALUE | DBPART_LENGTH | DBPART_STATUS;
		b->obStatus = next_offset; ADVANCE_OFFSET(next_offset, sizeof(DWORD));
		b->obLength = next_offset; ADVANCE_OFFSET(next_offset, sizeof(DBLENGTH));
		b->obValue = next_offset; ADVANCE_OFFSET(next_offset, b->cbMaxLen);
		if (m->wType & DBTYPE_BYREF) {
			b->dwMemOwner = DBMEMOWNER_PROVIDEROWNED;
		}
	}

	hr = CALL(CreateAccessor, pMetaIAccessor, DBACCESSOR_ROWDATA, meta_column_count, meta_bindings, 0, &hMetaAccessor, meta_bind_status);
	if (!hMetaAccessor) goto cleanup;

	/* allocate buffer */
	buffer = emalloc(next_offset);

	hRows = &hRow;
	while(hr = CALL(GetNextRows, pMetaIRowset, DB_NULL_HCHAPTER, 0, 1, &rows_obtained, &hRows), rows_obtained > 0) {
		hr = CALL(GetData, pMetaIRowset, hRow, hMetaAccessor, buffer);
		CALL(ReleaseRows, pMetaIRowset, rows_obtained, hRows, NULL, NULL, NULL);
		
		if (SUCCEEDED(hr)) {
			pdo_oledb_column *C = &S->columns[row_num];
			pdo_oledb_column_meta_data *D = ecalloc(1, sizeof(*D));
			WCHAR *possible_alias = NULL;

			C->metadata = D;

			for (i = 0; i < meta_column_count; i++) {
				DBBINDING *b = &meta_bindings[i];
				DBID *id = &meta_column_info[i].columnid;
				DWORD *pStatus = (DWORD *) (buffer + b->obStatus);
				DBLENGTH *pLength = (DBLENGTH *) (buffer + b->obLength);
				char *pValue = buffer + b->obValue;

				if (b->wType & DBTYPE_BYREF) {
					pValue = *((char **) pValue);
				}

				if (*pStatus == DBSTATUS_S_OK || *pStatus == DBSTATUS_S_TRUNCATED) {
					if (memcmp(id, &DBCOLUMN_BASECOLUMNNAME, sizeof(DBID)) == 0) {
						oledb_convert_bstr(S->conv, (WCHAR *) (pValue), *pLength / sizeof(WCHAR), &D->columnName, &D->columnNameLen, CONVERT_FROM_UNICODE_TO_OUTPUT);
						/* see if column was aliased */
						if (possible_alias && wcscmp(possible_alias, (WCHAR *) pValue) != 0) {
							C->flags |= ALIESED_COLUMN;
						}
					} else if (memcmp(id, &DBCOLUMN_BASETABLENAME, sizeof(DBID)) == 0) {
						oledb_convert_bstr(S->conv, (WCHAR *) (pValue), *pLength / sizeof(WCHAR), &D->tableName, &D->tableNameLen, CONVERT_FROM_UNICODE_TO_OUTPUT);
					} else if (memcmp(id, &DBCOLUMN_BASECATALOGNAME, sizeof(DBID)) == 0) {
						oledb_convert_bstr(S->conv, (WCHAR *) (pValue), *pLength / sizeof(WCHAR), &D->catalogName, &D->catalogNameLen, CONVERT_FROM_UNICODE_TO_OUTPUT);
					} else if (memcmp(id, &DBCOLUMN_COMPUTEMODE, sizeof(DBID)) == 0) {
						DWORD mode = *((DWORD *) pValue);
						D->isDynamic = (mode == DBCOMPUTEMODE_COMPUTED || mode == DBCOMPUTEMODE_DYNAMIC);
					} else if (memcmp(id, &DBCOLUMN_ISAUTOINCREMENT, sizeof(DBID)) == 0) {
						D->isAutoIncrement = *((VARIANT_BOOL *) pValue) ? TRUE : FALSE;
					} else if (memcmp(id, &DBCOLUMN_NAME, sizeof(DBID)) == 0) {
						possible_alias = (WCHAR *) pValue;
					} else if (memcmp(id, &DBCOLUMN_DEFAULTVALUE, sizeof(DBID)) == 0) {
						MAKE_STD_ZVAL(D->defaultValue);
						ZVAL_NULL(D->defaultValue);
					} else if (memcmp(id, &DBCOLUMN_ISUNIQUE, sizeof(DBID)) == 0) {
						D->isUnique = *((VARIANT_BOOL *) pValue) ? TRUE : FALSE;
					} else if (memcmp(id, &DBCOLUMN_KEYCOLUMN, sizeof(DBID)) == 0) {
						D->isKey = *((VARIANT_BOOL *) pValue) ? TRUE : FALSE;
					}
				}
			}			
			row_num++;
		} else {
			break;
		}
	}

cleanup:
	if (pMetaIAccessor) {
		if (hMetaAccessor) {
			CALL(ReleaseAccessor, pMetaIAccessor, hMetaAccessor, NULL);
		}
		RELEASE(pMetaIAccessor);
	}
	SAFE_RELEASE(pMetaIRowset);
	SAFE_RELEASE(pMetaIColumnsInfo);
	SAFE_RELEASE(pIColumnsRowset);
	SAFE_EFREE(buffer);
	SAFE_EFREE(meta_bindings);
	SAFE_EFREE(meta_bind_status);
	return hr;
}

BOOL oledb_stmt_build_column_name(pdo_oledb_stmt *S, pdo_oledb_column *C) {
	if (S->flags & (ADD_TABLE_NAME | ADD_CATALOG_NAME)) {
		if (!(C->flags & ALIESED_COLUMN)) {
			if (C->metadata) {
				char *parts[3];
				unsigned int lengths[3];
				int count = 0;

				if ((S->flags & ADD_CATALOG_NAME) && C->metadata->catalogName) {
					parts[count] = C->metadata->catalogName;
					lengths[count] = C->metadata->catalogNameLen;
					count++;
				}
				if ((S->flags & ADD_TABLE_NAME) && C->metadata->tableName) {
					parts[count] = C->metadata->tableName;
					lengths[count] = C->metadata->tableNameLen;
					count++;
				}
				if (C->metadata->columnName) {
					parts[count] = C->metadata->columnName;
					lengths[count] = C->metadata->columnNameLen;
					count++;
				}
				if (count > 0) {
					unsigned int total = 0;
					int i;
					char *p;
					for(i = 0; i < count; i++) {
						total += lengths[i] + 1;
					}
					C->name = p = emalloc(total);
					C->nameLen = total;
					for(i = 0; i < count; i++) {
						memcpy(p, parts[i], lengths[i]);
						p += lengths[i];
						*p++ = '.';
					}
					p[-1] = '\0';
					return TRUE;
				}
			}
		}
	}
	return FALSE;
}

static HRESULT oledb_stmt_create_columns(pdo_stmt_t *stmt TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;

	HRESULT hr;
	IColumnsInfo *pIColumnsInfo = NULL;

	DBCOLUMNINFO *col_info = NULL;
	DBORDINAL col_count;
	OLECHAR *col_info_buffer = NULL;
	int i;

	hr = QUERY_INTERFACE(S->pIRowset, IID_IColumnsInfo, pIColumnsInfo);
	if (!pIColumnsInfo) goto cleanup;

	hr = CALL(GetColumnInfo, pIColumnsInfo, &col_count, &col_info, &col_info_buffer);
	if (!SUCCEEDED(hr)) goto cleanup;

	S->columns = ecalloc(col_count, sizeof(*S->columns));
	stmt->column_count = col_count;

	if (S->flags & (ADD_TABLE_NAME | ADD_CATALOG_NAME)) {
		oledb_stmt_get_column_meta_data(stmt TSRMLS_CC);
	}

	for (i = 0; i < stmt->column_count; i++) {
		pdo_oledb_column *C = &S->columns[i];
		if (!oledb_stmt_build_column_name(S, C)) {
			oledb_convert_bstr(S->conv, col_info[i].pwszName, -1, &C->name, &C->nameLen, CONVERT_FROM_UNICODE_TO_OUTPUT);
		}
		C->columnLength = col_info[i].ulColumnSize;
		C->columnFlags = col_info[i].dwFlags;
		C->precision = col_info[i].bPrecision;
		C->maxLen = col_info[i].ulColumnSize;
		C->columnType = col_info[i].wType;
		C->ordinal = col_info[i].iOrdinal;
		oledb_copy_conversion_options(&C->conv, S->conv);
	}

cleanup:
	SAFE_RELEASE(pIColumnsInfo);
	CoTaskMemFree(col_info);
	CoTaskMemFree(col_info_buffer);
	return hr;
}

static HRESULT oledb_stmt_bind_columns(pdo_stmt_t *stmt TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;
	int i, j;

	HRESULT hr;
	DBBINDING *bindings = NULL;
	DBBINDSTATUS *bind_statuses = NULL;

	hr = oledb_stmt_create_columns(stmt TSRMLS_CC);
	if (!SUCCEEDED(hr)) goto cleanup;

	hr = QUERY_INTERFACE(S->pIRowset, IID_IAccessor, S->pIAccessorRowset);
	if (!S->pIAccessorRowset) goto cleanup;

	/* allocate the binding structure */
	bindings = ecalloc(stmt->column_count, sizeof(*bindings));
	bind_statuses = ecalloc(stmt->column_count, sizeof(*bind_statuses));

	S->nextOutputOffset = 0;
	for (i = 0, j = 0; i < stmt->column_count; i++) {
		pdo_oledb_column *C = &S->columns[i];
		hr = oledb_stmt_bind_column(stmt, C TSRMLS_CC);
		if (!SUCCEEDED(hr)) goto cleanup;

		if (C->byteCount > 0) {
			/* remember where the column is stored within the buffer */
			C->byteOffset = S->nextOutputOffset;

			/* set the binding type and column ordinal */
			bindings[j].iOrdinal = C->ordinal;
			bindings[j].wType = C->retrievalType;
			if (C->retrievalType & DBTYPE_BYREF) {
				bindings[j].dwMemOwner = DBMEMOWNER_PROVIDEROWNED;
			}

			/* retrieve the column status */
			bindings[j].dwPart = DBPART_STATUS;
			bindings[j].obStatus = S->nextOutputOffset;
			ADVANCE_OFFSET(S->nextOutputOffset, sizeof(DWORD));

			if (C->flags & VARIABLE_LENGTH) {
				/* retrieve the length too */
				bindings[j].dwPart |= DBPART_LENGTH;
				bindings[j].obLength = S->nextOutputOffset;
				ADVANCE_OFFSET(S->nextOutputOffset, sizeof(DBLENGTH));
			}

			bindings[j].dwPart |= DBPART_VALUE;
			bindings[j].obValue = S->nextOutputOffset;
			bindings[j].cbMaxLen = C->byteCount;
			ADVANCE_OFFSET(S->nextOutputOffset, C->byteCount);

			j++;
		}
	}
	hr = S_OK;

	if (j > 0) {
		/* create accessor for non-blob columns*/
		hr = CALL(CreateAccessor, S->pIAccessorRowset, DBACCESSOR_ROWDATA, j, bindings, 0, &S->hAccessorRowset, bind_statuses);
		if (!SUCCEEDED(hr)) goto cleanup;

		/* allocate buffer */
		S->outputBuffer = erealloc(S->outputBuffer, S->nextOutputOffset);
	} else {
		hr = S_OK;
	}

cleanup: 
	SAFE_EFREE(bindings);
	SAFE_EFREE(bind_statuses);
	if (!SUCCEEDED(hr)) {
		oledb_stmt_clear_rowset(stmt TSRMLS_CC);
	}
	return hr;
}

static void oledb_stmt_clear_param(pdo_stmt_t *stmt, struct pdo_bound_param_data *param TSRMLS_DC)
{
	pdo_oledb_param *P = (pdo_oledb_param*)param->driver_data;

	if (P) {
		SysFreeString(P->unicodeValue);
		SAFE_EFREE(P->varcharValue);
		if (P->stream) {
			if (RELEASE(P->stream) > 0) {
				/* release the stream a second time if the provider hasn't done so */
				RELEASE(P->stream);
			}
		}
		P->varcharValue = NULL;
		P->unicodeValue = NULL;
		P->stream = NULL;
	}
}

static BOOL oledb_stmt_is_variable_width_column(DWORD type) {
	switch(type) {
		case DBTYPE_STR:
		case DBTYPE_WSTR:
		case DBTYPE_BYTES:
		case DBTYPE_VARNUMERIC:
			return TRUE;
	}
	return FALSE;
}

static HRESULT oledb_stmt_bind_param(pdo_stmt_t *stmt, struct pdo_bound_param_data *param TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;
	pdo_oledb_param *P = (pdo_oledb_param*)param->driver_data;
	zval *value = param->parameter;

	HRESULT hr = S_OK;
	DBPARAMINFO param_info;
	
	if (S->paramInfo && param->paramno >= 0 && param->paramno < (int) S->paramCount) {
		/* we know something about the parameter */
		param_info = S->paramInfo[param->paramno];
	} else {
		/* we're in the dark...make an educated guess */
		ZeroMemory(&param_info, sizeof(param_info));
		param_info.dwFlags = DBPARAMFLAGS_ISINPUT;
		param_info.iOrdinal = param->paramno + 1;
	}

	/* don't allow output if flags isn't specified--to be consistent */
	if (param->param_type & PDO_PARAM_INPUT_OUTPUT) {
		param_info.dwFlags |= DBPARAMFLAGS_ISOUTPUT;
	} else {
		param_info.dwFlags &= ~DBPARAMFLAGS_ISOUTPUT;
	}

	/* see if string needs to be converted to Unicode or not */
	if (param_info.wType == DBTYPE_WSTR
	 || param_info.wType == DBTYPE_BSTR) {
		 /* the oledb provider said so... */
		P->flags |= STRING_AS_UNICODE;
	} 

	/* see if string contains binary data or not */
	if (param_info.wType == DBTYPE_BYTES) {
		P->flags |= STRING_AS_LOB;
	} 

	/* see if param is in or out */
	if (param_info.dwFlags & DBPARAMFLAGS_ISINPUT) {
		/* set the data type based on what's actually in the zval */
		if (Z_TYPE_P(value) == IS_STRING) {
			P->flags |= VARIABLE_LENGTH;
			if (P->flags & STRING_AS_UNICODE) {
				if (param_info.dwFlags & DBPARAMFLAGS_ISLONG) {
					P->dataType = L"DBTYPE_WLONGVARCHAR";
					P->dataTypeWidth = ~0;
					P->retrievalType = DBTYPE_IUNKNOWN;
					P->dataPointer = &P->stream;
					P->byteCount = sizeof(IUnknown);
					oledb_create_zval_stream(P->conv, value, CONVERT_FROM_INPUT_TO_UNICODE, &P->stream, &P->dataLength TSRMLS_CC);
				} else {
					UINT unicode_len;
					oledb_create_bstr(P->conv, Z_STRVAL_P(value), Z_STRLEN_P(value), &P->unicodeValue, &unicode_len, CONVERT_FROM_INPUT_TO_UNICODE);
					if (P->flags & TRUNCATE_STRING) {
						if (oledb_stmt_is_variable_width_column(param_info.wType) && param_info.ulParamSize < unicode_len) {
							unicode_len = param_info.ulParamSize;
						}
					}
					P->dataType = L"DBTYPE_WVARCHAR";
					P->dataTypeWidth = unicode_len;
					P->retrievalType = DBTYPE_WSTR;
					P->dataPointer = P->unicodeValue;
					P->byteCount = P->dataLength = unicode_len * sizeof(WCHAR);
				}
			} else {
				if (param_info.dwFlags & DBPARAMFLAGS_ISLONG) {
					P->dataType = (P->flags & STRING_AS_LOB) ? L"DBTYPE_VARBINARY" : L"DBTYPE_LONGVARCHAR";
					P->dataTypeWidth = ~0;
					P->retrievalType = DBTYPE_IUNKNOWN;
					P->dataPointer = &P->stream;
					P->byteCount = sizeof(IUnknown);
					oledb_create_zval_stream(P->conv, value, (P->flags & STRING_AS_LOB) ? -1 : CONVERT_FROM_INPUT_TO_VARCHAR, &P->stream, &P->dataLength TSRMLS_CC);
				} else {
					char *s = Z_STRVAL_P(value);
					UINT len = Z_STRLEN_P(value);
					if (P->flags & STRING_AS_LOB) {
						P->dataType = L"DBTYPE_BYTES";
					} else {
						P->dataType = L"DBTYPE_VARCHAR";

						/* need to convert */
						if (oledb_convert_string(P->conv, s, len, &s, &len, CONVERT_FROM_INPUT_TO_VARCHAR) != S_FALSE) {
							/* buffer needs to be freed */
							P->varcharValue = s;
						}
					}
					if (P->flags & TRUNCATE_STRING) {
						if (oledb_stmt_is_variable_width_column(param_info.wType) && param_info.ulParamSize < len) {
							if (P->flags & STRING_AS_LOB) {
								len = param_info.ulParamSize;
							} else {
								/* make sure multibyte strings don't get truncated mid-character */
								LPCSTR charset = (P->conv->varcharCharset) ? P->conv->varcharCharset : P->conv->charset;
								len = oledb_get_proper_truncated_length(s, param_info.ulParamSize, charset);
							}
						}
					}
					P->dataTypeWidth = len;
					P->retrievalType = DBTYPE_STR;
					P->dataPointer = s;
					P->byteCount = P->dataLength = len;
				}
			}
		} else if (Z_TYPE_P(value) == IS_RESOURCE) {
			P->flags |= VARIABLE_LENGTH;
			if (P->flags & STRING_AS_UNICODE) {
				P->dataType = L"DBTYPE_WLONGVARCHAR";
				P->dataTypeWidth = ~0;
				P->retrievalType = DBTYPE_IUNKNOWN;
				P->dataPointer = &P->stream;
				P->byteCount = sizeof(IUnknown);
				hr = oledb_create_zval_stream(P->conv, value, CONVERT_FROM_INPUT_TO_UNICODE, &P->stream, &P->dataLength TSRMLS_CC);
			} else {
				P->dataType = (P->flags & STRING_AS_LOB) ? L"DBTYPE_VARBINARY" : L"DBTYPE_LONGVARCHAR";
				P->dataTypeWidth = ~0;
				P->retrievalType = DBTYPE_IUNKNOWN;
				P->dataPointer = &P->stream;
				P->byteCount = sizeof(IUnknown);
				hr = oledb_create_zval_stream(P->conv, value, (P->flags & STRING_AS_LOB) ? -1 : CONVERT_FROM_INPUT_TO_VARCHAR, &P->stream, &P->dataLength TSRMLS_CC);
			}
		} else if (Z_TYPE_P(value) == IS_LONG) {
			P->dataType = L"DBTYPE_I4";
			P->dataTypeWidth = 4;
			P->retrievalType = DBTYPE_I4;
			P->intValue = Z_LVAL_P(value);
			P->dataPointer = &P->intValue;
			P->byteCount = sizeof(int);
			P->flags &= ~VARIABLE_LENGTH;
		} else if (Z_TYPE_P(value) == IS_BOOL) {
			P->dataType = L"DBTYPE_BOOL";
			P->dataTypeWidth = 4;
			P->retrievalType = DBTYPE_I4;
			P->intValue = Z_LVAL_P(value);
			P->dataPointer = &P->intValue;
			P->byteCount = sizeof(int);		
			P->flags &= ~VARIABLE_LENGTH;
		} else if (Z_TYPE_P(value) == IS_DOUBLE) {
			P->dataType = L"DBTYPE_R8";
			P->dataTypeWidth = 8;
			P->retrievalType = DBTYPE_R8;
			P->doubleValue = Z_DVAL_P(value);
			P->dataPointer = &P->doubleValue;
			P->byteCount = sizeof(double);
			P->flags &= ~VARIABLE_LENGTH;
		} else if (Z_TYPE_P(value) == IS_NULL) {
			if (!(param_info.dwFlags & DBPARAMFLAGS_ISINPUT)) {
				// set to empty string if it isn't an output param
				P->dataType = L"DBTYPE_WVARCHAR";
				P->retrievalType = DBTYPE_WSTR;
				P->dataPointer = NULL;
				P->byteCount = 0;
			}
		} else {
			hr = E_FAIL;
		}

		P->ioFlags |= DBPARAMIO_INPUT;
	}
	if (param_info.dwFlags & DBPARAMFLAGS_ISOUTPUT) {
		if (!P->dataType) {
			/* don't know what the param should be yet */
			if (PDO_PARAM_TYPE(param->param_type) == PDO_PARAM_STR) {
				P->flags |= VARIABLE_LENGTH;
				if (param_info.dwFlags & DBPARAMFLAGS_ISLONG) {
					/* retrieve with IUnknown if provider said so */
					P->dataType = (P->flags & STRING_AS_UNICODE) ? L"DBTYPE_WLONGVARCHAR" : L"DBTYPE_LONGVARCHAR";
					P->dataTypeWidth = ~0;
					P->retrievalType = DBTYPE_IUNKNOWN;
					P->byteCount = sizeof(IUnknown);
				} else {
					P->dataType = (P->flags & STRING_AS_UNICODE) ? L"DBTYPE_WVARCHAR" : L"DBTYPE_VARCHAR";
					P->dataTypeWidth = param->max_value_len;
					P->retrievalType = (P->flags & STRING_AS_UNICODE) ? DBTYPE_WSTR : DBTYPE_STR;
					P->byteCount = (P->flags & STRING_AS_UNICODE) ? param->max_value_len * sizeof(WCHAR) : param->max_value_len;
				}
			}
			else if (PDO_PARAM_TYPE(param->param_type) == PDO_PARAM_LOB) {
				P->flags |= VARIABLE_LENGTH;
				P->dataType = (P->flags & STRING_AS_UNICODE) ? L"DBTYPE_WLONGVARCHAR" : L"DBTYPE_LONGVARCHAR";
				P->dataTypeWidth = ~0;
				P->retrievalType = DBTYPE_IUNKNOWN;
				P->byteCount = sizeof(IUnknown *);
			}
			else if (PDO_PARAM_TYPE(param->param_type) == PDO_PARAM_INT) {
				P->dataType = L"DBTYPE_I4";
				P->dataTypeWidth = 4;
				P->retrievalType = DBTYPE_I4;
				P->byteCount = sizeof(int);
			}
			else if (PDO_PARAM_TYPE(param->param_type) == PDO_PARAM_BOOL) {
				P->dataType = L"DBTYPE_BOOL";
				P->dataTypeWidth = 4;
				P->retrievalType = DBTYPE_I4;
				P->byteCount = sizeof(int);
			} else {
				hr = E_FAIL;
			}
		} else {
			/* make sure there's enough space to accommodate the output as well */
			if (P->retrievalType == DBTYPE_STR) {
				P->byteCount = max((int) P->byteCount, param->max_value_len);
				P->dataTypeWidth = max((int) P->dataTypeWidth, param->max_value_len);
			} else if(P->retrievalType == DBTYPE_WSTR) {
				P->byteCount = max(P->byteCount, param->max_value_len * sizeof(WCHAR));
				P->dataTypeWidth = max((int) P->dataTypeWidth, param->max_value_len);
			}
		}

		P->ioFlags |= DBPARAMIO_OUTPUT;
	}
	P->ordinal = param_info.iOrdinal;

	if (SUCCEEDED(hr)) {
		if (S->pICommandWithParameters) {
			DBPARAMBINDINFO param_bind_info;
			DB_UPARAMS param_ordinal;
			ZeroMemory(&param_bind_info, sizeof(param_bind_info));
			param_bind_info.pwszDataSourceType = (LPWSTR) P->dataType;
			param_bind_info.ulParamSize = P->dataTypeWidth;
			param_bind_info.dwFlags = param_info.dwFlags;
			param_bind_info.bPrecision = param_info.bPrecision;
			param_bind_info.bScale = param_info.bScale;
			param_bind_info.pwszName = param_info.pwszName;
			param_ordinal = P->ordinal;

			hr = CALL(SetParameterInfo, S->pICommandWithParameters, 1, &param_ordinal, &param_bind_info);
		}
	} else {
		oledb_stmt_clear_param(stmt, param TSRMLS_CC);
	}
	return hr;
}

static HRESULT oledb_stmt_copy_bound_params(pdo_stmt_t *stmt, DBPARAMS *params TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;
	HashTable *ht = stmt->bound_params;
	struct pdo_bound_param_data *param;

	HRESULT hr = S_OK;
	DBBINDING *bindings = NULL;
	DBBINDSTATUS *bind_statuses = NULL;
	DBORDINAL param_count = ht->nNumOfElements;
	DBORDINAL i = 0;

	bindings = ecalloc(param_count, sizeof(*bindings));
	bind_statuses = ecalloc(param_count, sizeof(*bind_statuses));

	S->nextInputOffset = 0;

	zend_hash_internal_pointer_reset(ht);
	while (SUCCESS == zend_hash_get_current_data(ht, (void**)&param)) {
		pdo_oledb_param *P = (pdo_oledb_param*)param->driver_data;

		P->byteOffset = S->nextInputOffset;

		bindings[i].iOrdinal = P->ordinal;
		bindings[i].wType = P->retrievalType;
		bindings[i].eParamIO = P->ioFlags;
		if (P->retrievalType & DBTYPE_BYREF) {
			bindings[i].dwMemOwner = DBMEMOWNER_PROVIDEROWNED;
		}

		/* space for the column status */
		bindings[i].dwPart = DBPART_STATUS;
		bindings[i].obStatus = S->nextInputOffset;
		ADVANCE_OFFSET(S->nextInputOffset, sizeof(DWORD));

		if (P->flags & VARIABLE_LENGTH) {
			bindings[i].dwPart |= DBPART_LENGTH;
			bindings[i].obLength = S->nextInputOffset;
			bindings[i].cbMaxLen = P->byteCount;
			ADVANCE_OFFSET(S->nextInputOffset, sizeof(DBLENGTH));
		}

		if (P->byteCount || P->flags & VARIABLE_LENGTH) {
			bindings[i].dwPart |= DBPART_VALUE;
			bindings[i].obValue = S->nextInputOffset;
			ADVANCE_OFFSET(S->nextInputOffset, P->byteCount);
		}

		i++;
		zend_hash_move_forward(ht);
	}

	hr = CALL(CreateAccessor, S->pIAccessorCommand, DBACCESSOR_PARAMETERDATA, i, bindings, S->nextInputOffset, &S->hAccessorCommand, bind_statuses);
	if (!SUCCEEDED(hr)) goto cleanup;

	/* alloc a buffer large enough and copy the data into it */
	S->inputBuffer = erealloc(S->inputBuffer, S->nextInputOffset);
	zend_hash_internal_pointer_reset(ht);
	while (SUCCESS == zend_hash_get_current_data(ht, (void**)&param)) {
		pdo_oledb_param *P = (pdo_oledb_param*)param->driver_data;
		char *pBuffer = ((char *) S->inputBuffer) + P->byteOffset;

		if (P->ioFlags & DBPARAMIO_INPUT) {
			char *pValue;
			DBLENGTH *pLength = NULL;
			DWORD *pStatus = (DWORD *) pBuffer;

			if (P->byteCount) {
				*pStatus = DBSTATUS_S_OK;
				if (P->flags & VARIABLE_LENGTH) {
					pLength = (DBLENGTH *) (pBuffer + sizeof(DWORD));
					*pLength = P->dataLength;
					pValue = pBuffer + sizeof(DWORD) + sizeof(DBLENGTH);
				} else {
					pValue = pBuffer + sizeof(DWORD);
				}
				if (P->dataPointer) {
					memcpy(pValue, P->dataPointer, P->byteCount);
				}

				if (P->stream) {
					/* add ref to it so the stream stays put regardless of whether the operation succeeded or not */
					ADDREF(P->stream);
				}
			} else {
				*pStatus = DBSTATUS_S_ISNULL;
			}
		} else {
			/* just clear the memory if it's not an input parameter */
			ZeroMemory(pBuffer, sizeof(DWORD) + ((P->flags & VARIABLE_LENGTH) ? sizeof(DBLENGTH) : 0) + P->byteCount);
		}
		zend_hash_move_forward(ht);
	}

	params->pData = S->inputBuffer;
	params->cParamSets = 1;
	params->hAccessor = S->hAccessorCommand;

cleanup:
	SAFE_EFREE(bindings);
	SAFE_EFREE(bind_statuses);
	return hr;
}

static HRESULT oledb_stmt_sync_output_params(pdo_stmt_t *stmt TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;
	HashTable *ht = stmt->bound_params;
	struct pdo_bound_param_data *param;

	HRESULT hr = S_OK;
	zend_hash_internal_pointer_reset(ht);
	while (SUCCESS == zend_hash_get_current_data(ht, (void**)&param)) {
		pdo_oledb_param *P = (pdo_oledb_param*)param->driver_data;
		char *pBuffer = ((char *) S->inputBuffer) + P->byteOffset;

		if (P->ioFlags & DBPARAMIO_OUTPUT) {
			char *pBuffer = ((char *) S->inputBuffer) + P->byteOffset;
			DWORD *pStatus = (DWORD *) pBuffer;

			/* free previous value */
			zval_dtor(param->parameter);
			ZVAL_NULL(param->parameter);

			if (*pStatus == DBSTATUS_S_OK) {
				char *pValue, *s;
				php_stream *stream;
				DBLENGTH *pLength = NULL;
				unsigned int len;
				int conversion;

				if (P->flags & VARIABLE_LENGTH) {
					pLength = (DBLENGTH *) (pBuffer + sizeof(DWORD));
					pValue = pBuffer + sizeof(DWORD) + sizeof(DBLENGTH);
				} else {
					pValue = pBuffer + sizeof(DWORD);
				}

				switch (P->retrievalType) {
					case DBTYPE_STR:
						hr = oledb_convert_string(P->conv, pValue, *pLength, &s, &len, CONVERT_FROM_VARCHAR_TO_OUTPUT);
						ZVAL_STRINGL(param->parameter, s, len, (hr == S_FALSE));
					break;
					case DBTYPE_WSTR:
						hr = oledb_convert_bstr(P->conv, (BSTR) pValue, *pLength / 2, &s, &len, CONVERT_FROM_UNICODE_TO_OUTPUT);
						ZVAL_STRINGL(param->parameter, s, len, FALSE);
					break;
					case DBTYPE_IUNKNOWN:
						/* create PHPSTREAM */
						if (P->flags & STRING_AS_UNICODE) {
							conversion = (P->flags & STRING_AS_UNICODE);
						} else if (!P->flags & STRING_AS_BYTES) {
							conversion = -1;
						} else {
							conversion = CONVERT_FROM_VARCHAR_TO_OUTPUT;
						}
						hr = oledb_create_lob_stream(P->conv, *((IUnknown **) pValue), *pLength, conversion, stmt, &stream TSRMLS_CC);
						if (stream) {
							if (PDO_PARAM_TYPE(param->param_type) == PDO_PARAM_LOB) {
								php_stream_to_zval(stream, param->parameter);
							} else {
								/* copy stream contents into a string */
								len = php_stream_copy_to_mem(stream, &s, PHP_STREAM_COPY_ALL, 0);
								ZVAL_STRINGL(param->parameter, s, len, FALSE);
							}
						}
					break;
					case DBTYPE_I4:
						ZVAL_LONG(param->parameter, *((long *) pValue));
					break;
					case DBTYPE_R8:
						ZVAL_DOUBLE(param->parameter, *((double *) pValue));
					break;
				}
			}
		}
		zend_hash_move_forward(ht);
	}
	return hr;
}

static int oledb_stmt_execute(pdo_stmt_t *stmt TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;
	int ret = 0;
	BSTR sql_w = NULL;
	ICommandText *pICommandText = NULL;

	HRESULT hr;
	DBPARAMS params = { NULL, 0, 0 } ;

	/* release previous rowsets */
	oledb_stmt_clear_rowset(stmt TSRMLS_CC);
	SAFE_RELEASE(S->pIMultipleResults);
	S->pIMultipleResults = NULL;

	/* set command text now if provider doesn't support placeholders */
	if (stmt->supports_placeholders == PDO_PLACEHOLDER_NONE) {
		hr = QUERY_INTERFACE(S->pICommand, IID_ICommandText, pICommandText);
		if (!pICommandText) goto cleanup;

		oledb_create_bstr(S->conv, stmt->active_query_string, -1, &sql_w, NULL, CONVERT_FROM_INPUT_TO_QUERY);

		hr = CALL(SetCommandText, pICommandText, &DBGUID_DEFAULT, sql_w);
		if (!SUCCEEDED(hr)) goto cleanup;
	}

	/* handle bound params */
	if (stmt->bound_params && stmt->bound_params->nNumOfElements > 0) {
		hr = oledb_stmt_copy_bound_params(stmt, &params TSRMLS_CC);
		if (!SUCCEEDED(hr)) goto cleanup;
	}

	/* turn on DBPROP_UNIQUEROWS if we need aditional info */
	if (S->flags & (UNIQUE_ROWS | ADD_TABLE_NAME | ADD_CATALOG_NAME)) {
		ICommandProperties *pICommandProperties = NULL;
		hr = QUERY_INTERFACE(S->pICommand, IID_ICommandProperties, pICommandProperties);
		if (pICommandProperties) {
			DBPROP prop;
			DBPROPSET prop_set;
			prop_set.rgProperties = &prop;
			prop_set.cProperties = 1;
			prop_set.guidPropertySet = DBPROPSET_ROWSET;
			prop.dwOptions = DBPROPOPTIONS_REQUIRED;
			prop.dwPropertyID = DBPROP_UNIQUEROWS;
			prop.colid = DB_NULLID;
			VariantInit(&prop.vValue);
			V_VT(&prop.vValue) = VT_BOOL;
			V_I4(&prop.vValue) = VARIANT_TRUE;

			hr = CALL(SetProperties, pICommandProperties, 1, &prop_set);
			RELEASE(pICommandProperties);
		}
	}

	/* set the cursor type */
	if (S->flags & SERVER_SIDE_CURSOR) {
		ICommandProperties *pICommandProperties = NULL;
		hr = QUERY_INTERFACE(S->pICommand, IID_ICommandProperties, pICommandProperties);
		if (pICommandProperties) {
			DBPROP prop;
			DBPROPSET prop_set;
			prop_set.rgProperties = &prop;
			prop_set.cProperties = 1;
			prop_set.guidPropertySet = DBPROPSET_ROWSET;
			prop.dwOptions = DBPROPOPTIONS_REQUIRED;
			prop.dwPropertyID = DBPROP_SERVERCURSOR;
			prop.colid = DB_NULLID;
			VariantInit(&prop.vValue);
			V_VT(&prop.vValue) = VT_BOOL;
			V_I4(&prop.vValue) = VARIANT_TRUE;

			hr = CALL(SetProperties, pICommandProperties, 1, &prop_set);
			RELEASE(pICommandProperties);
			if (!SUCCEEDED(hr)) goto cleanup;
		}
	}

	/* set the cursor type */
	if (S->flags & SCROLLABLE_CURSOR) {
		ICommandProperties *pICommandProperties = NULL;
		hr = QUERY_INTERFACE(S->pICommand, IID_ICommandProperties, pICommandProperties);
		if (pICommandProperties) {
			DBPROP props[2];
			DBPROPSET prop_set;
			prop_set.rgProperties = props;
			prop_set.cProperties = 2;
			prop_set.guidPropertySet = DBPROPSET_ROWSET;

			props[0].dwOptions = DBPROPOPTIONS_REQUIRED;
			props[0].dwPropertyID = DBPROP_CANSCROLLBACKWARDS;
			props[0].colid = DB_NULLID;
			VariantInit(&props[0].vValue);
			V_VT(&props[0].vValue) = VT_BOOL;
			V_I4(&props[0].vValue) = VARIANT_TRUE;

			props[1].dwOptions = DBPROPOPTIONS_REQUIRED;
			props[1].dwPropertyID = DBPROP_CANFETCHBACKWARDS;
			props[1].colid = DB_NULLID;
			VariantInit(&props[1].vValue);
			V_VT(&props[1].vValue) = VT_BOOL;
			V_I4(&props[1].vValue) = VARIANT_TRUE;

			hr = CALL(SetProperties, pICommandProperties, 1, &prop_set);
			RELEASE(pICommandProperties);
			if (!SUCCEEDED(hr)) goto cleanup;
		}
	}

	if (FALSE && H->flags & MULTIPLE_RESULTS) {
		hr = CALL(Execute, S->pICommand, NULL, &IID_IMultipleResults, &params, &S->rowsAffected, (IUnknown **) &S->pIMultipleResults);
		if(!S->pIMultipleResults) goto cleanup;

		hr = CALL(GetResult, S->pIMultipleResults, NULL, DBRESULTFLAG_ROWSET, &IID_IRowset, &S->rowsAffected, (IUnknown **) &S->pIRowset);
		if (!SUCCEEDED(hr)) goto cleanup;
	} else {
		hr = CALL(Execute, S->pICommand, NULL, &IID_IRowset, &params, &S->rowsAffected, (IUnknown **) &S->pIRowset);
		if (!SUCCEEDED(hr)) goto cleanup;
	}

	/* copy any output values */
	if (stmt->bound_params) {
		hr = oledb_stmt_sync_output_params(stmt TSRMLS_CC);
		if (!SUCCEEDED(hr)) goto cleanup;
	}

	if (S->pIRowset) {
		hr = oledb_stmt_bind_columns(stmt TSRMLS_CC);
	}
	if (SUCCEEDED(hr)) {
		ret = 1;
	}

cleanup:
	SAFE_RELEASE(pICommandText);
	SysFreeString(sql_w);
	if (hr == DB_S_ERRORSOCCURRED) hr = E_FAIL;
	pdo_oledb_error_stmt(stmt, hr);
	return ret;
}

HRESULT oledb_stmt_set_param_driver_option(struct pdo_bound_param_data *param, long attr, zval *val TSRMLS_DC)
{
	pdo_oledb_param *P = (pdo_oledb_param*)param->driver_data;
	HRESULT hr = E_UNEXPECTED;

	switch (attr) {
		default:
			hr = oledb_set_conversion_option(&P->conv, attr, val, FALSE TSRMLS_CC);
			if (hr == S_FALSE) {
				DWORD mask = STRING_AS_UNICODE | STRING_AS_LOB | TRUNCATE_STRING | CONVERT_DATE_TIME;
				hr = oledb_set_internal_flag(attr, val, mask, &P->flags);
			}
	}
	return hr;
}

HRESULT oledb_stmt_set_param_driver_options(struct pdo_bound_param_data *param TSRMLS_DC)
{
	HRESULT hr = S_FALSE;
	if (param->driver_params) {
		HashTable *ht = Z_ARRVAL_P(param->driver_params);
		zval **v;
		char *str_key;
		int key;

		zend_hash_internal_pointer_reset(ht);
		while (SUCCESS == zend_hash_get_current_data(ht, (void**)&v)) {
			if (zend_hash_get_current_key(ht, &str_key, &key, FALSE) == HASH_KEY_IS_LONG) {
				SEPARATE_ZVAL_IF_NOT_REF(v);
				hr = oledb_stmt_set_param_driver_option(param, key, *v TSRMLS_CC);
				if (!SUCCEEDED(hr)) goto cleanup;
			}
			zend_hash_move_forward(ht);
		}
	}
cleanup:
	return hr;
}

HRESULT oledb_stmt_set_column_driver_option(pdo_oledb_column *C, long attr, zval *val TSRMLS_DC)
{
	HRESULT hr = E_UNEXPECTED;

	switch (attr) {
		default:
			hr = oledb_set_conversion_option(&C->conv, attr, val, FALSE TSRMLS_CC);
			if (hr == S_FALSE) {
				DWORD mask = STRING_AS_UNICODE | STRING_AS_LOB | TRUNCATE_STRING | CONVERT_DATE_TIME;
				hr = oledb_set_internal_flag(attr, val, mask, &C->flags);
			}
	}
	return hr;
}

HRESULT oledb_stmt_set_column_driver_options(struct pdo_bound_param_data *param, pdo_oledb_column *C TSRMLS_DC)
{
	HRESULT hr = S_FALSE;
	if (param->driver_params) {
		HashTable *ht = Z_ARRVAL_P(param->driver_params);
		zval **v;
		char *str_key;
		int key;

		zend_hash_internal_pointer_reset(ht);
		while (SUCCESS == zend_hash_get_current_data(ht, (void**)&v)) {
			if (zend_hash_get_current_key(ht, &str_key, &key, FALSE) == HASH_KEY_IS_LONG) {
				SEPARATE_ZVAL_IF_NOT_REF(v);
				hr = oledb_stmt_set_column_driver_option(C, key, *v TSRMLS_CC);
				if (!SUCCEEDED(hr)) goto cleanup;
			}
			zend_hash_move_forward(ht);
		}
	}
cleanup:
	return hr;
}

static int oledb_stmt_param_hook(pdo_stmt_t *stmt, struct pdo_bound_param_data *param,
		enum pdo_param_event event_type TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	HRESULT hr = S_OK;
	int ret = 1;
    
	if (param->is_param) {
		pdo_oledb_param *P = (pdo_oledb_param*)param->driver_data;
		if (!S->pICommandWithParameters || !S->pIAccessorCommand) {
			/* shouldn't happen */
			strcpy(stmt->error_code, "58004");
			return 0;
		}

		switch (event_type) {
			case PDO_PARAM_EVT_ALLOC:
				if (param->paramno < 0 || (S->paramInfo && param->paramno >= (int) S->paramCount)) {
					strcpy(stmt->error_code, "HY093");
					return 0;
				}
				P = (pdo_oledb_param*)ecalloc(1, sizeof(*P));
				P->flags = S->flags;
				oledb_copy_conversion_options(&P->conv, S->conv);
				hr = oledb_stmt_set_param_driver_options(param TSRMLS_CC);
				param->driver_data = P;
				break;
			case PDO_PARAM_EVT_FREE:
				if (P) {
					oledb_stmt_clear_param(stmt, param TSRMLS_CC);
					oledb_release_conversion_options(P->conv);
					efree(P);
				}
				break;
			case PDO_PARAM_EVT_EXEC_PRE:
				oledb_stmt_bind_param(stmt, param TSRMLS_CC);
				break;
			case PDO_PARAM_EVT_EXEC_POST:
				oledb_stmt_clear_param(stmt, param TSRMLS_CC);
				break;
		}
	} else {
		/* a bound column */
		if (param->paramno >= 0 && param->paramno < stmt->column_count) {
			pdo_oledb_column *C = &S->columns[param->paramno];
			switch (event_type) {
				case PDO_PARAM_EVT_EXEC_POST:
					hr = oledb_stmt_set_column_driver_options(param, C TSRMLS_CC);
					break;
			}
		} 
	}
	pdo_oledb_error_stmt(stmt, hr);
	return ret;
}

static int oledb_stmt_fetch(pdo_stmt_t *stmt,
	enum pdo_fetch_orientation ori, long offset TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	int ret = 0;

	HRESULT hr;
	DBCOUNTITEM rows_retrieved = 0;
	DBCOUNTITEM rows_to_fetch;
	DBCOUNTITEM row_offset, new_index;
	HROW *hRows = &S->hRow;

	if (!S->pIRowset) goto cleanup;

	if (*hRows) {
		hr = CALL(ReleaseRows, S->pIRowset, 1, hRows, NULL, NULL, NULL);
	}

	switch(ori) {
		case PDO_FETCH_ORI_FIRST:
			if (S->rowIndex > 1) {
				hr = CALL(RestartPosition, S->pIRowset, DB_NULL_HCHAPTER);
				if (!SUCCEEDED(hr)) goto cleanup;

				S->rowIndex = 0;
			}
			rows_to_fetch = 1;
			row_offset = 0;
			new_index = 1;
			break;
		case PDO_FETCH_ORI_ABS:
			if ((long) S->rowIndex > offset) {
				hr = CALL(RestartPosition, S->pIRowset, DB_NULL_HCHAPTER);
				if (!SUCCEEDED(hr)) goto cleanup;

				S->rowIndex = 0;
				row_offset = offset;
			} else {
				row_offset = offset - S->rowIndex - 1;
			}
			rows_to_fetch = 1;
			new_index = offset + 1;
			break;
		case PDO_FETCH_ORI_PRIOR:
			rows_to_fetch = -1;
			row_offset = offset;
			new_index = S->rowIndex;
			break;
		case PDO_FETCH_ORI_NEXT: 
			rows_to_fetch = 1;
			row_offset = offset;
			new_index = S->rowIndex + 1;
			break;
		case PDO_FETCH_ORI_REL:
			row_offset = offset;
			rows_to_fetch = 1;
			new_index = S->rowIndex + row_offset;
			break;
		case PDO_FETCH_ORI_LAST:
			oledb_set_automation_error(L"Cursor does not support scrolling to the last row.", L"42872");
			hr = E_NOTIMPL;
			goto cleanup;
			break;
	}


	/* get one row */
	hr = CALL(GetNextRows, S->pIRowset, DB_NULL_HCHAPTER, row_offset, rows_to_fetch, &rows_retrieved, &hRows);
	if (!SUCCEEDED(hr)) goto cleanup;

	S->rowIndex = new_index;

	/* get data for non-blob columns if necessary */
	if (rows_retrieved > 0) {
		if (S->hAccessorRowset) {
			hr = CALL(GetData, S->pIRowset, S->hRow, S->hAccessorRowset, S->outputBuffer);
			if (!SUCCEEDED(hr)) goto cleanup;
		}
		ret = 1;
	}

cleanup:
	pdo_oledb_error_stmt(stmt, hr);
	return ret;
}

static int oledb_stmt_describe(pdo_stmt_t *stmt, int colno TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	struct pdo_column_data *cols = stmt->columns;
	int i;

	if (!S->columns) {
		return 0;	
	}

	if (colno >= stmt->column_count) {
		/* error invalid column */
		return 0;
	}

	/* fetch all on demand, this seems easiest 
	** if we've been here before bail out 
	*/
	if (cols[0].name) {
		return 1;
	}
	for (i = 0; i < stmt->column_count; i++) {
		cols[i].precision = S->columns[i].precision;
		cols[i].maxlen = S->columns[i].maxLen;
		cols[i].namelen = S->columns[i].nameLen;
		cols[i].name = S->columns[i].name;
		cols[i].param_type = S->columns[i].pdoType;
	}
	return 1;
}

#define DBTYPE_XML 0x0000FFFF

static char *type_to_name_native(int type, int flags)
{
    switch (type) {
		case DBTYPE_I2: return "smallint";
		case DBTYPE_I4: return "int";
		case DBTYPE_I8: return "bigint";
		case DBTYPE_UI1: return "tinyint";
		case DBTYPE_I1:
		case DBTYPE_NUMERIC: return "numeric";
		case DBTYPE_R4: return "float";
		case DBTYPE_R8: return "real";
		case DBTYPE_DECIMAL: return "decimal";
		case DBTYPE_CY: return "money";
		case DBTYPE_WSTR:
		case DBTYPE_BSTR: 
			if (flags & DBCOLUMNFLAGS_ISLONG) {
				return "ntext";
			} else if(flags & DBCOLUMNFLAGS_ISFIXEDLENGTH) {
				return "nchar";
			} else {
				return "nvarchar";
			}
		case DBTYPE_BOOL: return "bit";
		case DBTYPE_VARIANT: return "variant";
 		case DBTYPE_GUID: return "uniqueidentifier";
		case DBTYPE_BYTES:
			if (flags & DBCOLUMNFLAGS_ISLONG) {
 				return "image";
			} else if(flags & DBCOLUMNFLAGS_ISFIXEDLENGTH) {
				return "binary";
			} else {
				return "varbinary";
			}
		case DBTYPE_STR:
			if (flags & DBCOLUMNFLAGS_ISLONG) {
 				return "text";
			} else if(flags & DBCOLUMNFLAGS_ISFIXEDLENGTH) {
				return "char";
			} else {
				return "varchar";
			}
			break;
		case DBTYPE_DATE:
		case DBTYPE_DBDATE:
		case DBTYPE_DBTIME:
		case DBTYPE_DBTIMESTAMP: return "datetime";
		case DBTYPE_XML: return "xml";
	}
	return NULL;
}

static int oledb_stmt_col_meta(pdo_stmt_t *stmt, long colno, zval *return_value TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_column *C;

	zval *flags;
	char *str;
	
	if (!S->columns) {
		return FAILURE;
	}
	if (colno >= stmt->column_count) {
		/* error invalid column */
		return FAILURE;
	}

	array_init(return_value);
	MAKE_STD_ZVAL(flags);
	array_init(flags);

	C = &S->columns[colno];

	if (!C->metadata) {
		oledb_stmt_get_column_meta_data(stmt TSRMLS_CC);
	}

	if (C->metadata && C->metadata->defaultValue) {
		Z_ADDREF_P(C->metadata->defaultValue);
		add_assoc_zval(return_value, "oledb:def", C->metadata->defaultValue);
	}
	if (!(C->columnFlags & DBCOLUMNFLAGS_ISNULLABLE)) {
		add_next_index_string(flags, "not_null", 1);
	}
	if (C->columnFlags & DBCOLUMNFLAGS_ISLONG) {
		add_next_index_string(flags, "blob", 1);
	}
	str = type_to_name_native(C->columnType, C->columnFlags);
	if (str) {
		add_assoc_string(return_value, "native_type", str, 1);
	}
	if (C->metadata) {
		if (C->metadata->columnName) {
			add_assoc_stringl(return_value, "column", C->metadata->columnName, C->metadata->columnNameLen, 1);
		}
		if (C->metadata->tableName) {
			add_assoc_stringl(return_value, "table", C->metadata->tableName, C->metadata->tableNameLen, 1);
		}
		if (C->metadata->catalogName) {
			add_assoc_stringl(return_value, "dbname", C->metadata->catalogName, C->metadata->catalogNameLen, 1);
		}
		if (C->metadata->isKey) {
			add_next_index_string(flags, "primary_key", 1);
		}
		if (C->metadata->isUnique) {
			add_next_index_string(flags, "unique_key", 1);
		}
		if (C->metadata->isAutoIncrement) {
			add_next_index_string(flags, "auto_increment", 1);
		}
		if (C->metadata->isDynamic) {
			add_next_index_string(flags, "computed", 1);
		}
	}
	add_assoc_zval(return_value, "flags", flags);
	return SUCCESS;

}

char *oledb_datetime_to_str(DBTIMESTAMP *ts)
{
	/* convert to YYYY-MM-DD hh:mm:ss format */
	char buffer[20];
	sprintf(buffer, "%04d-%02d-%02d %02d:%02d:%02d", ts->year, ts->month, ts->day, ts->hour, ts->minute, ts->second + (ts->fraction >= 500000000) ? 1 : 0);
	return estrdup(buffer);
}

int oledb_variant_to_string(pdo_oledb_conversion *conv, VARIANT *v, char **pS, unsigned long *pLen TSRMLS_DC) 
{
	VARIANT sv;
	HRESULT hr;
	VariantInit(&sv);
	hr = VariantChangeType(&sv, v, VARIANT_ALPHABOOL, VT_BSTR);
	if (SUCCEEDED(hr)) {
		oledb_convert_bstr(conv, V_BSTR(v), -1, pS, pLen, CONVERT_FROM_UNICODE_TO_OUTPUT);
		return 1;
	} 
	return 0;
}

static int oledb_stmt_get_col(pdo_stmt_t *stmt, int colno, char **ptr, unsigned long *len, int *caller_frees TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;
	pdo_oledb_column *C;
	int ret = 0;

	HRESULT hr = S_OK;

	if (colno >= stmt->column_count) {
		/* error invalid column */
		return 0;
	}

	C = &S->columns[colno];

	if(C->hAccessorColumn) {
		/* dealing with a blob here... */
		struct {
			DWORD status;
			DBLENGTH length;
			IUnknown *pUnknown;
		} buffer;

		hr = CALL(GetData, S->pIRowset, S->hRow, C->hAccessorColumn, &buffer);
		if (SUCCEEDED(hr)) {
			php_stream *stream = NULL;
			if (buffer.status == DBSTATUS_S_OK) {
				/* see if stream output should be converted from Unicode or another encoding */
				int conversion;
				if (C->columnType == DBTYPE_WSTR) {
					conversion = CONVERT_FROM_UNICODE_TO_OUTPUT;
				} else if (C->columnType == DBTYPE_STR) {
					conversion = CONVERT_FROM_VARCHAR_TO_OUTPUT;
				} else {
					conversion = -1;
				}
				hr = oledb_create_lob_stream(C->conv, buffer.pUnknown, buffer.length, conversion, stmt, &stream TSRMLS_CC);
				if (stream) {
					if (C->pdoType == PDO_PARAM_STR) {
						*len = php_stream_copy_to_mem(stream, ptr, PHP_STREAM_COPY_ALL, 0);
						php_stream_close(stream);
						*caller_frees = 1;
					} else {
						*ptr = (char *) stream;
					}
					ret = 1;
				}
			}
		} else {
			SAFE_RELEASE(buffer.pUnknown);
		}
	} else if(S->outputBuffer) {
		/* set offset into data buffer for column */
		char *p = ((char *) S->outputBuffer) + C->byteOffset;

		/* first 4 bytes contains the column status */
		DWORD *pStatus = (DWORD *) p;
		DBLENGTH *pLength = NULL;
		char *pValue;
		UINT value_len;

		if(*pStatus == DBSTATUS_S_OK || *pStatus == DBSTATUS_S_TRUNCATED) {
			if (C->flags & VARIABLE_LENGTH) {
				/* variable column see the length of the column */
				pLength = (DBLENGTH *) (p + sizeof(DWORD));
				pValue = p + sizeof(DWORD) + sizeof(DBLENGTH);
			}
			else {
				pValue = p + sizeof(DBLENGTH);
				pLength = NULL;
			}

			if (C->retrievalType & DBTYPE_BYREF) {
				/* a pointer is saved there, actually */
				pValue = *((char **) pValue);
			}

			switch (C->retrievalType & ~DBTYPE_BYREF) {
				case DBTYPE_BYTES:
					*ptr = pValue;
					if (C->flags & VARIABLE_LENGTH) {
						*len = *pLength;
					} else {
						*len = C->maxLen;
					}
				break;
				case DBTYPE_STR:
					if (C->columnType == DBTYPE_DBTIMESTAMP || C->columnType == DBTYPE_DATE) {
						/* the length field is bogus */
						value_len = strlen(pValue);
					} else {
						if (C->flags & VARIABLE_LENGTH) {
							value_len = *pLength;
						} else {
							value_len = C->maxLen;
						}
					}
					hr = oledb_convert_string(C->conv, pValue, value_len, ptr, len, CONVERT_FROM_VARCHAR_TO_OUTPUT);
					*caller_frees = (hr != S_FALSE);
				break;
				case DBTYPE_WSTR:
					oledb_convert_bstr(C->conv, (BSTR) pValue, *pLength / 2, ptr, len, CONVERT_FROM_UNICODE_TO_OUTPUT);
					*caller_frees = 1;
				break;
				case DBTYPE_VARIANT: 
					if (oledb_variant_to_string(C->conv, ((VARIANT *) pValue), ptr, len TSRMLS_CC)) {
						*caller_frees = 1;
					}
				break;
				case DBTYPE_DBTIMESTAMP:
					*ptr = oledb_datetime_to_str((DBTIMESTAMP *) pValue);
					*len = strlen(*ptr);
					*caller_frees = 1;
				break;
				case DBTYPE_UI1:
					*ptr = pValue;
					*len = sizeof(zend_bool);
				break;
				case DBTYPE_I4:
					*ptr = pValue;
					*len = sizeof(long);
				break;
			}
		} else {
			/* NULL value */
		}

		ret = 1;
	}

	pdo_oledb_error_stmt(stmt, hr);
	return ret;
}

HRESULT oledb_stmt_set_driver_option(pdo_stmt_t *stmt, long attr, zval *val TSRMLS_DC) {
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	HRESULT hr = E_UNEXPECTED;

	switch (attr) {
		default:
			hr = oledb_set_conversion_option(&S->conv, attr, val, FALSE TSRMLS_CC);
			if (hr == S_FALSE) {
				DWORD mask = UNIQUE_ROWS | SCROLLABLE_CURSOR | SERVER_SIDE_CURSOR | STRING_AS_UNICODE | STRING_AS_LOB | TRUNCATE_STRING | ADD_TABLE_NAME | ADD_CATALOG_NAME | CONVERT_DATE_TIME;
				hr = oledb_set_internal_flag(attr, val, mask, &S->flags);
			}
	}
	return hr;
}

static int oledb_stmt_set_attr(pdo_stmt_t *stmt, long attr, zval *val TSRMLS_DC)
{
	HRESULT hr = oledb_stmt_set_driver_option(stmt, attr, val TSRMLS_CC);
	pdo_oledb_error_stmt(stmt, hr);
	return SUCCEEDED(hr);
}

HRESULT oledb_stmt_set_driver_options(pdo_stmt_t *stmt, zval *driver_options TSRMLS_DC)
{
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
				hr = oledb_stmt_set_driver_option(stmt, key, *v TSRMLS_CC);
				if (!SUCCEEDED(hr)) goto cleanup;
			}
			zend_hash_move_forward(ht);
		}
	}
cleanup:
	return hr;
}

static int oledb_stmt_get_attr(pdo_stmt_t *stmt, long attr, zval *val TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	HRESULT hr = E_UNEXPECTED;

	switch (attr) {
		default:
			hr = oledb_set_conversion_option(&S->conv, attr, val, FALSE TSRMLS_CC);
			if (hr == S_FALSE) {
				hr = oledb_get_internal_flag(attr, S->flags, val);
			}
	}
	pdo_oledb_error_stmt(stmt, hr);
	return SUCCEEDED(hr);
}

static int oledb_stmt_next_rowset(pdo_stmt_t *stmt TSRMLS_DC)
{
	pdo_oledb_stmt *S = (pdo_oledb_stmt*)stmt->driver_data;
	pdo_oledb_db_handle *H = S->H;
	int ret = 0;
	HRESULT hr;

	oledb_stmt_clear_rowset(stmt TSRMLS_CC);
	if (S->pIMultipleResults) {

		/* try to get an IRowset */
		hr = CALL(GetResult, S->pIMultipleResults, NULL, DBRESULTFLAG_ROWSET, &IID_IRowset, &S->rowsAffected, (IUnknown **) &S->pIRowset);
		if (!S->pIRowset) goto cleanup;

		hr = oledb_stmt_bind_columns(stmt TSRMLS_CC);
		if (!SUCCEEDED(hr)) goto cleanup;

		ret = 1;
	} 

cleanup:
	pdo_oledb_error_stmt(stmt, hr);
	return ret;
}

struct pdo_stmt_methods oledb_stmt_methods = {
	oledb_stmt_dtor,
	oledb_stmt_execute,
	oledb_stmt_fetch,
	oledb_stmt_describe,
	oledb_stmt_get_col,
	oledb_stmt_param_hook,
	oledb_stmt_set_attr,
	oledb_stmt_get_attr,
	oledb_stmt_col_meta,
	oledb_stmt_next_rowset
};
