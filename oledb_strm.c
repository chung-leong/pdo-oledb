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

typedef struct {
	IStream *pIStream;
	ISequentialStream *pISequentialStream;
	IMLangConvertCharset *pIMLangConvertCharset;
	pdo_stmt_t *stmt;
	DBLENGTH length;
	char *bytes;
	int byte_count;
	int offset;
} oledb_lob_this;

static void oledb_blob_close_stream_resources(oledb_lob_this *this TSRMLS_DC) 
{
	SAFE_RELEASE(this->pIMLangConvertCharset);
	SAFE_RELEASE(this->pISequentialStream);
	SAFE_RELEASE(this->pIStream);
	if (this->bytes) {
		efree(this->bytes);
	}
	if (this->stmt) {
		_php_pdo_stmt_delref(this->stmt TSRMLS_CC);
	}
	ZeroMemory(this, sizeof(*this));
}

static size_t oledb_blob_write(php_stream *stream, const char *buf, size_t count TSRMLS_DC)
{
	return (size_t) - 1;
}

static size_t oledb_blob_read(php_stream *stream, char *buf, size_t count TSRMLS_DC)
{
	oledb_lob_this *this = (oledb_lob_this*)stream->abstract;

	HRESULT hr = E_UNEXPECTED;
	ULONG read = 0;

	if (this->pISequentialStream) {
		if (this->pIMLangConvertCharset) {
			unsigned remaining = count, total_len = 0;
			unsigned len_dest = remaining;
			unsigned len_src = this->byte_count - this->offset;
			BYTE *src = this->bytes + this->offset;
			BYTE *dest = (BYTE *) buf;
			if (len_src) {
				/* convert leftover stuff */
				hr = CALL(DoConversion, this->pIMLangConvertCharset, src, &len_src, dest, &len_dest);
				this->offset += len_src;
				remaining -= len_dest;
				total_len += len_dest;
				dest += len_dest;
			}
			src = this->bytes;
			while (remaining > 0 && this->byte_count > 0) {
				hr = CALL(Read, this->pISequentialStream, this->bytes, this->byte_count, &read);
				if (SUCCEEDED(hr)) {
					this->byte_count = read;
					if (read > 0) {
						len_src = read;
						len_dest = remaining;
						hr = CALL(DoConversion, this->pIMLangConvertCharset, src, &len_src, dest, &len_dest);
						remaining -= len_dest;
						total_len += len_dest;
						dest += len_dest;
					}
					this->offset = len_src;
				} else {
					break;
				}
			}
			read = total_len;
		} else {
			hr = CALL(Read, this->pISequentialStream, buf, count, &read);
			if (hr == S_FALSE && !this->pIStream) {
				/* We've reached the end of the stream and there's no way to seek back.
				   Release the resources now, in case the variable is just waiting for 
				   garbage collection. Some provider allows only one stream to be opened
				   at one time.
				*/
				oledb_blob_close_stream_resources(this TSRMLS_CC);
			}
		}
	}

	return SUCCEEDED(hr) ? read : (size_t) -1;
}

static int oledb_blob_close(php_stream *stream, int close_handle TSRMLS_DC)
{
	oledb_lob_this *this = (oledb_lob_this*)stream->abstract;

	if (close_handle) {
		oledb_blob_close_stream_resources(this TSRMLS_CC);
		efree(this);
	}

	return 0;
}

static int oledb_blob_flush(php_stream *stream TSRMLS_DC)
{
	/* do nothing */
	return 0;
}

static int oledb_blob_seek(php_stream *stream, off_t offset, int whence, off_t *newoffset TSRMLS_DC)
{
	oledb_lob_this *this = (oledb_lob_this*)stream->abstract;
    	
	LARGE_INTEGER large_offset;
	HRESULT hr = E_UNEXPECTED;

	large_offset.HighPart = 0;
	large_offset.LowPart = offset;

	if (this->pIStream) {
		/* try to seek if the object implements IStream */
		hr = CALL(Seek, this->pIStream, large_offset, whence, NULL);
	}

	return SUCCEEDED(hr) ? 0 : -1;
}

static int oledb_blob_stat(php_stream *stream, php_stream_statbuf *ssb TSRMLS_DC)
{
	oledb_lob_this *this = (oledb_lob_this*)stream->abstract;

	if (!this->pIMLangConvertCharset && this->length) {
		ZeroMemory(ssb, sizeof(*ssb));
		ssb->sb.st_size = this->length;
		return 0;
	} else {
		return -1;
	}
}

static php_stream_ops oledb_blob_stream_ops = {
	oledb_blob_write,
	oledb_blob_read,
	oledb_blob_close,
	oledb_blob_flush,
	"pdo_oledb blob stream",
	oledb_blob_seek,
	NULL,
	oledb_blob_stat,
	NULL
};

HRESULT oledb_create_lob_stream(pdo_oledb_conversion *conv, IUnknown *pUnk, DBLENGTH length, int conversion, pdo_stmt_t *stmt, php_stream **pStream TSRMLS_DC)
{
	oledb_lob_this *this = ecalloc(1, sizeof(*this));

	HRESULT hr;

	/* try to get an IStream interface (failure most likely) */
	hr = QUERY_INTERFACE(pUnk, IID_IStream, this->pIStream);
	if (this->pIStream) {
		/* downcast it to ISequentialStream */
		this->pISequentialStream = (ISequentialStream *) this->pIStream;
		ADDREF(this->pISequentialStream);
	} else {
		/* get an ISequentialStream instead */
		hr = QUERY_INTERFACE(pUnk, IID_ISequentialStream, this->pISequentialStream);
	}
	/* don't need IUnknown anymore as we're holding something else */
	RELEASE(pUnk);

	this->length = length;
	if (conversion != -1) {
		this->pIMLangConvertCharset = conv->pIMLangConvertCharsets[conversion]; 
	}
	if (this->pIMLangConvertCharset) {
		ADDREF(this->pIMLangConvertCharset);
		this->byte_count = 1024;
		this->bytes = emalloc(this->byte_count);
		this->offset = this->byte_count;
	}
	this->stmt = stmt;
	_php_pdo_stmt_addref(this->stmt TSRMLS_CC);

	*pStream = php_stream_alloc(&oledb_blob_stream_ops, this, 0, "r+b");

	if (!*pStream) {
		hr = E_FAIL;
		oledb_blob_close_stream_resources(this TSRMLS_CC);
		efree(this);
	}
	return hr;
}

typedef struct {
	ISequentialStreamVtbl *lpVtbl;
	char *bytes;
	unsigned int byte_count;
	unsigned int offset;
	php_stream *stream;
	int own_string;
	ULONG refcount;
	void *tsrm_ls;
} zval_stream;

#define DECLARE_THIS(vtbl, offset)		zval_stream *this = ((zval_stream *) &((IUnknown *) vtbl)[- offset])

static HRESULT STDMETHODCALLTYPE zval_stream_QueryInterface(ISequentialStream *ptr,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject)
{
	DECLARE_THIS(ptr, 0);
	if (IsEqualIID(riid, &IID_IUnknown) || IsEqualIID(riid, &IID_ISequentialStream)) {
		*ppvObject = &this->lpVtbl;
	} else {
		*ppvObject = NULL;
		return E_NOINTERFACE;
	}
	ADDREF((IUnknown *) *ppvObject);
	return S_OK;
}

static ULONG STDMETHODCALLTYPE zval_stream_AddRef(ISequentialStream *ptr)
{
	DECLARE_THIS(ptr, 0);
	return ++(this->refcount);
}

static ULONG STDMETHODCALLTYPE zval_stream_Release(ISequentialStream *ptr)
{
	DECLARE_THIS(ptr, 0);
	if(--(this->refcount) == 0) {
		if (this->own_string) {
			SAFE_EFREE(this->bytes);
		}
		CoTaskMemFree(this);
		return 0;
	}
	return this->refcount;
}

static HRESULT STDMETHODCALLTYPE zval_stream_Read(ISequentialStream *ptr,
			void *pv,
			ULONG cb,
			ULONG *pcbRead)
{
	DECLARE_THIS(ptr, 0);

	if (this->stream) {
		void *tsrm_ls = this->tsrm_ls;
		unsigned int len;
		len = php_stream_read(this->stream, (char *) pv, cb);
		*pcbRead = len;
		return (len > 0) ? S_OK : S_FALSE;
	} else if (this->bytes) {
		unsigned int len = min(cb, this->byte_count - this->offset);
		if (!len) {
			return S_FALSE;
		}
		memcpy(pv, this->bytes + this->offset, len);
		this->offset += len;
		*pcbRead = len;
		return S_OK;
	}
	return E_FAIL;
}

static HRESULT STDMETHODCALLTYPE zval_stream_Write(ISequentialStream *ptr,
			const void *pv,
			ULONG cb,
			ULONG *pcbWritten)
{
	DECLARE_THIS(ptr, 0);
	return E_NOTIMPL;
}

ISequentialStreamVtbl zval_stream_Vtbl = {
	zval_stream_QueryInterface,
	zval_stream_AddRef,
	zval_stream_Release,
	zval_stream_Read,
	zval_stream_Write
};

HRESULT oledb_create_zval_stream(pdo_oledb_conversion *conv, zval *value, int conversion, IUnknown **pUnk, DBLENGTH *pLength TSRMLS_DC)
{
	zval_stream *this;

	char *bytes = NULL;
	unsigned int byte_count = 0;
	php_stream *stream = NULL;
	php_stream_statbuf ssbuf;
	int own_string = 0;

	if (Z_TYPE_P(value) == IS_STRING) {
		bytes = Z_STRVAL_P(value);
		byte_count = Z_STRLEN_P(value);
	} else if (Z_TYPE_P(value) == IS_RESOURCE) {
		php_stream_from_zval_no_verify(stream, &value);

		if (!stream) {
			oledb_set_automation_error(L"Expected a stream resource", L"HY105");
			return ERROR_INVALID_PARAMETER;
		}
		
		/* we need to read the entire contents of the stream to determine the final length if:
			1. the contents is to be converted to Unicode or some other encoding
			2. the stream has filters attached (which might after the content length)
			3. the stream doesn't support stat
		*/
		if ((php_stream_stat(stream, &ssbuf) != 0 || !ssbuf.sb.st_size) || conversion != -1 || php_stream_is_filtered(stream)) {
			byte_count = php_stream_copy_to_mem(stream, &bytes, PHP_STREAM_COPY_ALL, 0);
			own_string = 1;
			stream = NULL;
		}
	} else {
		zval copy;
		copy = *value;
		zval_copy_ctor(&copy);
		convert_to_string(&copy);
		bytes = Z_STRVAL_P(value);
		byte_count = Z_STRLEN_P(value);
		own_string = 1;
	}

	if (conversion != -1) {
		/* convert bytes to Unicode or another encoding scheme */
		HRESULT hr;
		unsigned int len = byte_count;
		unsigned char *src = (unsigned char *) bytes;

		hr = oledb_convert_string(conv, src, len, &bytes, &byte_count, conversion);
		if (hr == S_FALSE) {
			if (own_string) {
				efree(src);
			}
		}
	}

	this = CoTaskMemAlloc(sizeof(*this));
	ZeroMemory(this, sizeof(*this));
	this->lpVtbl = &zval_stream_Vtbl;
#ifdef ZTS
	this->tsrm_ls = tsrm_ls;
#else
	this->tsrm_ls = NULL;
#endif
	this->stream = stream;
	this->bytes = bytes;
	this->byte_count = byte_count;
	this->own_string = own_string;
	
	if (stream) {
		*pLength = ssbuf.sb.st_size;
	} else {
		*pLength = this->byte_count;
	}
	return QUERY_INTERFACE((IUnknown *) this, IID_IUnknown, *pUnk);
}
