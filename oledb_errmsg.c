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

#include <ole2.h>
#include <oledberr.h>

const char *oledb_hresult_text(HRESULT hr) 
{
	switch(hr) {
		case DB_E_BADACCESSORHANDLE:
			return "Accessor is invalid.";
		case DB_E_ROWLIMITEXCEEDED:
			return "Row could not be inserted into the rowset without exceeding provider's maximum number of active rows.";
		case DB_E_READONLYACCESSOR:
			return "Accessor is read-only. Operation failed.";
		case DB_E_SCHEMAVIOLATION:
			return "Values violate the database schema.";
		case DB_E_BADROWHANDLE:
			return "Row handle is invalid.";
		case DB_E_OBJECTOPEN:
			return "Object was open.";
		case DB_E_BADCHAPTER:
			return "Chapter is invalid.";
		case DB_E_CANTCONVERTVALUE:
			return "Data or literal value could not be converted to the type of the column in the data source, and the provider was unable to determine which columns could not be converted.  Data overflow or sign mismatch was not the cause.";
		case DB_E_BADBINDINFO:
			return "Binding information is invalid.";
		case DB_SEC_E_PERMISSIONDENIED:
			return "Permission denied.";
		case DB_E_NOTAREFERENCECOLUMN:
			return "Column does not contain bookmarks or chapters.";
		case DB_E_LIMITREJECTED:
			return "Cost limits were rejected.";
		case DB_E_NOCOMMAND:
			return "Command text was not set for the command object.";
		case DB_E_COSTLIMIT:
			return "Query plan within the cost limit cannot be found.";
		case DB_E_BADBOOKMARK:
			return "Bookmark is invalid.";
		case DB_E_BADLOCKMODE:
			return "Lock mode is invalid.";
		case DB_E_PARAMNOTOPTIONAL:
			return "No value given for one or more required parameters.";
		case DB_E_BADCOLUMNID:
			return "Column ID is invalid.";
		case DB_E_BADRATIO:
			return "Numerator was greater than denominator. Values must express ratio between zero and 1.";
		case DB_E_BADVALUES:
			return "Value is invalid.";
		case DB_E_ERRORSINCOMMAND:
			return "One or more errors occurred during processing of command.";
		case DB_E_CANTCANCEL:
			return "Command cannot be canceled.";
		case DB_E_DIALECTNOTSUPPORTED:
			return "Command dialect is not supported by this provider.";
		case DB_E_DUPLICATEDATASOURCE:
			return "Data source object could not be created because the named data source already exists.";
		case DB_E_CANNOTRESTART:
			return "Rowset position cannot be restarted.";
		case DB_E_NOTFOUND:
			return "Object or data matching the name, range, or selection criteria was not found within the scope of this operation.";
		case DB_E_NEWLYINSERTED:
			return "Identity cannot be determined for newly inserted rows.";
		case DB_E_CANNOTFREE:
			return "Provider has ownership of this tree.";
		case DB_E_GOALREJECTED:
			return "Goal was rejected because no nonzero weights were specified for any goals supported. Current goal was not changed.";
		case DB_E_UNSUPPORTEDCONVERSION:
			return "Requested conversion is not supported.";
		case DB_E_BADSTARTPOSITION:
			return "No rows were returned because the offset value moves the position before the beginning or after the end of the rowset.";
		case DB_E_NOQUERY:
			return "Information was requested for a query and the query was not set.";
		case DB_E_NOTREENTRANT:
			return "Consumer's event handler called a non-reentrant method in the provider.";
		case DB_E_ERRORSOCCURRED:
			return "Multiple-step OLE DB operation generated errors. Check each OLE DB status value, if available. No work was done.";
		case DB_E_NOAGGREGATION:
			return "Non-NULL controlling IUnknown was specified, and either the requested interface was not IUnknown, or the provider does not support COM aggregation.";
		case DB_E_DELETEDROW:
			return "Row handle referred to a deleted row or a row marked for deletion.";
		case DB_E_CANTFETCHBACKWARDS:
			return "Rowset does not support fetching backward.";
		case DB_E_ROWSNOTRELEASED:
			return "Row handles must all be released before new ones can be obtained.";
		case DB_E_BADSTORAGEFLAG:
			return "One or more storage flags are not supported.";
		case DB_E_BADCOMPAREOP:
			return "Comparison operator is invalid.";
		case DB_E_BADSTATUSVALUE:
			return "Status flag was neither DBCOLUMNSTATUS_OK nor DBCOLUMNSTATUS_ISNULL.";
		case DB_E_CANTSCROLLBACKWARDS:
			return "Rowset does not support scrolling backward.";
		case DB_E_BADREGIONHANDLE:
			return "Region handle is invalid.";
		case DB_E_NONCONTIGUOUSRANGE:
			return "Set of rows is not contiguous to, or does not overlap, the rows in the watch region.";
		case DB_E_INVALIDTRANSITION:
			return "Transition from ALL* to MOVE* or EXTEND* was specified.";
		case DB_E_NOTASUBREGION:
			return "Region is not a proper subregion of the region identified by the watch region handle.";
		case DB_E_MULTIPLESTATEMENTS:
			return "Multiple-statement commands are not supported by this provider.";
		case DB_E_INTEGRITYVIOLATION:
			return "Value violated the integrity constraints for a column or table.";
		case DB_E_BADTYPENAME:
			return "Type name is invalid.";
		case DB_E_ABORTLIMITREACHED:
			return "Execution stopped because a resource limit was reached. No results were returned.";
		case DB_E_ROWSETINCOMMAND:
			return "Command object whose command tree contains a rowset or rowsets cannot be cloned.";
		case DB_E_CANTTRANSLATE:
			return "Current tree cannot be represented as text.";
		case DB_E_DUPLICATEINDEXID:
			return "Index already exists.";
		case DB_E_NOINDEX:
			return "Index does not exist.";
		case DB_E_INDEXINUSE:
			return "Index is in use.";
		case DB_E_NOTABLE:
			return "Table does not exist.";
		case DB_E_CONCURRENCYVIOLATION:
			return "Rowset used optimistic concurrency and the value of a column has changed since it was last read.";
		case DB_E_BADCOPY:
			return "Errors detected during the copy.";
		case DB_E_BADPRECISION:
			return "Precision is invalid.";
		case DB_E_BADSCALE:
			return "Scale is invalid.";
		case DB_E_BADTABLEID:
			return "Table ID is invalid.DB_E_BADID is deprecated; use DB_E_BADTABLEID instead";
		case DB_E_BADTYPE:
			return "Type is invalid.";
		case DB_E_DUPLICATECOLUMNID:
			return "Column ID already exists or occurred more than once in the array of columns.";
		case DB_E_DUPLICATETABLEID:
			return "Table already exists.";
		case DB_E_TABLEINUSE:
			return "Table is in use.";
		case DB_E_NOLOCALE:
			return "Locale ID is not supported.";
		case DB_E_BADRECORDNUM:
			return "Record number is invalid.";
		case DB_E_BOOKMARKSKIPPED:
			return "Form of bookmark is valid, but no row was found to match it.";
		case DB_E_BADPROPERTYVALUE:
			return "Property value is invalid.";
		case DB_E_INVALID:
			return "Rowset is not chaptered.";
		case DB_E_BADACCESSORFLAGS:
			return "One or more accessor flags were invalid.";
		case DB_E_BADSTORAGEFLAGS:
			return "One or more storage flags are invalid.";
		case DB_E_BYREFACCESSORNOTSUPPORTED:
			return "Reference accessors are not supported by this provider.";
		case DB_E_NULLACCESSORNOTSUPPORTED:
			return "Null accessors are not supported by this provider.";
		case DB_E_NOTPREPARED:
			return "Command was not prepared.";
		case DB_E_BADACCESSORTYPE:
			return "Accessor is not a parameter accessor.";
		case DB_E_WRITEONLYACCESSOR:
			return "Accessor is write-only.";
		case DB_SEC_E_AUTH_FAILED:
			return "Authentication failed.";
		case DB_E_CANCELED:
			return "Operation was canceled.";
		case DB_E_CHAPTERNOTRELEASED:
			return "Rowset is single-chaptered. The chapter was not released.";
		case DB_E_BADSOURCEHANDLE:
			return "Source handle is invalid.";
		case DB_E_PARAMUNAVAILABLE:
			return "Provider cannot derive parameter information and SetParameterInfo has not been called.";
		case DB_E_ALREADYINITIALIZED:
			return "Data source object is already initialized.";
		case DB_E_NOTSUPPORTED:
			return "Method is not supported by this provider.";
		case DB_E_MAXPENDCHANGESEXCEEDED:
			return "Number of rows with pending changes exceeded the limit.";
		case DB_E_BADORDINAL:
			return "Column does not exist.";
		case DB_E_PENDINGCHANGES:
			return "Pending changes exist on a row with a reference count of zero.";
		case DB_E_DATAOVERFLOW:
			return "Literal value in the command exceeded the range of the type of the associated column.";
		case DB_E_BADHRESULT:
			return "HRESULT is invalid.";
		case DB_E_BADLOOKUPID:
			return "Lookup ID is invalid.";
		case DB_E_BADDYNAMICERRORID:
			return "DynamicError ID is invalid.";
		case DB_E_PENDINGINSERT:
			return "Most recent data for a newly inserted row could not be retrieved because the insert is pending.";
		case DB_E_BADCONVERTFLAG:
			return "Conversion flag is invalid.";
		case DB_E_BADPARAMETERNAME:
			return "Parameter name is unrecognized.";
		case DB_E_MULTIPLESTORAGE:
			return "Multiple storage objects cannot be open simultaneously.";
		case DB_E_CANTFILTER:
			return "Filter cannot be opened.";
		case DB_E_CANTORDER:
			return "Order cannot be opened.";
		case MD_E_BADTUPLE:
			return "Tuple is invalid.";
		case MD_E_BADCOORDINATE:
			return "Coordinate is invalid.";
		case MD_E_INVALIDAXIS:
			return "Axis is invalid.";
		case MD_E_INVALIDCELLRANGE:
			return "One or more cell ordinals is invalid.";
		case DB_E_NOCOLUMN:
			return "Column ID is invalid.";
		case DB_E_COMMANDNOTPERSISTED:
			return "Command does not have a DBID.";
		case DB_E_DUPLICATEID:
			return "DBID already exists.";
		case DB_E_OBJECTCREATIONLIMITREACHED:
			return "Session cannot be created because maximum number of active sessions was already reached. Consumer must release one or more sessions before creating a new session object. ";
		case DB_E_BADINDEXID:
			return "Index ID is invalid.";
		case DB_E_BADINITSTRING:
			return "Format of the initialization string does not conform to the OLE DB specification.";
		case DB_E_NOPROVIDERSREGISTERED:
			return "No OLE DB providers of this source type are registered.";
		case DB_E_MISMATCHEDPROVIDER:
			return "Initialization string specifies a provider that does not match the active provider.";
		case DB_E_BADCOMMANDID:
			return "DBID is invalid.";
		case SEC_E_BADTRUSTEEID:
			return "Trustee is invalid.";
		case SEC_E_NOTRUSTEEID:
			return "Trustee was not recognized for this data source.";
		case SEC_E_NOMEMBERSHIPSUPPORT:
			return "Trustee does not support memberships or collections.";
		case SEC_E_INVALIDOBJECT:
			return "Object is invalid or unknown to the provider.";
		case SEC_E_NOOWNER:
			return "Object does not have an owner.";
		case SEC_E_INVALIDACCESSENTRYLIST:
			return "Access entry list is invalid.";
		case SEC_E_INVALIDOWNER:
			return "Trustee supplied as owner is invalid or unknown to the provider.";
		case SEC_E_INVALIDACCESSENTRY:
			return "Permission in the access entry list is invalid.";
		case DB_E_BADCONSTRAINTTYPE:
			return "ConstraintType is invalid or not supported by the provider.";
		case DB_E_BADCONSTRAINTFORM:
			return "ConstraintType is not DBCONSTRAINTTYPE_FOREIGNKEY and cForeignKeyColumns is not zero.";
		case DB_E_BADDEFERRABILITY:
			return "Specified deferrability flag is invalid or not supported by the provider.";
		case DB_E_BADMATCHTYPE:
			return "MatchType is invalid or the value is not supported by the provider.";
		case DB_E_BADUPDATEDELETERULE:
			return "Constraint update rule or delete rule is invalid.";
		case DB_E_BADCONSTRAINTID:
			return "Constraint ID is invalid.";
		case DB_E_BADCOMMANDFLAGS:
			return "Command persistence flag is invalid.";
		case DB_E_OBJECTMISMATCH:
			return "rguidColumnType points to a GUID that does not match the object type of this column, or this column was not set.";
		case DB_E_NOSOURCEOBJECT:
			return "Source row does not exist.";
		case DB_E_RESOURCELOCKED:
			return "OLE DB object represented by this URL is locked by one or more other processes.";
		case DB_E_NOTCOLLECTION:
			return "Client requested an object type that is valid only for a collection. ";
		case DB_E_READONLY:
			return "Caller requested write access to a read-only object.";
		case DB_E_ASYNCNOTSUPPORTED:
			return "Asynchronous binding is not supported by this provider.";
		case DB_E_CANNOTCONNECT:
			return "Connection to the server for this URL cannot be established.";
		case DB_E_TIMEOUT:
			return "Timeout occurred when attempting to bind to the object.";
		case DB_E_RESOURCEEXISTS:
			return "Object cannot be created at this URL because an object named by this URL already exists.";
		case DB_E_RESOURCEOUTOFSCOPE:
			return "URL is outside of scope.";
		case DB_E_DROPRESTRICTED:
			return "Column or constraint could not be dropped because it is referenced by a dependent view or constraint.";
		case DB_E_DUPLICATECONSTRAINTID:
			return "Constraint already exists.";
		case DB_E_OUTOFSPACE:
			return "Object cannot be created at this URL because the server is out of physical storage.";
		case DB_SEC_E_SAFEMODE_DENIED:
			return "Safety settings on this computer prohibit accessing a data source on another domain.";
		case DB_E_NOSTATISTIC:
			return "The specified statistic does not exist in the current data source or did not apply to the specified table or it does not support a histogram. ";
		case DB_E_ALTERRESTRICTED:
			return "Column or table could not be altered because it is referenced by a constraint.";
		case DB_E_RESOURCENOTSUPPORTED:
			return "Requested object type is not supported by the provider.";
		case DB_E_NOCONSTRAINT:
			return "Constraint does not exist.";
		case DB_E_COLUMNUNAVAILABLE:
			return "Requested column is valid, but could not be retrieved. This could be due to a forward only cursor attempting to go backwards in a row.";
	}
	return NULL;
}