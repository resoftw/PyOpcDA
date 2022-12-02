#include "Utils.h"
#include <iostream>
#include <string>
#include "opc/opcerror.h"

time_t  filetime_to_timet(const FILETIME& ft) { ULARGE_INTEGER ull;    ull.LowPart = ft.dwLowDateTime;    ull.HighPart = ft.dwHighDateTime;    return ull.QuadPart / 10000000ULL - 11644473600ULL; }

size_t GetProcessMemory()
{
	PROCESS_MEMORY_COUNTERS info;
	GetProcessMemoryInfo(GetCurrentProcess(), &info, sizeof(info));
	return (size_t)info.WorkingSetSize;
}

std::string VarTypeStr(ULONG vt)
{
	std::string r = "UNKNOWN";
	switch (vt)
	{
	case VT_EMPTY:
		r = "VT_EMPTY";
		break;
	case VT_I1:
		r = "VT_I1";
		break;
	case VT_I2:
		r = "VT_I2";
		break;
	case VT_I4:
		r = "VT_I4";
		break;
	case VT_R4:
		r = "VT_R4";
		break;
	case VT_R8:
		r = "VT_R8";
		break;
	case VT_CY:
		r = "VT_CY";
		break;
	case VT_DATE:
		r = "VT_DATE";
		break;
	case VT_BSTR:
		r = "VT_BSTR";
		break;
	case VT_DISPATCH:
		r = "VT_DISPATCH";
		break;
	case VT_BOOL:
		r = "VT_BOOL";
		break;
	case VT_UNKNOWN:
		r = "VT_UNKNOWN";
		break;
	case VT_UI1:
		r = "VT_UI1";
		break;
	case VT_UI2:
		r = "VT_UI2";
		break;
	case VT_UI4:
		r = "VT_UI4";
		break;
	case VT_LPSTR:
		r = "VT_LPSTR";
		break;
	case VT_LPWSTR:
		r = "VT_LPWSTR";
		break;
	case VT_STREAMED_OBJECT:
		r = "VT_STREAMED_OBJECT";
		break;
	case VT_STORED_OBJECT:
		r = "VT_STORED_OBJECT";
		break;
	case VT_VECTOR:
		r = "VT_VECTOR";
		break;
	case VT_ARRAY:
		r = "VT_ARRAY";
		break;
	case VT_ARRAY | VT_UI1:
		r = "VT_ARRAY|VT_UI1";
		break;
	case VT_VECTOR | VT_UI1:
		r = "VT_VECTOR|VT_UI1";
		break;
	}
	return r;
}
std::string Var2Str(const VARIANT& _In_ v)
{
	std::string r = "";
	std::wstring ws = L"";
	std::string tmp = "";
	switch (v.vt)
	{
		case VT_EMPTY:
			r = "";
			break;
		case VT_I1:
			r = std::to_string(v.cVal);
			break;
		case VT_I2:
			r = std::to_string(v.iVal);
			break;
		case VT_INT:
			r = std::to_string(v.intVal);
			break;
		case VT_I4:
			r = std::to_string(v.lVal);
			break;
		case VT_R4:
			r = std::to_string(v.fltVal);
			break;
		case VT_R8:
			r = std::to_string(v.dblVal);
			break;
		case VT_CY:
			r = std::to_string(v.llVal);
			break;
		case VT_DATE:
			SYSTEMTIME tm;
			VariantTimeToSystemTime(v.date, &tm);
			char dtm[20];
			sprintf_s(dtm, "%04d-%02d-%02d %02d:%02d:%02d", tm.wYear, tm.wMonth, tm.wDay, tm.wHour, tm.wMinute, tm.wSecond);
			r = dtm;
			break;
		case VT_BSTR:
			ws = v.bstrVal;
			r = std::string(ws.begin(), ws.end());
			break;
		case VT_BOOL:
			r = v.boolVal?"True":"False";
			break;
		case VT_UI1:
			r = std::to_string(v.bVal);
			break;
		case VT_UI2:
			r = std::to_string(v.uiVal);
			break;
		case VT_UINT:
			r = std::to_string(v.uintVal);
			break;
		case VT_UI4:
			r = std::to_string(v.ulVal);
			break;
		default:
			r = "UNK: " + VarTypeStr(v.vt);
			break;
	}
	return r;
}

wchar_t* convertMBSToWCS(char const* value) {
	size_t newSize = strlen(value) + 1;
	size_t convertedChars = 0;

	wchar_t* converted = new wchar_t[newSize];

	mbstowcs_s(&convertedChars, converted, newSize, value, _TRUNCATE);

	return converted;
}
std::string DBTime(time_t t)
{
	if (t < 0)return "1970-01-01 00:00:00";
	tm lt;
	char buf[50];
	localtime_s(&lt, &t);
	strftime(buf, 50, "%F %T", &lt);
	return std::string(buf);
}

std::string strTime(time_t t)
{
	if (t < 0)return "#n/a";
	tm lt;
	if (t == 0)t = time(0);
	char buf[50];
	localtime_s(&lt, &t);
	strftime(buf, 50, "%F %T", &lt);
	return std::string(buf);
}

std::string strTime(std::string const& fmt,time_t t)
{
	if (t < 0)return "#n/a";
	tm lt;
	if (t == 0)t = time(0);
	char buf[256];
	localtime_s(&lt, &t);
	strftime(buf, 256, fmt.c_str(), &lt);
	return std::string(buf);
}

std::string hResultStr(HRESULT h)
{
	switch (h)
	{
		case S_OK:
			return "S_OK";
		case S_FALSE:
			return "S_FALSE";
		case E_UNEXPECTED:
			return "E_UNEXPECTED";
		case E_NOTIMPL:
			return "E_NOTIMPL";
		case E_OUTOFMEMORY:
			return "E_OUTOFMEMORY";
		case E_INVALIDARG:
			return "E_INVALIDARG";
		case E_NOINTERFACE:
			return "E_NOINTERFACE";
		case E_POINTER:
			return "E_POINTER";
		case E_HANDLE:
			return "E_HANDLE";
		case E_ABORT:
			return "E_ABORT";
		case E_FAIL:
			return "E_FAIL";
		case E_ACCESSDENIED:
			return "E_ACCESSDENIED";
		case E_PENDING:
			return "E_PENDING";
		case E_BOUNDS:
			return "E_BOUNDS";
		case E_CHANGED_STATE:
			return "E_CHANGED_STATE";
		case E_ILLEGAL_STATE_CHANGE:
			return "E_ILLEGAL_STATE_CHANGE";
		case E_ILLEGAL_METHOD_CALL:
			return "E_ILLEGAL_METHOD_CALL";
		case RO_E_METADATA_NAME_NOT_FOUND:
			return "RO_E_METADATA_NAME_NOT_FOUND";
		case RO_E_METADATA_NAME_IS_NAMESPACE:
			return "RO_E_METADATA_NAME_IS_NAMESPACE";
		case RO_E_METADATA_INVALID_TYPE_FORMAT:
			return "RO_E_METADATA_INVALID_TYPE_FORMAT";
		case RO_E_INVALID_METADATA_FILE:
			return "RO_E_INVALID_METADATA_FILE";
		case RO_E_CLOSED:
			return "RO_E_CLOSED";
		case RO_E_EXCLUSIVE_WRITE:
			return "RO_E_EXCLUSIVE_WRITE";
		case RO_E_CHANGE_NOTIFICATION_IN_PROGRESS:
			return "RO_E_CHANGE_NOTIFICATION_IN_PROGRESS";
		case RO_E_ERROR_STRING_NOT_FOUND:
			return "RO_E_ERROR_STRING_NOT_FOUND";
		case E_STRING_NOT_NULL_TERMINATED:
			return "E_STRING_NOT_NULL_TERMINATED";
		case E_ILLEGAL_DELEGATE_ASSIGNMENT:
			return "E_ILLEGAL_DELEGATE_ASSIGNMENT";
		case E_ASYNC_OPERATION_NOT_STARTED:
			return "E_ASYNC_OPERATION_NOT_STARTED";
		case E_APPLICATION_EXITING:
			return "E_APPLICATION_EXITING";
		case E_APPLICATION_VIEW_EXITING:
			return "E_APPLICATION_VIEW_EXITING";
		case RO_E_MUST_BE_AGILE:
			return "RO_E_MUST_BE_AGILE";
		case RO_E_UNSUPPORTED_FROM_MTA:
			return "RO_E_UNSUPPORTED_FROM_MTA";
		case RO_E_COMMITTED:
			return "RO_E_COMMITTED";
		case RO_E_BLOCKED_CROSS_ASTA_CALL:
			return "RO_E_BLOCKED_CROSS_ASTA_CALL";
		case RO_E_CANNOT_ACTIVATE_FULL_TRUST_SERVER:
			return "RO_E_CANNOT_ACTIVATE_FULL_TRUST_SERVER";
		case RO_E_CANNOT_ACTIVATE_UNIVERSAL_APPLICATION_SERVER:
			return "RO_E_CANNOT_ACTIVATE_UNIVERSAL_APPLICATION_SERVER";
		case CO_E_INIT_TLS:
			return "CO_E_INIT_TLS";
		case CO_E_INIT_SHARED_ALLOCATOR:
			return "CO_E_INIT_SHARED_ALLOCATOR";
		case CO_E_INIT_MEMORY_ALLOCATOR:
			return "CO_E_INIT_MEMORY_ALLOCATOR";
		case CO_E_INIT_CLASS_CACHE:
			return "CO_E_INIT_CLASS_CACHE";
		case CO_E_INIT_RPC_CHANNEL:
			return "CO_E_INIT_RPC_CHANNEL";
		case CO_E_INIT_TLS_SET_CHANNEL_CONTROL:
			return "CO_E_INIT_TLS_SET_CHANNEL_CONTROL";
		case CO_E_INIT_TLS_CHANNEL_CONTROL:
			return "CO_E_INIT_TLS_CHANNEL_CONTROL";
		case CO_E_INIT_UNACCEPTED_USER_ALLOCATOR:
			return "CO_E_INIT_UNACCEPTED_USER_ALLOCATOR";
		case CO_E_INIT_SCM_MUTEX_EXISTS:
			return "CO_E_INIT_SCM_MUTEX_EXISTS";
		case CO_E_INIT_SCM_FILE_MAPPING_EXISTS:
			return "CO_E_INIT_SCM_FILE_MAPPING_EXISTS";
		case CO_E_INIT_SCM_MAP_VIEW_OF_FILE:
			return "CO_E_INIT_SCM_MAP_VIEW_OF_FILE";
		case CO_E_INIT_SCM_EXEC_FAILURE:
			return "CO_E_INIT_SCM_EXEC_FAILURE";
		case CO_E_INIT_ONLY_SINGLE_THREADED:
			return "CO_E_INIT_ONLY_SINGLE_THREADED";
		case CO_E_CANT_REMOTE:
			return "CO_E_CANT_REMOTE";
		case CO_E_BAD_SERVER_NAME:
			return "CO_E_BAD_SERVER_NAME";
		case CO_E_WRONG_SERVER_IDENTITY:
			return "CO_E_WRONG_SERVER_IDENTITY";
		case CO_E_OLE1DDE_DISABLED:
			return "CO_E_OLE1DDE_DISABLED";
		case CO_E_RUNAS_SYNTAX:
			return "CO_E_RUNAS_SYNTAX";
		case CO_E_CREATEPROCESS_FAILURE:
			return "CO_E_CREATEPROCESS_FAILURE";
		case CO_E_RUNAS_CREATEPROCESS_FAILURE:
			return "CO_E_RUNAS_CREATEPROCESS_FAILURE";
		case CO_E_RUNAS_LOGON_FAILURE:
			return "CO_E_RUNAS_LOGON_FAILURE";
		case CO_E_LAUNCH_PERMSSION_DENIED:
			return "CO_E_LAUNCH_PERMSSION_DENIED";
		case CO_E_START_SERVICE_FAILURE:
			return "CO_E_START_SERVICE_FAILURE";
		case CO_E_REMOTE_COMMUNICATION_FAILURE:
			return "CO_E_REMOTE_COMMUNICATION_FAILURE";
		case CO_E_SERVER_START_TIMEOUT:
			return "CO_E_SERVER_START_TIMEOUT";
		case CO_E_CLSREG_INCONSISTENT:
			return "CO_E_CLSREG_INCONSISTENT";
		case CO_E_IIDREG_INCONSISTENT:
			return "CO_E_IIDREG_INCONSISTENT";
		case CO_E_NOT_SUPPORTED:
			return "CO_E_NOT_SUPPORTED";
		case CO_E_RELOAD_DLL:
			return "CO_E_RELOAD_DLL";
		case CO_E_MSI_ERROR:
			return "CO_E_MSI_ERROR";
		case CO_E_ATTEMPT_TO_CREATE_OUTSIDE_CLIENT_CONTEXT:
			return "CO_E_ATTEMPT_TO_CREATE_OUTSIDE_CLIENT_CONTEXT";
		case CO_E_SERVER_PAUSED:
			return "CO_E_SERVER_PAUSED";
		case CO_E_SERVER_NOT_PAUSED:
			return "CO_E_SERVER_NOT_PAUSED";
		case CO_E_CLASS_DISABLED:
			return "CO_E_CLASS_DISABLED";
		case CO_E_CLRNOTAVAILABLE:
			return "CO_E_CLRNOTAVAILABLE";
		case CO_E_ASYNC_WORK_REJECTED:
			return "CO_E_ASYNC_WORK_REJECTED";
		case CO_E_SERVER_INIT_TIMEOUT:
			return "CO_E_SERVER_INIT_TIMEOUT";
		case CO_E_NO_SECCTX_IN_ACTIVATE:
			return "CO_E_NO_SECCTX_IN_ACTIVATE";
		case CO_E_TRACKER_CONFIG:
			return "CO_E_TRACKER_CONFIG";
		case CO_E_THREADPOOL_CONFIG:
			return "CO_E_THREADPOOL_CONFIG";
		case CO_E_SXS_CONFIG:
			return "CO_E_SXS_CONFIG";
		case CO_E_MALFORMED_SPN:
			return "CO_E_MALFORMED_SPN";
		case CO_E_UNREVOKED_REGISTRATION_ON_APARTMENT_SHUTDOWN:
			return "CO_E_UNREVOKED_REGISTRATION_ON_APARTMENT_SHUTDOWN";
		case CO_E_PREMATURE_STUB_RUNDOWN:
			return "CO_E_PREMATURE_STUB_RUNDOWN";
		case OLE_E_OLEVERB:
			return "OLE_E_OLEVERB";
		case OLE_E_ADVF:
			return "OLE_E_ADVF";
		case OLE_E_ENUM_NOMORE:
			return "OLE_E_ENUM_NOMORE";
		case OLE_E_ADVISENOTSUPPORTED:
			return "OLE_E_ADVISENOTSUPPORTED";
		case OLE_E_NOCONNECTION:
			return "OLE_E_NOCONNECTION";
		case OLE_E_NOTRUNNING:
			return "OLE_E_NOTRUNNING";
		case OLE_E_NOCACHE:
			return "OLE_E_NOCACHE";
		case OLE_E_BLANK:
			return "OLE_E_BLANK";
		case OLE_E_CLASSDIFF:
			return "OLE_E_CLASSDIFF";
		case OLE_E_CANT_GETMONIKER:
			return "OLE_E_CANT_GETMONIKER";
		case OLE_E_CANT_BINDTOSOURCE:
			return "OLE_E_CANT_BINDTOSOURCE";
		case OLE_E_STATIC:
			return "OLE_E_STATIC";
		case OLE_E_PROMPTSAVECANCELLED:
			return "OLE_E_PROMPTSAVECANCELLED";
		case OLE_E_INVALIDRECT:
			return "OLE_E_INVALIDRECT";
		case OLE_E_WRONGCOMPOBJ:
			return "OLE_E_WRONGCOMPOBJ";
		case OLE_E_INVALIDHWND:
			return "OLE_E_INVALIDHWND";
		case OLE_E_NOT_INPLACEACTIVE:
			return "OLE_E_NOT_INPLACEACTIVE";
		case OLE_E_CANTCONVERT:
			return "OLE_E_CANTCONVERT";
		case OLE_E_NOSTORAGE:
			return "OLE_E_NOSTORAGE";
		case DV_E_FORMATETC:
			return "DV_E_FORMATETC";
		case DV_E_DVTARGETDEVICE:
			return "DV_E_DVTARGETDEVICE";
		case DV_E_STGMEDIUM:
			return "DV_E_STGMEDIUM";
		case DV_E_STATDATA:
			return "DV_E_STATDATA";
		case DV_E_LINDEX:
			return "DV_E_LINDEX";
		case DV_E_TYMED:
			return "DV_E_TYMED";
		case DV_E_CLIPFORMAT:
			return "DV_E_CLIPFORMAT";
		case DV_E_DVASPECT:
			return "DV_E_DVASPECT";
		case DV_E_DVTARGETDEVICE_SIZE:
			return "DV_E_DVTARGETDEVICE_SIZE";
		case DV_E_NOIVIEWOBJECT:
			return "DV_E_NOIVIEWOBJECT";
		case OPC_E_BADRIGHTS:
			return "OPC_E_BADRIGHTS";
		case OPC_E_INVALIDHANDLE:
			return "OPC_E_INVALIDHANDLE";
		case OPC_E_UNKNOWNITEMID:
			return "OPC_E_UNKNOWNITEMID";
		case OPC_E_BADTYPE:
			return "OPC_E_BADTYPE";
		case OPC_E_PUBLIC:
			return "OPC_E_PUBLIC";
		case OPC_E_INVALIDITEMID:
			return "OPC_E_INVALIDITEMID";
		case OPC_E_INVALIDFILTER:
			return "OPC_E_INVALIDFILTER";
		case OPC_E_UNKNOWNPATH:
			return "OPC_E_UNKNOWNPATH";
		case OPC_E_RANGE:
			return "OPC_E_RANGE";
		case OPC_E_DUPLICATENAME:
			return "OPC_E_DUPLICATENAME";
		case OPC_S_UNSUPPORTEDRATE:
			return "OPC_S_UNSUPPORTEDRATE";
		case OPC_S_CLAMP:
			return "OPC_S_CLAMP";
		case OPC_S_INUSE:
			return "OPC_S_INUSE";
		case OPC_E_INVALIDCONFIGFILE:
			return "OPC_E_INVALIDCONFIGFILE";
		case OPC_E_NOTFOUND:
			return "OPC_E_NOTFOUND";
		case OPC_E_INVALID_PID:
			return "OPC_E_INVALID_PID";
		case OPC_E_DEADBANDNOTSET:
			return "OPC_E_DEADBANDNOTSET";
		case OPC_E_DEADBANDNOTSUPPORTED:
			return "OPC_E_DEADBANDNOTSUPPORTED";
		case OPC_E_NOBUFFERING:
			return "OPC_E_NOBUFFERING";
		case OPC_E_INVALIDCONTINUATIONPOINT:
			return "OPC_E_INVALIDCONTINUATIONPOINT";
		case OPC_S_DATAQUEUEOVERFLOW:
			return "OPC_S_DATAQUEUEOVERFLOW";
		case OPC_E_RATENOTSET:
			return "OPC_E_RATENOTSET";
		case OPC_E_NOTSUPPORTED:
			return "OPC_E_NOTSUPPORTED";
		case OPCCPX_E_TYPE_CHANGED:
			return "OPCCPX_E_TYPE_CHANGED";
		case OPCCPX_E_FILTER_DUPLICATE:
			return "OPCCPX_E_FILTER_DUPLICATE";
		case OPCCPX_E_FILTER_INVALID:
			return "OPCCPX_E_FILTER_INVALID";
		case OPCCPX_E_FILTER_ERROR:
			return "OPCCPX_E_FILTER_ERROR";
		case OPCCPX_S_FILTER_NO_DATA:
			return "OPCCPX_S_FILTER_NO_DATA";
		default:
			return "ERROR_UNKNOWN";
	}
}
