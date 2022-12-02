#include "OPC.h"
#include <windows.h>
#include <combaseapi.h>

HRESULT GetCLSID(LPWSTR host, LPWSTR prog, CLSID *opcid)
{
	IOPCServerList* m_spServerList = NULL;
	IOPCServerList2* m_spServerList2 = NULL;

	COSERVERINFO ServerInfo = { 0 };
	ServerInfo.pwszName = host;
	ServerInfo.pAuthInfo = NULL;

	MULTI_QI MultiQI[2] = { 0 };

	MultiQI[0].pIID = &IID_IOPCServerList;
	MultiQI[0].pItf = NULL;
	MultiQI[0].hr = S_OK;

	MultiQI[1].pIID = &IID_IOPCServerList2;
	MultiQI[1].pItf = NULL;
	MultiQI[1].hr = S_OK;

	//  Create the OPC server object and query for the IOPCServer interface of the object
	HRESULT hr = CoCreateInstanceEx((const IID* const)CLSID_OpcServerList, NULL, CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER, &ServerInfo, 1, MultiQI); // ,IID_IOPCServer, (void**)&m_IOPCServer);
	//hr = CoCreateInstance (CLSID_OpcServerList, NULL, CLSCTX_LOCAL_SERVER ,IID_IOPCServerList, (void**)&m_spServerList);
	if (hr == S_OK)
	{

		m_spServerList = (IOPCServerList*)MultiQI[0].pItf;
		//m_spServerList2 = (IOPCServerList2*)MultiQI[1].pItf;
	}
	else
	{
		return hr;
	}



	//try and get the class id of the OPC server on the remote host

	CLSID opcServerId;


	if (m_spServerList)
	{
		hr = m_spServerList->CLSIDFromProgID(prog, &opcServerId);
		m_spServerList->Release();
	}
	else
	{
		hr = S_FALSE;
		opcServerId.Data1 = 0;

	}

	//try to attach to an existing OPC Server (so our OPC server is a proxy)

	if (hr != S_OK)
	{
		//wprintf(L"Couldn't get class id for %s\n Return value: %p", prog, (void*)hr);
	}
	else
	{
		opcid = opcServerId;
		//printf("OPC Server clsid: %p %p %p %p%p%p%p%p%p%p%p\n", (void*)opcServerId.Data1, (void*)opcServerId.Data2, (void*)opcServerId.Data3, (void*)opcServerId.Data4[0], (void*)opcServerId.Data4[1], (void*)opcServerId.Data4[2], (void*)opcServerId.Data4[3], (void*)opcServerId.Data4[4], (void*)opcServerId.Data4[5], (void*)opcServerId.Data4[6], (void*)opcServerId.Data4[7]);
	}
	return hr;
}

