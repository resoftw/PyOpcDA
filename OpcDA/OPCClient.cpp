#include "OPCClient.h"
#define _CRTDBG_MAP_ALLOC

HRESULT InitSecurity(std::string d,std::string u,std::string p)
{
	//CoInitializeEx(NULL, COINIT_APARTMENTTHREADED);
	OleInitialize(NULL);
	std::wstring usr(u.begin(), u.end());
	std::wstring pas(p.begin(), p.end());
	std::wstring dom(d.begin(), d.end());
	
	SEC_WINNT_AUTH_IDENTITY_W authidentity;
	SecureZeroMemory(&authidentity, sizeof(authidentity));

	authidentity.User = (unsigned short*)usr.c_str();
	authidentity.UserLength = usr.length();
	authidentity.Domain = (unsigned short*)dom.c_str();
	authidentity.DomainLength = dom.length();
	authidentity.Password = (unsigned short*)pas.c_str();
	authidentity.PasswordLength = pas.length();
	authidentity.Flags = SEC_WINNT_AUTH_IDENTITY_UNICODE;

	SOLE_AUTHENTICATION_INFO         authninfo[2];
	SecureZeroMemory(authninfo, sizeof(SOLE_AUTHENTICATION_INFO) * 2);

	// NTLM Settings
	authninfo[0].dwAuthnSvc = RPC_C_AUTHN_WINNT;
	authninfo[0].dwAuthzSvc = RPC_C_AUTHZ_NONE;
	authninfo[0].pAuthInfo = &authidentity;

	// Kerberos Settings
	authninfo[1].dwAuthnSvc = RPC_C_AUTHN_GSS_KERBEROS;
	authninfo[1].dwAuthzSvc = RPC_C_AUTHZ_NONE;
	authninfo[1].pAuthInfo = &authidentity;

	SOLE_AUTHENTICATION_LIST    authentlist;

	authentlist.cAuthInfo = 2;
	authentlist.aAuthInfo = authninfo;

	HRESULT hr = CoInitializeSecurity(
		NULL,
		-1,
		NULL,
		NULL,
		RPC_C_AUTHN_LEVEL_CALL,
		RPC_C_IMP_LEVEL_IMPERSONATE,
		&authentlist,
		EOAC_NONE,
		NULL);
	return hr;
}


OPCClient::OPCClient() {
	HostName = L"";
	ProgID = L"";
	ipServer = NULL;
	ipSyncIO = NULL;
	ipGroup = NULL;
	hGroup = 0;
	OnData = NULL;
	_connected = false;
}

OPCClient::~OPCClient()
{
	if (_connected)	Disconnect();
	CoUninitialize();
}

void OPCClient::Error(std::string msg, long err)
{
	fprintf(stderr, "%s Error: 0x%08x\n", msg.c_str(), err);
}
void OPCClient::Error(std::wstring msg, long err)
{
	std::string s(msg.begin(), msg.end());
	fprintf(stderr, "%s wError: 0x%08x\n", s.c_str(), err);
}

std::string OPCClient::CheckStatus()
{
	std::string rslt = "OK";
	OPCSERVERSTATUS* ss = NULL;
	HRESULT h=ipServer->GetStatus(&ss);
	switch (h)
	{
		case E_FAIL:
			rslt = "CheckStatus: The operation failed.";
			break;
		case E_OUTOFMEMORY:
			rslt="CheckStatus: Not enough memory.";
			break;
		case E_INVALIDARG:
			rslt="CheckStatus: An argument to the function was invalid.";
			break;
		case S_OK:
			//con::Info("CheckStatus:\nStatus: ");
			switch (ss->dwServerState) {
				case OPC_STATUS_RUNNING:
					rslt="Running";
					break;
				case OPC_STATUS_FAILED:
					rslt="Failed";
					break;
				case OPC_STATUS_NOCONFIG:
					rslt="No Config";
					break;
				case OPC_STATUS_SUSPENDED:
					rslt="Suspended";
					break;
				case OPC_STATUS_TEST:
					rslt="Test mode";
					break;
				case OPC_STATUS_COMM_FAULT:
					rslt="Comm Fault";
					break;
				default:
					rslt="Unknown";
					break;
			}
			break;
		case (HRESULT)0x800706baL:
			rslt="RPC Server is not availabe.";
			//Reconnect();
			break;
		default:
			rslt = "CheckStatus: Unknown error " + std::to_string(h);
			break;
	}
	if (ss != NULL) {
		if (ss->szVendorInfo != NULL)CoTaskMemFree(ss->szVendorInfo);
		CoTaskMemFree(ss);
	}
	return rslt;
}

HRESULT OPCClient::GetStatus(OpcServerStatus& s)
{
	OPCSERVERSTATUS* ss = NULL;
	HRESULT h = ipServer->GetStatus(&ss);
	if (!FAILED(h)) 
	{
		s.StartTime = filetime_to_timet(ss->ftStartTime);
		s.CurrentTime = filetime_to_timet(ss->ftCurrentTime);
		s.LastUpdateTime = filetime_to_timet(ss->ftLastUpdateTime);
		s.ServerState = ss->dwServerState;
		s.GroupCount = ss->dwGroupCount;
		s.BandWidth = ss->dwBandWidth;
		s.MajorVersion = ss->wMajorVersion;
		s.MinorVersion = ss->wMinorVersion;
		s.BuildNumber = ss->wBuildNumber;
		std::wstring ws(ss->szVendorInfo);
		s.VendorInfo = std::string(ws.begin(), ws.end());
	}
	if (ss != NULL)
	{
		if (ss->szVendorInfo != NULL)CoTaskMemFree(ss->szVendorInfo);
		CoTaskMemFree(ss);
	}
	return h;
}

HRESULT OPCClient::SetGroupState(bool active)
{
	HRESULT hr;
	IOPCGroupStateMgt* grpstatemgr;
	DWORD revisedUpdateRate;
	BOOL activeFlag = active;

	hr = ipGroup->QueryInterface(__uuidof(grpstatemgr), (void**)&grpstatemgr);
	if (FAILED(hr)) {
		Error("SetGroupState: Could not obtain a pointer to IOPCGroupStateMgt",hr);
		return hr;
	}

	hr = grpstatemgr->SetState(NULL, &revisedUpdateRate, &activeFlag, NULL, NULL, NULL, NULL);
	if (FAILED(hr)) {
		Error("SetGroupState: Failed", hr);
	}
	grpstatemgr->Release();
	grpstatemgr = nullptr;
	return hr;
}

void OPCClient::Reconnect()
{
	Disconnect();
	Sleep(3000);//wait for 3secs
	Connect((LPWSTR)HostName.c_str(), (LPWSTR)ProgID.c_str());
	//std::vector<std::wstring> it;
	//it.swap(*itemids);
	//AddItems(it);
}

HRESULT GetCLSID(LPWSTR host, LPWSTR prog, CLSID &opcid)
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
	HRESULT hr = CoCreateInstanceEx(CLSID_OpcServerList, NULL, CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER, &ServerInfo, 1, MultiQI); // ,IID_IOPCServer, (void**)&m_IOPCServer);
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

HRESULT OPCClient::Connect(LPWSTR host, LPWSTR prog) 
{
	if (_connected) return -1;
	LPWSTR szProgID = prog;
	LPWSTR szHostName = host;
	HostName = host;
	ProgID = prog;
	HRESULT hResult = CoInitializeEx(NULL, COINIT_MULTITHREADED);
	if (FAILED(hResult)) {
		Error("CoInitializeEx failed",hResult);
		return hResult;
	}

	CLSID cClsid;
	hResult = GetCLSID(szHostName, szProgID, cClsid);
	if (FAILED(hResult))
	{
		if (UuidFromString((RPC_WSTR)szProgID, &cClsid) != RPC_S_OK) {
			Error(L"Could not resolve Prog ID/CLSID: "+ std::wstring(szProgID) ,hResult);
			return (HRESULT)-1;
		}
	}
	COSERVERINFO cInfo;
	ZeroMemory(&cInfo, sizeof(cInfo));
	cInfo.pwszName = szHostName;
	MULTI_QI cResults;
	ZeroMemory(&cResults, sizeof(cResults));
	cResults.pIID = &IID_IOPCServer;
	hResult = CoCreateInstanceEx(cClsid, NULL, CLSCTX_INPROC_SERVER | CLSCTX_LOCAL_SERVER | CLSCTX_REMOTE_SERVER, &cInfo, 1, &cResults);
	if (FAILED(hResult)) {
		Error("CoCreateInstanceEx failed",hResult);
		return hResult;
	}
	if (FAILED(cResults.hr)) {
		Error("QueryInterface IOPCServer failed",cResults.hr);
		return cResults.hr;
	}
	ipServer = (IOPCServer*)cResults.pItf;
	DWORD dwRevisedUpdateRate = 0;
	hResult = ipServer->AddGroup(L"", true, 1000, NULL, NULL, NULL, LOCALE_SYSTEM_DEFAULT, &hGroup, &dwRevisedUpdateRate, IID_IOPCItemMgt, (IUnknown**)&ipGroup);
	if (FAILED(hResult)) {
		ipServer->Release();
		Error("Add Group failed",hResult);
		return hResult;
	}

	hResult = ipGroup->QueryInterface(IID_IOPCSyncIO, (void**)&ipSyncIO);
	if (FAILED(hResult)) {
		ipGroup->Release();
		ipServer->Release();
		Error("Query Interface SyncIO failed "+ hResultStr(hResult),hResult);
		return hResult;
	}

	printf("Connected to %ls@%ls\n", szProgID, szHostName);
	_connected = true;
	return S_OK;
}

void OPCClient::Disconnect()
{
	if (ipGroup) {
		ipGroup->Release();
	}
	if (ipServer) {
		if (hGroup)	ipServer->RemoveGroup(hGroup, false);
		ipServer->Release();
	}
	ipGroup = NULL;
	hGroup = NULL;
	ipServer = NULL;
	printf("Disconnected from %ls@%ls\n",ProgID.c_str(), HostName.c_str());
	_connected = false;
}


HRESULT OPCClient::AddItem(int cid, std::string const& itemId, VARENUM type, OpcItem& addedInfo)
{
	// checks if an element with this key exists in the map...
	if (itemids.count(itemId) == 1)
	{
		return S_OK;
	}

	// item add result array.
	OPCITEMRESULT* result = nullptr;

	// item add errors array.
	HRESULT* errors = nullptr;

	wchar_t* id = convertMBSToWCS(itemId.c_str());

	OPCITEMDEF item{
		/*szAccessPath*/        NULL,
		/*szItemID*/            id,
		/*bActive*/             true,
		/*hClient*/             cid,
		/*dwBlobSize*/          0,
		/*pBlob*/               NULL,
		/*vtRequestedDataType*/ type,
		/*wReserved*/           0
	};

	// adds the items to the group.
	HRESULT hr = ipGroup->AddItems(1, &item, &result, &errors);


	// if succeeds, adds the item to the map
	if (hr == S_OK)
	{
		// creates the item info...
		addedInfo.ItemID = itemId;
		addedInfo.Handle = result[0].hServer;
		addedInfo.DataType = (VARENUM)result[0].vtCanonicalDataType;

		// adds the item handle to the map...
		iteminfos.push_back(addedInfo);
		addedInfo.index = iteminfos.size() - 1;

		itemids.emplace(itemId, addedInfo);
#ifdef DEBUG_ON
		std::cout<<itemid<<" Added.\n";
#endif
	}
	else
	{
		Error("Error adding "+itemId, hr);
	}

	// frees the memory allocated by the server.
	CoTaskMemFree(result->pBlob);
	CoTaskMemFree(result);
	result = nullptr;

	CoTaskMemFree(errors);
	errors = nullptr;

	return hr;
}

/* INTERNAL */
HRESULT OPCClient::_removeItem(OPCHANDLE const& handle)
{
	HRESULT* errors;
	HRESULT hr = ipGroup->RemoveItems(1, const_cast<OPCHANDLE*>(&handle), &errors);
	CoTaskMemFree(errors);
	errors = nullptr;
	return hr;
}


HRESULT OPCClient::RemoveItem(OpcItem const& item)
{
	HRESULT hr;
	if (itemids.count(item.ItemID) == 0)
	{
		Error("Remove item: "+item.ItemID+" not found!\n",-1);
		return S_FALSE;
	}
	OpcItem todel = itemids[item.ItemID];
	if (todel.Handle != item.Handle) {
		std::cerr << "Remove item : Item handle does not match! " << item.ItemID << item.Handle << " to " << todel.Handle;
		return S_FALSE;
	}
	if (hr = _removeItem(todel.Handle) == S_OK) {
		auto v = find_if(iteminfos.begin(), iteminfos.end(), 
			[&](OpcItem const& inf) 
			{
				return inf.Handle == item.Handle && inf.ItemID == item.ItemID;
			}
		);
		if (v != iteminfos.end())
			iteminfos.erase(v);
		itemids.erase(item.ItemID);
	}
	else {
		Error("RemoveItem: Error removing "+ item.ItemID, hr);
	}
}

HRESULT OPCClient::RemoveAllItems()
{
	HRESULT hr;
	for (auto it = itemids.begin(); it != itemids.end();)
	{
		hr = _removeItem(it->second.Handle);
		if (hr == S_OK)
		{
#ifdef DEBUG_ON
			//con::Info("%s removed.\n", it->second.ItemID.c_str());
#endif
			auto v = find_if(iteminfos.begin(), iteminfos.end(),
				[&](OpcItem const& inf)
				{
					return inf.Handle == it->second.Handle && inf.ItemID == it->second.ItemID;
				}
			);
			if (v != iteminfos.end())iteminfos.erase(v);
			it = itemids.erase(it);
		}
		else {
			Error("RemoveAllItems: Error removing "+it->second.ItemID, hr);
			++it;
		}
	}
	return itemids.empty() ? S_OK : S_FALSE;
}

HRESULT OPCClient::ReadItems()
{
	int nt = itemids.size();
	OPCHANDLE* opch = new OPCHANDLE[nt];
	int i = 0;
	for (auto it = itemids.begin(); it != itemids.end(); ++it)
	{
		opch[i] = it->second.Handle;
		++i;
	}
	HRESULT* errs = NULL;
	OPCITEMSTATE* vals = NULL;
	HRESULT h = ipSyncIO->Read(OPC_DS_CACHE, nt, opch, &vals, &errs);
	if (h == S_OK)
	{
		if (OnData != NULL)
		{
			std::vector<std::unique_ptr<OpcData>> Dat;
			for (i = 0; i < nt; ++i)
			{
				std::unique_ptr<OpcData> d = std::make_unique<OpcData>();
				time_t tm = filetime_to_timet(vals[i].ftTimeStamp);
				d->Timestamp = tm;
				d->Handle = vals[i].hClient;
				d->Quality = vals[i].wQuality;
				d->Value = vals[i].vDataValue;
				Dat.push_back(std::move(d));
			}
			OnData(Dat);
		}

	}
	else if (h == S_FALSE) {
		if (OnData != NULL)
		{
			std::vector<std::unique_ptr<OpcData>> Dat;
			for (i = 0; i < nt; ++i)
			{
				if (errs[i] == S_OK)
				{
					std::unique_ptr<OpcData> d = std::make_unique<OpcData>();
					time_t tm = filetime_to_timet(vals[i].ftTimeStamp);
					d->Timestamp = tm;
					d->Handle = vals[i].hClient;
					d->Quality = vals[i].wQuality;
					d->Value = vals[i].vDataValue;
					Dat.push_back(std::move(d));
				}
			}
			OnData(Dat);
		}
	}
	else {
		delete[] opch;
		Error("Read failed: "+ hResultStr(h), h);
		return h;
	}
	CoTaskMemFree(errs);
	for (i = 0; i < nt; ++i)
	{
		VariantClear(&vals[i].vDataValue);
	}
	CoTaskMemFree(vals);
	delete[] opch;
	return S_OK;
}

HRESULT OPCClient::Write(OpcItem const& item, VARIANT& value)
{
	HRESULT hr;
	if (itemids.count(item.ItemID) == 0)
	{
		Error("Write: " + item.ItemID + " not found!", -1);
		return S_FALSE;
	}
	OpcItem towrite = itemids[item.ItemID];
	if (towrite.Handle != item.Handle) {
		std::cerr << "Write item : Item handle does not match! " << item.ItemID << item.Handle << " to " << towrite.Handle;
		return S_FALSE;
	}
	HRESULT* errors = nullptr;
	hr = ipSyncIO->Write(1, const_cast<OPCHANDLE*>(&towrite.Handle), &value, &errors);

	if (hr != S_OK || &value == nullptr)
	{
		Error("Write "+ towrite.ItemID, hr);
		return hr;
	}

	//Release memeory allocated by the OPC server:
	CoTaskMemFree(errors);
	errors = nullptr;

	return hr;
}

