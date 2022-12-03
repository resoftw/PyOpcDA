#pragma once
#include <iostream>
#include <iomanip>
#include <time.h>
#include <atlconv.h>
#include <vector>
#include <sstream>
#include <unordered_map>
#include "Utils.h"
#include "opc/opcda.h"
#include "opc/OpcEnum.h"
#include "opc/opccomn.h"
#include "opc/opcerror.h"
#include <mutex>
#include <functional>

#pragma comment( lib, "rpcrt4.lib" )


struct OpcItem {
	OPCHANDLE Handle{ 0 };
	std::string ItemID;
	bool Active{ 0 };
	VARTYPE DataType{ 0 };
	int index{ 0 };
};


struct OpcData {
	OPCHANDLE Handle;
	time_t Timestamp;
	WORD Quality;
	VARIANT Value;
};


struct OpcServerStatus {
	time_t StartTime;
	time_t CurrentTime;
	time_t LastUpdateTime;
	DWORD ServerState;
	DWORD GroupCount;
	DWORD BandWidth;
	WORD MajorVersion;
	WORD MinorVersion;
	WORD BuildNumber;
	std::string VendorInfo;
};

typedef std::function<void(std::vector<std::unique_ptr<OpcData>> const&)> DataCB;

class OPCCallback : public IOPCDataCallback
{
private:
	ULONG m_ulRefs;
	static std::mutex opccbmutex;
	DataCB dataCBHandler;
public:
	OPCCallback(DataCB);
	STDMETHODIMP QueryInterface(REFIID iid, LPVOID* ppInterface);
	STDMETHODIMP_(ULONG) AddRef();
	STDMETHODIMP_(ULONG) Release();
	STDMETHODIMP OnDataChange(
		DWORD       dwTransid,
		OPCHANDLE   hGroup,
		HRESULT     hrMasterquality,
		HRESULT     hrMastererror,
		DWORD       dwCount,
		OPCHANDLE* phClientItems,
		VARIANT* pvValues,
		WORD* pwQualities,
		FILETIME* pftTimeStamps,
		HRESULT* pErrors
	);
	STDMETHODIMP OnReadComplete(
		DWORD       dwTransid,
		OPCHANDLE   hGroup,
		HRESULT     hrMasterquality,
		HRESULT     hrMastererror,
		DWORD       dwCount,
		OPCHANDLE* phClientItems,
		VARIANT* pvValues,
		WORD* pwQualities,
		FILETIME* pftTimeStamps,
		HRESULT* pErrors
	);
	STDMETHODIMP OnWriteComplete(
		DWORD       dwTransid,
		OPCHANDLE   hGroup,
		HRESULT     hrMastererr,
		DWORD       dwCount,
		OPCHANDLE* pClienthandles,
		HRESULT* pErrors
	);
	STDMETHODIMP OnCancelComplete(DWORD dwTransid, OPCHANDLE hGroup);
};


HRESULT InitSecurity(std::string d, std::string u, std::string p);

class OPCClient
{
private:
	std::wstring HostName;
	std::wstring ProgID;
	IOPCServer* ipServer;
	IOPCItemMgt* ipGroup;
	IOPCSyncIO* ipSyncIO;
	OPCHANDLE hGroup;
	bool _connected;
	std::unordered_map<std::string,OpcItem> itemids;
	std::vector<OpcItem> iteminfos;
	HRESULT _removeItem(OPCHANDLE const& handle);
public:
	OPCClient();
	~OPCClient();
	HRESULT Connect(LPWSTR host, LPWSTR prog);
	void Disconnect();
	void Reconnect();
	void Error(std::string msg, long err);
	void Error(std::wstring msg, long err);
	std::string CheckStatus();
	HRESULT GetStatus(OpcServerStatus&);
	HRESULT SetGroupState(bool active);
	DataCB OnData;
	HRESULT AddItem(int cid, std::string const& itemId, VARENUM type, OpcItem& addedInfo);
	HRESULT RemoveItem(OpcItem const& item);
	HRESULT RemoveAllItems();
	HRESULT ReadItems();
	HRESULT Write(OpcItem const& item, VARIANT& value);
};

