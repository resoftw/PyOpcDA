#include "OPCClient.h"

std::mutex OPCCallback::opccbmutex;

OPCCallback::OPCCallback(DataCB p) {
	m_ulRefs = 1;
	dataCBHandler = p;
}

STDMETHODIMP OPCCallback::QueryInterface(REFIID iid, LPVOID* ppInterface)
{
	if (ppInterface == NULL)
	{
		return E_INVALIDARG;
	}
		
	if (iid == IID_IUnknown)
	{
		*ppInterface = dynamic_cast<IUnknown*>(this);
		AddRef();
		return S_OK;
	}

	if (iid == IID_IOPCDataCallback)
	{
		*ppInterface = dynamic_cast<IOPCDataCallback*>(this);
		AddRef();
		return S_OK;
	}

	return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) OPCCallback::AddRef()
{
	return InterlockedIncrement((LONG*)&m_ulRefs);
}

STDMETHODIMP_(ULONG) OPCCallback::Release()
{
	ULONG ulRefs = InterlockedDecrement((LONG*)&m_ulRefs);

	if (ulRefs == 0)
	{
		delete this;
		return 0;
	}

	return ulRefs;
}

STDMETHODIMP OPCCallback::OnDataChange(
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
)
{
	std::lock_guard<std::mutex> lock(opccbmutex);
#ifdef DEBUG_ON
	//con::Out("CALLBACK IN %d | ",dwCount);
#endif

	if (dataCBHandler != NULL)
	{
		std::vector<std::unique_ptr<OpcData>> Dat;
		for (DWORD ii = 0; ii < dwCount; ii++)
		{
			std::unique_ptr<OpcData> d=std::make_unique<OpcData>();
			time_t tm = filetime_to_timet(pftTimeStamps[ii]);
			d->Timestamp = tm;
			d->Handle = phClientItems[ii];
			d->Quality = pwQualities[ii];
			//VariantInit(&d.Value);
			//if (VariantCopy(&d.Value, &pvValues[ii]) == S_OK)
			//{
			//	Dat->push_back(d);
			//};
			d->Value = pvValues[ii];
			Dat.push_back(std::move(d));
		}
		dataCBHandler(Dat);
	}
/*
		VARIANT vValue;
		VariantInit(&vValue);

		if (SUCCEEDED(VariantChangeType(&vValue, &(pvValues[ii]), NULL, VT_BSTR)))
		{
			//con::wInfo(L"Handle = %08X, Value = '%s'\n",phClientItems[ii],OLE2T(vValue.bstrVal));
			VariantClear(&vValue);
		} */
#ifdef DEBUG_ON
	//con::Out("CALLBACK OUT");
#endif

	return S_OK;
}

// OnReadComplete
STDMETHODIMP OPCCallback::OnReadComplete(
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
)
{
	return S_OK;
}

// OnWriteComplete
STDMETHODIMP OPCCallback::OnWriteComplete(
	DWORD       dwTransid,
	OPCHANDLE   hGroup,
	HRESULT     hrMastererr,
	DWORD       dwCount,
	OPCHANDLE* pClienthandles,
	HRESULT* pErrors
)
{
	return S_OK;
}


// OnCancelComplete
STDMETHODIMP OPCCallback::OnCancelComplete(
	DWORD       dwTransid,
	OPCHANDLE   hGroup
)
{
	return S_OK;
}

