// CToolTip32Server.h : Declaration of the CToolTip32Server

#pragma once
#include "resource.h"       // main symbols
#include <vector>
using namespace std;


#include "DeveloperExtensionsToolTip32Server_i.h"



using namespace ATL;


// CToolTip32Server

class ATL_NO_VTABLE CToolTip32Server :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CToolTip32Server, &CLSID_ToolTip32Server>,
	public IDispatchImpl<IDEToolTip32Server, &IID_IDEToolTip32Server, &LIBID_DeveloperExtensionsToolTip32ServerLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
	bool IsDllCOMServer(wchar_t *pwcs);
	CComBSTR IsRegistredInCOMPlus(wchar_t *pwcs);
	vector<CComBSTR> ContainsProgID(wchar_t *pwcs);
	DWORD GetBaseAddressOf(wchar_t *pwcs);

public:
	CToolTip32Server()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CTOOLTIP32SERVER)

DECLARE_NOT_AGGREGATABLE(CToolTip32Server)

BEGIN_COM_MAP(CToolTip32Server)
	COM_INTERFACE_ENTRY(IDEToolTip32Server)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

	// IDEToolTip32Server
	STDMETHOD(GetToolTip)(/*[in]*/ BSTR filename, /*[out, retval]*/ BSTR* tooltip);

};

OBJECT_ENTRY_AUTO(__uuidof(ToolTip32Server), CToolTip32Server)
