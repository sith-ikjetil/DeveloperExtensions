// CToolTip64Server.h : Declaration of the CToolTip64Server

#pragma once
#include "resource.h"       // main symbols
#include <vector>
using namespace std;


#include "DeveloperExtensionsToolTip64Server_i.h"



using namespace ATL;


// CToolTip64Server

class ATL_NO_VTABLE CToolTip64Server :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CToolTip64Server, &CLSID_ToolTip64Server>,
	public IDispatchImpl<IDEToolTip64Server, &IID_IDEToolTip64Server, &LIBID_DeveloperExtensionsToolTip64ServerLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
	bool IsDllCOMServer(wchar_t *pwcs);
	CComBSTR IsRegistredInCOMPlus(wchar_t *pwcs);
	vector<CComBSTR> ContainsProgID(wchar_t *pwcs);
	DWORD GetBaseAddressOf(wchar_t *pwcs);
public:
	CToolTip64Server()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CTOOLTIP64SERVER)

DECLARE_NOT_AGGREGATABLE(CToolTip64Server)

BEGIN_COM_MAP(CToolTip64Server)
	COM_INTERFACE_ENTRY(IDEToolTip64Server)
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

	// IDEToolTip64Server
	STDMETHOD(GetToolTip)(/*[in]*/ BSTR filename, /*[out, retval]*/ BSTR* tooltip);

};

OBJECT_ENTRY_AUTO(__uuidof(ToolTip64Server), CToolTip64Server)
