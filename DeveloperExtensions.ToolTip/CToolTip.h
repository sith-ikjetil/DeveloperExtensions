// CToolTip.h : Declaration of the CToolTip

#pragma once
#include "stdafx.h"
#include "resource.h"       // main symbols
#include "EnumFileType.h"
#include "comdef.h"
#include "shlobj.h"
#include "shlguid.h"
#include "DeveloperExtensionsToolTip_i.h"
#include <vector>
using namespace std;

#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;


// CToolTip

class ATL_NO_VTABLE CToolTip :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CToolTip, &CLSID_ToolTip>,	
	public IQueryInfo,
	public IPersistFile,
	public IDispatchImpl<IDEToolTip, &IID_IDEToolTip, &LIBID_DeveloperExtensionsToolTipLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
	EnumFileType m_fileType;
	wchar_t	m_wcsToolTipFilename[1024];

	HRESULT ExtractPath( /*[in]*/ BSTR full_path,/*[out, retval]*/ BSTR* path);
	HRESULT ExtractFilename( /*[in]*/ BSTR full_path,/*[out, retval]*/ BSTR* filename);

public:
	CToolTip()
	{
		this->m_fileType = EnumFileType::Unknwn;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CTOOLTIP)

DECLARE_NOT_AGGREGATABLE(CToolTip)

BEGIN_COM_MAP(CToolTip)
	COM_INTERFACE_ENTRY(IDEToolTip)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IQueryInfo)
	COM_INTERFACE_ENTRY(IPersistFile)	
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

	// IQueryInfo
	STDMETHOD(GetInfoFlags)(DWORD *pdwFlags);
	STDMETHOD(GetInfoTip)(DWORD dwFlags, LPWSTR *ppwszTip);

	// IPersistFile
	STDMETHOD(Load)(LPCOLESTR wszFile, DWORD dwMode);
	STDMETHOD(GetClassID)(LPCLSID pCLSID);
	STDMETHOD(IsDirty)(VOID);
	STDMETHOD(Save)(LPCOLESTR, BOOL);
	STDMETHOD(SaveCompleted)(LPCOLESTR);
	STDMETHOD(GetCurFile)(LPOLESTR FAR*);

};

OBJECT_ENTRY_AUTO(__uuidof(ToolTip), CToolTip)
