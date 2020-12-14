// CContextMenu.h : Declaration of the CContextMenu

#pragma once
#include "resource.h"       // main symbols
#include "comdef.h"
#include "shlobj.h"
#include "shlguid.h"
#include <vector>
#include <string>
#include "STarget.h"
#include "DeveloperExtensionsContextMenu_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace std;
using namespace ATL;

#define IDB_DEVELOPEREXTENSIONS_MENU    103
#define IDI_RUNCOMMAND                  105

// CContextMenu

class ATL_NO_VTABLE CContextMenu :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CContextMenu, &CLSID_ContextMenu>,
	public IContextMenu,
	public IShellExtInit,	
	public IPersistFile,
	public IDispatchImpl<IDEContextMenu, &IID_IDEContextMenu, &LIBID_DeveloperExtensionsContextMenuLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
	LPCITEMIDLIST		m_pidlFolder;
	vector<STarget>		m_vTargets;
	vector<UINT>		m_vMenuCommands;

	IDataObject			*m_pDataObject;
	bool				m_bIsComServer;
	bool				m_bIsManaged;
	bool				m_bIsDLL;
	bool				m_bIsTLB;
	bool				m_bIsEXE;
	bool				m_bIsNETMODULE;
	bool				m_bIsDirectory;
	HBITMAP				m_hBmpItem;
	HICON				m_hRunIcon;
	HMODULE				m_hModuleDll;
	//CConfigContextMenu* m_pCConfigContextMenu;
	wchar_t				m_wcsFileName[1024];	

	bool IsManaged(wchar_t *sFilename);
	void IsManagedVersion(const wchar_t* pwcsFilename, wchar_t* pwcsManagedVersion, int nCount);
	vector<CComBSTR> ContainsProgID(wchar_t *pwcs);
	CComBSTR IsRegistredInCOMPlus(wchar_t *pwcs);
	bool IsDllCOMServer(wchar_t *pwcs);
	DWORD GetBaseAddressOf(wchar_t *pwcs);
	wchar_t *BigNumToString(LONG lNum, wchar_t *wcsBuffer);
	wchar_t *BigNumToStringF(float fNum, wchar_t *wcsBuffer);
	void CheckIfXXX();
	HRESULT ExtractPath( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* path);
	HRESULT ExtractFilename( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* filename);
	HMENU CreateMenuForTarget(STarget target, UINT idCmdFirst);
	STarget GetCommonTarget();
	bool GetNETFolder(wchar_t* version, wchar_t* folder, rsize_t size);
	DWORD GetActualAddressFromRVA(IMAGE_SECTION_HEADER* pSectionHeader, IMAGE_NT_HEADERS* pNTHeaders, DWORD dwRVA);
	bool FindPathForCommand(wstring command, wstring& outputPathAndCommand);
	wstring SaveTargetsForEncDec();
public:
	CContextMenu();
	~CContextMenu();

	friend INT_PTR CALLBACK RunCommandDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);	
DECLARE_REGISTRY_RESOURCEID(IDR_CCONTEXTMENU)

DECLARE_NOT_AGGREGATABLE(CContextMenu)

BEGIN_COM_MAP(CContextMenu)	
	COM_INTERFACE_ENTRY(IDEContextMenu)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IContextMenu)
	COM_INTERFACE_ENTRY(IShellExtInit)
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

	// IDEContextMenu

	// IShellExtInit
	STDMETHOD(Initialize)(LPCITEMIDLIST pidlFolder, LPDATAOBJECT lpdobj, HKEY hkeyProgID);

	// IContextMenu
	STDMETHOD(GetCommandString)(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax);
	STDMETHOD(InvokeCommand)(LPCMINVOKECOMMANDINFO pici);
	STDMETHOD(QueryContextMenu)(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags);

	// IPersistFile
	STDMETHOD(Load)(LPCOLESTR wszFile, DWORD dwMode);
	STDMETHOD(GetClassID)(LPCLSID pCLSID);
	STDMETHOD(IsDirty)(VOID);
	STDMETHOD(Save)(LPCOLESTR, BOOL);
	STDMETHOD(SaveCompleted)(LPCOLESTR);
	STDMETHOD(GetCurFile)(LPOLESTR FAR*);

};

OBJECT_ENTRY_AUTO(__uuidof(ContextMenu), CContextMenu)
