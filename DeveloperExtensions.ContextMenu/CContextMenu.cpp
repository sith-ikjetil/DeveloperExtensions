// CContextMenu.cpp : Implementation of CContextMenu

#include "stdafx.h"
#include <MsXml6.h>
#include "windows.h"
#include "CContextMenu.h"
#include "atlstr.h"
#include "exdisp.h"
#include "oleauto.h"
#include "comadmin.h"
#include "tlhelp32.h"
#include "commctrl.h"
#include "objbase.h"
#include "STarget.h"
#include <cor.h>
#include <thread>
#include "../Include/itsoftware.h"
#include "../Include/itsoftware-com.h"
#include "../Include/itsoftware-win.h"


//
// using namespace
//
using namespace ItSoftware;
using namespace ItSoftware::COM;
using namespace ItSoftware::Win;

//STDAPI DllRegisterServerType
typedef HRESULT(WINAPI* DllRegisterServerType)(void);

//STDAPI DllUnregisterServer(void)
typedef HRESULT(WINAPI* DllUnregisterServerType)(void);

//__stdcall DevExtShowAboutDlg(HINSTANCE hInstance)
typedef void (WINAPI* DevExtShowAboutDlgType)(void);

// GetFileVersionType
typedef HRESULT(WINAPI* GetFileVersionType)(LPCWSTR szFilename, LPWSTR szBuffer, DWORD cchBuffer, DWORD *dwLength);

// Global extern variable
extern HINSTANCE g_hInstance;

// RUN Dialog Variables
INT_PTR CALLBACK RunCommandDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
HICON g_hRunIcon = NULL;
vector<STarget>* g_pvTargets;
CComBSTR g_bstrWorkingFolder;
CContextMenu* g_pCContextMenu = NULL;
CComBSTR g_bstrModuleDirectory;
//
// Constructor
//
CContextMenu::CContextMenu()
{
	m_pidlFolder = NULL;
	m_pDataObject = NULL;
	m_bIsDLL = false;
	m_bIsTLB = false;
	m_bIsComServer = false;
	m_bIsManaged = false;
	m_bIsNETMODULE = false;

	wchar_t buffer[1024];
	GetModuleFileNameW(g_hInstance, buffer, 1024);

	CComBSTR bstrFullPath(buffer);
	CComBSTR bstrPathOnly;
	this->ExtractPath(bstrFullPath, &bstrPathOnly);
	g_bstrModuleDirectory = bstrPathOnly.Copy();

	CComBSTR bstrLibrary(bstrPathOnly);
	if (bstrLibrary.operator LPWSTR()[bstrLibrary.Length() - 1] != L'\\') {
		bstrLibrary.Append(L"\\");
	}
	bstrLibrary.Append(L"DeveloperExtensions.Common.dll");
	m_hModuleDll = LoadLibraryW(bstrLibrary.operator LPWSTR());
	if (m_hModuleDll != NULL) {		
		m_hBmpItem = LoadBitmap(m_hModuleDll, MAKEINTRESOURCE(IDB_DEVELOPEREXTENSIONS_MENU));
		m_hRunIcon = LoadIcon(m_hModuleDll, MAKEINTRESOURCE(IDI_RUNCOMMAND));
		g_hRunIcon = this->m_hRunIcon;
	}

	g_pvTargets = &this->m_vTargets;
	g_pCContextMenu = this;
}

//
// Destructor
//
CContextMenu::~CContextMenu()
{
	if (m_hBmpItem) {
		DeleteObject(m_hBmpItem);
	}
	if (m_hModuleDll) {
		FreeLibrary(m_hModuleDll);
	}
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// IShellExtInit //////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Initialize
//
STDMETHODIMP CContextMenu::Initialize(LPCITEMIDLIST pidlFolder, LPDATAOBJECT lpdobj, HKEY hkeyProgID)
{
	if (this->m_hModuleDll == NULL) {
		MessageBoxW(GetActiveWindow(), L"Could not load DeveloperExtensions.Common.dll", L"Error Initialize on Developer Extensions", MB_OK | MB_ICONERROR);
		return Error(CComBSTR("Could not load DeveloperExtensions.Common.dll"), IID_IShellExtInit, E_FAIL);
	}

	//Store the PIDL.
	if (pidlFolder)
	{
		this->m_pidlFolder = pidlFolder;
	}

	// If Initialize has already been called, release the old
	// IDataObject pointer.
	if (m_pDataObject)
	{
		this->m_pDataObject->Release();
	}

	// Get data...
	if (lpdobj) {
		this->m_vTargets.clear();

		this->m_pDataObject = lpdobj;
		this->m_pDataObject->AddRef();

		STGMEDIUM stgm;
		FORMATETC fetc = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };

		HRESULT hr = lpdobj->GetData(&fetc, &stgm);

		switch (hr)
		{
		case S_OK:	// Data was successfully retrieved and placed in the storage medium provided.			
			break;
		case E_INVALIDARG: // Invalid argument		
		case E_UNEXPECTED:	// Unexpected uh oh.			
		case E_OUTOFMEMORY:	// Out of memory			
		case DV_E_LINDEX: // Invalid value for lindex; currently, only -1 is supported. 			
		case DV_E_FORMATETC: // Invalid value for pFormatetc. 
		case DV_E_TYMED: //Invalid tymed value. 
		case DV_E_DVASPECT://Invalid dwAspect value. 
		case OLE_E_NOTRUNNING:  // Object application is not running. 
		case STG_E_MEDIUMFULL: // An error occurred when allocating the medium
		default:
			return E_FAIL;
			break;
		};

		UINT uCount(0);

		// Get the file name from the CF_HDROP.
		uCount = DragQueryFile((HDROP)stgm.hGlobal, (UINT)-1, NULL, 0);
		for (UINT i = 0; i < uCount; i++) {
			STarget fn;
			DragQueryFileW((HDROP)stgm.hGlobal, i, fn.FileName, sizeof(fn.FileName));
			this->m_vTargets.push_back(fn);
		}

		ReleaseStgMedium(&stgm);
	}

	this->m_bIsDLL = false;
	this->m_bIsTLB = false;
	this->m_bIsEXE = false;
	this->m_bIsDirectory = false;
	this->m_bIsNETMODULE = false;

	this->CheckIfXXX();
	return NOERROR;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// IContextMenu ///////////////////////////////////////////////////////////////////////////////////////////////////////
//
// GetCommandString
//
STDMETHODIMP CContextMenu::GetCommandString(UINT_PTR idCmd, UINT uFlags, UINT *pwReserved, LPSTR pszName, UINT cchMax)
{
	return S_OK;
}
//
// InvokeCommand
//
STDMETHODIMP CContextMenu::InvokeCommand(LPCMINVOKECOMMANDINFO pici)
{
	if (m_hModuleDll == NULL) {
		MessageBoxW(GetActiveWindow(), L"Could not load devext.dll", L"Developer Extensions Error", MB_OK | MB_ICONERROR);
		return Error(CComBSTR("Could not load devext.dll"), IID_IContextMenu, E_FAIL);
	}

	// 0-based menu position.
#pragma warning(disable:4311)
#pragma warning(disable:4302)
	int iAction = reinterpret_cast<int>(pici->lpVerb);
	if (iAction > this->m_vMenuCommands.size() ||
		iAction < 0 ) {
		return E_FAIL;
	}

	if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_ABOUT) {
		DevExtShowAboutDlgType fnShowAboutDlg = (DevExtShowAboutDlgType)::GetProcAddress(m_hModuleDll, "DevExtShowAboutDlg");
		if (fnShowAboutDlg != NULL) {
			try {
				fnShowAboutDlg();
			}
			catch (...) {
				CComBSTR bstr("Exception caught while trying to call DevExtShowAboutDlg in devext.dll");
				MessageBoxW(GetActiveWindow(), bstr, L"Error", MB_OK | MB_ICONERROR);
			}
		}
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_COMMANDPROMPT) {
		CString path;
		path = m_vTargets[0].FileName;
		wchar_t bslash[2] = L"\\";
		int iIndex = path.ReverseFind(bslash[0]);
		CString directory = path.Left(iIndex + 1);

		wchar_t wcsCommand[255];
		::GetSystemDirectoryW(wcsCommand, 255);
		wcscat_s(wcsCommand, 255, L"\\cmd.exe");

		SHELLEXECUTEINFOW execInfo = { 0 };
		execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
		execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
		execInfo.lpVerb = L"Open";
		execInfo.nShow = SW_NORMAL;
		execInfo.lpFile = wcsCommand;
		execInfo.lpDirectory = directory.operator LPCWSTR();
		::ShellExecuteExW(&execInfo);
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_COMMANDPROMPT_ADMIN) {
		CString path;
		path = m_vTargets[0].FileName;
		wchar_t bslash[2] = L"\\";
		int iIndex = path.ReverseFind(bslash[0]);
		CString directory = path.Left(iIndex + 1);

		wchar_t wcsCommand[255];
		::GetSystemDirectoryW(wcsCommand, 255);
		wcscat_s(wcsCommand, 255, L"\\cmd.exe");

		wchar_t wcsParameters[255];
		wcscpy_s(wcsParameters, 255, L"/s /k pushd \"");
		wcscat_s(wcsParameters, 255, directory.operator LPCWSTR());
		wcscat_s(wcsParameters, 255, L"\"");

		SHELLEXECUTEINFOW execInfo = { 0 };
		execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
		execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;		
		execInfo.lpVerb = L"RunAs";
		execInfo.nShow = SW_NORMAL;
		execInfo.lpFile = wcsCommand;
		execInfo.lpDirectory = directory.operator LPCWSTR();
		execInfo.lpParameters = wcsParameters;
		::ShellExecuteExW(&execInfo);
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_INSTALLUTIL_I) {
		for (auto target : this->m_vTargets) {
			wchar_t wszPath[1024];
			if (!this->GetNETFolder(target.ManagedVersion,wszPath, 1024)) {
				continue;
			}
			wcscat_s(wszPath, 1024, L"\\");
			
			wcscat_s(wszPath, 1024, L"installutil.exe");
			
			CString path;
			path = target.FileName;
			wchar_t bslash[2] = L"\\";
			int iIndex = path.ReverseFind(bslash[0]);
			CString directory = path.Left(iIndex + 1);

			STARTUPINFOW si = { 0 };
			si.cb = sizeof(STARTUPINFO);
			si.lpReserved = NULL;
			si.lpTitle = L"InstallUtil /i";			

			wchar_t wcsCommand[255];
			wcscpy_s(wcsCommand, 255, wszPath);

			wchar_t wcsParameters[255];
			wcscpy_s(wcsParameters, 255, L"/i ");
			wcscat_s(wcsParameters, 255, target.FileName);
			
			SHELLEXECUTEINFOW execInfo = { 0 };
			execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
			execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
			execInfo.lpVerb = L"RunAs";
			execInfo.nShow = SW_NORMAL;
			execInfo.lpFile = wcsCommand;
			execInfo.lpDirectory = directory.operator LPCWSTR();
			execInfo.lpParameters = wcsParameters;
			::ShellExecuteExW(&execInfo);
		}
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_INSTALLUTIL_U) {
		for (auto target : this->m_vTargets) {
			wchar_t wszPath[1024];
			if (!this->GetNETFolder(target.ManagedVersion, wszPath, 1024)) {
				continue;
			}
			wcscat_s(wszPath, 1024, L"\\");

			wcscat_s(wszPath, 1024, L"installutil.exe");

			CString path;
			path = target.FileName;
			wchar_t bslash[2] = L"\\";
			int iIndex = path.ReverseFind(bslash[0]);
			CString directory = path.Left(iIndex + 1);

			STARTUPINFOW si = { 0 };
			si.cb = sizeof(STARTUPINFO);
			si.lpReserved = NULL;
			si.lpTitle = L"InstallUtil /u";

			wchar_t wcsCommand[255];
			wcscpy_s(wcsCommand, 255, wszPath);

			wchar_t wcsParameters[255];
			wcscpy_s(wcsParameters, 255, L"/u ");
			wcscat_s(wcsParameters, 255, target.FileName);

			SHELLEXECUTEINFOW execInfo = { 0 };
			execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
			execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
			execInfo.lpVerb = L"RunAs";
			execInfo.nShow = SW_NORMAL;
			execInfo.lpFile = wcsCommand;
			execInfo.lpDirectory = directory.operator LPCWSTR();
			execInfo.lpParameters = wcsParameters;
			::ShellExecuteExW(&execInfo);
		}
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_REGISTERCOMSERVER) {
		for (auto target : this->m_vTargets) {
			CString path;
			path = target.FileName;
			wchar_t bslash[2] = L"\\";
			int iIndex = path.ReverseFind(bslash[0]);
			CString directory = path.Left(iIndex + 1);			

			wchar_t wcsCommand[255];
			if (target.IsDLL) {
				::GetSystemDirectoryW(wcsCommand, 255);
				wcscat_s(wcsCommand, 255, L"\\regsvr32.exe");
			}
			else if (target.IsEXE) {
				wcscpy_s(wcsCommand, 255, L"\"");
				wcscat_s(wcsCommand, 255, target.FileName);
				wcscat_s(wcsCommand, 255, L"\"");
			}

			wchar_t wcsParameters[255];	
			if (target.IsDLL) {
				wcscpy_s(wcsParameters, 255, L"\"");
				wcscat_s(wcsParameters, 255, target.FileName);
				wcscat_s(wcsParameters, 255, L"\"");
			}
			else if (target.IsEXE) {
				wcscpy_s(wcsParameters, 255, L"/regserver");
			}
			
			SHELLEXECUTEINFOW execInfo = { 0 };
			execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
			execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
			execInfo.lpVerb = L"RunAs";
			execInfo.nShow = SW_NORMAL;
			execInfo.lpFile = wcsCommand;
			execInfo.lpDirectory = directory.operator LPCWSTR();
			execInfo.lpParameters = wcsParameters;
			::ShellExecuteExW(&execInfo);
		}
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_UNREGISTERCOMSERVER) {
		for (auto target : this->m_vTargets) {
			CString path;
			path = target.FileName;
			wchar_t bslash[2] = L"\\";
			int iIndex = path.ReverseFind(bslash[0]);
			CString directory = path.Left(iIndex + 1);

			wchar_t wcsCommand[255];
			if (target.IsDLL) {
				::GetSystemDirectoryW(wcsCommand, 255);
				wcscat_s(wcsCommand, 255, L"\\regsvr32.exe");
			}
			else if (target.IsEXE) {
				wcscpy_s(wcsCommand, 255, L"\"");
				wcscat_s(wcsCommand, 255, target.FileName);
				wcscat_s(wcsCommand, 255, L"\"");
			}

			wchar_t wcsParameters[255];
			if (target.IsDLL) {
				wcscpy_s(wcsParameters, 255, L"/u \"");
				wcscat_s(wcsParameters, 255, target.FileName);
				wcscat_s(wcsParameters, 255, L"\"");
			}
			else if (target.IsEXE) {
				wcscpy_s(wcsParameters, 255, L"/unregserver");
			}

			SHELLEXECUTEINFOW execInfo = { 0 };
			execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
			execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
			execInfo.lpVerb = L"RunAs";
			execInfo.nShow = SW_NORMAL;
			execInfo.lpFile = wcsCommand;
			execInfo.lpDirectory = directory.operator LPCWSTR();
			execInfo.lpParameters = wcsParameters;
			::ShellExecuteExW(&execInfo);
		}
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_CREATEHARDLINKTO) {
		wchar_t wcsBuffer[MAX_PATH];
		CString path;
		path = m_vTargets[0].FileName;

		CString drive;
		drive = path.Left(3);

		CComPtr<IShellFolder> pshf;
		SHGetDesktopFolder((IShellFolder **)&pshf);

		LPITEMIDLIST piil;		
		pshf->ParseDisplayName(GetActiveWindow(), NULL, drive.AllocSysString(), NULL, &piil, NULL);

		LPITEMIDLIST pidlSelected = NULL;
		BROWSEINFOW bi = { 0 };
		bi.hwndOwner = GetActiveWindow();
		bi.pidlRoot = piil;
		bi.pszDisplayName = wcsBuffer;
		bi.lpszTitle = L"Choose a folder. This folder will contain the new hard link.";
		bi.ulFlags = BIF_USENEWUI | BIF_NEWDIALOGSTYLE | BIF_RETURNONLYFSDIRS | BIF_RETURNONLYFSDIRS;
		bi.lpfn = NULL;
		bi.lParam = 0;

		pidlSelected = SHBrowseForFolderW(&bi);
		if (SHGetPathFromIDListW(pidlSelected, wcsBuffer)) {
			for (auto target : this->m_vTargets) {
				CString path;
				path = target.FileName;
				wchar_t bslash[2] = L"\\";
				int iIndex = path.ReverseFind(bslash[0]);
				CString filename = path.Right(path.GetLength() - iIndex);

				CComBSTR bstr(wcsBuffer);
				bstr += filename.operator LPCTSTR();

				if (!CreateHardLinkW(bstr, target.FileName, NULL)) {
					CComBSTR bstrError("Could not create hard link of file [");
					bstrError += target.FileName;
					bstrError.Append("] to file [");
					bstrError.AppendBSTR(bstr);
					bstrError.Append("]");
					MessageBoxW(GetActiveWindow(), bstrError, L"Error Creating Hard Link", MB_OK | MB_ICONERROR);
				}
			}
		}
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_RUN)
	{
		CString path;
		path = this->m_vTargets[	0].FileName;
		wchar_t bslash[2] = L"\\";
		int iIndex = path.ReverseFind(bslash[0]);
		CString directory = path.Left(iIndex + 1);

		g_bstrWorkingFolder = CComBSTR(directory.operator LPCWSTR());

		::DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_RUNCOMMAND), GetActiveWindow(), RunCommandDialogProc);
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_ENCRYPT)
	{
		RECT rect{ 0 };
		GetClientRect(GetForegroundWindow(), &rect);

		POINT pt{ 0 };
		pt.x = static_cast<LONG>((rect.right - rect.left) / 2);
		pt.y = static_cast<LONG>((rect.bottom - rect.top) / 2);
		ClientToScreen(GetForegroundWindow(), &pt);

		int wposx = pt.x;
		int wposy = pt.y;

		wstring filename = SaveTargetsForEncDec();
		wstringstream parameters;
		parameters << L"action:encrypt file:\"" << filename << L"\" wposx:" << wposx << L" wposy:" << wposy;
		wstring params = parameters.str();

		SHELLEXECUTEINFOW execInfo = { 0 };
		execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
		execInfo.lpVerb = L"Open";
		execInfo.nShow = SW_NORMAL;
		execInfo.lpFile = L"DeveloperExtensions.EncDec.exe";
		execInfo.lpDirectory = g_bstrModuleDirectory.operator LPWSTR();
		execInfo.lpParameters = params.c_str();
		::ShellExecuteExW(&execInfo);		
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_DECRYPT)
	{
		RECT rect{ 0 };
		GetClientRect(GetForegroundWindow(), &rect);

		POINT pt{ 0 };
		pt.x = static_cast<LONG>((rect.right - rect.left) / 2);
		pt.y = static_cast<LONG>((rect.bottom - rect.top) / 2);
		ClientToScreen(GetForegroundWindow(), &pt);

		int wposx = pt.x;
		int wposy = pt.y;

		wstring filename = SaveTargetsForEncDec();
		wstringstream parameters;
		parameters << L"action:decrypt file:\"" << filename << L"\" wposx:" << wposx << L" wposy:" << wposy;
		wstring params = parameters.str();

		SHELLEXECUTEINFOW execInfo = { 0 };
		execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
		execInfo.lpVerb = L"Open";
		execInfo.nShow = SW_NORMAL;
		execInfo.lpFile = L"DeveloperExtensions.EncDec.exe";
		execInfo.lpDirectory = g_bstrModuleDirectory.operator LPWSTR();
		execInfo.lpParameters = params.c_str();
		::ShellExecuteExW(&execInfo);
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_REGISTERTYPELIB) 
	{
		for (auto target : this->m_vTargets) 
		{
			if (!target.IsTLB)
			{
				continue;
			}
			CComPtr<ITypeLib> pITypeLib;
			HRESULT hr = ::LoadTypeLibEx(target.FileName, REGKIND_REGISTER, &pITypeLib);
			if (FAILED(hr)) {
				CComBSTR bstr("Failed to register type library '");
				bstr += target.FileName;
				bstr.Append("'.");
				MessageBoxW(GetActiveWindow(), bstr, L"Developer Extensions - Error", MB_OK | MB_ICONERROR);
			}		
		}
	}
	else if (this->m_vMenuCommands.at(iAction) == ID_DEVELOPEREXTENSION_UNREGISTERTYPELIB)
	{
		for (auto target : this->m_vTargets)
		{
			if (!target.IsTLB)
			{
				continue;
			}
			
			CComPtr<ITypeLib> pITypeLib;
			HRESULT hr = ::LoadTypeLib(target.FileName, &pITypeLib);
			if (FAILED(hr)) {
				CComBSTR bstr("Failed to load type library '");
				bstr += target.FileName;
				bstr.Append("'.");
				MessageBoxW(GetActiveWindow(), bstr, L"Developer Extensions - Error", MB_OK | MB_ICONERROR);
			}
			else {
				TLIBATTR *ptl;
				pITypeLib->GetLibAttr(&ptl);
				
				hr = ::UnRegisterTypeLib(ptl->guid, ptl->wMajorVerNum, ptl->wMinorVerNum, ptl->lcid, ptl->syskind);
				if (FAILED(hr)) {
					CComBSTR bstr("Failed to unregister type library '");
					bstr += target.FileName;
					bstr.Append("' .");
					MessageBoxW(GetActiveWindow(), bstr, L"Developer Extensions - Error", MB_OK | MB_ICONERROR);
				}
				if (pITypeLib != NULL) {
					pITypeLib->ReleaseTLibAttr(ptl);
				}
			}			
		}
	}

	return S_OK;
}

wstring CContextMenu::SaveTargetsForEncDec()
{
	wchar_t path[MAX_PATH];
	GetTempPath(MAX_PATH, path);

	wchar_t file[MAX_PATH];
	GetTempFileName(path, NULL, 0, file);

	wstring filename(file);

	CComPtr<IXMLDOMDocument2> pIXMLDOMDocument;
	HRESULT hr = pIXMLDOMDocument.CoCreateInstance(CLSID_DOMDocument60, NULL, CLSCTX::CLSCTX_INPROC_SERVER);
	if (FAILED(hr))
	{
		return L"";
	}	

	CComPtr<IXMLDOMProcessingInstruction> pi;
	hr = pIXMLDOMDocument->createProcessingInstruction(CComBSTR(L"xml"), CComBSTR(L"version='1.0' encoding='utf-8'"), &pi);
	CComPtr<IXMLDOMNode> npi;
	pIXMLDOMDocument->appendChild(pi, &npi);

	CComPtr<IXMLDOMElement> pIXDE;
	hr = pIXMLDOMDocument->createElement(CComBSTR(L"files"), &pIXDE);
	CComPtr<IXMLDOMNode> pIXDN;
	hr = pIXMLDOMDocument->appendChild(pIXDE, &pIXDN);

	for (auto item : *g_pvTargets)
	{
		if (!item.IsDirectory)
		{
			CComPtr<IXMLDOMElement> element;
			hr = pIXMLDOMDocument->createElement(CComBSTR(L"file"), &element);

			CComPtr<IXMLDOMNode> node;
			hr = pIXDE->appendChild(element, &node);

			hr = node->put_text(CComBSTR(item.FileName));
		}
	}	

	hr = pIXMLDOMDocument->save(CComVariant(filename.c_str()));

	return filename;
}
//
// QueryContextMenu
//
STDMETHODIMP CContextMenu::QueryContextMenu(HMENU hmenu, UINT indexMenu, UINT idCmdFirst, UINT idCmdLast, UINT uFlags)
{
	if (m_hModuleDll == NULL) {
		MessageBoxW(GetActiveWindow(), L"Could not load devext.dll", L"Error InvokeCommand on Developer Extensions", MB_OK | MB_ICONERROR);
		return Error(CComBSTR("Could not load devext.dll"), IID_IContextMenu, E_FAIL);
	}

	HRESULT hr(S_OK);
	HMENU hSubMenu = this->CreateMenuForTarget(this->GetCommonTarget(),idCmdFirst);
	int iMenuCount = GetMenuItemCount(hSubMenu);	
	if (hSubMenu) {	
		MENUITEMINFOW mii;
		mii.cbSize = sizeof(mii);
		mii.hSubMenu = hSubMenu;
		mii.fMask = MIIM_SUBMENU | MIIM_FTYPE | MIIM_STRING | MIIM_BITMAP;
		mii.fType = MFT_STRING;
		mii.fState = MFS_ENABLED | MFS_DEFAULT;
		mii.dwTypeData = L"Developer Extensions";
		mii.hbmpItem = m_hBmpItem;

		// set same space for checkmark or bitmap...
		MENUINFO mi = { 0 };
		mi.cbSize = sizeof(MENUINFO);
		mi.dwStyle = MNS_CHECKORBMP;
		mi.fMask = MIM_STYLE;
		GetMenuInfo(hmenu, &mi);

		mi.dwStyle |= MNS_CHECKORBMP;
		SetMenuInfo(hmenu, &mi);

		// ok insert it...
		BOOL b = InsertMenuItemW(hmenu, indexMenu, TRUE, &mii);
		hr = MAKE_HRESULT(SEVERITY_SUCCESS, 0, iMenuCount);
	}
	
	return hr;
}
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
//
// IPersistFile ///////////////////////////////////////////////////////////////////////////////////////////////////////
//
// Load
// 
STDMETHODIMP CContextMenu::Load(LPCOLESTR wszFile, DWORD dwMode)
{
	wcscpy_s(m_wcsFileName, 1024, wszFile);
	return S_OK;
};
//
// GetClassID
//
STDMETHODIMP CContextMenu::GetClassID(LPCLSID pCLSID)
{
	*pCLSID = CLSID_ContextMenu;
	return S_OK;
}
//
// IsDirty
//
STDMETHODIMP CContextMenu::IsDirty(VOID) 
{ 
	return E_NOTIMPL; 
}
//
// Save
//
STDMETHODIMP CContextMenu::Save(LPCOLESTR, BOOL) 
{ 
	return E_NOTIMPL; 
}
//
// SaveCompleted
//
STDMETHODIMP CContextMenu::SaveCompleted(LPCOLESTR) 
{ 
	return E_NOTIMPL; 
}
//
// GetCurFile
//
STDMETHODIMP CContextMenu::GetCurFile(LPOLESTR FAR*) 
{ 
	return E_NOTIMPL; 
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// IsManaged
//
bool CContextMenu::IsManaged(wchar_t *lpszImageName)
{
	bool bIsManaged = false;    //variable that indicates whether
								//managed or not.
	//TCHAR szPath[MAX_PATH];     //for convenience

	HANDLE hFile = CreateFile(lpszImageName, GENERIC_READ, FILE_SHARE_READ, NULL, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, NULL);			

	if (INVALID_HANDLE_VALUE != hFile)
	{
		//succeeded
		HANDLE hOpenFileMapping = CreateFileMapping(hFile, NULL,PAGE_READONLY, 0,0, NULL);
		if (hOpenFileMapping)
		{
			BYTE* lpBaseAddress = NULL;

			//Map the file, so it can be simply be acted on as a
			//contiguous array of bytes
			lpBaseAddress = (BYTE*)MapViewOfFile(hOpenFileMapping,FILE_MAP_READ, 0, 0, 0);

			if (lpBaseAddress)
			{
				//having mapped the executable, now start navigating
				//through the sections

					//DOS header is straightforward. It is the topmost
					//structure in the PE file
					//i.e. the one at the lowest offset into the file
					IMAGE_DOS_HEADER* pDOSHeader = (IMAGE_DOS_HEADER*)lpBaseAddress;
					
					//the only important data in the DOS header is the
					//e_lfanew
					//the e_lfanew points to the offset of the beginning
					//of NT Headers data
					IMAGE_NT_HEADERS* pNTHeaders = (IMAGE_NT_HEADERS*)((BYTE*)pDOSHeader + pDOSHeader->e_lfanew);

					//store the section header for future use. This will
					//later be need to check to see if metadata lies within
					//the area as indicated by the section headers
					IMAGE_SECTION_HEADER* pSectionHeader = (IMAGE_SECTION_HEADER*)((BYTE*)pNTHeaders + sizeof(IMAGE_NT_HEADERS));
					
					//Now, start parsing
					//First of all check if it is a PE file. All assemblies
					//are PE files.
					if (pNTHeaders->Signature == IMAGE_NT_SIGNATURE)
					{
						//start parsing COM table (this is what points to
						//the metadata and other information)
						DWORD dwNETHeaderTableLocation = pNTHeaders->OptionalHeader.DataDirectory[IMAGE_DIRECTORY_ENTRY_COM_DESCRIPTOR].VirtualAddress;

						if (dwNETHeaderTableLocation)
						{
							//.NET header data does exist for this module;
							//find its location in one of the sections
							IMAGE_COR20_HEADER* pNETHeader = (IMAGE_COR20_HEADER*)((BYTE*)pDOSHeader + GetActualAddressFromRVA(pSectionHeader,pNTHeaders, dwNETHeaderTableLocation));

							if (pNETHeader)
							{
								//valid address obtained. Suffice it to say,
								//this is good enough to identify this as a
								//valid managed component
								bIsManaged = true;
							}
						}
					}
					else
					{
						//cout << "Not PE file\r\n";
					}
					//cleanup
					UnmapViewOfFile(lpBaseAddress);
			}
			//cleanup
			CloseHandle(hOpenFileMapping);				
		}
		//cleanup
		CloseHandle(hFile);
	}

	return bIsManaged;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// IsManagedVersion
//
void CContextMenu::IsManagedVersion(const wchar_t* pwcsFilename, wchar_t* pwcsManagedVersion, int nCount) {

	wcscpy_s(pwcsManagedVersion, nCount, L"");

	HMODULE hModuleDll = LoadLibraryW(L"mscoree.dll");
	if (hModuleDll != NULL) {
		GetFileVersionType fnGetFileVersion = (GetFileVersionType)GetProcAddress(hModuleDll, "GetFileVersion");
		if (fnGetFileVersion != NULL) {
			DWORD dwLength(0);
			HRESULT hr = fnGetFileVersion(pwcsFilename, pwcsManagedVersion, nCount, &dwLength);
		}
	}
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// ContainstProgID
//
vector<CComBSTR> CContextMenu::ContainsProgID(wchar_t *pwcs)
{
	vector<CComBSTR> vProgIDs;
	CComPtr<ICOMAdminCatalog> pAdminCatalog;
	HRESULT hr = pAdminCatalog.CoCreateInstance(CComBSTR("COMAdmin.COMAdminCatalog.1"));//CLSID_COMAdminCatalog );
	if (FAILED(hr)) {
		return vProgIDs;
	}

	CComPtr<IDispatch> pDispatch;
	hr = pAdminCatalog->GetCollection(CComBSTR("InprocServers"), &pDispatch);
	if (FAILED(hr)) {
		return vProgIDs;
	}

	CComQIPtr<ICatalogCollection> pCollection;
	pCollection = pDispatch;

	pCollection->Populate();

	long lCount(0);
	hr = pCollection->get_Count(&lCount);

	long lComponents(0);
	CComBSTR bstrSrc;
	bstrSrc += pwcs;
	bstrSrc.ToLower();

	for (long l = 0; l < lCount; l++) {
		CComPtr<IDispatch> pDispatch;
		hr = pCollection->get_Item(l, &pDispatch);
		if (FAILED(hr)) {
			return vProgIDs;
		}

		CComQIPtr<ICatalogObject> pObject;
		pObject = pDispatch;
		pDispatch = NULL;

		CComVariant vtValue;
		CComBSTR bstrDLL("InprocServer32");
		hr = pObject->get_Value(bstrDLL, &vtValue);
		if (FAILED(hr))
			return vProgIDs;

		CComBSTR bstrValue;
		bstrValue.AppendBSTR(vtValue.bstrVal);
		if (bstrValue.Length()) {
			bstrValue.ToLower();
			if (bstrSrc == bstrValue) {
				CComVariant vtProgID;
				CComBSTR bstrProgID("ProgID");
				hr = pObject->get_Value(bstrProgID, &vtProgID);
				CComBSTR bstrRetVal;
				bstrRetVal.AppendBSTR(vtProgID.bstrVal);
				vProgIDs.push_back(bstrRetVal);
			}
		}
	}// for

	return vProgIDs;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// IsRegisteredInCOMPlus
//
CComBSTR CContextMenu::IsRegistredInCOMPlus(wchar_t *pwcs)
{
	CComBSTR bstrInprocServer;
	bstrInprocServer += pwcs;
	bstrInprocServer.ToLower();

	CComPtr<ICOMAdminCatalog> pAdminCatalog;
	HRESULT hr = pAdminCatalog.CoCreateInstance(CComBSTR("COMAdmin.COMAdminCatalog.1"));//CLSID_COMAdminCatalog );
	if (FAILED(hr)) {
		return NULL;
	}

	CComPtr<IDispatch> pDispatch;
	hr = pAdminCatalog->GetCollection(CComBSTR("Applications"), &pDispatch);
	if (FAILED(hr)) {
		return NULL;
	}

	CComQIPtr<ICatalogCollection> pCollection;
	pCollection = pDispatch;

	pCollection->Populate();

	long lCount(0);
	hr = pCollection->get_Count(&lCount);

	long lComponents(0);
	for (long l = 0; l < lCount; l++) {
		CComPtr<IDispatch> pDispatch;
		hr = pCollection->get_Item(l, &pDispatch);
		if (FAILED(hr)) {
			return NULL;
		}

		CComQIPtr<ICatalogObject> pObject;
		pObject = pDispatch;
		pDispatch = NULL;

		CComVariant vtName;
		hr = pObject->get_Value(CComBSTR("Name"), &vtName);
		if (FAILED(hr))
			return NULL;

		CComVariant vtKey;
		hr = pObject->get_Key(&vtKey);
		if (FAILED(hr))
			return NULL;

		hr = pCollection->GetCollection(CComBSTR("Components"), vtKey, &pDispatch);
		CComQIPtr<ICatalogCollection> pCollectionC;
		pCollectionC = pDispatch;
		pDispatch = NULL;

		pCollectionC->Populate();

		long lCount(0);
		hr = pCollectionC->get_Count(&lCount);
		for (int i = 0; i < lCount; i++) {
			CComPtr<IDispatch> pDispatch;
			hr = pCollectionC->get_Item(i, &pDispatch);
			if (FAILED(hr)) {
				return NULL;
			}

			CComQIPtr<ICatalogObject> pObject;
			pObject = pDispatch;
			pDispatch = NULL;

			CComVariant vtValue;
			CComBSTR bstrDLL("DLL");
			hr = pObject->get_Value(bstrDLL, &vtValue);
			if (FAILED(hr)) {
				return NULL;
			}

			CComBSTR bstrFile;
			bstrFile.AppendBSTR(vtValue.bstrVal);
			bstrFile.ToLower();

			if (bstrInprocServer == bstrFile) {
				CComBSTR bstr;
				bstr.AppendBSTR(vtName.bstrVal);
				return bstr;
			}
		}
	}// for

	return NULL;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// IsDllCOMServer
//
bool CContextMenu::IsDllCOMServer(wchar_t *pwcs)
{
	HMODULE hModule = ::LoadLibraryEx(pwcs, NULL, DONT_RESOLVE_DLL_REFERENCES);
	if (hModule != NULL) {
		DllRegisterServerType fnRegisterServer = (DllRegisterServerType) ::GetProcAddress(hModule, "DllRegisterServer");
		if (fnRegisterServer != NULL) {
			FreeLibrary(hModule);			
			return true;
		}
		FreeLibrary(hModule);
	}

	return false;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// GetBaseAddressOf
//
DWORD CContextMenu::GetBaseAddressOf(wchar_t *pwcs)
{
	HMODULE hModule = ::LoadLibraryEx(pwcs, NULL, DONT_RESOLVE_DLL_REFERENCES);
	if (hModule == NULL)
	{		
		return 0;
	}

	PVOID pvModulePreferredBaseAddr(NULL);
	IMAGE_DOS_HEADER idh = { 0 };
	IMAGE_NT_HEADERS inth = { 0 };

	HANDLE hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, GetCurrentProcessId());
	if (hSnapshot == INVALID_HANDLE_VALUE)
		return 0;

	MODULEENTRY32 me = { 0 };
	me.dwSize = sizeof(me);
	BOOL bRet = Module32First(hSnapshot, &me);
	if (me.hModule == hModule) {
		Toolhelp32ReadProcessMemory(me.th32ProcessID, me.modBaseAddr, &idh, sizeof(idh), NULL);
		if (idh.e_magic == IMAGE_DOS_SIGNATURE) {
			Toolhelp32ReadProcessMemory(me.th32ProcessID, me.modBaseAddr + idh.e_lfanew, &inth, sizeof(inth), NULL);
			if (inth.Signature == IMAGE_NT_SIGNATURE) {
				pvModulePreferredBaseAddr = (PVOID)inth.OptionalHeader.ImageBase;
			}
		}
	}
	else {
		while (Module32Next(hSnapshot, &me)) {
			if (me.hModule == hModule) {
				Toolhelp32ReadProcessMemory(me.th32ProcessID, me.modBaseAddr, &idh, sizeof(idh), NULL);
				if (idh.e_magic == IMAGE_DOS_SIGNATURE) {
					Toolhelp32ReadProcessMemory(me.th32ProcessID, me.modBaseAddr + idh.e_lfanew, &inth, sizeof(inth), NULL);
					if (inth.Signature == IMAGE_NT_SIGNATURE) {
						pvModulePreferredBaseAddr = (PVOID)inth.OptionalHeader.ImageBase;
					}
				}
			}
		}
	}


	FreeLibrary(hModule);
	return (DWORD)pvModulePreferredBaseAddr;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// GetBaseAddressOf
//
// This function accepts a number and converts it to a
// string, inserting commas where appropriate.
wchar_t *CContextMenu::BigNumToString(LONG lNum, wchar_t *wcsBuffer) {

	wchar_t wcsNum[100];
	//wsprintfW(wcsNum, TEXT(L"%d"), lNum);   
	wsprintfW(wcsNum, L"%d", lNum);
	NUMBERFMTW nf;
	nf.NumDigits = 0;
	nf.LeadingZero = FALSE;
	nf.Grouping = 3;
	nf.lpDecimalSep = L".";
	nf.lpThousandSep = L",";
	nf.NegativeOrder = 0;
	GetNumberFormatW(LOCALE_USER_DEFAULT, 0, wcsNum, &nf, wcsBuffer, 100);
	return(wcsBuffer);
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// GetBaseAddressOf
//
// This function accepts a number and converts it to a
// string, inserting commas where appropriate.
wchar_t *CContextMenu::BigNumToStringF(float fNum, wchar_t *wcsBuffer) {

	wchar_t wcsNum[100];
	_snwprintf_s(wcsNum, 100, L"%f", fNum);
	NUMBERFMTW nf;
	nf.NumDigits = 1;
	nf.LeadingZero = FALSE;
	nf.Grouping = 3;
	nf.lpDecimalSep = L".";
	nf.lpThousandSep = L",";
	nf.NegativeOrder = 0;
	GetNumberFormatW(LOCALE_USER_DEFAULT, 0, wcsNum, &nf, wcsBuffer, 100);
	return(wcsBuffer);
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// CheckIfXXX
//
void CContextMenu::CheckIfXXX()
{
	//m_bIsDLL = false;
	for (unsigned int i = 0; i < m_vTargets.size(); i++) {
		DWORD dwAttr = GetFileAttributesW(m_vTargets[i].FileName);
		if (dwAttr & FILE_ATTRIBUTE_DIRECTORY) {			
			m_bIsDirectory = true;
			m_vTargets[i].IsDirectory = true;
			m_vTargets[i].IsDLL = false;
			m_vTargets[i].IsComServer = false;
			m_vTargets[i].IsTLB = false;
			m_vTargets[i].IsEXE = false;
			m_vTargets[i].IsManaged = false;
			m_bIsManaged = false;
			CAtlString str;
			str = m_vTargets[i].FileName;
			CAtlString ch;
			ch = str.Right(1);
			if (ch != L"\\") {
				wcscat_s(m_vTargets[i].FileName, 1024, L"\\");
			}
		}
		else {
			CString str;
			str = m_vTargets[i].FileName;
			str.MakeLower();
			CString strCmpIdentDLL(".dll");
			CString strCmpIdentTLB(".tlb");
			CString strCmpIdentEXE(".exe");
			CString strCmpIdentNETMODULE(".netmodule");

			CString strCmp;
			strCmp = str.Right(4);

			if (strCmp == strCmpIdentDLL) {
				m_bIsDLL = true;
				m_bIsComServer = IsDllCOMServer(m_vTargets[i].FileName);
				m_vTargets[i].IsComServer = m_bIsComServer;
				m_bIsManaged = IsManaged(m_vTargets[i].FileName);				
				m_vTargets[i].IsManaged = m_bIsManaged;
				if (m_vTargets[i].IsManaged) {
					IsManagedVersion(m_vTargets[i].FileName, &m_vTargets[i].ManagedVersion[0], 32);
				}
				m_vTargets[i].IsDLL = true;
				m_vTargets[i].IsTLB = false;
				m_vTargets[i].IsEXE = false;
				m_vTargets[i].IsNETMODULE = false;
				m_vTargets[i].IsDirectory = false;
			}
			else if (strCmp == strCmpIdentNETMODULE) {
				m_bIsNETMODULE = true;
				m_bIsDLL = false;
				m_bIsComServer = IsDllCOMServer(m_vTargets[i].FileName);
				m_vTargets[i].IsComServer = m_bIsComServer;
				m_bIsManaged = IsManaged(m_vTargets[i].FileName);
				m_vTargets[i].IsManaged = m_bIsManaged;
				if (m_vTargets[i].IsManaged) {
					IsManagedVersion(m_vTargets[i].FileName, &m_vTargets[i].ManagedVersion[0], 32);
				}
				m_vTargets[i].IsNETMODULE = true;
				m_vTargets[i].IsDLL = false;
				m_vTargets[i].IsTLB = false;
				m_vTargets[i].IsEXE = false;
				m_vTargets[i].IsDirectory = false;
			}
			else if (strCmp == strCmpIdentTLB) {
				m_bIsTLB = true;
				m_bIsComServer = false;
				m_bIsManaged = false;
				m_vTargets[i].IsManaged = false;
				m_vTargets[i].IsComServer = false;
				m_vTargets[i].IsDLL = false;
				m_vTargets[i].IsTLB = true;
				m_vTargets[i].IsEXE = false;
				m_vTargets[i].IsDirectory = false;
				m_vTargets[i].IsNETMODULE = false;
			}
			else if (strCmp == strCmpIdentEXE) {
				m_bIsEXE = true;
				m_bIsComServer = true;// must be true unless a IsExeCOMServer method is created.
				m_vTargets[i].IsComServer = true;// must be true unless a IsExeCOMServer method is created.
				m_bIsManaged = IsManaged(m_vTargets[i].FileName);
				m_vTargets[i].IsManaged = m_bIsManaged;
				if (m_vTargets[i].IsManaged) {
					IsManagedVersion(m_vTargets[i].FileName, &m_vTargets[i].ManagedVersion[0], 32);
				}
				m_vTargets[i].IsDLL = false;
				m_vTargets[i].IsTLB = false;
				m_vTargets[i].IsEXE = true;
				m_vTargets[i].IsDirectory = false;
				m_vTargets[i].IsNETMODULE = false;
			}
			else {
				m_bIsManaged = false;
				m_vTargets[i].IsManaged = false;
				m_vTargets[i].IsComServer = false;
				m_vTargets[i].IsDLL = false;
				m_vTargets[i].IsTLB = false;
				m_vTargets[i].IsEXE = false;
				m_vTargets[i].IsDirectory = false;
				m_vTargets[i].IsNETMODULE = false;
			}
		}
	}
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// ExtractPath
//
STDMETHODIMP CContextMenu::ExtractPath( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* path)
{
	if (SysStringLen(full_path) == 0 || path == NULL) {
		return E_INVALIDARG;//Error(CComBSTR(L"ExtractPath Failed. Invalid Arguments."), IID_INCFPath, E_INVALIDARG);
	}

	long lLength = SysStringLen(full_path);
	long lLastHit(-1);
	for (long l = 0; l < lLength; l++) {
		if (full_path[l] == L'\\' || full_path[l] == L'/') {
			lLastHit = l;
		}
	}

	CComBSTR bstrRetval(L"");
	if (lLastHit != -1) {
		for (long l = 0; l <= lLastHit; l++) {
			bstrRetval.Append((wchar_t)full_path[l]);
		}
		*path = bstrRetval.Detach();
	}
	else {
		*path = bstrRetval.Detach();
	}

	return S_OK;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// ExtractFileName
//
STDMETHODIMP CContextMenu::ExtractFilename( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* filename)
{
	if (SysStringLen(full_path) == 0 || filename == NULL) {
		return E_INVALIDARG;//Error(CComBSTR(L"ExtractExtension Failed. Invalid Arguments."), IID_INCFPath, E_INVALIDARG);
	}

	long lLength = SysStringLen(full_path);
	long lLastHit(-1);
	for (long l = 0; l < lLength; l++) {
		if (full_path[l] == L'\\' || full_path[l] == L'/') {
			lLastHit = l;
		}
	}

	CComBSTR bstrRetval(L"");
	if (lLastHit != -1) {
		for (long l = lLastHit + 1; l < lLength; l++) {
			bstrRetval.Append((wchar_t)full_path[l]);
		}
		*filename = bstrRetval.Detach();
	}
	else {
		bstrRetval.AppendBSTR(full_path);
		*filename = bstrRetval.Detach();
	}

	return S_OK;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// CreateMenuForTarget
//
HMENU CContextMenu::CreateMenuForTarget(STarget target, UINT idCmdFirst) 
{		
	HMENU hMenu = CreatePopupMenu();

	this->m_vMenuCommands.clear();

	int iPos(0);
	if (target.IsDirectory) {
		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_COMMANDPROMPT);
		MENUITEMINFO mmi1 = { 0 };
		mmi1.cbSize = sizeof(MENUITEMINFO);
		mmi1.fMask = MIIM_STRING | MIIM_ID;
		mmi1.fType = MFT_STRING;
		mmi1.wID = idCmdFirst++;
		mmi1.dwTypeData = reinterpret_cast<LPWSTR>(L"Command Prompt");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi1);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_COMMANDPROMPT_ADMIN);
		MENUITEMINFO mmi2 = { 0 };
		mmi2.cbSize = sizeof(MENUITEMINFO);
		mmi2.fMask = MIIM_STRING | MIIM_ID;
		mmi2.fType = MFT_STRING;
		mmi2.wID = idCmdFirst++;
		mmi2.dwTypeData = reinterpret_cast<LPWSTR>(L"Command Prompt (Admin)");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi2);

		// Insert Separator
		MENUITEMINFO mmi3 = { 0 };
		mmi3.cbSize = sizeof(MENUITEMINFO);
		mmi3.fMask = MIIM_TYPE;
		mmi3.fType = MFT_MENUBREAK;
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi3);
	}

	if (target.IsTLB) {
		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_REGISTERTYPELIB);
		MENUITEMINFO mmi1 = { 0 };
		mmi1.cbSize = sizeof(MENUITEMINFO);
		mmi1.fMask = MIIM_STRING | MIIM_ID;
		mmi1.fType = MFT_STRING;
		mmi1.wID = idCmdFirst++;
		mmi1.dwTypeData = reinterpret_cast<LPWSTR>(L"Register TypeLib");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi1);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_UNREGISTERTYPELIB);
		MENUITEMINFO mmi2 = { 0 };
		mmi2.cbSize = sizeof(MENUITEMINFO);
		mmi2.fMask = MIIM_STRING | MIIM_ID;
		mmi2.fType = MFT_STRING;
		mmi2.wID = idCmdFirst++;
		mmi2.dwTypeData = reinterpret_cast<LPWSTR>(L"Unregister TypeLib");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi2);

		// Insert Separator
		MENUITEMINFO mmi3 = { 0 };
		mmi3.cbSize = sizeof(MENUITEMINFO);
		mmi3.fMask = MIIM_TYPE;
		mmi3.fType = MFT_MENUBREAK;
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi3);
	}
	if (target.IsComServer) {
		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_REGISTERCOMSERVER);
		MENUITEMINFO mmi1 = { 0 };
		mmi1.cbSize = sizeof(MENUITEMINFO);
		mmi1.fMask = MIIM_STRING | MIIM_ID;
		mmi1.fType = MFT_STRING;
		mmi1.wID = idCmdFirst++;
		mmi1.dwTypeData = reinterpret_cast<LPWSTR>(L"Register COM Server");		
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi1);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_UNREGISTERCOMSERVER);
		MENUITEMINFO mmi2 = { 0 };
		mmi2.cbSize = sizeof(MENUITEMINFO);
		mmi2.fMask = MIIM_STRING | MIIM_ID;
		mmi2.fType = MFT_STRING;
		mmi2.wID = idCmdFirst++;
		mmi2.dwTypeData = reinterpret_cast<LPWSTR>(L"Unregister COM Server");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi2);

		// Insert Separator
		MENUITEMINFO mmi3 = { 0 };
		mmi3.cbSize = sizeof(MENUITEMINFO);
		mmi3.fMask = MIIM_TYPE;
		mmi3.fType = MFT_MENUBREAK;
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi3);
	}

	if (target.IsDLL && target.IsManaged) {
		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_INSTALLUTIL_I);
		MENUITEMINFO mmi1 = { 0 };
		mmi1.cbSize = sizeof(MENUITEMINFO);
		mmi1.fMask = MIIM_STRING | MIIM_ID;
		mmi1.fType = MFT_STRING;
		mmi1.wID = idCmdFirst++;
		mmi1.dwTypeData = reinterpret_cast<LPWSTR>(L"InstallUtil /i");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi1);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_INSTALLUTIL_U);
		MENUITEMINFO mmi2 = { 0 };
		mmi2.cbSize = sizeof(MENUITEMINFO);
		mmi2.fMask = MIIM_STRING | MIIM_ID;
		mmi2.fType = MFT_STRING;
		mmi2.wID = idCmdFirst++;
		mmi2.dwTypeData = reinterpret_cast<LPWSTR>(L"InstallUtil /u");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi2);

		// Insert Separator
		MENUITEMINFO mmi3 = { 0 };
		mmi3.cbSize = sizeof(MENUITEMINFO);
		mmi3.fMask = MIIM_TYPE;
		mmi3.fType = MFT_MENUBREAK;
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi3);
	}

	if (!target.IsDirectory) {
		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_RUN);
		MENUITEMINFO mmi = { 0 };
		mmi.cbSize = sizeof(MENUITEMINFO);
		mmi.fMask = MIIM_STRING | MIIM_ID;
		mmi.fType = MFT_STRING;
		mmi.wID = idCmdFirst++;
		mmi.dwTypeData = reinterpret_cast<LPWSTR>(L"Run Command...");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_COMMANDPROMPT);
		MENUITEMINFO mmi1 = { 0 };
		mmi1.cbSize = sizeof(MENUITEMINFO);
		mmi1.fMask = MIIM_STRING | MIIM_ID;
		mmi1.fType = MFT_STRING;
		mmi1.wID = idCmdFirst++;
		mmi1.dwTypeData = reinterpret_cast<LPWSTR>(L"Command Prompt");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi1);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_COMMANDPROMPT_ADMIN);
		MENUITEMINFO mmi2 = { 0 };
		mmi2.cbSize = sizeof(MENUITEMINFO);
		mmi2.fMask = MIIM_STRING | MIIM_ID;
		mmi2.fType = MFT_STRING;
		mmi2.wID = idCmdFirst++;
		mmi2.dwTypeData = reinterpret_cast<LPWSTR>(L"Command Prompt (Admin)");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi2);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_CREATEHARDLINKTO);
		MENUITEMINFO mmi3 = { 0 };
		mmi3.cbSize = sizeof(MENUITEMINFO);
		mmi3.fMask = MIIM_STRING | MIIM_ID;
		mmi3.fType = MFT_STRING;
		mmi3.wID = idCmdFirst++;
		mmi3.dwTypeData = reinterpret_cast<LPWSTR>(L"Create Hard Link...");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi3);

		// Insert Separator
		MENUITEMINFO mmi4 = { 0 };
		mmi4.cbSize = sizeof(MENUITEMINFO);
		mmi4.fMask = MIIM_TYPE;
		mmi4.fType = MFT_MENUBREAK;
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi4);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_ENCRYPT);
		MENUITEMINFO mmi5 = { 0 };
		mmi5.cbSize = sizeof(MENUITEMINFO);
		mmi5.fMask = MIIM_STRING | MIIM_ID;
		mmi5.fType = MFT_STRING;
		mmi5.wID = idCmdFirst++;
		mmi5.dwTypeData = reinterpret_cast<LPWSTR>(L"Encrypt...");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi5);

		this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_DECRYPT);
		MENUITEMINFO mmi6 = { 0 };
		mmi6.cbSize = sizeof(MENUITEMINFO);
		mmi6.fMask = MIIM_STRING | MIIM_ID;
		mmi6.fType = MFT_STRING;
		mmi6.wID = idCmdFirst++;
		mmi6.dwTypeData = reinterpret_cast<LPWSTR>(L"Decrypt...");
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi6);

		// Insert Separator
		MENUITEMINFO mmi7 = { 0 };
		mmi7.cbSize = sizeof(MENUITEMINFO);
		mmi7.fMask = MIIM_TYPE;
		mmi7.fType = MFT_MENUBREAK;
		InsertMenuItem(hMenu, iPos++, TRUE, &mmi7);
	}	

	this->m_vMenuCommands.push_back(ID_DEVELOPEREXTENSION_ABOUT);
	MENUITEMINFO mmiAbout = { 0 };
	mmiAbout.cbSize = sizeof(MENUITEMINFO);
	mmiAbout.fMask = MIIM_STRING | MIIM_ID;
	mmiAbout.fType = MFT_STRING;
	mmiAbout.wID = idCmdFirst++;
	mmiAbout.dwTypeData = reinterpret_cast<LPWSTR>(L"About Developer Extensions");
	InsertMenuItem(hMenu, iPos++, TRUE, &mmiAbout);

	
	return hMenu;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// GetCommonTarget
//
STarget CContextMenu::GetCommonTarget()
{

	STarget target = this->m_vTargets[0];

	for (auto itr = this->m_vTargets.begin(); itr < this->m_vTargets.end(); itr++) 
	{		
		if (itr->IsComServer) {
			if (!target.IsComServer) {
				target.IsComServer = false;
			}
		}
		else if (!itr->IsComServer) {
			if (target.IsComServer) {
				target.IsComServer = false;
			}
		}

		if (itr->IsTLB) {
			if (!target.IsTLB) {
				target.IsTLB = false;
			}
		}
		else if (!itr->IsTLB) {
			if (target.IsTLB) {
				target.IsTLB = false;
			}
		}
		
		if (itr->IsDLL) {
			if (!target.IsDLL) {
				target.IsDLL = false;
			}
		}
		else if (!itr->IsDLL) {
			if (target.IsDLL) {
				target.IsDLL = false;
			}
		}
		
		if (itr->IsNETMODULE) {
			if (!target.IsNETMODULE) {
				target.IsNETMODULE = false;
			}
		}
		else if (!itr->IsNETMODULE) {
			if (target.IsNETMODULE) {
				target.IsNETMODULE = false;
			}
		}

		if (itr->IsEXE) {
			if (!target.IsEXE) {
				target.IsEXE = false;
			}
		}
		else if (!itr->IsEXE) {
			if (target.IsEXE) {
				target.IsEXE = false;
			}
		}

		if (itr->IsDirectory) {
			if (!target.IsDirectory) {
				target.IsDirectory = false;
			}
		}
		else if (!itr->IsDirectory) {
			if (target.IsDirectory) {
				target.IsDirectory = false;
			}
		}

		if (itr->IsManaged) {
			if (!target.IsManaged) {
				target.IsManaged = false;
			}
		}
		else if (!itr->IsManaged) {
			if (target.IsManaged) {
				target.IsManaged = false;
			}
		}
	}


	return target;
}
////////////////////////////////////////////////////////////////////////////////////////////////
//
// GetNETFolder
//
bool CContextMenu::GetNETFolder(wchar_t* version, wchar_t* folder, rsize_t size)
{	
	HKEY hKeyFramework;
	LSTATUS status = RegOpenKeyW(HKEY_LOCAL_MACHINE, L"SOFTWARE\\Microsoft\\.NETFramework", &hKeyFramework);	
	if (status == ERROR_SUCCESS) {
		BYTE* pData = new BYTE[255];
		DWORD cbData(255);				
		status = RegQueryValueEx(hKeyFramework, L"InstallRoot", 0, NULL, pData, &cbData);		
		if (status == ERROR_SUCCESS) {
			wcscpy_s(folder, size, reinterpret_cast<wchar_t*>(pData));
			
			DWORD dwIndex(0);
			wchar_t name[512];
			vector<wstring> versions;
			status = RegEnumKey(hKeyFramework, dwIndex++, name, 512);
			do {
				int length = static_cast<int>(wcslen(name));
				
				if (name[0] == L'v' && isdigit(name[length-1]) ) 
				{
					versions.push_back(wstring(name));
				}
				status = RegEnumKey(hKeyFramework, dwIndex++, name, 512);
			} while (status == ERROR_SUCCESS);

			if (versions.size() == 0) {
				return false;
			}

			bool bHit(false);
			for (auto str : versions) {
				if (str == version) {
					wcscat_s(folder, size, str.c_str());
					bHit = true;
					break;
				}
			}

			if (!bHit) {
				wcscat_s(folder, size, versions.back().c_str());
			}			

			return true;
		}		
		delete[] pData;
	}

	return false;
}

////////////////////////////////////////////////////////////////////////////////////////
//
// RunCommandDialogProc
//	
INT_PTR CALLBACK RunCommandDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{

	switch (uMsg)
	{
	case WM_INITDIALOG:
		ItsWin::CenterToItsParentWindow(hwndDlg);
		SendMessage(GetDlgItem(hwndDlg, IDC_STATICICON), STM_SETICON, (WPARAM)g_hRunIcon, 0);
		SetDlgItemTextW(hwndDlg, IDC_WORKINGFOLDER, g_bstrWorkingFolder);
		SetDlgItemTextW(hwndDlg, IDC_COMMAND, L"[command] %file%");
		return TRUE;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{					
		case IDOK:
		{
			wchar_t wcsWorkingFolder[512];
			wchar_t wcsCommand[1024];

			::GetDlgItemTextW(hwndDlg, IDC_WORKINGFOLDER, wcsWorkingFolder, 512);
			::GetDlgItemTextW(hwndDlg, IDC_COMMAND, wcsCommand, 1024);						
			BOOL isElevated = ::IsDlgButtonChecked(hwndDlg, IDC_CHECK_ELEVATED);

			CString atlStrCommand = wcsCommand;
			atlStrCommand = atlStrCommand.Trim();
			if (atlStrCommand.GetLength() == 0) {
				break;
			}

			if (wcslen(wcsWorkingFolder) > 0) {
				if (wcsWorkingFolder[wcslen(wcsWorkingFolder) - 1] != L'\\') {
					wcscat_s(wcsWorkingFolder, 512, L"\\");
				}
			}

			int i = 0;
			for (auto target : *g_pvTargets) 
			{			
				i++;

				CComBSTR bstrFileName;
				g_pCContextMenu->ExtractFilename(CComBSTR(target.FileName), &bstrFileName);

				wstring wcsFileToken(L"%file%");								
				wstring strCommand = atlStrCommand.operator LPCWSTR();				
				if (strCommand == L"%file%" && wcslen(wcsWorkingFolder) > 0) {					
					SHELLEXECUTEINFOW execInfo = { 0 };
					execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
					execInfo.lpVerb = (isElevated == TRUE) ? L"RunAs" : L"Open";
					execInfo.nShow = SW_NORMAL;
					execInfo.lpFile = target.FileName;
					if (wcslen(wcsWorkingFolder) > 0) {
						execInfo.lpDirectory = wcsWorkingFolder;
					}					
					::ShellExecuteExW(&execInfo);
					break;
				}
				

				size_t index = strCommand.find(L"%file%");
				while ( index != std::wstring::npos )
				{
					bstrFileName.ToLower( );
					strCommand.replace( index, 6, bstrFileName.operator LPWSTR( ) );
					index = strCommand.find( L"%file%" );					
				}

				wstring systemCommand = strCommand;

				wstring strParameters;
				index = strCommand.find_first_of(L' ');
				if (index != std::string::npos) {
					strParameters = strCommand.substr(index);
					strCommand = strCommand.substr(0, index);
				}

				bool bFound = false;
				wstring pathAndCommand;				 
				wstring pathSystemCommand;
				if (wcslen(wcsWorkingFolder) > 0) {
					pathAndCommand = wcsWorkingFolder + strCommand;										
					pathSystemCommand = wcsWorkingFolder + systemCommand;

					wstring temp;
					if (pathAndCommand.rfind(L".exe") == std::string::npos) {
						temp = pathAndCommand + L".exe";
						
						if (::PathFileExistsW(temp.c_str()))
						{
							bFound = true;
						}
					}
					
					if (pathAndCommand.rfind(L".bat") == std::string::npos) 
					{
						temp = pathAndCommand + L".bat";

						if (::PathFileExistsW(temp.c_str()))
						{
							bFound = true;
						}
					}
					
					if (!bFound) {
						temp = pathAndCommand;

						if (::PathFileExistsW(temp.c_str()))
						{
							bFound = true;
						}
					}
				}

				bool bHasRunAsSystem = false;

				if (!bFound) {
					if (!g_pCContextMenu->FindPathForCommand(strCommand, pathAndCommand))
					{
						wstring working_folder(wcsWorkingFolder);
						string cmd_working_folder(working_folder.begin(), working_folder.end());

						string cmd1(working_folder.begin(), working_folder.begin() + 2);
						string cmd2("cd ");
						cmd2 += cmd_working_folder;
						string cmd3(systemCommand.begin(), systemCommand.end());

						unique_ptr<char[]> pPath = make_unique<char[]>(MAX_PATH);

						stringstream ss;
						ss << "devext_" << i << "run.bat";
						string strFN = ss.str();

						GetTempPathA(MAX_PATH, pPath.get());
						strcat_s(pPath.get(), MAX_PATH, strFN.c_str());

						string sys_path(pPath.get());
						
						ItsTextFile file;
						if (!file.OpenOrCreateText(wstring(sys_path.begin(), sys_path.end()), L"w", L"rw", ItsFileOpenCreation::CreateAlways, ItsFileTextType::Ansi)) {
							MessageBox(GetActiveWindow(), L"Could not find command in %PATH%.", L"Developer Extensions - Error", MB_OK | MB_ICONERROR);
							break;
						}
						
						file.WriteTextLine("@ECHO OFF");
						file.WriteTextLine(cmd1);
						file.WriteTextLine(cmd2);
						file.WriteTextLine(cmd3);
						file.WriteTextLine("pause");

						file.Close();

						SHELLEXECUTEINFOW execInfo = { 0 };
						execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
						execInfo.lpVerb = (isElevated == TRUE) ? L"RunAs" : L"Open";
						execInfo.nShow = SW_NORMAL;
						execInfo.lpFile = wstring(sys_path.begin(),sys_path.end()).c_str();
						if (wcslen(wcsWorkingFolder) > 0) {
							execInfo.lpDirectory = wcsWorkingFolder;
						}
						execInfo.lpParameters = NULL;
						::ShellExecuteExW(&execInfo);
						//if (system(sys_path.c_str()) > 0) {
						//	MessageBox(GetActiveWindow(), L"Could not find command in %PATH%.", L"Developer Extensions - Error", MB_OK | MB_ICONERROR);
						//	break;
						//}						
						bHasRunAsSystem = true;
					}
				}

				if (!bHasRunAsSystem)
				{
					SHELLEXECUTEINFOW execInfo = { 0 };
					execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
					execInfo.lpVerb = (isElevated == TRUE) ? L"RunAs" : L"Open";
					execInfo.nShow = SW_NORMAL;
					execInfo.lpFile = pathAndCommand.c_str();
					if (wcslen(wcsWorkingFolder) > 0) {
						execInfo.lpDirectory = wcsWorkingFolder;
					}
					execInfo.lpParameters = strParameters.c_str();
					::ShellExecuteExW(&execInfo);
				}
			}

			EndDialog(hwndDlg, 0);
			return TRUE;
			break;
		}

		case IDCANCEL:
			EndDialog(hwndDlg, 0);
			return TRUE;
			break;

		case IDC_BROWSE:
		{
			wchar_t wcsBuffer[MAX_PATH];
			LPITEMIDLIST pidlSelected = NULL;
			BROWSEINFOW bi = { 0 };
			bi.hwndOwner = hwndDlg;
			bi.pidlRoot = NULL;
			bi.pszDisplayName = wcsBuffer;
			bi.lpszTitle = L"Choose a working directory.";
			bi.ulFlags = BIF_USENEWUI | BIF_NEWDIALOGSTYLE | BIF_RETURNONLYFSDIRS;
			bi.lpfn = NULL;
			bi.lParam = 0;

			pidlSelected = SHBrowseForFolderW(&bi);
			if (pidlSelected != NULL) {
				::SHGetPathFromIDListW(pidlSelected, wcsBuffer);
				if (wcsBuffer[wcslen(wcsBuffer) - 1] != L'\\') {
					wcscat_s(wcsBuffer, MAX_PATH, L"\\");
				}
				::SetDlgItemTextW(hwndDlg, IDC_WORKINGFOLDER, wcsBuffer);
			}
			break;
		}
		default:
			break;
		}
		break;

	default:
		break;
	}

	return FALSE;
}


DWORD CContextMenu::GetActualAddressFromRVA(IMAGE_SECTION_HEADER* pSectionHeader, IMAGE_NT_HEADERS* pNTHeaders, DWORD dwRVA)
{
	DWORD dwRet = 0;
	
	for (int j = 0; j < pNTHeaders->FileHeader.NumberOfSections; j++, pSectionHeader++)
	{
		DWORD cbMaxOnDisk = min(pSectionHeader->Misc.VirtualSize, pSectionHeader->SizeOfRawData);
		
		DWORD startSectRVA, endSectRVA;
		
		startSectRVA = pSectionHeader->VirtualAddress;
		endSectRVA = startSectRVA + cbMaxOnDisk;
		
		if ((dwRVA >= startSectRVA) && (dwRVA < endSectRVA))
		{
			dwRet = (pSectionHeader->PointerToRawData) + (dwRVA - startSectRVA);
			break;
		}
		
	}
	
	return dwRet;
}

bool CContextMenu::FindPathForCommand(wstring command, wstring& outputPathAndCommand) {
	DWORD dwSize = 16 * 1024;
	wchar_t* buffer = new wchar_t[dwSize];
	DWORD dwResult = ::GetEnvironmentVariableW(L"PATH", buffer, dwSize);
	if (dwResult == 0) {
		delete[] buffer;
		return false;
	}
	
	wstring temp(buffer);

	delete[] buffer;

	if (command.rfind(L".exe") == std::wstring::npos)
	{
		command += L".exe";
	}

	int index = 0;
	int indexLast = 0;
	
	while ((index = static_cast<int>(temp.find(L';', indexLast))) != std::wstring::npos)
	{
		wstring path = temp.substr(indexLast, index - indexLast);
		indexLast = index + 1;

		if (path[path.length() - 1] != L'\\') {
			path += L'\\';
		}

		wstring pathAndCommand = path + command;

		if (::PathFileExistsW(pathAndCommand.c_str()) )
		{
			outputPathAndCommand = pathAndCommand;
			return true;
		}
	}

	return false;
}

