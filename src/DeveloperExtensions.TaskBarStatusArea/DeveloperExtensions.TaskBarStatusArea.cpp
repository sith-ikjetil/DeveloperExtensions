//
// Include files
//
#include "stdafx.h"
#include <windows.h>
#include "resource.h"
#include <shellapi.h>
#include "commctrl.h"
#include <vector>
#include <atlbase.h>
#include <comadmin.h>
#include <process.h>
#include <stdlib.h>
#include <ShellScalingApi.h>
#include <VersionHelpers.h>

//
// using namespace
//
using namespace std;

// Data structure
struct COMPAPPLICATIONS
{
	UINT	m_uiMenuID;
	wchar_t m_wcsName[255];
};

// Global data
long const		g_lStatusAreaID = WM_USER + 1;
long const		g_lCallbackMessageID = WM_USER + 1;
long const		g_lShutdownApplicationID = WM_USER + 0x00001000;
long			g_lShutdownIndex = 0;	// Shutdown index
bool g_bAutoRun = false;
wchar_t const	g_wcsApplicationName[] = L"Developer Extensions";

HINSTANCE		g_hInstance(NULL);
NOTIFYICONDATA	g_notifyicondata = { 0 };
HMENU           g_hCoMenu(NULL);
HWND			g_hwndDlgShutdownApplication(NULL);
HWND			g_hwndDlgShutdownAllApplication(NULL);
bool g_bCancelShutdown = false;
vector<COMPAPPLICATIONS> g_vCoApplications;
HMODULE			g_hModuleDll(NULL);	// Handle to devext.dll

									// Function prototypes
LRESULT CALLBACK StatusAreaProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam); // Window function.
void BuildCOMPlusServerApplicationsMenu(HMENU& hSubMenu); // Util function to put COM+ server applications into the menu...
bool GetAllCoPlusServerApplications(); // Retrieve COM+ Server applications and add them to the g_vCoApplications vector.
bool ShutDownApplication(wchar_t* wcsApplicationName);
INT_PTR CALLBACK ShutdownApplicationProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
VOID DoShutdownApplication(PVOID pvoid);	// Thread function.
INT_PTR CALLBACK ShutdownAllApplicationsProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
VOID DoShutdownAllApplications(PVOID pvoid);	// Thread function.
bool StopIIS(PROCESS_INFORMATION* pi);
bool StartIIS(PROCESS_INFORMATION* pi);
//bool ResetIIS();
VOID DoRestartIIS(PVOID pvoid);
bool RefreshAllConfiguredComponents();
void AddStatusAreaIcon();
STDMETHODIMP ExtractPath( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* path);
HMENU CreateMainMenu();

//__stdcall DevExtShowAboutDlg(HINSTANCE hInstance)
typedef void (WINAPI* DevExtShowAboutDlgType)(void);

// wWinMain
int WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nShowCmd)
{

	wchar_t buffer[1024];
	GetModuleFileNameW(g_hInstance, buffer, 1024);

	CComBSTR bstrFullPath(buffer);
	CComBSTR bstrPathOnly;
	ExtractPath(bstrFullPath, &bstrPathOnly);

	CComBSTR bstrLibrary(bstrPathOnly);
	if (bstrLibrary.operator LPWSTR()[bstrLibrary.Length() - 1] != L'\\') {
		bstrLibrary.Append(L"\\");
	}
	bstrLibrary.Append(L"DeveloperExtensions.Common.dll");
	g_hModuleDll = LoadLibraryW(bstrLibrary.operator LPWSTR());
	if (g_hModuleDll == NULL) {
		MessageBox(GetActiveWindow(), L"Could not load library \"DeveloperExtensions.Common.dll\".", L"Error", MB_OK | MB_ICONERROR);
		return 0;
	}

	::CoInitialize(NULL);
	HINSTANCE g_hInstance = hInstance;

	HICON hIcon = LoadIcon(hInstance, MAKEINTRESOURCE(IDI_STATUSAREA_MAIN));

	// Setup up window class and create window.
	WNDCLASSEX wndclassex = { 0 };
	wndclassex.cbSize = sizeof(WNDCLASSEX);
	wndclassex.cbClsExtra = 0;
	wndclassex.cbWndExtra = 0;
	wndclassex.hbrBackground = (HBRUSH)GetStockObject(WHITE_BRUSH);
	wndclassex.hCursor = NULL;
	wndclassex.hIcon = NULL;
	wndclassex.hIconSm = NULL;
	wndclassex.hInstance = hInstance;
	wndclassex.lpfnWndProc = (WNDPROC)StatusAreaProc;
	wndclassex.lpszClassName = g_wcsApplicationName;
	wndclassex.lpszMenuName = NULL;
	wndclassex.style = 0;

	RegisterClassEx(&wndclassex);

	HWND hWnd = CreateWindowEx(0, g_wcsApplicationName, g_wcsApplicationName, 0, 0, 0, 0, 0, NULL, NULL, hInstance, (LPVOID)hInstance);

	// Show the window hidden.
	ShowWindow(hWnd, SW_HIDE);
	UpdateWindow(hWnd);

	// Set up the notifyicondata structure...
	g_notifyicondata.cbSize = sizeof(NOTIFYICONDATA);
	g_notifyicondata.hIcon = hIcon;
	g_notifyicondata.hWnd = hWnd;
	wcscpy_s(g_notifyicondata.szTip, 128, g_wcsApplicationName);
	g_notifyicondata.uCallbackMessage = g_lCallbackMessageID;
	g_notifyicondata.uFlags = NIF_MESSAGE | NIF_ICON | NIF_TIP;
	g_notifyicondata.uVersion = NOTIFYICON_VERSION;
	g_notifyicondata.dwState = NIS_SHAREDICON;
	g_notifyicondata.dwStateMask = NIS_SHAREDICON;
	g_notifyicondata.uID = g_lStatusAreaID;

	// Add the tastkabar status area icon to the status area.
	BOOL b(FALSE);
	b = ::Shell_NotifyIcon(NIM_ADD, &g_notifyicondata);
	b = ::Shell_NotifyIcon(NIM_SETVERSION, &g_notifyicondata);

	GetAllCoPlusServerApplications();
	// Set up the message loop.
	MSG msg;
	while (GetMessage(&msg, NULL, 0, 0)) {
		TranslateMessage(&msg);
		DispatchMessage(&msg);
	}

	b = ::Shell_NotifyIcon(NIM_DELETE, &g_notifyicondata);

	::CoUninitialize();

	FreeLibrary(g_hModuleDll);
	return static_cast<int>(msg.wParam);
}// int WINAPI wWinMain( HINSTANCE hInstance, HINSTANCE hPrevInstance, LPWSTR lpCmdLine, int nShowCmd )

 //////////////////////////////////////////////////////////////////////////////////////
 //
 // Function:	AddStatusAreaIcon
 //
 // Description: Adds the notify icon to the status area.
 //
void AddStatusAreaIcon() {
	Shell_NotifyIcon(NIM_ADD, &g_notifyicondata);
	Shell_NotifyIcon(NIM_SETVERSION, &g_notifyicondata);
}// AddStatusAreaIcon

 // StatusAreaProc. Window function for the status area icon window.
LRESULT CALLBACK StatusAreaProc(HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam)
{
	static HMENU hMenu(NULL);
	static HMENU hSubMenu(NULL);
	static UINT s_uTaskbarRestart;

	switch (iMsg)
	{
	case WM_CREATE:
		s_uTaskbarRestart = RegisterWindowMessage(L"TaskbarCreated");
		return 0;
		break;
	case WM_DESTROY:
		if (hMenu != NULL) {
			::DestroyMenu(::g_hCoMenu);
			::DestroyMenu(hMenu);
			::g_hCoMenu = NULL;
			hMenu = NULL;
			hSubMenu = NULL;
		}
		::PostQuitMessage(0);
		return 0;
		break;
	case g_lCallbackMessageID:
	{
		switch (lParam)
		{
		case WM_RBUTTONDOWN:
			if (hMenu != NULL) {
				::SendMessage(hWnd, WM_CANCELMODE, 0, 0);
				::DestroyMenu(hMenu);
				hMenu = NULL;
				hSubMenu = NULL;
			}
			return 0;
			break;
		case WM_CONTEXTMENU:
		{
			POINT pt;
			::GetCursorPos(&pt);

			hMenu = ::LoadMenu(g_hInstance, MAKEINTRESOURCE(IDR_MENUMAIN));
			hSubMenu = ::GetSubMenu(hMenu, 0);
			BuildCOMPlusServerApplicationsMenu(hSubMenu);
			SetForegroundWindow(hWnd);
			DWORD dwItem = ::TrackPopupMenuEx(hSubMenu, TPM_RIGHTALIGN | TPM_BOTTOMALIGN | TPM_RETURNCMD, pt.x, pt.y, hWnd, NULL);
			switch (dwItem)
			{
				// Stop IIS
			case ID_DEVELOPEREXTENSIONS_STOPIIS:
			{
				PROCESS_INFORMATION pi = { 0 };
				if (!StopIIS(&pi)) {
					MessageBox(GetActiveWindow(), L"Failed to stop IIS. Failed to execute \"net stop iisadmin /y\".", L"Error", MB_OK | MB_ICONERROR);
				}
				return 0;
				break;
			}
			// Start IIS
			case ID_DEVELOPEREXTENSIONS_STARTISS:
			{
				PROCESS_INFORMATION pi = { 0 };
				if (!StartIIS(&pi)) {
					MessageBox(GetActiveWindow(), L"Failed to start IIS. Failed to execute \"net start w3svc\".", L"Error", MB_OK | MB_ICONERROR);
				}
				return 0;
				break;
			}
			// Restart IIS
			case ID_DEVELOPEREXTENSIONS_RESTARTIIS:
			{
				_beginthread(DoRestartIIS, 0, NULL);
				return 0;
				break;
			}
			// Shutdown COM+ Application
			case ID_DEVELOPEREXTENSIONS_SHUTDOWNALLCOMAPPLICATIONS:
			{
				::DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_SHUTDOWNAPPLICATIONS), hWnd, ShutdownAllApplicationsProc);
				return 0;
				break;
			}
			// Refresh All Configured Components
			case ID_DEVELOPEREXTENSIONS_REFRESHALLCOMCOMPONENTS:
			{
				if (!RefreshAllConfiguredComponents()) {
					MessageBox(GetActiveWindow(), L"Failure to refresh all configured components.", L"Error", MB_OK | MB_ICONERROR);
				}
				else {
					MessageBox(GetActiveWindow(), L"All configured components refreshed.", L"Success", MB_OK | MB_ICONINFORMATION);
				}
				return 0;
				break;
			}
			// stop iis, shutdown all com+ server applications
			case ID_DEVELOPEREXTENSIONS_STOPIISSHUTDOWNALLCOMSERVERAPPLICATIONS:
			{
				PROCESS_INFORMATION pi = { 0 };
				if (!StopIIS(&pi)) {
					MessageBox(GetActiveWindow(), L"Failed to stop IIS. Failed to execute \"net stop iisadmin /y\".", L"Error", MB_OK | MB_ICONERROR);
					return 0;
					break;
				}
				WaitForSingleObject(pi.hProcess, INFINITE);
				g_bAutoRun = true;
				DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_SHUTDOWNAPPLICATIONS), hWnd, ShutdownAllApplicationsProc);
				g_bAutoRun = false;

				MessageBox(GetActiveWindow(), L"Finished successfully.", L"Success", MB_OK | MB_ICONINFORMATION);
				return 0;
				break;
			}
			// stop IIS, Shutdown All COM+ Server Applications, start iis
			case ID_DEVELOPEREXTENSIONS_RESTARTIISANDSHUTDOWNALLCOMAPPLICATIONS:
			{
				PROCESS_INFORMATION pi = { 0 };
				if (!StopIIS(&pi)) {
					MessageBox(GetActiveWindow(), L"Failed to stop IIS. Failed to execute \"net stop iisadmin /y\".", L"Error", MB_OK | MB_ICONERROR);
					return 0;
					break;
				}
				WaitForSingleObject(pi.hProcess, INFINITE);
				g_bAutoRun = true;
				DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_SHUTDOWNAPPLICATIONS), hWnd, ShutdownAllApplicationsProc);
				g_bAutoRun = false;

				/*	if ( !RefreshAllConfiguredComponents() ) {
				MessageBox( GetActiveWindow(), L"Failure to refresh all configured components.", L"Error", MB_OK | MB_ICONERROR );
				return 0;
				break;
				}
				*/
				//	else {
				//		MessageBox( GetActiveWindow(), L"All configured components refreshed.", L"Error", MB_OK | MB_ICONINFORMATION );
				//	}
				memset(&pi, 0, sizeof(PROCESS_INFORMATION));
				if (!StartIIS(&pi)) {
					MessageBox(GetActiveWindow(), L"Failed to start IIS. Failed to execute \"net start w3svc\".", L"Error", MB_OK | MB_ICONERROR);
					return 0;
					break;
				}
				WaitForSingleObject(pi.hProcess, INFINITE);
				MessageBox(GetActiveWindow(), L"IIS was stopped, all COM+ server applications were shutdown and IIS was started.", L"Success", MB_OK | MB_ICONINFORMATION);
				return 0;
				break;
			}
			// Exit
			case ID_DEVELOPEREXTENSIONS_EXIT:
			{
				::Shell_NotifyIcon(NIM_DELETE, &g_notifyicondata);
				::PostQuitMessage(0);
				return 0;
				break;
			}
			// About
			case ID_DEVELOPEREXTENSIONS_ABOUTDEVELOPEREXTENSIONS:
			{
				DevExtShowAboutDlgType fnShowAboutDlg = (DevExtShowAboutDlgType)::GetProcAddress(g_hModuleDll, "DevExtShowAboutDlg");
				if (fnShowAboutDlg != NULL) {
					try {
						fnShowAboutDlg();
					}
					catch (...) {
						CComBSTR bstr("Exception caught while trying to call DevExtShowAboutDlg in devext.dll");
						MessageBox(GetActiveWindow(), bstr, L"Error", MB_OK | MB_ICONERROR);
					}
				}
				else {
					MessageBox(GetActiveWindow(), L"Could not get address of exported function DevExtShowAboutDlg in devext.dll", L"Error", MB_OK | MB_ICONERROR);
				}
				return 0;
				break;
			}
			default:
				break;
			}; // switch ( dwItem )

			   // Do we have a shutdown application type menu buildup item...
			if (dwItem >= g_lShutdownApplicationID && (dwItem <= g_lShutdownApplicationID + ::g_vCoApplications.size())) {
				for (long l = 0; l < ::g_vCoApplications.size(); l++) {
					if (::g_vCoApplications[l].m_uiMenuID == dwItem) {
						g_lShutdownIndex = l;
						::DialogBox(g_hInstance, MAKEINTRESOURCE(IDD_SHUTDOWNAPPLICATIONS), hWnd, ShutdownApplicationProc);
						g_lShutdownIndex = 0;
						break;
					}
				}
			}

			::DestroyMenu(::g_hCoMenu);
			::DestroyMenu(hMenu);
			::g_hCoMenu = NULL;
			hMenu = NULL;
			hSubMenu = NULL;
			return 0;
			break;
		}
		default:
			break;
		};
	}
	return 0;
	break;
	default:
		if (iMsg == s_uTaskbarRestart) {
			AddStatusAreaIcon();
		}
		break;
	};


	return DefWindowProc(hWnd, iMsg, wParam, lParam);
}// LRESULT CALLBACK StatusAreaProc( HWND hWnd, UINT iMsg, WPARAM wParam, LPARAM lParam )

void BuildCOMPlusServerApplicationsMenu(HMENU& hSubMenu)
{
	// Get all COM+ Server applications...
	if (!GetAllCoPlusServerApplications()) {
		// gray out the menu item ....
		return;
	}

	// Create the menu
	::g_hCoMenu = CreateMenu();
	for (long l = 0; l < ::g_vCoApplications.size(); l++) {
		MENUITEMINFO mii = { 0 };
		mii.cbSize = sizeof(MENUITEMINFO);
		mii.fMask = MIIM_STRING | MIIM_ID | MIIM_DATA;
		mii.dwTypeData = ::g_vCoApplications[l].m_wcsName;
		mii.wID = ::g_vCoApplications[l].m_uiMenuID;
		mii.dwItemData = ::g_vCoApplications[l].m_uiMenuID;
		InsertMenuItem(g_hCoMenu, l, TRUE, &mii);
	}

	// Insert the menu into the std menu...
	MENUITEMINFO mii = { 0 };
	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STRING;
	GetMenuItemInfo(hSubMenu, ID_DEVELOPEREXTENSIONS_SHUTDOWNCOMPLUSAPPLICATION, FALSE, &mii);

	mii.cbSize = sizeof(MENUITEMINFO);
	mii.fMask = MIIM_STRING | MIIM_SUBMENU;
	mii.hSubMenu = ::g_hCoMenu;
	mii.dwTypeData = L"Shutdown COM+ Server Application";

	SetMenuItemInfo(hSubMenu, ID_DEVELOPEREXTENSIONS_SHUTDOWNCOMPLUSAPPLICATION, FALSE, &mii);
}

bool GetAllCoPlusServerApplications()
{
	if (g_vCoApplications.size() > 0)
		g_vCoApplications.clear();

	CComPtr<ICOMAdminCatalog> pAdminCatalog;
	HRESULT hr = pAdminCatalog.CoCreateInstance(CComBSTR("COMAdmin.COMAdminCatalog.1"));//CLSID_COMAdminCatalog );
	if (FAILED(hr)) {
		MessageBox(GetActiveWindow(), L"Failed to create the COMAdminCatalog.", L"Error", MB_OK | MB_ICONERROR);
		return false;
	}

	CComPtr<IDispatch> pDispatch;
	hr = pAdminCatalog->GetCollection(CComBSTR("Applications"), &pDispatch);
	if (FAILED(hr)) {
		MessageBox(GetActiveWindow(), L"Failed to retrieve the collection.", L"Error", MB_OK | MB_ICONERROR);
		return false;
	}

	CComQIPtr<ICatalogCollection> pCollection;
	pCollection = pDispatch;

	pCollection->Populate();

	long lCount(0);
	hr = pCollection->get_Count(&lCount);

	long lApplications(0);
	for (long l = 0; l < lCount; l++) {
		CComPtr<IDispatch> pDispatch;
		hr = pCollection->get_Item(l, &pDispatch);
		if (FAILED(hr)) {
			MessageBox(GetActiveWindow(), L"Failed to get catalog-object.", L"Error", MB_OK | MB_ICONERROR);
			return false;
		}

		CComQIPtr<ICatalogObject> pObject;
		pObject = pDispatch;

		CComVariant vtValue;
		hr = pObject->get_Value(CComBSTR("Activation"), &vtValue);

		if (vtValue.lVal == 1) {	// Is server Application....
			CComVariant vtName;
			hr = pObject->get_Name(&vtName);
			if (FAILED(hr)) {
				MessageBox(GetActiveWindow(), L"Failed to get catalog-object name.", L"Error", MB_OK | MB_ICONERROR);
				return false;
			}

			COMPAPPLICATIONS ca;
			ca.m_uiMenuID = g_lShutdownApplicationID + lApplications;
			lApplications++;

			wcscpy_s(ca.m_wcsName, 255, vtName.bstrVal);
			::g_vCoApplications.push_back(ca);
		}
	}
	return true;
}

bool ShutDownApplication(wchar_t* wcsApplicationName)
{
	CComPtr<ICOMAdminCatalog> pAdminCatalog;
	HRESULT hr = pAdminCatalog.CoCreateInstance(CComBSTR("COMAdmin.COMAdminCatalog.1"));//CLSID_COMAdminCatalog );
	if (FAILED(hr))
		return false;

	CComBSTR bstrApplicationName;
	bstrApplicationName += wcsApplicationName;

	hr = pAdminCatalog->ShutdownApplication(bstrApplicationName);
	if (FAILED(hr))
		return false;

	return true;
}

INT_PTR CALLBACK ShutdownApplicationProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		g_hwndDlgShutdownApplication = hwndDlg;
		EnableWindow(GetDlgItem(hwndDlg, IDOK), FALSE);
		_beginthread(DoShutdownApplication, 0, NULL);
		return TRUE;
		break;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDOK:
			EndDialog(hwndDlg, 0);
			g_hwndDlgShutdownApplication = NULL;
			return TRUE;
			break;
		default:
			break;
		}
		break;
	};
	return FALSE;
}

VOID DoShutdownApplication(PVOID pvoid)
{
	::CoInitialize(NULL);
	SendMessage(GetDlgItem(g_hwndDlgShutdownApplication, IDC_PROGRESS2), PBM_SETRANGE, 0, MAKELPARAM(0, 2));
	SendMessage(GetDlgItem(g_hwndDlgShutdownApplication, IDC_PROGRESS2), PBM_SETPOS, 1, 0);

	HWND hwndList = GetDlgItem(g_hwndDlgShutdownApplication, IDC_LISTAPPLICATIONS);
	SendMessage(hwndList, LB_ADDSTRING, 0, (LPARAM)g_vCoApplications[g_lShutdownIndex].m_wcsName);
	SendMessage(hwndList, LB_SETSEL, 1, 0);
	if (ShutDownApplication(g_vCoApplications[g_lShutdownIndex].m_wcsName)) {
		SendMessage(hwndList, LB_SETSEL, 0, 0);
	}

	SendMessage(GetDlgItem(g_hwndDlgShutdownApplication, IDC_PROGRESS2), PBM_SETPOS, 2, 0);
	EnableWindow(GetDlgItem(g_hwndDlgShutdownApplication, IDOK), TRUE);
	::CoUninitialize();
}


INT_PTR CALLBACK ShutdownAllApplicationsProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
		g_hwndDlgShutdownAllApplication = hwndDlg;
		SetDlgItemText(hwndDlg, IDOK, L"Cancel");
		_beginthread(DoShutdownAllApplications, 0, NULL);
		return TRUE;
		break;
	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDOK:
			if (g_hwndDlgShutdownAllApplication != NULL) {
				g_bCancelShutdown = true;
			}
			else {
				EndDialog(hwndDlg, 0);
				g_hwndDlgShutdownAllApplication = NULL;
			}
			return TRUE;
			break;
		default:
			break;
		}
		break;
	};
	return FALSE;
}

VOID DoShutdownAllApplications(PVOID pvoid)
{
	::CoInitialize(NULL);
	bool bError(false);
	SendMessage(GetDlgItem(g_hwndDlgShutdownAllApplication, IDC_PROGRESS2), PBM_SETRANGE, 0, MAKELPARAM(0, ::g_vCoApplications.size()));
	SendMessage(GetDlgItem(g_hwndDlgShutdownAllApplication, IDC_PROGRESS2), PBM_SETPOS, 1, 0);
	HWND hwndList = GetDlgItem(g_hwndDlgShutdownAllApplication, IDC_LISTAPPLICATIONS);
	HWND hwndButton = GetDlgItem(g_hwndDlgShutdownAllApplication, IDOK);
	for (long l = 0; l < g_vCoApplications.size(); l++) {
		if (l == g_vCoApplications.size() - 1) {
			EnableWindow(hwndButton, FALSE);
		}

		SendMessage(hwndList, LB_ADDSTRING, 0, (LPARAM)g_vCoApplications[l].m_wcsName);
		SendMessage(hwndList, LB_SETSEL, 1, l);
		if (ShutDownApplication(g_vCoApplications[l].m_wcsName)) {
			SendMessage(hwndList, LB_SETSEL, 0, l);
		}
		else {
			bError = true;
		}
		SendMessage(GetDlgItem(g_hwndDlgShutdownAllApplication, IDC_PROGRESS2), PBM_SETPOS, l + 1, 0);
		if (g_bCancelShutdown == true) {
			g_bCancelShutdown = false;
			INT iRes = MessageBox(GetActiveWindow(), L"Are you sure?", L"Confirmation", MB_YESNO);
			if (iRes == IDOK)
				break;
		}
	}
	EnableWindow(hwndButton, TRUE);
	SetDlgItemText(g_hwndDlgShutdownAllApplication, IDOK, L"OK");

	if (g_bAutoRun && !bError) {
		EndDialog(g_hwndDlgShutdownAllApplication, 0);
	}

	g_hwndDlgShutdownAllApplication = NULL;

	::CoUninitialize();
}

VOID DoRestartIIS(PVOID pvoid)
{
	::CoInitialize(NULL);

	//OSVERSIONINFO osvi = {0};
	//osvi.dwOSVersionInfoSize = sizeof(OSVERSIONINFO);

	//GetVersionEx(&osvi);

	//if ( osvi.dwMajorVersion >= 6 ) {
	//	ResetIIS();
	//}
	//else {					
	PROCESS_INFORMATION pi = { 0 };
	if (StopIIS(&pi)) {
		::WaitForSingleObject(pi.hProcess, INFINITE);
		::Sleep(1000);
		PROCESS_INFORMATION pi2 = { 0 };
		if (!StartIIS(&pi2)) {
			MessageBox(GetActiveWindow(), L"Failed to start IIS. Failed to execute \"iisreset /start\".", L"Error", MB_OK | MB_ICONERROR);
		}
	}
	else {
		MessageBox(GetActiveWindow(), L"Failed to stop IIS. Failed to execute \"iisreset /stop\".", L"Error", MB_OK | MB_ICONERROR);
	}
	//}
	::CoUninitialize();
}

bool StopIIS(PROCESS_INFORMATION* pi)
{
	if (IsWindows8OrGreater()) {
		SHELLEXECUTEINFOW execInfo = { 0 };
		execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
		execInfo.lpVerb = L"runas";
		execInfo.lpFile = L"iisreset";
		execInfo.lpParameters = L"/stop";
		execInfo.nShow = SW_SHOWNORMAL;
		execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
		bool bResult = ::ShellExecuteExW(&execInfo);
		pi->hProcess = execInfo.hProcess;
		return bResult;
	}
	else {
		wchar_t wcsCommand[] = L"iisreset /stop";//L"net stop iisadmin /y";		
		STARTUPINFO si = { 0 };
		si.cb = sizeof(STARTUPINFO);
		si.lpReserved = NULL;
		si.lpTitle = L"Stop IIS";
		
		BOOL b = CreateProcess(NULL, wcsCommand, NULL, NULL, FALSE, NORMAL_PRIORITY_CLASS, NULL, NULL, &si, pi);
		if (b == FALSE) {
			return false;
		}
	}
	return true;
}

//bool ResetIIS()
//{
//	SHELLEXECUTEINFOW execInfo = { 0 };
//	execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
//	execInfo.lpVerb = L"runas";
//	execInfo.lpFile = L"iisreset";
//	execInfo.nShow = SW_SHOWNORMAL;
//	execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
//	return (::ShellExecuteExW(&execInfo));
//}

bool StartIIS(PROCESS_INFORMATION* pi)
{
	if (IsWindows8OrGreater()) {
		SHELLEXECUTEINFOW execInfo = { 0 };
		execInfo.cbSize = sizeof(SHELLEXECUTEINFOW);
		execInfo.lpVerb = L"runas";
		execInfo.lpFile = L"iisreset";
		execInfo.lpParameters = L"/start";
		execInfo.nShow = SW_SHOWNORMAL;
		execInfo.fMask = SEE_MASK_NOCLOSEPROCESS;
		bool bResult = ::ShellExecuteExW(&execInfo);
		pi->hProcess = execInfo.hProcess;
		return bResult;
	}
	else {
		wchar_t wcsCommand[] = L"iisreset /start";//L"net start w3svc";						
		STARTUPINFO si = { 0 };
		si.cb = sizeof(STARTUPINFO);
		si.lpReserved = NULL;
		si.lpTitle = L"Start IIS";

		BOOL b = CreateProcess(NULL, wcsCommand, NULL, NULL, FALSE, NORMAL_PRIORITY_CLASS, NULL, NULL, &si, pi);
		if (b == FALSE) {
			return false;
		}
	}

	return true;
}

bool RefreshAllConfiguredComponents()
{
	if (g_vCoApplications.size() > 0)
		g_vCoApplications.clear();

	CComPtr<ICOMAdminCatalog> pAdminCatalog;
	HRESULT hr = pAdminCatalog.CoCreateInstance(CComBSTR("COMAdmin.COMAdminCatalog.1"));//CLSID_COMAdminCatalog );
	if (FAILED(hr)) {
		MessageBox(GetActiveWindow(), L"Failed to create the COMAdminCatalog.", L"Error", MB_OK | MB_ICONERROR);
		return false;
	}
	hr = pAdminCatalog->RefreshComponents();
	if (FAILED(hr)) {
		MessageBox(GetActiveWindow(), L"Failed to refresh components.", L"Error", MB_OK | MB_ICONERROR);
		return false;
	}

	return true;
}

////////////////////////////////////////////////////////////////////////////////////////////////
//
// ExtractPath
//
STDMETHODIMP ExtractPath( /*[in]*/ BSTR full_path, /*[out, retval]*/ BSTR* path)
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