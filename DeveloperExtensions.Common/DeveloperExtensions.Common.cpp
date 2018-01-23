// Include
#include <windows.h>
#include <sstream>
#include <string>
#include "resource.h"

using std::wstring;
using std::wstringstream;

// Global data
HINSTANCE g_hInstanceDll(NULL);



// Function prototypes
void WINAPI DevExtShowAboutDlg();
INT_PTR CALLBACK AboutDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam);
wstring GetFileVersion( );

// DllMain
BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpvReserved)
{
	switch (dwReason)
	{
	case DLL_PROCESS_ATTACH:
		::DisableThreadLibraryCalls(hInstance);
		g_hInstanceDll = hInstance;
		break;
	case DLL_THREAD_ATTACH:
		break;
	case DLL_THREAD_DETACH:
		break;
	case DLL_PROCESS_DETACH:
		break;
	};

	return TRUE;
}// BOOL DllMain



// DevExtShowAboutDlg
void __stdcall DevExtShowAboutDlg()
{
	DialogBox(g_hInstanceDll, MAKEINTRESOURCE(IDD_ABOUT), GetActiveWindow(), AboutDialogProc);
}// void DevExtShowAboutDlg



// AboutDialogProc
INT_PTR CALLBACK AboutDialogProc(HWND hwndDlg, UINT uMsg, WPARAM wParam, LPARAM lParam)
{
	switch (uMsg)
	{
	case WM_INITDIALOG:
	{
		wstring versionNumber = GetFileVersion( );
		wstring txt = L"Developer Extensions v";
		txt += versionNumber;

		SetDlgItemText( hwndDlg, IDC_STATIC_VERSION_STRING, txt.c_str( ) );

		RECT rectDlg{ 0 };
		GetWindowRect(hwndDlg, &rectDlg);

		RECT rectShell{ 0 };
		GetWindowRect(GetShellWindow(), &rectShell);

		int shellWidth = static_cast<int>((rectShell.right - rectShell.left) / 2);
		int shellHeight = static_cast<int>((rectShell.bottom - rectShell.top) / 2);

		int offsetx = static_cast<int>((rectDlg.right - rectDlg.left) / 2);
		int offsety = static_cast<int>((rectDlg.bottom - rectDlg.top) / 2);

		MoveWindow(hwndDlg, shellWidth - offsetx, shellHeight - offsety, rectDlg.right - rectDlg.left, rectDlg.bottom - rectDlg.top, TRUE);
		
		return TRUE;
	}

	case WM_COMMAND:
		switch (LOWORD(wParam))
		{
		case IDOK:
			EndDialog(hwndDlg, 0);
			return TRUE;
			break;
		default:
			break;
		}
		break;

	default:
		break;
	}

	return FALSE;
}// BOOL AboutDialogProc

 //
 // Method: GetFileVersion
 //
wstring GetFileVersion( )
{
	wchar_t wszFileName[ MAX_PATH ];
	GetModuleFileName( g_hInstanceDll, wszFileName, MAX_PATH );

	DWORD dwSize = GetFileVersionInfoSizeEx( FILE_VER_GET_NEUTRAL, wszFileName, NULL );
	BYTE* pData = new BYTE[ dwSize ];
	GetFileVersionInfoEx( FILE_VER_GET_NEUTRAL, wszFileName, NULL, dwSize, pData );

	LPVOID pVersionInfo( NULL );
	UINT uiLen( 0 );
	VerQueryValue( pData, L"\\", &pVersionInfo, &uiLen );

	VS_FIXEDFILEINFO* pVerInfo = (VS_FIXEDFILEINFO*) pVersionInfo;

	DWORD dwMajorVersionMsb = HIWORD( pVerInfo->dwFileVersionMS );
	DWORD dwMajorVersionLsb = LOWORD( pVerInfo->dwFileVersionMS );
	DWORD dwMinorVersionMsb = HIWORD( pVerInfo->dwFileVersionLS );
	DWORD dwMinorVersionLsb = LOWORD( pVerInfo->dwFileVersionLS );

	wstringstream wss;
	wss << dwMajorVersionMsb << L"." << dwMajorVersionLsb << L"." << dwMinorVersionMsb << L"." << dwMinorVersionLsb;

	wstring retVal;
	wss >> retVal;

	delete[] pData;

	return retVal;
}