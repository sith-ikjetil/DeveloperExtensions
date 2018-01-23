//
// #include
//
#include <Windows.h>
#include "../resource.h"
#include <string>
#include <vector>
#include "EAction.h"
#include "CApplication.h"
#include "../../Include/itsoftware-win.h"
#include "../../Include/itsoftware-win.cpp"

//
// #pragma
//
#pragma comment(linker, "/manifestdependency:\"type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='*' publicKeyToken='6595b64144ccf1df' language='*'\"")

//
// using
//
using DeveloperExtensions::CApplication;
using ItSoftware::Win::ItsWin;
using std::wstring;
using std::endl;
using std::vector;

//
// Global Data
//
vector<wstring> g_errorlog;

//
// Function: wWinMain
//
INT WINAPI wWinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, PWSTR pCmdLine, INT nCmdShow)
{
//	INITCOMMONCONTROLSEX icce{ 0 };
	//icce.dwSize = sizeof(INITCOMMONCONTROLSEX);
	//icce.dwICC = ICC_STANDARD_CLASSES;
	//InitCommonControlsEx(&icce);

	CApplication app;
	INT return_value = app.wWinMain(hInstance, hPrevInstance, pCmdLine, nCmdShow);
	if ( return_value ) {
		wstringstream msg;
		for (auto itr = g_errorlog.rbegin(); itr < g_errorlog.rend(); itr++)
		{
			msg << L"--> " << *itr << endl;
		}
		wstring error_message = msg.str();
		ItsWin::ShowErrorMessage(L"Developer Extensions", error_message);
	}
	return return_value;
}