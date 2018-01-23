// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "DeveloperExtensionsPropertySheet_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CDeveloperExtensionsPropertySheetModule _AtlModule;
HINSTANCE g_hInstanceDll = nullptr;
// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	g_hInstanceDll = hInstance;

#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(hInstance, dwReason, lpReserved))
		return FALSE;
#endif
	hInstance;
	return _AtlModule.DllMain(dwReason, lpReserved); 
}
