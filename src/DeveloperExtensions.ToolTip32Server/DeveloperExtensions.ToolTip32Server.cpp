// DeveloperExtensions.ToolTip32Server.cpp : Implementation of WinMain


#include "stdafx.h"
#include "resource.h"
#include "DeveloperExtensionsToolTip32Server_i.h"
#include "xdlldata.h"


using namespace ATL;


class CDeveloperExtensionsToolTip32ServerModule : public ATL::CAtlExeModuleT< CDeveloperExtensionsToolTip32ServerModule >
{
public :
	DECLARE_LIBID(LIBID_DeveloperExtensionsToolTip32ServerLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_DEVELOPEREXTENSIONSTOOLTIP32SERVER, "{16AE3732-1636-4967-AF8B-344AC3B0A510}")
	};

CDeveloperExtensionsToolTip32ServerModule _AtlModule;



//
extern "C" int WINAPI _tWinMain(HINSTANCE /*hInstance*/, HINSTANCE /*hPrevInstance*/, 
								LPTSTR /*lpCmdLine*/, int nShowCmd)
{
	return _AtlModule.WinMain(nShowCmd);
}

