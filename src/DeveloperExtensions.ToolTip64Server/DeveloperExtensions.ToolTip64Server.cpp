// DeveloperExtensions.ToolTip64Server.cpp : Implementation of WinMain


#include "stdafx.h"
#include "resource.h"
#include "DeveloperExtensionsToolTip64Server_i.h"
#include "xdlldata.h"


using namespace ATL;


class CDeveloperExtensionsToolTip64ServerModule : public ATL::CAtlExeModuleT< CDeveloperExtensionsToolTip64ServerModule >
{
public :
	DECLARE_LIBID(LIBID_DeveloperExtensionsToolTip64ServerLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_DEVELOPEREXTENSIONSTOOLTIP64SERVER, "{310A43F8-7C0A-4DF1-A084-82E2ECBE17B2}")
	};

CDeveloperExtensionsToolTip64ServerModule _AtlModule;



//
extern "C" int WINAPI _tWinMain(HINSTANCE /*hInstance*/, HINSTANCE /*hPrevInstance*/, 
								LPTSTR /*lpCmdLine*/, int nShowCmd)
{
	return _AtlModule.WinMain(nShowCmd);
}

