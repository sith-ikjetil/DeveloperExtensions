// dllmain.h : Declaration of module class.

class CDeveloperExtensionsToolTipModule : public ATL::CAtlDllModuleT< CDeveloperExtensionsToolTipModule >
{
public :
	DECLARE_LIBID(LIBID_DeveloperExtensionsToolTipLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_DEVELOPEREXTENSIONSTOOLTIP, "{A4B6F107-32C3-4F8A-8484-F526305134C2}")
};

extern class CDeveloperExtensionsToolTipModule _AtlModule;
