// dllmain.h : Declaration of module class.

class CDeveloperExtensionsContextMenuModule : public ATL::CAtlDllModuleT< CDeveloperExtensionsContextMenuModule >
{
public :
	DECLARE_LIBID(LIBID_DeveloperExtensionsContextMenuLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_DEVELOPEREXTENSIONSCONTEXTMENU, "{349C5C26-03CD-4886-B651-1E425A33672E}")
};

extern class CDeveloperExtensionsContextMenuModule _AtlModule;
