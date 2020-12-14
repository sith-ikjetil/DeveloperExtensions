// dllmain.h : Declaration of module class.

class CDeveloperExtensionsPropertySheetModule : public ATL::CAtlDllModuleT< CDeveloperExtensionsPropertySheetModule >
{
public :
	DECLARE_LIBID(LIBID_DeveloperExtensionsPropertySheetLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_DEVELOPEREXTENSIONSPROPERTYSHEET, "{C87AD07C-0D5F-4258-9EFF-023437BB5C98}")
};

extern class CDeveloperExtensionsPropertySheetModule _AtlModule;
