// dllmain.h : Declaration of module class.

class CCV82Module : public ATL::CAtlDllModuleT< CCV82Module >
{
public :
	DECLARE_LIBID(LIBID_CV82Lib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_CV82, "{5DF79B3E-EE00-4C18-B6D4-0034D21DF3C0}")
};

extern class CCV82Module _AtlModule;
