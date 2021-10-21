// dllmain.h : Declaration of module class.

class CTestAddinWithReplaceAllModule : public ATL::CAtlDllModuleT< CTestAddinWithReplaceAllModule >
{
public :
	DECLARE_LIBID(LIBID_TestAddinWithReplaceAllLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_TESTADDINWITHREPLACEALL, "{35343ED8-988B-4563-A651-2D6DDD1AA6EF}")
};

extern class CTestAddinWithReplaceAllModule _AtlModule;
