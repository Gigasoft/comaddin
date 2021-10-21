// dllmain.cpp : Implementation of DllMain.

#include "stdafx.h"
#include "resource.h"
#include "TestAddinWithReplaceAll_i.h"
#include "dllmain.h"
#include "xdlldata.h"

CTestAddinWithReplaceAllModule _AtlModule;

class CTestAddinWithReplaceAllApp : public CWinApp
{
public:

// Overrides
	virtual BOOL InitInstance();
	virtual int ExitInstance();

	DECLARE_MESSAGE_MAP()
};

BEGIN_MESSAGE_MAP(CTestAddinWithReplaceAllApp, CWinApp)
END_MESSAGE_MAP()

CTestAddinWithReplaceAllApp theApp;

BOOL CTestAddinWithReplaceAllApp::InitInstance()
{
#ifdef _MERGE_PROXYSTUB
	if (!PrxDllMain(m_hInstance, DLL_PROCESS_ATTACH, NULL))
		return FALSE;
#endif
	return CWinApp::InitInstance();
}

int CTestAddinWithReplaceAllApp::ExitInstance()
{
	return CWinApp::ExitInstance();
}
