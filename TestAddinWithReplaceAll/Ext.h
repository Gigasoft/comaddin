// Ext.h : Declaration of the CExt

#pragma once
#include "resource.h"       // main symbols



#include "TestAddinWithReplaceAll_i.h"



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

typedef IDispatchImpl<_IDTExtensibility2, &__uuidof(_IDTExtensibility2), &__uuidof(__AddInDesignerObjects), 1, 0>
IDTImpl;

typedef IDispatchImpl<IRibbonExtensibility, &__uuidof(IRibbonExtensibility), &__uuidof(__Office), /* wMajor = */ 2, /* wMinor = */ 5>
RibbonImpl;

typedef IDispatchImpl<_FormRegionStartup, &__uuidof(_FormRegionStartup), &__uuidof(__Outlook), /* wMajor = */ 9, /* wMinor = */ 4>
FormImpl;

typedef IDispatchImpl<IRibbonCallback, &__uuidof(IRibbonCallback)>
CallbackImpl;



// CExt

class ATL_NO_VTABLE CExt :
	public CComObjectRootEx<CComSingleThreadModel>,
	//public CComCoClass<CExt, &CLSID_Ext>,
	//public IDispatchImpl<IExt, &IID_IExt, &LIBID_TestAddinWithReplaceAllLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
	public CComCoClass<CExt, &CLSID_Ext>,
	//public CWindowImpl<CExt>,
	public IDispatchImpl<IExt, &IID_IExt, &LIBID_TestAddinWithReplaceAllLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public RibbonImpl,
	public IDTImpl,
	public CallbackImpl,
	public FormImpl
{
public:
	CExt();


DECLARE_REGISTRY_RESOURCEID(IDR_EXT)


BEGIN_COM_MAP(CExt)
	//COM_INTERFACE_ENTRY(IExt)
	//COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY2(IDispatch, IRibbonCallback)
	COM_INTERFACE_ENTRY(IExt)
	COM_INTERFACE_ENTRY(_IDTExtensibility2)
	COM_INTERFACE_ENTRY(_FormRegionStartup)
	COM_INTERFACE_ENTRY(IRibbonExtensibility)
	COM_INTERFACE_ENTRY(IRibbonCallback)
	COM_INTERFACE_ENTRY_IID(DIID_InspectorsEvents, IExt)
	COM_INTERFACE_ENTRY_IID(DIID_ExplorerEvents, IExt)
END_COM_MAP()

BEGIN_CATEGORY_MAP(CExt)
	IMPLEMENTED_CATEGORY(CATID_SafeForInitializing)
	IMPLEMENTED_CATEGORY(CATID_SafeForScripting)
END_CATEGORY_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct();
	void FinalRelease();

public:

	HWND m_hwndOutlook;
	CComQIPtr< Outlook::_Application > m_spApplication;

	CComPtr<IRibbonUI> m_spRibbonUI;

	static bool m_isPaneSet;
	static bool m_isRunningExchange;
	static int m_nOutlookMajorVersion;


	// IRibbonCallback
	STDMETHOD(ButtonClicked)(IDispatch * RibbonControl);
	STDMETHOD(GetImage)(IDispatch * RibbonControl, IDispatch ** ppdispImage);
	STDMETHOD(OnLoad)(IDispatch* RibbonIDl);
	STDMETHOD(GetEnabled)(IDispatch * RibbonControl, VARIANT_BOOL *pvarfEnabled);
	STDMETHOD(GetText)(IDispatch * RibbonControl, BSTR *pVal);
	STDMETHOD(GetCheckboxState)(IDispatch * RibbonControl, VARIANT_BOOL * state);
	STDMETHOD(CheckboxClicked)(IDispatch * RibbonControl, VARIANT_BOOL state);

	// IRibbonExtensibility
	STDMETHOD(GetCustomUI)(BSTR RibbonID, BSTR * RibbonXml);

	// IDTExtensibility2 implementation:
	STDMETHOD(OnConnection)(IDispatch * Application, ext_ConnectMode ConnectMode, IDispatch *AddInInst, SAFEARRAY **custom);
	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY **custom);
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY **custom)
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		return E_NOTIMPL;
	}

	STDMETHOD(OnStartupComplete)(SAFEARRAY **custom);//{return E_NOTIMPL;}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY **custom)
	{
		AFX_MANAGE_STATE(AfxGetStaticModuleState());
		return E_NOTIMPL;
	}

	// Event processing
	STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags,
		DISPPARAMS* pdispparams, VARIANT* pvarResult,
		EXCEPINFO* pexcepinfo, UINT* puArgErr);

	// _FormRegionStartup
	STDMETHOD(GetFormRegionStorage)(BSTR FormRegionName, LPDISPATCH Item, long LCID, OlFormRegionMode FormRegionMode, OlFormRegionSize FormRegionSize, VARIANT * Storage);
	STDMETHOD(BeforeFormRegionShow)(_FormRegion * FormRegion);
	STDMETHOD(GetFormRegionManifest)(BSTR FormRegionName, long LCID, VARIANT * Manifest);
	STDMETHOD(GetFormRegionIcon)(BSTR FormRegionName, long LCID, OlFormRegionIcon Icon, VARIANT * Result);
protected:
	_ExplorersPtr    m_spExplorers;
	CComPtr<Outlook::_Explorer> m_spExplorer;
	DWORD m_dwExplorerCookie;
	std::wstring GetStringResource(UINT id);
	SAFEARRAY* GetOFSResource(int nId);
	BSTR GetXMLResource(int nId);
	HRESULT HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes);
	// This is for the ribbon button and indicates that encryption is turned on for the compose message window
	CComPtr<IPictureDisp> m_spEncryptionOn;
	// This is for the ribbon button and indicates that encryption is turned off for the compose message window
	CComPtr<IPictureDisp> m_spEncryptionOff;

	CComPtr<IPictureDisp> GetPictureFromBitmap(HBITMAP hBitmap);


	Outlook::_InspectorsPtr m_spInspectors;
	DWORD m_dwInpectorsCookie;
	int GetOutlookVersion(CComQIPtr< Outlook::_Application > m_spApplication);

};

OBJECT_ENTRY_AUTO(__uuidof(Ext), CExt)
