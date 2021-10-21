// Ext.cpp : Implementation of CExt

#include "stdafx.h"
#include "Ext.h"


// CExt
#define ID_PREFIX L"preveil"
int g_verOLMajor = 0;


int CExt::m_nOutlookMajorVersion = 10;
bool  CExt::m_isRunningExchange = false;
bool  CExt::m_isPaneSet = false;

CExt::CExt()
	: m_hwndOutlook(NULL)
	, m_dwInpectorsCookie(0)
	, m_dwExplorerCookie(0)
{
}

HRESULT CExt::FinalConstruct()
{
	try
	{
	}
	catch (...)
	{
	}
	return S_OK;
}

void CExt::FinalRelease()
{

}

STDMETHODIMP CExt::OnConnection(IDispatch *pApplication, ext_ConnectMode ConnectMode, IDispatch *pAddInInst, SAFEARRAY ** /*custom*/)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());
	//g_log->LogTrace("CExt::OnConnection enter");

	HRESULT h = S_OK;
	try
	{


		// Get and save the application interface
		CComQIPtr<  Outlook::_Application > spApplication(pApplication);
		ATLASSERT(spApplication);
		m_spApplication = spApplication;

		// Get and set an advise on the inspectors
		m_spApplication->get_Inspectors(&m_spInspectors);
		HRESULT hr = AtlAdvise(m_spInspectors, GetUnknown(), DIID_InspectorsEvents, &m_dwInpectorsCookie);
		//::MessageBox(nullptr, L"ADVISE", L"new insp", MB_OK);

		m_spApplication->ActiveExplorer(&m_spExplorer);

		// Capture explorer events (selected folder)
		hr = AtlAdvise(m_spExplorer, GetUnknown(), DIID_ExplorerEvents, &m_dwExplorerCookie);


		// Get the version of Outlook
		m_nOutlookMajorVersion = GetOutlookVersion(m_spApplication);
		// if this is less than Outlook 2003 we are simply going return
		// without doing a thing.
		if (m_nOutlookMajorVersion < 11)
			return S_OK;


		// Create the contact monitor
		CComPtr<Outlook::_NameSpace> spNameSpace;
		m_spApplication->GetNamespace(_bstr_t(_T("MAPI")), &spNameSpace);

		CComPtr<IUnknown> spUnk;
		spNameSpace->get_MAPIOBJECT(&spUnk);


		OlExchangeConnectionMode mode;
		spNameSpace->get_ExchangeConnectionMode(&mode);

		if (mode == Outlook::olNoExchange)
		{
			CExt::m_isRunningExchange = false;
		}
		else
		{
			CExt::m_isRunningExchange = true;
		}




		if (FAILED(m_spApplication->get_Explorers(&m_spExplorers)))
			return E_FAIL;

	}
	catch (...)
	{
	}
	return S_OK;
}

STDMETHODIMP CExt::OnStartupComplete(SAFEARRAY ** /*custom*/)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	try
	{

	}
	catch (...)
	{
	}

	return S_OK;
}


STDMETHODIMP CExt::OnDisconnection(ext_DisconnectMode /*RemoveMode*/, SAFEARRAY ** /*custom*/)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	try
	{
		// Release advise on new inspectors
		if (m_dwInpectorsCookie != 0)
		{
			AtlUnadvise(m_spInspectors, DIID_InspectorsEvents, m_dwInpectorsCookie);
			m_dwInpectorsCookie = 0;
		}


	}
	catch (...)
	{
	}
	return S_OK;
}

STDMETHODIMP CExt::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags,
	DISPPARAMS* pdispparams, VARIANT* pvarResult,
	EXCEPINFO* pexcepinfo, UINT* puArgErr)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	ATLTRACE("\nCExt Event %x  %ld\n\n", dispidMember, dispidMember);
	HRESULT hr = DISP_E_MEMBERNOTFOUND;
	
	
	if (DISP_E_MEMBERNOTFOUND == hr)
	{
		hr = IDTImpl::Invoke(dispidMember,
			riid,
			lcid,
			wFlags,
			pdispparams,
			pvarResult,
			pexcepinfo,
			puArgErr);
		if (hr != DISP_E_MEMBERNOTFOUND)
			return hr;
	}

	
	if (DISP_E_MEMBERNOTFOUND == hr)
	{

		hr = CallbackImpl::Invoke(dispidMember,
			riid,
			lcid,
			wFlags,
			pdispparams,
			pvarResult,
			pexcepinfo,
			puArgErr);

		if (hr != DISP_E_MEMBERNOTFOUND)
			return hr;
	}
	if (DISP_E_MEMBERNOTFOUND == hr)
	{


	}
	hr = FormImpl::Invoke(dispidMember,
		riid,
		lcid,
		wFlags,
		pdispparams,
		pvarResult,
		pexcepinfo,
		puArgErr);

	if (hr != DISP_E_MEMBERNOTFOUND)
		return hr;


	return hr;
}





STDMETHODIMP CExt::GetCustomUI(BSTR RibbonID, BSTR * RibbonXml)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (RibbonID)
	{
		_bstr_t bt(RibbonID, true);
		std::wstring wsRibbonId(RibbonID, ::SysStringLen(RibbonID));

		if (wsRibbonId.length())
		{
			/*
			if (wsRibbonId == _T("Microsoft.Outlook.Mail.Compose"))
			{

				BSTR bstrRibbonXml = GetXMLResource(IDR_COMPOSE_RIBBON_FORMA);
				*RibbonXml = bstrRibbonXml;
				return S_OK;
			}
			*/
		}
	}

	return S_OK;
}




STDMETHODIMP CExt::GetFormRegionStorage(BSTR FormRegionName, LPDISPATCH Item, long LCID, OlFormRegionMode FormRegionMode, OlFormRegionSize FormRegionSize, VARIANT * Storage)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	try
	{
		V_VT(Storage) = VT_ARRAY | VT_UI1;
		if (FormRegionMode == olFormRegionCompose)
		{
			V_ARRAY(Storage) = GetOFSResource(IDR_MESSAGE_FORM_RGN_THIN/*IDR_BLANK*/);
			return S_OK;

		}
	}
	catch (const std::exception &e)
	{
	}
	return DISP_E_MEMBERNOTFOUND;
}


STDMETHODIMP CExt::BeforeFormRegionShow(_FormRegion * FormRegion)
{
	try
	{
		Outlook::OlFormRegionMode mode;
		FormRegion->get_FormRegionMode(&mode);

		if (mode == olFormRegionCompose)
		{

		}
	}
	catch (const std::exception &e)
	{
	}
	return S_OK;
}

STDMETHODIMP CExt::GetFormRegionManifest(BSTR FormRegionName, long LCID, VARIANT * Manifest)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	try
	{
		// NOTE TO PROGRAMMER. Microsoft debacle
		// There are two different formats for Outlook manifest files.
		// One for Outlook 2010 and the other for 2013 & 2016.
		//
		// 2010 ref: https://docs.microsoft.com/en-us/visualstudio/vsto/creating-outlook-form-regions#Adding
		// Outlook 2010 region types are Separate, Adjoining, Replacement and Replace-all
		//
		// 2013/2016 ref: https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/dd492010(v=office.12)
		// Outlook 2013/2016 reguib types are separate, adjoining, replace, replaceAll
		//
		//::MessageBox(nullptr, L"GetFormRegionManifest", L"Create form", MB_OK);

		V_VT(Manifest) = VT_BSTR;
		BSTR bstrManifest = nullptr;
		if (m_nOutlookMajorVersion > 14)
		{
			bstrManifest = GetXMLResource(IDR_FORM_REGION_MANIFEST);
			V_BSTR(Manifest) = bstrManifest;
		}
	}
	catch (const std::exception &e)
	{
	}
	return S_OK;
}

STDMETHODIMP CExt::GetFormRegionIcon(BSTR FormRegionName, long LCID, OlFormRegionIcon Icon, VARIANT * Result)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	return E_NOTIMPL;
}

std::wstring CExt::GetStringResource(UINT id)
{
	wchar_t wszStringResource[MAX_PATH] = { NULL };
	::LoadString(_AtlBaseModule.GetModuleInstance(), id, wszStringResource, sizeof(wszStringResource));

	return std::wstring(wszStringResource);
}

SAFEARRAY* CExt::GetOFSResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;
	if (FAILED(HrGetResource(nId, TEXT("OFS"),
		&pResourceData, &dwSizeInBytes)))
		return NULL;
	SAFEARRAY* psa;
	SAFEARRAYBOUND dim = { dwSizeInBytes, 0 };
	psa = SafeArrayCreate(VT_UI1, 1, &dim);
	if (psa == NULL)
		return NULL;
	BYTE* pSafeArrayData;
	SafeArrayAccessData(psa, (void**)&pSafeArrayData);
	memcpy((void*)pSafeArrayData, pResourceData, dwSizeInBytes);
	SafeArrayUnaccessData(psa);
	return psa;
}

BSTR CExt::GetXMLResource(int nId)
{
	LPVOID pResourceData = NULL;
	DWORD dwSizeInBytes = 0;
	HRESULT hr = HrGetResource(nId, TEXT("XML"),
		&pResourceData, &dwSizeInBytes);
	if (FAILED(hr))
		return NULL;
	// Assumes that the data is not stored in Unicode.
	CComBSTR cbstr(dwSizeInBytes, reinterpret_cast<LPCSTR>(pResourceData));
	return cbstr.Detach();
}


HRESULT CExt::HrGetResource(int nId, LPCTSTR lpType, LPVOID* ppvResourceData, DWORD* pdwSizeInBytes)
{
	HMODULE hModule = _AtlBaseModule.GetModuleInstance();
	if (!hModule)
		return E_UNEXPECTED;
	HRSRC hRsrc = FindResource(hModule, MAKEINTRESOURCE(nId), lpType);
	if (!hRsrc)
		return HRESULT_FROM_WIN32(GetLastError());
	HGLOBAL hGlobal = LoadResource(hModule, hRsrc);
	if (!hGlobal)
		return HRESULT_FROM_WIN32(GetLastError());
	*pdwSizeInBytes = SizeofResource(hModule, hRsrc);
	*ppvResourceData = LockResource(hGlobal);
	return S_OK;
}



STDMETHODIMP CExt::ButtonClicked(IDispatch * RibbonControl)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	try
	{
		//ATLTRACE("Ribbon callback >> ButtonClicked\n");
		CComPtr<Office1::IRibbonControl> spRibbonControl;
		HRESULT hr = RibbonControl->QueryInterface(__uuidof(Office1::IRibbonControl), reinterpret_cast<LPVOID*>(&spRibbonControl));
		if (SUCCEEDED(hr) && spRibbonControl)
		{

			BSTR bstrId = nullptr;
			hr = spRibbonControl->get_Id(&bstrId);
			if (SUCCEEDED(hr) && bstrId)
			{
				std::wstring wsId(bstrId, ::SysStringLen(bstrId));
				if (wsId == L"preveil_212")
				{
					m_spRibbonUI->InvalidateControl(CComBSTR("preveil_212"));
				}
				else if (wsId == L"preveil_211")
				{
					m_spRibbonUI->InvalidateControl(CComBSTR("preveil_211"));
				}
			}
		}
	}
	catch (const std::exception &e)
	{
	}
	return S_OK;
}

STDMETHODIMP CExt::GetImage(IDispatch * RibbonControl, IDispatch ** ppdispImage)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	try
	{
		//ATLTRACE("Ribbon callback >> GetImage\n");

		CComPtr<Office1::IRibbonControl> spRibbonControl;
		RibbonControl->QueryInterface(__uuidof(Office1::IRibbonControl), (LPVOID*)&spRibbonControl);

		BSTR bstrId = NULL;
		spRibbonControl->get_Id(&bstrId);
		if (bstrId)
		{
			std::wstring strID(bstrId, ::SysStringLen(bstrId));
			::SysFreeString(bstrId);

			UINT uiPos = static_cast<UINT>(strID.find(ID_PREFIX));
			if (uiPos != std::wstring::npos)
			{

				CComPtr<IDispatch> spDispInsp;
				CComPtr<Outlook::_Inspector> spInspector;
				HRESULT hr = spRibbonControl->get_Context(&spDispInsp);
				if (SUCCEEDED(hr) && spDispInsp)
				{
					hr = spDispInsp->QueryInterface(__uuidof(Outlook::_Inspector), (LPVOID*)&spInspector);


					//ATLTRACE("Inspector = %x  ribbon = %x\n", spInspector, spRibbonControl);
				}

				uiPos++;
				strID = strID.substr(uiPos + _tcslen(ID_PREFIX), -1);

				if (strID == L"212" || strID == L"211")
				{
					HICON hIcon = (HICON)::LoadImage(_AtlBaseModule.GetResourceInstance(),
						MAKEINTRESOURCE(IDI_ENCRYPTION_ON),
						IMAGE_ICON,
						32,
						32,
						0);

					if (nullptr == m_spEncryptionOn)
					{
						HICON hIcon = (HICON)::LoadImage(_AtlBaseModule.GetResourceInstance(),
							MAKEINTRESOURCE(IDI_ENCRYPTION_ON),
							IMAGE_ICON,
							32,
							32,
							0);


						ICONINFO iconInfo;
						::ZeroMemory(&iconInfo, sizeof(ICONINFO));
						GetIconInfo(hIcon, &iconInfo);
						m_spEncryptionOn = GetPictureFromBitmap(iconInfo.hbmColor);
					}

					IDispatch* spDisp = NULL;
					m_spEncryptionOn->QueryInterface(IID_IDispatch, (LPVOID*)&spDisp);
					*ppdispImage = spDisp;
				}

			}
		}
	}
	catch (const std::exception &e)
	{
	}

	return S_OK;
}

STDMETHODIMP CExt::OnLoad(IDispatch* RibbonIDl)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	if (m_spRibbonUI)
	{
		m_spRibbonUI.Release();
		m_spRibbonUI = NULL;
	}

	RibbonIDl->QueryInterface(__uuidof(Office1::IRibbonUI), (LPVOID*)&m_spRibbonUI);
	return S_OK;
}


STDMETHODIMP CExt::CheckboxClicked(IDispatch * RibbonControl, VARIANT_BOOL state)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	DWORD dwDisable = 0;
	if (state)
	{
	}
	else
	{
	}

	return S_OK;
}

STDMETHODIMP CExt::GetCheckboxState(IDispatch * RibbonControl, VARIANT_BOOL * state)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	CComPtr<IRibbonControl> spRibbonControl;
	HRESULT hr = RibbonControl->QueryInterface(__uuidof(Office1::IRibbonControl), reinterpret_cast<LPVOID*>(&spRibbonControl));
	if (SUCCEEDED(hr) && spRibbonControl)
	{
		BSTR bstrId = nullptr;
		hr = spRibbonControl->get_Id(&bstrId);
		if (SUCCEEDED(hr) && bstrId)
		{
			std::wstring wsId(bstrId, ::SysStringLen(bstrId));

			if (wsId == L"preveil_211")
			{

				if (true)
					*state = VARIANT_TRUE;
				else
					*state = VARIANT_FALSE;
			}
		}
	}

	return S_OK;
}


STDMETHODIMP CExt::GetEnabled(IDispatch * RibbonControl, VARIANT_BOOL *pvarfEnabled)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	try
	{
		//ATLTRACE("Ribbon callback >> GetEnabled\n");
		CComPtr<IRibbonControl> spRibbonControl;
		HRESULT hr = RibbonControl->QueryInterface(__uuidof(Office1::IRibbonControl), (LPVOID*)&spRibbonControl);
		if (SUCCEEDED(hr) && spRibbonControl)
		{
			CComPtr<IDispatch> spDisp;
			hr = spRibbonControl->get_Context(&spDisp);
			if (SUCCEEDED(hr) && spDisp)
			{

				BSTR bstrId = NULL;
				spRibbonControl->get_Id(&bstrId);
				if (bstrId)
				{
					std::wstring strID(bstrId, ::SysStringLen(bstrId));
					::SysFreeString(bstrId);
					*pvarfEnabled = VARIANT_TRUE;
				}
			}
		}

	}
	catch (const std::exception &e)
	{
	}

	return S_OK;
}

STDMETHODIMP CExt::GetText(IDispatch * RibbonControl, BSTR *pVal)
{
	AFX_MANAGE_STATE(AfxGetStaticModuleState());

	try
	{
		//ATLTRACE("Ribbon callback >> GetText\n");
		CComPtr<IRibbonControl> spRibbonControl;
		HRESULT hr = RibbonControl->QueryInterface(__uuidof(Office1::IRibbonControl), (LPVOID*)&spRibbonControl);
		if (SUCCEEDED(hr) && spRibbonControl)
		{
			CComPtr<IDispatch> spDisp;
			hr = spRibbonControl->get_Context(&spDisp);
			if (SUCCEEDED(hr) && spDisp)
			{
				CComPtr<Outlook::_Inspector> spInspector;
				hr = spDisp->QueryInterface(__uuidof(Outlook::_Inspector), (LPVOID*)&spInspector);

				BSTR bstrId = NULL;
				spRibbonControl->get_Id(&bstrId);
				if (bstrId)
				{
					std::wstring strID(bstrId, ::SysStringLen(bstrId));
					::SysFreeString(bstrId);
					if (strID == L"preveil_212")
					{
						*pVal = ::SysAllocString(L"Turn Off Encryption");
					}
					else if (strID == L"preveil_211")
					{
						*pVal = ::SysAllocString(L"Turn Off Encryption");
					
					}
				}
			}
		}

	}
	catch (const std::exception &e)
	{
	}
	return S_OK;
}


CComPtr<IPictureDisp> CExt::GetPictureFromBitmap(HBITMAP hBitmap)
{
	CComPtr<IPictureDisp> spPictDisp;
	PICTDESC pd;
	pd.cbSizeofstruct = sizeof(pd);
	pd.picType = PICTYPE_BITMAP;
	pd.bmp.hbitmap = hBitmap;

	OleCreatePictureIndirect(&pd,
		IID_IPictureDisp,
		TRUE,
		(LPVOID*)&spPictDisp);

	return spPictDisp;
}

int CExt::GetOutlookVersion(CComQIPtr< Outlook::_Application > m_spApplication)
{
	int nVersion = 10;
	try
	{
		BSTR bstrVersion = NULL;
		m_spApplication->get_Version(&bstrVersion);
		int v2, v3, v4;
		swscanf_s(bstrVersion, L"%d.%d.%d.%d", &g_verOLMajor, &v2, &v3, &v4);
		std::wstring str = (const TCHAR*)_bstr_t(bstrVersion, true);
		::SysFreeString(bstrVersion);

		std::wstring wsMsg = L"Outlook version: ";
		wsMsg += str;


		int nMajorVersion = 0;
		int nMinorRelease = 0;
		int nIncremental = 0;
		TCHAR szTmp[50] = { NULL };

		///////////////////////////////////////////////////
		// Get version information from the current system 
		// installation.
		///////////////////////////////////////////////////
#ifdef _UNICODE
		int nNumberOfFields = swscanf(str.c_str(), _T("%i.%i.%i.%s"),
			&nMajorVersion, &nMinorRelease, &nIncremental, szTmp);
#else
		int nNumberOfFields = sscanf(str.c_str(), _T("%i.%i.%i.%s"),
			&nMajorVersion, &nMinorRelease, &nIncremental, szTmp);
#endif

		// Set Outlook Version (we will only support 2002 & 2003)
		nVersion = nMajorVersion;
	}
	catch (const std::exception &e)
	{
	}

	return nVersion;
}

