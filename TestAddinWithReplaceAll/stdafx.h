// stdafx.h : include file for standard system include files,
// or project specific include files that are used frequently,
// but are changed infrequently

#pragma once

#ifndef STRICT
#define STRICT
#endif

#include "targetver.h"

#define _ATL_APARTMENT_THREADED

#define _ATL_NO_AUTOMATIC_NAMESPACE

#define _ATL_CSTRING_EXPLICIT_CONSTRUCTORS	// some CString constructors will be explicit

#define _CRT_NON_CONFORMING_SWPRINTFS
#define _SCL_SECURE_NO_WARNINGS

#define ISOLATION_AWARE_ENABLED 1
#define BOOST_USE_WINDOWS_H


#include <afxwin.h>
#include <afxext.h>
#include <afxole.h>
#include <afxodlgs.h>
#include <afxrich.h>
#include <afxhtml.h>
#include <afxcview.h>
#include <afxwinappex.h>
#include <afxframewndex.h>
#include <afxmdiframewndex.h>

#ifndef _AFX_NO_OLE_SUPPORT
#include <afxdisp.h>        // MFC Automation classes
#endif // _AFX_NO_OLE_SUPPORT

#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlbase.h>
#include <atlcom.h>
#include <atlctl.h>

#define ATL_NO_ASSERT_ON_DESTROY_NONEXISTENT_WINDOW

#include "resource.h"
#include <atlcomcli.h>
#include <atlcom.h>
#include <atlwin.h>
#include <atlenc.h>


#include <shlwapi.h>
#include <ExDisp.h>
#include <ExDispId.h>
#include <mshtml.h>
#include <mshtmdid.h>
#include <mshtmhst.h>
#include <commctrl.h>
#include <shellscalingapi.h>

#include <Dispex.h>

#include <atlstdthunk.h>

#include <comdef.h>

#include <map>
#include <list>
#include <string>
#include <vector>
#include <stack>
#include <utility>
#include <time.h>
#include <Psapi.h>
#include <regex>
#include <functional>
#include <algorithm>
#include <random>
#include <winhttp.h>
#include <iostream> 
#include <sstream> 
#include <fstream>
#include <mutex>
#include <thread>
#include <future>
#include <cstdarg>





#define REGISTRY_TARGET_KEY			L"Software\\AppDataLow\\MyTestOutlookAddin"

#import "libid:AC0714F2-3D04-11D1-AE7D-00A0C90F26F4"\
	auto_rename auto_search raw_interfaces_only rename_namespace("AddinDesign")
using namespace AddinDesign;
//The following #import imports MSO based on it's LIBID

#import "libid:2DF8D04C-5BFA-101B-BDE5-00AA0044DE52"\
	auto_rename auto_search raw_interfaces_only rename_namespace("Office1")
using namespace Office1;

// Forms type library (i.e. MSWORD.olb)
#import "libid:00020905-0000-0000-C000-000000000046"\
	auto_rename auto_search raw_interfaces_only  rename("Font", "FontWord") rename("Pages", "PagesWord") rename_namespace("Word")
using namespace Word;

// Outlook type library (i.e. msoutl.olb)
#import "libid:00062FFF-0000-0000-C000-000000000046"\
	auto_rename auto_search raw_interfaces_only named_guids rename("Font", "FontOutlook") rename("Pages", "PagesOutlook") rename_namespace("Outlook")
using namespace Outlook;

// Forms type library (i.e. fm20.dll)
#import "libid:0D452EE1-E08F-101A-852E-02608C4D0BB4"\
	auto_rename auto_search raw_interfaces_only  rename_namespace("Forms")
using namespace Forms;
