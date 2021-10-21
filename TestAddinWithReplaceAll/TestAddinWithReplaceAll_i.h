

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Thu Oct 21 15:58:22 2021
 */
/* Compiler settings for TestAddinWithReplaceAll.idl:
    Oicf, W1, Zp8, env=Win32 (32b run), target_arch=X86 8.00.0603 
    protocol : dce , ms_ext, c_ext, robust
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
/* @@MIDL_FILE_HEADING(  ) */

#pragma warning( disable: 4049 )  /* more than 64k source lines */


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 475
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RPCNDR_H_VERSION__
#error this stub requires an updated version of <rpcndr.h>
#endif // __RPCNDR_H_VERSION__

#ifndef COM_NO_WINDOWS_H
#include "windows.h"
#include "ole2.h"
#endif /*COM_NO_WINDOWS_H*/

#ifndef __TestAddinWithReplaceAll_i_h__
#define __TestAddinWithReplaceAll_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IExt_FWD_DEFINED__
#define __IExt_FWD_DEFINED__
typedef interface IExt IExt;

#endif 	/* __IExt_FWD_DEFINED__ */


#ifndef __IRibbonCallback_FWD_DEFINED__
#define __IRibbonCallback_FWD_DEFINED__
typedef interface IRibbonCallback IRibbonCallback;

#endif 	/* __IRibbonCallback_FWD_DEFINED__ */


#ifndef __Ext_FWD_DEFINED__
#define __Ext_FWD_DEFINED__

#ifdef __cplusplus
typedef class Ext Ext;
#else
typedef struct Ext Ext;
#endif /* __cplusplus */

#endif 	/* __Ext_FWD_DEFINED__ */


#ifndef __RibbonCallback_FWD_DEFINED__
#define __RibbonCallback_FWD_DEFINED__

#ifdef __cplusplus
typedef class RibbonCallback RibbonCallback;
#else
typedef struct RibbonCallback RibbonCallback;
#endif /* __cplusplus */

#endif 	/* __RibbonCallback_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"
#include "shobjidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IExt_INTERFACE_DEFINED__
#define __IExt_INTERFACE_DEFINED__

/* interface IExt */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IExt;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("47939ABD-7189-4C17-B55B-3D73B44FB70C")
    IExt : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct IExtVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IExt * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IExt * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IExt * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IExt * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IExt * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IExt * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IExt * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        END_INTERFACE
    } IExtVtbl;

    interface IExt
    {
        CONST_VTBL struct IExtVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IExt_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IExt_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IExt_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IExt_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IExt_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IExt_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IExt_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IExt_INTERFACE_DEFINED__ */


#ifndef __IRibbonCallback_INTERFACE_DEFINED__
#define __IRibbonCallback_INTERFACE_DEFINED__

/* interface IRibbonCallback */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IRibbonCallback;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("38B6566D-1DFA-4279-A282-DE50FDBC76FD")
    IRibbonCallback : public IDispatch
    {
    public:
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE OnLoad( 
            /* [in] */ IDispatch *spDispRibbonUI) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE ButtonClicked( 
            /* [in] */ IDispatch *RibbonControl) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetImage( 
            /* [in] */ IDispatch *RibbonControl,
            /* [retval][out] */ IDispatch **ppdispImage) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetEnabled( 
            /* [in] */ IDispatch *RibbonControl,
            /* [retval][out] */ VARIANT_BOOL *pvarfEnabled) = 0;
        
        virtual /* [helpstring][id] */ HRESULT STDMETHODCALLTYPE GetText( 
            /* [in] */ IDispatch *RibbonControl,
            /* [retval][out] */ BSTR *pVal) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IRibbonCallbackVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IRibbonCallback * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IRibbonCallback * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IRibbonCallback * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IRibbonCallback * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IRibbonCallback * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IRibbonCallback * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IRibbonCallback * This,
            /* [annotation][in] */ 
            _In_  DISPID dispIdMember,
            /* [annotation][in] */ 
            _In_  REFIID riid,
            /* [annotation][in] */ 
            _In_  LCID lcid,
            /* [annotation][in] */ 
            _In_  WORD wFlags,
            /* [annotation][out][in] */ 
            _In_  DISPPARAMS *pDispParams,
            /* [annotation][out] */ 
            _Out_opt_  VARIANT *pVarResult,
            /* [annotation][out] */ 
            _Out_opt_  EXCEPINFO *pExcepInfo,
            /* [annotation][out] */ 
            _Out_opt_  UINT *puArgErr);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *OnLoad )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *spDispRibbonUI);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *ButtonClicked )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *RibbonControl);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetImage )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *RibbonControl,
            /* [retval][out] */ IDispatch **ppdispImage);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetEnabled )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *RibbonControl,
            /* [retval][out] */ VARIANT_BOOL *pvarfEnabled);
        
        /* [helpstring][id] */ HRESULT ( STDMETHODCALLTYPE *GetText )( 
            IRibbonCallback * This,
            /* [in] */ IDispatch *RibbonControl,
            /* [retval][out] */ BSTR *pVal);
        
        END_INTERFACE
    } IRibbonCallbackVtbl;

    interface IRibbonCallback
    {
        CONST_VTBL struct IRibbonCallbackVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IRibbonCallback_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IRibbonCallback_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IRibbonCallback_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IRibbonCallback_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IRibbonCallback_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IRibbonCallback_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IRibbonCallback_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IRibbonCallback_OnLoad(This,spDispRibbonUI)	\
    ( (This)->lpVtbl -> OnLoad(This,spDispRibbonUI) ) 

#define IRibbonCallback_ButtonClicked(This,RibbonControl)	\
    ( (This)->lpVtbl -> ButtonClicked(This,RibbonControl) ) 

#define IRibbonCallback_GetImage(This,RibbonControl,ppdispImage)	\
    ( (This)->lpVtbl -> GetImage(This,RibbonControl,ppdispImage) ) 

#define IRibbonCallback_GetEnabled(This,RibbonControl,pvarfEnabled)	\
    ( (This)->lpVtbl -> GetEnabled(This,RibbonControl,pvarfEnabled) ) 

#define IRibbonCallback_GetText(This,RibbonControl,pVal)	\
    ( (This)->lpVtbl -> GetText(This,RibbonControl,pVal) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IRibbonCallback_INTERFACE_DEFINED__ */



#ifndef __TestAddinWithReplaceAllLib_LIBRARY_DEFINED__
#define __TestAddinWithReplaceAllLib_LIBRARY_DEFINED__

/* library TestAddinWithReplaceAllLib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_TestAddinWithReplaceAllLib;

EXTERN_C const CLSID CLSID_Ext;

#ifdef __cplusplus

class DECLSPEC_UUID("C00E7DE9-854A-4BA7-A008-33A811034C88")
Ext;
#endif

EXTERN_C const CLSID CLSID_RibbonCallback;

#ifdef __cplusplus

class DECLSPEC_UUID("22437F6E-068F-4D18-92B0-39EF2EC2988A")
RibbonCallback;
#endif
#endif /* __TestAddinWithReplaceAllLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

unsigned long             __RPC_USER  BSTR_UserSize(     unsigned long *, unsigned long            , BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserMarshal(  unsigned long *, unsigned char *, BSTR * ); 
unsigned char * __RPC_USER  BSTR_UserUnmarshal(unsigned long *, unsigned char *, BSTR * ); 
void                      __RPC_USER  BSTR_UserFree(     unsigned long *, BSTR * ); 

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


