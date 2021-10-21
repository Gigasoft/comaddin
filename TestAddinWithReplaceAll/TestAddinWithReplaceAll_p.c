

/* this ALWAYS GENERATED file contains the proxy stub code */


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

#if !defined(_M_IA64) && !defined(_M_AMD64) && !defined(_ARM_)


#pragma warning( disable: 4049 )  /* more than 64k source lines */
#if _MSC_VER >= 1200
#pragma warning(push)
#endif

#pragma warning( disable: 4211 )  /* redefine extern to static */
#pragma warning( disable: 4232 )  /* dllimport identity*/
#pragma warning( disable: 4024 )  /* array to pointer mapping*/
#pragma warning( disable: 4152 )  /* function/data pointer conversion in expression */
#pragma warning( disable: 4100 ) /* unreferenced arguments in x86 call */

#pragma optimize("", off ) 

#define USE_STUBLESS_PROXY


/* verify that the <rpcproxy.h> version is high enough to compile this file*/
#ifndef __REDQ_RPCPROXY_H_VERSION__
#define __REQUIRED_RPCPROXY_H_VERSION__ 475
#endif


#include "rpcproxy.h"
#ifndef __RPCPROXY_H_VERSION__
#error this stub requires an updated version of <rpcproxy.h>
#endif /* __RPCPROXY_H_VERSION__ */


#include "TestAddinWithReplaceAll_i.h"

#define TYPE_FORMAT_STRING_SIZE   69                                
#define PROC_FORMAT_STRING_SIZE   199                               
#define EXPR_FORMAT_STRING_SIZE   1                                 
#define TRANSMIT_AS_TABLE_SIZE    0            
#define WIRE_MARSHAL_TABLE_SIZE   1            

typedef struct _TestAddinWithReplaceAll_MIDL_TYPE_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ TYPE_FORMAT_STRING_SIZE ];
    } TestAddinWithReplaceAll_MIDL_TYPE_FORMAT_STRING;

typedef struct _TestAddinWithReplaceAll_MIDL_PROC_FORMAT_STRING
    {
    short          Pad;
    unsigned char  Format[ PROC_FORMAT_STRING_SIZE ];
    } TestAddinWithReplaceAll_MIDL_PROC_FORMAT_STRING;

typedef struct _TestAddinWithReplaceAll_MIDL_EXPR_FORMAT_STRING
    {
    long          Pad;
    unsigned char  Format[ EXPR_FORMAT_STRING_SIZE ];
    } TestAddinWithReplaceAll_MIDL_EXPR_FORMAT_STRING;


static const RPC_SYNTAX_IDENTIFIER  _RpcTransferSyntax = 
{{0x8A885D04,0x1CEB,0x11C9,{0x9F,0xE8,0x08,0x00,0x2B,0x10,0x48,0x60}},{2,0}};


extern const TestAddinWithReplaceAll_MIDL_TYPE_FORMAT_STRING TestAddinWithReplaceAll__MIDL_TypeFormatString;
extern const TestAddinWithReplaceAll_MIDL_PROC_FORMAT_STRING TestAddinWithReplaceAll__MIDL_ProcFormatString;
extern const TestAddinWithReplaceAll_MIDL_EXPR_FORMAT_STRING TestAddinWithReplaceAll__MIDL_ExprFormatString;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IExt_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IExt_ProxyInfo;


extern const MIDL_STUB_DESC Object_StubDesc;


extern const MIDL_SERVER_INFO IRibbonCallback_ServerInfo;
extern const MIDL_STUBLESS_PROXY_INFO IRibbonCallback_ProxyInfo;


extern const USER_MARSHAL_ROUTINE_QUADRUPLE UserMarshalRoutines[ WIRE_MARSHAL_TABLE_SIZE ];

#if !defined(__RPC_WIN32__)
#error  Invalid build platform for this stub.
#endif

#if !(TARGET_IS_NT50_OR_LATER)
#error You need Windows 2000 or later to run this stub because it uses these features:
#error   /robust command line switch.
#error However, your C/C++ compilation flags indicate you intend to run this app on earlier systems.
#error This app will fail with the RPC_X_WRONG_STUB_VERSION error.
#endif


static const TestAddinWithReplaceAll_MIDL_PROC_FORMAT_STRING TestAddinWithReplaceAll__MIDL_ProcFormatString =
    {
        0,
        {

	/* Procedure OnLoad */

			0x33,		/* FC_AUTO_HANDLE */
			0x6c,		/* Old Flags:  object, Oi2 */
/*  2 */	NdrFcLong( 0x0 ),	/* 0 */
/*  6 */	NdrFcShort( 0x7 ),	/* 7 */
/*  8 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 10 */	NdrFcShort( 0x0 ),	/* 0 */
/* 12 */	NdrFcShort( 0x8 ),	/* 8 */
/* 14 */	0x46,		/* Oi2 Flags:  clt must size, has return, has ext, */
			0x2,		/* 2 */
/* 16 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 18 */	NdrFcShort( 0x0 ),	/* 0 */
/* 20 */	NdrFcShort( 0x0 ),	/* 0 */
/* 22 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter spDispRibbonUI */

/* 24 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 26 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 28 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */

	/* Return value */

/* 30 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 32 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 34 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure ButtonClicked */

/* 36 */	0x33,		/* FC_AUTO_HANDLE */
			0x6c,		/* Old Flags:  object, Oi2 */
/* 38 */	NdrFcLong( 0x0 ),	/* 0 */
/* 42 */	NdrFcShort( 0x8 ),	/* 8 */
/* 44 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 46 */	NdrFcShort( 0x0 ),	/* 0 */
/* 48 */	NdrFcShort( 0x8 ),	/* 8 */
/* 50 */	0x46,		/* Oi2 Flags:  clt must size, has return, has ext, */
			0x2,		/* 2 */
/* 52 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 54 */	NdrFcShort( 0x0 ),	/* 0 */
/* 56 */	NdrFcShort( 0x0 ),	/* 0 */
/* 58 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter RibbonControl */

/* 60 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 62 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 64 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */

	/* Return value */

/* 66 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 68 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 70 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure GetImage */

/* 72 */	0x33,		/* FC_AUTO_HANDLE */
			0x6c,		/* Old Flags:  object, Oi2 */
/* 74 */	NdrFcLong( 0x0 ),	/* 0 */
/* 78 */	NdrFcShort( 0x9 ),	/* 9 */
/* 80 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 82 */	NdrFcShort( 0x0 ),	/* 0 */
/* 84 */	NdrFcShort( 0x8 ),	/* 8 */
/* 86 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x3,		/* 3 */
/* 88 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 90 */	NdrFcShort( 0x0 ),	/* 0 */
/* 92 */	NdrFcShort( 0x0 ),	/* 0 */
/* 94 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter RibbonControl */

/* 96 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 98 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 100 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */

	/* Parameter ppdispImage */

/* 102 */	NdrFcShort( 0x13 ),	/* Flags:  must size, must free, out, */
/* 104 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 106 */	NdrFcShort( 0x14 ),	/* Type Offset=20 */

	/* Return value */

/* 108 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 110 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 112 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure GetEnabled */

/* 114 */	0x33,		/* FC_AUTO_HANDLE */
			0x6c,		/* Old Flags:  object, Oi2 */
/* 116 */	NdrFcLong( 0x0 ),	/* 0 */
/* 120 */	NdrFcShort( 0xa ),	/* 10 */
/* 122 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 124 */	NdrFcShort( 0x0 ),	/* 0 */
/* 126 */	NdrFcShort( 0x22 ),	/* 34 */
/* 128 */	0x46,		/* Oi2 Flags:  clt must size, has return, has ext, */
			0x3,		/* 3 */
/* 130 */	0x8,		/* 8 */
			0x1,		/* Ext Flags:  new corr desc, */
/* 132 */	NdrFcShort( 0x0 ),	/* 0 */
/* 134 */	NdrFcShort( 0x0 ),	/* 0 */
/* 136 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter RibbonControl */

/* 138 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 140 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 142 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */

	/* Parameter pvarfEnabled */

/* 144 */	NdrFcShort( 0x2150 ),	/* Flags:  out, base type, simple ref, srv alloc size=8 */
/* 146 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 148 */	0x6,		/* FC_SHORT */
			0x0,		/* 0 */

	/* Return value */

/* 150 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 152 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 154 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

	/* Procedure GetText */

/* 156 */	0x33,		/* FC_AUTO_HANDLE */
			0x6c,		/* Old Flags:  object, Oi2 */
/* 158 */	NdrFcLong( 0x0 ),	/* 0 */
/* 162 */	NdrFcShort( 0xb ),	/* 11 */
/* 164 */	NdrFcShort( 0x10 ),	/* x86 Stack size/offset = 16 */
/* 166 */	NdrFcShort( 0x0 ),	/* 0 */
/* 168 */	NdrFcShort( 0x8 ),	/* 8 */
/* 170 */	0x47,		/* Oi2 Flags:  srv must size, clt must size, has return, has ext, */
			0x3,		/* 3 */
/* 172 */	0x8,		/* 8 */
			0x3,		/* Ext Flags:  new corr desc, clt corr check, */
/* 174 */	NdrFcShort( 0x1 ),	/* 1 */
/* 176 */	NdrFcShort( 0x0 ),	/* 0 */
/* 178 */	NdrFcShort( 0x0 ),	/* 0 */

	/* Parameter RibbonControl */

/* 180 */	NdrFcShort( 0xb ),	/* Flags:  must size, must free, in, */
/* 182 */	NdrFcShort( 0x4 ),	/* x86 Stack size/offset = 4 */
/* 184 */	NdrFcShort( 0x2 ),	/* Type Offset=2 */

	/* Parameter pVal */

/* 186 */	NdrFcShort( 0x2113 ),	/* Flags:  must size, must free, out, simple ref, srv alloc size=8 */
/* 188 */	NdrFcShort( 0x8 ),	/* x86 Stack size/offset = 8 */
/* 190 */	NdrFcShort( 0x3a ),	/* Type Offset=58 */

	/* Return value */

/* 192 */	NdrFcShort( 0x70 ),	/* Flags:  out, return, base type, */
/* 194 */	NdrFcShort( 0xc ),	/* x86 Stack size/offset = 12 */
/* 196 */	0x8,		/* FC_LONG */
			0x0,		/* 0 */

			0x0
        }
    };

static const TestAddinWithReplaceAll_MIDL_TYPE_FORMAT_STRING TestAddinWithReplaceAll__MIDL_TypeFormatString =
    {
        0,
        {
			NdrFcShort( 0x0 ),	/* 0 */
/*  2 */	
			0x2f,		/* FC_IP */
			0x5a,		/* FC_CONSTANT_IID */
/*  4 */	NdrFcLong( 0x20400 ),	/* 132096 */
/*  8 */	NdrFcShort( 0x0 ),	/* 0 */
/* 10 */	NdrFcShort( 0x0 ),	/* 0 */
/* 12 */	0xc0,		/* 192 */
			0x0,		/* 0 */
/* 14 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 16 */	0x0,		/* 0 */
			0x0,		/* 0 */
/* 18 */	0x0,		/* 0 */
			0x46,		/* 70 */
/* 20 */	
			0x11, 0x10,	/* FC_RP [pointer_deref] */
/* 22 */	NdrFcShort( 0xffec ),	/* Offset= -20 (2) */
/* 24 */	
			0x11, 0xc,	/* FC_RP [alloced_on_stack] [simple_pointer] */
/* 26 */	0x6,		/* FC_SHORT */
			0x5c,		/* FC_PAD */
/* 28 */	
			0x11, 0x4,	/* FC_RP [alloced_on_stack] */
/* 30 */	NdrFcShort( 0x1c ),	/* Offset= 28 (58) */
/* 32 */	
			0x13, 0x0,	/* FC_OP */
/* 34 */	NdrFcShort( 0xe ),	/* Offset= 14 (48) */
/* 36 */	
			0x1b,		/* FC_CARRAY */
			0x1,		/* 1 */
/* 38 */	NdrFcShort( 0x2 ),	/* 2 */
/* 40 */	0x9,		/* Corr desc: FC_ULONG */
			0x0,		/*  */
/* 42 */	NdrFcShort( 0xfffc ),	/* -4 */
/* 44 */	NdrFcShort( 0x1 ),	/* Corr flags:  early, */
/* 46 */	0x6,		/* FC_SHORT */
			0x5b,		/* FC_END */
/* 48 */	
			0x17,		/* FC_CSTRUCT */
			0x3,		/* 3 */
/* 50 */	NdrFcShort( 0x8 ),	/* 8 */
/* 52 */	NdrFcShort( 0xfff0 ),	/* Offset= -16 (36) */
/* 54 */	0x8,		/* FC_LONG */
			0x8,		/* FC_LONG */
/* 56 */	0x5c,		/* FC_PAD */
			0x5b,		/* FC_END */
/* 58 */	0xb4,		/* FC_USER_MARSHAL */
			0x83,		/* 131 */
/* 60 */	NdrFcShort( 0x0 ),	/* 0 */
/* 62 */	NdrFcShort( 0x4 ),	/* 4 */
/* 64 */	NdrFcShort( 0x0 ),	/* 0 */
/* 66 */	NdrFcShort( 0xffde ),	/* Offset= -34 (32) */

			0x0
        }
    };

static const USER_MARSHAL_ROUTINE_QUADRUPLE UserMarshalRoutines[ WIRE_MARSHAL_TABLE_SIZE ] = 
        {
            
            {
            BSTR_UserSize
            ,BSTR_UserMarshal
            ,BSTR_UserUnmarshal
            ,BSTR_UserFree
            }

        };



/* Object interface: IUnknown, ver. 0.0,
   GUID={0x00000000,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IDispatch, ver. 0.0,
   GUID={0x00020400,0x0000,0x0000,{0xC0,0x00,0x00,0x00,0x00,0x00,0x00,0x46}} */


/* Object interface: IExt, ver. 0.0,
   GUID={0x47939ABD,0x7189,0x4C17,{0xB5,0x5B,0x3D,0x73,0xB4,0x4F,0xB7,0x0C}} */

#pragma code_seg(".orpc")
static const unsigned short IExt_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0
    };

static const MIDL_STUBLESS_PROXY_INFO IExt_ProxyInfo =
    {
    &Object_StubDesc,
    TestAddinWithReplaceAll__MIDL_ProcFormatString.Format,
    &IExt_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IExt_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    TestAddinWithReplaceAll__MIDL_ProcFormatString.Format,
    &IExt_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(7) _IExtProxyVtbl = 
{
    0,
    &IID_IExt,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */
};


static const PRPC_STUB_FUNCTION IExt_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION
};

CInterfaceStubVtbl _IExtStubVtbl =
{
    &IID_IExt,
    &IExt_ServerInfo,
    7,
    &IExt_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};


/* Object interface: IRibbonCallback, ver. 0.0,
   GUID={0x38B6566D,0x1DFA,0x4279,{0xA2,0x82,0xDE,0x50,0xFD,0xBC,0x76,0xFD}} */

#pragma code_seg(".orpc")
static const unsigned short IRibbonCallback_FormatStringOffsetTable[] =
    {
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    (unsigned short) -1,
    0,
    36,
    72,
    114,
    156
    };

static const MIDL_STUBLESS_PROXY_INFO IRibbonCallback_ProxyInfo =
    {
    &Object_StubDesc,
    TestAddinWithReplaceAll__MIDL_ProcFormatString.Format,
    &IRibbonCallback_FormatStringOffsetTable[-3],
    0,
    0,
    0
    };


static const MIDL_SERVER_INFO IRibbonCallback_ServerInfo = 
    {
    &Object_StubDesc,
    0,
    TestAddinWithReplaceAll__MIDL_ProcFormatString.Format,
    &IRibbonCallback_FormatStringOffsetTable[-3],
    0,
    0,
    0,
    0};
CINTERFACE_PROXY_VTABLE(12) _IRibbonCallbackProxyVtbl = 
{
    &IRibbonCallback_ProxyInfo,
    &IID_IRibbonCallback,
    IUnknown_QueryInterface_Proxy,
    IUnknown_AddRef_Proxy,
    IUnknown_Release_Proxy ,
    0 /* IDispatch::GetTypeInfoCount */ ,
    0 /* IDispatch::GetTypeInfo */ ,
    0 /* IDispatch::GetIDsOfNames */ ,
    0 /* IDispatch_Invoke_Proxy */ ,
    (void *) (INT_PTR) -1 /* IRibbonCallback::OnLoad */ ,
    (void *) (INT_PTR) -1 /* IRibbonCallback::ButtonClicked */ ,
    (void *) (INT_PTR) -1 /* IRibbonCallback::GetImage */ ,
    (void *) (INT_PTR) -1 /* IRibbonCallback::GetEnabled */ ,
    (void *) (INT_PTR) -1 /* IRibbonCallback::GetText */
};


static const PRPC_STUB_FUNCTION IRibbonCallback_table[] =
{
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    STUB_FORWARDING_FUNCTION,
    NdrStubCall2,
    NdrStubCall2,
    NdrStubCall2,
    NdrStubCall2,
    NdrStubCall2
};

CInterfaceStubVtbl _IRibbonCallbackStubVtbl =
{
    &IID_IRibbonCallback,
    &IRibbonCallback_ServerInfo,
    12,
    &IRibbonCallback_table[-3],
    CStdStubBuffer_DELEGATING_METHODS
};

static const MIDL_STUB_DESC Object_StubDesc = 
    {
    0,
    NdrOleAllocate,
    NdrOleFree,
    0,
    0,
    0,
    0,
    0,
    TestAddinWithReplaceAll__MIDL_TypeFormatString.Format,
    1, /* -error bounds_check flag */
    0x50002, /* Ndr library version */
    0,
    0x800025b, /* MIDL Version 8.0.603 */
    0,
    UserMarshalRoutines,
    0,  /* notify & notify_flag routine table */
    0x1, /* MIDL flag */
    0, /* cs routines */
    0,   /* proxy/server info */
    0
    };

const CInterfaceProxyVtbl * const _TestAddinWithReplaceAll_ProxyVtblList[] = 
{
    ( CInterfaceProxyVtbl *) &_IRibbonCallbackProxyVtbl,
    ( CInterfaceProxyVtbl *) &_IExtProxyVtbl,
    0
};

const CInterfaceStubVtbl * const _TestAddinWithReplaceAll_StubVtblList[] = 
{
    ( CInterfaceStubVtbl *) &_IRibbonCallbackStubVtbl,
    ( CInterfaceStubVtbl *) &_IExtStubVtbl,
    0
};

PCInterfaceName const _TestAddinWithReplaceAll_InterfaceNamesList[] = 
{
    "IRibbonCallback",
    "IExt",
    0
};

const IID *  const _TestAddinWithReplaceAll_BaseIIDList[] = 
{
    &IID_IDispatch,
    &IID_IDispatch,
    0
};


#define _TestAddinWithReplaceAll_CHECK_IID(n)	IID_GENERIC_CHECK_IID( _TestAddinWithReplaceAll, pIID, n)

int __stdcall _TestAddinWithReplaceAll_IID_Lookup( const IID * pIID, int * pIndex )
{
    IID_BS_LOOKUP_SETUP

    IID_BS_LOOKUP_INITIAL_TEST( _TestAddinWithReplaceAll, 2, 1 )
    IID_BS_LOOKUP_RETURN_RESULT( _TestAddinWithReplaceAll, 2, *pIndex )
    
}

const ExtendedProxyFileInfo TestAddinWithReplaceAll_ProxyFileInfo = 
{
    (PCInterfaceProxyVtblList *) & _TestAddinWithReplaceAll_ProxyVtblList,
    (PCInterfaceStubVtblList *) & _TestAddinWithReplaceAll_StubVtblList,
    (const PCInterfaceName * ) & _TestAddinWithReplaceAll_InterfaceNamesList,
    (const IID ** ) & _TestAddinWithReplaceAll_BaseIIDList,
    & _TestAddinWithReplaceAll_IID_Lookup, 
    2,
    2,
    0, /* table of [async_uuid] interfaces */
    0, /* Filler1 */
    0, /* Filler2 */
    0  /* Filler3 */
};
#pragma optimize("", on )
#if _MSC_VER >= 1200
#pragma warning(pop)
#endif


#endif /* !defined(_M_IA64) && !defined(_M_AMD64) && !defined(_ARM_) */

