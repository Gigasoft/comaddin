

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


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


#ifdef __cplusplus
extern "C"{
#endif 


#include <rpc.h>
#include <rpcndr.h>

#ifdef _MIDL_USE_GUIDDEF_

#ifndef INITGUID
#define INITGUID
#include <guiddef.h>
#undef INITGUID
#else
#include <guiddef.h>
#endif

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        DEFINE_GUID(name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8)

#else // !_MIDL_USE_GUIDDEF_

#ifndef __IID_DEFINED__
#define __IID_DEFINED__

typedef struct _IID
{
    unsigned long x;
    unsigned short s1;
    unsigned short s2;
    unsigned char  c[8];
} IID;

#endif // __IID_DEFINED__

#ifndef CLSID_DEFINED
#define CLSID_DEFINED
typedef IID CLSID;
#endif // CLSID_DEFINED

#define MIDL_DEFINE_GUID(type,name,l,w1,w2,b1,b2,b3,b4,b5,b6,b7,b8) \
        const type name = {l,w1,w2,{b1,b2,b3,b4,b5,b6,b7,b8}}

#endif !_MIDL_USE_GUIDDEF_

MIDL_DEFINE_GUID(IID, IID_IExt,0x47939ABD,0x7189,0x4C17,0xB5,0x5B,0x3D,0x73,0xB4,0x4F,0xB7,0x0C);


MIDL_DEFINE_GUID(IID, IID_IRibbonCallback,0x38B6566D,0x1DFA,0x4279,0xA2,0x82,0xDE,0x50,0xFD,0xBC,0x76,0xFD);


MIDL_DEFINE_GUID(IID, LIBID_TestAddinWithReplaceAllLib,0x1873E461,0x6CD7,0x4338,0x9D,0x66,0xCE,0x99,0xB3,0xFB,0x30,0x12);


MIDL_DEFINE_GUID(CLSID, CLSID_Ext,0xC00E7DE9,0x854A,0x4BA7,0xA0,0x08,0x33,0xA8,0x11,0x03,0x4C,0x88);


MIDL_DEFINE_GUID(CLSID, CLSID_RibbonCallback,0x22437F6E,0x068F,0x4D18,0x92,0xB0,0x39,0xEF,0x2E,0xC2,0x98,0x8A);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



