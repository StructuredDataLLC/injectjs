

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Wed Apr 16 15:17:43 2014
 */
/* Compiler settings for CV82.idl:
    Oicf, W1, Zp8, env=Win64 (32b run), target_arch=AMD64 8.00.0603 
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

MIDL_DEFINE_GUID(IID, IID_IScripto,0x9038E84A,0x1885,0x4A40,0x8B,0x3D,0x9D,0x96,0x7A,0xF4,0x0E,0xDF);


MIDL_DEFINE_GUID(IID, IID_ICVEventHandler,0xCB998305,0xB733,0x429F,0xA3,0x0D,0xA8,0x65,0xB8,0x6F,0x0B,0xB8);


MIDL_DEFINE_GUID(IID, LIBID_CV82Lib,0xEEAEE929,0xA29C,0x48AC,0xA1,0x5E,0x92,0xCE,0x89,0x88,0x1F,0x80);


MIDL_DEFINE_GUID(IID, DIID__IScriptoEvents,0x325D8FE4,0xF4A1,0x429F,0xB4,0x93,0x85,0x10,0x99,0xB1,0x00,0xFD);


MIDL_DEFINE_GUID(CLSID, CLSID_Scripto,0x03A48679,0xF8A2,0x4AFE,0x8D,0x38,0x42,0x8C,0x21,0x91,0xC0,0xF7);


MIDL_DEFINE_GUID(CLSID, CLSID_CVEventHandler,0x76CEE870,0xE6E0,0x4B70,0x85,0xC7,0x5E,0xA2,0x00,0xF6,0xD6,0xDA);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



