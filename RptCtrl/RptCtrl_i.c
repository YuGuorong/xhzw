

/* this ALWAYS GENERATED file contains the IIDs and CLSIDs */

/* link this file in with the server and any clients */


 /* File created by MIDL compiler version 6.00.0366 */
/* at Wed Nov 24 23:45:16 2010
 */
/* Compiler settings for .\RptCtrl.idl:
    Oicf, W1, Zp8, env=Win32 (32b run)
    protocol : dce , ms_ext, c_ext
    error checks: allocation ref bounds_check enum stub_data 
    VC __declspec() decoration level: 
         __declspec(uuid()), __declspec(selectany), __declspec(novtable)
         DECLSPEC_UUID(), MIDL_INTERFACE()
*/
//@@MIDL_FILE_HEADING(  )

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

MIDL_DEFINE_GUID(IID, LIBID_RptCtrlLib,0x75914D81,0xEAB3,0x49B4,0xB4,0xD2,0x0A,0xD6,0x08,0x58,0x9F,0xB9);


MIDL_DEFINE_GUID(IID, DIID__DRptCtrl,0x615E2844,0x9D51,0x45D3,0x94,0xA3,0xA0,0xB7,0x20,0x80,0x33,0x21);


MIDL_DEFINE_GUID(IID, DIID__DRptCtrlEvents,0xE8D70003,0x4295,0x409A,0xB2,0x93,0x77,0x65,0x51,0x03,0x54,0x4D);


MIDL_DEFINE_GUID(CLSID, CLSID_RptCtrl,0xA7116AA5,0x99DF,0x4310,0x88,0x40,0xD0,0x28,0x69,0x14,0x9B,0x7D);

#undef MIDL_DEFINE_GUID

#ifdef __cplusplus
}
#endif



