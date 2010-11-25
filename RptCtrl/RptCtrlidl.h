

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


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


/* verify that the <rpcndr.h> version is high enough to compile this file*/
#ifndef __REQUIRED_RPCNDR_H_VERSION__
#define __REQUIRED_RPCNDR_H_VERSION__ 440
#endif

#include "rpc.h"
#include "rpcndr.h"

#ifndef __RptCtrlidl_h__
#define __RptCtrlidl_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef ___DRptCtrl_FWD_DEFINED__
#define ___DRptCtrl_FWD_DEFINED__
typedef interface _DRptCtrl _DRptCtrl;
#endif 	/* ___DRptCtrl_FWD_DEFINED__ */


#ifndef ___DRptCtrlEvents_FWD_DEFINED__
#define ___DRptCtrlEvents_FWD_DEFINED__
typedef interface _DRptCtrlEvents _DRptCtrlEvents;
#endif 	/* ___DRptCtrlEvents_FWD_DEFINED__ */


#ifndef __RptCtrl_FWD_DEFINED__
#define __RptCtrl_FWD_DEFINED__

#ifdef __cplusplus
typedef class RptCtrl RptCtrl;
#else
typedef struct RptCtrl RptCtrl;
#endif /* __cplusplus */

#endif 	/* __RptCtrl_FWD_DEFINED__ */


#ifdef __cplusplus
extern "C"{
#endif 

void * __RPC_USER MIDL_user_allocate(size_t);
void __RPC_USER MIDL_user_free( void * ); 


#ifndef __RptCtrlLib_LIBRARY_DEFINED__
#define __RptCtrlLib_LIBRARY_DEFINED__

/* library RptCtrlLib */
/* [control][helpstring][helpfile][version][uuid] */ 


EXTERN_C const IID LIBID_RptCtrlLib;

#ifndef ___DRptCtrl_DISPINTERFACE_DEFINED__
#define ___DRptCtrl_DISPINTERFACE_DEFINED__

/* dispinterface _DRptCtrl */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__DRptCtrl;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("615E2844-9D51-45D3-94A3-A0B720803321")
    _DRptCtrl : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DRptCtrlVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _DRptCtrl * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _DRptCtrl * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _DRptCtrl * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _DRptCtrl * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _DRptCtrl * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _DRptCtrl * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _DRptCtrl * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } _DRptCtrlVtbl;

    interface _DRptCtrl
    {
        CONST_VTBL struct _DRptCtrlVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DRptCtrl_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DRptCtrl_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DRptCtrl_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DRptCtrl_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DRptCtrl_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DRptCtrl_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DRptCtrl_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DRptCtrl_DISPINTERFACE_DEFINED__ */


#ifndef ___DRptCtrlEvents_DISPINTERFACE_DEFINED__
#define ___DRptCtrlEvents_DISPINTERFACE_DEFINED__

/* dispinterface _DRptCtrlEvents */
/* [helpstring][uuid] */ 


EXTERN_C const IID DIID__DRptCtrlEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("E8D70003-4295-409A-B293-77655103544D")
    _DRptCtrlEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _DRptCtrlEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _DRptCtrlEvents * This,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _DRptCtrlEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _DRptCtrlEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _DRptCtrlEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _DRptCtrlEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _DRptCtrlEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _DRptCtrlEvents * This,
            /* [in] */ DISPID dispIdMember,
            /* [in] */ REFIID riid,
            /* [in] */ LCID lcid,
            /* [in] */ WORD wFlags,
            /* [out][in] */ DISPPARAMS *pDispParams,
            /* [out] */ VARIANT *pVarResult,
            /* [out] */ EXCEPINFO *pExcepInfo,
            /* [out] */ UINT *puArgErr);
        
        END_INTERFACE
    } _DRptCtrlEventsVtbl;

    interface _DRptCtrlEvents
    {
        CONST_VTBL struct _DRptCtrlEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _DRptCtrlEvents_QueryInterface(This,riid,ppvObject)	\
    (This)->lpVtbl -> QueryInterface(This,riid,ppvObject)

#define _DRptCtrlEvents_AddRef(This)	\
    (This)->lpVtbl -> AddRef(This)

#define _DRptCtrlEvents_Release(This)	\
    (This)->lpVtbl -> Release(This)


#define _DRptCtrlEvents_GetTypeInfoCount(This,pctinfo)	\
    (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo)

#define _DRptCtrlEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo)

#define _DRptCtrlEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)

#define _DRptCtrlEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___DRptCtrlEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_RptCtrl;

#ifdef __cplusplus

class DECLSPEC_UUID("A7116AA5-99DF-4310-8840-D02869149B7D")
RptCtrl;
#endif
#endif /* __RptCtrlLib_LIBRARY_DEFINED__ */

/* Additional Prototypes for ALL interfaces */

/* end of Additional Prototypes */

#ifdef __cplusplus
}
#endif

#endif


