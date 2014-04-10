

/* this ALWAYS GENERATED file contains the definitions for the interfaces */


 /* File created by MIDL compiler version 8.00.0603 */
/* at Wed Apr 09 17:17:54 2014
 */
/* Compiler settings for CV82.idl:
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

#ifndef __CV82_i_h__
#define __CV82_i_h__

#if defined(_MSC_VER) && (_MSC_VER >= 1020)
#pragma once
#endif

/* Forward Declarations */ 

#ifndef __IScripto_FWD_DEFINED__
#define __IScripto_FWD_DEFINED__
typedef interface IScripto IScripto;

#endif 	/* __IScripto_FWD_DEFINED__ */


#ifndef __ICVEventHandler_FWD_DEFINED__
#define __ICVEventHandler_FWD_DEFINED__
typedef interface ICVEventHandler ICVEventHandler;

#endif 	/* __ICVEventHandler_FWD_DEFINED__ */


#ifndef ___IScriptoEvents_FWD_DEFINED__
#define ___IScriptoEvents_FWD_DEFINED__
typedef interface _IScriptoEvents _IScriptoEvents;

#endif 	/* ___IScriptoEvents_FWD_DEFINED__ */


#ifndef __Scripto_FWD_DEFINED__
#define __Scripto_FWD_DEFINED__

#ifdef __cplusplus
typedef class Scripto Scripto;
#else
typedef struct Scripto Scripto;
#endif /* __cplusplus */

#endif 	/* __Scripto_FWD_DEFINED__ */


#ifndef __CVEventHandler_FWD_DEFINED__
#define __CVEventHandler_FWD_DEFINED__

#ifdef __cplusplus
typedef class CVEventHandler CVEventHandler;
#else
typedef struct CVEventHandler CVEventHandler;
#endif /* __cplusplus */

#endif 	/* __CVEventHandler_FWD_DEFINED__ */


/* header files for imported files */
#include "oaidl.h"
#include "ocidl.h"

#ifdef __cplusplus
extern "C"{
#endif 


#ifndef __IScripto_INTERFACE_DEFINED__
#define __IScripto_INTERFACE_DEFINED__

/* interface IScripto */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_IScripto;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("9038E84A-1885-4A40-8B3D-9D967AF40EDF")
    IScripto : public IDispatch
    {
    public:
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE SetDispatch( 
            /* [in] */ IDispatch *Dispatch,
            /* [in] */ BSTR *Name) = 0;
        
        virtual /* [id] */ HRESULT STDMETHODCALLTYPE ExecString( 
            /* [in] */ BSTR *Script,
            /* [out] */ BSTR *Result,
            /* [retval][out] */ VARIANT_BOOL *Success) = 0;
        
    };
    
    
#else 	/* C style interface */

    typedef struct IScriptoVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            IScripto * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            IScripto * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            IScripto * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            IScripto * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            IScripto * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            IScripto * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            IScripto * This,
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
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *SetDispatch )( 
            IScripto * This,
            /* [in] */ IDispatch *Dispatch,
            /* [in] */ BSTR *Name);
        
        /* [id] */ HRESULT ( STDMETHODCALLTYPE *ExecString )( 
            IScripto * This,
            /* [in] */ BSTR *Script,
            /* [out] */ BSTR *Result,
            /* [retval][out] */ VARIANT_BOOL *Success);
        
        END_INTERFACE
    } IScriptoVtbl;

    interface IScripto
    {
        CONST_VTBL struct IScriptoVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define IScripto_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define IScripto_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define IScripto_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define IScripto_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define IScripto_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define IScripto_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define IScripto_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#define IScripto_SetDispatch(This,Dispatch,Name)	\
    ( (This)->lpVtbl -> SetDispatch(This,Dispatch,Name) ) 

#define IScripto_ExecString(This,Script,Result,Success)	\
    ( (This)->lpVtbl -> ExecString(This,Script,Result,Success) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __IScripto_INTERFACE_DEFINED__ */


#ifndef __ICVEventHandler_INTERFACE_DEFINED__
#define __ICVEventHandler_INTERFACE_DEFINED__

/* interface ICVEventHandler */
/* [unique][nonextensible][dual][uuid][object] */ 


EXTERN_C const IID IID_ICVEventHandler;

#if defined(__cplusplus) && !defined(CINTERFACE)
    
    MIDL_INTERFACE("CB998305-B733-429F-A30D-A865B86F0BB8")
    ICVEventHandler : public IDispatch
    {
    public:
    };
    
    
#else 	/* C style interface */

    typedef struct ICVEventHandlerVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            ICVEventHandler * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            ICVEventHandler * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            ICVEventHandler * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            ICVEventHandler * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            ICVEventHandler * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            ICVEventHandler * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            ICVEventHandler * This,
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
    } ICVEventHandlerVtbl;

    interface ICVEventHandler
    {
        CONST_VTBL struct ICVEventHandlerVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define ICVEventHandler_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define ICVEventHandler_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define ICVEventHandler_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define ICVEventHandler_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define ICVEventHandler_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define ICVEventHandler_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define ICVEventHandler_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 


#endif /* COBJMACROS */


#endif 	/* C style interface */




#endif 	/* __ICVEventHandler_INTERFACE_DEFINED__ */



#ifndef __CV82Lib_LIBRARY_DEFINED__
#define __CV82Lib_LIBRARY_DEFINED__

/* library CV82Lib */
/* [version][uuid] */ 


EXTERN_C const IID LIBID_CV82Lib;

#ifndef ___IScriptoEvents_DISPINTERFACE_DEFINED__
#define ___IScriptoEvents_DISPINTERFACE_DEFINED__

/* dispinterface _IScriptoEvents */
/* [uuid] */ 


EXTERN_C const IID DIID__IScriptoEvents;

#if defined(__cplusplus) && !defined(CINTERFACE)

    MIDL_INTERFACE("325D8FE4-F4A1-429F-B493-851099B100FD")
    _IScriptoEvents : public IDispatch
    {
    };
    
#else 	/* C style interface */

    typedef struct _IScriptoEventsVtbl
    {
        BEGIN_INTERFACE
        
        HRESULT ( STDMETHODCALLTYPE *QueryInterface )( 
            _IScriptoEvents * This,
            /* [in] */ REFIID riid,
            /* [annotation][iid_is][out] */ 
            _COM_Outptr_  void **ppvObject);
        
        ULONG ( STDMETHODCALLTYPE *AddRef )( 
            _IScriptoEvents * This);
        
        ULONG ( STDMETHODCALLTYPE *Release )( 
            _IScriptoEvents * This);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfoCount )( 
            _IScriptoEvents * This,
            /* [out] */ UINT *pctinfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetTypeInfo )( 
            _IScriptoEvents * This,
            /* [in] */ UINT iTInfo,
            /* [in] */ LCID lcid,
            /* [out] */ ITypeInfo **ppTInfo);
        
        HRESULT ( STDMETHODCALLTYPE *GetIDsOfNames )( 
            _IScriptoEvents * This,
            /* [in] */ REFIID riid,
            /* [size_is][in] */ LPOLESTR *rgszNames,
            /* [range][in] */ UINT cNames,
            /* [in] */ LCID lcid,
            /* [size_is][out] */ DISPID *rgDispId);
        
        /* [local] */ HRESULT ( STDMETHODCALLTYPE *Invoke )( 
            _IScriptoEvents * This,
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
    } _IScriptoEventsVtbl;

    interface _IScriptoEvents
    {
        CONST_VTBL struct _IScriptoEventsVtbl *lpVtbl;
    };

    

#ifdef COBJMACROS


#define _IScriptoEvents_QueryInterface(This,riid,ppvObject)	\
    ( (This)->lpVtbl -> QueryInterface(This,riid,ppvObject) ) 

#define _IScriptoEvents_AddRef(This)	\
    ( (This)->lpVtbl -> AddRef(This) ) 

#define _IScriptoEvents_Release(This)	\
    ( (This)->lpVtbl -> Release(This) ) 


#define _IScriptoEvents_GetTypeInfoCount(This,pctinfo)	\
    ( (This)->lpVtbl -> GetTypeInfoCount(This,pctinfo) ) 

#define _IScriptoEvents_GetTypeInfo(This,iTInfo,lcid,ppTInfo)	\
    ( (This)->lpVtbl -> GetTypeInfo(This,iTInfo,lcid,ppTInfo) ) 

#define _IScriptoEvents_GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId)	\
    ( (This)->lpVtbl -> GetIDsOfNames(This,riid,rgszNames,cNames,lcid,rgDispId) ) 

#define _IScriptoEvents_Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr)	\
    ( (This)->lpVtbl -> Invoke(This,dispIdMember,riid,lcid,wFlags,pDispParams,pVarResult,pExcepInfo,puArgErr) ) 

#endif /* COBJMACROS */


#endif 	/* C style interface */


#endif 	/* ___IScriptoEvents_DISPINTERFACE_DEFINED__ */


EXTERN_C const CLSID CLSID_Scripto;

#ifdef __cplusplus

class DECLSPEC_UUID("03A48679-F8A2-4AFE-8D38-428C2191C0F7")
Scripto;
#endif

EXTERN_C const CLSID CLSID_CVEventHandler;

#ifdef __cplusplus

class DECLSPEC_UUID("76CEE870-E6E0-4B70-85C7-5EA200F6D6DA")
CVEventHandler;
#endif
#endif /* __CV82Lib_LIBRARY_DEFINED__ */

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


