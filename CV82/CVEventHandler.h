// CVEventHandler.h : Declaration of the CCVEventHandler

#pragma once
#include "resource.h"       // main symbols



#include "CV82_i.h"
#include <v8.h>
#include <string>
#include <hash_map>

class CScripto; // fwd



#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif

using namespace ATL;

/** default traits are non-copyable, doesn't play well with stl containers */
typedef v8::Persistent<v8::Function, v8::CopyablePersistentTraits<v8::Function>> CopyablePersistent;

// CCVEventHandler

class ATL_NO_VTABLE CCVEventHandler :
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CCVEventHandler, &CLSID_CVEventHandler>,
	public IDispatchImpl<ICVEventHandler, &IID_ICVEventHandler, &LIBID_CV82Lib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CCVEventHandler()
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_CVEVENTHANDLER)


BEGIN_COM_MAP(CCVEventHandler)
	COM_INTERFACE_ENTRY(ICVEventHandler)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:

	std::hash_map < long, CopyablePersistent > func_map;
	CScripto *ptr;

	/**
	 * store a pointer to a function by dispid.  a single instance can be used
	 * for all sinks into this class, although at the moment we're not supporting
	 * multiple callbacks (should be no problem)
	 */
	void Store(v8::Isolate* isolate, long ID, v8::Local<v8::Function> value)
	{
		std::hash_map < long, CopyablePersistent >::iterator iter = func_map.find(ID);
		if (iter == func_map.end())
		{
			CopyablePersistent persistent(isolate, value);
			func_map.insert(std::pair< long, CopyablePersistent >(ID, persistent));
		}
		else
		{
			iter->second.Reset(isolate, value); // neat
		}
	}

	/**
	 * set pointer to CV class for callbacks
	 */
	STDMETHOD(SetPtr)(CScripto *ptr){ this->ptr = ptr; return S_OK; }

	/**
	 * interceptor for calls from the source interface.  we have dispids mapped to functions.
	 */
	STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr);
	

};

OBJECT_ENTRY_AUTO(__uuidof(CVEventHandler), CCVEventHandler)
