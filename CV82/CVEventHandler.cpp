// CVEventHandler.cpp : Implementation of CCVEventHandler

#include "stdafx.h"
#include "CVEventHandler.h"
#include "Scripto.h"

STDMETHODIMP CCVEventHandler::Invoke(DISPID dispidMember, REFIID riid, LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult, EXCEPINFO* pexcepinfo, UINT* puArgErr)
{
	std::hash_map < long, CopyablePersistent >::iterator iter = func_map.find(dispidMember);
	if (iter != func_map.end())
	{
		return ptr->EventCallback( iter->second );
	}
	return IDispatchImpl<ICVEventHandler, &IID_ICVEventHandler, &LIBID_CV82Lib, /*wMajor =*/ 1, /*wMinor =*/ 0>::
		Invoke(dispidMember, riid, lcid, wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr);
}

// CCVEventHandler

