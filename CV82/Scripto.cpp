// Scripto.cpp : Implementation of CScripto

#include "stdafx.h"
#include "Scripto.h"
#include "CVEventHandler.h"

#include <comdef.h>

#include <string>
#include <sstream>

// CScripto

//-----------------------------------------------------------
// functions from the v8 samples (shell.cc)
//-----------------------------------------------------------

const char* ToCString(const v8::String::Utf8Value& value) {
	return *value ? *value : "<string conversion failed>";
}

void FormatException(v8::Isolate* isolate, v8::TryCatch* try_catch, std::string &strreport) {

	v8::HandleScope handle_scope(isolate);
	v8::String::Utf8Value exception(try_catch->Exception());
	const char* exception_string = ToCString(exception);
	v8::Handle<v8::Message> message = try_catch->Message();
	if (message.IsEmpty()) {
		strreport = exception_string;
		strreport.append("\r\n");
	}
	else {
		v8::String::Utf8Value filename(message->GetScriptResourceName());
		const char* filename_string = ToCString(filename);
		int linenum = message->GetLineNumber();

		char sz[128];
		sprintf_s(sz, 128, "%s:%i: ", filename_string, linenum);
		strreport = sz;
		strreport.append(exception_string);
		strreport.append("\r\n");

		// Print line of source code.
		v8::String::Utf8Value sourceline(message->GetSourceLine());
		const char* sourceline_string = ToCString(sourceline);

		strreport.append(sourceline_string);
		strreport.append("\r\n");

		// Print wavy underline (GetUnderline is deprecated).
		int start = message->GetStartColumn();
		for (int i = 0; i < start; i++) {
			strreport.append(" ");
		}
		int end = message->GetEndColumn();
		for (int i = start; i < end; i++) {
			strreport.append("^");
		}
	}
}

//-----------------------------------------------------------
// utility functions FIXME: class.  also, improve them.
//-----------------------------------------------------------

void NarrowString(BSTR* bstr, std::string &str)
{
	CComBSTR wide(*bstr);
	int len = wide.Length();
	for (int i = 0; i < len; i++) str += (char)(wide.m_str[i] & 0xff);
}

void NarrowString2(LPCTSTR tstr, std::string &str)
{
	CComBSTR wide(tstr);
	int len = wide.Length();
	for (int i = 0; i < len; i++) str += (char)(wide.m_str[i] & 0xff);
}

void FormatCOMError(std::string &target, HRESULT hr, const char *msg, const char *symbol = 0)
{
	char szException[1024];

	_com_error err(hr);
	LPCTSTR errMsg = err.ErrorMessage();

	std::string str;

	NarrowString2(errMsg, str);

	target = msg ? msg : "COM Error";
	if (symbol)
	{
		target += " (";
		target += symbol;
		target += "): ";
	}

	sprintf_s(szException, 1024, "0x%x: ", hr);

	target += szException;
	target += str.c_str();

}

//-----------------------------------------------------------
//
//-----------------------------------------------------------


const char * vt_type(int vt)
{
	vt &= (~VT_BYREF);
	switch (vt){
	case 1:
		return "VT_NULL";
	case 2:
	case 3:
		return "int";
	case 4:
	case 5:
		return "double";
	case 6:
		return 		"VT_CY";// 6,
	case 7:
		return 		"VT_DATE";// 7,
	case 8:
		return 		"string";
	case 9:
		return 		"IDispatch";// 9,
	case 10:
		return 		"VT_ERROR";// 10,
	case 11:
		return 		"boolean";// 11,
	case 12:
		return 		"variant";// 12,
	case 13:
		return 		"IUnknown";// 13,
	case 14:
		return 		"VT_DECIMAL";// 14,
	case 16:
		return 		"VT_I1";// 16,
	case 17:
		return 		"VT_UI1";// 17,
	case 18:
		return 		"VT_UI2";// 18,
	case 19:
		return 		"uint";// 19,
	case 20:
		return 		"int";// 20,
	case 21:
		return 		"uint";// 21,
	case 22:
		return 		"int";// 22,
	case 23:
		return 		"uint";// 23,
	case 24:
		return 		"void";// 24,
	case 25:
		return 		"VT_HRESULT";// 25,
	case 26:
		return 		"VT_PTR";// 26,
	case 27:
		return 		"VT_SAFEARRAY";// 27,
	}
	return "OTHER";
}

void map_returntype(std::string &rtype, TYPEDESC *ptdesc, CComPtr<ITypeInfo> typeinfo)
{
	//HRESULT hr;

	CComPtr< ITypeInfo > spTypeInfo2;
	CComBSTR bstrName;

	if (ptdesc->vt == VT_PTR)
	{
		while ((ptdesc->vt == VT_PTR) && (ptdesc->lptdesc != 0)){
			ptdesc = ptdesc->lptdesc;
		}
	}

	if (ptdesc->hreftype != 0)
	{
		CComPtr<ITypeInfo> spTypeInfo2;
		if (ptdesc->vt == VT_USERDEFINED &&
			(SUCCEEDED(typeinfo->GetRefTypeInfo(ptdesc->hreftype, &spTypeInfo2))))
		{
			if (SUCCEEDED(spTypeInfo2->GetDocumentation(-1, &bstrName, 0, 0, 0)))
			{
				std::string refname;
				rtype.clear();
				NarrowString(&bstrName, rtype);
			}
			else
			{
				rtype = "(not found [1])";
			}
		}
		else
		{
			rtype = "(not found [2])"; // probably need to read the import list
		}
	}
	else if (ptdesc->vt)
	{
		rtype = vt_type(ptdesc->vt);
	}
	else
	{
		rtype = "(not found [3])"; // case?
	}

}


void mapCoClass(std::string &output, std::string &name, CComPtr<ITypeInfo> typeinfo, TYPEATTR *pTatt)
{
	std::stringstream ss;

	CComBSTR bstrName;
	std::string elementname;
	INT lImplTypeFlags;
	HREFTYPE handle;
	CComPtr< ITypeInfo > spTypeInfo2;

	ss << "\"" << name.c_str() << "\": { \n";
	ss << "\t\"type\": \"coclass\", \n";

	// FIXME: could cache source GUIDs here, would save lookup later

	for (UINT j = 0; j < pTatt->cImplTypes; j++)
	{
		HRESULT hr = typeinfo->GetImplTypeFlags(j, &lImplTypeFlags);
		if (SUCCEEDED(hr)) hr = typeinfo->GetRefTypeOfImplType(j, &handle);
		if (SUCCEEDED(hr)) hr = typeinfo->GetRefTypeInfo(handle, &spTypeInfo2);
		if (SUCCEEDED(hr))
		{
			if (SUCCEEDED(spTypeInfo2->GetDocumentation(-1, &bstrName, 0, 0, 0)))
			{
				elementname.clear();
				NarrowString(&bstrName, elementname);
				if (lImplTypeFlags == 1) ss << "\t\"default\": \"" << elementname << "\"";
				else ss << "\t\"source\": \"" << elementname << "\"";
				if (j < pTatt->cImplTypes - 1) ss << ",";
				ss << "\n";
			}
			spTypeInfo2 = 0;
		}
	}

	ss << "}";
	output = ss.str();

}

void mapEnum(std::string &output, std::string &name, CComPtr<ITypeInfo> typeinfo, TYPEATTR *pTatt, v8::Handle< v8::Context > context)
{
	std::stringstream ss;

	VARDESC *vd = 0;
	CComBSTR bstrName;
	std::string elementname;

	v8::Isolate * isolate = context->GetIsolate();
	v8::Local<v8::ObjectTemplate> tmpl = v8::ObjectTemplate::New(isolate);


	ss << "\"" << name.c_str() << "\": { \n";
	ss << "\t\"type\": \"enum\", \n";
	ss << "\t\"values\": { \n";

	for (UINT u = 0; u < pTatt->cVars; u++)
	{
		vd = 0;
		HRESULT hr = typeinfo->GetVarDesc(u, &vd);
		if (SUCCEEDED(hr))
		{
			if (vd->varkind != VAR_CONST
				|| vd->lpvarValue->vt != VT_I4)
			{
				ATLTRACE("enum type not const/I4\n");
			}

			elementname.clear();

			hr = typeinfo->GetDocumentation(vd->memid, &bstrName, 0, 0, 0);
			if (SUCCEEDED(hr))
			{
				NarrowString(&bstrName, elementname);
				ss << "\t\t\"" << elementname << "\": ";
				ss << vd->lpvarValue->intVal;
				if (u < pTatt->cVars - 1) ss << ",";
				ss << "\n";

				tmpl->Set(v8::String::NewFromUtf8(isolate, elementname.c_str()), v8::Integer::New(isolate, vd->lpvarValue->intVal));

			}
			typeinfo->ReleaseVarDesc(vd);
		}
	}

	ss << "\t}\n}";
	output = ss.str();

	v8::Local<v8::Object> inst = tmpl->NewInstance();
	context->Global()->Set(v8::String::NewFromUtf8(isolate, name.c_str()), inst);


}

void mapInterface(std::string &output, std::string &name, CComPtr<ITypeInfo> typeinfo, CComPtr<ITypeLib> typelib, TYPEATTR *pTatt, TYPEKIND tk)
{
	std::stringstream ss;
	std::string elementname;
	CComBSTR bstrName;

	FUNCDESC *fd;

	std::hash_map< std::string, MemberRep* > ifacemap;

	ss << "\"" << name.c_str() << "\": { \n";
	ss << "\t\"type\": \"" << (tk == TKIND_DISPATCH ? "dispatch" : "interface") << "\", \n";
	ss << "\t\"members\": {\n";

	for (UINT u = 0; u < pTatt->cFuncs; u++)
	{
		fd = 0;
		HRESULT hr = typeinfo->GetFuncDesc(u, &fd);
		if (SUCCEEDED(hr))
		{
			elementname.clear();

			hr = typeinfo->GetDocumentation(fd->memid, &bstrName, 0, 0, 0);
			if (SUCCEEDED(hr))
			{
				NarrowString(&bstrName, elementname);

				std::string rtype = "(unknown type)";

				if (fd->funckind == FUNC_DISPATCH) // && ( fd->invkind == INVOKE_PROPERTYGET || fd->invkind == INVOKE_FUNC ))
				{
					map_returntype(rtype, &(fd->elemdescFunc.tdesc), typeinfo);
				}
				else if (fd->lprgelemdescParam && (fd->lprgelemdescParam->paramdesc.wParamFlags & 0xa))
				{
					map_returntype(rtype, &(fd->lprgelemdescParam->tdesc), typeinfo);
				}
				else if ((fd->invkind == INVOKE_FUNC) && fd->lprgelemdescParam)
				{
					map_returntype(rtype, &(fd->lprgelemdescParam->tdesc), typeinfo);
				}
				else if (fd->invkind == INVOKE_FUNC) // && fd->lprgelemdescParam)
				{
					rtype = "void";
				}

				std::hash_map< std::string, MemberRep* >::iterator iter = ifacemap.find(elementname.c_str());
				if (iter != ifacemap.end())
				{
					MemberRep *mr = iter->second;

					// if this is get, set type
					if (fd->invkind == INVOKE_PROPERTYGET)
					{
						mr->type = rtype.c_str();
						mr->mrflags |= MRFLAG_PROPGET;
					}
					else if (fd->invkind == INVOKE_PROPERTYPUT || fd->invkind == INVOKE_PROPERTYPUTREF)
					{
						mr->mrflags |= MRFLAG_PROPPUT;
					}
					else
					{
						// this should not happen
						ATLTRACE("unexpected case mrflags\n");
					}

					ifacemap.erase(iter);
					ifacemap.insert(std::pair< std::string, MemberRep* >(elementname, mr));
				}
				else
				{
					MemberRep *mr = new MemberRep();
					mr->name = elementname.c_str();
					mr->type = rtype.c_str();
					if (fd->invkind == INVOKE_FUNC) mr->mrflags = MRFLAG_METHOD;
					else if (fd->invkind == INVOKE_PROPERTYGET) mr->mrflags = MRFLAG_PROPGET;
					else mr->mrflags = MRFLAG_PROPPUT;
					ifacemap.insert(std::pair< std::string, MemberRep* >(elementname, mr));
				}
			}
			typeinfo->ReleaseFuncDesc(fd);
		}
	}

	{
		std::hash_map< std::string, MemberRep* >::iterator iter = ifacemap.begin();
		for (int i = 0; i< ifacemap.size(); i++)
		{
			MemberRep *mr = (iter++)->second;
			ss << "\t\t\"" << mr->name << "\": { type:\"" << mr->type << "\", flags: " << mr->mrflags << " }";
			if (i < ifacemap.size() - 1) ss << ",";
			ss << "\n";
			delete mr;
		}
	}

	ss << "\t}\n}";

	output = ss.str();
}


//-----------------------------------------------------------
// callback functions
//-----------------------------------------------------------

void indexed_getter(uint32_t index, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = info.GetIsolate();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->Indexed_Getter(index, info);
}

void indexed_setter(uint32_t index, v8::Local<v8::Value> value, const v8::PropertyCallbackInfo<v8::Value>& info)
{

}

void getter(v8::Local<v8::String> property, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = info.GetIsolate();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->Getter(property, info);
}

void setter(v8::Local<v8::String> property, v8::Local<v8::Value> value, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = info.GetIsolate();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->Setter(property, value, info);
}

void invoker(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	v8::Isolate* isolate = args.GetIsolate();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(args.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->Invoker(args);

}

void ReleasePtr(const v8::WeakCallbackData<v8::Object, CScripto> &data )
{
	
	v8::Isolate* isolate = data.GetIsolate();
	v8::HandleScope handle_scope(isolate);

	CScripto *scripto = data.GetParameter();

	//if (data.GetParameter()) data.GetParameter()->Release();

	v8::Handle<v8::External> check = v8::Handle<v8::External>::Cast(data.GetValue()->GetInternalField(1));
	IDispatch *pdisp = (IDispatch*)(check->Value());
	if (pdisp)
	{
		int x = pdisp->Release();
		ATLTRACE("pdisp->Release returned %d\n", x);
	}
	scripto->RemovePersistentObj(pdisp);

	ATLTRACE("Release ptr on weak callback: 0x%x\n", pdisp);

	data.GetValue()->SetInternalField(1, v8::External::New(isolate, 0));

	CComObject<CCVEventHandler>* pSink;
	check = v8::Handle<v8::External>::Cast(data.GetValue()->GetInternalField(2));
	pSink = ((CComObject<CCVEventHandler>*)(check->Value()));
	if (pSink) pSink->Release();
	data.GetValue()->SetInternalField(2, v8::External::New(isolate, 0));



}

PWrap* CScripto::MapPersistentObj(v8::Isolate *isolate, IDispatch *pdisp)
{
	std::hash_map < long, PWrap* > ::iterator iter = object_map.find((long)pdisp);

	if (iter != object_map.end())
	{
		return iter->second;
	}

	// wrap dispatch will addref.  we want to release it at some point.
	// for the non-global objects, that should be up to the gc.

	// FIXME: we need to force the GC to be more aggressive.  I think there's
	// a build flag we can use.

	PWrap *wrap = new PWrap;

	{
		v8::HandleScope scope(isolate);
		v8::Local< v8::Object > local = WrapDispatch(isolate, pdisp);

		UINT ui;
		CComPtr< ITypeInfo > typeinfo;
		CComBSTR bstr;

		HRESULT hr = pdisp->GetTypeInfoCount(&ui);
		if (SUCCEEDED(hr) && ui > 0)
		{
			hr = pdisp->GetTypeInfo(0, 1033, &typeinfo);
		}
		if (SUCCEEDED(hr))
		{
			hr = typeinfo->GetDocumentation(-1, &bstr, 0, 0, 0);
		}
		if (SUCCEEDED(hr))
		{
			std::string narrow;
			NarrowString(&bstr, narrow);
			// local->SetHiddenValue(v8::String::NewFromUtf8(isolate, "internaltype"), v8::String::NewFromUtf8(isolate, narrow.c_str()));
			local->Set(v8::String::NewFromUtf8(isolate, "__internaltype"), v8::String::NewFromUtf8(isolate, narrow.c_str()));
		}

		wrap->p.Reset(isolate, local );
		wrap->p.SetWeak<CScripto>(this, ReleasePtr);
	}

	ATLTRACE("Set weak: 0x%x\n", (unsigned long)pdisp);

	object_map.insert(std::pair< long, PWrap* >((long)pdisp, wrap));

	return wrap;
}

void CScripto::RemovePersistentObj(IDispatch *pdisp)
{
	std::hash_map < long, PWrap* > ::iterator iter = object_map.find((long)pdisp);
	if (iter != object_map.end())
	{
		PWrap *p = iter->second;
		p->p.Reset();
		object_map.erase(iter);
		delete p;
	}
	
}

void CScripto::SetRetVal(v8::Isolate* isolate, CComVariant &var, v8::ReturnValue< v8::Value > &retval)
{
	switch (var.vt)
	{
	case VT_DISPATCH: // remember: can this actually be a persistent?
		retval.Set((MapPersistentObj(isolate, var.pdispVal))->p);
		break;
	case VT_EMPTY:
		retval.Set(v8::Undefined(isolate));
		break;
	case VT_NULL:
		retval.Set(v8::Null(isolate));
		break;
	case VT_I4:
		retval.Set(v8::Integer::New(isolate, var.intVal));
		break;
	case VT_R4:
	case VT_R8:
		retval.Set(v8::Number::New(isolate, var.dblVal));
		break;
	case VT_BSTR:
		{
			std::string strResult;
			NarrowString(&var.bstrVal, strResult);
			retval.Set(v8::String::NewFromUtf8(isolate, strResult.c_str()));
		}
		break;
	case VT_BOOL:
		retval.Set(var.boolVal);
		break;
	default:
		{
			std::string strMessage = "unexpected variant type: ";
			char sz[32];
			sprintf_s(sz, 32, "%d", var.vt);
			strMessage.append(sz);
			isolate->ThrowException(v8::String::NewFromUtf8(isolate, strMessage.c_str()));
		}
	}
}

void CScripto::Invoker(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	v8::Isolate* isolate = getInstanceIsolate();
	v8::HandleScope handle_scope(isolate);

	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(args.Holder()->GetInternalField(1));
	IDispatch *pdisp = (IDispatch*)(external->Value());
	v8::Handle<v8::Value> val1 = args.Callee()->GetHiddenValue(v8::String::NewFromUtf8(isolate, "dispid"));
	v8::Handle<v8::Value> val2 = args.Callee()->GetHiddenValue(v8::String::NewFromUtf8(isolate, "propget"));

	CComVariant cvResult;

	long dispid = val1->Int32Value();
	bool propget = val2->BooleanValue();

	DISPPARAMS dispparams;
	dispparams.cArgs = 0;
	dispparams.cNamedArgs = 0;

	HRESULT hr;

	
	dispparams.cArgs = args.Length();
	CComVariant *pcv = 0;
	CComBSTR *pbstr = 0;

	if (dispparams.cArgs > 0)
	{
		pcv = new CComVariant[dispparams.cArgs];
		pbstr = new CComBSTR[dispparams.cArgs];

		dispparams.rgvarg = pcv;
		int arglen = args.Length();

		for (int i = 0; i < arglen; i++)
		{
			int k = arglen - 1 - i;

			if (args[i]->IsBoolean()) pcv[k] = args[i]->BooleanValue();
			else if (args[i]->IsInt32()) pcv[k] = args[i]->Int32Value();
			else if (args[i]->IsNumber()) pcv[k] = args[i]->NumberValue();
			else if (args[i]->IsString())
			{
				v8::String::Utf8Value str(args[i]);
				const char* cstr = ToCString(str);
				pbstr[k] = cstr;
				pcv[k].vt = VT_BSTR | VT_BYREF;
				pcv[k].pbstrVal = &(pbstr[k]);
			}
			else if (args[i]->IsObject())
			{
				pcv[k].vt = VT_DISPATCH;

				external = v8::Handle<v8::External>::Cast(args[i]->ToObject()->GetInternalField(1));
				pcv[k].pdispVal = (IDispatch*)(external->Value());
			}
		}
	}

	hr = pdisp->Invoke(dispid, IID_NULL, 1033, propget ? DISPATCH_PROPERTYGET : DISPATCH_METHOD, &dispparams, &cvResult, NULL, NULL);

	if (pcv) delete[] pcv;
	if (pbstr) delete[] pbstr;


	if (FAILED(hr))
	{
		std::string msg;
		FormatCOMError(msg, hr, "COM Exception in Invoke");
		isolate->ThrowException(v8::String::NewFromUtf8(isolate, msg.c_str()));
	}
	else
	{
		SetRetVal(isolate, cvResult, args.GetReturnValue());
	}

}

HRESULT CScripto::EventCallback(const v8::Persistent<v8::Function, v8::CopyablePersistentTraits<v8::Function>> func)
{

	v8::Isolate* isolate = getInstanceIsolate();
	isolate->Enter();
	v8::Locker locker(isolate);

	v8::HandleScope handle_scope(isolate);

	v8::Local<v8::Context> context = v8::Local<v8::Context>::New(isolate, _context);

	context->Enter();

	v8::Context::Scope context_scope(context);

	v8::Handle< v8::Function > local = v8::Local<v8::Function>::New(isolate, func);
	v8::Handle< v8::Value > result = local->Call(context->Global(), 0, NULL);

	context->Exit();
	isolate->Exit();

	return S_OK;
}

HRESULT CScripto::GetCoClassForDispatch(ITypeInfo **ppCoClass, IDispatch *pdisp)
{
	CComPtr<ITypeInfo> spTypeInfo;
	CComPtr<ITypeLib> spTypeLib;

	bool matchIface = false;
	HRESULT hr = pdisp->GetTypeInfo(0, 0, &spTypeInfo);

	if (SUCCEEDED(hr) && spTypeInfo)
	{
		UINT tlidx;
		hr = spTypeInfo->GetContainingTypeLib(&spTypeLib, &tlidx);
	}

	if (SUCCEEDED(hr))
	{
		UINT tlcount = spTypeLib->GetTypeInfoCount();

		for (UINT u = 0; !matchIface && u < tlcount; u++)
		{
			TYPEATTR *pTatt = nullptr;
			CComPtr<ITypeInfo> spTypeInfo2;
			TYPEKIND ptk;

			if (SUCCEEDED(spTypeLib->GetTypeInfoType(u, &ptk)) && ptk == TKIND_COCLASS)
			{
				hr = spTypeLib->GetTypeInfo(u, &spTypeInfo2);
				if (SUCCEEDED(hr))
				{
					hr = spTypeInfo2->GetTypeAttr(&pTatt);
				}
				if (SUCCEEDED(hr))
				{
					for (UINT j = 0; !matchIface && j < pTatt->cImplTypes; j++)
					{
						INT lImplTypeFlags;
						hr = spTypeInfo2->GetImplTypeFlags(j, &lImplTypeFlags);
						if (SUCCEEDED(hr) && lImplTypeFlags == 1) // default interface, disp or dual
						{
							HREFTYPE handle;
							if (SUCCEEDED(spTypeInfo2->GetRefTypeOfImplType(j, &handle)))
							{
								CComPtr< ITypeInfo > spTypeInfo3;
								if (SUCCEEDED(spTypeInfo2->GetRefTypeInfo(handle, &spTypeInfo3)))
								{
									CComBSTR bstr;
									TYPEATTR *pTatt2 = nullptr;
									CComPtr<IUnknown> punk = 0;

									hr = spTypeInfo3->GetTypeAttr(&pTatt2);
									if (SUCCEEDED(hr)) hr = pdisp->QueryInterface(pTatt2->guid, (void**)&punk);
									if (SUCCEEDED(hr))
									{
										*ppCoClass = spTypeInfo2;
										// ... (*ppCoClass)->AddRef();
										matchIface = true;
									}

									if (pTatt2) spTypeInfo3->ReleaseTypeAttr(pTatt2);
								}
							}
						}
					}
				}
				if (pTatt) spTypeInfo2->ReleaseTypeAttr(pTatt);
			}
		}
	}

	return matchIface ? S_OK : E_FAIL;
}

void CScripto::Unsink(IDispatch *pdisp)
{
	long ID = (long)pdisp;
	std::hash_map < long, void* > ::iterator iter = sink_map.find(ID);
	CComObject<CCVEventHandler>* pSink = (CComObject<CCVEventHandler>*)(iter->second);
	if (pSink)
	{
		pSink->Reset();
		CComPtr<IUnknown> punk = 0;
		pSink->QueryInterface(IID_IUnknown, (void**)&punk);
		AtlUnadvise(punk, pSink->iid, pSink->cookie);
		pSink->Release();
	}
	sink_map.erase(iter);
}

void CScripto::Setter(v8::Local<v8::String> property, v8::Local<v8::Value> value, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = getInstanceIsolate();
	v8::HandleScope handle_scope(isolate);

	v8::Handle<v8::External> disp = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(1));
	CComPtr<IDispatch> pdisp((IDispatch*)(disp->Value()));

	std::string propname = ToCString(v8::String::Utf8Value(property));
	std::string valstr = ToCString(v8::String::Utf8Value(value));

	// ATLTRACE("setter: %s => %s\n", propname.c_str(), valstr.c_str());

	if (!pdisp) return;

	CComBSTR wide;
	wide.Append(propname.c_str());

	WCHAR *member = (LPWSTR)wide;
	DISPID dispid;

	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &member, 1, 1033, &dispid);

	if (FAILED(hr))
	{
		// it's not in the default interface.  it might be 
		// an event, if the script is trying to set a callback.
		// in that case, we need to find the appropriate
		// source interface.  we should map these ahead of 
		// time or at least cache mappings.

		/* if it's not a function, treat it as nullity - essentially clearing any existing mappings */

		/* if (!value->IsFunction()) return; */

		CComPtr<ITypeInfo> spTypeInfo;
		hr = pdisp->GetTypeInfo(0, 0, &spTypeInfo);

		if (SUCCEEDED(hr) && spTypeInfo)
		{
			CComPtr<ITypeLib> spTypeLib;
			UINT tlidx;
			hr = spTypeInfo->GetContainingTypeLib(&spTypeLib, &tlidx);
			if (SUCCEEDED(hr))
			{
				UINT tlcount = spTypeLib->GetTypeInfoCount();
				for (UINT u = 0; u < tlcount; u++)
				{
					TYPEKIND ptk;
					if (SUCCEEDED(spTypeLib->GetTypeInfoType(u, &ptk)) && ptk == TKIND_COCLASS)
					{
						CComPtr<ITypeInfo> spTypeInfo2;
						if (SUCCEEDED(spTypeLib->GetTypeInfo(u, &spTypeInfo2)))
						{
							TYPEATTR *pTatt = nullptr;
							if (SUCCEEDED(spTypeInfo2->GetTypeAttr(&pTatt)))
							{
								bool matchIface = false;

								for (UINT j = 0; j < pTatt->cImplTypes; j++)
								{
									INT lImplTypeFlags;
									if (SUCCEEDED(spTypeInfo2->GetImplTypeFlags(j, &lImplTypeFlags)))
									{
										if (lImplTypeFlags == 1) // default interface, disp or dual
										{
											HREFTYPE handle;
											if (SUCCEEDED(spTypeInfo2->GetRefTypeOfImplType(j, &handle)))
											{
												CComPtr< ITypeInfo > spTypeInfo3;
												if (SUCCEEDED(spTypeInfo2->GetRefTypeInfo(handle, &spTypeInfo3)))
												{
													CComBSTR bstr;
													TYPEATTR *pTatt2 = nullptr;

													hr = spTypeInfo3->GetDocumentation(-1, &bstr, 0, 0, 0);
													if( SUCCEEDED( hr )) hr = spTypeInfo3->GetTypeAttr(&pTatt2);
													if( SUCCEEDED(hr))
													{
														std::string iface_string;
														NarrowString( &bstr, iface_string );
														ATLTRACE("%s\n", iface_string.c_str());

														//ATLTRACE("Nested deep");
														CComPtr<IUnknown> punk = 0;

														if( SUCCEEDED( pdisp->QueryInterface(pTatt2->guid, (void**)&punk)))
														{
															matchIface = true;
														}

													}

													spTypeInfo3->ReleaseTypeAttr(pTatt2);

												}
											}

											// pdisp->QueryInterface()
										}
									}
								}

								if (matchIface)
								{
									for (UINT j = 0; j < pTatt->cImplTypes; j++)
									{
										INT lImplTypeFlags;
										if (SUCCEEDED(spTypeInfo2->GetImplTypeFlags(j, &lImplTypeFlags)))
										{
											if (lImplTypeFlags & 0x02) // source interface 
											{
												{
													// ok, it's a good method.

													HREFTYPE handle;
													if (SUCCEEDED(spTypeInfo2->GetRefTypeOfImplType(j, &handle)))
													{
														CComPtr< ITypeInfo > spTypeInfo3;
														if (SUCCEEDED(spTypeInfo2->GetRefTypeInfo(handle, &spTypeInfo3)))
														{
															CComBSTR bstr;
															TYPEATTR *pTatt2 = nullptr;
															CComObject<CCVEventHandler>* pSink = 0;
															MEMBERID memID2;
															hr = spTypeInfo3->GetIDsOfNames(&member, 1, &memID2);

															ATLTRACE("OBJECT DISP: 0x%x\r\n", pdisp);

															// we might have already sunk it

															// NOTE: this method is no good.  the dispatch is getting re-wrapped on
															// subsequent calls (in most cases).  so there's no use in stuffing the 
															// sink object into this local v8 object.

															// we need a global map.  NOTE: the *objects* don't have to be persistent,
															// to support COM calls, although we will have to release the sinks at 
															// some point.  What about the *functions*?

															// NOTE: functions are persisted in the event handler obj, (@see)

															//v8::Handle<v8::External> check = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(2));
															//pSink = ((CComObject<CCVEventHandler>*)(check->Value()));

															// FIXME: need to clean up this map at some point

															std::hash_map < long, void* > ::iterator iter = sink_map.find((long)pdisp.p);
															if (iter != sink_map.end()) pSink = (CComObject<CCVEventHandler>*)(iter->second);

															if (!pSink && value->IsFunction()) // ok sink
															{

																if (SUCCEEDED(hr)) hr = spTypeInfo3->GetDocumentation(-1, &bstr, 0, 0, 0);
																if (SUCCEEDED(hr)) hr = spTypeInfo3->GetTypeAttr(&pTatt2);
																if (SUCCEEDED(hr))
																{
																	// hr = pSinkClass->QueryInterface(IID_IUnknown, (LPVOID*)&m_pSinkUnk);


																	CComObject<CCVEventHandler>::CreateInstance(&pSink);
																	CComPtr<IUnknown> punk;
																	pSink->AddRef();
																	pSink->QueryInterface(IID_IUnknown, (void**)&punk);
																	pSink->SetPtr(this);
																	pSink->iid = pTatt2->guid;

																	hr = AtlAdvise(pdisp, punk, pTatt2->guid, &(pSink->cookie));

																	//info.Holder()->SetInternalField(2, v8::External::New(isolate, pSink));
																	//info.Holder()->SetInternalField(3, v8::Int32::New(isolate, dwCookie));

																	// FIXME: unadvise also

																	sink_map.insert(std::pair< long, void* >((long)pdisp.p, (void*)pSink));

																}
															}

															if (pSink && value->IsFunction())
															{
																// store the callback function in there
																pSink->Store(isolate, memID2, value.As<v8::Function>());
															}
															else if (pSink)
															{
																pSink->Clear(isolate, memID2);
																if (pSink->func_map.size() == 0) Unsink(pdisp);
															}

															spTypeInfo3->ReleaseTypeAttr(pTatt2);

														}
													}			//
												}				//
											}					// <--- bad code structure
										}						//
									}							//
								}								//

								spTypeInfo2->ReleaseTypeAttr(pTatt);
							}
						}
					}
				}
			}

			TYPEATTR *pTatt = nullptr;
			hr = spTypeInfo->GetTypeAttr(&pTatt);
			if (SUCCEEDED(hr) && pTatt)
			{

				spTypeInfo->ReleaseTypeAttr(pTatt);
			}
		}

		return;
	}

	if (FAILED(hr)) return;

	DISPPARAMS dispparams;
	DISPID dispidNamed = DISPID_PROPERTYPUT;

	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 1;
	dispparams.rgdispidNamedArgs = &dispidNamed;

	CComVariant cv;
	CComBSTR bstr;

	if (value->IsBoolean()) cv = value->BooleanValue();
	else if (value->IsInt32()) cv = value->Int32Value();
	else if (value->IsNumber()) cv = value->NumberValue();
	if (value->IsString())
	{
		v8::String::Utf8Value str(value);
		const char* cstr = ToCString(str);
		bstr = cstr;
		cv.vt = VT_BSTR | VT_BYREF;
		cv.pbstrVal = &(bstr);
	}
	else if (value->IsObject())
	{
		cv.vt = VT_DISPATCH;
		v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(value->ToObject()->GetInternalField(1));
		cv.pdispVal = (IDispatch*)(external->Value());
	}

	dispparams.rgvarg = &cv;

	hr = pdisp->Invoke(dispid, IID_NULL, 1033, DISPATCH_PROPERTYPUT, &dispparams, NULL, NULL, NULL);

	if (FAILED(hr))
	{
		std::string strMessage = "COM exception: propertyput: ";
		strMessage.append(propname);
		isolate->ThrowException(v8::String::NewFromUtf8(isolate, strMessage.c_str()));
	}
}

void CScripto::Indexed_Getter(uint32_t index, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = getInstanceIsolate();
	v8::HandleScope handle_scope(isolate);

	v8::Handle<v8::External> disp = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(1));
	CComPtr<IDispatch> pdisp((IDispatch*)(disp->Value()));

	if (!pdisp) return;

	CComBSTR wide;
	wide.Append("Item");

	WCHAR *member = (LPWSTR)wide;
	DISPID dispid;

	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &member, 1, 1033, &dispid);
	if (FAILED(hr)) return;

	DISPPARAMS dispparams;
	dispparams.cArgs = 1;
	dispparams.cNamedArgs = 0;

	CComVariant cv = index + 1; // breaking vb tradition // FIXME: make configurable?

	dispparams.rgvarg = &cv;

	CComVariant cvResult;

	hr = pdisp->Invoke(dispid, IID_NULL, 1033, DISPATCH_PROPERTYGET, &dispparams, &cvResult, NULL, NULL);

	if (SUCCEEDED(hr))
	{
		SetRetVal(isolate, cvResult, info.GetReturnValue());
	}
	else
	{
		std::string strMessage = "COM exception: propertyget: []";
		isolate->ThrowException(v8::String::NewFromUtf8(isolate, strMessage.c_str()));
	}

}

void CScripto::Getter(v8::Local<v8::String> property, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = getInstanceIsolate();
	v8::HandleScope handle_scope(isolate);

	v8::Handle<v8::External> disp = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(1));
	CComPtr<IDispatch> pdisp((IDispatch*)(disp->Value()));

	std::string propname = ToCString(v8::String::Utf8Value(property));
	//ATLTRACE("getter: %s\n", propname.c_str());

	if (!pdisp) return;

	CComBSTR wide;
	wide.Append(propname.c_str());

	WCHAR *member = (LPWSTR)wide;
	DISPID dispid;

	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &member, 1, 1033, &dispid);

	if (FAILED(hr)) return;

	CComPtr<ITypeInfo> spTypeInfo;
	hr = pdisp->GetTypeInfo(0, 0, &spTypeInfo);
	bool found = false;

	// FIXME: no need to iterate, you can use GetIDsOfNames from ITypeInfo 
	// to get the memid, then use that to get the funcdesc.  much more efficient.
	// ... how does that deal with propget/propput?

	if (SUCCEEDED(hr) && spTypeInfo)
	{
		TYPEATTR *pTatt = nullptr;
		hr = spTypeInfo->GetTypeAttr(&pTatt);
		if (SUCCEEDED(hr) && pTatt)
		{
			FUNCDESC * fd = nullptr;
			for (int i = 0; !found && i < pTatt->cFuncs; ++i)
			{
				hr = spTypeInfo->GetFuncDesc(i, &fd);
				if (SUCCEEDED(hr) && fd)
				{
					if (dispid == fd->memid 
						&& (fd->invkind == INVOKEKIND::INVOKE_FUNC
						|| fd->invkind == INVOKEKIND::INVOKE_PROPERTYGET))
					{
						UINT namecount;
						CComBSTR bstrNameList[32];
						spTypeInfo->GetNames(fd->memid, (BSTR*)&bstrNameList, 32, &namecount);

						found = true;
						if (fd->invkind == INVOKE_FUNC || ((fd->invkind == INVOKE_PROPERTYGET) && (fd->cParams - fd->cParamsOpt > 0)))
						{
							v8::Handle<v8::FunctionTemplate> tmpl = v8::FunctionTemplate::New(isolate, invoker);
							v8::Handle<v8::Function> func = tmpl->GetFunction();

								func->SetHiddenValue(v8::String::NewFromUtf8(isolate, "name"), v8::String::NewFromUtf8(isolate, propname.c_str()));
							func->SetHiddenValue(v8::String::NewFromUtf8(isolate, "dispid"), v8::Int32::New(isolate, dispid));
							func->SetHiddenValue(v8::String::NewFromUtf8(isolate, "propget"), v8::Boolean::New(isolate, fd->invkind == INVOKE_PROPERTYGET));

							info.GetReturnValue().Set(func);
						}
						else if (fd->invkind == INVOKE_PROPERTYGET)
						{

							DISPPARAMS dispparams;
							dispparams.cArgs = 0;
							dispparams.cNamedArgs = 0;

							CComVariant cvResult;

							HRESULT hr = pdisp->Invoke(dispid, IID_NULL, 1033, DISPATCH_PROPERTYGET, &dispparams, &cvResult, NULL, NULL);

							if (SUCCEEDED(hr))
							{
								SetRetVal(isolate, cvResult, info.GetReturnValue());
							}
							else
							{
								std::string strMessage = "COM exception: propertyget: ";
								strMessage.append(propname);
								isolate->ThrowException(v8::String::NewFromUtf8(isolate, strMessage.c_str()));
							}

						}
					}

					spTypeInfo->ReleaseFuncDesc(fd);
				}
			}

			spTypeInfo->ReleaseTypeAttr(pTatt);
		}
	}

	// info.GetReturnValue().Set(str);
}

void logmessage(const v8::FunctionCallbackInfo<v8::Value>& args) 
{
	v8::Isolate* isolate = args.GetIsolate();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(args.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->LogMessage(args);

}

void loginfo(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	v8::Isolate* isolate = args.GetIsolate();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(args.Holder()->GetInternalField(0));
	CScripto *p = (CScripto*)(external->Value());
	p->LogInfo(args);

}

void CScripto::LogInfo(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	CComBSTR bstr = "Invalid value";
	v8::Isolate* isolate = args.GetIsolate();

	if (args.Length() > 0)
	{
		if (!args[0]->IsObject())
		{
			LogMessage(args);
			return;
		}
		v8::Local< v8::Object > arg = v8::Handle<v8::Object>::Cast(args[0]);

		if (arg->InternalFieldCount() > 1)
		{
			v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(arg->GetInternalField(1));
			IDispatch *pdisp = (IDispatch*)(external->Value());
			if (pdisp)
			{
				// if (FAILED(MapTypeLib(pdisp, &bstr))) bstr = "COM error";

				UINT ui;
				CComPtr< ITypeInfo > typeinfo;
				CComPtr< ITypeLib > typelib;
				CComBSTR bstrName;
				MEMBERID memid;
				TYPEKIND tk;
				TYPEATTR *pTatt = nullptr;
				std::stringstream ss;

				HRESULT hr = pdisp->GetTypeInfoCount(&ui);
				if (SUCCEEDED(hr) && ui > 0)
				{
					hr = pdisp->GetTypeInfo(0, 1033, &typeinfo);
				}
				if (SUCCEEDED(hr))
				{
					hr = typeinfo->GetDocumentation(-1, &bstrName, 0, 0, 0);
				}
				if (SUCCEEDED(hr))
				{
					bstr = "object ";
					bstr.Append(bstrName);
					bstr.Append("\r\n");
				}
				if (SUCCEEDED(hr))
				{
					hr = typeinfo->GetContainingTypeLib(&typelib, &ui);
				}
				if (SUCCEEDED(hr)) hr = typeinfo->GetTypeAttr(&pTatt);

				CComPtr<ITypeInfo> ticc;
				GetCoClassForDispatch(&ticc, pdisp);
				if (ticc)
				{
					std::string strname;
					NarrowString(&bstrName, strname);
					std::string scc;
					TYPEATTR *pta = 0;
					ticc->GetTypeAttr(&pta);
					mapCoClass(scc, strname, ticc, pta);
					bstr += scc.c_str();

					////
					HREFTYPE handle;
					CComPtr< ITypeInfo > spTypeInfo2;
					for (UINT j = 0; j < pta->cImplTypes; j++)
					{
						TYPEATTR *pta2 = 0;

						hr = ticc->GetRefTypeOfImplType(j, &handle);
						if (SUCCEEDED(hr)) hr = ticc->GetRefTypeInfo(handle, &spTypeInfo2);
						if (SUCCEEDED(hr)) hr = spTypeInfo2->GetTypeAttr(&pta2);
						if (SUCCEEDED(hr))
						{
							CComBSTR bstrName2;
							if (SUCCEEDED(spTypeInfo2->GetDocumentation(-1, &bstrName2, 0, 0, 0)))
							{
								strname.clear();
								NarrowString(&bstrName2, strname);
								scc.clear();
								mapInterface(scc, strname, spTypeInfo2, typelib, pta2, pta2->typekind);
								bstr += scc.c_str();
								bstr += "\r\n";
							}
							spTypeInfo2 = 0;
						}
						if (pta2 && spTypeInfo2 ) spTypeInfo2->ReleaseTypeAttr(pta2);
					}

					ticc->ReleaseTypeAttr(pta);

				}
				else
				{
					std::string stri;
					std::string strname;
					NarrowString(&bstrName, strname);
					mapInterface(stri, strname, typeinfo, typelib, pTatt, pTatt->typekind);
					bstr += stri.c_str();
				}

				if (pTatt) typeinfo->ReleaseTypeAttr(pTatt);

				if (FAILED(hr)) bstr = "COM error";
				
			}
			else
			{
				bstr = "Not a dispatch interface";
			}
			Fire_OnConsolePrint(&bstr, 1);
		}
		else
		{
			if (!args[0]->IsObject())
			{
				LogMessage(args);
				return;
			}
		}
	}
	else Fire_OnConsolePrint(&bstr, 1);
}

void CScripto::LogMessage(const v8::FunctionCallbackInfo<v8::Value>& args) 
{
	// this is console.log

	CComBSTR bstr;

	for (int i = 0; i < args.Length(); i++)
	{
		if (i) bstr += " ";
		if (args[i]->IsUndefined()) bstr += "(undefined)";
		else if (args[i]->IsNull()) bstr += "(null)";
		else
		{
			v8::String::Utf8Value str(args[i]);
			const char* cstr = ToCString(str);
			bstr += cstr;
		}
	}
	
	Fire_OnConsolePrint(&bstr, 0);

}


void alert(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	v8::Isolate* isolate = args.GetIsolate();
	v8::HandleScope handle_scope(isolate);
	v8::Local< v8::External > external = v8::Handle<v8::External>::Cast(isolate->GetCallingContext()->Global()->GetHiddenValue(v8::String::NewFromUtf8(isolate, "inst")));
	CScripto *p = (CScripto*)(external->Value());
	p->Alert(args);
}

void confirm(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	v8::Isolate* isolate = args.GetIsolate();
	v8::HandleScope handle_scope(isolate);
	v8::Local< v8::External > external = v8::Handle<v8::External>::Cast(isolate->GetCallingContext()->Global()->GetHiddenValue(v8::String::NewFromUtf8(isolate, "inst")));
	CScripto *p = (CScripto*)(external->Value());
	p->Confirm(args);
}

void CScripto::Alert(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	CComBSTR bstr;

	for (int i = 0; i < args.Length(); i++)
	{
		if (i) bstr += " ";
		if (args[i]->IsUndefined()) bstr += "(undefined)";
		else if (args[i]->IsNull()) bstr += "(null)";
		else
		{
			v8::String::Utf8Value str(args[i]);
			const char* cstr = ToCString(str);
			bstr += cstr;
		}
	}

	Fire_OnAlert(&bstr);
}

void CScripto::Confirm(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	v8::Isolate* isolate = getInstanceIsolate();
	v8::HandleScope handle_scope(isolate);

	CComBSTR bstr;
	VARIANT_BOOL rslt;

	for (int i = 0; i < args.Length(); i++)
	{
		if (i) bstr += " ";
		if (args[i]->IsUndefined()) bstr += "(undefined)";
		else if (args[i]->IsNull()) bstr += "(null)";
		else
		{
			v8::String::Utf8Value str(args[i]);
			const char* cstr = ToCString(str);
			bstr += cstr;
		}
	}

	if (FAILED(Fire_OnConfirm(&bstr, &rslt))) rslt = VARIANT_FALSE;

	args.GetReturnValue().Set(v8::Boolean::New(isolate, rslt == VARIANT_TRUE ));

}


void CScripto::InitContext()
{
	v8::Isolate* isolate = getInstanceIsolate();

	v8::HandleScope handle_scope(isolate);

	v8::Handle< v8::ObjectTemplate > wrapper = v8::ObjectTemplate::New(isolate);
	wrapper->SetInternalFieldCount(4);
	wrapper->SetNamedPropertyHandler(getter, setter);
	wrapper->SetIndexedPropertyHandler(indexed_getter, indexed_setter);
	_wrapper.Reset(isolate, wrapper);

	v8::Local<v8::ObjectTemplate> global = v8::ObjectTemplate::New(isolate);

	global->Set(v8::String::NewFromUtf8(isolate, "alert"), v8::FunctionTemplate::New(isolate, alert));
	global->Set(v8::String::NewFromUtf8(isolate, "confirm"), v8::FunctionTemplate::New(isolate, confirm));

	v8::Handle<v8::Context> context = v8::Context::New(isolate, NULL, global);
	_context.Reset(isolate, context);

	context->Enter();

	context->Global()->SetHiddenValue(v8::String::NewFromUtf8(isolate, "inst"), v8::External::New(isolate, this));

	v8::Local<v8::ObjectTemplate> tmpl = v8::ObjectTemplate::New(isolate);
	tmpl->SetInternalFieldCount(1);
	tmpl->Set(v8::String::NewFromUtf8(isolate, "log"), v8::FunctionTemplate::New(isolate, logmessage));
	tmpl->Set(v8::String::NewFromUtf8(isolate, "info"), v8::FunctionTemplate::New(isolate, loginfo));

	v8::Local<v8::Object> console = tmpl->NewInstance();
	console->SetInternalField(0, v8::External::New(isolate, this));
	context->Global()->Set(v8::String::NewFromUtf8(isolate, "console"), console);

	_console.Reset(isolate, console);

	context->Exit();

}

void CScripto::ResetContext()
{
	_context.Reset();
	_wrapper.Reset();
	_console.Reset();
}

v8::Local< v8::Object > CScripto::WrapDispatch(v8::Isolate *isolate, IDispatch *pdisp)
{
	// v8::Isolate* isolate = getInstanceIsolate();
	// v8::HandleScope handle_scope(isolate);

	ATLTRACE("wrap dispatch: 0x%x\n", (unsigned long)pdisp);

	if (pdisp)
	{
		int x = 
			pdisp->AddRef();
		ATLTRACE("Addref returned %d\n", x);
	}

	v8::Local<v8::ObjectTemplate> wrapper = v8::Local<v8::ObjectTemplate>::New(isolate, _wrapper);
	v8::Local< v8::Object > instance = wrapper->NewInstance();

	instance->SetInternalField(0, v8::External::New(isolate, this));
	instance->SetInternalField(1, v8::External::New(isolate, pdisp));

	return instance;
}

//-----------------------------------------------------------
// exposed methods
//-----------------------------------------------------------

STDMETHODIMP CScripto::CleanUp()
{
	for (std::hash_map < long, void* > ::iterator iter = sink_map.begin();
		iter != sink_map.end();
		iter++)
	{
		CComObject<CCVEventHandler>* pSink = (CComObject<CCVEventHandler>*)(iter->second);
		if (pSink)
		{
			pSink->Reset();
			CComPtr<IUnknown> punk = 0;
			pSink->QueryInterface(IID_IUnknown, (void**)&punk);
			AtlUnadvise(punk, pSink->iid, pSink->cookie);
			pSink->Release();
		}
	}
	sink_map.clear();
	return S_OK;
}

STDMETHODIMP CScripto::SetGlobal(BSTR *JSON)
{
	return E_NOTIMPL;
}

STDMETHODIMP CScripto::GetGlobal(BSTR *JSON)
{
	// it's a start

	CComBSTR bstr = "JSON.stringify(this);";
	VARIANT_BOOL vb;

	return ExecString(&bstr, JSON, &vb);

}

STDMETHODIMP CScripto::MapTypeLib(IDispatch *Dispatch, BSTR *Description)
{
	
	CComPtr<ITypeInfo> spTypeInfo;
	CComPtr<ITypeLib> spTypeLib;
	
	HRESULT hr = Dispatch->GetTypeInfo(0, 0, &spTypeInfo);
	std::stringstream composite;
	std::list<std::string> elements;

	// init script: why? so we can map enums
	
	v8::Isolate* isolate = getInstanceIsolate();
	isolate->Enter();
	v8::Locker locker(isolate);

	// ATLTRACE("Isolate: %x\n", (unsigned long)isolate);

	v8::HandleScope handle_scope(isolate);
	if (_context.IsEmpty()) InitContext();
	v8::Local<v8::Context> context = v8::Local<v8::Context>::New(isolate, _context);

	context->Enter();

	// get type lib

	if (SUCCEEDED(hr) && spTypeInfo)
	{
		UINT tlidx;
		hr = spTypeInfo->GetContainingTypeLib(&spTypeLib, &tlidx);
	}

	if (SUCCEEDED(hr))
	{
		UINT tlcount = spTypeLib->GetTypeInfoCount();
		TYPEKIND tk;

		for (UINT u = 0; u < tlcount; u++)
		{
			if (SUCCEEDED(spTypeLib->GetTypeInfoType(u, &tk)))
			{
				CComBSTR bstrName;
				TYPEATTR *pTatt = nullptr;
				CComPtr<ITypeInfo> spTypeInfo2;
				hr = spTypeLib->GetTypeInfo(u, &spTypeInfo2);
				if (SUCCEEDED(hr)) hr = spTypeInfo2->GetTypeAttr(&pTatt);
				if (SUCCEEDED(hr)) hr = spTypeInfo2->GetDocumentation(-1, &bstrName, 0, 0, 0);
				if (SUCCEEDED(hr))
				{
					std::string tname;
					NarrowString(&bstrName, tname);
					
					
					std::string tmpstr;

					switch (tk)
					{
					case TKIND_ENUM:
						mapEnum(tmpstr, tname, spTypeInfo2, pTatt, context);
						composite << tmpstr;
						break;

					case TKIND_COCLASS:
						mapCoClass(tmpstr, tname, spTypeInfo2, pTatt);
						composite << tmpstr;
						break;

					case TKIND_INTERFACE:
					case TKIND_DISPATCH:
						mapInterface(tmpstr, tname, spTypeInfo2, spTypeLib, pTatt, tk);
						composite << tmpstr;
						break;

						//break;

					default:
						ATLTRACE("Unexpected tkind: %d\n", tk);
						break;
					}

					if (tmpstr.length() > 0) elements.push_back(tmpstr);

				}
				if (pTatt) spTypeInfo2->ReleaseTypeAttr(pTatt);
			}
		}
	}

	CComBSTR bstrComposite = "{\n";

	for (std::list< std::string > ::iterator iter = elements.begin(); iter != elements.end(); iter++)
	{
		if (bstrComposite.Length() > 2) bstrComposite.Append(",\n");
		bstrComposite.Append(iter->c_str());
	}

	bstrComposite.Append("\n}\n");
	bstrComposite.CopyTo(Description );

	context->Exit();
	isolate->Exit();

	return S_OK;
}

STDMETHODIMP CScripto::SetDispatch(IDispatch *Dispatch, BSTR* Name, VARIANT_BOOL MapEnums)
{
	std::string name;
	NarrowString(Name, name);

	v8::Isolate* isolate = getInstanceIsolate();
	isolate->Enter();
	v8::Locker locker(isolate);

	ATLTRACE("SetDispatch / Isolate: %x\n", (unsigned long)isolate);

	v8::HandleScope handle_scope(isolate);

	dispatch_list.insert(std::pair< std::string, IDispatch* >(name, Dispatch));

	if (_context.IsEmpty()) InitContext();
	v8::Local<v8::Context> context = v8::Local<v8::Context>::New(isolate, _context);
	context->Enter();
	
	v8::Local< v8::Object > instance = WrapDispatch(isolate, Dispatch);

	v8::Persistent<v8::Object> tmp;
	tmp.Reset(isolate, instance);

	context->Global()->Set(v8::String::NewFromUtf8(isolate, name.c_str()), instance);

	// do the mapping of enums into the global namespace here, 
	// since we're no longer consolidating that with the overall map
	// routine.

	// FIXME: again this could be done w/ JSON, since the object doesn't
	// change...

	if (MapEnums)
	{
		CComPtr<ITypeInfo> spTypeInfo;
		CComPtr<ITypeLib> spTypeLib;

		HRESULT hr = Dispatch->GetTypeInfo(0, 0, &spTypeInfo);

		if (SUCCEEDED(hr) && spTypeInfo)
		{
			UINT tlidx;
			hr = spTypeInfo->GetContainingTypeLib(&spTypeLib, &tlidx);
		}

		if (SUCCEEDED(hr))
		{
			UINT tlcount = spTypeLib->GetTypeInfoCount();
			TYPEKIND tk;

			for (UINT u = 0; u < tlcount; u++)
			{
				if (SUCCEEDED(spTypeLib->GetTypeInfoType(u, &tk)))
				{
					if ( tk == TKIND_ENUM )
					{
						CComBSTR bstrName;
						TYPEATTR *pTatt = nullptr;
						CComPtr<ITypeInfo> spTypeInfo2;
						hr = spTypeLib->GetTypeInfo(u, &spTypeInfo2);
						if (SUCCEEDED(hr)) hr = spTypeInfo2->GetTypeAttr(&pTatt);
						if (SUCCEEDED(hr)) hr = spTypeInfo2->GetDocumentation(-1, &bstrName, 0, 0, 0);
						if (SUCCEEDED(hr))
						{
							std::string tname;
							NarrowString(&bstrName, tname);
							std::string tmpstr;
							mapEnum(tmpstr, tname, spTypeInfo2, pTatt, context);
						}
						if (pTatt) spTypeInfo2->ReleaseTypeAttr(pTatt);
					}
				}
			}
		}
	}

	context->Exit();
	isolate->Exit();

	return S_OK;
}

STDMETHODIMP CScripto::ExecString(BSTR* Script, BSTR* Result, VARIANT_BOOL* Success)
{
	std::string result;
	bool success = false;

	// scoping v8
	{

		v8::Isolate* isolate = getInstanceIsolate();
		isolate->Enter();
		v8::Locker locker(isolate);

		v8::HandleScope handle_scope(isolate);

		if (_context.IsEmpty()) InitContext();

		v8::Local<v8::Context> context = v8::Local<v8::Context>::New(isolate, _context);

		CComBSTR wide(*Script);
		std::string narrow;

		int len = wide.Length();
		for (int i = 0; i < len; i++) narrow += (char)(wide.m_str[i] & 0xff);

		context->Enter();

		v8::Context::Scope context_scope(context);
		v8::Local<v8::String> name(v8::String::NewFromUtf8(context->GetIsolate(), "(shell)"));


		v8::TryCatch try_catch;
		v8::ScriptOrigin origin(name);
		v8::Handle<v8::Script> script = v8::Script::Compile(v8::String::NewFromUtf8(context->GetIsolate(), narrow.c_str()));

		if (script.IsEmpty()) {
			FormatException(isolate, &try_catch, result);
		}
		else {
			v8::Handle<v8::Value> script_result = script->Run();
			if (script_result.IsEmpty()) {
				FormatException(isolate, &try_catch, result);
			}
			else {
				if (!script_result->IsUndefined()) {

					v8::String::Utf8Value str(script_result);
					const char* cstr = ToCString(str);
					result = cstr;
				}
				success = true;
			}
		}

		context->Exit();

		v8::V8::LowMemoryNotification();
		while (!v8::V8::IdleNotification());

		isolate->Exit();
	}

	CComBSTR bstr = result.c_str();

	*Result = SysAllocString(bstr);
	*Success = success;

	return S_OK;
}

v8::Isolate* CScripto::getInstanceIsolate()
{
	if (instanceIsolate) return instanceIsolate;
	instanceIsolate = v8::Isolate::New();
	ATLTRACE("NEW ISOLATE: %x\n", (unsigned long)instanceIsolate);
	return instanceIsolate;
}

