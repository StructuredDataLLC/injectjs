// Scripto.cpp : Implementation of CScripto

#include "stdafx.h"
#include "Scripto.h"
#include "CVEventHandler.h"

#include <comdef.h>

#include <string>

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
// callback functions
//-----------------------------------------------------------

void indexed_getter(uint32_t index, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
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
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->Getter(property, info);
}

void setter(v8::Local<v8::String> property, v8::Local<v8::Value> value, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->Setter(property, value, info);
}

void invoker(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(args.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->Invoker(args);

}

void CScripto::SetRetVal(v8::Isolate* isolate, CComVariant &var, v8::ReturnValue< v8::Value > &retval)
{
	switch (var.vt)
	{
	case VT_DISPATCH:
		retval.Set(WrapDispatch(isolate, var.pdispVal));
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
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
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
	// printf("Event callback\n");

	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);

	v8::Local<v8::Context> context = v8::Local<v8::Context>::New(isolate, _context);

	context->Enter();

	v8::Context::Scope context_scope(context);

	v8::Handle< v8::Function > local = v8::Local<v8::Function>::New(isolate, func);
	v8::Handle< v8::Value > result = local->Call(context->Global(), 0, NULL);

	context->Exit();

	return S_OK;
}

void CScripto::Setter(v8::Local<v8::String> property, v8::Local<v8::Value> value, const v8::PropertyCallbackInfo<v8::Value>& info)
{
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);

	v8::Handle<v8::External> disp = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(1));
	CComPtr<IDispatch> pdisp((IDispatch*)(disp->Value()));

	std::string propname = ToCString(v8::String::Utf8Value(property));
	std::string valstr = ToCString(v8::String::Utf8Value(value));

	// printf("setter: %s => %s\n", propname.c_str(), valstr.c_str());

	if (!pdisp) return;

	CComBSTR wide;
	wide.Append(propname.c_str());

	WCHAR *member = (LPWSTR)wide;
	DISPID dispid;

	HRESULT hr = pdisp->GetIDsOfNames(IID_NULL, &member, 1, 1033, &dispid);

	if (FAILED(hr))
	{
		if (!value->IsFunction())
		{
			return;
		}
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
														printf("Nested deep");
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
															CComObject<CCVEventHandler>* pSink;
															MEMBERID memID2;
															hr = spTypeInfo3->GetIDsOfNames(&member, 1, &memID2);

															// we might have already sunk it
															v8::Handle<v8::External> check = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(2));
															pSink = ((CComObject<CCVEventHandler>*)(check->Value()));

															if (!pSink)
															{

																if (SUCCEEDED(hr)) hr = spTypeInfo3->GetDocumentation(-1, &bstr, 0, 0, 0);
																if (SUCCEEDED(hr)) hr = spTypeInfo3->GetTypeAttr(&pTatt2);
																if (SUCCEEDED(hr))
																{
																	// hr = pSinkClass->QueryInterface(IID_IUnknown, (LPVOID*)&m_pSinkUnk);


																	CComObject<CCVEventHandler>::CreateInstance(&pSink);
																	CComPtr<IUnknown> punk;
																	pSink->QueryInterface(IID_IUnknown, (void**)&punk);
																	pSink->SetPtr(this);

																	DWORD dwCookie;
																	hr = AtlAdvise(pdisp, punk, pTatt2->guid, &dwCookie);

																	info.Holder()->SetInternalField(2, v8::External::New(isolate, pSink));
																	info.Holder()->SetInternalField(3, v8::Int32::New(isolate, dwCookie));


																}
															}

															// now store the callback function in there

															pSink->Store(isolate, memID2, value.As<v8::Function>());

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
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
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
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);

	v8::Handle<v8::External> disp = v8::Handle<v8::External>::Cast(info.Holder()->GetInternalField(1));
	CComPtr<IDispatch> pdisp((IDispatch*)(disp->Value()));

	std::string propname = ToCString(v8::String::Utf8Value(property));
	printf("getter: %s\n", propname.c_str());

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
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);
	v8::Handle<v8::External> external = v8::Handle<v8::External>::Cast(args.Holder()->GetInternalField(0));

	CScripto *p = (CScripto*)(external->Value());
	p->LogMessage(args);

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
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);
	v8::Local< v8::External > external = v8::Handle<v8::External>::Cast(isolate->GetCallingContext()->Global()->GetHiddenValue(v8::String::NewFromUtf8(isolate, "inst")));
	CScripto *p = (CScripto*)(external->Value());
	p->Alert(args);
}

void confirm(const v8::FunctionCallbackInfo<v8::Value>& args)
{
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
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
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
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
	v8::V8::InitializeICU();
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
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
	// v8::Isolate* isolate = v8::Isolate::GetCurrent();
	// v8::HandleScope handle_scope(isolate);

	if (pdisp)
	{
		pdisp->AddRef();
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

STDMETHODIMP CScripto::SetDispatch(IDispatch *Dispatch, BSTR* Name)
{
	std::string name;
	NarrowString(Name, name);

	v8::Isolate* isolate = v8::Isolate::GetCurrent();
	v8::HandleScope handle_scope(isolate);

	dispatch_list.insert(std::pair< std::string, IDispatch* >(name, Dispatch));

	if (_context.IsEmpty()) InitContext();
	v8::Local<v8::Context> context = v8::Local<v8::Context>::New(isolate, _context);
	context->Enter();

	v8::Local< v8::Object > instance = WrapDispatch(isolate, Dispatch);

	v8::Persistent<v8::Object> tmp;
	tmp.Reset(isolate, instance);

	context->Global()->Set(v8::String::NewFromUtf8(isolate, name.c_str()), instance);

	context->Exit();

	return S_OK;
}

STDMETHODIMP CScripto::ExecString(BSTR* Script, BSTR* Result, VARIANT_BOOL* Success)
{
	v8::Isolate* isolate = v8::Isolate::GetCurrent();
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

	std::string result;

	v8::TryCatch try_catch;
	v8::ScriptOrigin origin(name);
	v8::Handle<v8::Script> script = v8::Script::Compile(v8::String::NewFromUtf8(context->GetIsolate(), narrow.c_str()));

	bool success = false;

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

	// printf("ok: %s\n", result.c_str());

	CComBSTR bstr = result.c_str();

	*Result = SysAllocString(bstr);
	*Success = success;

	return S_OK;
}


