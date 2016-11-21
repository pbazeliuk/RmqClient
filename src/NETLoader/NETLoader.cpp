/*
* ***** BEGIN LICENSE BLOCK *****
* Version: MIT
*
* Copyright (c) 2014-2016 Igor Medvedev
*
* Permission is hereby granted, free of charge, to any person
* obtaining a copy of this software and associated documentation
* files (the "Software"), to deal in the Software without
* restriction, including without limitation the rights to use, copy,
* modify, merge, publish, distribute, sublicense, and/or sell copies
* of the Software, and to permit persons to whom the Software is
* furnished to do so, subject to the following conditions:
*
* The above copyright notice and this permission notice shall be
* included in all copies or substantial portions of the Software.
*
* THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
* EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
* MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
* NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS
* BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN
* ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN
* CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
* SOFTWARE.
* ***** END LICENSE BLOCK *****
*/

#include "stdafx.h"

#include <stdio.h>
#include <wchar.h>
#include "NETLoader.h"
#include <windows.h>

static wchar_t *g_MethodNames[] = { L"GetLastError", L"GetInstalledRuntimes", L"CreateObject" };
static wchar_t *g_PropNames[] = { L"CurrentVersionCLR" };
//static const wchar_t g_kClassNames[] = L"NETLoader"; //|OtherClass1|OtherClass2";

uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len = 0);
uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len = 0);
uint32_t getLenShortWcharStr(const WCHAR_T* Source);

static AppCapabilities g_capabilities = eAppCapabilitiesInvalid;


using namespace std;

MethInfo::MethInfo()
{
	
}

MethInfo::~MethInfo()
{
	if (Name)
	{
		delete[] Name;
	}

	if (info)
	{
		info->Release();
		info = NULL;
	}
}

PropInfo::PropInfo()
{

}

PropInfo::~PropInfo()
{
	if (Name)
	{
		delete[] Name;
	}

	if (info)
	{
		info->Release();
		info = NULL;
	}
}

long GetClassObject(const wchar_t* wsName, IComponentBase** pInterface)
{
	if(!*pInterface)
    {
        *pInterface= new CAddInNative();
        return (long)*pInterface;
    }
    return 0;
}

AppCapabilities SetPlatformCapabilities(const AppCapabilities capabilities)
{
	g_capabilities = capabilities;
    return eAppCapabilitiesLast;
}

long DestroyObject(IComponentBase** pIntf)
{
   if(!*pIntf)
      return -1;

   delete *pIntf;
   *pIntf = 0;
   return 0;
}

const WCHAR_T* GetClassNames()
{
	return L"NETLoader";

	/*static WCHAR_T* names = 0;
    if (!names)
        ::convToShortWchar(&names, g_kClassNames);
    return names;*/
}

//CAddInNative
CAddInNative::CAddInNative()
{
	
}

CAddInNative::~CAddInNative()
{
	HRESULT hr;

	SetErrorDescription(NULL);

	//VariantClear(&NETObject);
	NETObject = NULL;
	
	if (NETType)
	{
		NETType->Release();
		NETType = NULL;
	}

	if (NETMethods)
	{
		delete[] NETMethods;
	}

	if (NETProperties)
	{
		delete[] NETProperties;
	}

	if (AppDomain)
	{
		hr = CorRuntimeHost->UnloadDomain(AppDomain);
		//AppDomain->Release();
		AppDomain = NULL;
	}

	if (CorRuntimeHost)
	{
		CorRuntimeHost->Release();
		CorRuntimeHost = NULL;
	}
	
	if (CurrentVersionCLR)
	{
		delete[] CurrentVersionCLR;
	}

	if (RemoveAppDirectory)
	{
		SHFILEOPSTRUCT fileop;
		fileop.hwnd = NULL;    // no status display
		fileop.wFunc = FO_DELETE;  // delete operation
		fileop.pFrom = RemoveAppDirectory;  // source file name as double null terminated string
		fileop.pTo = NULL;    // no destination needed
		fileop.fFlags = FOF_NOCONFIRMATION | FOF_SILENT;  // do not prompt the user

		/*if (!noRecycleBin)
			fileop.fFlags |= FOF_ALLOWUNDO;*/

		fileop.fAnyOperationsAborted = FALSE;
		fileop.lpszProgressTitle = NULL;
		fileop.hNameMappings = NULL;

		int ret = SHFileOperation(&fileop);
		
		delete[] RemoveAppDirectory;
		RemoveAppDirectory = NULL;

		
	}
}

bool CAddInNative::Init(void* pConnection)
{ 
	m_iConnect = (IAddInDefBase*)pConnection;
	return m_iConnect != NULL;

	return true;
}

long CAddInNative::GetInfo()
{ 
	return 2000; 
}

void CAddInNative::Done()
{
	
}

bool CAddInNative::RegisterExtensionAs(WCHAR_T** wsExtensionName)
{ 
	wchar_t *wsExtension = L"NETLoader";
    int iActualSize = ::wcslen(wsExtension) + 1;
    WCHAR_T* dest = 0;

    if (m_iMemory)
    {
        if(m_iMemory->AllocMemory((void**)wsExtensionName, iActualSize * sizeof(WCHAR_T)))
            ::convToShortWchar(wsExtensionName, wsExtension, iActualSize);
        return true;
    }

    return false; 
}

long CAddInNative::GetNProps()
{ 
	return eLastProp + NETPropertiesCount;
}

long CAddInNative::FindProp(const WCHAR_T* PropName)
{ 
	long PropNum = -1;
	wchar_t* name = 0;

	convFromShortWchar(&name, PropName);

	PropNum = findName(g_PropNames, name, eLastProp);

	delete[] name;

	if (PropNum >= 0)
	{
		return PropNum;
	}

	for (long i = 0; i < NETPropertiesCount; i++)
	{
		if (!wcscmp(NETProperties[i].Name, PropName))
		{
			PropNum = i + eLastProp;
			break;
		}
	}

	return PropNum;
}

const WCHAR_T* CAddInNative::GetPropName(long PropNum, long PropAlias)
{ 
	if (PropNum >= eLastProp + NETPropertiesCount)
		return NULL;

	wchar_t *CurrentName = NULL;
	WCHAR_T *PropName = NULL;
	int iActualSize = 0;

	switch (PropAlias)
	{
	case 0: // First language
		if (PropNum >= eLastProp)
		{
			CurrentName = NETProperties[PropNum - eLastProp].Name;
		}
		else
		{
			CurrentName = g_PropNames[PropNum];
		}

		break;
	default:
		return 0;
	}

	iActualSize = wcslen(CurrentName) + 1;

	if (m_iMemory && CurrentName)
	{
		if (m_iMemory->AllocMemory((void**)&PropName, iActualSize * sizeof(WCHAR_T)))
		{
			convToShortWchar(&PropName, CurrentName, iActualSize);
		}
	}

	return PropName;
}

bool CAddInNative::GetPropVal(const long PropNum, tVariant* PropVal)
{ 
	if (PropNum >= eLastProp)
	{
		VARIANT value;
		IUnknown* info;
		HRESULT hr;

		info = NETProperties[PropNum - eLastProp].info;

		if (NETProperties[PropNum - eLastProp].IsField)
		{
			hr = ((_FieldInfoPtr)info)->GetValue(NETObject, &value);
		}
		else
		{
			hr = ((_PropertyInfoPtr)info)->GetValue(NETObject, NULL, &value);
		}
		
		if (FAILED(hr))
		{
			SetErrorDescription();
			return false;
		}

		if (!ConvertFromVariantTo1C(&value, PropVal))
		{
			SetErrorDescription(L"Ошибка преобразования значения");
			return false;
		}
		return true;
	}
	else
	{
		switch (PropNum)
		{
		case eCurrentVersionCLR:
			TV_VT(PropVal) = VTYPE_PWSTR;
			PropVal->wstrLen = wcslen(CurrentVersionCLR);
			if (!m_iMemory->AllocMemory((void**)&PropVal->pwstrVal, (PropVal->wstrLen + 1) * sizeof(WCHAR_T)))
			{
				return false;
			}

			convToShortWchar(&PropVal->pwstrVal, CurrentVersionCLR, PropVal->wstrLen + 1);
			return true;
		}
	}

	return false;
}

bool CAddInNative::SetPropVal(const long PropNum, tVariant* PropVal)
{ 
	if (PropNum >= eLastProp)
	{
		VARIANT value;
		IUnknown* info;
		HRESULT hr;

		info = NETProperties[PropNum - eLastProp].info;

		if (!ConvertFrom1CToVariant(PropVal, &value))
		{
			SetErrorDescription(L"Ошибка преобразования значения");
			return false;
		}

		if (NETProperties[PropNum - eLastProp].IsField)
		{
			hr = ((_FieldInfoPtr)info)->SetValue_2(NETObject, value);
		}
		else
		{
			hr = ((_PropertyInfoPtr)info)->SetValue(NETObject, value, NULL);
		}

		
		if (FAILED(hr))
		{
			SetErrorDescription();
			return false;
		}

		return true;
	}
	
	return false;
}

bool CAddInNative::IsPropReadable(const long PropNum)
{ 
	if (PropNum >= eLastProp)
	{
		return NETProperties[PropNum - eLastProp].CanRead;
	}
	else
		switch (PropNum)
	{
		case eCurrentVersionCLR:
			return true;
	}

	return false;
}

bool CAddInNative::IsPropWritable(const long PropNum)
{
	if (PropNum >= eLastProp)
	{
		return NETProperties[PropNum - eLastProp].CanWrite;
	}

	return false;
}

long CAddInNative::GetNMethods()
{ 
	return eLastMethod + NETMethodsCount;
}

long CAddInNative::FindMethod(const WCHAR_T* MethodName)
{ 
	long MethodNum = -1;
    wchar_t* name = 0;

    convFromShortWchar(&name, MethodName);

    MethodNum = findName(g_MethodNames, name, eLastMethod);

    delete[] name;

	if (MethodNum >= 0)
	{
		return MethodNum;
	}

	for (long i = 0; i < NETMethodsCount; i++)
	{
		if (!wcscmp(NETMethods[i].Name, MethodName))
		{
			MethodNum = i + eLastMethod;
			break;
		}
	}

    return MethodNum;
}

const WCHAR_T* CAddInNative::GetMethodName(const long MethodNum, const long MethodAlias)
{ 
	if (MethodNum >= eLastMethod + NETMethodsCount)
        return NULL;

    wchar_t *CurrentName = NULL;
    WCHAR_T *MethodName = NULL;
    int iActualSize = 0;

    switch(MethodAlias)
    {
    case 0: // First language
		if (MethodNum >= eLastMethod)
		{
			CurrentName = NETMethods[MethodNum - eLastMethod].Name;
		}
		else
		{
			CurrentName = g_MethodNames[MethodNum];
		}
		
        break;
    default: 
        return 0;
    }

    iActualSize = wcslen(CurrentName)+1;

    if (m_iMemory && CurrentName)
    {
        if(m_iMemory->AllocMemory((void**)&MethodName, iActualSize * sizeof(WCHAR_T)))
            ::convToShortWchar(&MethodName, CurrentName, iActualSize);
    }

    return MethodName;
}

long CAddInNative::GetNParams(const long MethodNum)
{ 
	if (MethodNum >= eLastMethod)
	{
		return NETMethods[MethodNum - eLastMethod].MaxParam;
	}

	switch(MethodNum)
    { 
    case eGetLastError:
		return 0;
	case eGetInstalledRuntimes:
		return 0;
	case eCreateObject:
		return 5;
	default:
        return 0;
    }
    
    return 0;
}

bool CAddInNative::GetParamDefValue(const long MethodNum, const long ParamNum, tVariant *ParamDefValue)
{ 
	switch (MethodNum)
	{
	case eCreateObject:


		switch (ParamNum)
		{
		case 3:
			TV_VT(ParamDefValue) = VTYPE_PWSTR;
			ParamDefValue->wstrLen = 0;
			ParamDefValue->pwstrVal = NULL;
			return true;
		case 4:
			TV_VT(ParamDefValue) = VTYPE_BOOL;
			ParamDefValue->bVal = 0;
			return true;
		};
	}
	return false;
} 

bool CAddInNative::HasRetVal(const long MethodNum)
{ 
	if (MethodNum >= eLastMethod)
	{
		return NETMethods[MethodNum - eLastMethod].HasRetVal;
	}

	switch (MethodNum)
	{
	case eGetLastError:
		return true;
	case eGetInstalledRuntimes:
		return true;
	default:
		return false;
	}
}

bool CAddInNative::CallAsProc(const long MethodNum, tVariant* Params, const long SizeArray)
{ 
	if (MethodNum >= eLastMethod)
	{
		return InvokeMethod(MethodNum, Params, SizeArray, NULL);
	}

	switch(MethodNum)
    { 
    case eCreateObject:
		return CreateObject(Params, SizeArray);
	default:
        return false;
    }
}

bool CAddInNative::CallAsFunc(const long MethodNum, tVariant* RetValue, tVariant* Params, const long SizeArray)
{ 
	if (MethodNum >= eLastMethod)
	{
		return InvokeMethod(MethodNum, Params, SizeArray, RetValue);
	}

	switch (MethodNum)
	{
	case eGetLastError:
		return GetLastError(RetValue);
	case eGetInstalledRuntimes:
		return GetInstalledRuntimes(RetValue);
	default:
		return false;
	}
}

void CAddInNative::SetLocale(const WCHAR_T* loc)
{
	_wsetlocale(LC_ALL, loc);
}

bool CAddInNative::setMemManager(void* mem)
{
	m_iMemory = (IMemoryManager*)mem;
    return m_iMemory != 0;
}

long CAddInNative::findName(wchar_t* names[], const wchar_t* name, const uint32_t size) const
{
    long ret = -1;
    for (uint32_t i = 0; i < size; i++)
    {
        if (!wcscmp(names[i], name))
        {
            ret = i;
			break;
        }
    }
    return ret;
}

uint32_t convToShortWchar(WCHAR_T** Dest, const wchar_t* Source, uint32_t len)
{
    if (!len)
        len = ::wcslen(Source)+1;

    if (!*Dest)
        *Dest = new WCHAR_T[len];

    WCHAR_T* tmpShort = *Dest;
    wchar_t* tmpWChar = (wchar_t*) Source;
    uint32_t res = 0;

    ::memset(*Dest, 0, len*sizeof(WCHAR_T));
    do
    {
        *tmpShort++ = (WCHAR_T)*tmpWChar++;
        ++res;
    }
    while (len-- && *tmpWChar);

    return res;
}

uint32_t convFromShortWchar(wchar_t** Dest, const WCHAR_T* Source, uint32_t len)
{
    if (!len)
        len = getLenShortWcharStr(Source)+1;

    if (!*Dest)
        *Dest = new wchar_t[len];

    wchar_t* tmpWChar = *Dest;
    WCHAR_T* tmpShort = (WCHAR_T*)Source;
    uint32_t res = 0;

    ::memset(*Dest, 0, len*sizeof(wchar_t));
    do
    {
        *tmpWChar++ = (wchar_t)*tmpShort++;
        ++res;
    }
    while (len-- && *tmpShort);

    return res;
}

uint32_t getLenShortWcharStr(const WCHAR_T* Source)
{
    uint32_t res = 0;
    WCHAR_T *tmpShort = (WCHAR_T*)Source;

    while (*tmpShort++)
        ++res;

    return res;
}

void CAddInNative::SetErrorDescription()
{
	SetErrorDescription(NULL);

	IErrorInfo* err = NULL;
	HRESULT hr = GetErrorInfo(0, &err);
	if (hr != S_OK)
	{
		return;
	}

	BSTR desc = NULL;
	hr = err->GetDescription(&desc);
	if (hr != S_OK)
	{
		return;
	}

	//wchar_t * desc_(desc);
	
	SetErrorDescription(desc);

	SysFreeString(desc);
}

void CAddInNative::SetErrorDescription(wchar_t* desc)
{
	if (ErrorDescription)
	{
		delete[]ErrorDescription;
		ErrorDescription = NULL;
	}

	if (!desc)
	{
		return;
	}

	uint32_t Size = wcslen(desc) + 1;

	if (Size == 1)
	{
		return;
	}

	ErrorDescription = new wchar_t[Size];
	wcscpy_s(ErrorDescription, Size, desc);
}

bool CAddInNative::GetLastError(tVariant* RetValue)
{
	TV_VT(RetValue) = VTYPE_PWSTR;
	RetValue->wstrLen = 0;

	if (!ErrorDescription)
	{
		return true;
	}

	if (m_iMemory)
	{
		uint32_t Size = wcslen(ErrorDescription) + 1;
		if (m_iMemory->AllocMemory((void**)&TV_WSTR(RetValue), Size * sizeof(WCHAR_T)))
		{
			wcscpy_s(TV_WSTR(RetValue), Size, ErrorDescription);
			RetValue->wstrLen = Size - 1;
		}
	}

	return true;
}

bool CAddInNative::GetInstalledRuntimes(tVariant* RetValue)
{
	HRESULT hr;

	ICLRMetaHost *MetaHost = NULL;
	ICLRRuntimeInfo *RuntimeInfo = NULL;
	IUnknown* runtime = NULL;	

	// Load and start the .NET runtime.
	hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&MetaHost));
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	IEnumUnknown* runtimeEnumerator = nullptr;
	hr = MetaHost->EnumerateInstalledRuntimes(&runtimeEnumerator);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	DWORD bufsize;
	long wstrLen = 0;
	// первый проход, чтобы просто узнать размер строки
	while (runtimeEnumerator->Next(1, &runtime, NULL) == S_OK)
	{
		hr = runtime->QueryInterface(IID_PPV_ARGS(&RuntimeInfo));
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Cleanup;
		}

		RuntimeInfo->GetVersionString(NULL, &bufsize);
		wstrLen += bufsize; // последний символ NULL, но вместо него будем использовать ";"

		runtime->Release();
		RuntimeInfo->Release();
	}

	TV_VT(RetValue) = VTYPE_PWSTR;
	if (!m_iMemory->AllocMemory((void**)&RetValue->pwstrVal, wstrLen * sizeof(WCHAR_T)))
	{
		return false;
	}

	// сам буфер
	wchar_t buffer[50];
	// индекс, с которого нужно выводить строку
	long index = 0;

	bool ThisFirstRow = true;

	runtimeEnumerator->Reset();
	while (runtimeEnumerator->Next(1, &runtime, NULL) == S_OK)
	{
		hr = runtime->QueryInterface(IID_PPV_ARGS(&RuntimeInfo));
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Cleanup;
		}

		RuntimeInfo->GetVersionString(buffer, &bufsize);
		long len = wcslen(buffer);

		if (ThisFirstRow)
		{
			ThisFirstRow = false;
		}
		else{
			wcscpy(&RetValue->pwstrVal[index], L";");
			index++;
		}		

		wcscpy(&RetValue->pwstrVal[index], buffer);

		index += len;

		runtime->Release();
		runtime = NULL;
		RuntimeInfo->Release();
		RuntimeInfo = NULL;
	}

	RetValue->wstrLen = wcslen(RetValue->pwstrVal);

Cleanup:

	if (MetaHost)
	{
		MetaHost->Release();
		MetaHost = NULL;
	}

	if (runtime)
	{
		runtime->Release();
		runtime = NULL;
	}

	if (RuntimeInfo)
	{
		RuntimeInfo->Release();
		RuntimeInfo = NULL;
	}

	return SUCCEEDED(hr);
}

bool CAddInNative::FillMethods()
{
	HRESULT hr = S_OK;

	SAFEARRAY *Methods = NULL;
	IUnknown* var = NULL;
	BSTR name = NULL;
	SAFEARRAY *Parameters = NULL;
	MethInfo* methinfo = NULL;
	_Type *RetType = NULL;
	BSTR RetVal = NULL;


	hr = NETType->GetMethods(static_cast<BindingFlags>(BindingFlags_InvokeMethod | BindingFlags_Instance | BindingFlags_Public), &Methods);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Clean;
	}

	long last;
	
	SafeArrayGetUBound(Methods, SafeArrayGetDim(Methods), &last);

	NETMethodsCount = last + 1;

	NETMethods = new MethInfo[NETMethodsCount];

	//hr = SafeArrayAccessData(Methods, (void**)&var);
	for (long i = 0; i <= last; i++)
	{
		methinfo = &NETMethods[i];

		hr = SafeArrayGetElement(Methods, &i, &var);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}
		
		hr = var->QueryInterface(IID_PPV_ARGS(&methinfo->info));
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}
		
		// получение имени метода
		hr = methinfo->info->get_name(&name);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}		
		long NameSize = wcslen(name) + 1;
		methinfo->Name = new wchar_t[NameSize];
		wcscpy_s(methinfo->Name, NameSize, name);

		// получение параметров метода
		methinfo->info->GetParameters(&Parameters);
			
		hr = SafeArrayGetUBound(Parameters, SafeArrayGetDim(Parameters), &methinfo->MaxParam);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}
		methinfo->MaxParam++;

		// получение возвращаемого значения
		hr = methinfo->info->get_returnType(&RetType);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}

		RetType->get_ToString(&RetVal);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}

		methinfo->HasRetVal = wcscmp(RetVal, L"System.Void") != 0;

		// Очистка временных переменных

		if (var)
		{
			var->Release();
			var = NULL;
		}

		SysFreeString(name);
		name = NULL;

		SafeArrayDestroy(Parameters);
		Parameters = NULL;

		if (RetType)
		{
			RetType->Release();
			RetType = NULL;
		}

		SysFreeString(RetVal);
		RetVal = NULL;
	}
	
	//hr = SafeArrayUnaccessData(Methods);	

Clean:

	SafeArrayDestroy(Methods);

	if (var)
	{
		var->Release();
		var = NULL;
	}

	if (name)
	{
		SysFreeString(name);
		name = NULL;
	}

	if (Parameters)
	{
		SafeArrayDestroy(Parameters);
		Parameters = NULL;
	}

	if (RetType)
	{
		RetType->Release();
		RetType = NULL;
	}

	if (RetVal)
	{
		SysFreeString(RetVal);
		RetVal = NULL;
	}

	return SUCCEEDED(hr);
}

bool CAddInNative::FillProperties()
{
	HRESULT hr = S_OK;

	SAFEARRAY *Properties = NULL;
	SAFEARRAY *Fields = NULL;
	IUnknown* var = NULL;
	BSTR name = NULL;
	
	// вспомогательные переменные, память для них не нужно освобождать	
	PropInfo* propinfo = NULL;
	_PropertyInfoPtr p_info = NULL;
	_FieldInfoPtr f_info = NULL;

	hr = NETType->GetProperties(static_cast<BindingFlags>(BindingFlags_Instance | BindingFlags_Public), &Properties);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Clean;
	}

	hr = NETType->GetFields(static_cast<BindingFlags>(BindingFlags_Instance | BindingFlags_Public), &Fields);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Clean;
	}


	long p_last;
	long f_last;

	SafeArrayGetUBound(Properties, SafeArrayGetDim(Properties), &p_last);
	SafeArrayGetUBound(Fields, SafeArrayGetDim(Properties), &f_last);

	NETPropertiesCount = p_last + f_last + 2;

	NETProperties = new PropInfo[NETPropertiesCount];

	for (long i = 0; i <= p_last; i++)
	{
		propinfo = &NETProperties[i];

		hr = SafeArrayGetElement(Properties, &i, &var);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}

		hr = var->QueryInterface(IID_PPV_ARGS(&p_info));
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}

		propinfo->info = p_info;

		// получение имени свойства
		hr = p_info->get_name(&name);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}
		long NameSize = wcslen(name) + 1;
		propinfo->Name = new wchar_t[NameSize];
		wcscpy_s(propinfo->Name, NameSize, name);

		VARIANT_BOOL b;
		p_info->get_CanRead(&b);
		propinfo->CanRead = b == VARIANT_TRUE;

		p_info->get_CanWrite(&b);
		propinfo->CanWrite = b == VARIANT_TRUE;

		propinfo->IsField = false;

		// Очистка временных переменных

		if (var)
		{
			var->Release();
			var = NULL;
		}

		SysFreeString(name);
		name = NULL;
	}

	for (long i = 0; i <= f_last; i++)
	{
		propinfo = &NETProperties[i + p_last + 1];

		hr = SafeArrayGetElement(Fields, &i, &var);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}

		hr = var->QueryInterface(IID_PPV_ARGS(&f_info));
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}

		propinfo->info = f_info;

		// получение имени свойства
		hr = f_info->get_name(&name);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}
		long NameSize = wcslen(name) + 1;
		propinfo->Name = new wchar_t[NameSize];
		wcscpy_s(propinfo->Name, NameSize, name);

		propinfo->CanRead = true;
		propinfo->CanWrite = true;
		propinfo->IsField = true;

		// Очистка временных переменных

		if (var)
		{
			var->Release();
			var = NULL;
		}

		SysFreeString(name);
		name = NULL;
	}

Clean:

	SafeArrayDestroy(Properties);
	SafeArrayDestroy(Fields);

	if (var)
	{
		var->Release();
		var = NULL;
	}

	if (name)
	{
		SysFreeString(name);
		name = NULL;
	}

	return SUCCEEDED(hr);
}

void CAddInNative::CreateAppDomain(wchar_t* ApplicationBase, wchar_t* VersionCLR)
{
	HRESULT hr;	

	ICLRMetaHost *MetaHost = NULL;
	ICLRRuntimeInfo *RuntimeInfo = NULL;	
	IUnknownPtr DomainSetupThunk = NULL;
	IAppDomainSetupPtr DomainSetup = NULL;
	IUnknownPtr AppDomainThunk = NULL;
	
	// Load and start the .NET runtime.
	hr = CLRCreateInstance(CLSID_CLRMetaHost, IID_PPV_ARGS(&MetaHost));
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	if (VersionCLR && wcslen(VersionCLR))
	{
		hr = MetaHost->GetRuntime(VersionCLR, IID_PPV_ARGS(&RuntimeInfo));
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Cleanup;
		}
	}
	else
	{
		if (!GetLastRuntime(MetaHost, &RuntimeInfo))
		{
			goto Cleanup;
		}
	}	

	wchar_t currentRuntime[50];
	DWORD bufferSize = ARRAYSIZE(currentRuntime);

	RuntimeInfo->GetVersionString(currentRuntime, &bufferSize);
	long size = wcslen(currentRuntime) + 1;
	CurrentVersionCLR = new wchar_t[size];
	wcscpy(CurrentVersionCLR, currentRuntime);

	// Check if the specified runtime can be loaded into the process. This 
	// method will take into account other runtimes that may already be 
	// loaded into the process and set pbLoadable to TRUE if this runtime can 
	// be loaded in an in-process side-by-side fashion. 
	BOOL fLoadable;
	hr = RuntimeInfo->IsLoadable(&fLoadable);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	if (!fLoadable)
	{
		SetErrorDescription(L".NET runtime cannot be loaded");
		goto Cleanup;
	}

	// Load the CLR into the current process and return a runtime interface 
	// pointer. ICorRuntimeHost and ICLRRuntimeHost are the two CLR hosting  
	// interfaces supported by CLR 4.0. Here we demo the ICorRuntimeHost 
	// interface that was provided in .NET v1.x, and is compatible with all 
	// .NET Frameworks. 
	hr = RuntimeInfo->GetInterface(CLSID_CorRuntimeHost, IID_PPV_ARGS(&CorRuntimeHost));
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	// Start the CLR.
	hr = CorRuntimeHost->Start();
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}	

	hr = CorRuntimeHost->CreateDomainSetup(&DomainSetupThunk);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	hr = DomainSetupThunk->QueryInterface(IID_PPV_ARGS(&DomainSetup));
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	BSTR applicationBase = SysAllocString(ApplicationBase);
	DomainSetup->put_ApplicationBase(applicationBase);
	SysFreeString(applicationBase);

	// Get a pointer to the default AppDomain in the CLR.
	//hr = CorRuntimeHost->CreateDomain(L"NETLoader", NULL, &AppDomainThunk);
	hr = CorRuntimeHost->CreateDomainEx(L"NETLoader", DomainSetup, NULL, &AppDomainThunk);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	hr = AppDomainThunk->QueryInterface(IID_PPV_ARGS(&AppDomain));
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

Cleanup:

	if (MetaHost)
	{
		MetaHost->Release();
		MetaHost = NULL;
	}

	if (RuntimeInfo)
	{
		RuntimeInfo->Release();
		RuntimeInfo = NULL;
	}	

	if (DomainSetupThunk)
	{
		DomainSetupThunk->Release();
		DomainSetupThunk = NULL;
	}

	if (DomainSetup)
	{
		DomainSetup->Release();
		DomainSetup = NULL;
	}

	if (AppDomainThunk)
	{
		AppDomainThunk->Release();
		AppDomainThunk = NULL;
	}
}

bool CAddInNative::GetLastRuntime(ICLRMetaHost *MetaHost, ICLRRuntimeInfo **RuntimeInfo)
{
	HRESULT hr;	

	IEnumUnknown* runtimeEnumerator = NULL;
	IUnknown* runtime = NULL;

	hr = MetaHost->EnumerateInstalledRuntimes(&runtimeEnumerator);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	long last = 0;
	while (runtimeEnumerator->Skip(1) == S_OK)
	{
		last++;
	}	

	hr = runtimeEnumerator->Reset();
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	hr = runtimeEnumerator->Skip(last - 1);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	hr = runtimeEnumerator->Next(1, &runtime, NULL);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	hr = runtime->QueryInterface(IID_PPV_ARGS(RuntimeInfo));
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}
	
Cleanup:

	if (runtimeEnumerator)
	{
		runtimeEnumerator->Release();
		runtimeEnumerator = NULL;
	}

	if (runtime)
	{
		runtime->Release();
		runtime = NULL;
	}

	return SUCCEEDED(hr);
}

SAFEARRAY* CAddInNative::GetAssemblyData(wchar_t* AssemblyPath)
{
	// Чтение файла
	int ret = 0;

	struct _stat stat;
	ret = _wstat(AssemblyPath, &stat);
	if (ret)
	{
		SetErrorDescription(L"Assembly file not found!");
		return NULL;
	}

	ULONG FileDataSize = stat.st_size;

	BYTE *FileData = NULL;
	FileData = new BYTE[FileDataSize];
	if (!FileData)
	{
		SetErrorDescription(L"Assembly file not load: not enough memory!");
		return NULL;
	}

	FILE *file_in = _wfopen(AssemblyPath, L"rb");
	ret = fread(FileData, sizeof(BYTE), FileDataSize, file_in);
	if (ret != FileDataSize)
	{
		SetErrorDescription(L"Error in reading assembly file!");
		return NULL;
	}
	fclose(file_in);

	SAFEARRAY * rawAssembly = SafeArrayCreateVector(VT_UI1, 0, FileDataSize);

	BYTE *NewFileData;
	SafeArrayAccessData(rawAssembly, (void **)&NewFileData);

	for (unsigned long i = 0; i < FileDataSize; i++)
	{
		NewFileData[i] = FileData[i];
		//hr = SafeArrayPutElement(rawAssembly, &i, &FileData[i]);
	}

	SafeArrayUnaccessData(rawAssembly);

	if (FileData)
		delete FileData;

	return rawAssembly;
}

bool CAddInNative::CreateObject(tVariant* paParams, const long lSizeArray)
{
	if (NETObjectIsLoad)
	{
		SetErrorDescription(L"Объект уже создан");
		return false;
	}

	// ПРОВЕРКА ПАРАМЕТРОВ
	if (lSizeArray != 5)
	{
		SetErrorDescription(L"Должно быть 5 параметра");
		return false;
	}

	if (TV_VT(paParams) != VTYPE_PWSTR)
	{
		SetErrorDescription(L"Параметр № 1 должен быть типа Строка");
		return false;
	}

	if (TV_VT(&(paParams[1])) != VTYPE_PWSTR)
	{
		SetErrorDescription(L"Параметр № 2 должен быть типа строка");
		return false;
	}

	if (TV_VT(&(paParams[2])) != VTYPE_PWSTR)
	{
		SetErrorDescription(L"Параметр № 3 должен быть типа строка");
		return false;
	}

	if (TV_VT(&(paParams[3])) != VTYPE_PWSTR)
	{
		SetErrorDescription(L"Параметр № 4 должен быть типа строка");
		return false;
	}

	if (TV_VT(&(paParams[4])) != VTYPE_BOOL)
	{
		SetErrorDescription(L"Параметр № 5 должен быть типа булево");
		return false;
	}
	
	// ПОЛУЧЕНИЕ ДОМЕНА ПО УМОЛЧАНИЮ
	CreateAppDomain(paParams[0].pwstrVal, paParams[3].pwstrVal);
	if (!AppDomain)
	{
		return false;
	}

	// если при удалении объекта нужно удалить каталог, то сохраним его
	if (paParams[4].bVal)
	{
		long len = wcslen(paParams[0].pwstrVal);
		RemoveAppDirectory = new wchar_t[len + 2]; // должно быть два 2 в конце
		wcscpy(RemoveAppDirectory, paParams[0].pwstrVal);
		RemoveAppDirectory[len + 1] = 0;
	}

	// ЗАГРУЗКА СБОРКИ, СОЗДАНИЕ ОБЪЕКТА, ПОЛУЧЕНИЕ ИНФОРМАЦИИ ОБ ОБЪЕКТЕ
	HRESULT hr;
	_AssemblyPtr Assembly = NULL;
	

	BSTR AssemblyString = SysAllocString(paParams[1].pwstrVal);
	hr = AppDomain->Load_2(AssemblyString, &Assembly);
	SysFreeString(AssemblyString);
	if (FAILED(hr))
	{
		SetErrorDescription();		
		goto Cleanup;
	}	

	BSTR ClassName = SysAllocString(paParams[2].pwstrVal);
	hr = Assembly->GetType_2(ClassName, &NETType);
	SysFreeString(ClassName);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	if (!NETType)
	{
		SetErrorDescription(L"Класс с таким именем не найден");
		goto Cleanup;
	}

	hr = Assembly->CreateInstance(ClassName, &NETObject);
	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Cleanup;
	}

	if (!FillMethods())
	{
		goto Cleanup;
	}

	if (!FillProperties())
	{
		goto Cleanup;
	}

	NETObjectIsLoad = true;

Cleanup:

	//SafeArrayDestroy(rawAssembly);

	if (Assembly)
	{
		Assembly->Release();
		Assembly = NULL;
	}

	return SUCCEEDED(hr) && NETType;
}

bool CAddInNative::ConvertFrom1CToVariant(tVariant* Param, VARIANT* Var)
{
	LPSYSTEMTIME Time = NULL;
	int res;

	switch TV_VT(Param)
	{
		// Целое число
	case VTYPE_I2:
	case VTYPE_I4:
	case VTYPE_ERROR:
	case VTYPE_UI1:

		V_VT(Var) = VT_I4;
		V_I4(Var) = Param->lVal;

		return true;

		// Булево
	case VTYPE_BOOL:

		V_VT(Var) = VT_BOOL;
		V_BOOL(Var) = TV_BOOL(Param);
		return true;

		// Дробное число
	case VTYPE_R4:
	case VTYPE_R8:

		V_VT(Var) = VT_R8;
		V_R8(Var) = Param->dblVal;
		return true;

		// Дата
	case VTYPE_DATE:

		V_VT(Var) = VT_DATE;
		V_DATE(Var) = Param->date;
		return true;

	case VTYPE_TM:

		Time = new SYSTEMTIME();

		Time->wYear = Param->tmVal.tm_year + 1900;
		Time->wMonth = Param->tmVal.tm_mon;
		Time->wDay = Param->tmVal.tm_mday;
		Time->wDayOfWeek = Param->tmVal.tm_wday;
		Time->wHour = Param->tmVal.tm_hour;
		Time->wMinute = Param->tmVal.tm_min;
		Time->wSecond = Param->tmVal.tm_sec;
		Time->wMilliseconds = 0;

		V_VT(Var) = VT_DATE;
		res = SystemTimeToVariantTime(Time, &V_DATE(Var));

		delete Time;

		return res >= 0;
		break;

		// Строка
	case VTYPE_PSTR:

		return false;

	case VTYPE_PWSTR:

		V_VT(Var) = VT_BSTR;
		V_BSTR(Var) = SysAllocString(Param->pwstrVal);	

		return true;

		// ДвоичныеДанные 
	case VTYPE_BLOB:

		return false;

	}

	return false;
}

bool CAddInNative::ConvertFromVariantTo1C(VARIANT* Var, tVariant* Param)
{
	if (!Param)
	{
		return true;
	}

	BYTE *raw = NULL;

	switch V_VT(Var)
	{
		// Целое число
	case VT_I4:
		
		TV_VT(Param) = VTYPE_I4;
		Param->lVal = V_I4(Var);
		
		return true;

		// Булево
	case VT_BOOL:
		
		TV_VT(Param) = VTYPE_BOOL;
		TV_BOOL(Param) = V_BOOL(Var);

		return true;

		// Дробное число
	case VT_R4:
		
		TV_VT(Param) = VTYPE_R4;
		Param->dblVal = V_R4(Var);
		
		return true;

	case VT_R8:
		
		TV_VT(Param) = VTYPE_R8;
		Param->dblVal = V_R8(Var);
		
		return true;

		// Дата
	case VT_DATE:
		
		TV_VT(Param) = VTYPE_DATE;
		Param->date = V_DATE(Var);

		return true;

		// Строка
	case VT_BSTR:

		TV_VT(Param) = VTYPE_PWSTR;
		Param->wstrLen = wcslen(V_BSTR(Var));

		if (!m_iMemory->AllocMemory((void**)&Param->pwstrVal, (Param->wstrLen + 1) * sizeof(WCHAR_T)))
		{			
			return false;
		}

		convToShortWchar(&Param->pwstrVal, V_BSTR(Var), Param->wstrLen + 1);
		return true;
						
		
		// ДвоичныеДанные
		// массив байтов
	case VT_ARRAY | VT_UI1:

		TV_VT(Param) = VTYPE_BLOB;

		SafeArrayGetUBound(V_ARRAY(Var), SafeArrayGetDim(V_ARRAY(Var)), (LONG*)&Param->strLen);
		Param->strLen++;

		if (!m_iMemory->AllocMemory((void**)&Param->pstrVal, Param->strLen * sizeof(BYTE)))
		{
			return false;
		}		

		SafeArrayAccessData(V_ARRAY(Var), (void**)&raw);
		memcpy(Param->pstrVal, raw, Param->strLen);
		SafeArrayUnaccessData(V_ARRAY(Var));

		return true;	
	}
	

	return false;
}

bool CAddInNative::InvokeMethod(const long MethodNum, tVariant* Params, const long SizeArray, tVariant* RetValue)
{
	HRESULT hr;

	SAFEARRAY *MethodArgs = SafeArrayCreateVector(VT_VARIANT, 0, SizeArray);
	
	VARIANT RetVal;

	variant_t var;
	VariantInit(&var);

	for (long i = 0; i < SizeArray; i++)
	{
		if (!ConvertFrom1CToVariant(&Params[i], &var))
		{
			SetErrorDescription(L"Ошибка преобразования входящего параметра");
			hr = -1;
			goto Clean;
		}
		
		hr = SafeArrayPutElement(MethodArgs, &i, &var);
		if (FAILED(hr))
		{
			SetErrorDescription();
			goto Clean;
		}
	}

	hr = NETMethods[MethodNum - eLastMethod].info->Invoke_3(NETObject, MethodArgs, &RetVal);

	if (FAILED(hr))
	{
		SetErrorDescription();
		goto Clean;
	}

	
	if (!ConvertFromVariantTo1C(&RetVal, RetValue))
	{
		SetErrorDescription(L"Ошибка преобразования возвращаемого значения");
		hr = -1;
		goto Clean;
	}

Clean:

	if (MethodArgs)
	{
		SafeArrayDestroy(MethodArgs);
		MethodArgs = NULL;
	}

	return SUCCEEDED(hr);
}