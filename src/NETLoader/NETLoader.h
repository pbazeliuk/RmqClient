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

#ifndef __ADDINNATIVE_H__
#define __ADDINNATIVE_H__

#include <wtypes.h>

#include "ComponentBase.h"
#include "AddInDefBase.h"
#include "IMemoryManager.h"



#include <windows.h>

#include <metahost.h>
#pragma comment(lib, "mscoree.lib")

// Import mscorlib.tlb (Microsoft Common Language Runtime Class Library).
#import "mscorlib.tlb"											\
	raw_interfaces_only											\
    high_property_prefixes("_get","_put","_putref")				\
    rename("ReportEvent", "InteropServices_ReportEvent")		\


using namespace mscorlib;


class MethInfo
{
public:
	MethInfo();
	~MethInfo();

	wchar_t* Name = NULL;
	LONG MaxParam = 0;
	bool HasRetVal = false;
	_MethodInfoPtr info = NULL;
};

class PropInfo
{
public:
	PropInfo();
	~PropInfo();

	wchar_t* Name = NULL;
	bool CanRead = false;
	bool CanWrite = false;
	bool IsField = false;
	IUnknown* info = NULL;
};


///////////////////////////////////////////////////////////////////////////////
// class CAddInNative
class CAddInNative : public IComponentBase
{
public:
    enum Props
    {
		eCurrentVersionCLR,
        eLastProp      // Always last
    };

    enum Methods
    {
		eGetLastError,
		eGetInstalledRuntimes,
		eCreateObject,
		eLastMethod      // Always last
    };

    CAddInNative(void);
    virtual ~CAddInNative();
    // IInitDoneBase
    virtual bool ADDIN_API Init(void*);
    virtual bool ADDIN_API setMemManager(void* mem);
	virtual long ADDIN_API GetInfo();
    virtual void ADDIN_API Done();
    // ILanguageExtenderBase
    virtual bool ADDIN_API RegisterExtensionAs(WCHAR_T** wsLanguageExt);
    virtual long ADDIN_API GetNProps();
    virtual long ADDIN_API FindProp(const WCHAR_T* wsPropName);
    virtual const WCHAR_T* ADDIN_API GetPropName(long lPropNum, long lPropAlias);
    virtual bool ADDIN_API GetPropVal(const long lPropNum, tVariant* pvarPropVal);
    virtual bool ADDIN_API SetPropVal(const long lPropNum, tVariant* varPropVal);
    virtual bool ADDIN_API IsPropReadable(const long lPropNum);
    virtual bool ADDIN_API IsPropWritable(const long lPropNum);
    virtual long ADDIN_API GetNMethods();
    virtual long ADDIN_API FindMethod(const WCHAR_T* wsMethodName);
    virtual const WCHAR_T* ADDIN_API GetMethodName(const long lMethodNum, const long lMethodAlias);
    virtual long ADDIN_API GetNParams(const long lMethodNum);
    virtual bool ADDIN_API GetParamDefValue(const long lMethodNum, const long lParamNum, tVariant *pvarParamDefValue);
    virtual bool ADDIN_API HasRetVal(const long lMethodNum);
    virtual bool ADDIN_API CallAsProc(const long lMethodNum, tVariant* paParams, const long lSizeArray);
    virtual bool ADDIN_API CallAsFunc(const long lMethodNum, tVariant* pvarRetValue, tVariant* paParams, const long lSizeArray);
    operator IComponentBase*() { return (IComponentBase*)this; }
    // LocaleBase
    virtual void ADDIN_API SetLocale(const WCHAR_T* loc);
private:
	
	long findName(wchar_t* names[], const wchar_t* name, const uint32_t size) const;
	void SetErrorDescription();
	void SetErrorDescription(wchar_t* desc);
	bool GetLastError(tVariant* pvarRetValue);
	bool GetInstalledRuntimes(tVariant* RetValue);
	void CreateAppDomain(wchar_t* ApplicationBase, wchar_t *VersionCLR);
	bool GetLastRuntime(ICLRMetaHost *MetaHost, ICLRRuntimeInfo **RuntimeInfo);
	SAFEARRAY * GetAssemblyData(wchar_t* AssemblyPath);
	bool CreateObject(tVariant* paParams, const long lSizeArray);
	bool FillMethods();
	bool FillProperties();
	bool ConvertFrom1CToVariant(tVariant* Param, VARIANT* Var);
	bool ConvertFromVariantTo1C(VARIANT* Var, tVariant* Param);
	bool InvokeMethod(const long MethodNum, tVariant* Params, const long SizeArray, tVariant* RetValue);
	
	
	// Attributes
	IMemoryManager *m_iMemory = NULL;
	IAddInDefBase *m_iConnect = NULL;
	
	variant_t NETObject = NULL;
	
	wchar_t* ErrorDescription = NULL;

	ICorRuntimeHost *CorRuntimeHost = NULL;
	_AppDomain *AppDomain = NULL;
	_TypePtr NETType = NULL;
	
	bool NETObjectIsLoad = false;

	MethInfo* NETMethods = NULL;
	long NETMethodsCount = 0;
	
	PropInfo* NETProperties = NULL;
	long NETPropertiesCount = 0;

	wchar_t* CurrentVersionCLR = NULL;
	wchar_t* RemoveAppDirectory = NULL;
};
#endif //__ADDINNATIVE_H__


