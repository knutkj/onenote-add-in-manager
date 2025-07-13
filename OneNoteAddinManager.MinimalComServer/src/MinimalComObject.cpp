#include "MinimalComObject.h"

// IUnknown implementation

STDMETHODIMP MinimalComObject::QueryInterface(REFIID riid, void** ppv) {
    if (!ppv) return E_POINTER;
    if (riid == IID_IUnknown || riid == IID_IDispatch) {
        *ppv = static_cast<IDispatch*>(this);
        AddRef();
        return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) MinimalComObject::AddRef() { 
    return InterlockedIncrement(&refCount); 
}

STDMETHODIMP_(ULONG) MinimalComObject::Release() {
    ULONG res = InterlockedDecrement(&refCount);
    if (res == 0) delete this;
    return res;
}

// IDispatch implementation

STDMETHODIMP MinimalComObject::GetTypeInfoCount(UINT* pctinfo) { 
    if (pctinfo) *pctinfo = 0; 
    return S_OK; 
}

STDMETHODIMP MinimalComObject::GetTypeInfo(UINT, LCID, ITypeInfo**) { 
    return E_NOTIMPL; 
}

STDMETHODIMP MinimalComObject::GetIDsOfNames(REFIID, LPOLESTR*, UINT, LCID, DISPID*) { 
    return DISP_E_UNKNOWNNAME; 
}

STDMETHODIMP MinimalComObject::Invoke(DISPID, REFIID, LCID, WORD, DISPPARAMS*, VARIANT*, EXCEPINFO*, UINT*) { 
    return S_OK; 
}