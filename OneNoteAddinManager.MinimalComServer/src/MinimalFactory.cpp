#include "MinimalFactory.h"
#include "MinimalComObject.h"
#include "utils.h"
#include <oaidl.h>
#include <sstream>

// MinimalFactory implementation

/**
 * @brief QueryInterface implementation for MinimalFactory.
 *
 * Called by COM clients to request a pointer to a supported interface on
 * the class factory. Logs the requested IID for diagnostics. Returns a
 * pointer to IUnknown or IClassFactory if requested, otherwise returns
 * E_NOINTERFACE. If 'ppv' is null, returns E_POINTER.
 *
 * @param riid The IID of the requested interface (e.g., IID_IUnknown,
 * IID_IClassFactory).
 * @param ppv Address of pointer variable to receive the interface pointer.
 * @return HRESULT S_OK if the interface is supported, E_NOINTERFACE or
 * E_POINTER otherwise.
 */
STDMETHODIMP MinimalFactory::QueryInterface(REFIID riid, void** ppv) {
    // Log QueryInterface call and requested IID
    std::wstringstream msg;
    msg << L"MinimalFactory::QueryInterface called. IID: ";
    LPOLESTR iidStr = nullptr;
    StringFromIID(riid, &iidStr);
    if (iidStr) {
        msg << iidStr;
        CoTaskMemFree(iidStr);
    } else {
        msg << L"(IID conversion failed)";
    }
    Log(msg.str().c_str());

    if (!ppv) return E_POINTER;
    if (riid == IID_IUnknown || riid == IID_IClassFactory) {
        *ppv = static_cast<IClassFactory*>(this);
        AddRef();
        return S_OK;
    }
    *ppv = nullptr;
    return E_NOINTERFACE;
}

STDMETHODIMP_(ULONG) MinimalFactory::AddRef() { 
    return InterlockedIncrement(&refCount); 
}

STDMETHODIMP_(ULONG) MinimalFactory::Release() {
    ULONG res = InterlockedDecrement(&refCount);
    if (res == 0) delete this;
    return res;
}

/**
 * @brief Create an instance of the COM object.
 * 
 * CreateInstance is called by COM to create an instance of the object.
 * The 'riid' parameter specifies the IID (Interface Identifier) of the
 * interface the caller wants to use. For example, it may be:
 *
 *   - IID_IUnknown:     {00000000-0000-0000-C000-000000000046} (base COM interface)
 *   - IID_IDispatch:    {00020400-0000-0000-C000-000000000046} (late binding/scripting)
 *   - IID_IClassFactory:{00000001-0000-0000-C000-000000000046} (class factory)
 *   - IID_IMarshal:     {00000003-0000-0000-C000-000000000046} (marshalling)
 *   - IID_IPersist:     {0000010C-0000-0000-C000-000000000046} (persistence)
 *   - Or a custom interface IID.
 *
 * The object should return a pointer to the requested interface if
 * supported, or E_NOINTERFACE otherwise. The CLSID determines which object
 * to instantiate; the IID determines which interface to return.
 */
STDMETHODIMP MinimalFactory::CreateInstance(IUnknown* pUnkOuter, REFIID riid, void** ppv) {
    if (pUnkOuter) return CLASS_E_NOAGGREGATION;
    MinimalComObject* obj = new MinimalComObject();
    HRESULT hr = obj->QueryInterface(riid, ppv);
    obj->Release();

    // Log instance creation.
    std::wstringstream msg;
    msg << L"CreateInstance called. IID: ";
    LPOLESTR iidStr = nullptr;
    StringFromIID(riid, &iidStr);
    if (iidStr) {
        msg << iidStr;
        CoTaskMemFree(iidStr);
    } else {
        msg << L"(IID conversion failed)";
    }
    Log(msg.str().c_str());

    return hr;
}

STDMETHODIMP MinimalFactory::LockServer(BOOL) { 
    return S_OK; 
}