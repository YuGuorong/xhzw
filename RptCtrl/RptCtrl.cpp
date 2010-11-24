// RptCtrl.cpp : Implementation of CRptCtrlApp and DLL registration.

#include "stdafx.h"
#include "RptCtrl.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


CRptCtrlApp theApp;

const GUID CDECL BASED_CODE _tlid =
		{ 0x75914D81, 0xEAB3, 0x49B4, { 0xB4, 0xD2, 0xA, 0xD6, 0x8, 0x58, 0x9F, 0xB9 } };
const WORD _wVerMajor = 1;
const WORD _wVerMinor = 0;

      #include "comcat.h"

      // Helper function to create a component category and associated
      // description
      HRESULT CreateComponentCategory(CATID catid, WCHAR* catDescription)
      {
         ICatRegister* pcr = NULL ;
         HRESULT hr = S_OK ;

         hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr,
                               NULL,
                               CLSCTX_INPROC_SERVER,
                               IID_ICatRegister,
                               (void**)&pcr);
         if (FAILED(hr))
            return hr;

         // Make sure the HKCR\Component Categories\{..catid...}
         // key is registered
         CATEGORYINFO catinfo;
         catinfo.catid = catid;
         catinfo.lcid = 0x0409 ; // english

         // Make sure the provided description is not too long.
         // Only copy the first 127 characters if it is
         int len = wcslen(catDescription);
         if (len>127)
            len = 127;
         wcsncpy(catinfo.szDescription, catDescription, len);
         // Make sure the description is null terminated
         catinfo.szDescription[len] = '\0';

         hr = pcr->RegisterCategories(1, &catinfo);
         pcr->Release();

         return hr;
      }

      // Helper function to register a CLSID as belonging to a component
      // category
      HRESULT RegisterCLSIDInCategory(REFCLSID clsid, CATID catid)
      {
         // Register your component categories information.
         ICatRegister* pcr = NULL ;
         HRESULT hr = S_OK ;
         hr = CoCreateInstance(CLSID_StdComponentCategoriesMgr,
                               NULL,
                               CLSCTX_INPROC_SERVER,
                               IID_ICatRegister,
                               (void**)&pcr);
         if (SUCCEEDED(hr))
         {
            // Register this category as being "implemented" by
            // the class.
            CATID rgcatid[1] ;
            rgcatid[0] = catid;
            hr = pcr->RegisterClassImplCategories(clsid, 1, rgcatid);
         }

         if (pcr != NULL)
            pcr->Release();

         return hr;
      }

      const CATID CATID_SafeForScripting     =
      {0x7dd95801,0x9882,0x11cf,{0x9f,0xa9,0x00,0xaa,0x00,0x6c,0x42,0xc4}};
      const CATID CATID_SafeForInitializing  =
      {0x7dd95802,0x9882,0x11cf,{0x9f,0xa9,0x00,0xaa,0x00,0x6c,0x42,0xc4}};
			
     const GUID CDECL BASED_CODE _ctlid =
      //{ 0x43bd9e45, 0x328f, 0x11d0,{ 0xa6, 0xb9, 0x0, 0xaa, 0x0, 0xa7, 0xf, 0xc2 } };
     { 0xa7116aa5, 0x99df, 0x4310, {0x88, 0x40, 0xd0, 0x28, 0x69, 0x14, 0x9b, 0x7d}};
// CRptCtrlApp::InitInstance - DLL initialization

BOOL CRptCtrlApp::InitInstance()
{
	BOOL bInit = COleControlModule::InitInstance();

	if (bInit)
	{
		// TODO: Add your own module initialization code here.
	}

	return bInit;
}



// CRptCtrlApp::ExitInstance - DLL termination

int CRptCtrlApp::ExitInstance()
{
	// TODO: Add your own module termination code here.

	return COleControlModule::ExitInstance();
}



// DllRegisterServer - Adds entries to the system registry
      STDAPI DllRegisterServer(void)
      {
          AFX_MANAGE_STATE(_afxModuleAddrThis);

          if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
              return ResultFromScode(SELFREG_E_TYPELIB);

          if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
              return ResultFromScode(SELFREG_E_CLASS);

          if (FAILED( CreateComponentCategory(
                  CATID_SafeForScripting,
                  L"Controls that are safely scriptable") ))
                return ResultFromScode(SELFREG_E_CLASS);

          if (FAILED( CreateComponentCategory(
                  CATID_SafeForInitializing,
                  L"Controls safely initializable from persistent data") ))
                return ResultFromScode(SELFREG_E_CLASS);

          if (FAILED( RegisterCLSIDInCategory(
                  _ctlid, CATID_SafeForScripting) ))
                return ResultFromScode(SELFREG_E_CLASS);

          if (FAILED( RegisterCLSIDInCategory(
                  _ctlid, CATID_SafeForInitializing) ))
                return ResultFromScode(SELFREG_E_CLASS);

          return NOERROR;
      }
	
STDAPI DllRegisterServerx(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleRegisterTypeLib(AfxGetInstanceHandle(), _tlid))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(TRUE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}



// DllUnregisterServer - Removes entries from the system registry

STDAPI DllUnregisterServer(void)
{
	AFX_MANAGE_STATE(_afxModuleAddrThis);

	if (!AfxOleUnregisterTypeLib(_tlid, _wVerMajor, _wVerMinor))
		return ResultFromScode(SELFREG_E_TYPELIB);

	if (!COleObjectFactoryEx::UpdateRegistryAll(FALSE))
		return ResultFromScode(SELFREG_E_CLASS);

	return NOERROR;
}
