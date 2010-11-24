#pragma once

// RptCtrl.h : main header file for RptCtrl.DLL

#if !defined( __AFXCTL_H__ )
#error "include 'afxctl.h' before including this file"
#endif

#include "resource.h"       // main symbols

      #include "comcat.h"

      // Helper function to create a component category and associated
      // description
      HRESULT CreateComponentCategory(CATID catid, WCHAR* catDescription);

      // Helper function to register a CLSID as belonging to a component
      // category
      HRESULT RegisterCLSIDInCategory(REFCLSID clsid, CATID catid);
		
// CRptCtrlApp : See RptCtrl.cpp for implementation.

class CRptCtrlApp : public COleControlModule
{
public:
	BOOL InitInstance();
	int ExitInstance();
};

extern const GUID CDECL _tlid;
extern const WORD _wVerMajor;
extern const WORD _wVerMinor;

