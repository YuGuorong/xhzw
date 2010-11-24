// RptCtrlPropPage.cpp : Implementation of the CRptCtrlPropPage property page class.

#include "stdafx.h"
#include "RptCtrl.h"
#include "RptCtrlPropPage.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


IMPLEMENT_DYNCREATE(CRptCtrlPropPage, COlePropertyPage)



// Message map

BEGIN_MESSAGE_MAP(CRptCtrlPropPage, COlePropertyPage)
END_MESSAGE_MAP()



// Initialize class factory and guid

IMPLEMENT_OLECREATE_EX(CRptCtrlPropPage, "RPTCTRL.RptCtrlPropPage.1",
	0x4bf0caca, 0xc3eb, 0x4c42, 0xad, 0xf4, 0xad, 0x53, 0xa4, 0x7b, 0x4, 0xa5)



// CRptCtrlPropPage::CRptCtrlPropPageFactory::UpdateRegistry -
// Adds or removes system registry entries for CRptCtrlPropPage

BOOL CRptCtrlPropPage::CRptCtrlPropPageFactory::UpdateRegistry(BOOL bRegister)
{
	if (bRegister)
		return AfxOleRegisterPropertyPageClass(AfxGetInstanceHandle(),
			m_clsid, IDS_RPTCTRL_PPG);
	else
		return AfxOleUnregisterClass(m_clsid, NULL);
}



// CRptCtrlPropPage::CRptCtrlPropPage - Constructor

CRptCtrlPropPage::CRptCtrlPropPage() :
	COlePropertyPage(IDD, IDS_RPTCTRL_PPG_CAPTION)
{
}



// CRptCtrlPropPage::DoDataExchange - Moves data between page and properties

void CRptCtrlPropPage::DoDataExchange(CDataExchange* pDX)
{
	DDP_PostProcessing(pDX);
}



// CRptCtrlPropPage message handlers
