#pragma once

// RptCtrlPropPage.h : Declaration of the CRptCtrlPropPage property page class.


// CRptCtrlPropPage : See RptCtrlPropPage.cpp for implementation.

class CRptCtrlPropPage : public COlePropertyPage
{
	DECLARE_DYNCREATE(CRptCtrlPropPage)
	DECLARE_OLECREATE_EX(CRptCtrlPropPage)

// Constructor
public:
	CRptCtrlPropPage();

// Dialog Data
	enum { IDD = IDD_PROPPAGE_RPTCTRL };

// Implementation
protected:
	virtual void DoDataExchange(CDataExchange* pDX);    // DDX/DDV support

// Message maps
protected:
	DECLARE_MESSAGE_MAP()
};

