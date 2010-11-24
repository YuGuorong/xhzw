#pragma once

// RptCtrlCtrl.h : Declaration of the CRptCtrlCtrl ActiveX Control class.


// CRptCtrlCtrl : See RptCtrlCtrl.cpp for implementation.

class CRptCtrlCtrl : public COleControl
{
	DECLARE_DYNCREATE(CRptCtrlCtrl)

// Constructor
public:
	CRptCtrlCtrl();

// Overrides
public:
	virtual void OnDraw(CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid);
	virtual void DoPropExchange(CPropExchange* pPX);
	virtual void OnResetState();

// Implementation
protected:
	~CRptCtrlCtrl();

	DECLARE_OLECREATE_EX(CRptCtrlCtrl)    // Class factory and guid
	DECLARE_OLETYPELIB(CRptCtrlCtrl)      // GetTypeInfo
	DECLARE_PROPPAGEIDS(CRptCtrlCtrl)     // Property page IDs
	DECLARE_OLECTLTYPE(CRptCtrlCtrl)		// Type name and misc status

// Message maps
	DECLARE_MESSAGE_MAP()

// Dispatch maps
	DECLARE_DISPATCH_MAP()

	afx_msg void AboutBox();

// Event maps
	DECLARE_EVENT_MAP()

// Dispatch and event IDs
public:
	enum {
        dispidStrExportFileName = 4,
        dispidstrExpFile = 3,
        dispidAnalyzeFolder = 2L,
        dispidAnalyzeExcleFile = 1L
    };
protected:
    LONG AnalyzeExcleFile(LPCTSTR szExcleFileName);
    LONG AnalyzeFolder(LPCTSTR szFolderName);
    void OnstrExpFileChanged(void);
    CString m_strExpFile;
    void OnStrExportFileNameChanged(void);
    CString m_StrExportFileName;
};

