// ExcelRptDlg.h : 头文件
//

#pragma once


// CExcelRptDlg 对话框
class CExcelRptDlg : public CDialog
{
// 构造
public:
	CExcelRptDlg(CWnd* pParent = NULL);	// 标准构造函数

// 对话框数据
	enum { IDD = IDD_EXCELRPT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV 支持


// 实现
protected:
	HICON m_hIcon;

	// 生成的消息映射函数
	virtual BOOL OnInitDialog();
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	DECLARE_MESSAGE_MAP()
public:
    afx_msg void OnBnClickedBtnOpendir();
public:
    CString szFolderName;
    CString m_strExpFile;
public:
    CString m_expFIleName;
};
