// ExcelRptDlg.h : ͷ�ļ�
//

#pragma once


// CExcelRptDlg �Ի���
class CExcelRptDlg : public CDialog
{
// ����
public:
	CExcelRptDlg(CWnd* pParent = NULL);	// ��׼���캯��

// �Ի�������
	enum { IDD = IDD_EXCELRPT_DIALOG };

	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV ֧��


// ʵ��
protected:
	HICON m_hIcon;

	// ���ɵ���Ϣӳ�亯��
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
