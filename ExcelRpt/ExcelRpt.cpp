// ExcelRpt.cpp : ����Ӧ�ó��������Ϊ��
//

#include "stdafx.h"
#include "ExcelRpt.h"
#include "ExcelRptDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CExcelRptApp

BEGIN_MESSAGE_MAP(CExcelRptApp, CWinApp)
	ON_COMMAND(ID_HELP, &CWinApp::OnHelp)
END_MESSAGE_MAP()


// CExcelRptApp ����

CExcelRptApp::CExcelRptApp()
{
	// TODO: �ڴ˴���ӹ�����룬
	// ��������Ҫ�ĳ�ʼ�������� InitInstance ��
}


// Ψһ��һ�� CExcelRptApp ����

CExcelRptApp theApp;


// CExcelRptApp ��ʼ��

BOOL CExcelRptApp::InitInstance()
{
	CWinApp::InitInstance();

    if(!AfxOleInit())
    {
        return FALSE;
    }

	AfxEnableControlContainer();

	// ��׼��ʼ��
	// ���δʹ����Щ���ܲ�ϣ����С
	// ���տ�ִ���ļ��Ĵ�С����Ӧ�Ƴ�����
	// ����Ҫ���ض���ʼ������
	// �������ڴ洢���õ�ע�����
	// TODO: Ӧ�ʵ��޸ĸ��ַ�����
	// �����޸�Ϊ��˾����֯��
	SetRegistryKey(_T("Ӧ�ó��������ɵı���Ӧ�ó���"));

	CExcelRptDlg dlg;
	m_pMainWnd = &dlg;
	INT_PTR nResponse = dlg.DoModal();
	if (nResponse == IDOK)
	{
		// TODO: �ڴ˴����ô����ʱ�á�ȷ�������ر�
		//  �Ի���Ĵ���
	}
	else if (nResponse == IDCANCEL)
	{
		// TODO: �ڴ˷��ô����ʱ�á�ȡ�������ر�
		//  �Ի���Ĵ���
	}

	// ���ڶԻ����ѹرգ����Խ����� FALSE �Ա��˳�Ӧ�ó���
	//  ����������Ӧ�ó������Ϣ�á�
	return FALSE;
}
