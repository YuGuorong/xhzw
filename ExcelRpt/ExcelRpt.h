// ExcelRpt.h : PROJECT_NAME Ӧ�ó������ͷ�ļ�
//

#pragma once

#ifndef __AFXWIN_H__
	#error "�ڰ������ļ�֮ǰ������stdafx.h�������� PCH �ļ�"
#endif

#include "resource.h"		// ������


// CExcelRptApp:
// �йش����ʵ�֣������ ExcelRpt.cpp
//

class CExcelRptApp : public CWinApp
{
public:
	CExcelRptApp();

// ��д
	public:
	virtual BOOL InitInstance();

// ʵ��

	DECLARE_MESSAGE_MAP()
};

extern CExcelRptApp theApp;