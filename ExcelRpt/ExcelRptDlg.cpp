// ExcelRptDlg.cpp : 实现文件
//

#include "stdafx.h"
#include "ExcelRpt.h"
#include "ExcelRptDlg.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#endif


// CExcelRptDlg 对话框


#include "../RptCtrl/Includes/excel8.h"
#include "../RptCtrl/ParserProc.h"

#define SPLID_CHR  _T("$$")//_T("-")//
/////////////////////////////////////////////////////////////////////////////
// CRptAnaCtrl
// Excle write
// http://social.microsoft.com/Forums/zh-CN/visualcpluszhchs/thread/907a8e0c-74e8-4329-930a-a59f982e443d
// http://tech.e800.com.cn/articles/2009/527/1243392027658_1.html
// http://www.cppblog.com/sleepwom/archive/2009/10/03/97804.html

LPCTSTR  g_CT_ajudge_col[]= 
{
    {_T("省份")},{_T("原归属城市")},{_T("原城市区号")},{_T("现归属省份")},{_T("现归属城市")},{_T("现城市区号")}
};
LPCTSTR  g_CT_col[]= 
{
    {_T("省份")},{_T("城市")},{_T("城市区号")}
};
LPCTSTR  g_CM_col[]= 
{    
    {_T("省份")},{_T("城市")},{_T("区号")}
};
LPCTSTR  g_CNN_col[]= 
{
    {_T("省、直辖市")},{_T("所辖城市")},{_T("长途区号")}    
};

LPCTSTR g_strExpColumns[] = 
{
    _T("Link Name"),
    _T("NodeA"),
    _T("NodeB"),
    _T("SourceIP"),
    _T("DestIP"),
    _T("SampleDate"),
    _T("SampleTime"),
    _T("Latency"),
    _T("Loss"),
    _T("Sample Count")
};

PROCESS_TABLE  g_ProcTbl[] = 
{
    { UNINITIALIZE,      &OnInit,       NULL   }, //Got a valid cell
    { PROC_CT_ADJ,       &OnIndexRow,   g_CT_ajudge_col   },
    { PROC_CT_NORMAL,    &OnIndexRow,   g_CT_col   }, //Got a valid line
    { PROC_CM_NORMAL,    &OnIndexRow,   g_CM_col   },
    { PROC_CNN_NORMAL,   &OnIndexRow,   g_CNN_col   },
    { PROC_EXP_DATA,     &ProcExpData,  NULL        }

};

INT FindNextValidColumn(VARIANT * pval, int &col_beging, int col_end)
{
    while(col_beging++< col_end)  //col 0 never used!
    {
        if( pval[col_beging].vt != VT_EMPTY )
            return col_beging;
    }
    return 0;
}

////-------------------------------------------------------------------------------------------------------
TCHAR g_chr_brk[] = {
    _T('\0'), //UNINITIALIZE,      /*File Opened, read CONNECT_NAME*/ 
    _T('\0'), //PROC_CT_ADJ,    
    _T('、'), //PROC_CT_NORMAL, 
    _T(','),  //PROC_CM_NORMAL, 
    _T('、'), //PROC_CNN_NORMAL,
    _T('\0'), //PROC_EXP_DATA,
    _T('\0'), //STATE_END,       
    _T('\0')  //UNCHANGED,         /*Dose not change, for call back return check*/
};

RPOC_STATE OnIndexRow(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    RPOC_STATE state = UNCHANGED;
    int col = 0; 
    if( FindNextValidColumn(pval, col, num) )
    {
        if( pval[col].vt == VT_BSTR)
        {
            CString str = pval[col].bstrVal;
            if(str.Compare(_T("省份")) == 0 )
            {
                if( pParser->cur_state == PROC_CT_ADJ )
                {
                    col += 3;
                }
                else if( pParser->cur_state == UNINITIALIZE )
                {
                    pParser->cur_state = PROC_CM_NORMAL;
                }
            }
            else if(str.Compare(_T("省、直辖市")) == 0 )
            {
                if( pParser->cur_state == UNINITIALIZE )
                {
                    pParser->cur_state = PROC_CNN_NORMAL;
                }
            }
            else
            {
                return STATE_END;
            }

            pParser->ColIndex[0] = col++;
            pParser->ColIndex[1] = col++;
            pParser->ColIndex[2] = col++;
            //col = 6;
            pParser->ColNumberStart = col;
            while(col <= num )
            {
                str = pval[col++].bstrVal;
                str = str.Left(4);
                pParser->strInfo.Add(str);
            }
            pParser->app_state = pParser->cur_state;
            pParser->chr_brk = g_chr_brk[pParser->app_state];
            state = PROC_EXP_DATA;
        }
    }
    return state;
}

RPOC_STATE OnInit(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    int col = 0; 
    if( FindNextValidColumn(pval, col, num) )
    {
        if( pval[col].vt == VT_BSTR)
        {
            pParser->strInfo.RemoveAll();
            CString str = pval[col].bstrVal;
            if( str.Find( _T("调整")) != -1)
            {
                return PROC_CT_ADJ;
            }
            else if( (str.Find( _T("汇总") )!= -1 ) ||  ( str.Find( _T("新增")) != -1 ) )
            {
                return PROC_CT_NORMAL;
            }
            else 
                return OnIndexRow(pval, num , pstrKeyWords, pParser);           
        }
        return STATE_END;
    }
    return UNCHANGED;
}
//---------------------------------------------
int GetPartStr(CString &strCell, CString &part, RPT_PARSER * pParser)
{
    //TCHAR chrbrk = 
    int brk = strCell.Find(pParser->chr_brk);
    if( brk != -1 )
    {
        part = strCell.Left(brk);
        strCell.Delete(0, brk+1);
        return 1;
    }
    return 0;
}

int ProcParseCell( VARIANT * pval, int col, RPT_PARSER * pParser)
{
    if( pval[col].vt == VT_BSTR)
    {
        CString strpart;
        CString str = //_T("280-289、301、311、321、331、650-653、655-656");//cnn //pval[col].bstrVal;
        _T("402-403,406-407,472-473,477,484,487");//cm
        //_T("311-312、369、383-385、409、834");//ct
        while( GetPartStr(str, strpart, pParser) )
        {
            int begin, end = -1;
            _stscanf(str,_T("%d-%d"),&begin, &end);
            if( end == -1) end = begin;
            while( begin <= end)
            {
                
            }
        }
    }    
    return 1;
}
RPOC_STATE ProcExpData(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    int col = 0;
    for( int idx = 0; idx<3; idx++)
    {
        pParser->ptrInfos[idx] = _T("");
        col = pParser->ColIndex[idx];
        if( pval[col].vt == VT_BSTR )
        {
            pParser->ptrInfos[idx] = pval[col].bstrVal;
        }
    }

    for(col = pParser->ColNumberStart; col < num ; col++)
    {
        ProcParseCell(pval, col, pParser);
    }
    return UNCHANGED;
}

////-------------------------------------------------------------------------------------------------------

////-------------------------------------------------------------------------------------------------------

int OnReadRow(VARIANT * pval, int num, int repeat, void * param)
{
    RPT_PARSER * pParser = (RPT_PARSER * )param;
    if( pParser->cur_state >= UNCHANGED )
        return 0;

    int state_id = (int)(pParser->cur_state);

    PROCESS_TABLE * pTblProc = &g_ProcTbl[state_id];

    if( pTblProc->pProc )
    {
        RPOC_STATE next_statu = pTblProc->pProc(pval, num, pTblProc->pszKeyWords, pParser);
        if( next_statu != UNCHANGED )
            pParser->cur_state = next_statu;
        if( next_statu == STATE_END )
            return 0;
    }
    pParser->nCurReadRow++;

    return 1;
}


BOOL GetRangString(TCHAR * buff, int len, int cols, int rows)
{
    int pos = 1;
    if( cols > 26 * 26 ) return FALSE;
    if( cols > 26 )
    {
        buff[0] = _T('A') + cols / 26;
        buff[1] = _T('A') + cols % 26;
        pos = 2;
    }
    else 
    {
        buff[0] = _T('A') + cols;
    }

    _stprintf(&buff[pos], _T("%d"), rows);
    return TRUE;
}

void ParseNodeByFileName(LPCTSTR szExcleFileName,RPT_PARSER * pParser)
{
    CString str = szExcleFileName;
    int pos=str.ReverseFind(_T('\\'));
    if( pos != -1) str.Delete(0, pos+1);

    pos=str.Find(SPLID_CHR, 0);
    pParser->strInfo[NodeA] = _T("");
    pParser->strInfo[NodeB] = _T("");
    if( pos != -1)
    {
        CString strNode = str.Left(pos);
        pParser->strInfo[NodeA] = strNode;
        str.Delete(0, pos+2);
        pos = str.Find(_T('.'));
        if( pos != -1)
        {
            strNode = str.Left(pos);
            pParser->strInfo[NodeB] = strNode;
        }
    }
}


BOOL ExcleRead( RPT_PARSER * pParser , LPCTSTR szExcleFileName, cbExcelRead pfnRead)
{
    //   OLE   Variant   for   Optional. 
    BOOL ret_val = TRUE;

    COleVariant   VOptional((long)DISP_E_PARAMNOTFOUND,   VT_ERROR); 

    Range   objRange; 
    VARIANT   ret; 
    INT row = pParser->nCurWriteRow;

  
    try 
    {        
        if( _taccess(szExcleFileName,4) != 0 )
            return FALSE;

        //   Instantiate   Excel   and   open   an   existing   workbook. 
        pParser->objBook   =   pParser->objBooks.Open( szExcleFileName, 
            VOptional,   VOptional,   VOptional,   VOptional, 
            VOptional,   VOptional,   VOptional,   VOptional, 
            VOptional,   VOptional,   VOptional,   VOptional);
        
        pParser->objSheets   =   pParser->objBook.GetWorksheets(); 
        pParser->objSheet   =   pParser->objSheets.GetItem(COleVariant((short)1)); 
        
        //Get   the   range   object   for   which   you   wish   to   retrieve   the 
        //data   and   then   retrieve   the   data   (as   a   variant   array,   ret). 
        int col_num = GetColumnCount(pParser->objSheet);
        int rows  =  GetRowCount(pParser->objSheet);


        TCHAR strend[8];
        if( GetRangString(strend, 8, col_num, rows) )
        {
            objRange   =   pParser->objSheet.GetRange(COleVariant( _T("A1 ")),   COleVariant( strend )); 
            ret   =   objRange.GetValue(); 
            
            //Create   the   SAFEARRAY   from   the   VARIANT   ret. 
            COleSafeArray   sa(ret); 
            
            //Determine   the   array 's   dimensions. 
            long   lNumRows; 
            long   lNumCols; 
            sa.GetUBound(1,   &lNumRows); 
            sa.GetUBound(2,   &lNumCols); 
            
            //Display   the   elements   in   the   SAFEARRAY. 
            long   index[2]; 
            VARIANT *  val = new VARIANT[lNumCols+1]; 
            int   r =1,   c,  rep =0; 
            
            pParser->cur_state = UNINITIALIZE;
            pParser->nToalCols = col_num;
            pParser->nToalRows = rows;
            pParser->nCurReadRow =1;

            pParser->strInfo.RemoveAll();
            pParser->strInfo.SetSize(MAX_COLLECTIONS);
            ParseNodeByFileName(szExcleFileName,pParser);


            while(r <=lNumRows) 
            { 
                for(c=1;c <=lNumCols;c++) 
                { 
                    index[0]=r; 
                    index[1]=c; 
                    sa.GetElement(index,   &val[c]); 
                } 
                int flag = pfnRead(val, col_num, rep, pParser);
                if( flag == 0 )
                    break;
                r++;
            } 
            delete val;
        }
        //Close   the   workbook   without   saving   changes 
        //and   quit   Microsoft   Excel. 
        pParser->objBook.Close(COleVariant((short)FALSE),   VOptional,   VOptional); 
    }
    catch(CException * e)
    {
        ret_val =  FALSE;
    }

    if( row != pParser->nCurWriteRow )
        InsertRowData(pParser, NULL, 0);

    return ret_val;
}


BOOL InitExcelApp( RPT_PARSER * pParser )
{
    pParser->objApp.CreateDispatch(_T("Excel.Application"));
    if (!pParser->objApp)
    {
        AfxMessageBox(_T("Can not start Excel."));
        return FALSE;
    }
    
    //   Instantiate   Excel   and   open   an   existing   workbook. 
    pParser->objApp.SetVisible(FALSE);
    pParser->objApp.SetUserControl(TRUE);
    pParser->objBooks = pParser->objApp.GetWorkbooks(); 


    pParser->cur_state = UNINITIALIZE;

    pParser->nExpCols = sizeof(g_strExpColumns)/sizeof(LPCTSTR);
    pParser->nCurWriteRow = 0;

    return TRUE;

}

void ParseExcle(LPCTSTR szExcleFileName, LPCTSTR szDir, CString &strExp)
{
    if(szExcleFileName == NULL )
        return;

    RPT_PARSER  parser;
    COleVariant   VOptional((long)DISP_E_PARAMNOTFOUND,   VT_ERROR); 

    InitExcelApp(&parser);
    if( szDir == NULL )
    {
        CString strdir = szExcleFileName;
        int pos = strdir.ReverseFind(_T('\\'));
        if( pos == -1 )
        {
            parser.strRptDir = _T(".\\");
        }
        else
        {
            parser.strRptDir  = strdir.Left(pos + 1);
        }
    }
    else
    {
        parser.strRptDir = szDir;
    }



    if( ExcleRead(&parser,szExcleFileName, OnReadRow) )
    {
        if( parser.nCurWriteRow )
        {
            
            Range range;
            Range cols;
            TCHAR buff[8]; 
            GetRangString(buff, 8, parser.nExpCols , 2);
            range=parser.ExpSheet.GetRange(COleVariant(_T("A1")),COleVariant(buff));
            cols=range.GetEntireColumn();
            cols.AutoFit();            
            
            ExportDataDone(&parser);
        }
    }
    parser.objApp.Quit(); 
    
}

CString OnBrowserDir()
{
    // TODO: Add your control notification handler code here
    CString strdir;
    TCHAR szDir[MAX_PATH];
    ZeroMemory(szDir, MAX_PATH);
    BROWSEINFO bi;
    ZeroMemory(&bi, sizeof(BROWSEINFO));
    ITEMIDLIST *pidl;
    bi.hwndOwner = ::AfxGetMainWnd()->GetSafeHwnd();
    bi.pidlRoot = NULL;
    bi.pszDisplayName = szDir;
    bi.lpszTitle = _T("请选择目录");
    bi.ulFlags = BIF_STATUSTEXT | BIF_RETURNONLYFSDIRS;
    bi.lpfn = NULL;
    bi.lParam = 0;
    bi.iImage = 0;
    pidl = SHBrowseForFolder(&bi);
    if(pidl == NULL)  return strdir;
    if(SHGetPathFromIDList(pidl, szDir))   
        strdir = szDir;
    return strdir;
}


CExcelRptDlg::CExcelRptDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CExcelRptDlg::IDD, pParent)
    , szFolderName(_T(""))
    , m_expFIleName(_T(""))
{
	m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
}

void CExcelRptDlg::DoDataExchange(CDataExchange* pDX)
{
    CDialog::DoDataExchange(pDX);
    DDX_Text(pDX, IDC_TEXT_OUTPUT, m_expFIleName);
    DDX_Text(pDX, IDC_EDIT1, szFolderName);
}

BEGIN_MESSAGE_MAP(CExcelRptDlg, CDialog)
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
    ON_BN_CLICKED(IDC_BTN_OPENDIR, &CExcelRptDlg::OnBnClickedBtnOpendir)
END_MESSAGE_MAP()


// CExcelRptDlg 消息处理程序

BOOL CExcelRptDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// 设置此对话框的图标。当应用程序主窗口不是对话框时，框架将自动
	//  执行此操作
	SetIcon(m_hIcon, TRUE);			// 设置大图标
	SetIcon(m_hIcon, FALSE);		// 设置小图标

	// TODO: 在此添加额外的初始化代码

	return TRUE;  // 除非将焦点设置到控件，否则返回 TRUE
}

// 如果向对话框添加最小化按钮，则需要下面的代码
//  来绘制该图标。对于使用文档/视图模型的 MFC 应用程序，
//  这将由框架自动完成。

void CExcelRptDlg::OnPaint()
{
	if (IsIconic())
	{
		CPaintDC dc(this); // 用于绘制的设备上下文

		SendMessage(WM_ICONERASEBKGND, reinterpret_cast<WPARAM>(dc.GetSafeHdc()), 0);

		// 使图标在工作矩形中居中
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// 绘制图标
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

//当用户拖动最小化窗口时系统调用此函数取得光标显示。
//
HCURSOR CExcelRptDlg::OnQueryDragIcon()
{
	return static_cast<HCURSOR>(m_hIcon);
}

#undef S_OK
#define S_OK  
extern  int g_nExpFileType ;

void CExcelRptDlg::OnBnClickedBtnOpendir()
{
    this->UpdateData();

   CString strfilt = szFolderName;
    if( strfilt.IsEmpty() )
        strfilt = OnBrowserDir();
    else 
    {
        CString str= strfilt.Right(4);
        if( str.CompareNoCase(_T(".xls")) == 0 )
        {
            CString strexp;
            ParseExcle(strfilt,NULL,strexp );
            m_strExpFile = strexp;
            return S_OK;
        }        
    }

    if( strfilt.IsEmpty() )
    {
        return S_OK;
    }

    if( strfilt.GetAt(strfilt.GetLength() -1 ) == _T('\\') ) 
        strfilt.Delete(strfilt.GetLength() -1, 1);

    CString strDir = strfilt;


    DWORD tick = ::GetTickCount();
    strfilt += _T("\\*.xls"); 
    WIN32_FIND_DATA  find;
    HANDLE  hfile = FindFirstFile(strfilt,&find);
    if( hfile != INVALID_HANDLE_VALUE  )
    {
        RPT_PARSER  parser;
        COleVariant   VOptional((long)DISP_E_PARAMNOTFOUND,   VT_ERROR); 
        
        InitExcelApp(&parser);
        
        BOOL bnext = TRUE;
        while( bnext )
        {
            CString strXls = strDir + _T("\\") + find.cFileName;
            parser.strRptDir  = strDir;
            TRACE(_T("%ws\n"), strXls);//这里是所有找到的文件名
            ExcleRead(&parser,strXls, OnReadRow);
            bnext = FindNextFile(hfile, &find);        
        }

        if( parser.nCurWriteRow )
        {
            if( g_nExpFileType == 0 )
            {
                Range range;
                Range cols;
                TCHAR buff[8]; 
                GetRangString(buff, 8, parser.nExpCols , 2);
                range=parser.ExpSheet.GetRange(COleVariant(_T("A1")),COleVariant(buff));
                cols=range.GetEntireColumn();
                cols.AutoFit();
            }
            ExportDataDone(&parser);            
            m_expFIleName = parser.strExpFileName;
            szFolderName = m_expFIleName.Left(m_expFIleName.ReverseFind(_T('\\')));

        }
        
        parser.objApp.Quit(); 
    }
    
    TRACE(_T("Time used: %d\n"), GetTickCount()-tick );
    g_nExpFileType = 1;
    this->UpdateData(false);
}
