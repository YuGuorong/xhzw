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

COL_FLAG g_strColFlags[] = 
{
    {SAMPLE_TIME,      _T("Sample Time")        },
    {LATENCY,          _T("Latency in ms")      },
    {PACKET_LOSS,      _T("Packet Loss (%)")    },
    {SAMPLE_COUNT,     _T("Sample Count")       },
};

COL_FLAG  g_CollectIP[]= 
{
    {IP_SOURCE,        _T("Source Address:")},
    {IP_DEST,          _T("Destination Address:")},
};

LPCTSTR g_strHeaderStart[] = { _T("Report Details"),   NULL };
LPCTSTR g_strHeaderEnd[]   = { _T("Sample Count")  ,   NULL };
LPCTSTR g_strTableEnd[]    = { _T("Report Generated"), _T("Total Samples"),NULL };

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
    { UNINITIALIZE,      &OnInit,          NULL               }, //Got a valid cell
    { LOOKUP_IPS,        &OnLookupIPs,     g_strHeaderStart   },
    { LOOKUP_TBL_HEADER, &OnLookupHead,    g_strHeaderEnd     }, //Got a valid line
    { SET_COL_INDEX,     &OnFirstDataRow,  g_strTableEnd      },
    { LOOKUP_TBL_DATA,   &OnExpDatas,      g_strTableEnd      },
    { LOOKUP_NEXT_TBL,   &OnLookupNextTbl, g_strHeaderStart   },
    { SKIP_TBLE_HEADER,  &OnSkipHead,      g_strHeaderEnd     },
};

INT CheckEndConditon(LPCTSTR pstrKeyWords[], CString &str)
{
    if( pstrKeyWords == NULL ) return 0;
    int i=0;
    str.TrimLeft(_T(' '));
    str.TrimRight(_T(' '));

    while( pstrKeyWords[i] )
    {
        if( str.Compare(pstrKeyWords[i]) == 0 )
            return 1;
        i++;
    }
    return 0;
}

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
RPOC_STATE OnInit(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    int col = 0; 
    if( FindNextValidColumn(pval, col, num) )
    {
        if( pval[col].vt == VT_BSTR)
        {
            pParser->strInfo.SetAt(CONNECT_NAME,  pval[col].bstrVal);
        }
        return LOOKUP_IPS;
    }
    return UNCHANGED;
}
//---------------------------------------------
RPOC_STATE OnLookupIPs(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    int col = 0 , nfindIp = -1; 
    while( FindNextValidColumn(pval, col, num) )
    {
        if( pval[col].vt == VT_BSTR )
        {
            CString str = pval[col].bstrVal;
            if( nfindIp >= 0 )
            {
                pParser->strInfo.SetAt(nfindIp, str);
            }

            if( CheckEndConditon(pstrKeyWords, str) )
                return LOOKUP_TBL_HEADER;
            for(int j=0; j<sizeof(g_CollectIP)/sizeof(COL_FLAG);j++)
            {
                if( str.Compare(g_CollectIP[j].strFlag) == 0 )
                {
                    nfindIp = g_CollectIP[j].index;
                    
                }                
            }
        }
    }
    return UNCHANGED;
}
//---------------------------------------------

RPOC_STATE OnLookupHead(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    int col =0 ;
    while( FindNextValidColumn(pval, col, num) )
    {
        if( pval[col].vt == VT_BSTR )
        {
            CString str = pval[col].bstrVal;
            if( CheckEndConditon(pstrKeyWords, str) )
            {
                //TODO: Export header here
                if( pParser->nCurWriteRow == 0 )
                {
                    PrepareExportFile(pParser,
                        g_strExpColumns, pParser->nExpCols);
                }
                TRACE(_T("Find Table, %d\n"), pParser->nCurReadRow);
                return SET_COL_INDEX;
            }
        }
    }
    return UNCHANGED;
}
//---------------------------------------------
RPOC_STATE OnFirstDataRow(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    int col =0 , idx = 0;
    while( FindNextValidColumn(pval, col, num) )
    {
        pParser->ColIndex[idx++] = col; 

    }
    if( idx )
    {
        pParser->nImportCols = idx;
        OnExpDatas(pval, num, pstrKeyWords,  pParser);
        TRACE(_T("first row: %d\n"),pParser->nCurReadRow);
        return LOOKUP_TBL_DATA;
    }
    return UNCHANGED;
}

//---------------------------------------------
RPOC_STATE OnExpDatas(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    if( pParser->nImportCols )
    {
        if( pval[pParser->ColIndex[0]].vt == VT_EMPTY )
        {
            int col = 0; 
            while( FindNextValidColumn(pval, col, num) )
            {
                if( pval[col].vt == VT_BSTR )
                {
                    TRACE(_T("table end %d\n"),pParser->nCurReadRow);
                    return LOOKUP_NEXT_TBL;
                }
            }
        }
        else
        {
            //TODO: Export data here
            //INT InsertRowData(RPT_PARSER *pParser, LPCTSTR szCol[], int nCols);
            CStringArray str;
            str.Append(pParser->strInfo);
            int date_col = pParser->ColIndex[0];
            if(pval[date_col].vt == VT_BSTR)
            {
                TCHAR szDt[64];
                _tcsncpy_s(szDt,64, (LPCTSTR)pval[date_col].bstrVal, 63);
                LPCTSTR szTime = _tcschr(szDt, _T(' '));
                if( szTime)
                {
                    szDt[szTime-szDt] =_T('\0');
                    str.Add(szDt);
                    szTime++;
                    while( *szTime && *szTime == ' ') szTime++;                    

                    int h, min;
                    TCHAR hd;
                    if( _stscanf(szTime, _T("%d:%d%c"), &h, &min, &hd) != 0)
                    {
                        if( h>=12) h =0 ;
                        if( hd == _T('P') ) h += 12;
                        _stprintf(szDt,_T("%d:%d\0"), h, min);
                        szTime = szDt;
                    }
                    str.Add(szTime);
                }      
                else
                {
                    str.Add(_T(" "));
                    str.Add(_T(" "));
                }
            }

            for(int i=1; i<pParser->nImportCols; i++)
            {
                if( pval[pParser->ColIndex[i]].vt == VT_BSTR)
                {
                    str.Add(pval[pParser->ColIndex[i]].bstrVal);
                }
                else if(pval[pParser->ColIndex[i]].vt == VT_R8 )
                {
                    TCHAR buff[64];
                    _stprintf(buff, _T("%.02f"), pval[pParser->ColIndex[i]].dblVal);
                    str.Add(buff);
                }
                else if( pval[pParser->ColIndex[i]].vt == VT_EMPTY)
                {
                    str.Add(_T(""));
                }

            }
            InsertRowData(pParser, str);

        }
    }
    return UNCHANGED;
}
//---------------------------------------------
RPOC_STATE OnLookupNextTbl(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    int col = 0; 
    while( FindNextValidColumn(pval, col, num) )
    {
        if( pval[col].vt == VT_BSTR )
        {
            CString str = pval[col].bstrVal;
            if( CheckEndConditon(pstrKeyWords, str) )
                return SKIP_TBLE_HEADER;
        }
    }
    return UNCHANGED;
}
//---------------------------------------------
RPOC_STATE OnSkipHead(VARIANT * pval, int num, LPCTSTR pstrKeyWords[], RPT_PARSER * pParser)
{
    int col =0 ;
    while( FindNextValidColumn(pval, col, num) )
    {
        if( pval[col].vt == VT_BSTR )
        {
            CString str = pval[col].bstrVal;
            if( CheckEndConditon(pstrKeyWords, str) )
            {
                TRACE(_T("table start %d\n"),pParser->nCurReadRow+1);
                return LOOKUP_TBL_DATA;
            }
        }

    }
    return UNCHANGED;
}
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
