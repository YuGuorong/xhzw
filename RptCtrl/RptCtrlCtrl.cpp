// RptCtrlCtrl.cpp : Implementation of the CRptCtrlCtrl ActiveX Control class.

#include "stdafx.h"
#include "RptCtrl.h"
#include "RptCtrlCtrl.h"
#include "RptCtrlPropPage.h"


#ifdef _DEBUG
#define new DEBUG_NEW
#endif



#include "includes/excel8.h"
#include "ParserProc.h"

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

IMPLEMENT_DYNCREATE(CRptCtrlCtrl, COleControl)



// Message map

BEGIN_MESSAGE_MAP(CRptCtrlCtrl, COleControl)
	ON_OLEVERB(AFX_IDS_VERB_PROPERTIES, OnProperties)
END_MESSAGE_MAP()



// Dispatch map

BEGIN_DISPATCH_MAP(CRptCtrlCtrl, COleControl)
	DISP_FUNCTION_ID(CRptCtrlCtrl, "AboutBox", DISPID_ABOUTBOX, AboutBox, VT_EMPTY, VTS_NONE)
    DISP_FUNCTION_ID(CRptCtrlCtrl, "AnalyzeExcleFile", dispidAnalyzeExcleFile, AnalyzeExcleFile, VT_I4, VTS_BSTR)
    DISP_FUNCTION_ID(CRptCtrlCtrl, "AnalyzeFolder", dispidAnalyzeFolder, AnalyzeFolder, VT_I4, VTS_BSTR)
    DISP_PROPERTY_NOTIFY_ID(CRptCtrlCtrl, "strExpFile", dispidstrExpFile, m_strExpFile, OnstrExpFileChanged, VT_BSTR)
    DISP_PROPERTY_NOTIFY_ID(CRptCtrlCtrl, "StrExportFileName", dispidStrExportFileName, m_StrExportFileName, OnStrExportFileNameChanged, VT_BSTR)
END_DISPATCH_MAP()



// Event map

BEGIN_EVENT_MAP(CRptCtrlCtrl, COleControl)
END_EVENT_MAP()



// Property pages

// TODO: Add more property pages as needed.  Remember to increase the count!
BEGIN_PROPPAGEIDS(CRptCtrlCtrl, 1)
	PROPPAGEID(CRptCtrlPropPage::guid)
END_PROPPAGEIDS(CRptCtrlCtrl)



// Initialize class factory and guid
//A7116AA5-99DF-4310-8840-D02869149B7D
IMPLEMENT_OLECREATE_EX(CRptCtrlCtrl, "RPTCTRL.RptCtrlCtrl.1",
	0xa7116aa5, 0x99df, 0x4310, 0x88, 0x40, 0xd0, 0x28, 0x69, 0x14, 0x9b, 0x7d)



// Type library ID and version

IMPLEMENT_OLETYPELIB(CRptCtrlCtrl, _tlid, _wVerMajor, _wVerMinor)



// Interface IDs

const IID BASED_CODE IID_DRptCtrl =
		{ 0x615E2844, 0x9D51, 0x45D3, { 0x94, 0xA3, 0xA0, 0xB7, 0x20, 0x80, 0x33, 0x21 } };
const IID BASED_CODE IID_DRptCtrlEvents =
		{ 0xE8D70003, 0x4295, 0x409A, { 0xB2, 0x93, 0x77, 0x65, 0x51, 0x3, 0x54, 0x4D } };



// Control type information

static const DWORD BASED_CODE _dwRptCtrlOleMisc =
	OLEMISC_ACTIVATEWHENVISIBLE |
	OLEMISC_SETCLIENTSITEFIRST |
	OLEMISC_INSIDEOUT |
	OLEMISC_CANTLINKINSIDE |
	OLEMISC_RECOMPOSEONRESIZE;

IMPLEMENT_OLECTLTYPE(CRptCtrlCtrl, IDS_RPTCTRL, _dwRptCtrlOleMisc)



// CRptCtrlCtrl::CRptCtrlCtrlFactory::UpdateRegistry -
// Adds or removes system registry entries for CRptCtrlCtrl

BOOL CRptCtrlCtrl::CRptCtrlCtrlFactory::UpdateRegistry(BOOL bRegister)
{
	// TODO: Verify that your control follows apartment-model threading rules.
	// Refer to MFC TechNote 64 for more information.
	// If your control does not conform to the apartment-model rules, then
	// you must modify the code below, changing the 6th parameter from
	// afxRegApartmentThreading to 0.

	if (bRegister)
		return AfxOleRegisterControlClass(
			AfxGetInstanceHandle(),
			m_clsid,
			m_lpszProgID,
			IDS_RPTCTRL,
			IDB_RPTCTRL,
			afxRegApartmentThreading,
			_dwRptCtrlOleMisc,
			_tlid,
			_wVerMajor,
			_wVerMinor);
	else
		return AfxOleUnregisterClass(m_clsid, m_lpszProgID);
}



// CRptCtrlCtrl::CRptCtrlCtrl - Constructor

CRptCtrlCtrl::CRptCtrlCtrl()
{
	InitializeIIDs(&IID_DRptCtrl, &IID_DRptCtrlEvents);
	// TODO: Initialize your control's instance data here.
}



// CRptCtrlCtrl::~CRptCtrlCtrl - Destructor

CRptCtrlCtrl::~CRptCtrlCtrl()
{
	// TODO: Cleanup your control's instance data here.
}



// CRptCtrlCtrl::OnDraw - Drawing function

void CRptCtrlCtrl::OnDraw(
			CDC* pdc, const CRect& rcBounds, const CRect& rcInvalid)
{
	if (!pdc)
		return;

	// TODO: Replace the following code with your own drawing code.
	pdc->FillRect(rcBounds, CBrush::FromHandle((HBRUSH)GetStockObject(WHITE_BRUSH)));
	//pdc->Ellipse(rcBounds);
    pdc->FrameRect(rcBounds, &CBrush::CBrush(RGB(0,0,0)));
}



// CRptCtrlCtrl::DoPropExchange - Persistence support

void CRptCtrlCtrl::DoPropExchange(CPropExchange* pPX)
{
	ExchangeVersion(pPX, MAKELONG(_wVerMinor, _wVerMajor));
	COleControl::DoPropExchange(pPX);

	// TODO: Call PX_ functions for each persistent custom property.
}



// CRptCtrlCtrl::OnResetState - Reset control to default state

void CRptCtrlCtrl::OnResetState()
{
	COleControl::OnResetState();  // Resets defaults found in DoPropExchange

	// TODO: Reset any other control state here.
}



// CRptCtrlCtrl::AboutBox - Display an "About" box to the user

void CRptCtrlCtrl::AboutBox()
{
	CDialog dlgAbout(IDD_ABOUTBOX_RPTCTRL);
	dlgAbout.DoModal();
}



// CRptCtrlCtrl message handlers
extern  int g_nExpFileType ;

LONG CRptCtrlCtrl::AnalyzeExcleFile(LPCTSTR szExcleFileName)
{
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

	// TODO: Add your implementation code here
    CString strExpFileName;
    ParseExcle(szExcleFileName,  NULL, strExpFileName);//_T("D:\\temp\\cnbjce02$$cnbjce04.xls")
    m_strExpFile = strExpFileName;
    g_nExpFileType = 1;
    return 0;
}

LONG CRptCtrlCtrl::AnalyzeFolder(LPCTSTR szFolderName)
{
    AFX_MANAGE_STATE(AfxGetStaticModuleState());
 //   ::MessageBox(NULL, szFolderName,L"folder",MB_OK);
    // TODO: Add your dispatch handler code here
//查找文件的函数调用 
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
            m_StrExportFileName = parser.strExpFileName;
        }
        
        parser.objApp.Quit(); 
    }
    
    TRACE(_T("Time used: %d\n"), GetTickCount()-tick );
    g_nExpFileType = 1;

    return 0;
}

void CRptCtrlCtrl::OnstrExpFileChanged(void)
{
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    // TODO: Add your property handler code here

    SetModifiedFlag();
}

void CRptCtrlCtrl::OnStrExportFileNameChanged(void)
{
    AFX_MANAGE_STATE(AfxGetStaticModuleState());

    // TODO: Add your property handler code here

    SetModifiedFlag();
}
