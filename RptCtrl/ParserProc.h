#ifndef INCLUDE_PARSERPROC_H
#define INCLUDE_PARSERPROC_H

typedef int (*cbExcelRead)(VARIANT * pval, int num, int repeat, void * param);

//Export columns
typedef enum  tag_storage_string_info
{
    CONNECT_NAME,
    NodeA,
    NodeB,
    IP_SOURCE,
    IP_DEST,
    MAX_COLLECTIONS,

    SAMPLE_TIME = MAX_COLLECTIONS,
    LATENCY,
    PACKET_LOSS,
    SAMPLE_COUNT,
    MAX_EXP_COL
}EXP_COLS;

typedef struct tag_column_idx_string
{
    EXP_COLS  index;
    LPCTSTR   strFlag;
}COL_FLAG;

typedef enum tag_porcess_status
{
    UNINITIALIZE,      /*File Opened, read CONNECT_NAME*/ 
    PROC_CT_ADJ,    
    PROC_CT_NORMAL, 
    PROC_CM_NORMAL, 
    PROC_CNN_NORMAL,
    PROC_EXP_DATA,
    STATE_END,       
    UNCHANGED,         /*Dose not change, for call back return check*/
}RPOC_STATE;

typedef struct  tag_report_parser
{
    VOID *       handle_export;
    RPOC_STATE   cur_state;
    
    DWORD        nToalCols;
    DWORD        nToalRows;
    DWORD        nCurReadRow;
    DWORD        nExpCols;

    DWORD        nExpRow;
    DWORD        nImportCols;
    INT          ColIndex[MAX_EXP_COL - MAX_COLLECTIONS];

    CStringArray strInfo;  //strInfo[MAX_COLLECTIONS];
    CString      strRptDir;
    CString      strExpFileName;
    EXP_COLS     tblExpCols[MAX_EXP_COL];


    _Application objApp; 
    Workbooks    objBooks; 

    _Workbook    objBook; 
    Worksheets   objSheets; 
    _Worksheet   objSheet; 

    _Workbook    ExpBook;
    Worksheets   ExpSheets;
    _Worksheet   ExpSheet;
    DWORD        nCurWriteRow;

}RPT_PARSER;


typedef RPOC_STATE (*cbProcRow)(VARIANT * pval, int num, LPCTSTR strKeyWords[], RPT_PARSER * param);
#define DECLARE_PROCESS(proc)  RPOC_STATE proc(VARIANT * pval, int num, LPCTSTR strKeyWords[], RPT_PARSER * pParser);

DECLARE_PROCESS(OnInit);
DECLARE_PROCESS(OnIndexRow);
DECLARE_PROCESS(ProcCtTbl   );
DECLARE_PROCESS(ProcCmTbl   );
DECLARE_PROCESS(ProcCnnTbl  );
DECLARE_PROCESS(ProcExpData );

typedef struct tag_process_table
{
    RPOC_STATE  state;
    cbProcRow   pProc;
    LPCTSTR  *  pszKeyWords;
}PROCESS_TABLE;


class eCException: public CException
{
public:
    eCException(){};
    virtual ~eCException() {};
};

BOOL GetRangString(TCHAR * buff, int len, int cols, int rows);
INT PrepareExportFile(RPT_PARSER *pParser, LPCTSTR szCol[], int cols);
INT InsertRowData(RPT_PARSER *pParser, LPCTSTR szCol[], int cols);
INT InsertRowData(RPT_PARSER *pParser, CStringArray &astr);
void ExportDataDone(RPT_PARSER *pParser)    ;
BOOL GetDefaultXlsFileName(LPCTSTR sDirName, CString& sExcelFile);
BOOL MakeSurePathExists( CString &Path,bool Write0Read1);

#endif /*INCLUDE_PARSERPROC_H*/


/*
int OnExcelReadRow(VARIANT * pval, int num, int repeat, void * param)
{
    VARIANT val;
    static int state;
    for( int i=0 ; i<num; i++)
    {
        val = pval[i];
        switch(val.vt) 
        { 
        case VT_R4:
            TRACE( _T("r4\n"));        
            break;
        case VT_CY:
            TRACE( _T("cy\n"));        
            break;
        case VT_I8:
            TRACE( _T("i8\n"));        
            break;
        case   VT_R8: 
            { 
                //http://zhidao.baidu.com/question/122953167
                double vl = (double)val.dblVal;
                TRACE( _T("\t\t%.2lf "),   vl); 
            } 
            break; 
        case   VT_BSTR: 
            { 
                
                //TRACE( _T("\t\t%s "),(CString)val.bstrVal); 
            } 
            break; 
        case   VT_EMPTY: 
            { 
                //TRACE( _T("\t\t <empty> ")); 
            } 
            break; 
        case VT_DATE:
            TRACE( _T("DATE\n"));        
            break;
        default:
            break;
        } 
    }
   // TRACE( _T("\n ")); 
    return 1;
}
*/
