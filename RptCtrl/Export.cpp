#include "stdafx.h"
#include <io.h>
#include <odbcinst.h>
#include <afxdb.h>

#include "Includes\excel8.h"
#include "ParserProc.h"

///////////////////////////////////////////////////////////////////////////////
//	BOOL MakeSurePathExists( CString &Path,bool FilenameIncluded)
//	参数：
//		Path				路径
//		FilenameIncluded	路径是否包含文件名
//	返回值:
//		文件是否存在
//	说明:
//		判断Path文件(FilenameIncluded=true)是否存在,存在返回TURE，不存在返回FALSE
//		自动创建目录
//
///////////////////////////////////////////////////////////////////////////////
TCHAR g_FileExt[][5] = {{_T(".xls")},  {_T(".txt")}};
 int g_nExpFileType = 1;

BOOL MakeSurePathExists( CString &Path, bool Write0Read1)
{
    INT mode = ( Write0Read1 ) ? 4 : 2;
	return !_taccess(Path,mode);//return !_access(Path,0); //	
}

//获得默认的文件名
BOOL GetDefaultXlsFileName(LPCTSTR sDirName, CString& sExcelFile)
{
	///默认文件名：yyyymmddhhmmss.xls
	CString timeStr;
	CTime day;
	day=CTime::GetCurrentTime();
	int filenameday,filenamemonth,filenameyear,filehour,filemin,filesec;
	filenameday=day.GetDay();//dd
	filenamemonth=day.GetMonth();//mm月份
	filenameyear=day.GetYear();//yyyy
	filehour=day.GetHour();//hh
	filemin=day.GetMinute();//mm分钟
	filesec=day.GetSecond();//ss
	timeStr.Format(_T("%04d%02d%02d%02d%02d%02d"),filenameyear,filenamemonth,filenameday,filehour,filemin,filesec);
	
	sExcelFile =  timeStr + g_FileExt[g_nExpFileType];

    //if( sDirName == NULL)

    CString strRpt = sDirName;
    sExcelFile = strRpt + _T("\\") + sExcelFile ;
    
    if (MakeSurePathExists(sExcelFile,true)) 
    {
		if(!DeleteFile(sExcelFile)) 
        {    // delete the file
			return FALSE;
		}
	}
	return TRUE;
}

CFile exp_f;
INT PrepareExportFile(RPT_PARSER *pParser, LPCTSTR szCol , int ncols)
{
    CString strExpFile;
    GetDefaultXlsFileName(pParser->strRptDir, strExpFile);
    pParser->strExpFileName = strExpFile;

    if( g_nExpFileType == 0 )
    {
        COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
        
        pParser->ExpBook= pParser->objBooks.Add(covOptional);
        pParser->ExpSheets=pParser->ExpBook.GetSheets();
        pParser->ExpSheet=pParser->ExpSheets.GetItem(COleVariant((short)1));
        
        //InsertRowData(pParser, szCol,ncols);
        
        pParser->ExpBook.SaveAs(COleVariant(strExpFile),covOptional,
            covOptional,covOptional,
            covOptional,covOptional,(long)0,covOptional,covOptional,covOptional,
            covOptional);
        
        strExpFile.Empty();
        return 0;
    }
    else
    {
        if( exp_f.Open(strExpFile, CFile::modeReadWrite|CFile::modeCreate) )
        {
            pParser->handle_export = &exp_f;
            return 1;
        }                
    }
    return -1;
    
}

/*"insert into hcode(hcode,postcode,cityName) values('1894297', '716', '荆州');"*/
char  g_expPreText[] = {("insert into hcode(hcode,postcode,cityName) values(\'")};
INT InsertRowData(RPT_PARSER *pParser, LPCTSTR szCol,  int nCols)
{
    // TODO: Add your control notification handler code here
    pParser->nCurWriteRow++;
    if( szCol )
    {
        if( g_nExpFileType == 0 )
        {
            int row = pParser->nCurWriteRow;
            for(int i=0;i<nCols;i++)
            {
                Range range;
                TCHAR buff[9]; 
                GetRangString(buff, 8, i, row);
                range=pParser->ExpSheet.GetRange(COleVariant(buff),COleVariant(buff));
                range.SetValue(COleVariant(szCol));
            }
        }
        else
        {
            exp_f.Write(g_expPreText, strlen(g_expPreText));

            //char tmp[1024];
            //DWORD dwNum  = ::WideCharToMultiByte(CP_ACP,0,szCol[i],-1,tmp,1024,0,0);
            char tmps[1024];            
            sprintf(tmps,("%S%03d\',\'%s\',\'%s\');\r\n"),szCol,nCols,pParser->pszInfo[2],pParser->pszInfo[1]);
            exp_f.Write(tmps, strlen(tmps));          
        }

    }

    return 0;
}

INT InsertRowData(RPT_PARSER *pParser, CStringArray &astr)
{
    pParser->nCurWriteRow++;
    if( astr.GetSize() )
    {
        char tmp[1024];
        int row = pParser->nCurWriteRow;
        int nCols = astr.GetSize();
        
        if( g_nExpFileType == 0 )
        {
            for(int i=0;i<nCols;i++)
            {
                Range range;
                TCHAR buff[9]; 
                GetRangString(buff, 8, i, row);
                
                range=pParser->ExpSheet.GetRange(COleVariant(buff),COleVariant(buff));
                range.SetValue(  COleVariant(astr.GetAt(i))  );
            }
        }
        else
        { 
            for(int i=0;i<nCols;i++)
            {
                DWORD dwNum  = ::WideCharToMultiByte(CP_ACP,0,astr.GetAt(i),-1,tmp,1024,0,0);
                exp_f.Write(tmp, dwNum-1 );                
                exp_f.Write( ",\t",1);
            }
             exp_f.Write( ("\r\n"),2);
       }
    }
    return 0;

}

void ExportDataDone(RPT_PARSER *pParser)    
{
    if( g_nExpFileType == 0 )
    {
        COleVariant VOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);
        pParser->ExpBook.Close(COleVariant((short)TRUE),   VOptional,   VOptional);         
    }
    else
    {
        exp_f.Write( ("\r\n"),2);
        exp_f.Close();
    }
}


int ExportExcelFile()
{
    return 0 ;
}
