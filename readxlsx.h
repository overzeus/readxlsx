#ifndef _READXLSX_H_
#define _READXLSX_H_

#include <stdio.h>
#include <string.h>
#include <stdlib.h>
#include <libxml/parser.h>
#include <libxml/tree.h>
#include <errno.h>
#include <ctype.h>
#include <zlib.h>
#include <sys/stat.h>
#include <sys/types.h>
#include <unistd.h>
#include <iconv.h>
#include "zip.h"
#include "unzip.h"
#include "ioapi.h"
#include "miniunz.h"



#ifndef _IN_
#define _IN_
#endif


#ifndef _OUT_
#define _OUT_
#endif

#ifndef _IN_OUT_
#define _IN_OUT_
#endif


#define SHEET_1_IDX 1
#define SHEET_2_IDX 2
#define SHEET_3_IDX 3
#define SHEET_4_IDX 4
#define SHEET_5_IDX 5
#define SHEET_6_IDX 6
#define SHEET_7_IDX 7
#define SHEET_8_IDX 8


#define COL_A_IDX 0
#define COL_B_IDX 1
#define COL_C_IDX 2
#define COL_D_IDX 3
#define COL_E_IDX 4
#define COL_F_IDX 5
#define COL_G_IDX 6
#define COL_H_IDX 7
#define COL_I_IDX 8
#define COL_J_IDX 9
#define COL_K_IDX 10
#define COL_L_IDX 11
#define COL_M_IDX 12
#define COL_N_IDX 13
#define COL_O_IDX 14
#define COL_P_IDX 15
#define COL_Q_IDX 16
#define COL_R_IDX 17
#define COL_S_IDX 18
#define COL_T_IDX 19
#define COL_U_IDX 20
#define COL_V_IDX 21
#define COL_W_IDX 22
#define COL_X_IDX 23
#define COL_Y_IDX 24
#define COL_Z_IDX 25


#define _DEBUG_XLSX_
#ifdef _DEBUG_XLSX_

#define _DEBUG_STR_LN_(STR) printf(#STR":[%s]\n",(STR))
#define _DEBUG_INT_LN_(INTNUM) printf(#INTNUM":[%d]\n",(INTNUM))
#define _DEBUG_CHAR_LN_(CH) printf(#CH":[%c]\n",(CH))
#endif

#define XL_DIR "xl/"
#define WORKSHEETS_XML_DIR       XL_DIR"worksheets/"

#define WORKBOOK_XML_RELS_PATH   XL_DIR"_rels/workbook.xml.rels"
#define WORKBOOK_XML_PATH        XL_DIR"workbook.xml"
#define SHAREDSTRINGS_XML_PATH   XL_DIR"sharedStrings.xml"

#define TO_UTF8 "UTF-8//IGNORE"
#define TO_GB2312 "GB2312//IGNORE"
#define FROM_UTF8 "UTF-8"
#define FROM_GB2312 "GB2312"


#define CELL_VALUE_LEN  4096
#define UTF_8_SPACE     100

typedef unsigned char bytes_t;

typedef struct _CELL_
{
    char value[CELL_VALUE_LEN+ UTF_8_SPACE];/** 如果转UTF8可能会需要多一点的空间，所以先预留100个字节 **/
}Cell;

typedef struct _SHEET_
{
    char sheetName[100+1]; /* sheet页名字 */
    char id[10+1];  /* sheet页id,对应用层无影响 */
    Cell **cells;   /* 二维单元格数组 */
    int  cellrows;  /* 单元格总行数 */
    int  cellcols;  /* 单元格总列数 */
    char target[1024+1]; /* 解压缩后sheet[1...n].xml的绝对路径 */
}WorkSheet;

typedef struct _SHARED_STRING_
{
    char value[2048];

}SharedString;

typedef struct _BOOK_
{
    WorkSheet *workSheets; /* sheet页数组 */
    int sheetcnt; /* sheet页的总页数 */
    char unziprootdir[4096+1];/* xlsx 解压缩的根目录 */
    SharedString *sharedstrings;/* 单元格中字符串类型的值 */
}WorkBook;


int ParseWorkBookXml(const char * _IN_ xmlcontent,const int _IN_ size, const char * _IN_ tocode,WorkBook * _OUT_ wBook);
int ContainsId(char * _IN_ id, WorkSheet * _IN_ wsheets, int _IN_ sheetscnt);
int ParseWorkBookXmlRels(const char * _IN_ xmlcontent,const int _IN_ size, WorkBook * _OUT_ wBook);
int ParseSharedStringsXml( const char * _IN_ xmlcontent, const int _IN_ size, const char *_IN_ tocode, WorkBook *wBook);
int CalcRowsAndCols(char * _IN_ dimension, int * _OUT_ rows, int * _OUT_ cols);
int AlphaColToIndex(char _IN_ ch,int * _OUT_ index);
int AlphaColToNum(char * _IN_ pos,  int * _OUT_ pcol);
int SplitAlphaStr(char * _IN_ str, char * _OUT_ as );
int ParseSheetXml(const char* _IN_ xmlcontent, const int _IN_ size, const int _IN_ wsheetidx, const char* tocode, WorkBook* _OUT_ wBook);
int ParseWorkSheets(WorkBook * _IN_OUT_ wBook, const char* tocode);
int PrintWorkBook(WorkBook _IN_ *wBook);
int CloseXlsx(WorkBook _IN_ *wBook);
int GetXmlFileContent(const char * _IN_ path, char ** _OUT_ xmlcontent, size_t * _OUT_ outsize );
int PrintWorkSheet(WorkSheet *_IN_ wsheet);
WorkSheet *GetWorkSheetByName(const WorkBook _IN_ *wBook, const char * _IN_ sheetname);
WorkSheet *GetWorkSheetByIndex(const WorkBook _IN_ *wBook, const int _IN_ index);
WorkBook* OpenXlsx(const char _IN_ *path,const char _IN_ *tocode);

#endif

