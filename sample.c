#include "readxlsx.h"

int main()
{

    /*
    convert to gb2312
    */
    WorkBook *wBook = OpenXlsx("test.xlsx","gb2312");
//    PrintWorkBook(wBook);
    
    /* 
     worksheet no[1...n]
    */

    int ii = 0;
    WorkSheet* wsheet = NULL;
    for (ii = SHEET_1_IDX; ii < wBook->sheetcnt; ii++) {
        wsheet = GetWorkSheetByIndex(wBook,ii);
        /*
        worksheet is Two-dimensional array .
        */
        PrintWorkSheet(wsheet);
        //printf("%s\n",wsheet->cells[0][1].value);
    }

    /* 
    通过sheet名字查找指定sheet
    */
    wsheet = GetWorkSheetByName(wBook,"SHEET1");
    PrintWorkSheet(wsheet);




    CloseXlsx(wBook);
    return 0;
}
