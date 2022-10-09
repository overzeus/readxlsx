#include "readxlsx.h"

#define XML_CHARSET_CODE "UTF-8"

/*
函数：利用iconv函数库转换编码
备注：char _OUT_ *out_buf 需要事先申请空间，特别是UTF-8转GB2312时，长度最好是原来的3倍

*/
int Charset_Convert(const char _IN_ *from_charset, const char _IN_ *to_charset,
    char _IN_ *in_buf, size_t _IN_ in_left, char _OUT_ *out_buf, size_t _IN_ out_left) {
        iconv_t icd;
        char *pin = in_buf;
        char *pout = out_buf;
        size_t out_len = out_left;
        if ((iconv_t)-1 == (icd = iconv_open(to_charset,from_charset))) {
            return -1;
        }
        if ((size_t)-1 == iconv(icd, &pin, &in_left, &pout, &out_left)) {
            iconv_close(icd);
            return -1;
        }
        out_buf[out_len - out_left] = 0;
        iconv_close(icd);
        return (int)out_len - out_left;
}


/**
* 函数：输入缓冲区内直接转换编码
* 
**/
int CharsetConvertByCellValueLen(const char _IN_* from_charset, const char _IN_* to_charset, char _IN_OUT_*inbuf, size_t _IN_ inlen) {
    iconv_t icd;
    char* pin = inbuf;
    size_t in_left = inlen;
    char  out[CELL_VALUE_LEN];
    char* pout = out;
    size_t out_len = sizeof(out);
    size_t out_left = out_len;

    bzero(out,sizeof(out));

    if ((iconv_t)-1 == (icd = iconv_open(to_charset, from_charset))) {
        return -1;
    }
    if ((size_t)-1 == iconv(icd, &pin, &in_left, &pout, &out_left)) {
        iconv_close(icd);
        return -1;
    }
    out[out_len - out_left] = 0;
    iconv_close(icd);
    snprintf(inbuf,inlen,"%s",out);
    return (int)out_len - out_left;
}


size_t GetFileSize(char *path)
{
    FILE *fp = NULL;
    int  fileSize = 0;

    if( path == NULL )
        return -1;

    fp = fopen(path,"r");
    if( fp == NULL )
        return -1;

    fseek(fp,0,SEEK_END);
    fileSize=ftell(fp);
    rewind(fp);

    fclose(fp);

    return fileSize;
}


/*

递归创建????

*/
// int mkdirp(char *new_path, int perms)
// {
//     if( new_path == NULL )
//         return -1;

//     char *saved_path, *cp;
//     int saved_ch;
//     struct stat st;
//     int rc = 0;
 
//     cp = saved_path = strdup(new_path);
//     while (*cp && *cp == '/') ++cp;
    
//     while (1) {
//         while (*cp && *cp != '/') ++cp;
//         if ((saved_ch = *cp) != 0)
//             *cp = 0;
            
//         if ((rc = stat(saved_path, &st)) >= 0) {
//             if (!S_ISDIR(st.st_mode)) {
//                 errno = ENOTDIR;
//                 rc = -errno;
//                 break;
//             }
//         } 
//         else {
//             if (errno != ENOENT ) {
//                break;
//             }
            
//             if ((rc = mkdir(saved_path, perms)) < 0 ) {
//                 if (errno != EEXIST)
//                     break;
                
//                 if ((rc = stat(saved_path, &st)) < 0)
//                     break;
                    
//                 if (!S_ISDIR(st.st_mode)) {
//                     errno = ENOTDIR;
//                     rc = -errno;
//                     break;
//                 }                
//             }
//         }
        
//         if (saved_ch != 0)
//             *cp = saved_ch;
        
//         while (*cp && *cp == '/') ++cp;
//         if (*cp == 0)
//             break;

//     }//end while
    
//     free(saved_path);
//     return rc;
// }



#define MAXFILENAME 4096

/*

函数：解压缩xlsx文件
返回值：0-成功 <0失败
备注：输出参数为解压缩的绝对路径，需要自己手动释放
*/
char* UnZipXlsx(const char * _IN_ srczippath)
{
    if( srczippath == NULL || *srczippath == '\0' )
        return NULL;
    
    const char *zipfilename=NULL;
    const char *filename_to_extract=NULL;
    const char *password=NULL;
    char filename_try[MAXFILENAME+1] = "";
    int ret_value=0;
    int opt_do_extract_withoutpath=0, opt_do_list, opt_extractdir;
    int opt_do_extract = 1;
    int opt_overwrite=0;
    int ret = 0;
    int len = 0;
    char *dirname=NULL;
    unzFile uf=NULL;
    
    len = strlen(srczippath) + strlen(".dir") + 1;
    dirname = (char *)malloc(len);
    if( dirname == NULL )
        return NULL;

    snprintf( dirname, len, "%s.dir", srczippath );        


    opt_do_list = 1;
    opt_overwrite = 1;
    opt_extractdir = 1;
    ret = makedir(dirname);
    if( ret < 1 ){
        if(errno != EEXIST)
            return NULL;
    }

    
    zipfilename = srczippath;
    strncpy(filename_try, zipfilename,MAXFILENAME-1);
    /* strncpy doesnt append the trailing NULL, of the string is too long. */
    filename_try[ MAXFILENAME ] = '\0';
    uf = unzOpen64(zipfilename);
    if (uf==NULL)
    {
        strcat(filename_try,".zip");
        uf = unzOpen64(filename_try);
    }

    if (uf==NULL)
    {
        printf("Cannot open %s or %s.zip\n",zipfilename,zipfilename);
        return NULL;
    }
    printf("%s opened\n",filename_try);

    chdir(dirname);
    if (filename_to_extract == NULL)
        ret_value = do_extract(uf, opt_do_extract_withoutpath, opt_overwrite, password);
    else
        ret_value = do_extract_onefile(uf, filename_to_extract, opt_do_extract_withoutpath, opt_overwrite, password);


    unzClose(uf);
    if( ret_value != 0 )
        return NULL;    
    chdir("..");
    return dirname;
}





/*

函数：转换workbook.xml
返回值：0-成功??-1失败

*/
int ParseWorkBookXml(const char * _IN_ xmlcontent,const int _IN_ size, const char * _IN_ tocode,WorkBook * _OUT_ wBook) 
{   
    if( xmlcontent == NULL || size <= 0 || wBook == NULL )
        return -1;

    xmlDocPtr  doc = NULL;
    xmlNodePtr root = NULL,node = NULL, sheetsnode = NULL;
    xmlChar *  sheetname = NULL, *id = NULL;
    int        sheetcnt = 0;
    int        sheetidx = 0;
    int        wsheetsize = 0;
    size_t     tocodebufsize = 0,sheetnamelen = 0;
    char *     tocodebuf = NULL;


    doc = xmlParseMemory(xmlcontent,size);    //parse xml in memory    
    if( doc == NULL )
        return -1;

    root = xmlDocGetRootElement(doc);

    for( node = root->children; node; node = node->next ){
        if(xmlStrcasecmp(node->name,BAD_CAST"sheets")==0)
            break;
    }       

    if(node==NULL){
        return -1;
    }
    
    sheetsnode = node;
    
    sheetcnt = 0;
    for(node=node->children;node;node=node->next){
        sheetcnt++;
    }

    wsheetsize = (sizeof(WorkSheet))*(sheetcnt+1);
    wBook->workSheets = (WorkSheet *)malloc(wsheetsize);
    if( wBook->workSheets == NULL )
        return -errno;

    memset(wBook->workSheets, 0x00, wsheetsize);
    wBook->sheetcnt = sheetcnt+1;

    node = sheetsnode;
    for(node=node->children;node;node=node->next){
        if(xmlStrcasecmp(node->name,BAD_CAST"sheet")==0){
            sheetname = xmlGetProp(node,BAD_CAST"name");
            sheetidx = atoi((char *)xmlGetProp(node,BAD_CAST"sheetId"));
            id = xmlGetProp(node,BAD_CAST"id");
            if( tocode != NULL ){
                sheetnamelen = strlen((char *)sheetname);
                tocodebufsize = sheetnamelen * 4;
                tocodebuf = (char *)malloc(tocodebufsize);
                if( tocodebuf == NULL ){
                    fprintf(stderr,"sheetname tocode %s malloc err!\n",tocode);
                    return -errno;
                }        
                Charset_Convert(XML_CHARSET_CODE,tocode,(char *)sheetname,sheetnamelen,tocodebuf,tocodebufsize);
                //sheetname = tocodebuf;
            }
            if(tocodebuf != NULL)
                strcpy(wBook->workSheets[sheetidx].sheetName, tocodebuf);
            if( tocodebuf != NULL ){
                free(tocodebuf);
                tocodebuf = NULL;
            }
            strcpy(wBook->workSheets[sheetidx].id,(char *)id);
        }

    }

    return 0;
}


/*

函数：workboo.xml.res??有每个sheet对应的id，需要通过id去找对应的sheet
返回值：0-成功??<0失败

*/
int ContainsId(char * _IN_ id, WorkSheet * _IN_ wsheets, int _IN_ sheetscnt)
{
    if( wsheets == NULL || sheetscnt <= 0 )
        return 0;

    int ii = 0;

    for(ii = SHEET_1_IDX; ii < sheetscnt; ii++ ){
        if( strcmp(wsheets[ii].id, id) == 0 ){
            return ii;
        }

    }

    return 0;
}


/*

函数：解析workboo.xml.res
返回值：0-成功??<0失败

*/
int ParseWorkBookXmlRels(const char * _IN_ xmlcontent,const int _IN_ size, WorkBook * _OUT_ wBook)
{
    if( xmlcontent == NULL || size <= 0 || wBook == NULL )
        return -1;

    xmlDocPtr  doc = NULL;
    xmlNodePtr root = NULL,node = NULL;
    xmlChar* id = NULL, * target = NULL;
    char     *pend = NULL;
    int        sheetcnt = 0;
    int        idx = 0;     


    doc = xmlParseMemory(xmlcontent,size);    //parse xml in memory    
    if( doc == NULL )
        return -1;

    root = xmlDocGetRootElement(doc);
    for( node = root->children; node; node = node->next ){
        if(xmlStrcasecmp(node->name,BAD_CAST"Relationship")==0){
            id = xmlGetProp(node,BAD_CAST"Id");
            if((idx=ContainsId((char *)id,wBook->workSheets,wBook->sheetcnt))){
                target = xmlGetProp(node,BAD_CAST"Target");
                pend = strstr((char *)target,"/");
                pend++;
                snprintf(wBook->workSheets[idx].target,sizeof(wBook->workSheets[idx].target),"%s/%s/%s",wBook->unziprootdir,WORKSHEETS_XML_DIR,pend);
                sheetcnt++;

            }
        }
        
    }       

    return (sheetcnt > 0 ? 0:-1 );
}

/*
函数：从sharedstring.xml 获取单元格的值
返回值：0-成功 <0失败
*/
int ParseSharedStringsXml( const char * _IN_ xmlcontent, const int _IN_ size, const char *_IN_ tocode,WorkBook *wBook)
{
    if( xmlcontent == NULL || size <= 0 || wBook == NULL)
        return -1;

    xmlDocPtr  doc = NULL;
    xmlNodePtr root = NULL,node = NULL, tnode = NULL;
    xmlChar *  t = NULL, *scnt = NULL;
    int count = 0;
    int ii = 0;
    size_t tocodebufsize = 0,tlen = 0;
    char *tocodebuf = NULL;


    doc = xmlParseMemory(xmlcontent,size);    //parse xml in memory    
    if( doc == NULL )
        return -1;

    root = xmlDocGetRootElement(doc);
    scnt = xmlGetProp(root,BAD_CAST"count");

    if( scnt == NULL )
        return 0;

    count = atoi((char *)scnt);
    if(count == 0)
        return 0;


    wBook->sharedstrings = (SharedString *)malloc(sizeof(SharedString)*count);
    if(wBook->sharedstrings == NULL)
        return -errno;

    ii = 0;
    for( node = root->children; node; node = node->next ){
        if(xmlStrcasecmp(node->name,BAD_CAST"si")==0){
            for(tnode=node->children; tnode ; tnode=tnode->next ){ 
                if(xmlStrcasecmp(tnode->name,BAD_CAST"t")==0){
                    t=xmlNodeGetContent(tnode);
                    if(tocode != NULL){
                        tlen = strlen((char *)t);
                        tocodebufsize = tlen*4;
                        tocodebuf = (char *)malloc(tocodebufsize);
                        if( tocodebuf == NULL ){
                            fprintf(stderr,"sharedstrings tocode %s malloc err!",tocode);
                            return -errno;
                        }
                        Charset_Convert(XML_CHARSET_CODE,tocode,(char *)t,tlen,tocodebuf,tocodebufsize);
                        //t = tocodebuf;
                    }
                    if (tocodebuf != NULL)
                        strcpy( wBook->sharedstrings[ii].value, tocodebuf);
                    ii++;
                    if( tocodebuf != NULL ){
                        free(tocodebuf);
                        tocodebuf = NULL;
                    }
                }
                
            }    
        }    
    
    }
    return 0;
}

/*

函数：通过传入的dimension计算出单个sheet总共的行数，列数。例 A1:AB10,即共10行27列
返回值：0-成功 <0失败

*/
int CalcRowsAndCols(char * _IN_ dimension, int * _OUT_ rows, int * _OUT_ cols)
{
    if( dimension == NULL || rows == NULL || cols == NULL )
        return -1;

    char *p = NULL;
    char start[10+1];
    char end[10+1];

    char startrow[10+1];
    char endrow[10+1];
    char startcol[10+1];
    char endcol[10+1];
    int ii = 0;
    int jj = 0;
    int iendcol = 0;

    memset(start,0x00,sizeof(start));
    memset(end,0x00,sizeof(end));

    memset(startrow,0x00,sizeof(startrow));
    memset(endrow,0x00,sizeof(endrow));
    memset(startcol,0x00,sizeof(startcol));
    memset(endcol,0x00,sizeof(endcol));

    /*
    dimension?取值举??
    dimension? A1:E5
    dimension? A1:A9
    dimension? A1
    */
    p = strstr(dimension,":");
    if( p != NULL ){
        strncpy( start, dimension, p-dimension ); 
        start[p-dimension] = '\0';
        strcpy( end, p+1 );
    }else{
        strcpy( start, dimension);
        strcpy( end, dimension);

    }


    ii = 0;
    jj = 0;
    for( p = start; *p != '\0'; p++ ){
        if( isdigit(*p) )
            startrow[ii++] = *p;
        else if(isalpha(*p))
            startcol[jj++] = *p;
    }

    ii = 0;
    jj = 0;
    for( p = end; *p != '\0'; p++ ){

        if( isdigit(*p) )
            endrow[ii++] = *p;
        else if(isalpha(*p))
            endcol[jj++] = *p;

    }

        
    *rows = atoi(endrow);
    /* 
    计算当前的英文列号，????几列，从0开始算，例如，AA??26??
    */
    AlphaColToNum(endcol,&iendcol);
    *cols = iendcol + 1;

    return 0;
}

/*

函数：只计算A~Z列应的下标，例：A为0列，如果要计算AB，AZ，ABC这样的列对应的下标，需要调用AlphaColToNum
返回值：0-成功 <0失败

*/
int AlphaColToIndex(char _IN_ ch,int * _OUT_ index)
{
    char tch = toupper(ch);
    if( (tch < 'A' || tch > 'Z') || index == NULL )
        return -1;

    switch (tch)
    {
    case 'A':
        *index = COL_A_IDX; 
        break;    
    case 'B':
        *index = COL_B_IDX; 
        break;
    case 'C':
        *index = COL_C_IDX; 
        break;        
    case 'D':
        *index = COL_D_IDX; 
        break;
    case 'E':
        *index = COL_E_IDX; 
        break;
    case 'F':
        *index = COL_F_IDX; 
        break;
    case 'G':
        *index = COL_G_IDX; 
        break;
    case 'H':
        *index = COL_H_IDX; 
        break;
    case 'I':
        *index = COL_I_IDX; 
        break;
    case 'J':
        *index = COL_J_IDX; 
        break;
    case 'K':
        *index = COL_K_IDX; 
        break;
    case 'L':
        *index = COL_L_IDX; 
        break;
    case 'M':
        *index = COL_M_IDX; 
        break;
    case 'N':
        *index = COL_N_IDX; 
        break;
    case 'O':
        *index = COL_O_IDX; 
        break;
    case 'P':
        *index = COL_P_IDX; 
        break;
    case 'Q':
        *index = COL_Q_IDX; 
        break;
    case 'R':
        *index = COL_R_IDX; 
        break;
    case 'S':
        *index = COL_S_IDX; 
        break;
    case 'T':
        *index = COL_T_IDX; 
        break;
    case 'U':
        *index = COL_U_IDX; 
        break;
    case 'V':
        *index = COL_V_IDX; 
        break;
    case 'W':
        *index = COL_W_IDX; 
        break;
    case 'X':
        *index = COL_X_IDX; 
        break;
    case 'Y':
        *index = COL_Y_IDX; 
        break;
    case 'Z':
        *index = COL_Z_IDX; 
        break;
    default:
        *index = -1;
        break;
    }

    return (*index == -1?-1:0);
}

/*

函数：通过输入的英文字母（列号），计算对应的数字。如果其中有数字则失败
返回值：0-成功 <0失败
备注?? 
输入参数 char * _IN_ pos  英文字母为AB
输出参数 int *_OUT_ pcol 转出数字为27

*/
int AlphaColToNum(char * _IN_ pos, int * _OUT_ pcol)
{
    if( pos == NULL )
        return -1;

    int len = 0, ival = 0,index = 0;
    char scol[5+1];
    char *curp = pos;

    memset(scol,0x00,sizeof(scol));

    len = strlen(pos);
    for( ;*curp != '\0'; curp++ ){
        if(!isalpha(*curp))
            return -1;

        AlphaColToIndex(*curp,&index);
        if( len > 1 )
            ival += (26*len-26)*(index+1);
        else if( len ==1 )
            ival += index;    
        len--;
    }

    *pcol = ival;

    return 0;
}

/*

函数：将输入字符串的英文字母分解出
返回值：0-成功 <0失败
备注?? 该函数只用来解析xlsx对应的A1或E19之类的下标

*/
int SplitAlphaStr(char * _IN_ str, char * _OUT_ as )
{
    if( str == NULL || as == NULL )
        return -1;

    int ii = 0;
    for(; *str != '\0'; str++ )
        if(isalpha(*str))
            as[ii++] = *str;

    return 0;
}

/**
* 交换变量
**/
static inline void swapint(int *i1,int *i2)
{
    int tmp;

    tmp = *i1;
    *i1 = *i2;
    *i2 = tmp;

    return;
}

/***
* 
* 函数：通过xml文档中的<row>标签及子节点的<c>标签，计算总行数和列数
* 返回值：0-成功 <0失败
* 备注：不要相信sheet[1...n].xml中的<dimension ref="A1">中的值，因为可能会有很多行和列，我踩过的坑不会骗你
**/
int CalRowsAndColsByCountRowNode(xmlNodePtr root,int *rows,int *cols)
{
    xmlNodePtr curnode = NULL, rnode = NULL, cnode = NULL;
    xmlChar *rnum = NULL, *cnum = NULL;
    int newrows = 0, newcols = 0, largercols = 0;

    for (curnode = root->children; curnode; curnode = curnode->next) {

        if (xmlStrcasecmp(curnode->name,BAD_CAST"sheetData") == 0) {
       
            /*
            <sheetData>
                <row r="1" spans="1:3">
                    <c r="A1" t="s">
                        <v>0</v>
                    </c>
                    <c r="B1">
                        <v>111</v>
                    </c>
                    <c r="C1" t="s">
                        <v>1</v>
                    </c>
                </row>
            </sheetData>
            */
            for (rnode = curnode->children; rnode; rnode = rnode->next) {/*  <row r="1" spans="1:3"> </row>  */
                rnum = xmlGetProp(rnode, BAD_CAST"r");/** 获取当前的行号 **/
                
                for (cnode = rnode->children; cnode; cnode = cnode->next) { /* <c r="A1" t="s"> 获取列属性*/
                    cnum = xmlGetProp(cnode, BAD_CAST"r");
                    if (cnum != NULL) {
                        newrows = atoi((char *)rnum);
                        newcols++;
                    
                    }//end if  cnum != NULL      
                }//end if cnode
//                printf("newcols = %d\tlargercols = %d\n", newcols, largercols);
                if (newcols > largercols)
                    swapint(&newcols,&largercols);
                newcols = 0;
            } // end if rnode
        }//end if(xmlStrcasecmp(curnode->name,BAD_CAST"sheetData")==0)
    }//end for( curnode = root->children; curnode ;.....
    *rows = newrows;
    *cols = largercols;
    return 0;
}




/*

函数：解析sheet[1...n].xml并写入WorkBook类型的结构体
返回值： 0-成功 <0失败
备注

*/
int ParseSheetXml(const char * _IN_ xmlcontent, const int _IN_ size,const int _IN_ wsheetidx, const char *tocode,WorkBook * _OUT_ wBook)
{
    if( xmlcontent == NULL || size <= 0 || wBook == NULL || wsheetidx < SHEET_1_IDX)
        return -1;

    xmlDocPtr  doc = NULL;
    xmlNodePtr root = NULL, curnode = NULL, rnode = NULL, cnode = NULL, vnode = NULL, isnode = NULL, istnode = NULL;
    xmlChar    * diref = NULL, * rnum = NULL, * cnum = NULL, * t = NULL, * v = NULL, * ist = NULL;
    int  rows = 0, cols = 0, irowpos = 0, icolpos = 0, ivpos = 0, newrows = 0, newcols = 0;
    int ii = 0,jj = 0;
    char strcol[5+1];



    memset(strcol,0x00,sizeof(strcol));

    doc = xmlParseMemory(xmlcontent,size);    //parse xml in memory    
    if( doc == NULL )
        return -1;

    root = xmlDocGetRootElement(doc);
    for( curnode = root->children; curnode; curnode = curnode->next ){
        
        if(xmlStrcasecmp(curnode->name,BAD_CAST"dimension")==0){
            diref = xmlGetProp(curnode,BAD_CAST"ref");

            CalcRowsAndCols((char *)diref,&rows,&cols);
//            printf("rows=%d\tcols=%d by dimension ref\n",rows,cols);
            CalRowsAndColsByCountRowNode(root, &newrows, &newcols);
//            printf("newrows = %d\tnewcols = %d\n", newrows, newcols);

            if (rows < newrows)
                rows = newrows;
            if (cols < newcols)
                cols = newcols;


            wBook->workSheets[wsheetidx].cellrows = rows;
            wBook->workSheets[wsheetidx].cellcols = cols;
            printf("sheet[%s] has rows:[%d] cols:[%d]\n", wBook->workSheets[wsheetidx].sheetName, rows, cols);

            wBook->workSheets[wsheetidx].cells = (Cell **)malloc(rows*sizeof(Cell*));
            memset(wBook->workSheets[wsheetidx].cells, 0x00, rows*sizeof(Cell*));

            for( ii = 0; ii < rows; ii++ ){
                wBook->workSheets[wsheetidx].cells[ii] = (Cell*)malloc(cols*sizeof(Cell));
                memset(wBook->workSheets[wsheetidx].cells[ii], 0x00, cols*sizeof(Cell));
                for(jj = 0; jj < cols; jj++){
                    memset( wBook->workSheets[wsheetidx].cells[ii][jj].value, 0x00, sizeof(wBook->workSheets[wsheetidx].cells[ii][jj].value) );
                }    
            }
            break;
        }    
    }

    for( curnode = root->children; curnode; curnode = curnode->next ){       
        if(xmlStrcasecmp(curnode->name,BAD_CAST"sheetData")==0){
            /*<sheetData>里面有两种方式表示数据
            * 1.EXCEL标准
            <sheetData>
                <row r="1" spans="1:3">
                    <c r="A1" t="s">
                        <v>0</v>
                    </c>
                    <c r="B1">  如果没有属性值t，则直接取v标签的值
                        <v>111</v>
                    </c>
                    <c r="C1" t="s">
                        <v>1</v>
                    </c>
                </row>    
            </sheetData>
            2.SXSSF标准
            <sheetData>
                <row r="1">
                    <c r="A1" s="1" t="inlineStr">
                        <is>
                            <t>&#26631;&#30340;</t>
                        </is>
                    </c>
                    <c r="B1" s="1" t="inlineStr"><is><t>&#26399;&#26435;&#26399;&#38480;</t></is></c>
                    <c r="C1" s="1" t="inlineStr"><is><t>0.0355</t></is></c>
                    <c r="D1" s="1" t="inlineStr"><is><t>0.0360</t></is></c>
                    <c r="E1" s="1" t="inlineStr"><is><t>0.0365</t></is></c>
                    <c r="F1" s="1" t="inlineStr"><is><t>0.0370</t></is></c>
                    <c r="G1" s="1" t="inlineStr"><is><t>0.0375</t></is></c>
                    <c r="H1" s="1" t="inlineStr"><is><t>0.0380</t></is></c>
                    <c r="I1" s="1" t="inlineStr"><is><t>0.0385</t></is></c>
                </row>
            </sheetData>
            */
            for( rnode=curnode->children; rnode; rnode = rnode->next ){/*  <row r="1" spans="1:3"> </row>  行标签*/
                rnum = xmlGetProp(rnode,BAD_CAST"r");               
                for(cnode = rnode->children; cnode; cnode = cnode->next){ /* <c r="A1" t="s"> 列标签,要注意excel模式和sxssf模式的区别 */
                    cnum = xmlGetProp(cnode,BAD_CAST"r");
                    if( cnum != NULL ){
                        irowpos = atoi((char *)rnum) - 1; 
                        SplitAlphaStr((char *)cnum,strcol);
                        AlphaColToNum(strcol,&icolpos); 

                        t = xmlGetProp(cnode,BAD_CAST"t");
                        vnode = cnode->children;
                        v = xmlNodeGetContent(vnode);
                        if (t != NULL) {
                            if (xmlStrcasecmp(t, BAD_CAST"s") == 0) { /** excel模式：t="s"时取sharedstring.xml中的值 **/
                                if (v != NULL) {
                                    ivpos = atoi((char *)v);
                                    strcpy(wBook->workSheets[wsheetidx].cells[irowpos][icolpos].value, wBook->sharedstrings[ivpos].value);
                                }
                            
                            }else if (xmlStrcasecmp(t, BAD_CAST"inlineStr") == 0) { /** sxssf模式：t="inlineStr"时取<is><t>111</is></t>的值 **/
                                isnode = cnode->children;
                                if(isnode != NULL)
                                    istnode = isnode->children;
                                if(istnode != NULL)
                                    ist = xmlNodeGetContent(istnode);

                                if(ist != NULL)
                                    strcpy(wBook->workSheets[wsheetidx].cells[irowpos][icolpos].value, (char *)ist);
                                ist = NULL;
                            }
                        }
                        else { /** excel模式：没有t属性时，直接取v标签中的值 <v>aaa</v> **/
                            strcpy(wBook->workSheets[wsheetidx].cells[irowpos][icolpos].value, (char *)v);
                            v = NULL;
                        
                        }//end if t != NULL else ...
                        if (wBook->workSheets[wsheetidx].cells[irowpos][icolpos].value[0] != '\0')
                            CharsetConvertByCellValueLen(FROM_UTF8, TO_GB2312, wBook->workSheets[wsheetidx].cells[irowpos][icolpos].value, CELL_VALUE_LEN);
                    }//end if  cnum != NULL      
                }//end if cnode
            } // end if rnode
        }//end if(xmlStrcasecmp(curnode->name,BAD_CAST"sheetData")==0)
    }//end for( curnode = root->children; curnode ;.....
   
    return 0;
}



/*

函数：从sheet[1...n].xml获取所有的sheet信息
返回值：0-成功 <-0失败

*/
int ParseWorkSheets(WorkBook * _IN_OUT_ wBook,const char *tocode)
{
    if( wBook == NULL || wBook->sheetcnt < 1)
        return -1;

    int ii = 0, ret = 0;
    char *path = NULL,*xmlbuf = NULL;
    size_t bufsize = 0;

    for( ii = 0; ii < wBook->sheetcnt; ii++ ){

        if( wBook->workSheets[ii].sheetName[0] != '\0' ){
            path = wBook->workSheets[ii].target;
            ret = GetXmlFileContent(path, &xmlbuf, &bufsize);
            if( ret < 0 )
                return -errno;

            ParseSheetXml(xmlbuf,bufsize,ii,tocode,wBook);

        }

    }

    return 0;
}


/*

函数：打印单个worksheet

*/
int PrintWorkSheet(WorkSheet *wsheet)
{
    if(wsheet == NULL)
        return -1;

    if( wsheet->sheetName[0] == '\0' )
        return -1;

    int row = 0;
    int col = 0;

    printf("sheet[%s]\n",wsheet->sheetName);

    for( row = 0; row < wsheet->cellrows; row++ ){
        for (col = 0; col < wsheet->cellcols; col++){
            printf("%s\t",wsheet->cells[row][col].value);                  
        }
        printf("\n");
    }    


    return 0;
}

/*

函数：打印整个xlsx

*/
int PrintWorkBook(WorkBook *wBook)
{
    if( wBook == NULL )
        return -1;

    int ii = 0;
    int row = 0, col = 0; 
    WorkSheet *wsheet = NULL;

    
    for( ii = 0; ii < wBook->sheetcnt; ii++){
        wsheet = wBook->workSheets+ii;
        if( wsheet->sheetName[0] != '\0' ){
            printf("sheetname:[%s]\n",wsheet->sheetName);
            for( row = 0; row < wsheet->cellrows; row++ ){
                for (col = 0; col < wsheet->cellcols; col++){
                    printf("%s\t",wsheet->cells[row][col].value);                  
                }
                printf("\n");
            }
        }
    }

    return 0;
}

/*

函数：释放workbook空间

*/

int CloseXlsx(WorkBook *wBook)
{
    if( wBook == NULL )
        return -1;
        
    int ii = 0, row = 0;
    WorkSheet *wsheet = NULL;

    if( wBook->workSheets != NULL ){
        for( ii = 0; ii < wBook->sheetcnt; ii++ ){
            wsheet = wBook->workSheets+ii;
            if( wsheet != NULL ){
                if( wsheet->cells != NULL ){
                    for( row = 0; row <wsheet->cellrows; row++ ){
                        if(wsheet->cells[row] != NULL){
                            free(wsheet->cells[row]);
                            wsheet->cells[row] = NULL;
                        }    
                    }
                    free(wsheet->cells);
                    wsheet->cells = NULL;
                }    
            }    
        }

        free(wBook->workSheets);
        wBook->workSheets = NULL;
    }    

    if(wBook->sharedstrings != NULL){
        free(wBook->sharedstrings);
        wBook->sharedstrings = NULL;
    }    

    free(wBook);
    return 0;
}





/*
函数：获取xml的内容，并计算文件大小
返回值：0-成功 <0失败
*/

int GetXmlFileContent(const char * _IN_ path, char ** _OUT_ xmlcontent, size_t * _OUT_ outsize )
{
    char *content = NULL;
    FILE *fp = NULL;
    int  fileSize = 0;
    if( path == NULL )
        return -1;

    fp = fopen(path,"r");
    if( fp == NULL )
        return -1;

    fseek(fp,0,SEEK_END);
    fileSize=ftell(fp);
    rewind(fp);
    
    content = (char *)malloc(fileSize);
    if( content == NULL )
        return -1;
    bzero(content,fileSize);

    fread(content,1,fileSize,fp);
    fclose(fp);

    *xmlcontent = content;
    *outsize = fileSize;

    return 0;
}


/*

函数：打开一个xlsx文件
返回值：NULL则失败
备注：
打开XLSX有两步骤：1用zip解压缩成一个目录；2.对该目录中的xml进行解析

*/
WorkBook* OpenXlsx(const char _IN_ *path,const char _IN_ *tocode)
{
    if( path == NULL )
        return NULL;

    WorkBook *wBook = NULL;
//    SharedString *sharestrings = NULL;
    char *xmlbuf = NULL;
    char xmlpath[4096 + 1];
    char *dstdir = NULL;
    char rootdir[1024 + 1];
    int ret = 0;
    size_t bufsize = 0;

    bzero(rootdir,sizeof(rootdir));
    bzero(xmlpath,sizeof(xmlpath));

    wBook = (WorkBook*)malloc(sizeof(WorkBook));
    if( wBook == NULL ){
        fprintf(stderr,"WorkBook malloc Memory failed!\n");
        return NULL;
    }

    bzero(wBook, sizeof(WorkBook));
    
    dstdir = UnZipXlsx(path);
    if( dstdir == NULL ){
        fprintf(stderr,"Unzip xlsx file failed!\n");
        return NULL;
    }
    
    if(dstdir[0] == '/')
        snprintf(wBook->unziprootdir, sizeof(wBook->unziprootdir), "%s",dstdir);
    else {
        getcwd(rootdir,sizeof(rootdir));
        snprintf(wBook->unziprootdir, sizeof(wBook->unziprootdir), "%s/%s", rootdir, dstdir);
    }
    free(dstdir);
    dstdir = NULL;

    bzero(xmlpath,sizeof(xmlpath));
    snprintf(xmlpath,sizeof(xmlpath),"%s/%s",wBook->unziprootdir,WORKBOOK_XML_PATH);
    ret = GetXmlFileContent(xmlpath,&xmlbuf,&bufsize);
    if( ret < 0 ){
        fprintf(stderr,"Load %s failed,exit\n", xmlpath);
        return NULL;
    }

    ret = ParseWorkBookXml(xmlbuf, bufsize, tocode, wBook);
    if( ret < 0 ){
        fprintf(stderr,"Parse WorkBook.Xml[%s] failed\n", xmlpath);
        goto error;
    }     
    free(xmlbuf);
    xmlbuf = NULL;


    bzero(xmlpath,sizeof(xmlpath));
    snprintf(xmlpath,sizeof(xmlpath),"%s/%s",wBook->unziprootdir,WORKBOOK_XML_RELS_PATH);
    ret = GetXmlFileContent(xmlpath,&xmlbuf,&bufsize);
    if( ret < 0 ){
        fprintf(stderr,"Load [%s] failed,exit\n", xmlpath);
        return NULL;
    }


    ret = ParseWorkBookXmlRels(xmlbuf,bufsize,wBook);
    if( ret < 0 ){
        fprintf(stderr,"Parse WorkBook.Xml.Rels[%s] failed\n", xmlpath);
        goto error;
    }
    free(xmlbuf);
    xmlbuf = NULL;



    snprintf(xmlpath,sizeof(xmlpath),"%s/%s",wBook->unziprootdir,SHAREDSTRINGS_XML_PATH);
    ret = GetXmlFileContent(xmlpath,&xmlbuf,&bufsize);  
    if( ret < 0 ){
        fprintf(stderr,"Load [%s] failed,exit\n", xmlpath);
        return NULL;
    }


    ret = ParseSharedStringsXml(xmlbuf,bufsize,tocode,wBook);
    if( ret < 0 ){
        fprintf(stderr,"Parse sharedString.Xml[%s] failed\n", xmlpath);
        goto error;
    }    
    free(xmlbuf);
    xmlbuf = NULL;

    ret = ParseWorkSheets(wBook,tocode); 
    if( ret < 0 ){
        fprintf(stderr,"Parse All sheet.Xml[%s] failed\n", xmlpath);
        goto error;
    }

    




    return wBook;

error:
    free(xmlbuf);
    xmlbuf = NULL;
    return NULL;
}



WorkSheet *GetWorkSheetByName(const WorkBook _IN_ *wBook, const char * _IN_ sheetname)
{
    if( wBook == NULL || sheetname == NULL || wBook->sheetcnt < 1 )
        return NULL;

    WorkSheet *wSheet = NULL;
    int ii = 0;

    for( ii = 0; ii < wBook->sheetcnt; ii++ ){
        if( strcmp(sheetname,wBook->workSheets[ii].sheetName) == 0 ){
            wSheet = &(wBook->workSheets[ii]);
            break;
        }
    }

    return wSheet;
}

WorkSheet *GetWorkSheetByIndex(const WorkBook _IN_ *wBook, const int _IN_ index)
{
    if( wBook == NULL || index < 0 || wBook->sheetcnt < 1 )
        return NULL;

    WorkSheet *wSheet = NULL;
    wSheet = &(wBook->workSheets[index]);
    return wSheet;
}




