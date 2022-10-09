#ifndef _MINIUZ_H_
#define _MINIUZ_H_
#include "unzip.h"



int do_list(unzFile uf);

int makedir (char *newdir);


int mymkdir(const char*dirname);

int do_extract(unzFile uf,int opt_extract_without_path,int opt_overwrite,const char*password); 

int do_extract_currentfile(unzFile uf,const int* popt_extract_without_path,int* popt_overwrite,const char*password);

int do_extract_onefile(unzFile uf,const char* filename,int opt_extract_without_path,int opt_overwrite,const char* password);

#endif
