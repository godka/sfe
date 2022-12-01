#include <windows.h>
#include <stdio.h>
//#include <math.h>


typedef int (*myfunc)();

short	WINAPI	INTrol(short a, short n);

short	WINAPI	INTror(short a, short n);

char	WINAPI	BYTEror(char a, char n);

char	WINAPI	BYTErol(char a, char n);

void	WINAPI	getbyteItem(unsigned char a[] , int LENGTH,short tkey );

void	WINAPI	convertCOLOR(unsigned int *mcolor_RGB , unsigned int *data,int WW,int HH,int MaskColor);

unsigned char WINAPI get256(unsigned int *mcolor_RGB , unsigned int d);

void	WINAPI	convertCOLOR2(unsigned int *mcolor_RGB , unsigned int *data,int WW,int HH,unsigned int MaskColor1);

void WINAPI drawPic(short MapNum, unsigned int X,unsigned int Y,int datalong,unsigned char *data,unsigned int CenterX,unsigned int CenterY,HDC hdc);

void WINAPI cpymem(int *dest,int value);

void WINAPI ReadDataFromFile(unsigned char* data,int num,int *count,int *w,int* h,int*x,int*y,int*black,int*BeginAddress,int*len);

void WINAPI ReadPicFromBytes(unsigned char *data ,int Size,unsigned int* Picdata);