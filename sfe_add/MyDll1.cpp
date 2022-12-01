#include "MyDll1.h"
#include "SDL_video.h"
#include "SDL_RWops.h"
//定义一个函数指针
 typedef void (  * TULIPFUNC )(void);  
 #pragma comment(lib,"sdl.lib")

 




int squ(int a)
{
	return a*a;	
}

unsigned short WINAPI INTrol(unsigned short a, unsigned short n)
{
	unsigned short b;
	__asm
	{
		mov ax,a
		mov cx,n
		rol ax,cl
		mov b,ax
	}
	return b;
}

unsigned short WINAPI INTror(unsigned short a,unsigned short n)
{
	unsigned short b;
	__asm
	{
		mov ax,a
		mov cx,n
		ror ax,cl
		mov b,ax
	}
	return b;
}


unsigned char WINAPI BYTEror(unsigned char a,unsigned char n)
{
	unsigned char b;
	__asm
	{
		mov al,a
		mov cl,n
		rol al,cl
		mov b,al
	}
	return b;
}

unsigned char WINAPI BYTErol(unsigned char a, unsigned char n)
{
	unsigned char b;
	__asm
	{
		mov al,a
		mov cl,n
		ror al,cl
		mov b,al
	}
	return b;
}

void WINAPI getbyteItem(unsigned char *a , int LENGTH,short tkey )
{
	int tmpInt; 
	for (int i=0;i<=LENGTH;i+=2)
	{
		tmpInt = (*(a+i+1) << 8) | *(a+i);
        tmpInt = tmpInt ^ tkey;
        if (tmpInt < 0)
       { 
			tmpInt = tmpInt + 65536;
		}        
		
		*(a + i) =  tmpInt & 255;
        *(a + i + 1) = tmpInt>>8;
    } 
}

void WINAPI convertCOLOR(unsigned int *mcolor_RGB , unsigned int *data,int WW,int HH,unsigned int MaskColor1)
{
int i,c;
unsigned int rr,gg,bb;
unsigned int rc[256],gc[256],bc[256];
int vmin,v,nn;
    //Me.MousePointer = vbHourglass
	//FILE *fp=fopen("debug.txt","w+");
    for(i=0;i<256;i++){
		rc[i]=*(mcolor_RGB+i) >>16 & 255 ;
		gc[i]=*(mcolor_RGB+i) >> 8 & 255;
		bc[i]=*(mcolor_RGB+i)      & 255;
	//	fprintf(fp,"RGB[%d]=%d\n",i,*(mcolor_RGB+i));

	}
//	fprintf(fp,"MaskColor1=%d\n",MaskColor1);
	for(i=0;i<WW*HH;i++){
		if(*(data + i) != MaskColor1){
//			fprintf(fp,"before data[%d]=%d\n",i,*(data + i));
			vmin = 100000;
			rr = (*(data + i))	& 255;
			gg = (*(data + i)) >> 8 & 255;
			bb = (*(data + i)) >>16 & 255;
			for(c=0;c<256;c++){
				v = squ(rc[c] - rr) + squ(bb - bc[c]) + squ(gg - gc[c]);
				if(v < vmin){
					vmin = v;
					nn = c;
				}
                }
		*(data + i) = bc[nn]*65536l+gc[nn]*256+rc[nn] ;
		//fprintf(fp,"after data[%d]=%d\n",i,*(data + i));
            }
	}
//	fclose(fp);
}


unsigned char WINAPI get256(unsigned int *mcolor_RGB , unsigned int d)
{
	int rr,gg,bb;
	int r2,g2,b2;
	int i;

	b2=d>>16 & 255;
	g2=d>> 8 & 255;
	r2=d     & 255;
	for(i=0;i<256;i++){
		rr=*(mcolor_RGB+i)>>16 & 255;
		gg=*(mcolor_RGB+i)>> 8 & 255;
		bb=*(mcolor_RGB+i)	   & 255;
		if ((r2==rr) && (g2==gg) && (b2==bb)){ 
			return i;
			break;
		}
	}
    return i;
}

void WINAPI convertCOLOR2(unsigned int *mcolor_RGB , unsigned int *data,int WW,int HH,unsigned int MaskColor1)
{
	int i,c;
	unsigned int rr,gg,bb;
	double yy,uu,vv;
	unsigned int rc[256],gc[256],bc[256];
	double yc[256],uc[256],vc[256];
	int vmin,v,nn;

    //Me.MousePointer = vbHourglass
    for(i=0;i<256;i++){
		rc[i]=*(mcolor_RGB+i) >>16 & 255 ;
		gc[i]=*(mcolor_RGB+i) >>8 & 255;
		bc[i]=*(mcolor_RGB+i) & 255;
		yc[i]= 0.299 * rc[i] + 0.587 * gc[i] + 0.114 * bc[i];
		uc[i] = -0.1687 * rc[i] - 0.3313 * gc[i] + 0.5 * bc[i] + 128;
		vc[i] = 0.5 * rc[i] - 0.4187 * gc[i] - 0.0813 * bc[i] + 128;
	}
    
	for(i=0;i<WW*HH;i++){
			if(*(data + i) != MaskColor1){
				vmin = 100000;
				rr = *(data + i) & 255;
				gg = *(data + i) >>8 & 255;
				bb = *(data + i) >>16 & 255;
                yy = 0.299 * rr + 0.587 * gg + 0.114 * bb;
                uu = -0.1687 * rr - 0.3313 * gg + 0.5 * bb + 128;
                vv = 0.5 * rr - 0.4187 * gg - 0.0813 * bb + 128;
                
				for(c=0;c<256;c++){
					v = 2*squ(rc[c] - rr) + squ(bb - bc[c]) + squ(gg - gc[c]);
					if(v < vmin){
						vmin = v;
						nn = c;
					}
                }
				//*(data + i) = (unsigned int)((rc[nn]<<16)|(gc[nn]<<8)|bc[nn] );
				*(data + i) = bc[nn]*65536l+gc[nn]*256+rc[nn] ;
			//	*(data+i)=0;
			}
	  }
}

void WINAPI drawPic(short MapNum, unsigned int X,unsigned int Y,int datalong,unsigned char *data,unsigned int CenterX,unsigned int CenterY,HDC hdc)
{
	int i=0;
	char PicWidth;
	int piclong,MaskLong;
	int dx;

    //If MapNum <= 0 Or MapNum >= MMAPNUM Then Exit Sub
    //X = X - MmapPic(MapNum).X
    //Y = Y - MmapPic(MapNum).Y
    dx = X;

    while(i < datalong){
        PicWidth = *(data+i);
        i = i++;
        piclong = PicWidth + i;
        while(i < piclong){
            X = X + *(data+i);
            i++;
            MaskLong = i + 1 + *(data+i);
            i++;
            while(i < MaskLong){
                if((X < CenterX * 2)&&(Y < CenterY * 2)&&(X >= 0)&&(Y >= 0)){
                    SetPixel(hdc, X, Y,*(data+i));
                }
                    X++;i++;
            }
        }
        Y++;
        X = dx;
    }
}

void WINAPI cpymem(int *dest,int value)
{
	*dest=value;
}

void WINAPI ReadDataFromFile(unsigned char* data,int num,int *count,int *w,int* h,int*x,int*y,int*black,int*BeginAddress,int*len)
{ 
	int* Filedata = (int*)data;
	*count = *Filedata;
	*len = *(Filedata + num + 1);

	if(0 == num)
		*BeginAddress = (*count + 1) * 4;
	else 
		*BeginAddress = *(Filedata + num); 

	*len = *len - *BeginAddress - 12;
	
    Filedata = (int*)(data + *BeginAddress);

	*x = *(Filedata);
	*y = *(Filedata + 1);
	*black = *(Filedata + 2);
	*BeginAddress=*BeginAddress+12;

	unsigned char temp ;
	*w = 0;
	*h = 0; int i=0 ;
	for (  i=0 ; i<4; i++)
	{
		temp =*(data + *BeginAddress + 16+i);
		*w=(*w<<8)|temp;
	}
	for (i=0; i<4; i++)
	{
		temp =*(data + *BeginAddress + 20+i);
		*h=(*h<<8)|temp;
	}
}

void WINAPI ReadPicFromBytes(unsigned char *data ,int Size,unsigned int* Picdata)
{ 

	data=new unsigned char[Size]; 
	SDL_RWops *Rwops=SDL_RWFromMem(data, Size);
	SDL_Surface* Pic=SDL_LoadBMP_RW(Rwops, 1); 
	int w = (*Pic).w;
	int h = (*Pic).h; 
    short bpp = (*(*Pic).format).BytesPerPixel; 
    Picdata=new unsigned int[w * h];

    for (int i1 = 0 ; i1 < h ; i1++)
	{ 
		for (int i2 = 0 ; i2 < w ; i2++)
		{
			*(Picdata+i1 * w + i2) = *((unsigned int*)((*Pic).pixels) + i1 * (*Pic).pitch + i2 * bpp);
		}
	}
	SDL_FreeSurface(Pic); 

}
