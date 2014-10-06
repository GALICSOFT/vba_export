//
//
//
//
////////////////////////////////////////////////////////////////////////////////

#if defined(_MSC_VER) && _MSC_VER >= 1200
  #pragma warning( disable : 4786)
  #pragma warning( disable : 4996)
#endif


#include <stdio.h>
#include <stdlib.h>
#include <string.h>
#include <io.h>
#include <direct.h>
#include <errno.h>



enum
{
	DF_DATA_NONE	= 0,
	DF_DATA_INT		= 1,
	DF_DATA_DOUBLE	= 2,
	DF_DATA_STRING	= 3,
	DF_DATA_FUNCTION= 4,

	DF_VERSION_SIZE	= 16,
	DF_OFFSET_INFO	= 32,
};

struct Tschm
{
    int		t;
    char	n[28];
};

struct Tcell
{
	short	t;			// type
	short	l;			// string length

	union
	{
		int		n;		// int data
		double	d;		// double data
		char*	s;		// string or function data
	};
};


char		m_sVer[DF_VERSION_SIZE]={0};

short		m_infoRow	= 0;
short		m_infoCol	= 0;
Tschm*		m_infoSch	= NULL;
Tcell*		m_infoCel	= NULL;

short		m_dataRow	= 0;
short		m_dataCol	= 0;
Tschm*		m_dataSch	= NULL;
Tcell*		m_dataCel	= NULL;

int  ReadTbl(char* sFile);
void ReadByType(Tcell* pCel, int nT, FILE* fp);
int  Confirm(char* sFile);
void Destroy();




int main(int argc, char* argv[])
{
	int hr = 0;

	char curdir[_MAX_PATH]={0};

	char drive[_MAX_DRIVE]={0};
	char fdir [_MAX_DIR  ]={0};
	char fname[_MAX_FNAME]={0};
	char fext [_MAX_EXT  ]={0};

	_getcwd(curdir, _MAX_PATH);

	if(1 == argc)
	{
		char  fbuf [_MAX_PATH ]={0};
		char  fsrc [_MAX_PATH ]={0};
		char  ftxt [_MAX_PATH ]={0};


		sprintf(fbuf, "%s/*.bin", curdir);

		_finddata_t fd;
		long handle;
		int result=1;

		handle=_findfirst(fbuf, &fd);
		if (handle == -1)
			return 0;

		while (result != -1)
		{
			sprintf(fsrc, "%s/%s", curdir,	fd.name);

			_splitpath(fsrc, drive, fdir, fname, fext );
			sprintf(ftxt, "%s%s%s%s", drive, fdir, fname, ".txt");

			hr = ReadTbl(fsrc);
			hr = Confirm(ftxt);
			Destroy();

			result=_findnext(handle,&fd);
		}
		_findclose(handle);

		return 0;
	}

	char  fsrc [_MAX_PATH ]={0};
	char  ftxt [_MAX_PATH ]={0};

	strcpy(fsrc, argv[2]);

	_splitpath(fsrc, drive, fdir, fname, fext );
	sprintf(ftxt, "%s%s%s%s", drive, fdir, fname, ".txt");

	hr = ReadTbl(fsrc);
	hr = Confirm(ftxt);
	Destroy();

	return 0;
}


int ReadTbl(char* sFile)
{ 
	FILE*	fp = NULL;
	int		i=0, j=0, n= 0;
	int		nT =0;
	Tcell*	pCel = NULL;
	
	fp = fopen(sFile, "rb");
	if(!fp)
		return -1;


	fread(m_sVer, 1, DF_VERSION_SIZE, fp);

	m_infoRow = 1;
	fread(&m_infoCol, 2, 1, fp);
	fread(&m_dataRow, 2, 1, fp);
	fread(&m_dataCol, 2, 1, fp);

	if (0 >= m_infoRow ||0 >= m_infoCol ||
		0 >= m_dataRow ||0 >= m_dataCol)
	{
		fclose(fp);
		return -1;
	}

	m_infoSch = (Tschm*)malloc(sizeof(Tschm) * m_infoCol);
	m_infoCel = (Tcell*)malloc(sizeof(Tcell) * m_infoCol * m_infoRow);
	memset(m_infoSch, 0, sizeof(Tschm) * m_infoCol);
	memset(m_infoCel, 0, sizeof(Tcell) * m_infoCol * m_infoRow);


	m_dataSch = (Tschm*)malloc(sizeof(Tschm) * m_dataCol);
	m_dataCel = (Tcell*)malloc(sizeof(Tcell) * m_dataCol * m_dataRow);
	memset(m_dataSch, 0, sizeof(Tschm) * m_dataCol);
	memset(m_dataCel, 0, sizeof(Tcell) * m_dataCol * m_dataRow);

	fseek(fp, DF_OFFSET_INFO, SEEK_SET);

	// read the schema
	fread(m_infoSch, sizeof(Tschm), m_infoCol, fp);
	fread(m_dataSch, sizeof(Tschm), m_dataCol, fp);

	// read info
	for(j=0; j<m_infoCol; ++j)
	{
		nT   =  m_infoSch[j].t;
		pCel = &m_infoCel[j];

		ReadByType(pCel, nT, fp);
	}

	// read data
	for(i=0; i<m_dataRow; ++i)
	{
		for(j=0; j<m_dataCol; ++j)
		{
			n	 =  i * m_dataCol + j;
			nT   =  m_dataSch[j].t;
			pCel = &m_dataCel[n];

			ReadByType(pCel, nT, fp);
		}
	}

	fclose(fp);

	return 0;
}


void ReadByType(Tcell* pCel, int nT, FILE* fp)
{
	// write type
	pCel->t = nT;

	// read integer
	if     (DF_DATA_INT == nT)
	{
		int v =0;
		fread(&v, sizeof(v), 1, fp);

		pCel->n = v;
	}
	else if(DF_DATA_DOUBLE == nT)
	{
		double v =0;
		fread(&v, sizeof(v), 1, fp);

		pCel->d = v;
	}
	else if(DF_DATA_FUNCTION == nT || DF_DATA_STRING == nT)
	{
		short v =0;
		char* s =NULL;
		int   l =0;
		fread(&v, 2, 1, fp);

		l = (int)((v + 8)/4);
		l *=4;
		s = (char*)malloc(l);
		memset(s, 0, l);

		fread( s, 1, v, fp);

		pCel->l = l;
		pCel->s = s;
	}

}


int Confirm(char* sFile)
{ 
	FILE*	fp = NULL;
	int		i=0, j=0, n=0, nT;
	Tcell*	pCel = NULL;
	
	fp = fopen(sFile, "wt");
	if(!fp)
		return -1;


	fprintf(fp, "\n// Confirm read db\n\n");

	fprintf(fp, "ver: %d\n", m_sVer);
	fprintf(fp, "info: %d, %d\n", m_infoRow, m_infoCol);
	fprintf(fp, "data: %d, %d\n\n", m_dataRow, m_dataCol);

	// write info
	for(j=0; j<m_infoCol; ++j)
	{
		nT   =  m_infoSch[j].t;
		pCel = &m_infoCel[j];

		fprintf(fp, "%-24s %d ", m_infoSch[j].n, m_infoSch[j].t);
	
		if     (DF_DATA_INT == nT)
			fprintf(fp, ": %d\n", pCel->n);

		else if(DF_DATA_DOUBLE == nT)
			fprintf(fp, ": %6.4f\n", pCel->d);

		else if(DF_DATA_FUNCTION == nT)
			fprintf(fp, ": %s\n", pCel->s);

		else if(DF_DATA_STRING == nT)
			fprintf(fp, ": %s\n", pCel->s);

	}

	fprintf(fp, "\n\n");

	for(j=0; j<m_dataCol; ++j)
	{
		fprintf(fp, "[%s:%d] ", m_dataSch[j].n, m_dataSch[j].t);
	}

	fprintf(fp, "\n\n");


	// write data
	for(i=0; i<m_dataRow; ++i)
	{
		fprintf(fp, "[%3d] ", i);

		for(j=0; j<m_dataCol; ++j)
		{
			n	 =  i * m_dataCol + j;
			nT   =  m_dataSch[j].t;
			pCel = &m_dataCel[n];

			pCel->t = nT;

			if     (DF_DATA_INT == nT)
				fprintf(fp, "%4d", pCel->n);

			else if(DF_DATA_DOUBLE == nT)
				fprintf(fp, "%6.4f", pCel->d);

			else if(DF_DATA_FUNCTION == nT)
				fprintf(fp, "%-12s", pCel->s);

			else if(DF_DATA_STRING == nT)
				fprintf(fp, "%-12s", pCel->s);


			if(j != m_dataCol-1)
				fprintf(fp, ", ");
		}

		fprintf(fp, "\n");
	}

	fclose(fp);

	return 0;
}


void Destroy()
{
	int		i=0, j=0, n= 0;
	int		nT =0;
	Tcell*	pCel = NULL;

	if(NULL == m_infoSch)
		return;

	// release the info
	for(j=0; j<m_infoCol; ++j)
	{
		nT   =  m_infoSch[j].t;
		pCel = &m_infoCel[j];

		if(DF_DATA_FUNCTION == nT || DF_DATA_STRING == nT)
		{
			free(pCel->s);
			pCel->s = NULL;
		}
	}

	// release the data
	for(i=0; i<m_dataRow; ++i)
	{
		for(j=0; j<m_dataCol; ++j)
		{
			n	 =  i * m_dataCol + j;
			nT   =  m_dataSch[j].t;
			pCel = &m_dataCel[n];

			if(DF_DATA_FUNCTION == nT || DF_DATA_STRING == nT)
			{
				free(pCel->s);
				pCel->s = NULL;
			}
		}
	}

	free(m_infoSch);	m_infoSch = NULL;
	free(m_infoCel);	m_infoCel = NULL;
					
	free(m_dataSch);	m_dataSch = NULL;
	free(m_dataCel);	m_dataCel = NULL;
}


