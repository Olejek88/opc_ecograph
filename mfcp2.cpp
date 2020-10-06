#define _WIN32_DCOM				// Enables DCOM extensions
#define INITGUID				// Initialize OLE constants

#include <stdio.h>
#include <math.h>				// some mathematical function
#include "afxcmn.h"				// MFC function such as CString,....
#include "ecobus.h"				// no comments
#include "ecodrv.h"				// no comments
#include "unilog.h"				// universal utilites for creating log-files
#include <locale.h>				// set russian codepage
#include <opcda.h>				// basic function for OPC:DA
#include "lightopc.h"			// light OPC library header file
#include "serialport.h"			// function for work with Serial Port

#define ECL_SID "opc.ecograph"	// identificator of OPC server
#define TAGNAME_LEN 64			// tag lenght

// OPC vendor info (Major/Minor/Build/Reserv)
static const loVendorInfo vendor = {0,11,0,0,"EcoGraph OPC Server" };
loService *my_service;			// name of light OPC Service
static void a_server_finished(void*, loService*, loClient*);// OnServer finish his work
static int OPCstatus=OPC_STATUS_RUNNING;					// status of OPC server
//---------------------------------------------------------------------------------
char* ReadParam (char *SectionName,char *Value);// read parametr from .ini file
int tTotal=ECOCOM_NUM_MAX;					// total quantity of tags
int com_num	=1;					// COM-port number
static Eco *devp;				// pointer to tag structure
static loTagId ti[ECOCOM_NUM_MAX];			// Tag id
static loTagValue tv[ECOCOM_NUM_MAX];		// Tag value
static char *tn[ECOCOM_NUM_MAX];			// Tag name
unsigned int timetg[ECOCOM_NUM_MAX];		// command priority timer
BOOL active_channal[11];		// active channal (1-7 analog, 8-11 digital)
BOOL digital=TRUE;
BOOL analog=TRUE;
BOOL integrator=FALSE;			// server options
BOOL intellect=TRUE;			// intellegence server
BOOL rvaltags=TRUE;				// show rVal tags
char argv0[FILENAME_MAX + 32];	// lenght of command line (file+path (260+32))
unilog *logg=NULL;				// new structure of unilog
static OPCcfg cfg;				// new structure of cfg
CSerialPort port;				// com-port
void addCommToPoll();			// add commands to read list
UINT ScanBus();					// bus scanned programm
UINT PollDevice(int device);	// polling single device
UINT InitDriver();				// func of initialising port and creating service
UINT DestroyDriver();			// function of detroying driver and service
HRESULT WriteDevice(int device,const unsigned cmdnum,LPSTR data);	// write tag to device
FILE *CfgFile;						// pointer to .ini file
static CRITICAL_SECTION lk_values;	// protects ti[] from simultaneous access 
CRITICAL_SECTION drv_access;
BOOL checkprop (int i);			// check tag property

static int mymain(HINSTANCE hInstance, int argc, char *argv[]);
static int show_error(LPCTSTR msg);		// just show messagebox with error
static int show_msg(LPCTSTR msg);		// just show messagebox with message
static void poll_device(void);			// function polling device
void dataToTag(int device);				// copy data buffer to tag

// {A96598F3-9B45-49f9-A515-1607E2D7F332}
DEFINE_GUID(GID_eOPCserverDll, 
0xa96598f3, 0x9b45, 0x49f9, 0xa5, 0x15, 0x16, 0x7, 0xe2, 0xd7, 0xf3, 0x32);
// {36D92C9E-AA9E-46f3-88F6-E7B98A34B8B2}
DEFINE_GUID(GID_eOPCserverExe, 
0x36d92c9e, 0xaa9e, 0x46f3, 0x88, 0xf6, 0xe7, 0xb9, 0x8a, 0x34, 0xb8, 0xb2);
//---------------------------------------------------------------------------------
void timetToFileTime( time_t t, LPFILETIME pft )
{
    LONGLONG ll = Int32x32To64(t, 10000000) + 116444736000000000;
    pft->dwLowDateTime = (DWORD) ll;
    pft->dwHighDateTime =  (unsigned long)(ll >>32);	//(unsigned long)
}

char *absPath(char *fileName)					// return abs path of file
{
  static char path[sizeof(argv0)]="\0";			// path - massive of comline
  char *cp;										  
  if(*path=='\0') strcpy(path, argv0);			
  if(NULL==(cp=strrchr(path,'\\'))) cp=path; else cp++;
  cp=strcpy(cp,fileName);
  return path;
}

inline void init_common(void)		// create log-file
{
  logg = unilog_Create(ECL_SID, absPath(LOG_FNAME), NULL, 0, ll_DEBUG); // level [ll_FATAL...ll_DEBUG] 
  UL_INFO((LOGID, "Star1t"));		// write in log
}

inline void cleanup_common(void)	// delete log-file
{
  UL_INFO((LOGID, "Finish"));
  unilog_Delete(logg); logg = NULL; // + logs was not currently
}

int show_error(LPCTSTR msg)			// just show messagebox with error
{
  ::MessageBox(NULL, msg, ECL_SID, MB_ICONSTOP|MB_OK);
  return 1;
}
int show_msg(LPCTSTR msg)			// just show messagebox with message
{
  ::MessageBox(NULL, msg, ECL_SID, MB_OK);
  return 1;
}
inline void cleanup_all(DWORD objid)
{ // Informs OLE that a class object, previously registered is no longer available for use
  if (FAILED(CoRevokeClassObject(objid)))  UL_WARNING((LOGID, "CoRevokeClassObject() failed..."));
  DestroyDriver();					// close port and destroy driver
  CoUninitialize();					// Closes the COM library on the current thread
  cleanup_common();					// delete log-file
  //fclose(CfgFile);					// close .ini file
}
//---------------------------------------------------------------------------------
class myClassFactory: public IClassFactory	
{
public:
  LONG RefCount;			// reference counter
  LONG server_count;		// server counter
  CRITICAL_SECTION lk;		// protect RefCount 
  
  myClassFactory(): RefCount(0), server_count(0)	// when creating interface zero all counter
  {  InitializeCriticalSection(&lk);	  }
  ~myClassFactory()
  {  DeleteCriticalSection(&lk); }

// IUnknown realisation
  STDMETHODIMP QueryInterface(REFIID, LPVOID*);
  STDMETHODIMP_(ULONG) AddRef( void);
  STDMETHODIMP_(ULONG) Release( void);
// IClassFactory realisation
  STDMETHODIMP CreateInstance(LPUNKNOWN, REFIID, LPVOID*);
  STDMETHODIMP LockServer(BOOL);
  
  inline LONG getRefCount(void)
  {
    LONG rc;
    EnterCriticalSection(&lk);		// attempt recieve variable whom may be used by another threads
    rc = RefCount;					// rc = client counter
    LeaveCriticalSection(&lk);                                    
    return rc;
  }

  inline int in_use(void)
  {
    int rv;
    EnterCriticalSection(&lk);
    rv = RefCount | server_count;
    LeaveCriticalSection(&lk);
    return rv;
  }

  inline void serverAdd(void)
  {
    InterlockedIncrement(&server_count);	// increment server counter
  }

  inline void serverRemove(void)
  {
    InterlockedDecrement(&server_count);	// decrement server counter
  }
};
//----- IUnknown -------------------------------------------------------------------------
STDMETHODIMP myClassFactory::QueryInterface(REFIID iid, LPVOID* ppInterface)
{
  if (ppInterface == NULL) return E_INVALIDARG;	// pointer to interface missed (NULL)

  if (iid == IID_IUnknown || iid == IID_IClassFactory)	// legal IID
    {
      UL_DEBUG((LOGID, "myClassFactory::QueryInterface() Ok"));	// write to log
      *ppInterface = this;		// interface succesfully returned
      AddRef();					// adding reference to interface
      return S_OK;				// return succesfully
    }
  UL_ERROR((LOGID, "myClassFactory::QueryInterface() Failed"));
  *ppInterface = NULL;			// no interface returned
  return E_NOINTERFACE;			// error = No Interface
}

STDMETHODIMP_(ULONG) myClassFactory::AddRef(void)	// new client was connected
{
  ULONG rv;
  EnterCriticalSection(&lk);
  rv = (ULONG)++RefCount;							// increment counter of client
  LeaveCriticalSection(&lk);                                    
  UL_DEBUG((LOGID, "myClassFactory::AddRef(%ld)", rv));	// write to log number
  return rv;
}

STDMETHODIMP_(ULONG) myClassFactory::Release(void)	// client has been disconnected
{
  ULONG rv;
  EnterCriticalSection(&lk);
  rv = (ULONG)--RefCount;							// decrement client counter
  LeaveCriticalSection(&lk);
  UL_DEBUG((LOGID, "myClassFactory::Release(%d)", rv));	// write to log number
  return rv;
}

//----- IClassFactory ----------------------------------------------------------------------
STDMETHODIMP myClassFactory::LockServer(BOOL fLock)
{
  if (fLock)	AddRef();
  else		    Release();
  UL_DEBUG((LOGID, "myClassFactory::LockServer(%d)", fLock)); 
  return S_OK;
}

STDMETHODIMP myClassFactory::CreateInstance(LPUNKNOWN pUnkOuter, REFIID riid, LPVOID* ppvObject)
{
  if (pUnkOuter != NULL)
    return CLASS_E_NOAGGREGATION; // Aggregation is not supported by this code

  IUnknown *server = 0;

  AddRef(); // for a_server_finished()
  if (loClientCreate(my_service, (loClient**)&server, 0, &vendor, a_server_finished, this))
    {
      UL_DEBUG((LOGID, "myClassFactory::loCreateClient() failed"));
      Release();
      return E_OUTOFMEMORY;	
    }
  serverAdd();

  HRESULT hr = server->QueryInterface(riid, ppvObject);
  if (FAILED(hr))
      UL_DEBUG((LOGID, "myClassFactory::loClient QueryInterface() failed"));
  else
    {
      loSetState(my_service, (loClient*)server, loOP_OPERATE, OPCstatus, 0);
      UL_DEBUG((LOGID, "myClassFactory::server_count = %ld", server_count));
    }
  server->Release();
  return hr;
}

static myClassFactory my_CF;
static void a_server_finished(void *a, loService *b, loClient *c)
{
  my_CF.serverRemove();						
  if (a) ((myClassFactory*)a)->Release();
  UL_DEBUG((LOGID, "a_server_finished(%lu)", my_CF.server_count));
}

//---------------------------------------------------------------------------------
int APIENTRY WinMain(HINSTANCE hInstance,HINSTANCE hPrevInstance,LPSTR lpCmdLine,int nCmdShow)
{
  static char *argv[3] = { "dummy.exe", NULL, NULL };	// defaults arguments
  argv[1] = lpCmdLine;									// comandline - progs keys
  return mymain(hInstance, 2, argv);
}

int main(int argc, char *argv[])
{
  return mymain(GetModuleHandle(NULL), argc, argv);
}

int mymain(HINSTANCE hInstance, int argc, char *argv[]) 
{
const char eClsidName [] = ECL_SID;				// desription 
const char eProgID [] = ECL_SID;				// name
DWORD objid;									// fully qualified path for the specified module
char *cp;
objid=::GetModuleFileName(NULL, argv0, sizeof(argv0));	// function retrieves the fully qualified path for the specified module
if(objid==0 || objid+50 > sizeof(argv0)) return 0;		// not in border

init_common();									// create log-file
if(NULL==(cp = setlocale(LC_ALL, ".1251")))		// sets all categories, returning only the string cp-1251
	{ 
	UL_ERROR((LOGID, "setlocale() - Can't set 1251 code page"));	// in bad case write error in log
	cleanup_common();							// delete log-file
    return 0;
	}
cp = argv[1];		
if(cp)						// check keys of command line 
	{
    int finish = 1;			// flag of comlection
    if (strstr(cp, "/r"))	//	attempt registred server
		{
	     if (loServerRegister(&GID_eOPCserverExe, eProgID, eClsidName, argv0, 0)) 
			{ show_error("Registration Failed");
			  UL_ERROR((LOGID, "Registration <%s> <%s> Failed", eProgID, argv0));  } 
		 else 
			{ show_msg("EcoGraph OPC Registration Ok");
			 UL_INFO((LOGID, "Registration <%s> <%s> Ok", eProgID, argv0));		}
		} 
	else 
		if (strstr(cp, "/u")) 
			{
			 if (loServerUnregister(&GID_eOPCserverExe, eClsidName)) 
				{ show_error("UnRegistration Failed");
				 UL_ERROR((LOGID, "UnReg <%s> <%s> Failed", eClsidName, argv0)); } 
			 else 
				{ show_msg("EcoGraph OPC Server Unregistered");
				 UL_INFO((LOGID, "UnReg <%s> <%s> Ok", eClsidName, argv0));		}
			} 
		else  // only /r and /u options
			if (strstr(cp, "/?")) 
				 show_msg("Use: \nKey /r to register server.\nKey /u to unregister server.\nKey /? to show this help.");
				 else
					{
					 UL_WARNING((LOGID, "Ignore unknown option <%s>", cp));
					 finish = 0;		// nehren delat
					}
		if (finish) {      cleanup_common();      return 0;    } 
	}
if ((CfgFile = fopen(CFG_FILE, "r+")) == NULL )	
	{
	 show_error("Error open .ini file");
	 UL_ERROR((LOGID, "Error open .ini file"));	// in bad case write error in log
	 return 0;
	}
if (FAILED(CoInitializeEx(NULL, COINIT_MULTITHREADED))) {	// Initializes the COM library for use by the calling thread
    UL_ERROR((LOGID, "CoInitializeEx() failed. Exiting..."));
    cleanup_common();	// close log-file
    return 0;
  }
UL_INFO((LOGID, "CoInitializeEx() Ok...."));	// write to log
devp = new Eco();
if (InitDriver()) {		// open and set com-port
    CoUninitialize();	// Closes the COM library on the current thread
    cleanup_common();	// close log-file
    return 0;
  }
UL_INFO((LOGID, "InitDriver() Ok...."));	// write to log

if (FAILED(CoRegisterClassObject(GID_eOPCserverExe, &my_CF, 
				   CLSCTX_LOCAL_SERVER|CLSCTX_REMOTE_SERVER|CLSCTX_INPROC_SERVER, 
				   REGCLS_MULTIPLEUSE, &objid)))
    { UL_ERROR((LOGID, "CoRegisterClassObject() failed. Exiting..."));
      cleanup_all(objid);		// close comport and unload all librares
      return 0; }
UL_INFO((LOGID, "CoRegisterClassObject() Ok...."));	// write to log

Sleep(1000);			
my_CF.Release();		// avoid locking by CoRegisterClassObject() 

if (OPCstatus!=OPC_STATUS_RUNNING)	// ???? maybe Status changed and OPC not currently running??
	{	while(my_CF.in_use())      Sleep(1000);	// wait
		cleanup_all(objid);
		return 0;	}
addCommToPoll();		// check tags list and list who need
while(my_CF.in_use())	// while server created or client connected
    poll_device();		// polling devices else do nothing (and be nothing)

cleanup_all(objid);		// destroy himself
return 0;
}
//-------------------------------------------------------------------
void poll_device(void)
{
  FILETIME ft;
  int devi, i,j, ecode;
  for (devi = 0, i = 0; devi < devp->idnum; devi++) 
	{
	  UL_DEBUG((LOGID, "Driver poll <%d>", devp->ids[i]));
	  ecode=PollDevice(devp->ids[i]);
	  if (ecode)
		{
		 dataToTag (devp->ids[i]);
		 UL_DEBUG((LOGID, "Copy data to tag success"));
		 time(&devp->mtime);
		 timetToFileTime(devp->mtime, &ft);
		}
  	  else
		{
		 for (j=0; j<devp->cv_size; j++)
		    devp->cv_status[j] = FALSE;
		 GetSystemTimeAsFileTime(&ft);
		}	 
	 EnterCriticalSection(&lk_values);
	 for (int ci = 0; ci < devp->cv_size; ci++, i++) 
		{
	      WCHAR buf[64];
		  LCID lcid = MAKELCID(0x0409, SORT_DEFAULT); // This macro creates a locale identifier from a language identifier. Specifies how dates, times, and currencies are formatted
 	 	  MultiByteToWideChar(CP_ACP,	 // ANSI code page
									  0, // flags
						   devp->cv[ci], // points to the character string to be converted
				 strlen(devp->cv[ci])+1, // size in bytes of the string pointed to 
									buf, // Points to a buffer that receives the translated string
			  sizeof(buf)/sizeof(buf[0])); // function maps a character string to a wide-character (Unicode) string
// 	  	  UL_DEBUG((LOGID, "set tag <i,status,value>=<%d,%d,%s,%s>",i, devp->cv_status[ci], devp->cv[ci],buf));	
		  if (devp->cv_status[ci]) 
			{	
			  VARTYPE tvVt = tv[i].tvValue.vt;
			  VariantClear(&tv[i].tvValue);			  
			  switch (tvVt) 
					{
				 	  case VT_UI2:
								UINT vi1;
								vi1 = *devp->cv[ci];
								vi1=atoi(devp->cv[ci]);
								V_UI2(&tv[i].tvValue) = vi1;
								break;
				 	  case VT_I2:
								short vi2;
								vi2 = *devp->cv[ci];
								VarI2FromStr(buf, lcid, 0, &vi2);
								V_I2(&tv[i].tvValue) = vi2;
								break;
					  case VT_I4:
								ULONG vi4;								
								vi4=atoi(devp->cv[ci]);
								V_I4(&tv[i].tvValue) = vi4;
								break;
					  case VT_DATE:
						  	    DATE date; 
							    VarDateFromStr(buf, lcid, 0, &date);
								V_DATE(&tv[i].tvValue) = date;
								break;
					  case VT_BSTR:
					  default:
								V_BSTR(&tv[i].tvValue) = SysAllocString(buf);
					}
				V_VT(&tv[i].tvValue) = tvVt;
		  	    tv[i].tvState.tsQuality = OPC_QUALITY_GOOD;
			}
			if (ecode == 0)
				 tv[i].tvState.tsQuality = OPC_QUALITY_UNCERTAIN;
			if (ecode == 2)
				 tv[i].tvState.tsQuality = OPC_QUALITY_DEVICE_FAILURE;
			tv[i].tvState.tsTime = ft;
		}
	 loCacheUpdate(my_service, tTotal, tv, 0);
     LeaveCriticalSection(&lk_values);
	}
  Sleep(1000);
}
//-------------------------------------------------------------------
loTrid ReadTags(const loCaller *, unsigned  count, loTagPair taglist[],
		VARIANT   values[],	WORD      qualities[],	FILETIME  stamps[],
		HRESULT   errs[],	HRESULT  *master_err,	HRESULT  *master_qual,
		const VARTYPE vtype[],	LCID lcid)
{  return loDR_STORED; }
//-------------------------------------------------------------------
int WriteTags(const loCaller *ca,
              unsigned count, loTagPair taglist[],
              VARIANT values[], HRESULT error[], HRESULT *master, LCID lcid)
{  
 unsigned ii,ci,devi; int i;
 char cmdData[50],buff[50],bufg[50];	// data to save massive
 char *ppbuf = cmdData;
 VARIANT v;				// input data - variant type
 char ldm;				
 struct lconv *lcp;		// Contains formatting rules for numeric values in different countries/regions
 lcp = localeconv();	// Gets detailed information on locale settings.	
 ldm = *(lcp->decimal_point);	// decimal point (i nah ona nujna?)
 VariantInit(&v);				// Init variant type
 UL_TRACE((LOGID, "WriteTags (%d) invoked", count));	
 EnterCriticalSection(&lk_values);	
 for(ii = 0; ii < count; ii++) 
	{
      HRESULT hr = 0;
	  loTagId clean = 0;
      cmdData[0] = '\0';
      cmdData[ECOCOM_DATALEN_MAX] = '\0';
      UL_TRACE((LOGID,  "WriteTags(Rt=%u Ti=%u)", taglist[ii].tpRt, taglist[ii].tpTi));	  
      i = (unsigned)taglist[ii].tpRt - 1;
      ci = i % devp->cv_size;
      devi = i / devp->cv_size;
      if (!taglist[ii].tpTi || !taglist[ii].tpRt || i >= tTotal)
			continue;
      VARTYPE tvVt = tv[i].tvValue.vt;
      hr = VariantChangeType(&v, &values[ii], 0, tvVt);
      if (hr == S_OK) 
		{
			switch (tvVt) 
				{	
				 case VT_UI2:
							 _snprintf(cmdData, ECOCOM_DATALEN_MAX, "%u",v.uiVal);
 							 sprintf(buff,"com%d/id%0.2d/Settings/Normal feed speed",com_num, devp->ids[devi], tn[i]);
							 sprintf(bufg,"com%d/id%0.2d/Settings/Active feed speed",com_num, devp->ids[devi], tn[i]);
							 if ((!strcmp(tn[i],buff))||(!strcmp(tn[i],bufg)))
								{
								 UINT speed=atoi(cmdData);
								 if (speed<2) { sprintf(cmdData,"00"); break;}
								 if (speed<8) { sprintf(cmdData,"01"); break;}	
								 if (speed<15) { sprintf(cmdData,"02"); break;}	
								 if (speed<25) { sprintf(cmdData,"03"); break;}	
								 if (speed<45) { sprintf(cmdData,"04"); break;}	
								 if (speed<90) { sprintf(cmdData,"05"); break;}	
								 if (speed<180) { sprintf(cmdData,"06"); break;}	
								 if (speed<270) { sprintf(cmdData,"07"); break;}	
								 if (speed<450) { sprintf(cmdData,"08"); break;}	
								 if (speed<800) { sprintf(cmdData,"09"); break;}	
								 if (speed>799) { sprintf(cmdData,"10"); break;}	
								}
							 break;
				 case VT_I2: _snprintf(cmdData, ECOCOM_DATALEN_MAX, "%d", v.iVal);
							 break; // Write formatted data to a string
				 case VT_R4: UL_TRACE((LOGID, "Number input (%f)",v.fltVal));
							 _snprintf(cmdData, ECOCOM_DATALEN_MAX, "%f", v.fltVal);
//							 ppbuf = cFloatCtoIEEE754 (v.fltVal);
							 strcpy(cmdData, ppbuf);
							 //if (ldm != '.' && (dm = strchr(cmdData, ldm)))  *dm = '.';
							 break;
				 case VT_UI1:_snprintf(cmdData, ECOCOM_DATALEN_MAX, "%c", v.bVal);
							 break;
				 case VT_BSTR:
							 sprintf(buff,"com%d/id%0.2d/Settings/Feed Unit",com_num, devp->ids[devi], tn[i]);				 
							 UL_DEBUG((LOGID, "buff %s / tn = %s",buff,tn[i]));
							 WideCharToMultiByte(CP_ACP,0,v.bstrVal,-1,cmdData,ECOCOM_DATALEN_MAX,NULL, NULL);							 
							 if (!strcmp(tn[i],buff))
								{
								 UL_DEBUG((LOGID, "inchi nah %s | cmdData %s", v.bstrVal,cmdData));
								 if ((!strcmp(cmdData,"mm/h"))||(!strcmp(cmdData,"mm"))||(!strcmp(cmdData,"00")))
										sprintf(cmdData,"00");
								 if ((!strcmp(cmdData,"in/h"))||(!strcmp(cmdData,"inch"))||(!strcmp(cmdData,"01")))
										sprintf(cmdData,"01");	
								}							 

							 break;
				 default:	  
							 WideCharToMultiByte(CP_ACP,0,v.bstrVal,-1,cmdData,ECOCOM_DATALEN_MAX,NULL, NULL);
				}
		 UL_TRACE((LOGID, "%!l WriteTags(Rt=%u Ti=%u %s %s)",hr, taglist[ii].tpRt, taglist[ii].tpTi, tn[i], cmdData));
 		 hr = WriteDevice(devp->ids[devi], devp->cv_cmdid[ci], cmdData);
	  }
       VariantClear(&v);
	   if (S_OK != hr) 
			{
			 *master = S_FALSE;
			 error[ii] = hr;
			 UL_TRACE((LOGID, "Write failed"));
			}
	   else
			UL_TRACE((LOGID, "Write success"));
       taglist[ii].tpTi = clean; 
  }
 LeaveCriticalSection(&lk_values); 
 return loDW_TOCACHE; 
}
//-------------------------------------------------------------------
void activation_monitor(const loCaller *ca, int count, loTagPair *til)
{}
//-------------------------------------------------------------------
int tcount=0;
unsigned char DeviceDataBuffer[16][ECOCOM_NUM_MAX][40];
char mBuf[ECOCOM_NUM_MAX][6];
//----------------------------------------------------------------------------
void addCommToPoll()
{
int	max_com=0,i,j,flajok;	
sprintf(mBuf[0],"VVVV");
for (i=0;i<sizeof(EcoCommU)/sizeof(EcoCom);i++)
{	  
  if ((EcoCommU[i].getCmd!="nein")&&(!strstr(EcoCommU[i].name,"(rVal)")))
	if (checkprop(i))
	 {
	 flajok=0;
	 for (j=0;j<=tcount;j++)
		 if (!strcmp (EcoCommU[i].getCmd,mBuf[j]))
			{ flajok=1; break;}
	 if (flajok==0) 
		{		 
		 sprintf(mBuf[tcount],"%s",EcoCommU[i].getCmd);
		 UL_DEBUG((LOGID, "Add command %s to poll. Priority %d. Total %d command added.",mBuf[tcount],EcoCommU[i].priority,tcount));
		 tcount++;
		}
	}
}}
//-----------------------------------------------------------------------------------
void dataToTag(int device)
{
int	max_com=0,i,k;
char *l;
unsigned int j;	
char buf[50];
char *pbuf=buf; 
*pbuf = '\0';
for (i=0; i<devp->cv_size; i++)			
	{
	 if (EcoCommU[i].scan_it)
	 {
	 for (j=0; j<EcoCommU[i].num; j++)
		{
		 for (k=0;k<tcount;k++)  
			{
			 if (!strcmp(EcoCommU[i].getCmd,mBuf[k])) 
				{ 
//				 UL_DEBUG((LOGID, "tag: %s / com: %s / sym: %d",EcoCommU[i].name,EcoCommU[i].getCmd,DeviceDataBuffer[device][k][j+1+EcoCommU[i].start])); 
				 break;
				}
			}
		 buf[j]=DeviceDataBuffer[device][k][j+1+EcoCommU[i].start];
		}
	 buf[j] = '\0';	
//---------------------------------------------------------------------------------------------
	 if (EcoCommU[i].name=="Analog input/1/Type") 
		 if (!strcmp(buf,"00"))	 active_channal[1]=FALSE;
		 else					 active_channal[1]=TRUE;
	 if (EcoCommU[i].name=="Analog input/2/Type") 
		 if (!strcmp(buf,"00"))	 active_channal[2]=FALSE;
		 else					 active_channal[2]=TRUE;
	 if (EcoCommU[i].name=="Analog input/3/Type") 
		 if (!strcmp(buf,"00"))	 active_channal[3]=FALSE;
		 else					 active_channal[3]=TRUE;
	 if (EcoCommU[i].name=="Analog input/4/Type") 
		 if (!strcmp(buf,"00"))	 active_channal[4]=FALSE;
		 else					 active_channal[4]=TRUE;
	 if (EcoCommU[i].name=="Analog input/5/Type") 
		 if (!strcmp(buf,"00"))	 active_channal[5]=FALSE;
		 else					 active_channal[5]=TRUE;
	 if (EcoCommU[i].name=="Analog input/6/Type") 
		 if (!strcmp(buf,"00"))	 active_channal[6]=FALSE;
		 else					 active_channal[6]=TRUE;
	 if (EcoCommU[i].name=="Digital input/1/Input function")
		 if (!strcmp(buf,"00"))	 active_channal[7]=FALSE;
		 else					 active_channal[7]=TRUE;
	 if (EcoCommU[i].name=="Digital input/2/Input function")
		 if (!strcmp(buf,"00"))	 active_channal[8]=FALSE;
		 else					 active_channal[8]=TRUE;
	 if (EcoCommU[i].name=="Digital input/3/Input function")
		 if (!strcmp(buf,"00"))	 active_channal[9]=FALSE;
		 else					 active_channal[9]=TRUE;
	 if (EcoCommU[i].name=="Digital input/4/Input function")
		 if (!strcmp(buf,"00"))	 active_channal[10]=FALSE;
		 else					 active_channal[10]=TRUE;
//---------------------------------------------------------------------------------------------
	 if (EcoCommU[i].name=="Settings/Feed Unit")
		  if (!strcmp(buf,"00"))	sprintf(buf,"mm/h");	else sprintf(buf,"in/h");
	 if ((EcoCommU[i].name=="Settings/Normal feed speed")||(EcoCommU[i].name=="Settings/Active feed speed"))
		{
		  if (!strcmp(buf,"00")) sprintf(buf,"0");
		  if (!strcmp(buf,"01")) sprintf(buf,"5");
		  if (!strcmp(buf,"02")) sprintf(buf,"10");
		  if (!strcmp(buf,"03")) sprintf(buf,"20");
		  if (!strcmp(buf,"04")) sprintf(buf,"30");
		  if (!strcmp(buf,"05")) sprintf(buf,"60");
		  if (!strcmp(buf,"06")) sprintf(buf,"120");
		  if (!strcmp(buf,"07")) sprintf(buf,"240");
		  if (!strcmp(buf,"08")) sprintf(buf,"300");
		  if (!strcmp(buf,"09")) sprintf(buf,"600");
		  if (!strcmp(buf,"10")) sprintf(buf,"1000");
		}
//---------------------------------------------------------------------------------------------
	  strcpy (devp->cv[i],buf);
	 }	 
	 else
		{
 		 for (k=0;k<tcount;k++)  
			 if (!strcmp(EcoCommU[i].getCmd,mBuf[k])) 
				 break;
	 	 for (j=0; j<EcoCommU[i].num; j++)
			 buf[j]=DeviceDataBuffer[device][k][j+1+EcoCommU[i].start]; 
		 buf[j] = '\0';

		 l=strstr(EcoCommU[i].name,"(rVal)");
		 if (l) // noscan and name with rVal
			{
				// parse getCmd and setCmd num raz
		//	 UL_DEBUG((LOGID, "getCmd: %s / setCmd: %s / raz: %d",EcoCommU[i].getCmd,EcoCommU[i].setCmd,EcoCommU[i].num));
			 char buff[550];
			 sprintf (buff,EcoCommU[i].name);
			 buff[l-EcoCommU[i].name-1]='\0';
//			 sprintf(bufg,"com%d/id%0.2d/%s",com_num, devp->ids[i], buff);
//			 UL_DEBUG((LOGID, "attempt to find: %s",buff));
			 int pos=0;
			 for (k=0;k<devp->cv_size;k++)  
				{
		//		 UL_DEBUG((LOGID, "attempt to find: %s / %s",buff,EcoCommU[k].name));
				 if (!strcmp(EcoCommU[k].name,buff))
					{
				     //UL_DEBUG((LOGID, "found at %d",k)); 
					 pos=1;
					 break;
					}
				}
			 if (pos)
			 {
			  char key[ECOCOM_MAX_ENUMERATE][20];
			  char value[ECOCOM_MAX_ENUMERATE][50];
			  char *token;			 
			  pos=0;
//			  UL_DEBUG((LOGID, "data %s",devp->cv[k]));
//			  UL_DEBUG((LOGID, "find token in %s and %s",EcoCommU[i].getCmd,EcoCommU[i].setCmd));
			  sprintf (buff,EcoCommU[i].getCmd);
			  token = strtok(buff,"|\n");
			  while(token != NULL )
				{				 
				 sprintf(key[pos],token);
		//		 UL_DEBUG((LOGID, "keys %s",key[pos]));
				 token = strtok(NULL,"|\n");
				 pos++;
				}			 			 
 			  pos=0;
			  sprintf (buff,EcoCommU[i].setCmd);
			  token = strtok(buff,"|\n");
			  while(token != NULL )
				{				 			
				 sprintf(value[pos],token);
		//		 UL_DEBUG((LOGID, "values %s",value[pos]));
				 token = strtok(NULL,"|\n");
				 pos++;
				}			 			 			 
 	 		  // get data
//			  UL_DEBUG((LOGID, "data %s",devp->cv[k]));
//			  for (j=0; j<EcoCommU[k].num; j++)
//				 buf[j]=DeviceDataBuffer[device][k][j+1+EcoCommU[k].start];
//			  buf[j] = '\0';

			  for (int kk=0;kk<pos;kk++)
				{
		//		 UL_DEBUG((LOGID, "current data %s / keys %s / value %s",devp->cv[k],key[kk],value[kk]));
				 if (!strcmp(key[kk],devp->cv[k]))
					{
					 sprintf(buf,value[kk]);
					 break;
					}
				}
			 }
			}
		 if ((EcoCommU[i].name=="Information/Module board 1")||(EcoCommU[i].name=="Information/Module board 2")) 
			 if (!strcmp(buf,"1")) sprintf(buf,"Wetzer analogue board"); 
				else sprintf(buf,"not available");
 		 if ((EcoCommU[i].name=="Information/Digital IO")||(EcoCommU[i].name=="Information/RS485")||(EcoCommU[i].name=="Information/RS485-Profibus")||(EcoCommU[i].name=="Information/Math channel")) 
			 if (!strcmp(buf,"1"))	sprintf(buf,"available"); 
				else  sprintf(buf,"not available");
 		 if (EcoCommU[i].name=="Information/Data memory") 
			{
			 if (!strcmp(buf,"0"))	sprintf(buf,"not available"); 
			 if (!strcmp(buf,"1"))  sprintf(buf,"Floppy drive available"); 
			 if (!strcmp(buf,"2"))  sprintf(buf,"ATAFlash");
			}
 		 if (EcoCommU[i].name=="Information/Internal memory") 
			 if (!strcmp(buf,"0"))	sprintf(buf,"1Mb"); 
				else sprintf(buf,"2Mb");
 		 if (EcoCommU[i].name=="Information/Integration") 
			 if (!strcmp(buf,"0"))	sprintf(buf,"nicht da"); 
				else sprintf(buf,"vorhanden");
 		 if ((EcoCommU[i].name=="Information/Digital board 1")||(EcoCommU[i].name=="Information/Digital board 2")) 
			 if (!strcmp(buf,"0"))	sprintf(buf,"not available"); 
				else sprintf(buf,"Wetzer digital board");
 		 if (EcoCommU[i].name=="Information/Data interface")
			 if (!strcmp(buf,"0"))	sprintf(buf,"nicht da"); 
				else sprintf(buf,"ext profibuscoupler");
		 strcpy (devp->cv[i],buf);
		 devp->cv_status[i] = TRUE;
		}
	 UL_DEBUG((LOGID, "Copy data to tag %s. Num: %d. Start: %d.",devp->cv[i],EcoCommU[i].num,EcoCommU[i].start));
	}
}
//-----------------------------------------------------------------------------------
HRESULT WriteDevice(int device,const unsigned cmdnum,LPSTR data)
{
	int nump,cnt_false,cnt;
	unsigned int k,j,i;
	unsigned char Out[45],sBuf1[40];
//	unsigned char sBuf1[40];
	const int sBuf[] = {0x01,0x30,0x31,0x2,0x57,0x31,0x30,0x30,0x30,0x3,0x50};	// short frame

	unsigned char *Outt = Out,*Int = sBuf1; 
	DWORD dwStoredFlags;
	COMSTAT comstat;
	dwStoredFlags = EV_RXCHAR | EV_TXEMPTY | EV_RXFLAG;	
	UL_DEBUG((LOGID, "command send (write) %d",cmdnum));
	port.SetMask (dwStoredFlags);
	port.GetStatus (comstat);	

	for (nump=0;nump<=1;nump++)
	{	
	port.SetRTS ();	port.SetDTR ();
	for (i=0;i<=13;i++) 	Out[i] = (char) sBuf[i];
	Out [2]=48+device; Out[4]=*(EcoCommU[cmdnum].setCmd); Out[5]=*(EcoCommU[cmdnum].setCmd+1);
	Out[6]=*(EcoCommU[cmdnum].setCmd+2); Out[7]=*(EcoCommU[cmdnum].setCmd+3); Out[8]=*(EcoCommU[cmdnum].setCmd+4);
	for (j=9;j<9+EcoCommU[cmdnum].num;j++)  Out[j] = data[j-9]; Out[j]=32; j++;
	Out[j]=3; j++; Out[j]=0; for (k=4;k<j;k++) Out[j]=Out[j]^Out[k]; j++;

	for (i=0;i<=j;i++)
	{	
	 port.Write(Outt+i, 1);	port.WaitEvent (dwStoredFlags);
	 UL_DEBUG((LOGID, "send sym (write): %d = %d",i,Out[i]));
	}
	cnt_false=0;  Int = sBuf1;
	for (cnt=0;cnt<40;cnt++)
	{
	if (port.Read(Int+cnt, 1) == FALSE)
		{cnt_false++; if (cnt_false>3) break;}
	else cnt_false=0;
	UL_DEBUG((LOGID, "recieve sym (write): %d = %d.",cnt,sBuf1[cnt]));
	port.GetStatus (comstat);
	}
//-----------------------------------------------------------------------------------
	BOOL bcFF_OK = FALSE; BOOL bcFF_06 = FALSE; int chff=0;
	for (cnt=0;cnt<40;cnt++)
		{ 
		  if (sBuf1[cnt]==0x1)			// ответ от подчиненного???
			 if (sBuf1[cnt+1]==0x30)	// да еще и тот адрес
				if (sBuf1[cnt+2]==48+device)		// да еще и на команду идентификации!
					if ((sBuf1[cnt+4]==0x30)||(sBuf1[cnt+4]==0x31))		// все правильно
						bcFF_06 = TRUE; 
					else
						{
						 switch (sBuf1[cnt+4])
							{	
								case 50: UL_DEBUG((LOGID, "Write: address cannot be edited")); break;
								case 51: UL_DEBUG((LOGID, "Write: address does not be exist")); break;
								case 52: UL_DEBUG((LOGID, "Write: option not available for this address")); break;
								case 53: UL_DEBUG((LOGID, "Write: address not used at the moment")); break;
								case 54: UL_DEBUG((LOGID, "Write: address not allowed  using serial interface")); break;
								case 55: UL_DEBUG((LOGID, "Write: not allowed characters in the parametr")); break;
								case 56: UL_DEBUG((LOGID, "Write: parametr logically incorect")); break;
								case 57: UL_DEBUG((LOGID, "Write: invalid data format")); break;
								case 58: UL_DEBUG((LOGID, "Write: invalid time format")); break;
								case 59: UL_DEBUG((LOGID, "Write: value not available in selection list")); break;
							}
						}
		}
	if (bcFF_06)
		  return S_OK;
	}
	return E_FAIL;
}
//-----------------------------------------------------------------------------------
UINT PollDevice(int device)
{
try {
	int cnt=0,c0m=0, chff=0, num_bytes=0, startid=0, cnt_false=0;
	const int sBuf[] = {0x01,0x30,0x31,0x2,0x56,0x30,0x30,0x30,0x30,0x3,0x50};	// short frame
	unsigned char sBuf1[55],Out[15],DId[80],Data[55];
//	unsigned char Out[15];
//	unsigned char DId[80];
//	unsigned char Data[55];
	unsigned char *Outt = Out,*DeviceId = DId,*Int = sBuf1;
	char *token;	
	char key[10][40], buff[100];
	int pos=0;
//	char buff[100];
	COMSTAT comstat;
	DWORD dwStoredFlags;
	BOOL bcFF_OK = FALSE;
	BOOL bcFF_06 = FALSE;
	BOOL intel_ok = TRUE;	
	dwStoredFlags = EV_RXCHAR | EV_TXEMPTY | EV_RXFLAG;	
	port.SetMask (dwStoredFlags);
	port.GetStatus (comstat);	
//-----------------------------------------------------------------------------------	
	for (c0m=0;c0m<tcount;c0m++) 
	{		 
	 if (intellect)
		{
		 intel_ok=FALSE;
		 for (int k=0;k<=devp->cv_size;k++)  
			 if (!strcmp(EcoCommU[k].getCmd,mBuf[c0m]))
				break;
		 pos=0; sprintf (buff,EcoCommU[k].name);
		 token = strtok(buff,"/\n");
		 while(token != NULL )
			{				 
			 sprintf(key[pos],token); pos++;
			 token = strtok(NULL,"/\n"); 
			}			 			 
		 //UL_DEBUG((LOGID, "key0 %s key1 %d key2 %s | channal %d",key[0],atoi(key[1]),key[2],active_channal[1]));
		 if (!strcmp(key[0],"Analog input"))
		    if ((active_channal[atoi(key[1])]==TRUE)||(!strcmp(key[2],"Type")))
				intel_ok=TRUE;
			else	intel_ok=FALSE;
		 else intel_ok=TRUE;
		 if (!strcmp(key[0],"Digital input"))
		    if ((active_channal[atoi(key[1])+6]==TRUE)||(!strcmp(key[2],"Input function")))
				intel_ok=TRUE;
			else	intel_ok=FALSE;
		 else intel_ok=TRUE;
		 }
	 BOOL prior_ok = FALSE;	 
	 for (int k=0;k<devp->cv_size;k++)  
		{
//		 UL_DEBUG((LOGID, "%s,%d,%s,%d,%d",mBuf[c0m],k,EcoCommU[k].getCmd,EcoCommU[k].scan_it,intel_ok));
		 if ((!strcmp(EcoCommU[k].getCmd,mBuf[c0m]))&&(EcoCommU[k].scan_it)&&intel_ok)
			{
//			 UL_DEBUG((LOGID, "%s | %d | %d | %d",EcoCommU[k].getCmd,EcoCommU[k].scan_it,intel_ok,timetg[k]));
			 timetg[k]--; 
			 if (!timetg[k]) 
				{ prior_ok=TRUE; timetg[k]=EcoCommU[k].priority;}
			 break;
			}
		}
	 if (prior_ok)
	 {	
	 UL_DEBUG((LOGID, "command request %s",mBuf[c0m]));
	 port.SetRTS (); port.SetDTR ();
	 for (int i=0;i<=11;i++) 	Out[i] = (char) sBuf[i]; // 5 первых ff + стандарт команды
	 Out [2]=48+device; Out[4]=mBuf[c0m][0]; Out[5]=mBuf[c0m][1]; Out[6]=mBuf[c0m][2];
	 Out[7]=mBuf[c0m][3]; Out[8]=mBuf[c0m][4]; Out[10]=Out[4]^Out[5]^Out[6]^Out[7]^Out[8]^Out[9];
	 for (i=0;i<=10;i++) 
		{	port.Write(Outt+i, 1);	port.WaitEvent (dwStoredFlags);	}

	 Int = sBuf1;	cnt_false=0;
	 for (cnt=0;cnt<50;cnt++) sBuf1[cnt]=0;
	 for (cnt=0;cnt<50;cnt++)
			{
			 if (port.Read(Int+cnt, 1) == FALSE)
				{ cnt_false++; if (cnt_false>0) break;}		//!!! (4)
			 else cnt_false=0;
				port.GetStatus (comstat);
			}
//-----------------------------------------------------------------------------------
	 chff=0;
	 for (cnt=0;cnt<50;cnt++)
		{ 
		  if (sBuf1[cnt]==0x1)			// ответ от подчиненного???
			 if (sBuf1[cnt+1]==0x30)	// да еще и тот адрес
				if (sBuf1[cnt+2]==48+device)		// да еще и на команду идентификации!
					if ((sBuf1[cnt+4]==0x30)||(sBuf1[cnt+4]==0x31)||(sBuf1[cnt+4]==0x32))		// все правильно
						bcFF_06 = TRUE; 
					else
						{
						 bcFF_06 = FALSE; 
						 switch (sBuf1[cnt+4])
							{	//1
								case 50: if (Out[4]=='R') UL_DEBUG((LOGID, "Read: address cannot be edited")); break;			
								case 51: if (Out[4]=='R') UL_DEBUG((LOGID, "Read: address does not be exist")); break;
								case 52: if (Out[4]=='R') UL_DEBUG((LOGID, "Read: option not available for this address")); break;
								case 53: if (Out[4]=='R') UL_DEBUG((LOGID, "Read: address not used at the moment")); break;
								case 54: if (Out[4]=='R') UL_DEBUG((LOGID, "Read: address not allowed  using serial interface")); break;
								case 55: if (Out[4]=='R') UL_DEBUG((LOGID, "Read: parametr length incorrect")); break;
								case 57: if (Out[4]=='V') UL_DEBUG((LOGID, "Version: error"));
								break;
							}
						}
		  if (sBuf1[cnt]==0x3)
			{
			  num_bytes=cnt-4;
			  break;
			}
		}
	 startid=4;	 
	 if (bcFF_06)
		{
		unsigned char ks=0;
//		UL_DEBUG((LOGID, "num_bytes=%d",num_bytes));
		for (i=startid;i<startid+num_bytes+1;i++)
//			{
			 ks=ks^sBuf1[i];
			 //UL_DEBUG((LOGID, "byte=%d",sBuf1[i]));
//			}
		if (ks!=sBuf1[i])	
			{ 
			  bcFF_06=FALSE; 
			  UL_DEBUG((LOGID, "ks error. ks:%d // need:%d",ks,sBuf1[i]));
			}
		else			
			 UL_DEBUG((LOGID, "ks ok. ks:%d // need:%d",ks,sBuf1[i]));
		for (cnt=0;cnt<50;cnt++) Data[cnt]=0;
		for (cnt=0;cnt<num_bytes;cnt++) 
			{ 
				Data[cnt]=sBuf1[startid+cnt]; 
				DeviceDataBuffer[device][c0m][cnt] = sBuf1[startid+cnt];
//				UL_DEBUG((LOGID, "c0m:%d sym: %d = %d",c0m,cnt,DeviceDataBuffer[device][c0m][cnt]));
			}		
 		 for (int k=0;k<devp->cv_size;k++)  
			if (!strcmp(EcoCommU[k].getCmd,mBuf[c0m])) 	
			  devp->cv_status[k]=TRUE;
		}
		else 
			{
//			 UL_DEBUG((LOGID, "need:%s",mBuf[c0m]));
	  		 for (int k=0;k<devp->cv_size;k++)  
				if (!strcmp(EcoCommU[k].getCmd,mBuf[c0m])) 
					break;
			 devp->cv_status[k]=FALSE;
			 UL_DEBUG((LOGID, "False status for command: %s tag:%d",EcoCommU[k].getCmd,k));
			}
		}
//-----------------------------------------------------------------------------------
	}	
	if (bcFF_06) 
		 return 1;
//-----------------------------------------------------------------------------------	
	return 1;
}
catch (CSerialException* pEx)
{
    TRACE(_T("Handle Exception, Message:%s\n"), pEx->GetErrorMessage());
    pEx->Delete();
	return 2;
}
}
//-------------------------------------------------------------------
int init_tags(void)
{
  FILETIME ft;	//  64-bit value representing the number of 100-ns intervals since January 1,1601
  int i,devi;		
  unsigned rights=0;	// tag type (read/write)
  int ecode;
  GetSystemTimeAsFileTime(&ft);	// retrieves the current system date and time
  EnterCriticalSection(&lk_values);
  for (i=0; i < tTotal; i++)    
	  tn[i] = new char[TAGNAME_LEN];	// reserve memory for massive
  for (devi = 0, i = 0; devi < devp->idnum; devi++) 
  {
    int ci;
    for (ci=0;ci<devp->cv_size; ci++, i++) 
		{
		 //UL_TRACE((LOGID, "loAddRealTag %d %d = %d",ci,i,devp->cv_size));
		 int cmdid = devp->cv_cmdid[ci];
		 sprintf(tn[i], "com%d/id%0.2d/%s",com_num, devp->ids[devi], EcoCommU[cmdid].name);
		 rights=0;
		 if (strcmp(EcoCommU[cmdid].getCmd,"nein")) rights = rights | OPC_READABLE;
	     if (strcmp(EcoCommU[cmdid].setCmd,"nein")) rights = rights | OPC_WRITEABLE;
		 if (strstr(EcoCommU[cmdid].name,"(rVal)")) rights = OPC_READABLE;
	//	 UL_DEBUG((LOGID, "Set priority %d for tags %s num %d",EcoCommU[cmdid].priority,EcoCommU[cmdid].name,ci));
		 timetg[ci]=ci%20+1;
		 VariantInit(&tv[i].tvValue);
		 if ((rvaltags&&strstr(EcoCommU[cmdid].name,"(rVal)"))||!strstr(EcoCommU[cmdid].name,"(rVal)"))
		 if (checkprop(cmdid))
		 {
		 switch (EcoCommU[ci].dtype)
		 {
			 case VT_UI2:
				 V_UI2(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_UI2;
				 ecode = loAddRealTag_a(my_service, &ti[i], // returned TagId 
			       (loRealTag)(i+1), tn[i],
			       0,rights, &tv[i].tvValue, -9999, 9999); break;
			 case VT_I2:
				 V_I2(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_I2;
				 ecode = loAddRealTag_a(my_service, &ti[i], // returned TagId 
			       (loRealTag)(i+1), tn[i],
			       0,rights, &tv[i].tvValue, -99999, 99999); break;
			 case VT_I4:
				 V_I4(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_I4;
				 ecode = loAddRealTag_a(my_service, &ti[i], // returned TagId 
				  (loRealTag)(i+1), tn[i],
			      0,rights, &tv[i].tvValue, -99999, 99999); break;
			 case VT_R4:
				 V_R4(&tv[i].tvValue) = 0.0;
				 V_VT(&tv[i].tvValue) = VT_R4;
				 ecode = loAddRealTag_a(my_service, &ti[i], (loRealTag)(i+1), // != 0
			       tn[i], 0, rights, &tv[i].tvValue, -99999.0, 99999.0); break;
			 case VT_UI1:
				 V_UI1(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_UI1;
				 ecode = loAddRealTag_a(my_service, &ti[i], (loRealTag)(i+1), // != 0
			       tn[i], 0, rights, &tv[i].tvValue, -99, 99); break;
			 case VT_DATE:
				 V_DATE(&tv[i].tvValue) = 0;
				 V_VT(&tv[i].tvValue) = VT_DATE;
				 ecode = loAddRealTag_a(my_service, &ti[i], (loRealTag)(i+1), // != 0
			     tn[i], 0, rights, &tv[i].tvValue, 0, 0); break;

			 default:
				 V_BSTR(&tv[i].tvValue) = SysAllocString(L"");
				 V_VT(&tv[i].tvValue) = VT_BSTR;
				 ecode = loAddRealTag(my_service, &ti[i], (loRealTag)(i+1), 
			     tn[i], 0, rights, &tv[i].tvValue, 0, 0);
		 }      
		 tv[i].tvTi = ti[i];
		 tv[i].tvState.tsTime = ft;
		 tv[i].tvState.tsError = S_OK;
		 tv[i].tvState.tsQuality = OPC_QUALITY_NOT_CONNECTED;
		 UL_TRACE((LOGID, "%!e loAddRealTag(%s) = %u", ecode, tn[i], ti[i]));
		}
	}
  } 
  LeaveCriticalSection(&lk_values);
  if(ecode) 
  {
    UL_ERROR((LOGID, "%!e driver_init()=", ecode));
    return -1;
  }
  return 0;
}
//-------------------------------------------------------------------
UINT DestroyDriver()
{
  if (my_service)		
    {
      int ecode = loServiceDestroy(my_service);
      UL_INFO((LOGID, "%!e loServiceDestroy(%p) = ", ecode));	// destroy derver
      DeleteCriticalSection(&lk_values);						// destroy CS
      my_service = 0;		
    }
try {
	 port.Close();
	 UL_INFO((LOGID, "Close COM-port"));						// write in log
	 return	1;
	}
catch (CSerialException* pEx)
	{
	 UL_ERROR((LOGID, "Error when closing COM-port"));			// write in log
	 pEx->Delete();
	 return 0;
	}
}
//-------------------------------------------------------------------
UINT InitDriver()
{
 loDriver ld;		// structure of driver description
 LONG ecode;		// error code 
 tTotal = ECOCOM_NUM_MAX;		// total tag quantity
 if (my_service) {	
      UL_ERROR((LOGID, "Driver already initialized!"));
      return 0;
  }
 memset(&ld, 0, sizeof(ld));   
 ld.ldRefreshRate =10000;		// polling time 
 ld.ldRefreshRate_min = 7000;	// minimum polling time
 ld.ldWriteTags = WriteTags;	// pointer to function write tag
 ld.ldReadTags = ReadTags;		// pointer to function read tag
 ld.ldSubscribe = activation_monitor;	// callback of tag activity
 ld.ldFlags = loDF_IGNCASE;				// ignore case
 ld.ldBranchSep = '/';					// hierarchial branch separator
 ecode = loServiceCreate(&my_service, &ld, tTotal);		//	creating loService
 UL_TRACE((LOGID, "%!e loServiceCreate()=", ecode));	// write to log returning code
 if (ecode) return 1;									// error to create service	
 InitializeCriticalSection(&lk_values);
 try {
			COMMTIMEOUTS timeouts;
			char buf[50]; char *pbuf=buf; unsigned int speed=200;
			pbuf = ReadParam ("Port","COM");
			com_num = atoi(pbuf); 
			//port.Open(atoi(pbuf), 19200, CSerialPort::NoParity, 8, CSerialPort::OneStopBit, CSerialPort::NoFlowControl, FALSE);
			pbuf = ReadParam ("Port","Speed");			
			speed = atoi(pbuf);
			pbuf = ReadParam ("Server","Intellect");
			if (!strcmp(pbuf,"On")) intellect = TRUE;
			else intellect = FALSE;
			pbuf = ReadParam ("Server","rVal tags");
			if (!strcmp(pbuf,"On")) rvaltags = TRUE;
			else rvaltags = FALSE;

			pbuf = ReadParam ("Server","Digital");
			if (!strcmp(pbuf,"On")) digital = TRUE;
			else digital = FALSE;
			pbuf = ReadParam ("Server","Integrator");
			if (!strcmp(pbuf,"On")) integrator = TRUE;
			else integrator = FALSE;
			pbuf = ReadParam ("Server","Analog");
			if (!strcmp(pbuf,"6")) analog = TRUE;
			else analog = FALSE;
			UL_DEBUG((LOGID, "Digital %d",digital));
			UL_DEBUG((LOGID, "Integrator %d",integrator));
			UL_DEBUG((LOGID, "Analog %d",analog));
			UL_DEBUG((LOGID, "Intellect %d",intellect));
			UL_DEBUG((LOGID, "rVal tags %d",rvaltags));	
			UL_INFO((LOGID, "Opening port COM%d on speed %d",com_num,speed));	
			port.Open(com_num,speed, CSerialPort::NoParity, 8, CSerialPort::OneStopBit, CSerialPort::NoFlowControl, FALSE);
			timeouts.ReadIntervalTimeout = 50;
			timeouts.ReadTotalTimeoutMultiplier = 0; 
			timeouts.ReadTotalTimeoutConstant = 80;				// !!! (180)
			timeouts.WriteTotalTimeoutMultiplier = 0; 
			timeouts.WriteTotalTimeoutConstant = 50; 
			port.SetTimeouts(timeouts);
			UL_INFO((LOGID, "Set COM-port timeouts %d:%d:%d:%d:%d",timeouts.ReadIntervalTimeout,timeouts.ReadTotalTimeoutMultiplier,timeouts.ReadTotalTimeoutConstant,timeouts.WriteTotalTimeoutMultiplier,timeouts.WriteTotalTimeoutConstant));
		}
	catch (CSerialException* pEx)
		{
			TRACE(_T("Handle Exception, Message:%s\n"), pEx->GetErrorMessage());
			UL_ERROR((LOGID, "Unable open COM-port"));		// write in log
//			last_error=0x18; // OPC_QUALITY_COMM_FAILURE
			pEx->Delete();
			return 1;
		}
	try {
 		 UL_INFO((LOGID, "Scan bus"));		// write in log 
		 if (ScanBus()) { 
							UL_INFO((LOGID, "Total %d devices found",devp->idnum)); 
							if (init_tags()) 
									return 1; 
							else return 0;
						}
			else		{ 
//							last_error=0x8; // OPC_QUALITY_NOT_CONNECTED
							UL_ERROR((LOGID, "No devices found")); 
							return 1; 
						}
		}
	catch (CSerialException* pEx)
		{
			UL_ERROR((LOGID, "ScanBus unknown error"));		// write in log
//			last_error=0x18; // OPC_QUALITY_COMM_FAILURE
			pEx->Delete();
			return 1;	
		}
}
//----------------------------------------------------------------------------------------
UINT ScanBus()
{
	devp->idnum = 0;
	devp->ids[devp->idnum] = 1; devp->idnum++;
return 1;
const int sBuf[] = {0x01,0x30,0x31,0x2,0x56,0x30,0x30,0x30,0x30,0x3,0x50};	// short frame
unsigned char Out[15],sBuf1[40];
unsigned char *Outt = Out,*Int = sBuf1;
int cnt,cnt_false, chff=0;
DWORD dwStoredFlags;
dwStoredFlags = EV_RXCHAR | EV_TXEMPTY | EV_RXFLAG;	
COMSTAT comstat;
port.SetMask (dwStoredFlags);
port.GetStatus (comstat);	
port.Read(Int, 1);
devp->idnum = 0;
port.SetRTS (); port.SetDTR ();
//------------------------------------------------------------------------------------
for (int adr=1;adr<=ECOCOM_ID_MAX;adr++)
	for (int nump=0;nump<=0;nump++)
		{
		 port.SetRTS (); port.SetDTR ();
		 for (int i=0;i<=11;i++) 	Out[i] = (char) sBuf[i]; // 5 первых ff + стандарт команды
		 Out [1] = 48+adr/10; Out [2] = 48+adr%10;
		 Out[10]=Out[4]^Out[5]^Out[6]^Out[7]^Out[8]^Out[9];
		 for (i=0;i<=10;i++) 
			{	port.Write(Outt+i, 1);	port.WaitEvent (dwStoredFlags);	}

		 Int = sBuf1;	cnt_false=0;
		 for (cnt=0;cnt<39;cnt++)
			{
			 if (port.Read(Int+cnt, 1) == FALSE)
				{ cnt_false++; if (cnt_false>1) break;}		//!!! (4)
			 else cnt_false=0;
				port.GetStatus (comstat);
			}
//-----------------------------------------------------------------------------------
	BOOL bcFF_OK = FALSE;	BOOL bcFF_06 = FALSE;
	for (cnt=0;cnt<36;cnt++)
		{ 
		  if (sBuf1[cnt]==0x1)			// ответ от подчиненного???
			 if (sBuf1[cnt+1]==0x30)	// да еще и тот адрес
				if (sBuf1[cnt+2]==48+adr)		// да еще и на команду идентификации!
					if (sBuf1[cnt+4]==0x30)		// все правильно
						bcFF_06 = TRUE; 
		}	
    if (bcFF_06) 
		{	
			devp->ids[devp->idnum] = adr; devp->idnum++;
			UL_INFO((LOGID, "Device found on address %d",adr));	// write in log
			break;
		}
	}	 
return devp->idnum;
}
//-----------------------------------------------------------------------------------
char* ReadParam (char *SectionName,char *Value)
{
char buf[150]; 
char string1[50];
char *pbuf=buf;
unsigned int s_ok=0;
sprintf(string1,"[%s]",SectionName);
rewind (CfgFile);
	  while(!feof(CfgFile))
//		{		 
		 if(fgets(buf,50,CfgFile)!=NULL)
			if (strstr(buf,string1))
				{
				 s_ok=1; break;				 
				}
//		}
if (s_ok)
	{
	 while(!feof(CfgFile))
		{
		 if(fgets(buf,100,CfgFile)!=NULL&&strstr(buf,"[")==NULL&&strstr(buf,"]")==NULL)
			{
	//		 UL_INFO((LOGID, "do %s / ishem %s",buf,Value));	// write in log
			 for (s_ok=0;s_ok<strlen(buf)-1;s_ok++)
				 if (buf[s_ok]==';') buf[s_ok+1]='\0';
	//		 UL_INFO((LOGID, "value %s",buf));	// write in log
			 if (strstr(buf,Value))
				{
				 for (s_ok=0;s_ok<strlen(buf)-1;s_ok++)
					if (s_ok>strlen(Value)) buf[s_ok-strlen(Value)-1]=buf[s_ok];
						 buf[s_ok-strlen(Value)-1]='\0';
				 UL_INFO((LOGID, "Section name %s, value %s, che %s",SectionName,Value,buf));	// write in log
				 return pbuf;
				}
			}
		}	 	
 	 if (SectionName=="Port")	{ buf[0]='1'; buf[1]='\0';}
	 return pbuf;
	}
else{
	 sprintf(buf, "error");			// if something go wrong return error
	 return pbuf;
	}	
}
//------------------------------------------------------------------------------------
BOOL checkprop (int cmdid)
{
 if ((analog&&strstr(EcoCommU[cmdid].name,"Analog input/4"))||!strstr(EcoCommU[cmdid].name,"Analog input/4"))
 if ((analog&&strstr(EcoCommU[cmdid].name,"Analog input/5"))||!strstr(EcoCommU[cmdid].name,"Analog input/5"))
 if ((analog&&strstr(EcoCommU[cmdid].name,"Analog input/6"))||!strstr(EcoCommU[cmdid].name,"Analog input/6"))
 if ((digital&&strstr(EcoCommU[cmdid].name,"Digital input/1"))||!strstr(EcoCommU[cmdid].name,"Digital input/1"))
 if ((digital&&strstr(EcoCommU[cmdid].name,"Digital input/2"))||!strstr(EcoCommU[cmdid].name,"Digital input/2"))
 if ((digital&&strstr(EcoCommU[cmdid].name,"Digital input/3"))||!strstr(EcoCommU[cmdid].name,"Digital input/3"))
 if ((digital&&strstr(EcoCommU[cmdid].name,"Digital input/4"))||!strstr(EcoCommU[cmdid].name,"Digital input/4"))
 if ((analog&&strstr(EcoCommU[cmdid].name,"Analog channal 4"))||!strstr(EcoCommU[cmdid].name,"Analog channal 4"))
 if ((analog&&strstr(EcoCommU[cmdid].name,"Analog channal 5"))||!strstr(EcoCommU[cmdid].name,"Analog channal 5"))
 if ((analog&&strstr(EcoCommU[cmdid].name,"Analog channal 6"))||!strstr(EcoCommU[cmdid].name,"Analog channal 6"))
 if ((digital&&strstr(EcoCommU[cmdid].name,"Digital channal 1"))||!strstr(EcoCommU[cmdid].name,"Digital channal 1"))
 if ((digital&&strstr(EcoCommU[cmdid].name,"Digital channal 2"))||!strstr(EcoCommU[cmdid].name,"Digital channal 2"))
 if ((digital&&strstr(EcoCommU[cmdid].name,"Digital channal 3"))||!strstr(EcoCommU[cmdid].name,"Digital channal 3"))
 if ((digital&&strstr(EcoCommU[cmdid].name,"Digital channal 4"))||!strstr(EcoCommU[cmdid].name,"Digital channal 4"))
 if ((integrator&&strstr(EcoCommU[cmdid].name,"(intermediate)"))||!strstr(EcoCommU[cmdid].name,"(intermediate)"))
 if ((integrator&&strstr(EcoCommU[cmdid].name,"(daily)"))||!strstr(EcoCommU[cmdid].name,"(daily)"))
 if ((integrator&&strstr(EcoCommU[cmdid].name,"(monthly)"))||!strstr(EcoCommU[cmdid].name,"(monthly)"))
 if ((integrator&&strstr(EcoCommU[cmdid].name,"(yearly)"))||!strstr(EcoCommU[cmdid].name,"(yearly)"))
		return TRUE;
 return FALSE;
}