// XunJianDlg.cpp : implementation file
//
#include "stdafx.h"
#include "XunJian.h"
#include "XunJianDlg.h"
#include "excel.h"
#include "ExcelFile.h"
#include "ReadTxt.h"

#ifdef _DEBUG
#define new DEBUG_NEW
#undef THIS_FILE
static char THIS_FILE[] = __FILE__;
#endif

bool opened = false;

CString PROJ[/*31*/]={"     �ڵ���Ϣ     ","��λ","   �忨����   ","  �忨���к�  ","�忨��������",
			"   �忨�¶�   ","Ӳ�����","�忨����\r\n�û���","�忨����\r\n��·��","������\r\n�û���","������\r\n��·��",
			"��ʷ���\r\n�û���","LINCENSE��","   ����״̬   ","����xcrp״̬","crash�ļ�","  ����״̬  ",
			"cpu������(5s,1m,5m)","�ڴ�������","  Disk������  ","��ȫ���","","","","","�豸����������",
			"���ж˿�link-damp�Ƿ�����","�˿�Э���Ƿ�ǿ�ƹر�","�Ƿ�Ϊ֧��ģ��",
			"������־�Ƿ�����","log�Ƿ�����"};
CString	COMMAND[]={ "term len 0",                                    // 0
					"show clock",                                    // 1
					"show hard | begin N/A",                         // 2
					"'Temp.- Inlet Air' | exclude Wavelen",          // 3 show hard de | include Temp | exclude 'Temp.- Inlet Air' | exclude Wavelength
					"show diag pod backplane de | begin Back",       // 4
					"show diag pod fantray de | begin Fan",          // 5
					"show diag pod | begin N/A",                     // 6
					"show sub sum all | include Total",              // 7
					"show circuit sum | include total",              // 8
					"show ip route sum all | include Subscriber",    // 9
					"show licenses",                                 // 10
					"show system status",                            // 11
					"show redundancy | inc NO | count",              // 12
					"show redundancy | inc FAIL | count",            // 13
					"show crash",                                    // 14
					"show process | begin NAME",                     // 15
					"show process cpu | include Total",              // 16
					"show disk | begin Internal",                    // 17
					"show port | inc Up | count",                    // 18
					"show port | include Up",                        // 19
					"show port detail",                              // 20
					"RedbackApproved | exclude Yes | count",         // 21 show hard detail | include RedbackApproved | exclude Yes | count
					"show memory",									 // 22
					"show redundancy",								 // 23
					"Num Circuits (TOTAL):"							 // 24 show ism gl | begin after 4 'Num Circuits (TOTAL'
};

CString PPPOESUB[]={"show pppoe sub ",
					"show pppoe sub 1 | count",                     // 
					"show pppoe sub 2 | count",                     // 
					"show pppoe sub 3 | count",                     // 
					"show pppoe sub 4 | count",                     // 
					"show pppoe sub 5 | count",                     // 
					"show pppoe sub 6 | count",                     // 
					"show pppoe sub 7 | count",                     // 
					"show pppoe sub 8 | count",                     // 
					"show pppoe sub 9 | count",                     // 
					"show pppoe sub 10 | count",                    // 
					"show pppoe sub 11 | count",                    // 
					"show pppoe sub 12 | count",                    // 
					"show pppoe sub 13 | count",                    // 
					"show pppoe sub 14 | count"                     // 
};

/************************************************************************
CString CIRSUB[]={	"show circuit ",
					"show circuit 1 sum",             // "show circuit 1 sum | include total"
					"show circuit 2 sum",             // 
					"show circuit 3 sum",             // 
					"show circuit 4 sum",             // 
					"show circuit 5 sum",             // 
					"show circuit 6 sum",             // 
					"show circuit 7 sum",             // 
					"show circuit 8 sum",             // 
					"show circuit 9 sum",             // 
					"show circuit 10 sum",            // 
					"show circuit 11 sum",            // 
					"show circuit 12 sum",            // 
					"show circuit 13 sum",            // 
					"show circuit 14 sum"             // 
};
**************************************************************************/

CXunJianDlg::CXunJianDlg(CWnd* pParent /*=NULL*/)
	: CDialog(CXunJianDlg::IDD, pParent)
{
	//{{AFX_DATA_INIT(CXunJianDlg)
		// NOTE: the ClassWizard will add member initialization here
	//}}AFX_DATA_INIT
	// Note that LoadIcon does not require a subsequent DestroyIcon in Win32
	xj_ipListFileName = _T("ip_list.txt");
}

void CXunJianDlg::DoDataExchange(CDataExchange* pDX)
{
	CDialog::DoDataExchange(pDX);
	//{{AFX_DATA_MAP(CXunJianDlg)
		// NOTE: the ClassWizard will add DDX and DDV calls here
	//}}AFX_DATA_MAP
	DDX_Text(pDX, IDC_HOSTLISTFILENAME, xj_ipListFileName);
}

BEGIN_MESSAGE_MAP(CXunJianDlg, CDialog)
	//{{AFX_MSG_MAP(CXunJianDlg)
	ON_WM_SYSCOMMAND()
	ON_WM_PAINT()
	ON_WM_QUERYDRAGICON()
	//}}AFX_MSG_MAP
	ON_BN_CLICKED(IDOK, &CXunJianDlg::OnBnClickedOk)
	ON_EN_CHANGE(IDC_HOSTLISTFILENAME, &CXunJianDlg::OnEnChangeHostlistfilename)
	ON_BN_CLICKED(IDC_OPENFILE, &CXunJianDlg::OnBnClickedOpenfile)
//	ON_EN_CHANGE(IDC_EDIT1, &CXunJianDlg::OnEnChangeEdit1)
	ON_LBN_SELCHANGE(IDC_HOSTLIST, &CXunJianDlg::OnLbnSelchangeHostlist)
	ON_BN_CLICKED(IDCANCEL, &CXunJianDlg::OnBnClickedCancel)
	ON_NOTIFY(NM_THEMECHANGED, IDC_SCROLLBAR1, &CXunJianDlg::OnNMThemeChangedScrollbar1)
END_MESSAGE_MAP()

/////////////////////////////////////////////////////////////////////////////
// CXunJianDlg message handlers

BOOL CXunJianDlg::OnInitDialog()
{
	CDialog::OnInitDialog();

	// Set the icon for this dialog.  The framework does this automatically
	//  when the application's main window is not a dialog
	SetIcon(m_hIcon, TRUE);			// Set big icon
	SetIcon(m_hIcon, FALSE);		// Set small icon
	
	// TODO: Add extra initialization here
	
	return TRUE;  // return TRUE  unless you set the focus to a control
}

void CXunJianDlg::OnPaint() 
{
	if (IsIconic())
	{
		CPaintDC dc(this); // device context for painting

		SendMessage(WM_ICONERASEBKGND, (WPARAM) dc.GetSafeHdc(), 0);

		// Center icon in client rectangle
		int cxIcon = GetSystemMetrics(SM_CXICON);
		int cyIcon = GetSystemMetrics(SM_CYICON);
		CRect rect;
		GetClientRect(&rect);
		int x = (rect.Width() - cxIcon + 1) / 2;
		int y = (rect.Height() - cyIcon + 1) / 2;

		// Draw the icon
		dc.DrawIcon(x, y, m_hIcon);
	}
	else
	{
		CDialog::OnPaint();
	}
}

// The system calls this to obtain the cursor to display while the user drags
//  the minimized window.
HCURSOR CXunJianDlg::OnQueryDragIcon()
{
	return (HCURSOR) m_hIcon;
}

void CXunJianDlg::OnOK() 
{
	try{
	pList= (CListBox *)GetDlgItem(IDC_HOSTLIST);
	if( pList->GetTextLen(0)>15 || pList->GetTextLen(0)<7 )
	{
		MessageBox("�����б��ļ�δ���أ�������ѡ��");
		return;
	}

	_GUID clsid;
	IUnknown *pUnk;
	IDispatch *pDisp;
	LPDISPATCH lpDisp;

	_Application app;
	Workbooks xj_books;
	_Workbook xj_book;
	Worksheets xj_sheets;
	_Worksheet xj_sheet;
	Range range;
	Range unionRange;
	Range cols;

	Font font;
//	COleVariant background;

	COleVariant covOptional((long)DISP_E_PARAMNOTFOUND, VT_ERROR);

	::CLSIDFromProgID(L"Excel.Application",&clsid); // from registry
	if(GetActiveObject(clsid, NULL,&pUnk) == S_OK)
	{
		VERIFY(pUnk->QueryInterface(IID_IDispatch,(void**) &pDisp) == S_OK);
		app.AttachDispatch(pDisp);
		pUnk->Release();
	} 
	else
	{  
		if(!app.CreateDispatch("Excel.Application"))
		{
			MessageBox("Excel program not found");     
			app.Quit();     
			return;
		}
	}

	xj_books=app.GetWorkbooks();
	xj_book=   xj_books.Add(covOptional);
	xj_sheets= xj_book.GetSheets();
	xj_sheet=  xj_sheets.GetItem(COleVariant((short)1));
	
/************************����ʼ�滭*************************************/
	range=xj_sheet.GetRange(COleVariant("A1"),COleVariant("Z1"));
	range.Merge(COleVariant((long)0));  //�ϲ���һ��,��䣬�Ӵ֣�18��������
	range.SetValue2(COleVariant("Ѳ�챨��"));
	range.SetHorizontalAlignment(COleVariant((long)-4108));
	font=range.GetFont();
	font.SetName(COleVariant("����"));
	font.SetBold(COleVariant((short)TRUE));
	font.SetSize(COleVariant((short)18));

	int i;
	Range item;
	range=xj_sheet.GetRange(COleVariant("A2"),COleVariant("Z2"));
	for(i= 0; i < 26; i++)
	{
		item.AttachDispatch(range.GetItem(COleVariant((long)1),COleVariant((long)i+1)).pdispVal);
		item.SetValue2(COleVariant(PROJ[i]));
	}  //����һ��Ŀ¼

	range=xj_sheet.GetRange(COleVariant("U3"),COleVariant("Y3"));
	for(i=1;i<6;i++)
	{
		item.AttachDispatch(range.GetItem(COleVariant((long)1),COleVariant((long)i)).pdispVal);
		item.SetValue2(COleVariant(PROJ[i+25]));
	}  //��氲ȫ�����Ŀ¼

	range=xj_sheet.GetRange(COleVariant("U2"),COleVariant("Y2"));
	range.Merge(COleVariant((long)0));  //�ϲ���ȫ��鵥Ԫ��

	range=xj_sheet.GetRange(COleVariant("A2"),COleVariant("Z3"));
	lpDisp=range.GetInterior();
	Interior   cellinterior;
	cellinterior.AttachDispatch(lpDisp);
	cellinterior.SetColor(COleVariant((long)0xc0c0c0));  //���ñ���ɫΪ��ɫ
	cellinterior.ReleaseDispatch();
	range=xj_sheet.GetRange(COleVariant("A1"),COleVariant("Z3"));
	range.SetHorizontalAlignment(COleVariant((long)-4108)); //ȫ������
	Borders bord;
	bord=range.GetBorders();
	bord.SetLineStyle(COleVariant((short)1));  //���ñ߿�
	range=xj_sheet.GetRange(COleVariant("A2"),COleVariant("Z3"));
	cols=range.GetEntireColumn();
	cols.AutoFit();  //�Զ�����

	range.AttachDispatch(xj_sheet.GetCells());
	for( i= 1; i < 21; i++)  //�ϲ�2��3��
	{
		unionRange.AttachDispatch(range.GetItem(COleVariant((long)2),COleVariant((long)i)).pdispVal);
		unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)2),COleVariant((long)1)));
		unionRange.Merge(COleVariant((long)0));
	}
	unionRange.AttachDispatch(range.GetItem(COleVariant((long)2),COleVariant((long)26)).pdispVal);
	unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)2),COleVariant((long)1)));
	unionRange.Merge(COleVariant((long)0));

/**************************����ʼ�滭���************************************/

	long usedRowNum; //�м���
	CString handleFile;
	CString hostFileName,hostip;
	bool error = false;
	CString infos,info;
	ExcelFile excelFile;
	ReadTxt xj_txt;
	xj_HostCount=pList->GetCount();
	for(int n_host=0;n_host<xj_HostCount;n_host++)  //��ѭ����һ���ļ�һ��ѭ����
	{		
		pList->GetText(n_host,hostFileName);
		hostip = hostFileName;
		handleFile = hostFileName + _T(" ���ڴ���...");
		pList->DeleteString(n_host);
		pList->InsertString(n_host,handleFile);
		pList->SetCurSel(n_host);
		pList->UpdateWindow();

		hostFileName = xj_FilePath + hostFileName;
		hostFileName += _T(".txt");
		CStdioFile hostFile;
		if(!hostFile.Open(hostFileName,CFile::modeRead,0))
		{  //��¼�������ļ���
			handleFile.Replace("���ڴ���...","ʧ�ܣ�");
			error = true;
			pList->DeleteString(n_host);
			pList->InsertString(n_host,handleFile);
			pList->UpdateWindow();
			continue;
		}
		usedRowNum = excelFile.GetRowCount(xj_sheet);
		range.AttachDispatch(xj_sheet.GetCells());

/**************************ԭ�����㷨v1.2**************************
		info = xj_txt.ReadLine(&hostFile,COMMAND[0]);
		info.Delete(0,7);
		info.Replace("#"+COMMAND[0],"");  //��ȡ�ڵ�����
		info += "\r\n" + hostip;
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(1)),COleVariant(info));

		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(2)),COleVariant("N/A"));
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(3)),COleVariant("backplane"));
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(2)),COleVariant("N/A"));
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(3)),COleVariant("fan tray"));
		infos = xj_txt.ReadLines(&hostFile,COMMAND[2]);  //show hard | begin N/A
		int token = 15;
		info = infos.Tokenize(_T(" "),token);
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(4)),COleVariant(info.Trim()));//�������к�
		token = 50;
		info = infos.Tokenize(_T(" "),token);
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(5)),COleVariant(info.Trim()));//������������
		infos = infos.Mid(76);  //ȥ����һ��
		token = 15;
		info = infos.Tokenize(_T(" "),token);
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(4)),COleVariant(info.Trim()));//�������к�
		token = 50;
		info = infos.Tokenize(_T(" "),token);
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(5)),COleVariant(info.Trim()));//������������
		
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(6)),COleVariant("N/A"));  //�����¶�
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(8)),COleVariant((long)0));//���������û���
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(9)),COleVariant((long)0));//�����·��
		
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(6)),COleVariant("N/A"));  //�����¶�
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(8)),COleVariant((long)0));//���������û���
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(9)),COleVariant((long)0));//���ȵ�·��
		
		i=0;
		infos = infos.Mid(76);
		token = 0;
		while(token > -1)
		{
			info = infos.Tokenize(_T(" "),token);
			if(info == "") break;
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant(long(2)),COleVariant(info.Trim())); //��λ
			info = infos.Tokenize(_T(" "),token);
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant(long(3)),COleVariant(info.Trim())); //�忨����
			info = infos.Tokenize(_T(" "),token);
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant(long(4)),COleVariant(info.Trim())); //�忨���к�
			//info = infos.Tokenize(_T(" "),token);
			//info = infos.Tokenize(_T(" "),token);			
			while( info.Find("-") == -1) info = infos.Tokenize(_T(" "),token);
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant(long(5)),COleVariant(info.Trim())); //��������
			i++;
			
			info = infos.Tokenize(_T(" "),token);
			info = infos.Tokenize(_T(" "),token);
		}
		int cardCount = i;  //�忨��Ŀ���������棬����������ͷ�����
**********************************************************************************/

//********************************���㷨v1.3************************************
		info.Format( _T("%d"), n_host+1);
		info += ". \r\n" + xj_txt.ReadHostName(&hostFile,COMMAND[0],COMMAND[1]);  //��ȡ�ڵ�����
		info += "\r\n" + hostip;
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(1)),COleVariant(info));

		CString t_slot,t_cardtype,t_sn,t_date;
		int cardCount = -2;   //�忨��Ŀ���������棬����������ͷ�����
		for( i=1; xj_txt.ReadHardware( xj_txt.ReadNextLine(&hostFile,COMMAND[2],i), t_slot, t_cardtype, t_sn, t_date); i++ )
		{
			range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(2)),COleVariant(t_slot.Trim())); //��λ
			range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(3)),COleVariant(t_cardtype.Trim())); //�忨����
			range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(4)),COleVariant(t_sn.Trim())); //�忨���к�
			range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(5)),COleVariant(t_date.Trim())); //��������
			if( t_cardtype == "unknown")
			{
				range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(6)),COleVariant("N/A"));   //����忨��ʶ�𣬰忨�¶�ΪN/A
				range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(7)),COleVariant("FAIL"));  //����忨��ʶ�𣬣�Ӳ�����ΪFAIL
			}
			cardCount++;
		}

		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(6)),COleVariant("N/A"));  //�����¶�
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(8)),COleVariant((long)0));//���������û���
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(9)),COleVariant((long)0));//�����·��
		
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(6)),COleVariant("N/A"));  //�����¶�
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(8)),COleVariant((long)0));//���������û���
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(9)),COleVariant((long)0));//���ȵ�·��
//******************************���㷨end************************************

		infos = xj_txt.ReadLines(&hostFile,COMMAND[3]);  //�忨�¶Ȳ���
		infos.Replace("POD Status          : Not Run","");
		infos.Replace("POD Status          : Success","");
		infos.Replace("Temperature        ","");
		infos.Replace("Temp.- Processor    ","");
		infos.Replace("Voltage             : N/A               ","");
		int token = 2;
		COleVariant getitem;
		for( i=3; token > -1;i++)
		{	
			getitem = range.GetItem( COleVariant(usedRowNum+i), COleVariant((long)6));
			getitem.ChangeType(VT_BSTR);
			if( COleVariant("") == getitem )  //�����ⵥԪ��Ϊ��������
			{
				info = infos.Tokenize( _T(":"),token);
				range.SetItem(COleVariant(usedRowNum+i),COleVariant((long)6),COleVariant(info.Trim()));
			}
		}

		info = xj_txt.ReadFailString(&hostFile,COMMAND[4]);
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(7)),COleVariant(info));  //����Ӳ�����
		info = xj_txt.ReadFailString(&hostFile,COMMAND[5]);
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(7)),COleVariant(info));  //����Ӳ�����
		
		CStringArray infoarray ;

		xj_txt.ReadDiagPod(&hostFile,COMMAND[6],infoarray);
		int j;
		for(i = 0, j = 3; i < infoarray.GetCount(); j++)
		{
			getitem = range.GetItem( COleVariant(usedRowNum+j),COleVariant((long)7));
			getitem.ChangeType(VT_BSTR);
			if( getitem == COleVariant("") )  //�����ⵥԪ��Ϊ��������
			{
				info = infoarray.GetAt(i++);
				range.SetItem(COleVariant(usedRowNum+j),COleVariant((long)7),COleVariant(info.Trim()));//�忨Ӳ�����
			}
		}
		
		infoarray.RemoveAll();
		bool slotcirsum = xj_txt.ReadCircuitSum(&hostFile,COMMAND[24],infoarray);
		COleVariant nslot;
		for(i=0; i< cardCount; i++)
		{  //�忨�����û��������ߵ�·��
			nslot = range.GetItem(COleVariant(usedRowNum+3+i),COleVariant((long)2));  //ȡ��λ��
			nslot.ChangeType(VT_INT);
			info = xj_txt.ReadNextLine(&hostFile,PPPOESUB[nslot.intVal]);
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant((long)8),COleVariant(info.Trim()));  //�忨�����û���
			if(slotcirsum)
			{
				info = infoarray.GetAt(nslot.intVal - 1);
				range.SetItem(COleVariant(usedRowNum+3+i),COleVariant((long)9),COleVariant(info.Trim()));  //�忨���ߵ�·��
			}
		}

		info = xj_txt.ReadNextLine(&hostFile,COMMAND[7]);  //�������û���

		info.Replace("Total=","");
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)10),COleVariant(info.Trim())); 

		info = xj_txt.ReadNextLine(&hostFile,COMMAND[8]);  //�����ߵ�·��
		info.Replace("total:","");
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)11),COleVariant(info.Trim())); 

		infos = xj_txt.ReadNextLine(&hostFile,COMMAND[9]);  //��ʷ��������û���
		if( infos == _T("") ) info = _T("0");
		else
		{
			token = 0;
			for(i = 0; i < 5 ; i++) info = infos.Tokenize(" ",token);
		}
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)12),COleVariant(info.Trim()));

		infos = xj_txt.ReadNextLine(&hostFile,COMMAND[10]);  //License��
		if( infos == _T("") ) info = _T("0");
		else 
		{
			token = 0;
			for(i = 0; i < 6 ; i++) info = infos.Tokenize(" ",token);
		}
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)13),COleVariant(info.Trim()));

		infos = xj_txt.ReadLines(&hostFile,COMMAND[11]);  //����״̬
		if(infos.Find("OK") > -1) info = _T("OK");
		else info = infos;
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)14),COleVariant(info.Trim()));

		info = xj_txt.ReadRedundancy(&hostFile,COMMAND[12],COMMAND[13],COMMAND[23]);  //����xcrp״̬
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)15),COleVariant(info.Trim()));

		info = xj_txt.ReadLines(&hostFile,COMMAND[14]);    //Crash�ļ�
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)16),COleVariant(info.Trim()));

		info = xj_txt.ReadProcess(&hostFile,COMMAND[15]);  //����״̬
		if(info == "") info = _T("OK");
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)17),COleVariant(info.Trim()));

		infos = xj_txt.ReadNextLine(&hostFile,COMMAND[16]); //CPU������
		token = 39;
		info = infos.Tokenize(":",token);
		info.Replace(" ","");
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)18),COleVariant(info.Trim()));

		info = xj_txt.ReadMemory(&hostFile,COMMAND[22]);   //�ڴ�������
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)19),COleVariant(info.Trim()));

		info = xj_txt.ReadDisk(&hostFile,COMMAND[17]); //Disk������
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)20),COleVariant(info.Trim()));

		info = xj_txt.ReadPortLinkDampening(&hostFile,COMMAND[20]); //Link-Dampening
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)21),COleVariant(info));

		info = xj_txt.ReadPortAutoNegotiate(&hostFile,COMMAND[20]); //Auto-negotiate
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)22),COleVariant(info));

		info = xj_txt.ReadSfp(&hostFile,COMMAND[21]); //��֧�ֹ�ģ��
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)23),COleVariant(info));

		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)24),COleVariant("����"));  //������־

		hostFile.Close();

		unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)1)).pdispVal);
		unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)cardCount+2),COleVariant((long)1)));
		unionRange.Merge(COleVariant((long)0));  //�ڵ����Ƶ�Ԫ��ϲ�
		for( i= 10; i < 27; i++)  //�ϲ�������Ԫ��
		{
			unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)i)).pdispVal);
			unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)cardCount+2),COleVariant((long)1)));
			unionRange.Merge(COleVariant((long)0));
		}

		unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)1)).pdispVal);
		unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)cardCount+2),COleVariant((long)26)));
		unionRange.SetRowHeight(COleVariant(13.5));
		bord = unionRange.GetBorders();
		bord.SetLineStyle(COleVariant((short)1));  //���ñ߿�

		handleFile.Replace("���ڴ���...","�����");
		pList->DeleteString(n_host);
		pList->InsertString(n_host,handleFile);
		pList->UpdateWindow();
	}

	CTime time;
	time = time.GetCurrentTime();
	infos = time.Format("%Y%m%d%H%M%S");  //time.Format();
	info = _T("Ѳ�챨��") + infos + _T(".xlsx");
	info = xj_FilePath + info;
	info.Replace("\\\\","\\");

	xj_book.SaveAs(COleVariant(info),covOptional,covOptional,covOptional,covOptional,covOptional,0,covOptional,covOptional,covOptional,covOptional,covOptional);
	
	if(error == true )
	{
		MessageBox("Ѳ�챨������ɣ��ѱ��浽\r\n" + info + "\r\n���ļ��򿪴��󣬵��\"ȷ��\"���ز鿴","���ļ��򿪴���",MB_OK|MB_ICONWARNING);
		app.Quit();
	}
	else 
	{
		if(MessageBox("Ѳ�챨������ɣ��ѱ��浽\r\n" + info + "\r\n���\"ȷ��\"���ļ��鿴","���ɱ������",MB_OKCANCEL) == IDOK)
		{
			app.SetVisible(TRUE);
			app.SetUserControl(TRUE);
		}
		else app.Quit();
	}
}
catch (CFileException* e)
    {
        e->ReportError();
        e->Delete();
    }
	//CDialog::OnOK();
}

void CXunJianDlg::OnBnClickedOk()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	OnOK();
}

void CXunJianDlg::OnEnChangeHostlistfilename()
{
	// TODO:  ����ÿؼ��� RICHEDIT �ؼ�������������
	// ���͸�֪ͨ��������д CDialog::OnInitDialog()
	// ���������� CRichEditCtrl().SetEventMask()��
	// ͬʱ�� ENM_CHANGE ��־�������㵽�����С�

	// TODO:  �ڴ���ӿؼ�֪ͨ����������
}

void CXunJianDlg::OnBnClickedOpenfile()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	CStdioFile listfile;
	if( opened == true || listfile.Open(xj_ipListFileName,CFile::modeRead,0) == false)
	{
		//listfile.Close();
		m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
		CFileDialog FileDlg(true, _T("txt"),NULL,OFN_FILEMUSTEXIST|OFN_HIDEREADONLY, 
										 "�ı��ļ�(*.TXT)|*.TXT|All Files(*.*)|*.*||"); 
		if( FileDlg.DoModal() == IDOK )
		{ 
			xj_ipListFileName = FileDlg.GetFileName();
			listfile.Open(xj_ipListFileName,CFile::modeRead,0);
		}
		else return;
	}
	xj_ipListFileName = listfile.GetFileName();//FileDlg.GetFileName();
	xj_FilePath = listfile.GetFilePath();//FileDlg.GetPathName();
	xj_FilePath.Replace(xj_ipListFileName,"");
	xj_FilePath.Replace("\\","\\\\");
	pList= (CListBox *)GetDlgItem(IDC_HOSTLIST);
    pList->ResetContent();   
    CString str;
    while(listfile.ReadString(str))   
	{
		if(str.Find('.') > -1 && str.Find("\'") == -1)  //���в����� ,�� '���Ų����룬�൱��ע�͵�����
		{
			str.Replace(" ","");  //ȥ�����пո�
			pList->AddString(str);
		}
    }
	opened = true;
    listfile.Close(); 
	UpdateData(false);
}


void CXunJianDlg::OnLbnSelchangeHostlist()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
}

void CXunJianDlg::OnBnClickedCancel()
{
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	OnCancel();
}

void CXunJianDlg::OnNMThemeChangedScrollbar1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// �ù���Ҫ��ʹ�� Windows XP ����߰汾��
	// ���� _WIN32_WINNT ���� >= 0x0501��
	// TODO: �ڴ���ӿؼ�֪ͨ����������
	*pResult = 0;
}
