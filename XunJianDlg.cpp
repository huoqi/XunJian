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

CString PROJ[/*31*/]={"     节点信息     ","槽位","   板卡类型   ","  板卡序列号  ","板卡生产日期",
			"   板卡温度   ","硬件检测","板卡在线\r\n用户数","板卡在线\r\n电路数","总在线\r\n用户数","总在线\r\n电路数",
			"历史最大\r\n用户数","LINCENSE数","   运行状态   ","主备xcrp状态","crash文件","  进程状态  ",
			"cpu利用率(5s,1m,5m)","内存利用率","  Disk利用率  ","安全检查","","","","","设备评估及建议",
			"下行端口link-damp是否配置","端口协商是否强制关闭","是否为支持模块",
			"启动日志是否正常","log是否正常"};
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
		MessageBox("主机列表文件未加载，请重新选择！");
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
	
/************************表格初始绘画*************************************/
	range=xj_sheet.GetRange(COleVariant("A1"),COleVariant("Z1"));
	range.Merge(COleVariant((long)0));  //合并第一行,填充，加粗，18磅，居中
	range.SetValue2(COleVariant("巡检报表"));
	range.SetHorizontalAlignment(COleVariant((long)-4108));
	font=range.GetFont();
	font.SetName(COleVariant("黑体"));
	font.SetBold(COleVariant((short)TRUE));
	font.SetSize(COleVariant((short)18));

	int i;
	Range item;
	range=xj_sheet.GetRange(COleVariant("A2"),COleVariant("Z2"));
	for(i= 0; i < 26; i++)
	{
		item.AttachDispatch(range.GetItem(COleVariant((long)1),COleVariant((long)i+1)).pdispVal);
		item.SetValue2(COleVariant(PROJ[i]));
	}  //描绘第一行目录

	range=xj_sheet.GetRange(COleVariant("U3"),COleVariant("Y3"));
	for(i=1;i<6;i++)
	{
		item.AttachDispatch(range.GetItem(COleVariant((long)1),COleVariant((long)i)).pdispVal);
		item.SetValue2(COleVariant(PROJ[i+25]));
	}  //描绘安全检查子目录

	range=xj_sheet.GetRange(COleVariant("U2"),COleVariant("Y2"));
	range.Merge(COleVariant((long)0));  //合并安全检查单元格

	range=xj_sheet.GetRange(COleVariant("A2"),COleVariant("Z3"));
	lpDisp=range.GetInterior();
	Interior   cellinterior;
	cellinterior.AttachDispatch(lpDisp);
	cellinterior.SetColor(COleVariant((long)0xc0c0c0));  //设置背景色为灰色
	cellinterior.ReleaseDispatch();
	range=xj_sheet.GetRange(COleVariant("A1"),COleVariant("Z3"));
	range.SetHorizontalAlignment(COleVariant((long)-4108)); //全部居中
	Borders bord;
	bord=range.GetBorders();
	bord.SetLineStyle(COleVariant((short)1));  //设置边框
	range=xj_sheet.GetRange(COleVariant("A2"),COleVariant("Z3"));
	cols=range.GetEntireColumn();
	cols.AutoFit();  //自动调整

	range.AttachDispatch(xj_sheet.GetCells());
	for( i= 1; i < 21; i++)  //合并2、3行
	{
		unionRange.AttachDispatch(range.GetItem(COleVariant((long)2),COleVariant((long)i)).pdispVal);
		unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)2),COleVariant((long)1)));
		unionRange.Merge(COleVariant((long)0));
	}
	unionRange.AttachDispatch(range.GetItem(COleVariant((long)2),COleVariant((long)26)).pdispVal);
	unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)2),COleVariant((long)1)));
	unionRange.Merge(COleVariant((long)0));

/**************************表格初始绘画完成************************************/

	long usedRowNum; //行计数
	CString handleFile;
	CString hostFileName,hostip;
	bool error = false;
	CString infos,info;
	ExcelFile excelFile;
	ReadTxt xj_txt;
	xj_HostCount=pList->GetCount();
	for(int n_host=0;n_host<xj_HostCount;n_host++)  //主循环，一个文件一次循环。
	{		
		pList->GetText(n_host,hostFileName);
		hostip = hostFileName;
		handleFile = hostFileName + _T(" 正在处理...");
		pList->DeleteString(n_host);
		pList->InsertString(n_host,handleFile);
		pList->SetCurSel(n_host);
		pList->UpdateWindow();

		hostFileName = xj_FilePath + hostFileName;
		hostFileName += _T(".txt");
		CStdioFile hostFile;
		if(!hostFile.Open(hostFileName,CFile::modeRead,0))
		{  //记录不存在文件名
			handleFile.Replace("正在处理...","失败！");
			error = true;
			pList->DeleteString(n_host);
			pList->InsertString(n_host,handleFile);
			pList->UpdateWindow();
			continue;
		}
		usedRowNum = excelFile.GetRowCount(xj_sheet);
		range.AttachDispatch(xj_sheet.GetCells());

/**************************原来的算法v1.2**************************
		info = xj_txt.ReadLine(&hostFile,COMMAND[0]);
		info.Delete(0,7);
		info.Replace("#"+COMMAND[0],"");  //获取节点名称
		info += "\r\n" + hostip;
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(1)),COleVariant(info));

		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(2)),COleVariant("N/A"));
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(3)),COleVariant("backplane"));
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(2)),COleVariant("N/A"));
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(3)),COleVariant("fan tray"));
		infos = xj_txt.ReadLines(&hostFile,COMMAND[2]);  //show hard | begin N/A
		int token = 15;
		info = infos.Tokenize(_T(" "),token);
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(4)),COleVariant(info.Trim()));//背板序列号
		token = 50;
		info = infos.Tokenize(_T(" "),token);
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(5)),COleVariant(info.Trim()));//备板生产日期
		infos = infos.Mid(76);  //去掉第一行
		token = 15;
		info = infos.Tokenize(_T(" "),token);
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(4)),COleVariant(info.Trim()));//风扇序列号
		token = 50;
		info = infos.Tokenize(_T(" "),token);
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(5)),COleVariant(info.Trim()));//风扇生产日期
		
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(6)),COleVariant("N/A"));  //背板温度
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(8)),COleVariant((long)0));//背板在线用户数
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(9)),COleVariant((long)0));//背板电路数
		
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(6)),COleVariant("N/A"));  //风扇温度
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(8)),COleVariant((long)0));//风扇在线用户数
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(9)),COleVariant((long)0));//风扇电路数
		
		i=0;
		infos = infos.Mid(76);
		token = 0;
		while(token > -1)
		{
			info = infos.Tokenize(_T(" "),token);
			if(info == "") break;
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant(long(2)),COleVariant(info.Trim())); //槽位
			info = infos.Tokenize(_T(" "),token);
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant(long(3)),COleVariant(info.Trim())); //板卡类型
			info = infos.Tokenize(_T(" "),token);
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant(long(4)),COleVariant(info.Trim())); //板卡序列号
			//info = infos.Tokenize(_T(" "),token);
			//info = infos.Tokenize(_T(" "),token);			
			while( info.Find("-") == -1) info = infos.Tokenize(_T(" "),token);
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant(long(5)),COleVariant(info.Trim())); //生产日期
			i++;
			
			info = infos.Tokenize(_T(" "),token);
			info = infos.Tokenize(_T(" "),token);
		}
		int cardCount = i;  //板卡数目。包括引擎，不包括背板和风扇数
**********************************************************************************/

//********************************新算法v1.3************************************
		info.Format( _T("%d"), n_host+1);
		info += ". \r\n" + xj_txt.ReadHostName(&hostFile,COMMAND[0],COMMAND[1]);  //获取节点名称
		info += "\r\n" + hostip;
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(1)),COleVariant(info));

		CString t_slot,t_cardtype,t_sn,t_date;
		int cardCount = -2;   //板卡数目。包括引擎，不包括背板和风扇数
		for( i=1; xj_txt.ReadHardware( xj_txt.ReadNextLine(&hostFile,COMMAND[2],i), t_slot, t_cardtype, t_sn, t_date); i++ )
		{
			range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(2)),COleVariant(t_slot.Trim())); //槽位
			range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(3)),COleVariant(t_cardtype.Trim())); //板卡类型
			range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(4)),COleVariant(t_sn.Trim())); //板卡序列号
			range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(5)),COleVariant(t_date.Trim())); //生产日期
			if( t_cardtype == "unknown")
			{
				range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(6)),COleVariant("N/A"));   //如果板卡不识别，板卡温度为N/A
				range.SetItem(COleVariant(usedRowNum+i),COleVariant(long(7)),COleVariant("FAIL"));  //如果板卡不识别，，硬件检测为FAIL
			}
			cardCount++;
		}

		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(6)),COleVariant("N/A"));  //背板温度
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(8)),COleVariant((long)0));//背板在线用户数
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(9)),COleVariant((long)0));//背板电路数
		
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(6)),COleVariant("N/A"));  //风扇温度
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(8)),COleVariant((long)0));//风扇在线用户数
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(9)),COleVariant((long)0));//风扇电路数
//******************************新算法end************************************

		infos = xj_txt.ReadLines(&hostFile,COMMAND[3]);  //板卡温度部分
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
			if( COleVariant("") == getitem )  //如果检测单元格为空则填入
			{
				info = infos.Tokenize( _T(":"),token);
				range.SetItem(COleVariant(usedRowNum+i),COleVariant((long)6),COleVariant(info.Trim()));
			}
		}

		info = xj_txt.ReadFailString(&hostFile,COMMAND[4]);
		range.SetItem(COleVariant(usedRowNum+1),COleVariant(long(7)),COleVariant(info));  //背板硬件检测
		info = xj_txt.ReadFailString(&hostFile,COMMAND[5]);
		range.SetItem(COleVariant(usedRowNum+2),COleVariant(long(7)),COleVariant(info));  //风扇硬件检测
		
		CStringArray infoarray ;

		xj_txt.ReadDiagPod(&hostFile,COMMAND[6],infoarray);
		int j;
		for(i = 0, j = 3; i < infoarray.GetCount(); j++)
		{
			getitem = range.GetItem( COleVariant(usedRowNum+j),COleVariant((long)7));
			getitem.ChangeType(VT_BSTR);
			if( getitem == COleVariant("") )  //如果检测单元格为空则填入
			{
				info = infoarray.GetAt(i++);
				range.SetItem(COleVariant(usedRowNum+j),COleVariant((long)7),COleVariant(info.Trim()));//板卡硬件检测
			}
		}
		
		infoarray.RemoveAll();
		bool slotcirsum = xj_txt.ReadCircuitSum(&hostFile,COMMAND[24],infoarray);
		COleVariant nslot;
		for(i=0; i< cardCount; i++)
		{  //板卡在线用户数、在线电路数
			nslot = range.GetItem(COleVariant(usedRowNum+3+i),COleVariant((long)2));  //取槽位号
			nslot.ChangeType(VT_INT);
			info = xj_txt.ReadNextLine(&hostFile,PPPOESUB[nslot.intVal]);
			range.SetItem(COleVariant(usedRowNum+3+i),COleVariant((long)8),COleVariant(info.Trim()));  //板卡在线用户数
			if(slotcirsum)
			{
				info = infoarray.GetAt(nslot.intVal - 1);
				range.SetItem(COleVariant(usedRowNum+3+i),COleVariant((long)9),COleVariant(info.Trim()));  //板卡在线电路数
			}
		}

		info = xj_txt.ReadNextLine(&hostFile,COMMAND[7]);  //总在线用户数

		info.Replace("Total=","");
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)10),COleVariant(info.Trim())); 

		info = xj_txt.ReadNextLine(&hostFile,COMMAND[8]);  //总在线电路数
		info.Replace("total:","");
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)11),COleVariant(info.Trim())); 

		infos = xj_txt.ReadNextLine(&hostFile,COMMAND[9]);  //历史在线最大用户数
		if( infos == _T("") ) info = _T("0");
		else
		{
			token = 0;
			for(i = 0; i < 5 ; i++) info = infos.Tokenize(" ",token);
		}
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)12),COleVariant(info.Trim()));

		infos = xj_txt.ReadNextLine(&hostFile,COMMAND[10]);  //License数
		if( infos == _T("") ) info = _T("0");
		else 
		{
			token = 0;
			for(i = 0; i < 6 ; i++) info = infos.Tokenize(" ",token);
		}
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)13),COleVariant(info.Trim()));

		infos = xj_txt.ReadLines(&hostFile,COMMAND[11]);  //运行状态
		if(infos.Find("OK") > -1) info = _T("OK");
		else info = infos;
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)14),COleVariant(info.Trim()));

		info = xj_txt.ReadRedundancy(&hostFile,COMMAND[12],COMMAND[13],COMMAND[23]);  //主备xcrp状态
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)15),COleVariant(info.Trim()));

		info = xj_txt.ReadLines(&hostFile,COMMAND[14]);    //Crash文件
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)16),COleVariant(info.Trim()));

		info = xj_txt.ReadProcess(&hostFile,COMMAND[15]);  //进程状态
		if(info == "") info = _T("OK");
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)17),COleVariant(info.Trim()));

		infos = xj_txt.ReadNextLine(&hostFile,COMMAND[16]); //CPU利用率
		token = 39;
		info = infos.Tokenize(":",token);
		info.Replace(" ","");
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)18),COleVariant(info.Trim()));

		info = xj_txt.ReadMemory(&hostFile,COMMAND[22]);   //内存利用率
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)19),COleVariant(info.Trim()));

		info = xj_txt.ReadDisk(&hostFile,COMMAND[17]); //Disk利用率
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)20),COleVariant(info.Trim()));

		info = xj_txt.ReadPortLinkDampening(&hostFile,COMMAND[20]); //Link-Dampening
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)21),COleVariant(info));

		info = xj_txt.ReadPortAutoNegotiate(&hostFile,COMMAND[20]); //Auto-negotiate
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)22),COleVariant(info));

		info = xj_txt.ReadSfp(&hostFile,COMMAND[21]); //不支持光模块
		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)23),COleVariant(info));

		range.SetItem(COleVariant(usedRowNum+1),COleVariant((long)24),COleVariant("正常"));  //启动日志

		hostFile.Close();

		unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)1)).pdispVal);
		unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)cardCount+2),COleVariant((long)1)));
		unionRange.Merge(COleVariant((long)0));  //节点名称单元格合并
		for( i= 10; i < 27; i++)  //合并其他单元格
		{
			unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)i)).pdispVal);
			unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)cardCount+2),COleVariant((long)1)));
			unionRange.Merge(COleVariant((long)0));
		}

		unionRange.AttachDispatch(range.GetItem(COleVariant(usedRowNum+1),COleVariant((long)1)).pdispVal);
		unionRange.AttachDispatch(unionRange.GetResize(COleVariant((long)cardCount+2),COleVariant((long)26)));
		unionRange.SetRowHeight(COleVariant(13.5));
		bord = unionRange.GetBorders();
		bord.SetLineStyle(COleVariant((short)1));  //设置边框

		handleFile.Replace("正在处理...","已完成");
		pList->DeleteString(n_host);
		pList->InsertString(n_host,handleFile);
		pList->UpdateWindow();
	}

	CTime time;
	time = time.GetCurrentTime();
	infos = time.Format("%Y%m%d%H%M%S");  //time.Format();
	info = _T("巡检报表") + infos + _T(".xlsx");
	info = xj_FilePath + info;
	info.Replace("\\\\","\\");

	xj_book.SaveAs(COleVariant(info),covOptional,covOptional,covOptional,covOptional,covOptional,0,covOptional,covOptional,covOptional,covOptional,covOptional);
	
	if(error == true )
	{
		MessageBox("巡检报表已完成，已保存到\r\n" + info + "\r\n有文件打开错误，点击\"确定\"返回查看","有文件打开错误！",MB_OK|MB_ICONWARNING);
		app.Quit();
	}
	else 
	{
		if(MessageBox("巡检报表已完成，已保存到\r\n" + info + "\r\n点击\"确定\"打开文件查看","生成报表完成",MB_OKCANCEL) == IDOK)
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
	// TODO: 在此添加控件通知处理程序代码
	OnOK();
}

void CXunJianDlg::OnEnChangeHostlistfilename()
{
	// TODO:  如果该控件是 RICHEDIT 控件，则它将不会
	// 发送该通知，除非重写 CDialog::OnInitDialog()
	// 函数并调用 CRichEditCtrl().SetEventMask()，
	// 同时将 ENM_CHANGE 标志“或”运算到掩码中。

	// TODO:  在此添加控件通知处理程序代码
}

void CXunJianDlg::OnBnClickedOpenfile()
{
	// TODO: 在此添加控件通知处理程序代码
	CStdioFile listfile;
	if( opened == true || listfile.Open(xj_ipListFileName,CFile::modeRead,0) == false)
	{
		//listfile.Close();
		m_hIcon = AfxGetApp()->LoadIcon(IDR_MAINFRAME);
		CFileDialog FileDlg(true, _T("txt"),NULL,OFN_FILEMUSTEXIST|OFN_HIDEREADONLY, 
										 "文本文件(*.TXT)|*.TXT|All Files(*.*)|*.*||"); 
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
		if(str.Find('.') > -1 && str.Find("\'") == -1)  //空行不加入 ,有 '符号不加入，相当于注释掉此行
		{
			str.Replace(" ","");  //去除所有空格
			pList->AddString(str);
		}
    }
	opened = true;
    listfile.Close(); 
	UpdateData(false);
}


void CXunJianDlg::OnLbnSelchangeHostlist()
{
	// TODO: 在此添加控件通知处理程序代码
}

void CXunJianDlg::OnBnClickedCancel()
{
	// TODO: 在此添加控件通知处理程序代码
	OnCancel();
}

void CXunJianDlg::OnNMThemeChangedScrollbar1(NMHDR *pNMHDR, LRESULT *pResult)
{
	// 该功能要求使用 Windows XP 或更高版本。
	// 符号 _WIN32_WINNT 必须 >= 0x0501。
	// TODO: 在此添加控件通知处理程序代码
	*pResult = 0;
}
