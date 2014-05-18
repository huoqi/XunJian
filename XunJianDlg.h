// XunJianDlg.h : header file
//

#if !defined(AFX_XUNJIANDLG_H__DEC771C6_5F9F_4510_A59E_AD3714369B84__INCLUDED_)
#define AFX_XUNJIANDLG_H__DEC771C6_5F9F_4510_A59E_AD3714369B84__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

/////////////////////////////////////////////////////////////////////////////
// CXunJianDlg dialog

class CXunJianDlg : public CDialog
{
// Construction
public:
	CXunJianDlg(CWnd* pParent = NULL);	// standard constructor

// Dialog Data
	//{{AFX_DATA(CXunJianDlg)
	enum { IDD = IDD_XUNJIAN_DIALOG };
		// NOTE: the ClassWizard will add data members here
	//}}AFX_DATA

	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CXunJianDlg)
	protected:
	virtual void DoDataExchange(CDataExchange* pDX);	// DDX/DDV support
	//}}AFX_VIRTUAL

// Implementation
protected:
	HICON m_hIcon;

	// Generated message map functions
	//{{AFX_MSG(CXunJianDlg)
	virtual BOOL OnInitDialog();
//	afx_msg void OnSysCommand(UINT nID, LPARAM lParam);
	afx_msg void OnPaint();
	afx_msg HCURSOR OnQueryDragIcon();
	virtual void OnOK();
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
public:
	afx_msg void OnBnClickedOk();
	afx_msg void OnEnChangeHostlistfilename();
	afx_msg void OnBnClickedOpenfile();

	void CXunJianDlg::OpenListFile();

//自己加的全局变量
public:
	CString xj_ipListFileName;
	//CStringArray xj_HostList;
	//CString xj_HostName;
	CString xj_FilePath;
	CListBox *pList; 
	int xj_HostCount;

	//afx_msg void OnEnChangeEdit1();
	afx_msg void OnLbnSelchangeHostlist();
	afx_msg void OnBnClickedCancel();
	afx_msg void OnNMThemeChangedScrollbar1(NMHDR *pNMHDR, LRESULT *pResult);
};

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_XUNJIANDLG_H__DEC771C6_5F9F_4510_A59E_AD3714369B84__INCLUDED_)
