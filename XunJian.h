// XunJian.h : main header file for the XUNJIAN application
//

#if !defined(AFX_XUNJIAN_H__CF30FBFC_0333_48C7_9839_05CBC4908036__INCLUDED_)
#define AFX_XUNJIAN_H__CF30FBFC_0333_48C7_9839_05CBC4908036__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#ifndef __AFXWIN_H__
	#error include 'stdafx.h' before including this file for PCH
#endif

#include "resource.h"		// main symbols

/////////////////////////////////////////////////////////////////////////////
// CXunJianApp:
// See XunJian.cpp for the implementation of this class
//
;
class CXunJianApp : public CWinApp
{
public:
	CXunJianApp();

// Overrides
	// ClassWizard generated virtual function overrides
	//{{AFX_VIRTUAL(CXunJianApp)
	public:
	virtual BOOL InitInstance();
	//}}AFX_VIRTUAL

// Implementation

	//{{AFX_MSG(CXunJianApp)
		// NOTE - the ClassWizard will add and remove member functions here.
		//    DO NOT EDIT what you see in these blocks of generated code !
	//}}AFX_MSG
	DECLARE_MESSAGE_MAP()
};


/////////////////////////////////////////////////////////////////////////////

//{{AFX_INSERT_LOCATION}}
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.

#endif // !defined(AFX_XUNJIAN_H__CF30FBFC_0333_48C7_9839_05CBC4908036__INCLUDED_)
