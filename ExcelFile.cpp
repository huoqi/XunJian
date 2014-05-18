#include "stdafx.h"
#include "excel.h"
#include "ExcelFile.h"

ExcelFile::ExcelFile(void)
{
}

ExcelFile::~ExcelFile(void)
{
}

int ExcelFile::GetColumnCount(_Worksheet m_sheet)   
{   
	Range rg;   
	Range usedRange;   
	usedRange.AttachDispatch(m_sheet.GetUsedRange(), true);   
	rg.AttachDispatch(usedRange.GetColumns(), true);   
	int count = rg.GetCount();   
	usedRange.ReleaseDispatch();   
	rg.ReleaseDispatch();   
	return count;   
}  
 
int ExcelFile::GetRowCount(_Worksheet m_sheet)   
{   
	Range rg;   
	Range usedRange;   
	usedRange.AttachDispatch(m_sheet.GetUsedRange(), true);   
	rg.AttachDispatch(usedRange.GetRows(), true);   
	int count = rg.GetCount();   
	usedRange.ReleaseDispatch();   
	rg.ReleaseDispatch();   
	return count;   
}
