#pragma once

class ExcelFile
{
public:
	ExcelFile(void);
	~ExcelFile(void);
	int GetRowCount(_Worksheet);   
	int GetColumnCount(_Worksheet); 
};
