#pragma once

class ReadTxt
{
public:
	ReadTxt(void);
	~ReadTxt(void);

	CString ReadLine(CStdioFile *s_file,CString findstr);
	CString ReadLines(CStdioFile *s_file,CString startstr);
	CString ReadLines(CStdioFile *s_file,CString startstr,CString endstr);
	CString ReadNextLine(CStdioFile *s_file,CString startstr,int next=1);
	CString ReadNextLines(CStdioFile *s_file,CString startstr,int next=1);
	CString ReadHostName(CStdioFile *s_file,CString cmd1,CString cmd2);  //��ȡ����hostname
	int ReadHardware(CString instr,CString &slot,CString &cardtype,CString &sn,CString &date);  //��ȡshow hard��Ϣ
	CString ReadFailString(CStdioFile *s_file,CString cmdstr);  //���塢����Ӳ�����
	bool ReadCircuitSum(CStdioFile *s_file,CString cmdstr,CStringArray &rstr);
	CString ReadRedundancy(CStdioFile *s_file,CString cmd1,CString cmd2,CString cmd3);
	CString ReadProcess(CStdioFile *s_file,CString cmdprocess);
	CString ReadMemory(CStdioFile *s_file,CString cmdmem);
	CString ReadDisk(CStdioFile *s_file,CString cmddisk);
	CString ReadPortLinkDampening(CStdioFile * s_file,CString shportde);
	CString ReadPortAutoNegotiate(CStdioFile * s_file,CString shportde);
	CString ReadSfp(CStdioFile *s_file,CString cmdsfp);
	void ReadDiagPod(CStdioFile *s_file,CString cmdstr,CStringArray &rstr);

};
