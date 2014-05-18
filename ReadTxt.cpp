#include "stdafx.h"
#include "ReadTxt.h"

ReadTxt::ReadTxt(void)
{
}

ReadTxt::~ReadTxt(void)
{
}

CString ReadTxt::ReadLine(CStdioFile *s_file,CString findstr)
{
	CString str;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(findstr) > -1) break;
	return str.Trim();
}

CString ReadTxt::ReadLines(CStdioFile *s_file,CString startstr)
{
	CString str,rstr;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(startstr) > -1) break;
	while(s_file->ReadString(str) && str.Find("#") == -1)
	{
		rstr += str.Trim();
		rstr += " \r\n";
	}
	return rstr.Trim();
}

CString ReadTxt::ReadLines(CStdioFile *s_file,CString startstr,CString endstr)
{
	CString str,rstr;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(startstr) > -1) break;
	while(s_file->ReadString(str) && str.Find(endstr) == -1)
	{
		rstr += str.Trim();
		rstr += " \r\n";
	}
	return rstr.Trim();
}

CString ReadTxt::ReadNextLine(CStdioFile *s_file,CString startstr,int next)
{
	CString str;
	int i = next;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(startstr) > -1) break;
	for( ; i > 0 && s_file->ReadString(str); i--) 
		if( str.Find("#") > -1 )
		{
			str = _T("");
			break;
		}
	return str.Trim();
}

CString ReadTxt::ReadNextLines(CStdioFile *s_file,CString startstr,int next)
{
	CString str,rstr;
	int i = next;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(startstr) > -1) break;
	for( ; i>1; i--) s_file->ReadString(str);
	while(s_file->ReadString(str) && str.Find("#") == -1)
	{
		rstr += str;
		rstr += " \r\n";
	}
	return rstr.Trim();
}

CString ReadTxt::ReadHostName(CStdioFile *s_file,CString cmd1,CString cmd2)
{
	CString str;
	//ReadTxt t;
	str = ReadLine(s_file,cmd1);
	if (str.Find("[local]") == -1) 
	{
		str = ReadLine(s_file,cmd2);
		cmd1 = cmd2;
		//t.~ReadTxt();
	}
	if (str.Find("[local]") > -1 )
	{
		str.Replace("[local]","");
		str.Replace("#","");
		str.Replace(cmd1,"");
	}
	else str = _T("");
	return str.Trim();
}

int ReadTxt::ReadHardware(CString instr,CString &slot,CString &cardtype,CString &sn,CString &date)
{
	int token =0;
	if( "" == instr ) return 0;
	slot = instr.Tokenize( _T(" "),token);
	cardtype = instr.Tokenize( _T(" "),token);
	if( cardtype == "fan" ) 
	{
		cardtype = _T("fan tray");
		sn = instr.Tokenize( _T(" "),token);
	}
	if( cardtype == "unknown" )
	{
		sn = _T("N/A");
		date = _T("N/A");
	}
	else 
	{
		sn = instr.Tokenize( _T(" "),token);
		date = instr.Tokenize( _T(" "),token);
		while( date.Find("-") == -1) date = instr.Tokenize(_T(" "),token);
	}
	return 1;
}

void ReadTxt::ReadDiagPod(CStdioFile *s_file,CString cmdstr,CStringArray &rstr)
{
	CString str;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(cmdstr) > -1) break;
	s_file->ReadString(str);
	s_file->ReadString(str);

	while(s_file->ReadString(str) && str.Find("#") == -1)
	{
		if(str.Find("FAIL") > -1) rstr.Add("FAIL");
		else rstr.Add("PASS");
	}
}

CString ReadTxt::ReadFailString(CStdioFile *s_file,CString cmdstr)
{
	CString str,rstr;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(cmdstr) > -1) break;
	while(s_file->ReadString(str) &&str.Find("#") == -1)
	{
		str.MakeLower();
		if(str.Find("fail") > -1)
		{
			rstr += str.Trim();
			rstr += "\r\n";
		}
	}
	if(rstr == "") rstr=_T("PASS");
	return rstr.Trim();
}

bool ReadTxt::ReadCircuitSum(CStdioFile *s_file,CString cmdstr,CStringArray &rstr)
{
	CString str,strs;
	bool rtrn = false;
	int i=0;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(cmdstr) > -1) 
		{
			rtrn = true;
			break;
		}
	if(rtrn)
	{
		for(i = 0;i < 3;i++)
		{
			s_file->ReadString(str);
			strs += str;
		}
		int token = 0;
		for( i = 0; i < 14; i++)
		{
			strs.Tokenize(" ",token);
			str = strs.Tokenize(" ",token);
			str.Replace(",","");  //去掉逗号
			rstr.Add(str);
		}
	}
	return rtrn;
}

CString ReadTxt::ReadRedundancy(CStdioFile *s_file,CString cmd1,CString cmd2,CString cmd3)
{
	CString str1,str2,rstr;
	s_file->SeekToBegin();
	while(s_file->ReadString(str1))
		if(str1.Find(cmd1) > -1) break;
	s_file->ReadString(str1);
	str1.Trim();
	s_file->ReadString(str2);
	s_file->ReadString(str2);
	str2.Trim();
	if(str1 == "0" && str2 == "0") rstr = "OK";
	else return "FAIL";
	while(s_file->ReadString(str1))
		if(str1.Find(cmd3) > -1) break;
	while(s_file->ReadString(str1) && str1.Find("#") == -1)
	{
		if(str1.Find("Switchover") > -1) 
			rstr += _T("\r\nxcrp Switchover");
		if(str1.Find( ")->(" ) > -1)
			rstr += "\r\n" + str1;
	}
	return rstr;
}


CString ReadTxt::ReadProcess(CStdioFile *s_file,CString cmdprocess)
{
	CString str,str_u,rstr;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(cmdprocess) > -1) break;
	s_file->ReadString(str);
	int token = 53;
	while(s_file->ReadString(str) && str.Find("#") == -1)
	{
		token = 53;
		str_u = str.Tokenize(" ",token);
		if(str_u.Find("%") > -1) 
		{
			str_u.Replace("%","");
			str_u.Trim();
			if(atof(str_u) > 5.0) 
			{
				token = 0;
				rstr += str.Tokenize(" ",token);
				rstr += _T(" ");
				token = 53;
				rstr += str.Tokenize(" ",token);
				rstr += "\r\n";
			}
		}
		else continue;
	}
	return rstr;
}

CString ReadTxt::ReadMemory(CStdioFile *s_file,CString cmdmem)
{
	CString str,m_total,m_used,rstr;
	while(s_file->ReadString(str))
		if(str.Find("#") > -1 && str.Find(cmdmem) > -1) break;
	s_file->ReadString(str);
	if(str.Find("k,") > -1)
	{
		int token = 0;
		do{
			m_total = str.Tokenize(" ",token);
		}while(m_total.Find("k,") == -1 && token > -1) ;

		m_total.Replace("k,","");
		m_total.Replace(" ","");
		do{
			m_used = str.Tokenize(" ",token);
		}while(m_used.Find("k,") == -1) ;
		m_used.Replace("k,","");
		m_used.Replace(" ","");

		rstr.Format("%.2f",atof(m_used)/atof(m_total)*100);
		return rstr+_T("%");
	}
	else return "Error";
}


CString ReadTxt::ReadDisk(CStdioFile *s_file,CString cmddisk)
{
	CString str,rstr;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(cmddisk) > -1) break;
	s_file->ReadString(str);
	rstr += _T("Internal ");
	int token = 41;
	rstr += str.Tokenize(" ",token);
	rstr += "\r\n";

	s_file->ReadString(str);
	rstr += _T("External ");
	if(str.Find("%") > -1)
	{
		token = 41;
		rstr += str.Tokenize(" ",token);
	}
	else rstr += _T(" Error");
	return rstr;
}

CString ReadTxt::ReadPortLinkDampening(CStdioFile * s_file,CString shportde)
{
	CString str,porttype,nport,linkd;
	int token;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(shportde) > -1) break;
	while(s_file->ReadString(str) && str.Find("#") == -1)
	{
		if(str.Find("state is Up") > -1 && str.Find("7/1") == -1) 
		{
			token = 0;
			porttype = str.Tokenize(" ",token);
			nport = str.Tokenize(" ",token); //取端口号
			s_file->ReadString(str);
			str.Replace(" ","");
			str.MakeLower();
			if(str.Find("7609")>-1 || str.Find("8512")>-1 || str.Find("9312")>-1 || str.Find("7810")>-1 ||
				(str.Find("crs")==-1 && str.Find("csr")==-1 && str.Find("5k")==-1 && str.Find("5000")==-1 && str.Find("gsr")==-1))
			{  //判断为下行端口就检查link-dampening是否开启
				while(s_file->ReadString(str))
					if(str.Find("Link Dampening") > -1) break;
				if(str.Find("disabled") > -1)
				{
					linkd += nport;
					linkd += _T(" ");
				}
			}
		}
	}
	if(linkd == "") linkd = _T("是");
	else linkd += _T("\r\n未开启");
	return linkd;
}

CString ReadTxt::ReadPortAutoNegotiate(CStdioFile * s_file,CString shportde)
{
	CString str,porttype,nport,autonegotiate;
	int token;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(shportde) > -1) break;
	while(s_file->ReadString(str) && str.Find("#") == -1)
	{
		if(str.Find("state is Up") > -1) 
		{
			token = 0;
			porttype = str.Tokenize(" ",token);
			nport = str.Tokenize(" ",token); //取端口号
			
			if(porttype.Find("atm") > -1) continue;  //atm板卡无协商相关
			while(s_file->ReadString(str) && str.Find("Active Alarms") == -1)
				if(str.Find("negotiation") > -1) break;
			if(str.Find(": on") > -1)
			{
				autonegotiate += nport;
				autonegotiate += _T(" ");
			}
		}
	}

	if(autonegotiate == "") autonegotiate = _T("是");
	else autonegotiate += _T("\r\n为auto-negotiate");
	return autonegotiate;
}

CString ReadTxt::ReadSfp(CStdioFile *s_file,CString cmdsfp)
{
	CString str,cardtype,slot,port,rstr;
	int token=0;
	s_file->SeekToBegin();
	while(s_file->ReadString(str))
		if(str.Find(cmdsfp) > -1) break;
	s_file->ReadString(str);
	if(str.Trim() == "0") return _T("是");

	s_file->ReadString(str);
	while(s_file->ReadString(str) && str.Find("#") == -1)
	{
		if(str.Find("Slot") > -1)
		{
			token = 62;
			cardtype = str.Tokenize(" ",token); //取板卡类型
			cardtype.Trim();
			if(cardtype.Find("10-port") > -1 || cardtype.Find("10ge") > -1) //万兆和10-port支持查看
			{
				token = 22;
				slot = str.Tokenize(" ",token); //取槽位号
				slot.Trim();
				while(s_file->ReadString(str) && str.Find("[local]") == -1)
				{
					if(str == "")
					{
						s_file->ReadString(str);
						if(str == "" || str.Find("[local]") > -1) break;
						if(str.Find("Port ") > -1)
						{
							token = 22;
							port = str.Tokenize(" ",token);  //取端口号
							port.Trim();
							while(s_file->ReadString(str) && str.Find("[local]") == -1)
							{
								if(str.Find("RedbackApproved") > -1)
								{
									if(str.Find("No") > -1)	rstr += slot + _T("/") + port + _T(" ");
									break;
								}
							}
						}
					}
				}
			}
		}
	}
	if(rstr == "") rstr = _T("是");
	else rstr += _T("\r\n不支持");
	return rstr;
}
