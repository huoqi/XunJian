#$language = "VBScript"
#$interface = "1.0"

' SecureCRT script to automate the inventory retrieval on Ericsson SmartEdge
' Works with SecureCRT version:
'   v5.2.2
'   v6.1.4
' Does NOT work with version:
'   v5.0.3
'   v6.5

' Needed library : none
' COMMAND LINE PARAMETERS : none
' INPUT : a plain text file with one SmartEgde IP address or hostname per line
' OUTPUT : a plain text file with all the command ouput, one file per node

'Version History
' Version 2.1 2011-12-07:
' adding additional commands
' Version 2.0 2011-11-01:
' adding support for Juniper 
' Version 0.1.1 2010-12-15:
'	adding choice for telnet/ssh connections
' Version 0.1.0 2010-12-15:
'	Initial version

' Constants for setting MessageBox options
Const ICON_STOP = 16		' Critical message; displays STOP icon.
Const ICON_QUESTION = 32	' Warning query; displays '?' icon.

Const BUTTON_OK = 0		' OK button only

Const DEFBUTTON1 = 0	' First button is default

Const IDOK = 1		' OK button clicked

Function BrowseWithIE
    ' Uses "FindFile()" class that employs IE to perform a
    ' file browse operation.  FindFile class is referenced
    ' below.
    Set objBrowser = New FindFile
    BrowseWithIE = objBrowser.Browse
'    if szFile = "" then
'        MsgBox "Cancel"
'    else
'        MsgBox szFile
'    end if

End Function

' FindFile class part of script example "Reboot A
' List Of Workstations" posted to the Win32 Scripting
' examples site (www.netreach.net), but modified to work
' within a script hosted by SecureCRT/CRT:
'    - removed "WScript." from function calls like "CreateObject()"
'    - Changed "WScript.Sleep 500" to "crt.Sleep 500"
'    - removed "Cancel" button from the IE temp FileBrowser.html to
'              eliminate security warnings from IE about scripts that
'              could potentially do damage.  Cancel can still be
'              achieved by clicking on the 'X' to close the window.
'
' script code base for FindFile from:
'   http://cwashington.netreach.net/depo/view.asp?Index=922&ScriptType=vbscript

Class FindFile
            Private fso, sPath1

      '--FileSystemObject needed to check file path:
           Private Sub Class_Initialize()
                Set fso = CreateObject("Scripting.FileSystemObject")
           end sub

           Private Sub Class_Terminate()
              Set FSO = Nothing
           End sub

      '-- the one public function in class:

           Public Function Browse()
              'on error resume next
               sPath1 = GetPath
               Browse = sPath1
           end function

Private Function GetPath()
    Dim Ftemp, ts, IE, sPath, sStatus

    '--Get the TEMP folder path and create a text file in it:

        Ftemp = fso.GetSpecialFolder(2)
        Ftemp = Ftemp & "\FileBrowser.html"
        set ts = fso.CreateTextFile(Ftemp, true)

    '--write the webpage needed for file browsing window:
        ts.WriteLine "<HTML><HEAD><TITLE></TITLE></HEAD>"
        ts.WriteLine "<BODY BGCOLOR=" & chr(34) & "#F3F3F8" & chr(34) & " TEXT=" & chr(34) & "black" & chr(34) & ">"
        ts.WriteLine "<DIV ALIGN=" & chr(34) & "center" & chr(34) & ">"
        ts.WriteLine "<FONT FACE=" & Chr(34) & "arial" & Chr(34) & " SIZE=2>"
        ts.WriteLine "<B>Select File with SmartEdge list:</B><BR>"
        ts.WriteLine "<FORM>"

                     '--this is the file browsing box in webpage:
                 ts.WriteLine "<INPUT TYPE=" & chr(34) & "file" & chr(34) & "></input>"
                      '--this puts a bit of space between buttons:
              ts.WriteLine chr(38) & chr(35) & "160" & chr(59) & " " & chr(38) & chr(35) & "160" & chr(59)
        ts.WriteLine "<BR>"
        ts.WriteLine "</FORM>"
        ts.WriteLine "</FONT></DIV>"
        ts.WriteLine "</BODY></HTML>"
        ts.Close
        set ts = nothing

    '--webpage is written. now have IE open it:
        Set IE = CreateObject("InternetExplorer.Application")
        IE.Navigate "file:///" & Ftemp
        IE.AddressBar = false
        IE.menubar = false
        IE.ToolBar = false
        IE.StatusBar = false
        IE.width = 400
        IE.height = 150
        IE.resizable = true
        IE.visible = true

    '--do a loop every 1/2 second until either:
    '-- the browsing window value is a valid file path or
    '--IE is closed from the control box.

        Do while IE.visible = true
            spath = ie.document.forms(0).elements(0).value         '--get browsing text box value.

            if fso.FileExists(spath) = true or IE.visible = false then exit do
            crt.Sleep 500
        Loop

        IE.visible = false
        IE.Quit
        set IE = nothing
        if fso.FileExists(spath) = true then
            GetPath =  spath
        else
            GetPath = ""
        end if
  End Function

End Class

Sub main

	Dim login, pass, enablepass,SEList, JumboxPrompt, result, method, vendor, promptString

	login = "admin"
	pass = "redback"
	enablepass = "redback"
	SEList = "IP_list.txt"
	method = "telnet"
	JumpboxPrompt = "#"

	Dim connect

	' Define the way to connect on the node itself
	connect = "telnet"
	' Set connect = "ssh"

	'Ask for the vendor of the equipment you are going to connect to
    'result = crt.Dialog.Prompt("Vendor/Type of all nodes (1-SE, 2-JUNIPER SRX/M SERIES", "vendor", "1", False)
	'result = crt.Dialog.Prompt("Vendor/Type of all nodes (1-SE, 2-JUNIPER SRX/M SERIES, 3-JUNIPER NETSCREEN", "vendor", "1", False)
	
	'if result = "1" then
	vendor = "SE"
	promptString = "[local]"
                'else 
                  'result = "2" 
	  'vendor = "JUNIPER"
	  'promptString = ">"
  	'end if 
  
  'Ask for the file containing the list of nodes to connect to
	SEList = BrowseWithIE

	if SEList = "" then
		result = crt.Dialog.MessageBox("You didn't select a file!", "Exiting script", BUTTON_OK Or ICON_STOP)
		crt.Quit
	end if

	'Ask for the login to connect on all the nodes
	result = crt.Dialog.Prompt("Enter login", "Login", "admin", False)
	
	if result <> "" then
		login = result
	end if
	
	'Ask for the password to connect on all the nodes
	result = crt.Dialog.Prompt("Enter password", "Password", "", True)
	
	if result <> "" then
		pass = result
	end if
	
	'Ask for the enable password to connect on all the nodes
	'result = crt.Dialog.Prompt("Enter enable password", "Password", "", True)
	
	'if result <> "" then
		'enablepass = result
	'end if
	
	'Ask for the connection method to all the nodes
	' text = "Connection method to the SE " & VbCr & "1 for telnet"  & VbCr & "2 for SSH from a Unix host (ssh <login>@1.1.1.1)" & VbCr & "3 for SSH from a SmartEdge (ssh 1.1.1.1 admin <login>)"
	result = crt.Dialog.Prompt("Connection method to the node (1-telnet 2-ssh from unix 3-ssh from node)", "Connection Method", "1", False)
	
	if result = "2" Or result = "3" then
		method = "ssh"
	end if
	
	' turn on synchronous mode so we don't miss any data
	crt.Screen.Synchronous = True
	
	dim nRow
	
	' Get the shell prompt so that we can know what to look for when
    ' determining if the command is completed. Won't work if the prompt
    ' is dynamic (e.g. changes according to number of commands entered, etc)
    crt.Screen.Send VbCr
	nRow = crt.Screen.CurrentRow
    JumpboxPrompt = crt.screen.Get(nRow, 0, nRow, crt.Screen.CurrentColumn - 1)
    JumpboxPrompt = Trim(JumpboxPrompt)
	
	Const ForReading = 1, ForWriting = 2

	Dim fso, MyFile, mydate

	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.OpenTextFile(SEList, ForReading)

	Do While MyFile.AtEndOfStream <> True

		Host = MyFile.ReadLine
		
		' Try to ping the loopback address
		' crt.Screen.Send "ping " & Host & VbCr
		' If crt.screen.WaitForString("alive", 30) <> True Then			
			' crt.Dialog.MessageBox(Host & " is not reachable. Passing to next host.", "Host unreachable", BUTTON_OK Or ICON_STOP)
			
		' Else
			'Desactivate the logging if it's already enabled
			crt.Session.Log(False)
			
			if result = "2" then
				' Connect through SSH to the node
				crt.Screen.Send method & " " & Host & " " & login & VbCr
			else 
				if result = "3" then
					' Connect through SSH to the node
					crt.Screen.Send method & " " & Host & " admin " & login & VbCr
				else 
					' Connect through telnet to the node
					crt.Screen.Send method & " " & Host & VbCr
					
					' Wait for the string login
					if crt.Screen.WaitForString ("login:", 50) = True then
					
						' get the current date and time
						'mydate = Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) 
                        '& "_" & Hour(Now()) & "-" & Minute(Now())
						
						'activate the logging with the right filename
						crt.Session.LogFileName = Host & ".txt"
						crt.Session.Log(True)
						
						' Send login followed by a carriage return
						crt.Screen.Send login & VbCr
						
					else
						' get the current date and time
						'mydate = Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) 
                        '& "_" & Hour(Now()) & "-" & Minute(Now())
						
						'activate the logging with the right filename
						crt.Session.LogFileName =  Host & ".txt"
						crt.Session.Log(True)
						
						'Send the escape character "^C" or decimal value "003" to terminate the telnet session
						crt.screen.Send Chr(3)
						crt.Screen.Send VbCr
						crt.Screen.WaitForString JumpboxPrompt
						
					end if					
				end if
			end if

				' Wait for the string password: or Password:
			if crt.Screen.WaitForString ("assword:", 30) = True then
				
				if method = "ssh" then
						' get the current date and time
						'mydate = Year(Now()) & "-" & Month(Now()) & "-" & Day(Now()) 
						'& "_" & Hour(Now()) & "-" & Minute(Now())
						
						'activate the logging with the right filename
					crt.Session.LogFileName = Host  & ".txt"
					crt.Session.Log(True)
				end if
						
					' Send password followed by a carriage return
				crt.Screen.Send pass & VbCr
				'crt.Screen.WaitForString promptString
						 
				 ' SmartEdge commands
				'crt.Screen.Send "enable" & chr(13)
				'crt.Screen.WaitForString "Password: "
				
				'crt.Screen.Send enablepass & chr(13)

				Do while crt.Screen.WaitForString(">", 5) = true
				'Ask for the enable to connect on all the nodes
					result = crt.Dialog.Prompt("Please enable password", "password", "", True)
					if result <> "" then
						enablepass = result
					end if
					crt.Screen.Send "enable" & chr(13)
					crt.Screen.WaitForString "Password: "
				
					crt.Screen.Send enablepass & chr(13)
					crt.Screen.WaitForString promptString
				Loop
					
				crt.Screen.Send "term len 0" & VbCr 
				crt.Screen.WaitForString promptString   
                                    
				crt.Screen.Send "show clock" & VbCr 
				crt.Screen.WaitForString promptString   
                                      
				crt.Screen.Send "show hard | begin N/A" & VbCr 
				crt.Screen.WaitForString promptString   
										
				crt.Screen.Send "show hard de | include Temp | exclude 'Temp.- Inlet Air' | exclude Wavelength" & VbCr
				crt.Screen.WaitForString promptString   
										
				crt.Screen.Send "show diag pod backplane de | begin Back" & VbCr 
				crt.Screen.WaitForString promptString  

				crt.Screen.Send "show diag pod fantray de | begin Fan" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show diag pod | begin N/A" & VbCr 
				crt.Screen.WaitForString promptString 
										
				crt.Screen.Send "show pppoe sub 1 | count" & VbCr 
				crt.Screen.WaitForString promptString 
	  
				crt.Screen.Send "show pppoe sub 2 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 3 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 4 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 5 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 6 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 7 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 8 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 9 | count" & VbCr 
				crt.Screen.WaitForString promptString

				crt.Screen.Send "show pppoe sub 10 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 11 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 12 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 13 | count" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show pppoe sub 14 | count" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 1 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 2 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 3 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 4 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 5 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 6 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 7 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 8 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 9 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 10 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 11 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 12 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 13 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show circuit 14 sum | include total" & VbCr 
				crt.Screen.WaitForString promptString 

				crt.Screen.Send "show sub sum all | include Total" & VbCr 
				crt.Screen.WaitForString promptString     
	 
				crt.Screen.Send "show circuit sum | include total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show ip route sum all | include Subscriber " & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show licenses | include Total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show system status" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show redundancy | inc NO | count" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show redundancy | inc FAIL | count" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show crash" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show process | begin NAME" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show process cpu | include Total" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show disk | begin Internal" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show port detail" & VbCr 
				crt.Screen.WaitForString promptString      

				crt.Screen.Send "show hard detail | include RedbackApproved | exclude Yes | count" & VbCr 
				crt.Screen.WaitForString promptString   

				crt.Screen.Send "show hard detail" & VbCr 
				crt.Screen.WaitForString promptString
									
				crt.Screen.Send "show tech" & VbCr  


			end if 
					
			Do While crt.screen.WaitForString(promptString, 50) <> True
				crt.Screen.Send " "
			Loop
					
		' Logout of the node
		if method = "telnet" then
			crt.Screen.Send "exit" & VbCr
			crt.Screen.WaitForString JumpboxPrompt					
			crt.Sleep 500
					
		else 
			if method = "ssh" then
				'Send the escape character "^C" or decimal value "003" to terminate the telnet session
				crt.screen.Send Chr(3)
				crt.Screen.Send VbCr
				crt.Screen.WaitForString JumpboxPrompt
			end if
		End If

	Loop

	'Disable logging
	crt.Session.Log(False)

	'Close the File
	MyFile.Close

	' turn off synchronous mode to restore normal input processing
	crt.Screen.Synchronous = False
  
End Sub
