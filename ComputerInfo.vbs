CONST AS400FilePath = "C:\Program Files\IBM\Client Access\Emulator\Private\"
'CONST AS400FilePath = "C:\"
CONST DestinationPath = "C:\OCS SETUP\"
 
Const FOR_READING = 1
Const FOR_WRITING = 2
Const FOR_APPENDING = 8 
 
on error resume next
Dim FSO
Dim DesktopPath,WshNetwork
 
' ---------------------------------------------------
' Read AS/400 WorkStationID
' ---------------------------------------------------
Function ReadAS400WorkStationID(fName)
          Dim AS400FSO
          Dim f
          Dim line,data
          Dim count
          'Wscript.Echo "File : " & fName
          Set AS400FSO = CreateObject("Scripting.FileSystemObject") 
          set f = AS400FSO.OpenTextFile(fName , FOR_READING)
          Do Until f.AtEndOfStream 
                   line= f.ReadLine
                   data = split(line,"=")
                   IF (data(0) = "WorkStationID") THEN
                             ReadAS400WorkStationID = line
                   END IF
          LOOP
          
END FUNCTION
 
 
 
 
' ---------------------------------------------------
' Get Computer information
'----------------------------------------------------
 
Function GetIPaddress() 
          dim arrIPaddr()
          dim i, Count, ArrNum
          Dim IPAddressString,strComputer
          Dim bolIsRightSubnet
    Dim objWMIService
    Dim IPConfig,IPConfigSet 
    Dim CmpCnt
    Dim CmpFlag
 
          bolIsRightSubnet = False
          strComputer = "."
          Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")
          Set IPConfigSet = objWMIService.ExecQuery("Select IPAddress from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
 
          ArrNum =0
          CmpCnt =0
          Count = 0
          For Each IPConfig in IPConfigSet
          If Not IsNull(IPConfig.IPAddress) Then
                   ArrNum = ArrNum + 1
          End If
          Next
 
          Redim arrIPaddr(ArrNum)
          IPAddressString=""
          For Each IPConfig in IPConfigSet
          If Not IsNull(IPConfig.IPAddress) Then
                   'If Not IsNull(IPConfig.IPAddress) Then 
                             For i=LBound(IPConfig.IPAddress) to UBound(IPConfig.IPAddress)
                'WScript.Echo IPConfig.IPAddress(i)
                                      CmpFlag = false
                                      if Count >= 1 then
 
                                                for CmpCnt = 0 to Count
                                                          if (IPConfig.IPAddress(i) = arrIPaddr(cmpCnt) ) then
                                                                   CmpFlag = true
                                                          end if
                                                next 
                                      end if
                                      
                                      if CmpFlag = false then
                                                arrIPaddr(Count) = IPConfig.IPAddress(i)
 
                                                if count=0 then
                                                          IPAddressString = IPConfig.IPAddress(0)
                                                else
                                                          IPAddressString = IPAddressString & vbCrLF & "                      " & IPConfig.IPAddress(i)
                                                end if
                                                Count = Count + 1             
                                      end if
                             Next
                   'End If
          End If
          Next
                   'GetIPaddress =   Join(arrIPaddr, vbCrLF & "                      ")
          GetIPaddress = IPAddressString
End Function
 
Dim WorkStationID
Dim objFSO,objFolders, colFiles, objFile
Dim FileExt
Dim Count
 
Set WshNetwork = WScript.CreateObject("WScript.Network")
 
Set objFSO = CreateObject("Scripting.FileSystemObject")
IF objFSO.FolderExists(AS400FilePath) Then
          Set objFolders = objFSO.GetFolder(AS400FilePath)
          Set colFiles = objFolders.Files
          Count =0 
 
          For Each objFile in colFiles
                   FileExt = right(objFile.Name,2)
                   IF FileExt = "WS" THEN
                             Count = Count + 1
                             WorkStationID = WorkStationID & vbCrLf & "                  " & ReadAS400WorkStationID (AS400FilePath & "\" & objFile.Name)
                   END IF
          NEXT  
ELSE
          WorkStationID="NO System i"
END IF
'OS and version
strComputer = "."
Set objWMIService = GetObject("winmgmts:" _
    & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colSettings = objWMIService.ExecQuery _
    ("Select * from Win32_OperatingSystem")
For Each objOperatingSystem in colSettings 
   OSName = objOperatingSystem.Name
    ServicePack = objOperatingSystem.ServicePackMajorVersion _
            & "." & objOperatingSystem.ServicePackMinorVersion
next
 
tmp = cstr(OSName)
OSname = split(OSName,"|")
 
' CPU SPeed--------------------------------------------------------------
dim cpu
Set colItems = objWMIService.ExecQuery("Select * from Win32_Processor")
For Each objItem in colItems
'Wscript.Echo "CPU: " & objItem.Name
cpu = objItem.Name
next
 
'Ram---------------------------------------------------------------------
dim ram,ram1,ram2,tmp
Set colSettings = objWMIService.ExecQuery("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings 
    'Wscript.Echo "Total Physical Memory: " & objComputer.TotalPhysicalMemory
ram = objComputer.TotalPhysicalMemory
'ram=ram/1048576
ram = Round(FormatNumber(ram / 1024 / 1024 / 1024, 2), 2)
'tmp=cstr(ram1)
'ram = split(tmp,".")
ram=FormatNumber(ram,1)
Next
 
'hard disk----------------------------------------------------------------
dim harddisk ,num

Const wbemFlagReturnImmediately = &h10
Const wbemFlagForwardOnly       = &h20
dim DriveInfor

strComputer = "."

Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")
Set colInstances = objWMIService.ExecQuery( "SELECT * FROM Win32_LogicalDisk", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly )
For each objDiskDrive in colInstances

	if instr(objDiskDrive.description,"Local Fixed Disk")then
		'wscript.echo objDiskDrive.DeviceID
		DriveLetter = objDiskDrive.DeviceID
		'Get Free Space
		FreeSpace = ConvertSize(objDiskDrive.FreeSpace)
		'Get Disk Size
		DiskSize  = ConvertSize(objDiskDrive.Size)
		'Get Used Size
		UsedSize = ConvertSize(objDiskDrive.Size - objDiskDrive.FreeSpace)
		'Get full diskPlay
		DriveInfor = DriveInfor & "Drive_"& DriveLetter & ">" & vbTab & "Size: "& DiskSize & vbTab &"Used: " & UsedSize & vbTab & "Free: " & FreeSpace & vbCrLf
		
	elseif instr(objDiskDrive.DeviceID,"H:") then
		'Get Home Size
		HomeSize  = ConvertSize(objDiskDrive.Size)	
		Set fso = CreateObject("Scripting.FileSystemObject")	 
        Set f = fso.GetFolder(objDiskDrive.DeviceID)
		'Get Used Home Size
		HomeDriveUsed =  ConvertSize(f.size) 
        'Get Free Home Size 
		HomeDriveFree =  ConvertSize(objDiskDrive.Size - f.size)     
 
	end if
	
next

filespec = "H:\" 
 
Dim  f, s

Set fso = CreateObject("Scripting.FileSystemObject")
 
	if fso.DriveExists(filespec) then   
        Set f = fso.GetFolder(filespec)
		 
        s = "My_Drive Thin:>   Size: " & HomeSize & vbTab & "Used: " & HomeDriveUsed & vbTab & "Free: " & HomeDriveFree
 
	else
        s = "My_Drive Thin used : Not have My_Drive Thin(H:)"
end if




function ConvertSize(Bytes)	

		If Bytes >= 1073741824 Then
			SetBytesC = Round(FormatNumber(Bytes / 1024 / 1024 / 1024, 2), 2) & " GB"
		ElseIf Bytes >= 1048576 Then
			SetBytesC = Round(FormatNumber(Bytes / 1024 / 1024, 2), 2) & " MB"
		ElseIf Bytes >= 1024 Then
			SetBytesC = Round(FormatNumber(Bytes / 1024, 2), 2) & " KB"
		ElseIf Bytes < 1024 Then
			SetBytesC = Bytes & " Bytes"
		Else
			SetBytesC = "0 Bytes"
		End If

		ConvertSize = SetBytesC
		
End function


 
WScript.Echo "User Information" & vbCrLF _
          & "Computer Name = " & WshNetwork.ComputerName & vbCrLF _
          & "User Name = " & WshNetwork.UserName & vbCrLf _
          & "IP Address = " & GetIPaddress & vbCrLf _
          & "Domain = " & WshNetwork.UserDomain  & vbCrLf _
          & "System i WorkStation ID : " & WorkStationID& vbCrLf _
          & "OS Name = " & OSname(0) & vbCrLf _
          & "Service Pack = "&ServicePack & vbCrLf _
          & "CPU = "&cpu & vbCrLf _
          & "RAM = "&ram& " GB"& vbCrLf & vbCrLf _
          & DriveInfor & vbCrLf _
          & s
          
 
' -----------------------------------------------------
' End of Get computer information
' -----------------------------------------------------
 
 