Attribute VB_Name = "Module1"
' ##############################################################################################
' OpenVPN configuration file CONVERTER for OpenwWRT
'
' /*
' * Copyright (C) 2019 Robert Nio (06-14-2019)
' *
' * Licensed under the Apache License, Version 2.0 (the "License");
' * you may not use this file except in compliance with the License.
' * You may obtain a copy of the License at
' *
' * http://www.apache.org/licenses/LICENSE-2.0
' *
' * Unless required by applicable law or agreed to in writing, software
' * distributed under the License is distributed on an "AS IS" BASIS,
' * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' * See the License for the specific language governing permissions and
' * limitations under the License.
' */
'
' ##############################################################################################
'
'#  OpenVPN configuration file CONVERTER for OpenWRT
'
' This will CONVERT your ".OVPN" configuration files from "<vpn config file>.ovpn" to something, which
' OpenWRT can work with. You can select (within this code) to create single CONF files for each of the OVPN files. (default is ONE BIG config file)
'
' The current configuration will convert the ".ovpn" files into a config similar to "/etc/config/openvpn" and merge this config file with your existing "/etc/config/openvpn" file.
'
' It will create crypto information from any of your ".ovpn" files (e.g. client.key / client.cert etc.) This information is copied out of the ".ovpn" files and put in separate files under a crypto directory
'
' Once all that is done, it will "SFTP" (Secure Copy from Windows to Router using SSH) all files to the router. All files will land in "/etc/openvpn".
'
' Then it will merge the created "openvpn.conf" files with your "/etc/config/openvpn" file
'
'### Why NOT use the NEW ,ovpn file upload feature within OpenWRT 19.07?
'
'       This can process and upload MULTIPLE files at once.
'
'       Easy to cut out functionality to REMOVE those configurations (see in VBA code)
'
'       I had only 18.06 to work with a couple days ago ... thus whipped up this macro
'       (NO Upload feature available in 18.06)
'
' ## Features:
'
' - Reads .ovpn files / extracts information / creates corresponding OpenWRT config file and insert it into /etc/config
' - Works with a directory FULL of separate .ovpn files
' - Creates individual files with KEY / CERT / AUTH info from ,ovpn files and stores them in CRYPTO dir
' - Tested with OpenWRT 18.06 (KongPRO) / PuTTY 0.70 / Excel 2016 /
' - "Plays" well on your systems, does not leave any crumbs behind.
'
' ## WHAT YOU NEED:
'
' - EXCEL
' - current copy of PuTTY & TOOLS (plink / psftp): http://www.putty.org/ (Tested with Version: 0.70)
' - The OpenWRT router Version 18.06 (This is the one I tested this on. Other versions should probably work as well)
'
'    ### WHY EXCEL?
'    It provides an easy, interactive IDE.
'
'    One could wrap this into a compiled exe ... just need to create  some forms etc. to allow basic interactions (e.g. configure settings / start conversion etc.).
'
'     If I would do it ... I would definitely have all settings in an external TXT file ...
'
'     I work with Excel almost everyday ... so just adding another button to do some router magic is not a big deal :)
'
'## GET STARTED:
'
'    Create / Open an EMPTY Excel Workbook
'
'    Create a SIMPLE Macro within Excel ... just start recording a macro ...change a cell ... stop recording.
'    Make sure you have the macro STORED with the current / new workbook.
'
'    Once that is done, Exel will have created a MODULE with a simple macro.
'
'    Press ALT-11 to invoke the IDE
'
'    look on your LEFT side, select the module1 which belongs to your just created workbook.
'
'    cut and paste the macro code from here into the module1 in excel.
'
'    CONFIRM / ADJUST the settings in this VBA module to YOUR environment
'
'    Execute the "Convert_OpenVPN_Config_Files" Macro
'
'### License
'----
'
'Apache License, Version 2.0

'
'
'
' ##############################################################################################
' ##############################################################################################
' ##############################################################################################
' ##############################################################################################
' ##############################################################################################
'
' CUSTOMIZATIONS:
'
' Please ADJUST / MODIFY these seeetings to reflect your environment
'
' ##############################################################################################
' ##############################################################################################
'
' The Family Jewels: All your CRYPTO settings
'
' IP of your router on your LAN network
' e.g. "192.168.99.1"
Public Const RouterIP = "<YOUR LAN IP>"

' We use SSH to access the router
' Set the PORT number
' e.g. "22"
Public Const RouterSSHPort = "<YOUR ROUTER SSH PORT>"

' We will use PuTTY saved sessions in place of hostnames.
' This allows us to keep session info within PuTTY and NOT in this code :)
' e.g. "PuTTY_seesion_name"
Public Const PuTTYSession = "<Putty SESSION NAME TO ACCESS YOUR ROUTER>"

' The ID / PW of your VPN provider
' These are used to verify the CERT / KEY etc. settings
' ==> additional layer of security
Public Const UserID = "<USER ID FROM VPN VENDOR>"
Public Const Password = "<PASSWORD FROM VPN VENDOR>"

' ##############################################################################################
' ##############################################################################################
'
' ON YOUR WINDOWS MACHINE:
'
' DIRECTORY LOCATION for PUTTY installation (on your WINDOWS system)
' e.g. "C:\Program Files\PuTTY"
Public Const PUTTYDIR = "<LOCATION OF PuTTY BINARIES>"

' DIRECTORY LOCATION for temp files (on your WINDOWS system)
' e.g. "C:\temp\"
Public Const TEMPDIRWIN = "<TEMP DIR ON WINDOWS MACHINE>"

' The LOCATION of the SOURCE directories
' e.g. "C:\OpenVpn\"
Public Const SOURCEDIR = "<LOCATION OF WHERE YOUR '.OVPN' FILES ARE>"

' The LOCATION of the TARGET directories
' e.g. "C:\OpenVpn\OpenWRT\"
Public Const TARGETDIR = TEMPDIRWIN & "converted\"
Public Const CRYPTODIR = TARGETDIR & "crypto\"

' The NAME of the resulting file with the OpenWRT configuration info
Public Const OpenWRT_Config_File_Name = "openvpn"

' The LOCATION of the resulting file with the OpenWRT configuration info
'Public Const OpenWRT_Config_File_Location = "/etc/config/" & OpenWRT_Config_File_Name
' e.g. "C:\OpenVpn\OpenWRT\openvpn"
Public Const OpenWRT_Config_File_Location = TARGETDIR & OpenWRT_Config_File_Name

' ##############################################################################################
'
' ON YOUR ROUTER:
'

' The "OpenVPN" file location on your ROUTER
' e.g. "/etc/openvpn/"
Public Const OpenWRT_OpenVPN_Location = "/etc/openvpn/"

' Where we place the CRYPTO files (Auth, Cert, Key etc.) on your router
' e.g. "/etc/openvpn/crypto/"
Public Const OpenWRT_Crypto_Location = "/etc/openvpn/crypto/"

' DIRECTORY LOCATION for temp files (on your ROUTER
' e.g. "/tmp/openvpn/"
Public Const TEMPDIRRT = "/tmp/openvpn/"



' ##############################################################################################
' ##############################################################################################
' ##############################################################################################
' ##############################################################################################
' ##############################################################################################


' ##############################################################################################
' Some DECLARATIONS needed for supervised command executions (remote SHELL-SCRIPTS on router)
' ##############################################################################################

Private Type STARTUPINFO
   cb As Long
   lpReserved As String
   lpDesktop As String
   lpTitle As String
   dwX As Long
   dwY As Long
   dwXSize As Long
   dwYSize As Long
   dwXCountChars As Long
   dwYCountChars As Long
   dwFillAttribute As Long
   dwFlags As Long
   wShowWindow As Integer
   cbReserved2 As Integer
   lpReserved2 As Long
   hStdInput As Long
   hStdOutput As Long
   hStdError As Long
End Type

Private Type PROCESS_INFORMATION
   hProcess As Long
   hThread As Long
   dwProcessID As Long
   dwThreadID As Long
End Type

Private Declare PtrSafe Function WaitForSingleObject Lib "kernel32" (ByVal _
   hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare PtrSafe Function CreateProcessA Lib "kernel32" (ByVal _
   lpApplicationName As Long, ByVal lpCommandLine As String, ByVal _
   lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
   ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
   ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As Long, _
   lpStartupInfo As STARTUPINFO, lpProcessInformation As _
   PROCESS_INFORMATION) As Long

Private Declare PtrSafe Function CloseHandle Lib "kernel32" (ByVal _
   hObject As Long) As Long
   
Private Declare PtrSafe Function FindWindow Lib "user32.dll" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Private Declare PtrSafe Function CloseWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&


Sub Convert_OpenVPN_Config_Files()
'
' Convert_OpenVPN_Config_Files Macro
'

'Get current state of various Excel settings; put this at the beginning of your code

screenUpdateState = Application.ScreenUpdating

statusBarState = Application.DisplayStatusBar

calcState = Application.Calculation

eventsState = Application.EnableEvents

'displayPageBreakState = ActiveSheet.DisplayPageBreaks 'note this is a sheet-level setting

'turn off some Excel functionality so your code runs faster

Application.ScreenUpdating = False

Application.DisplayStatusBar = False

'Application.Calculation = xlCalculationManual

Application.EnableEvents = False

'ActiveSheet.DisplayPageBreaks = False 'note this is a sheet-level setting

On Error GoTo Processing_Err

Const NO_FILES_IN_DIR As Long = 9
Const INVALID_DIR As Long = 13

Dim FSO As Object
Dim AlleFiles As Variant
Dim i, j As Integer

' The Variables for the OpenVPN file
'Dim InputDataLine As String
Dim OpenVPN_File As Integer
Dim OpenVPN_File_Name As String
Dim OpenVPN_File_Location As String
Dim InputBuffer As String ' can be up to 2 GB big :)
Dim InputLines() As String

' The Variables for the OpenWRT Config file
Dim OpenWRT_Config_File As Integer

' The Variables for the CONFIG File
Dim OutputDataLine As String
Dim Config_File As Integer
Dim Config_File_Name As String
Dim Config_File_Location As String
Dim ConfigBuffer As String ' can be up to 2 GB big :)

' The Variables for the AUTH file with ID/PW
Dim Config_Auth_File As Integer
Dim Config_Auth_File_Name As String
Dim Config_Auth_File_Location As String

' The Variables for the CERTIFICATE file
Dim Config_Cert_File As Integer
Dim Config_Cert_File_Name As String
Dim Config_Cert_File_Location As String

' The Variables for the KEY file
Dim Config_Key_File As Integer
Dim Config_Key_File_Name As String
Dim Config_Key_File_Location As String

' The Variables for the TLS-AUTH file
Dim Config_TLS_AUTH_File As Integer
Dim Config_TLS_AUTH_File_Name As String
Dim Config_TLS_AUTH_File_Location As String

' The Variables for the CERTIFICATE-AUTHORITY file
Dim Config_CA_File As Integer
Dim Config_CA_File_Name As String
Dim Config_CA_File_Location As String

Dim FileExtention As String

' OpenVPN configuration file VPN-location
Dim DeleteFolderName As String

Dim Elemente() As String

' We create a Scripting.FileSystemObject and FSO's reference count of this new object is now 1
Set FSO = CreateObject("Scripting.FileSystemObject")

' ##############################################################################################
' ##############################################################################################

' We process all '.ovpn" files and create the config file and crypto files based on these.

' ##############################################################################################
' ##############################################################################################

    ' We START with a CLEAN slate
    ' we will DELETE everything which was previously done on the WINDOWS machine.
    ' FSO.DeleteFolder(TARGETDIR) ' delete all files and directories in Target Location
    
    ' Set Object
    ' Clean the TARGETDIR
    DeleteFolderName = Left(TARGETDIR, Len(TARGETDIR) - 1)
    If FSO.FolderExists(DeleteFolderName) Then Result = FSO.DeleteFolder(DeleteFolderName, True)
    
    ' Clean the TEMPDIRWIN
    DeleteFolderName = Left(TEMPDIRWIN, Len(TEMPDIRWIN) - 1)
    If FSO.FolderExists(DeleteFolderName) Then Result = FSO.DeleteFolder(DeleteFolderName, True)
    
    
    ' Make sure we have a destination folders
    If Not (FSO.FolderExists(TEMPDIRWIN)) Then FSO.CreateFolder (TEMPDIRWIN)
    If Not (FSO.FolderExists(TARGETDIR)) Then FSO.CreateFolder (TARGETDIR)
    If Not (FSO.FolderExists(CRYPTODIR)) Then FSO.CreateFolder (CRYPTODIR)
    
    
    If FSO.FolderExists(SOURCEDIR) Then ' Check if we have a valid DIR
        
        ChDir SOURCEDIR
            
        AlleFiles = GetAllFilesInDir(SOURCEDIR)
        
        ' Put a DUMMY starter line in the beginning; this will help to better
        ' remove and insert the configuration without deleting other data in
        ' the config file.
        ' Create HEADER of conf file
        Config_File_Name = "openvpn" & ".conf"
        Config_File_Location = TARGETDIR & Config_File_Name
        Config_File = FreeFile
        Open Config_File_Location For Append As Config_File
        
        OutputDataLine = Chr(10) & "config" & Chr(9) & "openvpn" & Chr(9) & "'DUMMY_Start'" & Chr(10)
        OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "config" & Chr(9) & "'dummystart.conf'" & Chr(10) & Chr(10) & Chr(10)
        
        Print #Config_File, OutputDataLine;
            
        Close #Config_File
        
            
        For i = 0 To UBound(AlleFiles)
            
            'Check if we have a VALID OpenVPN file ... check etxtenson ".ovpn"
            FileExtention = Right(AlleFiles(i), Len(AlleFiles(i)) - InStrRev(AlleFiles(i), "."))
            
            If LCase(FileExtention) = "ovpn" Then
            
                ' This NEEDS to be adjusted for your specific VPN vendor
                ' Get the LOCATION name of this VPN config file
                ' My ovpn NAMEs are "my_expressvpn_argentina_udp.ovpn"
                OpenVPN_File_Name = Mid(AlleFiles(i), InStr(5, AlleFiles(i), "_", vbTextCompare) + 1, Len(AlleFiles(i)))
                OpenVPN_File_Name = Left(OpenVPN_File_Name, Len(OpenVPN_File_Name) - 5)
                OpenVPN_File_Name = Replace(OpenVPN_File_Name, "_-_", "_")
                OpenVPN_File_Location = SOURCEDIR & OpenVPN_File_Name
                
                ' ###########################################################
                ' We want all configurations in ONE file
                ' If INDIVIDUAL files are needed change the TWO lines below:
                
                'Config_File_Name = OpenVPN_File_Name & ".conf"
                Config_File_Name = "openvpn" & ".conf"
                
                ' ###########################################################
                
                Config_Auth_File_Name = OpenVPN_File_Name & ".auth"
                Config_Cert_File_Name = OpenVPN_File_Name & "_client.crt"
                Config_Key_File_Name = OpenVPN_File_Name & "_client.key"
                Config_TLS_AUTH_File_Name = OpenVPN_File_Name & "_ta.key"
                Config_CA_File_Name = OpenVPN_File_Name & "_ca2.key"
                
                Config_File_Location = TARGETDIR & Config_File_Name
                Config_Auth_File_Location = CRYPTODIR & OpenVPN_File_Name & "\" & Config_Auth_File_Name
                Config_Cert_File_Location = CRYPTODIR & OpenVPN_File_Name & "\" & Config_Cert_File_Name
                Config_Key_File_Location = CRYPTODIR & OpenVPN_File_Name & "\" & Config_Key_File_Name
                Config_TLS_AUTH_File_Location = CRYPTODIR & OpenVPN_File_Name & "\" & Config_TLS_AUTH_File_Name
                Config_CA_File_Location = CRYPTODIR & OpenVPN_File_Name & "\" & Config_CA_File_Name
                
                'READ the Source File
                OpenVPN_File_Location = SOURCEDIR & AlleFiles(i)
                OpenVPN_File = FreeFile
                Open OpenVPN_File_Location For Input As #OpenVPN_File
                '//load all
                InputBuffer = Input$(LOF(1), #OpenVPN_File)
                Close #OpenVPN_File
                
                
                'Test which line terminator is used.
                'CONVERT the BLOB of data into lines
                If InStr(1, InputBuffer, vbCrLf) > 0 Then
                        'This file uses CRLF as the line terminator
                        InputLines = Split(InputBuffer, vbCrLf)
                    
                    ElseIf InStr(1, InputBuffer, vbCr) > 0 Then
                        'This file uses CR as the line terminator
                        InputLines = Split(InputBuffer, vbCr)
                        
                    ElseIf InStr(1, InputBuffer, vbLf) > 0 Then
                        'This file uses LF as the line terminator
                        InputLines = Split(InputBuffer, vbLf)
                        
                    Else
                        'This file doesn't use a standard line terminator
                        GoTo Ende
                    
                End If
                
                ' Make sure we have CRYPTO folder
                If Not (FSO.FolderExists(CRYPTODIR & OpenVPN_File_Name)) Then FSO.CreateFolder (CRYPTODIR & OpenVPN_File_Name)
                
                
                
                ' Write the AUTH file with ID/PW
                Config_Auth_File = FreeFile
                Open Config_Auth_File_Location For Output As Config_Auth_File
                
                OutputDataLine = UserID & Chr(10) & Password & Chr(10)
                
                Print #Config_Auth_File, OutputDataLine;
                
                Close #Config_Auth_File
                
                
                ' ###########################################################
                
'                ' This seems NOT to work. The SAME config inside the "/etc/config/openvpn" file works
'                ' ... outside, called with "config  openvpn "argentina_udp"
'                '                                   option  config  "/etc/openvpn/argentina_udp.conf"
'                ' ==> does NOT work ?! same options / settings just two different deliveries.
'                ' Populate the OpenWRT Config File "/etc/config/openvpn"
'                OpenWRT_Config_File = FreeFile
'                Open OpenWRT_Config_File_Location For Append As OpenWRT_Config_File
'
'                ' Populate the OpenWRT Config file with the first few lines
'                OutputDataLine = "config" & Chr(9) & "openvpn" & Chr(9) & "'" & OpenVPN_File_Name & "'" & Chr(10)
'                OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "config" & Chr(9) & "'" & OpenWRT_OpenVPN_Location & Config_File_Name & "'" & Chr(10) & Chr(10)
'
'                Print #OpenWRT_Config_File, OutputDataLine;
'
'                Close #OpenWRT_Config_File
                
                ' ###########################################################
                
                
                Config_File = FreeFile
                Open Config_File_Location For Append As Config_File
                
                ' Populate the OpenWRT Config file with the first few lines
                OutputDataLine = "config" & Chr(9) & "openvpn" & Chr(9) & "'" & OpenVPN_File_Name & "'" & Chr(10) & Chr(10)
                OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "client" & Chr(9) & "'1'" & Chr(10)
                OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "auth_user_pass" & Chr(9) & "'" & OpenWRT_Crypto_Location & OpenVPN_File_Name & "/" & Config_Auth_File_Name & "'" & Chr(10)
                OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "cert" & Chr(9) & "'" & OpenWRT_Crypto_Location & OpenVPN_File_Name & "/" & Config_Cert_File_Name & "'" & Chr(10)
                OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "ca" & Chr(9) & "'" & OpenWRT_Crypto_Location & OpenVPN_File_Name & "/" & Config_CA_File_Name & "'" & Chr(10)
                OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "tls_auth" & Chr(9) & "'" & OpenWRT_Crypto_Location & OpenVPN_File_Name & "/" & Config_TLS_AUTH_File_Name & "'" & Chr(10)
                OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "key" & Chr(9) & "'" & OpenWRT_Crypto_Location & OpenVPN_File_Name & "/" & Config_Key_File_Name & "'" & Chr(10)
                
                Print #Config_File, OutputDataLine;
                
                
                For k = 0 To UBound(InputLines)
                    
                    If InputLines(k) = "" Then GoTo NEXTROUND
                    
                    ' Check for CLIENT-CERT section
                    If InputLines(k) = "<cert>" Then
                    
                        Config_Cert_File = FreeFile
                        Open Config_Cert_File_Location For Output As #Config_Cert_File
                        
                        k = k + 1 ' Line by Line
                        
                        While InputLines(k) <> "</cert>"
                        
                            Print #Config_Cert_File, InputLines(k) & Chr(10);
                            
                            k = k + 1 ' Line by Line
                            
                        Wend
                        
                        Close #Config_Cert_File
                        
                        GoTo NEXTROUND
                    
                    End If ' Check for CLIENT-CERT
                    
                    
                    ' Check for CLIENT-KEY section
                    If InputLines(k) = "<key>" Then
                    
                        Config_Key_File = FreeFile
                        Open Config_Key_File_Location For Output As #Config_Key_File
                        
                        k = k + 1 ' Line by Line
                        
                        While InputLines(k) <> "</key>"
                        
                            Print #Config_Key_File, InputLines(k) & Chr(10);
                            
                            k = k + 1 ' Line by Line
                            
                        Wend
                        
                        Close #Config_Key_File
                    
                        GoTo NEXTROUND
                    
                    End If ' Check for CLIENT-KEY
                    
                    
                    ' Check for TLS-AUTH section
                    If InputLines(k) = "<tls-auth>" Then
                    
                        Config_TLS_AUTH_File = FreeFile
                        Open Config_TLS_AUTH_File_Location For Output As #Config_TLS_AUTH_File
                        
                        k = k + 1 ' Line by Line
                        
                        While InputLines(k) <> "</tls-auth>"
                        
                            Print #Config_TLS_AUTH_File, InputLines(k) & Chr(10);
                            
                            k = k + 1 ' Line by Line
                            
                        Wend
                        
                        Close #Config_TLS_AUTH_File
                        
                        GoTo NEXTROUND
                    
                    End If ' Check for TLS-AUTH
                    
                    
                    ' Check for CERTIFICATE-AUTHORITY section
                    If InputLines(k) = "<ca>" Then
                    
                        Config_CA_File = FreeFile
                        Open Config_CA_File_Location For Output As #Config_CA_File
                        
                        k = k + 1 ' Line by Line
                        
                        While InputLines(k) <> "</ca>"
                        
                            Print #Config_CA_File, InputLines(k) & Chr(10);
                            
                            k = k + 1 ' Line by Line
                            
                        Wend
                        
                        Close #Config_CA_File
                        
                        GoTo NEXTROUND
                    
                    End If ' Check for CERTIFICATE-AUTHORITY
                    
                    
                    Elemente = Split(InputLines(k))
                    
                    
                    ' Do we have the REMOTE server ... and a PORT?
                    If Elemente(0) = "remote" Then
                                
                                OutputDataLine = Chr(9) & "option" & Chr(9) & Elemente(0) & Chr(9) & "'" & Elemente(1) & "'" & Chr(10)
                                
                                If UBound(Elemente) = 2 Then
                                    
                                    OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "port" & Chr(9) & "'" & Elemente(2) & "'" & Chr(10)
                                    
                                End If ' Check if we have a PORT#
                                
                                Print #Config_File, OutputDataLine;
                                
                                GoTo NEXTROUND
                                
                        ' We want to leave the CYPHER untouched ... no conversion of "-" to "_"
                        ElseIf Elemente(0) = "cipher" Then
                                
                                OutputDataLine = Chr(9) & "option" & Chr(9) & Elemente(0) & Chr(9) & "'" & Elemente(1)
                                
                                If UBound(Elemente) = 2 Then
                                    
                                        OutputDataLine = OutputDataLine & " " & Elemente(2) & "'" & Chr(10)
                                    
                                    Else
                                    
                                        OutputDataLine = OutputDataLine & "'" & Chr(10)
                                    
                                End If ' Check if we have a 2 options for this setting
                                
                                Print #Config_File, OutputDataLine;
                                
                                GoTo NEXTROUND
                                
                                
                        ' The --comp-lzo option would only enable the LZO compression algorithm.
                        ' The --compress option allows also to use the improved LZ4 algorithm instead.
                        ' This will allow --compress to be pushed by the server on a per-client basis.
                        ' Providing just --compress without an algorithm is the equivalent of --comp-lzo
                        ' no which disables compression but enables the packet framing for compression.
                        ' Contrary to prior statements --comp-lzo no is not compatible with the --compress
                        ' counterpart. Therefore openvpn needs to keep supporting --comp-lzo no for backward
                        ' compatibility.
                        ElseIf Elemente(0) = "comp-lzo" Then

                                OutputDataLine = Chr(9) & "option" & Chr(9) & "compress"

                                If Elemente(1) = "yes" Then

                                        OutputDataLine = OutputDataLine & Chr(9) & "'" & "lzo" & "'" & Chr(10)

                                    Else

                                        OutputDataLine = OutputDataLine & Chr(10)

                                End If ' Check if we have a 2 options for this setting

                                ' we will FORCE compression, eventhough our VPN vendor does not suggest it.
                                If InStr(OutputDataLine, "lzo") = 0 And InStr(OutputDataLine, "lz4") = 0 Then OutputDataLine = Chr(9) & "option" & Chr(9) & "compress" & Chr(9) & "'" & "lz4" & "'" & Chr(10)
                                
                                Print #Config_File, OutputDataLine;

                                GoTo NEXTROUND
                                
                                
                        ' As of OpenSSL v1.1, the nsCertType extension in X.509 certificates are no longer
                        ' supported. This extension is old and have been deprecated for a long time.
                        ' The replacement option, ---remote-cert-tls is a macro which sets the --remote-cert-ku
                        ' and --remote-cert-eku to appropriate values, depending on it is wanted to check if the
                        ' remote provided certificate is a server or client certificate. As the extended key
                        ' usage extension is far more commonly used today, this is effectively the equivalent
                        ' of --ns-cert-type. For the time being, if --ns-cert-type is used in OpenVPN v2.5 or
                        ' later, it will currently be re-mapped to --remote-cert-tls and complain about a deprecated
                        ' option being used.
                        ElseIf Elemente(0) = "ns-cert-type" Then
                                
                                OutputDataLine = Chr(9) & "option" & Chr(9) & "remote_cert_tls" & Chr(9) & "'" & Elemente(1) & "'" & Chr(10)
                                
                                Print #Config_File, OutputDataLine;
                                
                                GoTo NEXTROUND
                                
                                
                        ElseIf Elemente(0) = "verify-x509-name" Then
                                
                                OutputDataLine = Chr(9) & "option" & Chr(9) & Clean(Elemente(0)) & Chr(9) & "'" & Elemente(1)
                                
                                If UBound(Elemente) = 2 Then
                                    
                                        OutputDataLine = OutputDataLine & " " & Elemente(2) & "'" & Chr(10)
                                    
                                    Else
                                    
                                        OutputDataLine = OutputDataLine & "'" & Chr(10)
                                    
                                End If ' Check if we have a 2 options for this setting
                                
                                Print #Config_File, OutputDataLine;
                                
                                GoTo NEXTROUND
                        
                        
                        ' Do we have a SINGLE option?
                        ElseIf UBound(Elemente) = 0 Then
                    
                                OutputDataLine = Chr(9) & "option" & Chr(9) & Clean(Elemente(0)) & Chr(10)
                    
                                GoTo NEXTROUND
                        
                        ' Do we have a TWO elements of an OPTION?
                        ElseIf UBound(Elemente) = 1 Then
                        
                                OutputDataLine = Chr(9) & "option" & Chr(9) & Clean(Elemente(0)) & Chr(9) & "'" & Clean(Elemente(1)) & "'" & Chr(10)
                                
                                Print #Config_File, OutputDataLine;
                                
                                GoTo NEXTROUND
                        
                            Else
                            
                                OutputDataLine = "option"
                                
                                For j = 1 To UBound(Elemente)
                                
                                    OutputDataLine = OutputDataLine & Chr(9) & "'" & Clean(Elemente(j)) & "'"
                                
                                Next j
                                
                                OutputDataLine = OutputDataLine & Chr(10)
                                
                                Print #Config_File, OutputDataLine;
                        
                    End If ' check # of Elements
                    
                    
NEXTROUND:
                    
                Next k ' walk through lines of input data
                
                'We want to have 2 x empty lines between configurations
                Print #Config_File, Chr(10) & Chr(10);
                
            End If ' Check if correct file TYPE
                
            Close #Config_File
            Close #OpenVPN_File
            
        Next i 'Process one file at a time
        
        Config_File = FreeFile
        Open Config_File_Location For Append As Config_File
                
        ' Put a DUMMY ENDE line at the end; this will help to better
        ' remove and insert the configuration without deleting other data in
        ' the config file.
        OutputDataLine = Chr(10) & "config" & Chr(9) & "openvpn" & Chr(9) & "'DUMMY_Ende'" & Chr(10)
        OutputDataLine = OutputDataLine & Chr(9) & "option" & Chr(9) & "config" & Chr(9) & "'dummyend.conf'" & Chr(10) & Chr(10)
        Print #Config_File, OutputDataLine;
        
        Close #Config_File
            
    End If ' Check if we have a valid DIR
    
' ##############################################################################################
' ##############################################################################################

' We will COPY all generated files onto the router

' ##############################################################################################
' ##############################################################################################
    
    ' Prep the ROUTER
    ' EXECUTE remote OpenWRT command:
    '
    ' Using psftp:
    '
    ' rm -r /tmp/openvpn
    ' mkdir /tmp/openvpn
    
    Call ExecuteShellCommand("rm -r /tmp/openvpn ; mkdir /tmp/openvpn ; rm " & OpenWRT_OpenVPN_Location & Config_File_Name, "cli")
    
    ' COPY all files from Windows to Router
    ' EXECUTE remote OpenWRT command:
    '
    ' Using psftp:
    '
    ' cd /etc/openvpn (from Variable: OpenWRT_OpenVPN_Location)
    ' put -r C:\temp\OpenVPN (from Variable: TARGETDIR)
        
    Call ExecuteShellCommand("cd " & OpenWRT_OpenVPN_Location & Chr(10) & "put -r " & TARGETDIR, "sftp")
    
    ' ##############################################################################################
    ' ##############################################################################################
    '
    ' This can be used to REMOVE the an OLD configuration independently !
    ' you just need PuTTY tools and the sub "ExecuteShellCommand"
    '
    ' ##############################################################################################
    ' ##############################################################################################
    '
    ' Check if we have an OLD VPN config from this tool & remove it
    ' EXECUTE remote OpenWRT command:
    '
    ' Using psftp:
    '
    ' DELETE the previous configurations: (the 100,000 is just an abritary number ... making sure we also catch BIG configuration files)
    '
    ' grep -B 100000 -f /etc/openvpn/PStart.TXT /etc/config/openvpn | head -n -2 >/tmp/openvpn/openvpn.tmp
    ' && grep -B 100000 -f /etc/openvpn/PStart.TXT /etc/config/openvpn | tail -n 1  >>/tmp/openvpn/openvpn.tmp
    ' && mv /tmp/openvpn/openvpn.tmp /etc/config/openvpn
    '
    ' ATTENTION this will OVERWRITE the current /etc/config/openvpn file !!!
    
    Call ExecuteShellCommand("grep -B 100000 DUMMY_Start" & " /etc/config/openvpn | head -n -2 >" & TEMPDIRRT & "openvpn.tmp" _
                       & " && grep -A 100000 DUMMY_Ende" & " /etc/config/openvpn | tail -n 1 >>" & TEMPDIRRT & "openvpn.tmp" _
                       & " && mv " & TEMPDIRRT & "openvpn.tmp" & " /etc/config/openvpn" & Chr(10), "cli")
                       
    ' ##############################################################################################
    ' ##############################################################################################
    
    
    ' APPEND new config to current config
    ' EXECUTE remote OpenWRT command:
    '
    ' Using psftp:
    '
    ' cat /etc/config/openvpn /etc/openvpn/openvpn.conf >/tmp/openvpn/openvpn.tmp
    ' && mv /tmp/openvpn/openvpn.tmp /etc/config/openvpn
        
    Call ExecuteShellCommand("cat /etc/config/openvpn " & OpenWRT_OpenVPN_Location & Config_File_Name & " > " & TEMPDIRRT & "openvpn.tmp" _
                       & " && mv " & TEMPDIRRT & "openvpn.tmp" & " /etc/config/openvpn" & Chr(10), "cli")
    

    ' Final CLEANUP on the ROUTER
    ' EXECUTE remote OpenWRT command:
    '
    ' Using psftp:
    '
    ' rm -r /tmp/openvpn
    ' mkdir /tmp/openvpn
    '
    ' ATTENTION this will OVERWRITE any previously copied files !!!
    Call ExecuteShellCommand("rm -r /tmp/openvpn", "cli")
    

    ' Final CLEANUP on WINDOWS
    ' Clean the TARGETDIR
    DeleteFolderName = Left(TARGETDIR, Len(TARGETDIR) - 1)
    If FSO.FolderExists(DeleteFolderName) Then Result = FSO.DeleteFolder(DeleteFolderName, True)
    
    ' Clean the TEMPDIRWIN
    DeleteFolderName = Left(TEMPDIRWIN, Len(TEMPDIRWIN) - 1)
    If FSO.FolderExists(DeleteFolderName) Then Result = FSO.DeleteFolder(DeleteFolderName, True)
    
'We set FSO to nothing ==> it now has no longer any active references, so VBA safely GCs it.
Set FSO = Nothing

'after your code runs, restore state; put this at the end of your code

Application.ScreenUpdating = screenUpdateState

Application.DisplayStatusBar = statusBarState

'Application.Calculation = xlCalculationAutomatic

Application.EnableEvents = eventsState

'ActiveSheet.DisplayPageBreaks = displayPageBreaksState 'note this is a sheet-level setting

GoTo Ende

Processing_Err:

    Select Case Err.Number
        Case NO_FILES_IN_DIR
            MsgBox "The directory named '" & StartDirectory _
                & "' contains no files."
        Case INVALID_DIR
            MsgBox "'" & StartDirectory & "' is not a valid directory."
        Case 0
        Case Else
            MsgBox "Error #" & Err.Number & " - " & Err.Description
    
    
    End Select
    
    Resume

Ende:

End Sub

Private Function GetAllFilesInDir(ByVal strDirPath As String) As Variant
    ' Loop through the directory specified in strDirPath and save each
    ' file name in an array, then return that array to the calling
    ' procedure.
    ' Return False if strDirPath is not a valid directory.
    Dim strTempName As String
    Dim varFiles() As Variant
    Dim lngFileCount As Long
    
    On Error GoTo GetAllFiles_Err
    
    ' Make sure that strDirPath ends with a "\" character.
    If Right$(strDirPath, 1) <> "\" Then
        strDirPath = strDirPath & "\"
    End If
    
    ' Make sure strDirPath is a directory.
    If ((GetAttr(strDirPath) = vbDirectory) Or (GetAttr(strDirPath) = vbDirectory + vbReadOnly)) Then
        strTempName = Dir(strDirPath, vbDirectory)
        Do Until Len(strTempName) = 0
            ' Exclude ".", "..".
            If (strTempName <> ".") And (strTempName <> "..") Then
                ' Make sure we do not have a sub-directory name.
                If (GetAttr(strDirPath & strTempName) _
                    And vbDirectory) <> vbDirectory Then
                    ' Increase the size of the array
                    ' to accommodate the found filename
                    ' and add the filename to the array.
                    ReDim Preserve varFiles(lngFileCount)
                    varFiles(lngFileCount) = strTempName
                    lngFileCount = lngFileCount + 1
                End If
            End If
            ' Use the Dir function to find the next filename.
            strTempName = Dir()
        Loop
        ' Return the array of found files.
        GetAllFilesInDir = varFiles
    End If
GetAllFiles_End:
    Exit Function
GetAllFiles_Err:
    GetAllFilesInDir = False
    Resume GetAllFiles_End
End Function

Private Function Clean(InputString As String) As String
    ' In OpenWRT some options come along with "_" ... in some VPN config files, we see "-" instead
    ' Thus we CONVERT all "-" into "_" ... making OpenWRT happy.
    Clean = Replace(InputString, "-", "_")
End Function

Private Function Quote(str As String) As String
  If Left(str, 1) = """" Then
    Quote = str
    Else: Quote = """" & str & """"
  End If
End Function


Private Sub ExecuteShellCommand(commando As String, tool As String)

    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim ReturnValue As Integer
    
    ' We create a Scripting.FileSystemObject and FSO's reference count of this new object is now 1
    Set FSO = CreateObject("Scripting.FileSystemObject")


    ' Initialize the STARTUPINFO structure:
    start.cb = Len(start)

    ' EXECUTE remote OpenWRT command:
    '
    ' FIRST write the command to a TEXT file on the Windows Machine

    Komando = TEMPDIRWIN & "C.TXT"

    Select Case tool

        Case "sftp"
            software = PUTTYDIR & "psftp.exe"
            closer = "quit"
            switches = " -C -batch -P " & RouterSSHPort & " " & PuTTYSession & " -b "
    
        Case "cli"
            software = PUTTYDIR & "plink.exe"
            closer = "exit"
            switches = " -batch -load " & PuTTYSession & " -m "
    
    End Select

    Komando_File = FreeFile
    Open Komando For Output As #Komando_File

    Print #Komando_File, commando
    Print #Komando_File, closer

    Close #Komando_File

    ' NOW configure the PuTTY command to execute the newly created TXT file on the router
    '
    myFile = software
    Komando = Komando
    Options = switches & Komando
    cmdline = myFile & Options

    ' Execute PuTTY command
    ' e.g. "E:\Web\Hardware\ROUTER\Tools\Putty\psftp.exe -C -batch -P 2212 psftp -b C:\Temp\OpenVPN\C.TXT"
    '
    ' Start the shelled application:
    ReturnValue = CreateProcessA(0&, cmdline, 0&, 0&, 1&, _
    NORMAL_PRIORITY_CLASS, 0&, 0&, start, proc)

    ' Wait for the shelled application to finish:
    Do
        ReturnValue = WaitForSingleObject(proc.hProcess, 0)
        DoEvents
    Loop Until ReturnValue <> 258

    ReturnValue = CloseHandle(proc.hProcess)

    ' cleaning up ... deleteing the previously created Komando file in TEMPDIRWIN
    If FSO.FolderExists(Komando) Then FSO.DeleteFile (Komando)
    
    'We set FSO to nothing ==> it now has no longer any active references, so VBA safely GCs it.
    Set FSO = Nothing
    
End Sub
