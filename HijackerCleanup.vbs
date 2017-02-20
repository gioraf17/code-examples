'********************************************************************************
'* Script name: HijackerCleanup.vbs
'* Created on:  3/26/2016
'* Author:      Giorgio Rafaelle
'* Purpose:     Recursively searches the current user subdirectories and restores
'*              default shortcut targets for various web browsers, and
'*              optionally restores select IE registry subkeys to their
'*              default values.
'* History:     2/19/2017 updated comments, style asjustments
'********************************************************************************

'* run with elevated privilages
If Not WScript.Arguments.Named.Exists("elevate") Then
  CreateObject("Shell.Application").ShellExecute WScript.FullName _
    , WScript.ScriptFullName & " /elevate", "", "runas", 1
  WScript.Quit
End If

'* initialize script shell and file system object
set wshell = Createobject("WScript.Shell")
set fso= createobject("Scripting.FileSystemObject")

'* get current user name
strUser = wshell.ExpandEnvironmentStrings("%username%")

'* default empty shortcut target parameter used to reset browser shortcuts
Const ShortcutArg = ""

'* shortcut names for the supported web browsers
Const IEShortcut = "internet explorer.lnk"
Const FFShortcut = "mozilla firefox.lnk"
Const GCShortcut = "google chrome.lnk"
Const EdgeShortcut = "microsoft edge.lnk"

'* searchbar registry subkey for IE - to be deleted as it may be related to malicious code
Const SearchBarSubkey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Search Bar"

'* registry subkeys for IE that are to be reset to their default values
Const SearchURLDefaultSubkey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchURL\(Default)"
Const SearchURLProviderSubkey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\SearchURL\Provider"
Const UserSearchPageSubkey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\Search Page"
Const MachineSearchPageSubkey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Search Page"
Const StartPageSubkey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Start Page"
Const SearchHooksSubkey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\URLSearchHooks\(Default)"
Const SearchHooksOtherSubkey = "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\URLSearchHooks\{CFBFAE00-17A6-11D0-99CB-00C04FD64497}"
Const DefaultPageSubkey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Default_Page_URL"
Const DefaultSearchSubkey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Main\Default_Search_URL"
Const CustomizSearchSubkey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Search\CustomizeSearch"
Const SearchAssistantSubkey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\Search\SearchAssistant"
Const AboutURLsSubKey = "HKEY_LOCAL_MACHINE\Software\Microsoft\Internet Explorer\AboutURLs\blank"

'* default subkey values for IE
Const SearchPageVal = "http://go.microsoft.com/fwlink/?LinkId=54896"
Const DefaultStartPageVal = "http://go.microsoft.com/fwlink/p/?LinkId=255141"
Const AboutURLsVal = "res://mshtml.dll/blank.htm"
Const CustSearchURL = "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchcust.htm"
Const AssistantURL = "http://ie.search.msn.com/{SUB_RFC1766}/srchasst/srchasst.htm"

'*********************************************************
'* Function Name: Is32BitOS
'* Purpose:       Determines if CPU architecture is 32 bit
'* Return Value:  Boolean value
'*********************************************************
Function Is32BitOS()
    Const Path = "winmgmts:root\cimv2:Win32_Processor='cpu0'"
    Is32BitOS = (GetObject(Path).AddressWidth = 32)
End Function

'*********************************************************
'* Function Name: Is64BitOS
'* Purpose:       Determines if CPU architecture is 64 bit
'* Return Value:  Boolean value
'*********************************************************
Function Is64BitOS()
    Const Path = "winmgmts:root\cimv2:Win32_Processor='cpu0'"
    Is64BitOS = (GetObject(Path).AddressWidth = 64)
End Function

'*********************************************************
'* Function Name:      KeyExists
'* Purpose:            Check if given registry subkey exists
'* Arguments Supplied: Key, the registry key to check for
'* Return Value:       Boolean value
'*********************************************************
Function KeyExists(Key)
  Dim entry

  On Error Resume Next
  entry = wshell.RegRead(Key)
  If Err.Number = 0 then
    Err.Clear
    KeyExists = False
  Else
    Err.Clear
    KeyExists = True
  End If
End Function

'* default paths for IE - 32 or 64 bit
Const IE32bit = "C:\Program Files (x86)\Internet Explorer\iexplore.exe"
Const IE64bit = "C:\Program Files\Internet Explorer\iexplore.exe"

'* set path for IE based upon CPU architecture
IEPath = ""
If Is32BitOS() then
  IEPath = IE32bit
ElseIf Is64BitOS() then
  IEPath = IE64bit
Else
  FFPath = ""
End If

'* default path for Microsoft Edge
Const EdgePath = "C:\Windows\SystemApps\Microsoft.MicrosoftEdge_8wekyb3d8bbwe\MicrosoftEdge.exe"

'* default paths for Firefox
Const FF32bit = "C:\Program Files\Mozilla Firefox\firefox.exe"
Const FF64bit = "C:\Program Files (x86)\Mozilla Firefox\firefox.exe"

'* set the correct path for Firefox installation
FFPath = ""
If fso.FileExists(FF32bit) then
  FFPath = FF32bit
ElseIf fso.FileExists(FF64bit) then
  FFPath = FF64bit
Else
  FFPath = ""
End If

'* default paths for Google Chrome installation
ChromeWin7 = "\Users\"&strUser&"\AppData\Local\Google\Chrome\Application\chrome.exe"
Const ChromeWin8 = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"

'* set correct path for Google Chrome instllation
ChromePath = ""
If fso.FileExists(ChromeWin7) then
  ChromePath = ChromeWin7
ElseIf fso.FileExists(ChromeWin8) then
  ChromePath = ChromeWin8
Else
  ChromePath = ""
End If

'* search every subdirectory of the current user for all browser shortcuts and modify
RecursiveTargetMod fso.GetFolder("C:\Users\"&strUser&"\")
RecursiveTargetMod fso.GetFolder("C:\Users\Public\")

'***********************************************************************
'* Subroutine Name:      RecursiveTargetMod
'* Purpose:              Recursively search the given directory, and all
'*                       subdirectories and modify all found web browser
'*                       shortcut targets.
'* Arguments Supplied:   objFolder, the directory to be searched
'***********************************************************************
Sub RecursiveTargetMod(objFolder)
  '* declare file and subfolder variables
  Dim objFile, objSubFolder
  '* check each file in the current folder against the specified
  '* web browser shortcut filename, and modify the target
  For Each objFile In objFolder.Files
    '* check the current file name against the IE shortcut filename
    If LCase(objFile.Name) = IEShortcut then
      '* reset target of the shorcut
      Set shortcut = wshell.CreateShortcut(objFile.Path)
      shortcut.TargetPath = IEPath
      shortcut.Arguments = ShortcutArg
      shortcut.Save
    '* check the current file name against the Firefox shortcut filename
    ElseIf LCase(objFile.Name) = FFShortcut then
      '* reset target of the shorcut
      Set shortcut = wshell.CreateShortcut(objFile.Path)
      shortcut.TargetPath = FFPath
      shortcut.Arguments = ShortcutArg
      shortcut.Save
    '* check the current file name against the Google Chrome shortcut filename
    ElseIf LCase(objFile.Name) = GCShortcut then
      '* reset target of the shorcut
      Set shortcut = wshell.CreateShortcut(objFile.Path)
      shortcut.TargetPath = ChromePath
      shortcut.Arguments = ShortcutArg
      shortcut.Save
    '* check the current file name against the Microsoft Edge shortcut filename
    ElseIf LCase(objFile.Name) = EdgeShortcut then
      '* reset target of the shorcut
      Set shortcut = wshell.CreateShortcut(objFile.Path)
      shortcut.TargetPath = EdgePath
      shortcut.Arguments = ShortcutArg
      shortcut.Save
    End If
  Next

  '* search each subfolder
  For Each objSubFolder In objFolder.SubFolders
    On Error Resume Next
    RecursiveTargetMod objSubFolder
    Err.Clear
  Next
End Sub

'* promt user upon completion of shortcut reset, ask to reset IE subkeys
prompt = MsgBox("Shortcuts reset! Do you wish to reset Internet Explorer registry defaults? (This should only be done if IE is still redirecting after shortcut reset)", 1, "Hijacker Cleanup Tool")

'* exit script if user answers no, reset IE registry subkeys if user answers yes
If prompt = 2 then
  WScript.Quit
ElseIf prompt = 1 then
  '* reset search bar subkey if it exists
  If KeyExists(SearchBarSubkey) = 0 then
    wshell.RegDelete SearchBarSubkey
  End If
  '* reset SearchURLDefault subkey if it exists
  If KeyExists(SearchURLDefaultSubkey) = 0 then
    wshell.RegWrite SearchURLDefaultSubkey, "(value not set)", "REG_SZ"
  End If
  '* reset SearchURLProvider subkey if it exists
  If KeyExists(SearchURLProviderSubkey) = 0 then
    wshell.RegWrite SearchURLProviderSubkey, "no value", "REG_SZ"
  End If
  '* reset UserSearchPage subkey if it exists
  If KeyExists(UserSearchPageSubkey) = 0 then
    wshell.RegWrite UserSearchPageSubkey, SearchPageVal, "REG_SZ"
  End If
  '* reset MachineSearchPage subkey if it exists
  If KeyExists(MachineSearchPageSubkey) = 0 then
    wshell.RegWrite MachineSearchPageSubkey, SearchPageVal, "REG_SZ"
  End If
  '* reset StartPage subkey if it exists
  If KeyExists(StartPageSubkey) = 0 then
    wshell.RegWrite StartPageSubkey, DefaultStartPageVal, "REG_SZ"
  End If
  '* reset SearchHooks subkey if it exists
  If KeyExists(SearchHooksSubkey) = 0 then
    wshell.RegWrite SearchHooksSubkey, "(value not set)", "REG_SZ"
  End If
  '* reset SearchHooksOther subkey if it exists
  If KeyExists(SearchHooksOtherSubkey) = 0 then
    wshell.RegWrite SearchHooksOtherSubkey, "", "REG_SZ"
  End If
  '* reset DefaultPage subkey if it exists
  If KeyExists(DefaultPageSubkey) = 0 then
    wshell.RegWrite DefaultPageSubkey, DefaultStartPageVal, "REG_SZ"
  End If
  '* reset DefaultSearch subkey if it exists
  If KeyExists(DefaultSearchSubkey) = 0 then
    wshell.RegWrite DefaultSearchSubkey, SearchPageVal, "REG_SZ"
  End If
  '* reset CustomizeSearch subkey if it exists
  If KeyExists(CustomizSearchSubkey) = 0 then
    wshell.RegWrite CustomizSearchSubkey, CustSearchURL, "REG_SZ"
  End If
  '* reset SearchAssistant subkey if it exists
  If KeyExists(SearchAssistantSubkey) = 0 then
    wshell.RegWrite SearchAssistantSubkey, AssistantURL, "REG_SZ"
  End If
  '* reset AboutURLs subkey if it exists
  If KeyExists(AboutURLsSubKey) = 0 then
    wshell.RegWrite AboutURLsSubKey, AboutURLsVal, "REG_SZ"
  End If
  '* prompt user when subkeys are reset
  MsgBox("You need to restart and reset Internet Options once again for changes to take effect.")
End If
