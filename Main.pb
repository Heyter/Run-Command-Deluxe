; TODO: Add start hidden option
; TODO: Add large icon mode
; TODO: Add appdata path var

;
;- Setup - Compiler Stuff
;

CompilerIf Not #PB_Compiler_OS = #PB_OS_Windows
	CompilerError "This program can only be compiled for Windows"
CompilerElseIf Not #PB_Compiler_OS = #PB_OS_Windows_10
	;CompilerWarning "This program might encounter errors on things, fuck this..."
CompilerEndIf


;
;- Setup - Imports
;

XIncludeFile ".\PB-Utils\Strings.pb"
XIncludeFile ".\PB-Utils\Logger.pb"
;XIncludeFile ".\PB-Gadgets\ButtonIconGadget.pb"
XIncludeFile ".\IconGrabber.pb"
XIncludeFile ".\Lang.pb"

ConfigureDebugWindowLogLevel(#LoggingLevel_Any | #LoggingLevel_Trace)
ConfigureDebugWindowLogLevel(#LoggingLevel_Any)
;ConfigureConsoleLogLevel(#LoggingLevel_Any)


;
;- Setup - Enumerations & Constants
;

Global OSArchitecture.s = GetEnvironmentVariable("PROCESSOR_ARCHITECTURE")

Enumeration ErrorCode
	#ERR_ParsingFailure
	#ERR_CFG_FileIsDir
	#ERR_DataDirIsFile
	#ERR_DataDirCreationFailure
EndEnumeration

Enumeration
	#CGF_DataDirectory
	
	#CFG_entries
EndEnumeration
Global Dim CFG.s(#CFG_entries-1)


;
;- Setup - Preferences
;

PreferenceFile.s = GetFilePart(ProgramFilename(), #PB_FileSystem_NoExtension)+".ini"
PreferenceFile = ReplaceString(ReplaceString(PreferenceFile, "_x86", ""), "_x64", "")

DefaultDataDirPath.s = ".\"+GetFilePart(ProgramFilename(), #PB_FileSystem_NoExtension)+"_Data"
DefaultDataDirPath = ReplaceString(ReplaceString(DefaultDataDirPath, "_x86", ""), "_x64", "")

If FileSize(PreferenceFile) = -2
	LogFatal("The config file is a directory", #ERR_CFG_FileIsDir, #True, "Fatal Error", "The config file is a directory")
EndIf

If FileSize(PreferenceFile) = -1
	CreatePreferences(PreferenceFile, #PB_Preference_GroupSeparator)
	WritePreferenceString("DataDirectory", DefaultDataDirPath)
Else
	OpenPreferences(PreferenceFile, #PB_Preference_GroupSeparator)
EndIf

LogTrace("Loading the data directory path and preparing it...")
CFG(#CGF_DataDirectory) = ReadPreferenceString("DataDirectory", DefaultDataDirPath)
If FileSize(CFG(#CGF_DataDirectory)) >= 0
	LogFatal("The data directory is a file ("+CFG(#CGF_DataDirectory)+")", #ERR_DataDirIsFile, #True, "Fatal Error", "The data directory is a file ("+CFG(#CGF_DataDirectory)+")")
ElseIf FileSize(CFG(#CGF_DataDirectory)) = -1
	If Not CreateDirectory(CFG(#CGF_DataDirectory))
		LogFatal("Unable to create the data directory ("+CFG(#CGF_DataDirectory)+")", #ERR_DataDirCreationFailure, #True, "Fatal Error", "Unable to create the data directory ("+CFG(#CGF_DataDirectory)+")")
	EndIf
EndIf

ClosePreferences()


;
;- Setup - Path variables
;

LogDebug("Reading path variables")
Global NewMap PathPlaceholders.s()
Dim _PathPlaceholders.s(0)
Restore TextDefaultPathPlaceholders
Read.s _PathPlaceholdersDefinition$
_PathPlaceholdersDefinition$ = _PathPlaceholdersDefinition$ + #CRLF$ + "%data%;"+CFG(#CGF_DataDirectory)

LogTrace("Attempting to read custom path variables...")
If FileSize(CFG(#CGF_DataDirectory)+"\pathvars.cfg")
	If ReadFile(0, CFG(#CGF_DataDirectory)+"\pathvars.cfg")
		While Eof(0) = 0
			_PathPlaceholdersDefinition$ = _PathPlaceholdersDefinition$ + #CRLF$+ ReadString(0, #PB_UTF8)
		Wend
		CloseFile(0)
	Else
		LogError("Unable to open pathvars.cfg", #True, "Error", "Unable to open pathvars.cfg")
	EndIf
EndIf

LogTrace("Preparing and cleanning raw path variables data...")
_PathPlaceholdersDefinition$ = ReplaceString(_PathPlaceholdersDefinition$, #CR$, "")
ExplodeStringToArray(_PathPlaceholders(), _PathPlaceholdersDefinition$, #LF$)

LogTrace("Parsing path variables...")
For i=0 To ArraySize(_PathPlaceholders())-1
	If IsNullOrEmpty(_PathPlaceholders(i))
		Continue
	EndIf
	
	Dim _PlaceholderInfos.s(0)
	If ExplodeStringToArray(_PlaceholderInfos(), _PathPlaceholders(i), ";") <> 2
		LogFatal("Unable to parse path placeholders correctly ("+_PathPlaceholders(i)+")", #ERR_ParsingFailure, #True, "Fatal Error", "Unable to parse path placeholders correctly ("+_PathPlaceholders(i)+")"+#CRLF$+"If you are using a custom definition file, please check it.")
	EndIf
	LogTrace("PathVar: "+_PlaceholderInfos(0)+" -> "+_PlaceholderInfos(1))
	
	PathPlaceholders(_PlaceholderInfos(0)) = _PlaceholderInfos(1)
	FreeArray(_PlaceholderInfos())
Next
FreeArray(_PathPlaceholders())


;
;- Images
;

LogDebug("Loading images and icons...")
UsePNGImageDecoder()

Global ImageMainIcon, ImageRun, ImageFolder, ImageUsers, ImageUsers, ImageAdminShell, ImageSettings, ImageInfo
ImageMainIcon = CatchImage(#PB_Any, ?ImageIcon, ?ImageIconEND - ?ImageIcon)
ImageRun = Icon_GetHdl("c:\windows\system32\shell32.dll", 16, 25)
ImageFolder = Icon_GetRealHdl("C:\Windows\System32\imageres.dll", 16, 4)
ImageAdminShell = Icon_GetRealHdl("C:\Windows\System32\imageres.dll", 16, 74)
ImageUsers = Icon_GetRealHdl("C:\Windows\System32\imageres.dll", 16, 74)
ImageUser = Icon_GetRealHdl("C:\Windows\System32\imageres.dll", 16, 118)
ImageSettings = Icon_GetRealHdl("C:\Windows\System32\imageres.dll", 16, 110)
ImageInfo = Icon_GetRealHdl("C:\Windows\System32\imageres.dll", 16, 77)


;
;- Procedures
;

Procedure RunProgramAsAdmin(ProgramName$, Parameters$ = "", WorkingDirectory$ = "")
	Protected shExecInfo.SHELLEXECUTEINFO
	
	With shExecInfo
		\cbSize = SizeOf(SHELLEXECUTEINFO)
		\lpVerb = @"runas"
		\lpParameters = @Parameters$
		\lpFile = @ProgramName$
		\lpDirectory = @WorkingDirectory$
		\nShow = #SW_NORMAL
	EndWith
	
	ProcedureReturn ShellExecuteEx_(shExecInfo)
EndProcedure


;
;- Window
;

Enumeration
	#EXE_RunAsAdmin = %00000001
	#EXE_DisableFileCheck = %00000010
EndEnumeration

Enumeration
	#EXE_Flags = 0
	#EXE_Path = 1
	#EXE_Name = 2
	#EXE_Description = 3
	#EXE_WorkingDir = 4
	#EXE_Parameters = 5
	#EXE_IconPath = 6
	#EXE_IconNbr = 7
EndEnumeration

Global NewList QuickLaunchAppIDs.i()
Global NewList QuickLaunchAppStartInfo.s()
#IMAGEBUTTONSIZE = 24
#IMAGEBUTTONSTARTY = 120
#IMAGEBUTTONMAXX = 20

#Window_Size_Main_X = 10*2+#IMAGEBUTTONSIZE*#IMAGEBUTTONMAXX

Procedure ClickQuickLaunchButton(event, EventGadgetID)
	SelectElement(QuickLaunchAppStartInfo(), GetGadgetData(EventGadgetID)-1)
	Debug "Launching -> "+QuickLaunchAppStartInfo()
	
	Protected _FullPath.s = Right(QuickLaunchAppStartInfo(), Len(QuickLaunchAppStartInfo())-2)
	
	If Val(Left(QuickLaunchAppStartInfo(), 1)) & #EXE_RunAsAdmin
		RunProgramAsAdmin(_FullPath, "", "")
	Else
		RunProgram(_FullPath)
	EndIf
EndProcedure

Procedure LoadExecutablesDefinitions(WindowID)
	LogDebug("Reading executables definitions files...")
	_RawExecutablesDefinitions.s = ""
	HasReadFiles.b = #False
	
	LogTrace("Trying to read executables.cfg")
	If FileSize(CFG(#CGF_DataDirectory) + "\" + "executables.cfg") > 0
		If ReadFile(0, CFG(#CGF_DataDirectory) + "\" + "executables.cfg")
			LogTrace("Reading executables.cfg...")
			While Eof(0) = 0
				_RawExecutablesDefinitions = _RawExecutablesDefinitions + ReadString(0, #PB_UTF8) + #CRLF$
			Wend
			CloseFile(0)
			HasReadFiles = #True
		Else
			LogError("Unable to open executables.cfg", #True, "Error", "Unable To open executables.cfg")
		EndIf
	EndIf
	
	LogTrace("Walking "+CFG(#CGF_DataDirectory)+" for execs_*.cfg files...")
	If ExamineDirectory(0, CFG(#CGF_DataDirectory), "execs_*.cfg")
		While NextDirectoryEntry(0)
			If DirectoryEntryType(0) = #PB_DirectoryEntry_File And DirectoryEntrySize(0)
				If ReadFile(0, CFG(#CGF_DataDirectory) + "\" + DirectoryEntryName(0))
					LogTrace("Reading "+DirectoryEntryName(0)+"...")
					If HasReadFiles
						_RawExecutablesDefinitions = _RawExecutablesDefinitions + "NEWLINE" + #CRLF$
					EndIf
					
					While Eof(0) = 0
						_RawExecutablesDefinitions = _RawExecutablesDefinitions + ReadString(0, #PB_UTF8) + #CRLF$
					Wend
					CloseFile(0)
					HasReadFiles = #True
				Else
					LogError("Unable to open "+DirectoryEntryName(0), #True, "Error", "Unable to open "+DirectoryEntryName(0))
				EndIf
			EndIf
		Wend
		FinishDirectory(0)
	EndIf
	
	If Not HasReadFiles
		LogDebug("No executables definitions files found, loading default values...")
		Restore TextDefaultExecutablesInfos
		Read.s _RawExecutablesDefinitions
	EndIf
	
	LogDebug("Parsing executables definitions...")
	Dim _ExecutablesDefinitions.s(0)
	_RawExecutablesDefinitions = ReplaceString(_RawExecutablesDefinitions, #CR$, "")
	ExplodeStringToArray(_ExecutablesDefinitions(), _RawExecutablesDefinitions, #LF$)
	_PosY.b = 0
	_PosX.b = 0
	_GadgetID.i = 0
	
	LogTrace("Looping trough exec. defs. ...")
	For i=0 To ArraySize(_ExecutablesDefinitions())-1
		If IsNullOrEmpty(_ExecutablesDefinitions(i)) Or _ExecutablesDefinitions(i) = #CR$ Or _ExecutablesDefinitions(i) = #LF$
			Continue
		EndIf
		
		If Left(_ExecutablesDefinitions(i), 1) = "_"
			Continue
		EndIf
		
		Dim _ExecutableInfos.s(0)
		ExplodeStringToArray(_ExecutableInfos(), _ExecutablesDefinitions(i), ";")
		
		If _ExecutablesDefinitions(i) = "NEWLINE"
			_PosY = _PosY + 1
			_PosX = 0
			Continue
		EndIf
		
		If _ExecutablesDefinitions(i) = "SPACE"
			_PosX = _PosX + 1
			Continue
		EndIf
		
		If _ExecutablesDefinitions(i) = "END"
			Break
		EndIf
		
		If _PosX = #IMAGEBUTTONMAXX
			_PosX = 0
			_PosY = _Posy +1
		EndIf
		
		If Not (ArraySize(_ExecutableInfos()) = 7 Or ArraySize(_ExecutableInfos()) = 8)
			LogFatal("Unable to parse: "+_ExecutablesDefinitions(i), #ERR_ParsingFailure, #True, "Fatal Error", "Unable to parse: "+Chr(34)+_ExecutablesDefinitions(i)+Chr(34)+#CRLF$+"Make sure the files are encoded in AINSI or unsigned UTF-8, or maybe ASCII."+#CRLF$+"And make sure you follow the right [notation]")
		EndIf
		
		_ImgNbr = 1
		If ArraySize(_ExecutableInfos()) = 8
			_ImgNbr = Val(_ExecutableInfos(#EXE_IconNbr))
		EndIf
		
		ForEach PathPlaceholders()
			_ExecutableInfos(#EXE_Path) = ReplaceString(_ExecutableInfos(#EXE_Path), MapKey(PathPlaceholders()), PathPlaceholders())
			_ExecutableInfos(#EXE_IconPath) = ReplaceString(_ExecutableInfos(#EXE_IconPath), MapKey(PathPlaceholders()), PathPlaceholders())
		Next
		
		; Checking if 64bit exe is given, and more...
		IsCompatible.b = #True
		If FindString(_ExecutableInfos(#EXE_Path), "??")
			; If 32-bit only (disabled for 64-bits)
			; TODO: Verify if it really works
			If OSArchitecture = "x86"
				_ExecutableInfos(1) = Left(_ExecutableInfos(#EXE_Path), FindString(_ExecutableInfos(#EXE_Path), "??"))
			Else
				IsCompatible = #False
			EndIf
		ElseIf FindString(_ExecutableInfos(#EXE_Path), "?")
			; If a 64-bit executable is available
			If FindString(_ExecutableInfos(#EXE_Path), "?") = 1
				; If 64bit only
				If OSArchitecture = "AMD64"
					_ExecutableInfos(#EXE_Path) = Right(_ExecutableInfos(#EXE_Path), Len(_ExecutableInfos(#EXE_Path))-1)
				Else
					IsCompatible = #False
				EndIf
			Else
				If OSArchitecture = "AMD64"
					_ExecutableInfos(#EXE_Path) = Right(_ExecutableInfos(#EXE_Path), Len(_ExecutableInfos(#EXE_Path))-FindString(_ExecutableInfos(#EXE_Path), "?"))
				Else
					_ExecutableInfos(#EXE_Path) = Right(_ExecutableInfos(#EXE_Path), FindString(_ExecutableInfos(#EXE_Path), "?")-1)
				EndIf
			EndIf
		EndIf
		
		If IsCompatible And Not Val(_ExecutableInfos(#EXE_Flags)) & #EXE_DisableFileCheck
			If FileSize(_ExecutableInfos(#EXE_Path)) < 0
				LogWarn("Unable to find "+_ExecutableInfos(#EXE_Path))
				Continue
			EndIf
		EndIf
		
		_GadgetID = ButtonImageGadget(#PB_Any, 10+#IMAGEBUTTONSIZE*_PosX, #IMAGEBUTTONSTARTY+#IMAGEBUTTONSIZE*_PosY, #IMAGEBUTTONSIZE, #IMAGEBUTTONSIZE, Icon_GetHdl(_ExecutableInfos(#EXE_IconPath), 16, _ImgNbr))
		If _GadgetID
			AddElement(QuickLaunchAppIDs())
			QuickLaunchAppIDs() = _GadgetID
			
			AddElement(QuickLaunchAppStartInfo())
			QuickLaunchAppStartInfo() = _ExecutableInfos(#EXE_Flags)+";"+_ExecutableInfos(#EXE_Path)
			SetGadgetData(_GadgetID, ListSize(QuickLaunchAppStartInfo()))
			
			;Debug Str(Len(_ExecutableInfos(#EXE_Description))) +" -> "+_ExecutableInfos(#EXE_Description)
			If Len(_ExecutableInfos(#EXE_Description))
				;http://www.purebasic.fr/english/viewtopic.php?t=14482
				GadgetToolTip(_GadgetID, _ExecutableInfos(#EXE_Name)+" ("+GetFilePart(_ExecutableInfos(#EXE_Path))+")"+#CRLF$+_ExecutableInfos(#EXE_Description))
			Else
				GadgetToolTip(_GadgetID, _ExecutableInfos(#EXE_Name)+" ("+GetFilePart(_ExecutableInfos(#EXE_Path))+")")
			EndIf
			
			;SetGadgetColor(_GadgetID, #PB_Gadget_BackColor, RGB(240, 240, 240))
			;SetGadgetColor(_GadgetID, #PB_Gadget_FrontColor, RGB(240, 240, 240))
			;SetGadgetColor(_GadgetID, #PB_Gadget_LineColor, RGB(240, 240, 240))
			
			If Not IsCompatible
				DisableGadget(_GadgetID, 1)
			EndIf
		Else
			LogError("Unable to add ButtonImageGadget for "+_ExecutableInfos(#EXE_Name)+" ("+GetFilePart(_ExecutableInfos(#EXE_Path))+")")
		EndIf
		
		_PosX = _PosX + 1
		FreeArray(_ExecutableInfos())
	Next
	FreeArray(_ExecutablesDefinitions())
	
	ResizeWindow(WindowID, #PB_Ignore, #PB_Ignore, #PB_Ignore, #IMAGEBUTTONSTARTY+#IMAGEBUTTONSIZE*(_PosY+2)+8)
EndProcedure

XIncludeFile "WindowMain.pbf"
OpenWindow_Main(0, 0, #Window_Size_Main_X)

Repeat
	Event = WaitWindowEvent()
	
	If GetAsyncKeyState_(#VK_F1) & $8000
		;Open About window
	EndIf
	
	Select EventWindow()
		Case #Window_Main
			Select Event
				Case #PB_Event_CloseWindow
					Break
				Default
					Window_Main_Events(Event)
			EndSelect
		Default
			Debug "What ?"
	EndSelect
ForEver

End 0


;
;- DataSection
;

DataSection
	ImageIcon:
	IncludeBinary "./icon_32.png"
	ImageIconEND:
	
	TextDefaultPathPlaceholders:
	Data.s "%windir%;C:\Windows"+#LF$+
	       "%sys32%;C:\Windows\System32"+#LF$+
	       "%syswow%;C:\Windows\SysWOW64"+#LF$+
	       "%prg%;C:\Program Files"+#LF$+
	       "%prgx86%;C:\Program Files (x86)"
	
	TextDefaultExecutablesInfos:
	Data.s "2;taskmgr.exe;Task Manager;;;;%sys32%\taskmgr.exe;1"+#LF$+
	       "2;cmd.exe;Command Prompt;;;;%sys32%\cmd.exe;1"+#LF$+
	       "2;powershell.exe;Powershell;;;;%sys32%\WindowsPowerShell\v1.0\powershell.exe"+#LF$+
	       "2;control.exe;Control Panel;;;;%sys32%\control.exe"+#LF$+
	       "3;regedit.exe;Registry Editor;;;;%windir%\regedit.exe"+#LF$+
	       "2;compmgmt.msc;Computer Management;;;;.\data\compmgmt.ico"+#LF$+
	       "2;diskmgmt.msc;Disk Management;;;;%sys32%\dmdskres.dll"+#LF$+
	       "2;devmgmt.msc;Device Manager;;;;%sys32%\devmgr.dll;6"
EndDataSection

; IDE Options = PureBasic 5.60 (Windows - x86)
; CursorPosition = 428
; FirstLine = 383
; Folding = -
; EnableXP
; CompileSourceDirectory