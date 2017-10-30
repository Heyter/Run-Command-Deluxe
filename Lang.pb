
Enumeration
	#Lng_Window_Title
	#Lng_Text_Open
	#Lng_Text_Default
	#Lng_Button_Run
	#Lng_Button_RunAdmin
	#Lng_Button_Browse
	
	; Counter
	#Lng_entries
EndEnumeration

Global Dim LNG.s(#Lng_entries-1)

Procedure.b LoadLanguage(PrefFile$)
	OpenPreferences(PrefFile$)
	PreferenceGroup("Window")
	LNG(#Lng_Window_Title) = ReadPreferenceString("Title", "Default Title")
	ClosePreferences()
	ProcedureReturn #True
EndProcedure

; IDE Options = PureBasic 5.60 (Windows - x86)
; CursorPosition = 5
; Folding = -
; EnableXP
; CompileSourceDirectory