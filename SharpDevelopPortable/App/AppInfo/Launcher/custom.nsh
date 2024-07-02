
;= LAUNCHER 
;= ################
;!define XML_PLUGIN
;!define DIRECTORIES_MOVE
 
;= VARIABLES 
;= ################
 
;= DEFINES
;= ################
!define APP			`SharpDevelop`
!define FULLNAME	`SharpDevelop`
!define APPDIR		`$EXEDIR\App\${APP}`
!define V			5.1.0.5216
!define C			`IC#Code`
!define XML			`${DATA}\ICSharpCode\SharpDevelop5\SharpDevelopProperties.xml`
!define DEFXML		`${DEFDATA}\ICSharpCode\SharpDevelop5\SharpDevelopProperties.xml`
!define XPATH		`//[@key='CoreProperties.UILanguage']`
!define PFM			`$0\PortableApps.com\PortableAppsPlatform.exe`
;!define DOTNET		`SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full`
!define ICSC		`${DATA}\ICSharpCode`
!define DEFICSC		`${DEFDATA}\ICSharpCode`

;= FUNCTIONS
;= ################
Function GetLCID
	!macro _GetLCID _LNG _LCID
		Push ${_LNG}
		Call GetLCID
		Pop ${_LCID}
	!macroend
	!define GetLCID "!insertmacro _GetLCID"
	Exch $0
	StrCmp $0 en 0 +3
	StrCpy $0 1033
		Goto EndLCID
	StrCmp $0 zh 0 +3
	StrCpy $0 2052
		Goto EndLCID
	StrCmp $0 zh 0 +3
	StrCpy $0 1028
		Goto EndLCID
	StrCmp $0 cs 0 +3
	StrCpy $0 1029
		Goto EndLCID		
	StrCmp $0 nl 0 +3
	StrCpy $0 1043
		Goto EndLCID		
	StrCmp $0 fi 0 +3
	StrCpy $0 1035
		Goto EndLCID		
	StrCmp $0 fr 0 +3
	StrCpy $0 1036
		Goto EndLCID		
	StrCmp $0 de 0 +3
	StrCpy $0 1031
		Goto EndLCID		
	StrCmp $0 hu 0 +3
	StrCpy $0 1038
		Goto EndLCID		
	StrCmp $0 it 0 +3
	StrCpy $0 1040
		Goto EndLCID		
	StrCmp $0 ja 0 +3
	StrCpy $0 1041
		Goto EndLCID		
	StrCmp $0 ko 0 +3
	StrCpy $0 1042
		Goto EndLCID		
	StrCmp $0 nb 0 +3
	StrCpy $0 1044
		Goto EndLCID		
	StrCmp $0 pl 0 +3
	StrCpy $0 1045
		Goto EndLCID		
	StrCmp $0 pt 0 +3
	StrCpy $0 2070
		Goto EndLCID		
	StrCmp $0 pt-br 0 +3
	StrCpy $0 1046
		Goto EndLCID		
	StrCmp $0 ru 0 +3
	StrCpy $0 1049
		Goto EndLCID		
	StrCmp $0 es 0 +3
	StrCpy $0 1034
		Goto EndLCID		
	StrCmp $0 es-mx 0 +3
	StrCpy $0 3082
		Goto EndLCID		
	StrCmp $0 sv 0 +3
	StrCpy $0 1053
		Goto EndLCID		
	StrCmp $0 tr 0 +3
	StrCpy $0 1055
		Goto +2
	StrCpy $0 1033
	EndLCID:
	Exch $0
FunctionEnd
Function GetLNG
	!macro _GetLNG _LCID _LNG
		Push ${_LCID}
		Call GetLNG
		Pop ${_LNG}
	!macroend
	!define GetLNG "!insertmacro _GetLNG"
	Exch $0
	StrCmp $0 1033 0 +3
	StrCpy $0 en
		Goto EndLNG
	StrCmp $0 2052 0 +3
	StrCpy $0 zh
		Goto EndLNG
	StrCmp $0 1028 0 +3
	StrCpy $0 zh
		Goto EndLNG
	StrCmp $0 1029 0 +3
	StrCpy $0 cs
		Goto EndLNG		
	StrCmp $0 1043 0 +3
	StrCpy $0 nl
		Goto EndLNG		
	StrCmp $0 1035 0 +3
	StrCpy $0 fi
		Goto EndLNG		
	StrCmp $0 1036 0 +3
	StrCpy $0 fr
		Goto EndLNG		
	StrCmp $0 1031 0 +3
	StrCpy $0 de
		Goto EndLNG		
	StrCmp $0 1038 0 +3
	StrCpy $0 hu
		Goto EndLNG		
	StrCmp $0 1040 0 +3
	StrCpy $0 it
		Goto EndLNG		
	StrCmp $0 1041 0 +3
	StrCpy $0 ja
		Goto EndLNG		
	StrCmp $0 1042 0 +3
	StrCpy $0 ko
		Goto EndLNG		
	StrCmp $0 1044 0 +3
	StrCpy $0 nb
		Goto EndLNG		
	StrCmp $0 1045 0 +3
	StrCpy $0 pl
		Goto EndLNG		
	StrCmp $0 2070 0 +3
	StrCpy $0 pt
		Goto EndLNG		
	StrCmp $0 1046 0 +3
	StrCpy $0 pt-br
		Goto EndLNG		
	StrCmp $0 1049 0 +3
	StrCpy $0 ru
		Goto EndLNG		
	StrCmp $0 1034 0 +3
	StrCpy $0 es
		Goto EndLNG		
	StrCmp $0 3082 0 +3
	StrCpy $0 es-mx
		Goto EndLNG		
	StrCmp $0 1053 0 +3
	StrCpy $0 sv
		Goto EndLNG		
	StrCmp $0 1055 0 +3
	StrCpy $0 tr
		Goto +2
	StrCpy $0 error
	EndLNG:
	Exch $0
FunctionEnd
Function dotNETCheck
	!define CheckDOTNET "!insertmacro _CheckDOTNET"
	!macro _CheckDOTNET _RESULT _VALUE
		Push `${_VALUE}`
		Call dotNETCheck
		Pop ${_RESULT}
	!macroend
	Exch $1
	Push $0
	Push $1
	IntCmp $1 460798 0 +3 0
	StrCpy $0 4.7
		Goto +21
	IntCmp $1 394802 0 +3 0
	StrCpy $0 4.6.2
		Goto +18
	IntCmp $1 394254 0 +3 0
	StrCpy $0 4.6.1
		Goto +15
	IntCmp $1 393295 0 +3 0
	StrCpy $0 4.6
		Goto +12
	IntCmp $1 379893 0 +3 0
	StrCpy $0 4.5.2
		Goto +9
	IntCmp $1 378675 0 +3 0
	StrCpy $0 4.5.1
		Goto +6
	IntCmp $1 378389 0 +3 0
	StrCpy $0 4.5
		Goto +3
	StrCpy $0 ""
	SetErrors
	Pop $1
	Exch $0
FunctionEnd

;= MACROS
;= ################
!define MsgBox "!insertmacro _MsgBox"
!macro _MsgBox
	MessageBox MB_ICONSTOP|MB_TOPMOST `$(LauncherFileNotFound)`
	Call Unload
	Quit
!macroend
!define PortableApps.comLocaleID "!insertmacro _PortableApps.comLocaleID"
!macro _PortableApps.comLocaleID
	${If} ${FileExists} `${XML}`
		${XMLReadText} `${XML}` ${XPATH} $R0
	${Else}
		${XMLReadText} `${DEFXML}` ${XPATH} $R0   
	${EndIf}
	${GetLCID} $R0 $R0
	${SetEnvironmentVariable} PortableApps.comLocaleID $R0
!macroend

;= CUSTOM 
;= ################
${SegmentFile}
${Segment.OnInit}
	${CheckDOTNET} $0 378389
	IfErrors 0 +4
	MessageBox MB_ICONSTOP|MB_TOPMOST `You must have v$0 or greater of the .NET Framework installed. Launcher aborting!`
	Call Unload
	Quit
	${If} ${IsNT}
		${IfNot} ${AtLeastWinVista}
			StrCpy $MissingFileOrPath `Windows Vista or newer`
			${MsgBox}
		${EndIf}
	${Else}
		StrCpy $MissingFileOrPath `Windows Vista or newer`
		${MsgBox}
	${EndIf}
	IfFileExists `${ICSC}\*.*` +3 0
	CreateDirectory `${ICSC}`
	CopyFiles /SILENT `${DEFICSC}\*.*` `${ICSC}`
	Push $0
	${GetParent} `$EXEDIR` $0
	${If} ${FileExists} `${PFM}`
		ReadEnvStr $0 PortableApps.comLocaleID
		${GetLNG} $0 $0
		${If} $0 == error
			${PortableApps.comLocaleID}
		${EndIf}
	${Else}
		${PortableApps.comLocaleID}
	${EndIf}
	Pop $0
	Push $0
	ReadEnvStr $0 PortableApps.comLocaleID
	${GetLNG} $0 $0
	${SetEnvironmentVariable} PAL:LanguageCustom $0
	Pop $0
!macroend
${SegmentUnload}
	FindFirst $0 $1 `$LOCALAPPDATA\Microsoft\*` 
	StrCmpS $0 "" +12
	StrCmpS $1 "" +11
	StrCmpS $1 "." +8
	StrCmpS $1 ".." +7
	StrCpy $2 $1 3
	StrCmpS $2 CLR 0 +5
	IfFileExists `$LOCALAPPDATA\Microsoft\$1\UsageLogs\${APP}.exe.log` 0 +2
	Delete `$LOCALAPPDATA\Microsoft\$1\UsageLogs\*.log`
	RMDir `$LOCALAPPDATA\Microsoft\$1\UsageLogs`
	RMDir `$LOCALAPPDATA\Microsoft\$1`
	FindNext $0 $1
	Goto -10
	FindClose $0
!macroend
