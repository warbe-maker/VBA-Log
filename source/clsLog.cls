VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' -----------------------------------------------------------------------------------
' Class Module clsLog
'
' Methods/services:
' -----------------
' - ColsAlignmnetHeader
' - ColsAlignmnetLine
' - ColsHeader
' - ColsWidth             Specifies each columns width, defaults to the width of the
'                         corresponding column header string, is adjusted to the
'                         width of the corresponding header string when already
'                         provided.
' - Dsply
' - Entry                 Appends an entry to the log file. A single line when only
'                         one string is provided or a column aligned string when more
'                         than one string is provided (ParamArray)'
' Properties:
' -----------
' - ColsMargin            Defaults to " ", may be set to vbNullString, when provided
'                         adds to the width of the header string
' - FileFullName  Get/Let
' - FileName      Let
' - KeepDays      Let
' - LogFile       Get
' - Path          Let
' - Title         Let     Triggers the automated writing of a header with the firs
'                         call of the Entry method
' - WithTimeStamp Let     Prefix for log entries when True
'
' W. Rauschenberger, Berlin Apr 2023
' -----------------------------------------------------------------------------------
Private Const COLS_ALGN_DFLT_HEADER As String = "C"
Private Const COLS_ALGN_DFLT_LINE   As String = "L"
Private Const COLS_DELIMITER_DFLT   As String = "|"
Private Const COLS_MARGIN_DFLT      As String = " "
Private Const LOG_FILE_NAME_DFLT    As String = "Service.log"

Private bHeaderDue          As Boolean
Private bNewLog             As Boolean
Private bWithTimeStamp      As Boolean
Private lKeepDays           As Long
Private sColsMargin         As String
Private sEntry              As String
Private sFileFullName       As String
Private sFileName           As String
Private sPath               As String
Private sHeaderText         As String
Private sTitle              As String
Private sServiceDelimiter   As String
Private sServicedItem       As String
Private sServicedItemName   As String
Private sServicedItemType   As String
Public vColsWidth           As Variant ' Public for test purpose only
Public vColsHeader          As Variant ' Public for test purpose only
Public vColsAlignmnetHeader As Variant ' Public for test purpose only
Public vColsAlignmnetLine   As Variant ' Public for test purpose only
Private vEntry              As Variant

#If Not MsgComp = 1 Then
    ' -------------------------------------------------------------------------------
    ' The 'minimum error handling' aproach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed which must
    ' be indicated by the Conditional Compile Argument MsgComp = 1.
    ' See https://github.com/warbe-maker/Common-VBA-Message-Service
    ' -------------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Private Declare PtrSafe Function apiShellExecute Lib "shell32.dll" _
    Alias "ShellExecuteA" _
    (ByVal hWnd As Long, _
    ByVal lpOperation As String, _
    ByVal lpFile As String, _
    ByVal lpParameters As String, _
    ByVal lpDirectory As String, _
    ByVal nShowCmd As Long) _
    As Long

Private Sub Class_Initialize()
' -----------------------------------------------------------------------------------
' Initializing defaults
' -----------------------------------------------------------------------------------
    bHeaderDue = True
    bNewLog = True
    bWithTimeStamp = True
    lKeepDays = 1
    sColsMargin = COLS_MARGIN_DFLT
    sFileFullName = ThisWorkbook.Path & "\" & LOG_FILE_NAME_DFLT
    sFileName = LOG_FILE_NAME_DFLT
    sPath = ThisWorkbook.Path
End Sub

Public Function TimeStamp() As String ' Public for test purpose only
    If bWithTimeStamp Then
        If ArrayIsAllocated(vColsHeader) _
        Then TimeStamp = Format(Now(), "yy-mm-dd-hh:mm:ss") & " " & COLS_DELIMITER_DFLT _
        Else TimeStamp = Format(Now(), "yy-mm-dd-hh:mm:ss") & " "
    End If
End Function

Public Property Let ColsMargin(ByVal s As String)
    sColsMargin = s
End Property

Public Sub Entry(ParamArray e() As Variant)
' -----------------------------------------------------------------------------------
' Append an entry to the log file.
' -----------------------------------------------------------------------------------
    Static sItem    As String
    Dim sCompName   As String
    
    If bHeaderDue Then WriteHeader
    vEntry = e
    If LBound(vEntry) = UBound(vEntry) _
    Then WriteEntry vEntry(0) _
    Else WriteColumned
    
End Sub

Friend Property Get FileFullName() As String
    FileFullName = sFileFullName
End Property

Friend Property Let FileFullName(ByVal s As String)
' ----------------------------------------------------------------------------
' Explicitely specifies the log file's name and location. This is an
' alternative to the provision of FileName and Path
' ----------------------------------------------------------------------------
    Dim fso As New FileSystemObject
    With fso
        sFileName = .GetFileName(s)
        sPath = .GetParentFolderName(s)
        If Not .FileExists(sFileFullName) Then .CreateTextFile sFileFullName
    End With
    Set fso = Nothing
End Property

Public Property Let FileName(ByVal s As String)
    sFileName = s
    sFileFullName = Replace(sPath & "\" & sFileName, "\\", "\")
End Property

Public Property Let KeepDays(ByVal l As Long): lKeepDays = l: End Property

Friend Property Get LogFile() As File
' -----------------------------------------------------------------------------------
' Returns the log file as file object
' -----------------------------------------------------------------------------------
    With New FileSystemObject
        If Not .FileExists(sFileFullName) Then .CreateTextFile sFileFullName
        Set LogFile = .GetFile(sFileFullName)
    End With

End Property

Public Property Let Path(ByVal v As Variant)
' -----------------------------------------------------------------------------------
' Specifies the location (folder) for the log file based on the provided information
' which may be a string, a Workbook, or a folder object.
' -----------------------------------------------------------------------------------
    Const PROC = "Path-Let"
    Dim fso As New FileSystemObject
    Dim wbk As Workbook
    Dim fld As Folder
    
    Select Case VarType(v)
        Case VarType(v) = vbString
            sPath = v
        Case VarType(v) = vbObject
            If TypeOf v Is Workbook Then
                Set wbk = v
                sPath = wbk.Path
            ElseIf TypeOf v Is Folder Then
                Set fld = v
                sPath = fld.Path
            Else
                Err.Raise AppErr(1), ErrSrc(PROC), "The provided argument is neither a string specifying a " & _
                                                   "folder's path, nor a Workbook object, nor a Folder object!"
            End If
    End Select
    
End Property

Public Property Let WithTimeStamp(ByVal b As Boolean)
    bWithTimeStamp = b
End Property

Public Property Get WithTimeStamp() As Boolean
    WithTimeStamp = bWithTimeStamp
End Property

Public Property Let Title(ByVal s As String)
' ----------------------------------------------------------------------------
' Alternatively to the "Service" property!
' ----------------------------------------------------------------------------
    bHeaderDue = s <> sTitle
    sTitle = s
    vColsHeader = vbNullString
    vColsWidth = vbNullString
    vEntry = vbNullString
End Property

Private Sub WriteColumned()
' -----------------------------------------------------------------------------------
' Add a line to the log file aligned in columns.
' - Any string exceeding the number of provided column headers and column widths is
'   ignored!
' - When no column headers had been provided (method ColsHeader) an error is raised.
' -----------------------------------------------------------------------------------
    Const PROC = "WriteColumned"
    
    On Error GoTo eh
    Dim i           As Long
    Dim s           As String
    Dim sColDelim   As String
    Dim v           As Variant
    Dim sElement    As String
    
    If Not ArrayIsAllocated(vColsHeader) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No column header provided. This service writes log entries in " & _
                                            "columns for which column header is obligatory " & _
                                            "(for unformatted of preformatted log entries use the WriteEntry service)!"
    ProvideColsAlignmnetLine
    v = Split(sHeaderText, COLS_DELIMITER_DFLT)
    For i = LBound(vEntry) To Min(UBound(vEntry), UBound(vColsWidth))
        sElement = vEntry(i)
        If Len(sElement) > 0 Then
            If Left(sElement, 1) <> sColsMargin Then
                sElement = sColsMargin & sElement
            End If
        End If
        s = s & sColDelim & Align(sElement, Len(v(i)), vColsAlignmnetLine(i), sColsMargin)
        sColDelim = COLS_DELIMITER_DFLT
    Next i
    WriteEntry s
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Align(ByVal a_strng As String, _
                       ByVal a_lngth As Long, _
              Optional ByVal a_mode As String = "L", _
              Optional ByVal a_margin As String = vbNullString, _
              Optional ByVal a_fill As String = " ") As String
' ----------------------------------------------------------------------------
' Returns a string (a_strng) with a lenght (a_lngth) aligned (a_mode) filled
' with characters (a_fill).
' ----------------------------------------------------------------------------
    Dim SpaceLeft       As Long
    Dim LengthRemaining As Long
    
    Select Case Left(a_mode, 1)
        Case "L"
            If Len(a_strng & a_margin) >= a_lngth _
            Then Align = VBA.Left$(a_strng & a_margin, a_lngth) _
            Else Align = a_strng & a_margin & VBA.String$(a_lngth - (Len(a_strng & a_margin)), a_fill)
        Case "R"
            If Len(a_margin & a_strng) >= a_lngth _
            Then Align = VBA.Left$(a_margin & a_strng, a_lngth) _
            Else Align = VBA.String$(a_lngth - (Len(a_margin & a_strng)), a_fill) & a_margin & a_strng
        Case "C"
            If Len(a_margin & a_strng & a_margin) >= a_lngth Then
                Align = a_margin & Left$(a_strng, a_lngth) & a_margin
            Else
                SpaceLeft = Max(1, ((a_lngth - Len(a_strng) - (2 * Len(a_margin))) / 2))
                Align = VBA.String$(SpaceLeft, a_fill) & a_margin & a_strng & a_margin & VBA.String$(SpaceLeft, a_fill)
                Align = VBA.Right$(Align, a_lngth)
            End If
    End Select

End Function

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    AppErr = IIf(app_err_no < 0, app_err_no - vbObjectError, vbObjectError - app_err_no)
End Function

Private Function ArrayIsAllocated(arr As Variant) As Boolean
    
    On Error Resume Next
    ArrayIsAllocated = _
    IsArray(arr) _
    And Not IsError(LBound(arr, 1)) _
    And LBound(arr, 1) <= UBound(arr, 1)
    
End Function

Public Sub ColsAlignmnetHeader(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' When ColWidths are not provided the columns width defaults to the width of the
' colum headers plus 2 margin spaces.
' -----------------------------------------------------------------------------------
    Dim i As Long
    vColsAlignmnetHeader = c_align
    '~~ Set any invalid entry to the default "L" (align left)
    For i = LBound(vColsAlignmnetHeader) To UBound(vColsAlignmnetHeader)
        If Len(vColsAlignmnetHeader(i)) >= 1 Then
            Select Case Left(vColsAlignmnetHeader(i), 1)
                Case "L", "R", "C"
                Case Else: vColsAlignmnetHeader(i) = COLS_ALGN_DFLT_HEADER
            End Select
        Else
            vColsAlignmnetHeader(i) = COLS_ALGN_DFLT_HEADER
        End If
    Next i
    
End Sub

Public Sub ColsAlignmnetLine(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' When ColWidths are not provided the columns width defaults to the width of the
' colum headers plus 2 margin spaces.
' -----------------------------------------------------------------------------------
    Dim i As Long
    vColsAlignmnetLine = c_align
    '~~ Set any invalid entry to the default "L" (align left)
    For i = LBound(vColsAlignmnetLine) To UBound(vColsAlignmnetLine)
        If Len(vColsAlignmnetLine(i)) >= 1 Then
            Select Case Left(vColsAlignmnetLine(i), 1)
                Case "L", "R", "C"
                Case Else: vColsAlignmnetLine(i) = COLS_ALGN_DFLT_LINE
            End Select
        Else
            vColsAlignmnetLine(i) = COLS_ALGN_DFLT_LINE
        End If
    Next i
    
End Sub

Public Sub ColsHeader(ParamArray c_headers() As Variant)
' -----------------------------------------------------------------------------------
' When ColWidths are not provided the columns width defaults to the width of the
' colum headers plus 2 margin spaces.
' -----------------------------------------------------------------------------------
    Dim i As Long
    Dim l As Long
    
    vColsHeader = vbNullString
    vColsHeader = c_headers
    If Not ArrayIsAllocated(vColsWidth) Then
        ReDim vColsWidth(UBound(vColsHeader))
        For i = LBound(vColsHeader) To UBound(vColsHeader)
            vColsWidth(i) = Len(vColsHeader(i))
        Next i
    Else
        For i = LBound(vColsHeader) To Min(UBound(vColsHeader), UBound(vColsWidth))
            l = vColsWidth(i)
            vColsWidth(i) = Max(Len(vColsHeader(i)), l)
        Next i
    End If
    
End Sub

Public Sub ColsWidth(ParamArray c_widths() As Variant)
' -----------------------------------------------------------------------------------
' Specifies the width of n columns. When not provided the column width defaults to
' width of the column headers
' -----------------------------------------------------------------------------------
    Dim i As Long
    Dim l As Long
    
    vColsWidth = vbNullString
    vColsWidth = c_widths
    If ArrayIsAllocated(vColsHeader) Then
        For i = LBound(vColsHeader) To Min(UBound(vColsHeader), UBound(vColsWidth))
            l = vColsWidth(i)
            vColsWidth(i) = Max(Len(vColsHeader(i)), l)
        Next i
    End If

End Sub

Public Sub Dsply()
' -----------------------------------------------------------------------------------
'
' -----------------------------------------------------------------------------------
    ShellRun sFileFullName
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays a debugging option
' button when the Conditional Compile Argument 'Debugging = 1', displays an
' optional additional "About:" section when the err_dscrptn has an additional
' string concatenated by two vertical bars (||), and displays the error message
' by means of VBA.MsgBox when neither the Common Component mErH (indicated by
' the Conditional Compile Argument "ErHComp = 1", nor the Common Component mMsg
' (idicated by the Conditional Compile Argument "MsgComp = 1") is installed.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'       ErrSrc  To provide an unambiguous procedure name by prefixing is with
'               the module name.
'
' W. Rauschenberger Berlin, Apr 2023
'
' See: https://github.com/warbe-maker/Common-VBA-Error-Services
' ------------------------------------------------------------------------------' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Consider extra information is provided with the error description
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsLog" & "." & sProc
End Function

Private Sub WriteEntry(ByVal s As String)
' ----------------------------------------------------------------------------
' Writes the string (s) into the file (ft_file) which might be a file
' object or a file's full name.
' Note: ft_split is not used but specified to comply with Property Get.
' ----------------------------------------------------------------------------
    Const PROC = "WriteEntry"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    Dim ts  As TextStream
    Dim sFl As String
   
    ProvideLogFile
    With fso
        sFl = sFileFullName
        Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForAppending)
    End With
    
    ts.WriteLine TimeStamp & s

xt: ts.Close
    Set fso = Nothing
    Set ts = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function Max(ParamArray va() As Variant) As Variant
' ----------------------------------------------------------------------------
' Returns the maximum value of all values provided (va).
' ----------------------------------------------------------------------------
    Dim v As Variant
    
    Max = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v > Max Then Max = v
    Next v
    
End Function

Private Function Min(ParamArray va() As Variant) As Variant
' --------------------------------------------------------
' Returns the minimum (smallest) of all provided values.
' --------------------------------------------------------
    Dim v As Variant
    
    Min = va(LBound(va)): If LBound(va) = UBound(va) Then Exit Function
    For Each v In va
        If v < Min Then Min = v
    Next v
    
End Function

Private Sub ProvideColsAlignmnetHeader()
    Dim i   As Long
    If Not ArrayIsAllocated(vColsAlignmnetHeader) Then
        ReDim vColsAlignmnetHeader(UBound(vColsHeader))
        For i = LBound(vColsAlignmnetHeader) To UBound(vColsAlignmnetHeader)
            vColsAlignmnetHeader(i) = COLS_ALGN_DFLT_HEADER
        Next i
    End If
End Sub

Private Sub ProvideColsAlignmnetLine()
    Dim i   As Long
    If Not ArrayIsAllocated(vColsAlignmnetLine) Then
        ReDim vColsAlignmnetLine(UBound(vColsHeader))
        For i = LBound(vColsAlignmnetLine) To UBound(vColsAlignmnetLine)
            vColsAlignmnetLine(i) = COLS_ALGN_DFLT_LINE
        Next i
    End If
End Sub

Private Sub ProvideLogFile()
    Dim fso As New FileSystemObject
    With fso
        If .FileExists(sFileFullName) Then
            If VBA.DateDiff("d", .GetFile(sFileFullName).DateCreated, Now()) > lKeepDays _
            Then .DeleteFile sFileFullName
        End If
        If Not .FileExists(sFileFullName) Then
            .CreateTextFile sFileFullName
            sServiceDelimiter = vbNullString
        Else
            If .GetFile(sFileFullName).Size = 0 _
            Then sServiceDelimiter = vbNullString _
            Else sServiceDelimiter = "="
        End If
    End With
    Set fso = Nothing
End Sub

Private Sub ShellRun(ByVal sr_string As String, _
            Optional ByVal sr_show_how As Long = 1)
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------
    Const PROC = "ShellRun"
    Const ERROR_SUCCESS = 32&
    Const ERROR_NO_ASSOC = 31&
    Const ERROR_OUT_OF_MEM = 0&
    Const ERROR_FILE_NOT_FOUND = 2&
    Const ERROR_PATH_NOT_FOUND = 3&
    Const ERROR_BAD_FORMAT = 11&
    
    On Error GoTo eh
    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, sr_string, vbNullString, vbNullString, sr_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       Err.Raise lRet, ErrSrc(PROC), "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   Err.Raise lRet, ErrSrc(PROC), "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   Err.Raise lRet, ErrSrc(PROC), "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       Err.Raise lRet, ErrSrc(PROC), "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & sr_string, 1)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub WriteHeader()
' -----------------------------------------------------------------------------------
' When not performed explicitely it will be performed along with the first _Entry_
' method call when bHeaderDue is True.
' - Writes a header based on the strings provided in their width by default, by the
'   width explicitely provided otherwise
' - When the KeepDays limit is reached writing the header establishes a new log file,
'   else a ==== delimiting line in the width of the header line is written, followed
'   by a service header when Service had been provided
' -----------------------------------------------------------------------------------
    Const PROC = "WriteHeader"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim sHeaderLine As String
    Dim sHeaderSrvc As String
    
    Dim i           As Long
    Dim sColDelim   As String
    Dim v           As Variant
    
    ProvideLogFile
    If Not ArrayIsAllocated(vColsHeader) Then GoTo xt
    ProvideColsAlignmnetHeader
    
    sHeaderText = vbNullString
    For i = LBound(vColsHeader) To Min(UBound(vColsHeader), UBound(vColsWidth))
        sHeaderText = sHeaderText & sColDelim & Align(vColsHeader(i), vColsWidth(i), vColsAlignmnetHeader(i), sColsMargin)
        sColDelim = COLS_DELIMITER_DFLT
    Next i
    sColDelim = vbNullString
    
    v = Split(sHeaderText, COLS_DELIMITER_DFLT)
    For i = LBound(v) To UBound(v)
        sHeaderLine = sHeaderLine & sColDelim & String(Len(v(i)), "-")
        sColDelim = "+"
    Next i
    If sTitle <> vbNullString Then
        sHeaderSrvc = Align(sTitle, Len(sHeaderText), "C", " ", "-")
    End If
    
    If sServiceDelimiter <> vbNullString Then
        WriteEntry String(Len(sHeaderText), sServiceDelimiter)
        sServiceDelimiter = vbNullString
    End If
    If sHeaderSrvc <> vbNullString Then WriteEntry sHeaderSrvc
    WriteEntry sHeaderText
    WriteEntry sHeaderLine
    bHeaderDue = False ' may be set to True again when a new Title is provided
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

