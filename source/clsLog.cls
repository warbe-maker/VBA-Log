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
Option Base 1 ' ensures the index conforms with the column number
' -----------------------------------------------------------------------------------
' Class Module clsLog
'
' Methods/services:
' -----------------
' - ColsAlignmentHeader
' - ColsAlignmentItem
' - ColsHeader
' - ColsWidth             Specifies each columns width, defaults to the width of the
'                         corresponding column header string, is adjusted to the
'                         width of the corresponding header string when already
'                         provided.
' - Dsply
' - Entry                 Appends a line to the log file.
' - Items                 Appends a log line with the provided items aligned in columns
'
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
'                         call of the Items method
' - WithTimeStamp Let     Prefix for log entries when True
'
' W. Rauschenberger, Berlin Apr 2023
' -----------------------------------------------------------------------------------
Private Const DEFAULT_COL_ALIGNMENT_HEADER  As String = "C"
Private Const DEFAULT_COL_ALIGNMENT_ITEM    As String = "L"
Private Const DEFAULT_COL_DELIMITER         As String = "|"
Private Const DEFAULT_COL_MARGIN            As String = " "

Private sColsDelimiter      As String
Private fso                 As New FileSystemObject
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
Public vColsAlignmentHeader As Variant ' Public for test purpose only
Public vColsAlignmentItem   As Variant ' Public for test purpose only
Private vItems              As Variant

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
    sColsMargin = DEFAULT_COL_MARGIN
    sFileFullName = ActiveWorkbook.Path & "\" & DefaultLogFileName
    sFileName = DefaultLogFileName
    sPath = ActiveWorkbook.Path
    sColsDelimiter = DEFAULT_COL_DELIMITER
End Sub

Public Property Get ColsDelimiter(ByVal s As String):   sColsDelimiter = s: End Property

Private Property Get ColHeader(Optional ByVal c_col As Long) As String
    Dim s   As String
    Dim l   As Long
    
    If ArrayIsAllocated(vColsHeader) Then
        If c_col <= UBound(vColsHeader) Then
            s = sColsMargin & vColsHeader(c_col) & sColsMargin
            l = Len(s)
            ColWidth(c_col) = l
            ColHeader = Align(vColsHeader(c_col), ColWidth(c_col), ColAlignmentHeader(c_col), sColsMargin)
        End If
    End If
End Property

Private Property Let ColHeader(Optional ByVal c_col As Long, _
                                        ByVal c_header As String)
    
    If Not ArrayIsAllocated(vColsHeader) Then
        ReDim vColsHeader(c_col)
    Else
        ReDim Preserve vColsHeader(Max(UBound(vColsHeader), c_col))
    End If
    vColsHeader(c_col) = c_header
    
End Property

Private Property Get ColAlignmentHeader(Optional ByVal c_col As Long) As String
    If Not ArrayIsAllocated(vColsAlignmentItem) Then
        ColAlignmentHeader = DEFAULT_COL_ALIGNMENT_HEADER
    Else
        If c_col <= UBound(vColsAlignmentItem) Then
            If vColsAlignmentItem(c_col) = vbNullString Then
                ColAlignmentHeader = DEFAULT_COL_ALIGNMENT_HEADER
            Else
                ColAlignmentHeader = vColsAlignmentItem(c_col)
            End If
        Else
            ColAlignmentHeader = DEFAULT_COL_ALIGNMENT_HEADER
        End If
    End If
            
End Property

Private Property Let ColAlignmentHeader(Optional ByVal c_col As Long, _
                                                ByVal c_align As String)
    Dim lAlignmentHeader As Long
    
    If Not ArrayIsAllocated(vColsAlignmentItem) Then
        ReDim vColsAlignmentItem(c_col)
    Else
        ReDim Preserve vColsAlignmentItem(Max(UBound(vColsAlignmentItem), c_col))
    End If
    If c_align = vbNullString Then
        c_align = DEFAULT_COL_ALIGNMENT_HEADER
    Else
        Select Case UCase(Left(c_align, 1))
            Case "L", "C", "R": vColsAlignmentItem(c_col) = c_align
            Case Else:          vColsAlignmentItem(c_col) = DEFAULT_COL_ALIGNMENT_HEADER
        End Select
    End If
    
End Property

Private Property Get ColAlignmentItem(Optional ByVal c_col As Long) As String
    If Not ArrayIsAllocated(vColsAlignmentItem) Then
        ColAlignmentItem = DEFAULT_COL_ALIGNMENT_ITEM
    Else
        If UBound(vColsAlignmentItem) <= c_col Then
            If vColsAlignmentItem(c_col) = vbNullString Then
                ColAlignmentItem = DEFAULT_COL_ALIGNMENT_ITEM
            Else
                ColAlignmentItem = vColsAlignmentItem(c_col)
            End If
        Else
            ColAlignmentItem = DEFAULT_COL_ALIGNMENT_ITEM
        End If
    End If
            
End Property

Private Property Let ColAlignmentItem(Optional ByVal c_col As Long, _
                                               ByVal c_align As String)
    Dim lAlignmentLine As Long
    
    If Not ArrayIsAllocated(vColsAlignmentItem) Then
        ReDim vColsAlignmentItem(c_col)
    Else
        ReDim Preserve vColsAlignmentItem(Max(UBound(vColsAlignmentItem), c_col))
    End If
    If c_align = vbNullString Then
        c_align = DEFAULT_COL_ALIGNMENT_ITEM
    Else
        Select Case UCase(Left(c_align, 1))
            Case "L", "C", "R": vColsAlignmentItem(c_col) = c_align
            Case Else:          vColsAlignmentItem(c_col) = DEFAULT_COL_ALIGNMENT_ITEM
        End Select
    End If
    
End Property

Public Property Let ColsMargin(ByVal s As String):      sColsMargin = s:        End Property

Private Property Get ColWidth(Optional ByVal c_col As Long) As Long
    If Not ArrayIsAllocated(vColsWidth) Then
        If ArrayIsAllocated(vItems) Then
            If c_col <= UBound(vItems) Then
                ColWidth = Len(vItems(c_col))
            End If
        End If
    Else
        If c_col <= UBound(vColsWidth) Then
            ColWidth = vColsWidth(c_col)
        Else
            If ArrayIsAllocated(vItems) Then
                If c_col <= UBound(vItems) Then
                    ColWidth = Len(vItems(c_col))
                End If
            End If
        End If
    End If
End Property

Private Property Let ColWidth(Optional ByVal c_col As Long, _
                                       ByVal c_width As Long)
    Dim lWidth As Long
    
    If Not ArrayIsAllocated(vColsWidth) Then
        ReDim vColsWidth(c_col)
    Else
        If c_col > UBound(vColsWidth) Then
            ReDim Preserve vColsWidth(Max(UBound(vColsWidth), c_col))
        End If
    End If
    lWidth = vColsWidth(c_col)
    vColsWidth(c_col) = Max(lWidth, c_width)

End Property

Private Property Get DefaultLogFileName() As String
    DefaultLogFileName = fso.GetBaseName(ActiveWorkbook.Name) & ".log"
End Property

Public Property Let Entry(ByVal s As String):   WriteEntry s:                   End Property

Public Property Get FileFullName() As String:   FileFullName = sFileFullName:   End Property

Public Property Let FileFullName(ByVal s As String)
' ----------------------------------------------------------------------------
' Explicitely specifies the log file's name and location. This is an
' alternative to the provision of FileName and Path
' ----------------------------------------------------------------------------
    With fso
        sFileName = .GetFileName(s)
        sPath = .GetParentFolderName(s)
        If Not .FileExists(sFileFullName) Then .CreateTextFile sFileFullName
    End With
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

Public Property Let Title(ByVal s As String)
' ----------------------------------------------------------------------------
' Alternatively to the "Service" property!
' ----------------------------------------------------------------------------
    bHeaderDue = s <> sTitle
    sTitle = s
    vItems = vbNullString
End Property

Public Property Let WithTimeStamp(ByVal b As Boolean)
    bWithTimeStamp = b
End Property

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
    Dim sItem           As String
    Dim sFill           As String
    
    Select Case Left(a_mode, 1)
        Case "L"
            sItem = a_margin & Trim(a_strng) & a_margin
            If Len(sItem) >= a_lngth Then
                Align = VBA.Left$(sItem, a_lngth)
            Else
                sFill = VBA.String$(a_lngth - (Len(sItem)), a_fill)
                Align = VBA.Left(sItem & sFill, a_lngth)
            End If
        Case "C"
            sItem = a_margin & Trim(a_strng) & a_margin
            If Len(sItem) >= a_lngth Then
                Align = Left$(sItem, a_lngth)
            Else
                sFill = VBA.String$(Int((a_lngth - Len(sItem)) / 2), a_fill)
                Align = VBA.Right$(sFill & sItem & sFill, a_lngth)
            End If
            If Len(Align) < a_lngth Then
                Align = Align & VBA.String$(a_lngth - Len(Align), a_fill)
            End If
        Case "R"
            sItem = a_margin & Trim(a_strng) & a_margin
            If Len(sItem) >= a_lngth Then
                Align = VBA.Right$(sItem, a_lngth)
            Else
                sFill = VBA.String$(a_lngth - (Len(sItem)), a_fill)
                Align = VBA.Right$(sFill & sItem & sFill, a_lngth)
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

Public Sub ColsAlignmentHeader(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' When ColWidths are not provided the columns width defaults to the width of the
' colum headers plus 2 margin spaces.
' -----------------------------------------------------------------------------------
    Dim i As Long
    
    For i = LBound(c_align) To UBound(c_align)
        ColAlignmentHeader(i + 1) = c_align(i)
    Next i
    
End Sub

Public Sub ColsAlignmentItem(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' When ColWidths are not provided the columns width defaults to the width of the
' colum headers plus 2 margin spaces.
' -----------------------------------------------------------------------------------
    Dim i As Long
    
    For i = LBound(c_align) To UBound(c_align)
        ColAlignmentItem(i + 1) = c_align(i)
    Next i
    
End Sub

Private Property Let ColItem(Optional ByVal c_col As Long, _
                                     ByVal c_item As String)
    If Not ArrayIsAllocated(vItems) Then
        ReDim vItems(c_col)
    Else
        ReDim Preserve vItems(Max(UBound(vItems), c_col))
    End If
    vItems(c_col) = c_item

End Property

Private Property Get ColItem(Optional ByVal c_col As Long) As String
    If ArrayIsAllocated(vItems) Then
        If c_col <= UBound(vItems) Then
            ColItem = Align(vItems(c_col), ColWidth(c_col), ColAlignmentItem(c_col), sColsMargin)
        End If
    End If
End Property

Public Sub ColsHeader(ParamArray c_headers() As Variant)
' -----------------------------------------------------------------------------------
' Note: ColsWidth defaults to the maximum of an already specified width and the width
'       if the corresponding header string.
' -----------------------------------------------------------------------------------
    Const PROC = "ColsHeader"
    
    On Error GoTo eh
    Dim i As Long
    
    vColsHeader = vbNullString
    For i = LBound(c_headers) To UBound(c_headers)
        ColHeader(i + 1) = c_headers(i)
        ColWidth(i + 1) = Len(c_headers(i))
    Next i
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ColsWidth(ParamArray c_widths() As Variant)
' -----------------------------------------------------------------------------------
' Specifies the width of n columns. When not provided the column width defaults to
' width of the column headers
' -----------------------------------------------------------------------------------
    Const PROC = "ColsWidth"
    
    On Error GoTo eh
    Dim i As Long
    Dim l As Long
    
    vColsWidth = vbNullString
    For i = LBound(c_widths) To UBound(c_widths)
        l = ColWidth(i + 1)
        If ArrayIsAllocated(vColsHeader) Then
            If i + 1 <= UBound(vColsHeader) Then
                ColWidth(i + 1) = Max(Len(vColsHeader(i + 1)), l, c_widths(i))
            Else
                ColWidth(i + 1) = Max(l, c_widths(i))
            End If
        Else
            ColWidth(i + 1) = Max(l, c_widths(i))
        End If
    Next i

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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
    If err_source = vbNullString Then err_source = Err.Source
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

Public Sub Items(ParamArray e() As Variant)
' -----------------------------------------------------------------------------------
' Append the items (i) to the log file aligned in columns.
' -----------------------------------------------------------------------------------
    Dim i As Long
    
    If bHeaderDue Then WriteHeader
    For i = LBound(e) To UBound(e)
        ColItem(i + 1) = e(i)
    Next i
    WriteItems
    
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

Private Sub ProvideColsAlignmentHeader()
    Dim i   As Long
    If Not ArrayIsAllocated(vColsAlignmentHeader) Then
        ReDim vColsAlignmentHeader(UBound(vColsHeader))
        For i = LBound(vColsAlignmentHeader) To UBound(vColsAlignmentHeader)
            vColsAlignmentHeader(i) = DEFAULT_COL_ALIGNMENT_HEADER
        Next i
    End If
End Sub

Private Sub ProvideColsAlignmentItem()
    Dim i   As Long
    If Not ArrayIsAllocated(vColsAlignmentItem) Then
        ReDim vColsAlignmentItem(UBound(vColsHeader))
        For i = LBound(vColsAlignmentItem) To UBound(vColsAlignmentItem)
            vColsAlignmentItem(i) = DEFAULT_COL_ALIGNMENT_ITEM
        Next i
    End If
End Sub

Private Sub ProvideLogFile()
    With fso
        If Not .FileExists(sFileFullName) Then
            .CreateTextFile sFileFullName
        Else
            If VBA.DateDiff("d", .GetFile(sFileFullName).DateCreated, Now()) > lKeepDays Then
                .DeleteFile sFileFullName
                .CreateTextFile sFileFullName
            End If
        End If
        If .GetFile(sFileFullName).Size = 0 _
        Then sServiceDelimiter = vbNullString _
        Else sServiceDelimiter = "="
    End With
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

Private Function TimeStamp() As String ' Public for test purpose only
    If bWithTimeStamp Then
        If ArrayIsAllocated(vColsHeader) _
        Then TimeStamp = Format(Now(), "yy-mm-dd-hh:mm:ss") & " " & sColsDelimiter _
        Else TimeStamp = Format(Now(), "yy-mm-dd-hh:mm:ss") & " "
    End If
End Function

Private Sub WriteItems()
' -----------------------------------------------------------------------------------
' Add a line to the log file aligned in columns.
' - Any string exceeding the number of provided column headers and column widths is
'   ignored!
' - When no column headers had been provided (method ColsHeader) an error is raised.
' -----------------------------------------------------------------------------------
    Const PROC = "WriteItems"
    
    On Error GoTo eh
    Dim i           As Long
    Dim s           As String
    Dim sColDelim   As String
    Dim v           As Variant
    Dim sElement    As String
    
    v = Split(sHeaderText, sColsDelimiter)
    For i = LBound(vItems) To UBound(vItems)
        sElement = vItems(i)
        If Len(sElement) > 0 Then
            If Left(sElement, 1) <> sColsMargin Then
                sElement = sColsMargin & sElement
            End If
        End If
        s = s & sColDelim & ColItem(i)
        sColDelim = sColsDelimiter
    Next i
    WriteLog s
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub WriteLog(ByVal s As String)
' ----------------------------------------------------------------------------
' Writes the string (s) into the file (ft_file) which might be a file
' object or a file's full name.
' Note: ft_split is not used but specified to comply with Property Get.
' ----------------------------------------------------------------------------
    Const PROC = "WriteLog"
    
    On Error GoTo eh
    Dim ts  As TextStream
   
    ProvideLogFile
    Set ts = fso.OpenTextFile(FileName:=sFileFullName, IOMode:=ForAppending)
    ts.WriteLine TimeStamp & s

xt: ts.Close
    Set ts = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub WriteEntry(ByVal s As String)
' ----------------------------------------------------------------------------
' Writes the string (s) into the file (ft_file) which might be a file
' object or a file's full name.
' Note: ft_split is not used but specified to comply with Property Get.
' ----------------------------------------------------------------------------
    Const PROC = "WriteEntry"
    
    On Error GoTo eh
    Dim ts  As TextStream
   
    If bHeaderDue Then WriteHeader
    ProvideLogFile
    Set ts = fso.OpenTextFile(FileName:=sFileFullName, IOMode:=ForAppending)
    ts.WriteLine TimeStamp & s

xt: ts.Close
    Set ts = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub WriteHeader()
' -----------------------------------------------------------------------------------
' When not performed explicitely it will be performed along with the first _Items_
' method call when bHeaderDue is True.
' - Writes a header based on the strings provided in their width by default, by the
'   width explicitely provided otherwise
' - When the KeepDays limit is reached writing the header establishes a new log file,
'   else a ==== delimiting line in the width of the header line is written, followed
'   by a service header when Service had been provided
' -----------------------------------------------------------------------------------
    Const PROC = "WriteHeader"
    
    On Error GoTo eh
    Dim sHeaderLine As String
    Dim i           As Long
    Dim sColDelim   As String
    Dim v           As Variant
    
'    ProvideLogFile
    If Not ArrayIsAllocated(vColsHeader) Then
        If sServiceDelimiter <> vbNullString Then
            WriteEntry String(Len(sHeaderText), sServiceDelimiter)
            sServiceDelimiter = vbNullString
        End If
        GoTo xt
    End If
    
    sHeaderText = vbNullString
    For i = LBound(vColsHeader) To UBound(vColsHeader)
        sHeaderText = sHeaderText & sColDelim & ColHeader(i)
        sColDelim = sColsDelimiter
    Next i
    sColDelim = vbNullString
    
    v = Split(sHeaderText, sColsDelimiter)
    For i = LBound(v) To UBound(v)
        sHeaderLine = sHeaderLine & sColDelim & String(Len(v(i)), "-")
        sColDelim = "+"
    Next i
    If sTitle <> vbNullString Then
        sTitle = Align(sTitle, Len(sHeaderText), "C", " ", "-")
    End If
    
    If sServiceDelimiter <> vbNullString Then
        WriteLog String(Len(sHeaderText), sServiceDelimiter)
        sServiceDelimiter = vbNullString
    End If
    If sTitle <> vbNullString Then WriteLog sTitle
    WriteLog sHeaderText
    WriteLog sHeaderLine
    bHeaderDue = False ' may be set to True again when a new Title is provided
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
    Set vColsAlignmentHeader = Nothing
    Set vColsAlignmentItem = Nothing
    Set vColsHeader = Nothing
    Set vColsWidth = Nothing
    Set vItems = Nothing
End Sub
