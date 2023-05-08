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
' - BeginOfNewLogEntries  Delimiter line, title, header, imlicitely when applicable
'                         or explicitely when desired.
' - Dsply
' - Entry                 - When a single string without any | is provided the string
'                          is appended as a new line to the log file
'                         - When multiple strings are provided or the string contains
'                           vertical bars (|) the indicated items are writen aligned
'                           in columns to the log-file
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

Private aColsAlignmentHeader As Variant ' Public for test purpose only
Private aColsAlignmentItem   As Variant ' Public for test purpose only
Private aColsHeader          As Variant ' Public for test purpose only
Private aColsWidth           As Variant ' Public for test purpose only
Private aItems              As Variant
Private aTitle              As Variant
Private bAlignedByColumns   As Boolean
Private bNewLog             As Boolean
Private bWithTimeStamp      As Boolean
Private fso                 As New FileSystemObject
Private lKeepDays           As Long
Private sColsDelimiter      As String
Private sColsMargin         As String
Private sEntry              As String
Private sFileFullName       As String
Private sFileName           As String
Private sHeaderLine         As String
Private sPath               As String
Private sServicedItem       As String
Private sServicedItemName   As String
Private sServicedItemType   As String

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
    bNewLog = True
    bWithTimeStamp = True
    lKeepDays = 1
    sColsMargin = DEFAULT_COL_MARGIN
    sFileFullName = ActiveWorkbook.Path & "\" & DefaultLogFileName
    sFileName = DefaultLogFileName
    sPath = ActiveWorkbook.Path
    sColsDelimiter = DEFAULT_COL_DELIMITER
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
    Set aColsAlignmentHeader = Nothing
    Set aColsAlignmentItem = Nothing
    Set aColsHeader = Nothing
    Set aColsWidth = Nothing
    Set aItems = Nothing
    Set aTitle = Nothing
End Sub

Public Property Get ColsDelimiter(ByVal s As String):   sColsDelimiter = s: End Property

Private Property Get ColHeader(Optional ByVal c_col As Long) As String
    Dim s   As String
    Dim l   As Long
    
    If IsAllocated(aColsHeader) Then
        If c_col <= UBound(aColsHeader) Then
            s = sColsMargin & aColsHeader(c_col) & sColsMargin
            l = Len(s)
            l = Max(ColWidth(c_col), l) ' width with marging may be greater than originally calculated
            ColWidth(c_col) = l
            ColHeader = Align(aColsHeader(c_col), l, ColAlignmentHeader(c_col), sColsMargin)
        End If
    End If
End Property

Private Property Let ColHeader(Optional ByVal c_col As Long, _
                                        ByVal c_header As String)
    
    If Not IsAllocated(aColsHeader) Then
        ReDim aColsHeader(c_col)
    Else
        ReDim Preserve aColsHeader(Max(UBound(aColsHeader), c_col))
    End If
    aColsHeader(c_col) = c_header
    
    
End Property

Private Property Get ColAlignmentHeader(Optional ByVal c_col As Long) As String
    If IsAllocated(aColsAlignmentHeader) Then
        If c_col <= UBound(aColsAlignmentHeader) Then
            ColAlignmentHeader = aColsAlignmentHeader(c_col)
        End If
    End If
End Property

Private Property Let ColAlignmentHeader(Optional ByVal c_col As Long, _
                                                ByVal c_align As String)
    If Not IsAllocated(aColsAlignmentHeader) Then
        ReDim aColsAlignmentHeader(c_col)
    ElseIf c_col > UBound(aColsAlignmentHeader) Then
        ReDim Preserve aColsAlignmentHeader(c_col)
    End If
    aColsAlignmentHeader(c_col) = c_align
End Property

Private Property Get ColAlignmentItem(Optional ByVal c_col As Long) As String
    If IsAllocated(aColsAlignmentItem) Then
        If c_col <= UBound(aColsAlignmentItem) Then
            ColAlignmentItem = aColsAlignmentItem(c_col)
        Else
            ColAlignmentItem = DEFAULT_COL_ALIGNMENT_ITEM
        End If
    Else
        ColAlignmentItem = DEFAULT_COL_ALIGNMENT_ITEM
    End If
End Property

Private Property Let ColAlignmentItem(Optional ByVal c_col As Long, _
                                               ByVal c_align As String)
    Dim lAlignmentLine As Long
    
    If Not IsAllocated(aColsAlignmentItem) Then
        ReDim aColsAlignmentItem(c_col)
    ElseIf c_col > UBound(aColsAlignmentItem) Then
        ReDim Preserve aColsAlignmentItem(c_col)
    End If
    aColsAlignmentItem(c_col) = c_align
    
End Property

Public Property Let ColsMargin(ByVal s As String):      sColsMargin = s:        End Property

Private Property Get ColWidth(Optional ByVal c_col As Long) As Long
    If Not IsAllocated(aColsWidth) Then
        If IsAllocated(aItems) Then
            If c_col <= UBound(aItems) Then
                ColWidth = Len(aItems(c_col))
            End If
        End If
    Else
        If c_col <= UBound(aColsWidth) Then
            ColWidth = aColsWidth(c_col)
        Else
            If IsAllocated(aItems) Then
                If c_col <= UBound(aItems) Then
                    ColWidth = Len(aItems(c_col))
                End If
            End If
        End If
    End If
End Property

Private Property Let ColWidth(Optional ByVal c_col As Long, _
                                       ByVal c_width As Long)
    Dim lWidth As Long
    
    If Not IsAllocated(aColsWidth) Then
        ReDim aColsWidth(c_col)
    Else
        If c_col > UBound(aColsWidth) Then
            ReDim Preserve aColsWidth(Max(UBound(aColsWidth), c_col))
        End If
    End If
    lWidth = aColsWidth(c_col)
    aColsWidth(c_col) = Max(lWidth, c_width)

End Property

Private Property Get DefaultLogFileName() As String
    DefaultLogFileName = fso.GetBaseName(ActiveWorkbook.Name) & ".log"
End Property

Public Sub Entry(ParamArray e() As Variant)
' ----------------------------------------------------------------------------
' - When a single string without any | is provided the string is appended as a
'   new line to the log file
' - When multiple strings are provided or a string contains vertical bars (|)
'   the indicated items are writen aligned in columns to the log-file
' ----------------------------------------------------------------------------
    Dim i       As Long
    Dim iCol    As Long
    Dim s       As String
    Dim v       As Variant
    
    v = e
    If LBound(v) = UBound(v) Then
        '~~ Only a single string has been provided
        s = v(LBound(v))
        If InStr(s, "|") <> 0 Then
            '~~ The | in the string indicates items for being column aligned
            AlignedByColumns = True
            For i = LBound(e) To UBound(e)
                iCol = i + 1
                ColItem(iCol) = e(i)
            Next i
            WriteItems
        Else
            AlignedByColumns = False
    
            If bNewLog Then
                BeginOfNewLogEntries n_items:=False, n_line:=s
                bNewLog = False
            End If
            WriteLine s
        End If
    Else
        '~~ When more than one string is provided this indicates items for being column aligned
        AlignedByColumns = True
        For i = LBound(e) To UBound(e)
            iCol = i + 1
            ColItem(iCol) = e(i)
        Next i
        WriteItems
    End If
        
xt: Exit Sub
End Sub

Private Property Let AlignedByColumns(ByVal b As Boolean)
' ----------------------------------------------------------------------------
' The change from one mode to another (from bAlignedByColumns = True to False
' and vice versa) indicates the begin of a new series of log entries.
' ----------------------------------------------------------------------------
    If Not bAlignedByColumns = b Then
        
        bAlignedByColumns = b
        bNewLog = True
    End If
End Property

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

Public Sub Title(ParamArray t() As Variant)
' ----------------------------------------------------------------------------
' Alternatively to the "Service" property!
' ----------------------------------------------------------------------------
    bNewLog = True
    aTitle = t
    aItems = vbNullString
End Sub

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
                Align = VBA.Right$(sFill & sItem, a_lngth)
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

Private Function IsAllocated(arr As Variant) As Boolean
' -----------------------------------------------------------------------------------
' Retruns TRUE when an array (arr) is allocated, i.e. when it has a valid LBound and
' and an LBound which is less or equal its UBound.
' -----------------------------------------------------------------------------------
    On Error Resume Next
    IsAllocated = IsArray(arr) _
              And Not IsError(LBound(arr, 1)) _
              And LBound(arr, 1) <= UBound(arr, 1)
End Function

Private Function Aligned(ByVal a_string As String) As String
' -----------------------------------------------------------------------------------
' Returns the implicit alignment of a string (a_string) which is either "L"eft,
' "C"entered, "R"ight depending on the number of sColMargin(s) at the left and the
' right of the string. Any left or right column delimiter (|) is ignored.
'
' Examples:  |    Left   | Centered |   Right   |
'            |-----------+----------+-----------|
'            | "xxx "    | "xxx"    | " xxxx"   |
'            | " xxxx  " | " xxxx " | "  xxxx " |
'
' -----------------------------------------------------------------------------------
    Dim lLeft   As Long
    Dim lRight  As Long
    Dim v       As Variant
    Dim i       As Long
    Dim s       As String
    
    a_string = Replace(a_string, "|", vbNullString)
    v = Split(a_string, sColsMargin)
    For i = LBound(v) To UBound(v)
        If v(i) <> vbNullString Then Exit For
        lLeft = lLeft + 1
    Next i
    For i = UBound(v) To LBound(v) Step -1
        If v(i) <> vbNullString Then Exit For
        lRight = lRight + 1
    Next i
    Select Case True
        Case lLeft = lRight:    Aligned = "C"
        Case lLeft < lRight:    Aligned = "L"
        Case lLeft > lRight:    Aligned = "R"
    End Select

End Function

Public Sub ColsAlignmentHeader(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' This method explicitely specifies the columns alignment of the column headers.
' Only when no explicit header alignment had been specified, the alignment is considered
' "implicit", indicated by the number of left and right spaces.
' -----------------------------------------------------------------------------------
    Dim i       As Long
    Dim iCol    As Long
    
    For i = LBound(c_align) To UBound(c_align)
        iCol = i + 1
        ColAlignmentHeader(iCol) = Aligned(c_align(i))
    Next i
    
End Sub

Public Sub ColsAlignmentItem(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' When ColWidths are not provided the columns width defaults to the width of the
' colum headers plus 2 margin spaces.
' -----------------------------------------------------------------------------------
    Dim i       As Long
    Dim iCol    As Long
    
    For i = LBound(c_align) To UBound(c_align)
        iCol = i + 1
        ColAlignmentItem(iCol) = c_align(i)
    Next i
    
End Sub

Private Property Let ColItem(Optional ByVal c_col As Long, _
                                     ByVal c_item As String)
    If Not IsAllocated(aItems) Then
        ReDim aItems(c_col)
    ElseIf c_col > UBound(aItems) Then
        ReDim Preserve aItems(c_col)
    End If
    aItems(c_col) = c_item

End Property

Private Property Get ColItem(Optional ByVal c_col As Long) As String
' -----------------------------------------------------------------------------------
' Returns the item of a column (c_col):
' - Aligned as explicitely specified or, when not specified explicitely, aligned as
'   indicated implicit by the very first.
' - In the width explicitely specified and/or specified by the header's width or when
'   no width is specified by the width of the very first row's item width.
' Note 1: Any column(c_col) above the UBound(aItems) is ignored, i.e. returns a
'         vbNullString.
' -----------------------------------------------------------------------------------
    Dim lWidth As Long
    
    If IsAllocated(aItems) Then
        Select Case True
            Case c_col < UBound(aItems)
                If ColWidth(c_col) = 0 Then
                    '~~ When yet no column widht is specified, the first one specifies it
                    lWidth = Len(aItems(c_col))
                    ColWidth(c_col) = lWidth
                Else
                    lWidth = ColWidth(c_col)
                End If
            Case c_col = UBound(aItems)
                '~~ The rightmost column's width may be expanded to the lenght of the item
                lWidth = ColWidth(c_col)
                lWidth = Max(lWidth, Len(sColsMargin & aItems(c_col)))
                ColWidth(c_col) = lWidth
        End Select
        ColItem = Align(aItems(c_col), lWidth, ColAlignmentItem(c_col), sColsMargin)
    End If
End Property

Public Sub Headers(ParamArray c_headers() As Variant)
' -----------------------------------------------------------------------------------
' Note: ColsWidth defaults to the maximum of an already specified width and the width
'       if the corresponding header string.
' -----------------------------------------------------------------------------------
    Const PROC = "Headers"
    
    On Error GoTo eh
    Dim i       As Long
    Dim j       As Long
    Dim iCol    As Long
    Dim s       As String
    Dim v       As Variant
    
    aColsHeader = vbNullString
    For i = LBound(c_headers) To UBound(c_headers)
        s = c_headers(i)
        If InStr(s, sColsDelimiter) <> 0 Then
            v = Split(s, sColsDelimiter)
            For j = LBound(v) To UBound(v)
                If v(j) <> vbNullString Then
                    iCol = j + 1
                    ColHeader(iCol) = v(j)
                    ColWidth(iCol) = Len(v(j)) ' enclosed in || the width is implicit
                End If
            Next j
        Else
            iCol = i + 1
            ColHeader(iCol) = c_headers(i)
            ColWidth(iCol) = Len(sColsMargin & Trim(c_headers(i)) & sColsMargin)
            
            If ColAlignmentHeader(iCol) = vbNullString Then
                '~~ Implicit alignment when none has been specified explicitely
                ColAlignmentHeader(iCol) = Aligned(c_headers(i))
            End If
        End If
    Next i
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub ColsWidth(ParamArray c_widths() As Variant)
' -----------------------------------------------------------------------------------
' Specifies the width of n columns which is the maximum of an already specified and
' this one specified width.
' -----------------------------------------------------------------------------------
    Const PROC = "ColsWidth"
    
    On Error GoTo eh
    Dim i       As Long
    Dim iCol    As Long
    
    For i = LBound(c_widths) To UBound(c_widths)
        iCol = i + 1
        ColWidth(iCol) = c_widths(i)
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
        If IsAllocated(aColsHeader) _
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
    Dim sElement    As String
    
    For i = LBound(aItems) To UBound(aItems)
        s = s & sColDelim & ColItem(i)
        sColDelim = sColsDelimiter
    Next i
    
    If bNewLog Then
        BeginOfNewLogEntries True, s
        bNewLog = False
    End If
    
    WriteLine s
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub WriteLine(ByVal w_string As String)
' ----------------------------------------------------------------------------
' Writes a string (w_string) into the file (sFileFullName).
' ----------------------------------------------------------------------------
    Const PROC = "WriteLine"
    
    On Error GoTo eh
    Dim ts  As TextStream
   
    ProvideLogFile
    Set ts = fso.OpenTextFile(FileName:=sFileFullName, IOMode:=ForAppending)
    ts.WriteLine TimeStamp & w_string

xt: ts.Close
    Set ts = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub BeginOfNewLogEntries(ByVal n_items As Boolean, _
                                 ByVal n_line As String)
' -----------------------------------------------------------------------------------
' Implicitely performed with the writing of the first line of a new series of log
' entries provided:
' - A new instance of the Class-Module had been established
' - The Title Property has been asigned a (new) title
' - The _ColsHeader_ method had been used to specify new column headers
' - A new log-file has been stablished
'
' Writes: - A delimiting line (=====) when the log-file already has entries, in the
'           width of the about to be written line (n_line)
'         - The specified Title(s) if any, , in the width of the about to be written
'           line (n_line)
'         - A column header, provided the about to be writen line (n_line) is for
'           items aligned in columns (n_items = True)
' -----------------------------------------------------------------------------------
    Const PROC = "BeginOfNewLogEntries"
    
    On Error GoTo eh
    Dim s           As String
    Dim sDelimiter  As String
    Dim i           As Long
    Dim sColDelim   As String
    Dim v           As Variant
    Dim lLenTitle   As Long
    
    ProvideLogFile
    If fso.GetFile(sFileFullName).Size <> 0 Then
        '~~ Log entries delimiter
        WriteLine String(Len(n_line), "=")
    End If
    
    If IsAllocated(aTitle) Then
        '~~ Log title
        For i = LBound(aTitle) To UBound(aTitle)
            s = Align(aTitle(i), Len(n_line), "C", " ", "-")
            WriteLine s
        Next i
        WriteLine String(Len(s), "-") ' Title undeline
    End If
      
    If IsAllocated(aColsHeader) Then
        '~~ Column header
        s = vbNullString
        For i = LBound(aColsHeader) To UBound(aColsHeader)
            s = s & sColDelim & ColHeader(i)
            sColDelim = sColsDelimiter
        Next i
        WriteLine s
    
        '~~ Column header under-line based on the header line written
        v = Split(s, sColsDelimiter)
        s = vbNullString
        sColDelim = vbNullString
        For i = LBound(v) To UBound(v)
            s = s & sColDelim & String(Len(v(i)), "-")
            sColDelim = "+"
        Next i
        WriteLine s
    End If
    
    bNewLog = False ' may be set to True again when a new Title is provided
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub NewLog()
    bNewLog = True
End Sub
