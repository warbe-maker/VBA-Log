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
' Class Module clsLog: Common services for writing a log file including writing
' ==================== entries aligned in columns.
'
' Methods/services:
' -----------------
' - AlignmentHeaders Explicitly specifies the alignment for each column's
'                    header, defaults to "C"entered
' - AlignmentItems   Explicitly specifies the alignment for each column's
'                    item, defaults to "L"eft adjusted when not provided.
' - ColsDelimiter    Defaults to | when Headers are specified and defaults to a
'                    single space when no Headers are specified.
' - Dsply            Displays the Log-Files content by means of ShellRun.
' - Entry            Either writes a single string, i.e. one without any any |
'                    and or subsequent string elements of writes an number of
'                    provided strings or one string with vertical bars, both
'                    indicating items to be writen alingned in columns.
' - Headers          Specifies the column headers
' - MinColWidths     Explicitly specifies each columns minimum width. The width of, defaults to the
'                    width of the corresponding column header string, is
'                    adjusted to the width of the corresponding header string
'                    when one is provided.
' - NewLog           Explicit indication that the next Entry is the first of a
'                    new series of log entries. I.e. with the next Entry a
'                    Delimiter line is written (provided its not a new log file).
'                    In most cases this method is unnecessary because the begin
'                    of a new series of log entries is implicitly considered
'                    with the Title method, the Headers method and in case the
'                    Entry changes from a single string to column aligned items
'                    or vice versa.
' - Title            Specifies the - optionally multiline - title of a new series
'                    of log entries. Triggers the writing of the column headers
'                    provided specified.
'
' Properties:
' -----------
' - FileFullName  r/w     Specifies the full name of the log-file. The location
'                         defaults to the ActiveWorkbook's parent folder, the name
'                         defaults to the ActiveWorkbook's BaseName with a .log
'                         file extention.
' - FileName      w       Specifies the log-file's name, defaults to the
'                         ActiveWorkbook's BaseName with a .log file extention.
' - KeepDays      w       Specifies the number of days a new created log-file is
'                         kept before it is deleted and re-created, defaults to 1.
' - LogFile       r       The log-file as file object
' - Path          w       Sepcifies the log-file's location/folder, defaults to the
'                         ActiveWorkbook's parent folder
' - WithTimeStamp w       Prefixes each log entry with a timestamp in the format
'                         yy-mm-dd-hh:mm:ss when True
'
' W. Rauschenberger, Berlin Apr 2023
' -----------------------------------------------------------------------------------
Private aAlignmentsHeader   As Variant ' Public for test purpose only
Private aAlignmentsItem     As Variant ' Public for test purpose only
Private aHeaders            As Variant ' Public for test purpose only
Private aItems              As Variant
Private aTitle              As Variant
Private aMinColWidths             As Variant ' Public for test purpose only
Private aFills()            As String
Private bColsAligned        As Boolean
Private bNewLog             As Boolean
Private bWithTimeStamp      As Boolean
Private fso                 As New FileSystemObject
Private lColumns            As Long
Private lKeepDays           As Long
Private lLenTitle           As Long
Private sAlignmentTitle     As String
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
    Init
    bWithTimeStamp = False
    sFileFullName = ActiveWorkbook.Path & "\" & DefaultLogFileName
    sFileName = DefaultLogFileName
    sPath = ActiveWorkbook.Path
End Sub

Private Sub Class_Terminate()
    Set fso = Nothing
    Set aAlignmentsHeader = Nothing
    Set aAlignmentsItem = Nothing
    Set aHeaders = Nothing
    Set aMinColWidths = Nothing
    Set aItems = Nothing
    Set aTitle = Nothing
    sColsDelimiter = vbNullString
End Sub

Private Property Let AlignedByColumns(ByVal b As Boolean)
' ----------------------------------------------------------------------------
' The change from one mode to another (from bColsAligned = True to False
' and vice versa) indicates the begin of a new series of log entries.
' ----------------------------------------------------------------------------
    If Not bColsAligned = b Then
        bColsAligned = b
        If Not bNewLog Then Init
    End If
End Property

Private Property Get ColAlignmentHeader(Optional ByVal c_col As Long) As String
    If IsAllocated(aAlignmentsHeader) Then
        If c_col <= UBound(aAlignmentsHeader) Then
            ColAlignmentHeader = aAlignmentsHeader(c_col)
        End If
    End If
End Property

Private Property Let ColAlignmentHeader(Optional ByVal c_col As Long, _
                                                ByVal c_align As String)
    If Not IsAllocated(aAlignmentsHeader) Then
        ReDim aAlignmentsHeader(c_col)
    ElseIf c_col > UBound(aAlignmentsHeader) Then
        ReDim Preserve aAlignmentsHeader(c_col)
    End If
    aAlignmentsHeader(c_col) = c_align
    
    Debug.Print "Head-Col " & c_col & ": Alignment=" & c_align

End Property

Private Property Get ColAlignmentItem(Optional ByVal c_col As Long) As String
    If IsAllocated(aAlignmentsItem) Then
        If c_col <= UBound(aAlignmentsItem) _
        Then ColAlignmentItem = aAlignmentsItem(c_col)
    End If
End Property

Private Property Get ColFill(Optional ByVal c_col As Long) As String
    If IsAllocated(aFills) Then
        If c_col <= UBound(aFills) _
        Then ColFill = aFills(c_col)
    End If
End Property

Private Property Let ColFill(Optional ByVal c_col As Long, _
                                      ByVal c_fill As String)
    If Not IsAllocated(aFills) Then
        ReDim aFills(c_col)
    ElseIf c_col > UBound(aFills) Then
        ReDim Preserve aFills(c_col)
    End If
    aFills(c_col) = c_fill
    
End Property

Private Property Let ColAlignmentItem(Optional ByVal c_col As Long, _
                                               ByVal c_align As String)
    Dim lAlignmentLine As Long
    
    If Not IsAllocated(aAlignmentsItem) Then
        ReDim aAlignmentsItem(c_col)
    ElseIf c_col > UBound(aAlignmentsItem) Then
        ReDim Preserve aAlignmentsItem(c_col)
    End If
    If aAlignmentsItem(c_col) = vbNullString _
    Then aAlignmentsItem(c_col) = c_align
    
    Debug.Print "Item-Col " & c_col & ": Alignment=" & c_align
End Property

Private Property Get ColHeader(Optional ByVal c_col As Long) As String
    Dim s   As String
    Dim l   As Long
    
    If IsAllocated(aHeaders) Then
        If c_col <= UBound(aHeaders) Then
            s = WithMargins(aHeaders(c_col), sColsMargin)
            l = Len(s)
            l = Max(ColWidth(c_col), l) ' width with marging may be greater than originally calculated
            ColWidth(c_col) = l
            ColHeader = Align(aHeaders(c_col), l, ColAlignmentHeader(c_col), sColsMargin)
        End If
    End If
End Property

Private Property Let ColHeader(Optional ByVal c_col As Long, _
                                        ByVal c_header As String)
    
    If Not IsAllocated(aHeaders) Then
        ReDim aHeaders(c_col)
    Else
        ReDim Preserve aHeaders(Max(UBound(aHeaders), c_col))
    End If
    aHeaders(c_col) = c_header
    
    
End Property

Private Property Get ColItem(Optional ByVal c_col As Long) As String
' -----------------------------------------------------------------------------------
' Returns the item of a column (c_col):
' - AlignedImplicitly as explicitly specified or, when not specified explicitly, aligned as
'   indicated implicit by the very first.
' - In the width explicitly specified and/or specified by the header's width or when
'   no width is specified by the width of the very first row's item width.
' Note 1: Any column(c_col) beyond UBound(aItems) is ignored, i.e. returns a
'         vbNullString.
' -----------------------------------------------------------------------------------
    Dim lWidth  As Long
    Dim sFill   As String
    
    If IsAllocated(aItems) Then
        If bNewLog Then
            '~~ When this is the first line of a new series of log entries the width
            '~~ may still be expanded when the item's wdth is greater than the
            '~~ already specified - in case
            lWidth = ColWidth(c_col)
            ColWidth(c_col) = Max(Len(aItems(c_col)), lWidth)
        End If
        sFill = ColFill(c_col)
        If sFill = vbNullString Then sFill = " "
        If c_col < lColumns Then
            ColItem = Align(aItems(c_col), ColWidth(c_col), ColAlignmentItem(c_col), sColsMargin, sFill)
        ElseIf c_col = lColumns Then
            '~~ The width of the rightmost column is unlimited. A specified header however
            '~~ will consider the columns width which includes the width of the first lines's item.
            ColItem = Align(aItems(c_col), Len(aItems(c_col)) + Len(sColsMargin), "L", sColsMargin, sFill)
        End If
    End If
    
End Property

Private Property Let ColItem(Optional ByVal c_col As Long, _
                                      ByVal c_item As String)
    If Not IsAllocated(aItems) Then
        ReDim aItems(c_col)
    ElseIf c_col > UBound(aItems) Then
        ReDim Preserve aItems(c_col)
    End If
    aItems(c_col) = c_item

End Property

Public Property Let ColsDelimiter(ByVal s As String)
    sColsDelimiter = s
End Property

Private Property Get ColWidth(Optional ByVal c_col As Long) As Long
    If IsAllocated(aMinColWidths) Then
        If c_col <= UBound(aMinColWidths) Then
            ColWidth = aMinColWidths(c_col)
        End If
    End If
End Property

Private Property Let ColWidth(Optional ByVal c_col As Long, _
                                       ByVal c_width As Long)
    Dim lWidth As Long
    
    If Not IsAllocated(aMinColWidths) Then
        ReDim aMinColWidths(c_col)
    ElseIf c_col > UBound(aMinColWidths) Then
        ReDim Preserve aMinColWidths(c_col)
    End If
    aMinColWidths(c_col) = c_width
    Debug.Print "WidthCol " & c_col & ": Width    =" & c_width

End Property

Private Property Get DefaultLogFileName() As String
    DefaultLogFileName = fso.GetBaseName(ActiveWorkbook.Name) & ".log"
End Property

Public Property Get FileFullName() As String:   FileFullName = sFileFullName:   End Property

Public Property Let FileFullName(ByVal s As String)
' ----------------------------------------------------------------------------
' Explicitly specifies the log file's name and location. This is an
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

Public Property Let WithTimeStamp(ByVal b As Boolean)
    bWithTimeStamp = b
End Property

Public Function Align(ByVal a_strng As String, _
                      ByVal a_lngth As Long, _
             Optional ByVal a_mode As String = "L", _
             Optional ByVal a_margin As String = vbNullString, _
             Optional ByVal a_fill As String = " ") As String
' ----------------------------------------------------------------------------
' Returns a string (a_strng) with a lenght (a_lngth) aligned (a_mode) filled
' with characters (a_fill).
' Note 1: The length (a_lngth) of the string (a_strng) is considered including
'         already the length of the column margin (a_margin) - if any.
' Note 2: A fill (a_fill) = "." will fill with ".....:"
' ----------------------------------------------------------------------------
    Dim sFill   As String
    Dim s       As String
    
    Select Case Left(UCase(a_mode), 1)
        Case "L"
            If a_fill = " " Then
                Align = Left$(a_margin & LTrim$(a_strng) & a_margin & String(a_lngth, a_fill), a_lngth)
            ElseIf a_fill = "." Then
                Align = Left$(a_margin & LTrim$(a_strng) & " " & String(a_lngth, a_fill), a_lngth) & ":"
            End If
        Case "R"
            Align = Right$(String(a_lngth, a_fill) & a_margin & RTrim$(a_strng) & a_margin, a_lngth)
        Case "C"
            s = a_margin & Trim$(a_strng) & a_margin
            sFill = String$(Int((a_lngth - Len(s)) / 2), a_fill)
            Align = sFill & s & sFill
            If Len(Align) < a_lngth Then
                Align = Align & VBA.String$(a_lngth - Len(Align), a_fill)
            End If
    End Select

End Function

Public Function AlignedImplicitly(ByVal a_string As String) As String
' -----------------------------------------------------------------------------------
' Function returns the alignment specified implicitly in accordance with the
' following rules: - (L)eft adjusted  = Leading spaces > trailing spaces
'                  - (C)entered       = Leading spaces = trailing spaces
'                  - (R)ight adjusted = Leading spaces < trailing spaces.
'
' Examples:  |   (L)eft  |(C)entered|  (R)ight  |
'            |-----------+----------+-----------|
'            | "xxx "    | "xxx"    | " xxxx"   |
'            | " xxxx  " | " xxxx " | "  xxxx " |
'
' -----------------------------------------------------------------------------------
    Dim lLeft   As Long
    Dim lRight  As Long
    Dim i       As Long
    
    For i = 1 To Len(a_string)
        If i = 1 Then
            If Left(a_string, 1) <> " " Then Exit For
        Else
            If Mid$(a_string, i, 1) <> " " Then Exit For
        End If
        lLeft = lLeft + 1
    Next i
    For i = Len(a_string) To 1 Step -1
        If i = Len(a_string) Then
            If Right(a_string, 1) <> " " Then Exit For
        Else
            If Mid$(a_string, i, 1) <> " " Then Exit For
        End If
        lRight = lRight + 1
    Next i
    Select Case True
        Case lLeft = lRight:    AlignedImplicitly = "C"
        Case lLeft < lRight:    AlignedImplicitly = "L"
        Case lLeft > lRight:    AlignedImplicitly = "R"
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

Public Function AsArray(ByVal a_string As String, _
                        ByVal a_split_by As String) As String()
' -----------------------------------------------------------------------------------
' Returns a string (a_string) as an array of strings considering a split character
' (a_split_by). Any leading or trailing split character is removed before the split.
' -----------------------------------------------------------------------------------
    If InStr(a_string, a_split_by) <> 0 Then
        If Left(a_string, 1) = a_split_by Then a_string = Right(a_string, Len(a_string) - 1)
        If Right(a_string, 1) = a_split_by Then a_string = Left(a_string, Len(a_string) - 1)
    End If
    AsArray = Split(a_string, a_split_by)
End Function

Public Sub Dsply()
' -----------------------------------------------------------------------------------
' Displays the Log-Files content.
' -----------------------------------------------------------------------------------
    ShellRun sFileFullName
End Sub

Public Sub Entry(ParamArray e() As Variant)
' ----------------------------------------------------------------------------
' - When a single string without any | is provided the string is appended as a
'   new line to the log file
' - When multiple strings are provided or a string contains vertical bars (|)
'   the indicated items are writen aligned in columns to the log-file
' ----------------------------------------------------------------------------
    Const PROC = "Entry"
    
    On Error GoTo eh
    Dim i       As Long
    Dim iCol    As Long
    Dim s       As String
    Dim v       As Variant
    Dim lWidth  As Long
    
    If LBound(e) = UBound(e) Then
        '~~ Only a single string has been provided
        s = e(LBound(e))
        If InStr(s, "|") <> 0 Then
            '~~ The | in the string indicates items for being column aligned
            '~~ and additionally the implicit width
            If Left(s, 1) = "|" Then s = Right(s, Len(s) - 1)
            If Right(s, 1) = "|" Then s = Left(s, Len(s) - 1)
            v = Split(s, "|")
            AlignedByColumns = True
            lColumns = UBound(v)
            For i = LBound(v) To UBound(v)
                iCol = i + 1
                ColItem(iCol) = CStr(v(i))
                If ColWidth(iCol) = 0 Then
                    ColWidth(iCol) = Len(v(i))
                End If
            Next i
            WriteItems
        Else
            AlignedByColumns = False
            If bNewLog _
            Then NewLog n_items:=False, n_line:=s
            WriteLine s
            bNewLog = False
        End If
    Else
        '~~ When more than one string is provided this indicates items for being aligned in columns
        AlignedByColumns = True
        lColumns = UBound(e) + 1
        For i = LBound(e) To UBound(e)
            iCol = i + 1
            If bNewLog Then
                '~~ When yet no alignment had been specified explicitly
                '~~ the first row determines the items alignment
                If ColAlignmentItem(iCol) = vbNullString _
                Then ColAlignmentItem(iCol) = AlignedImplicitly(e(i))
            End If
            ColItem(iCol) = sColsMargin & Trim$(CStr(e(i))) & sColsMargin
            
            If ColWidth(iCol) = 0 Then
                ColWidth(iCol) = Len(ColItem(iCol))
            End If
        Next i
        WriteItems
    End If
        
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AlignmentItems(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' Explicit specification of the items column alignment. When no explicit alignment is
' is specified or the first character is neither L, nor C, nor R, the alignment
' default to the implicit specification of the first line's items.
' Valid specifications: Spec        Meaning
'                       ----------- -------------------------------
'                       L           Left adjusted no fill
'                       L.          Left adjusted filled with ....:
'                       R           Right adjusted
'                       C           Centered
'                       -C-, C-, -C Centered with - fills
'
' -----------------------------------------------------------------------------------
    Const PROC = "AlignmentItems"
    Dim i       As Long
    Dim iCol    As Long
    Dim s       As String
    Dim sAlign  As String
    Dim sFill   As String
    
    On Error GoTo eh
    For i = LBound(c_align) To UBound(c_align)
        If Not IsError(c_align(i)) Then
            '~~ IsError indicates a columns alignment 'not explicitly specified'.
            '~~ The corresponding columns will be aligned implicitly specified
            '~~ with the first entry's items specifications
            iCol = i + 1
            sFill = " " ' the default in case nothing else is specified
            Select Case True
                Case Left$(c_align(i), 1) = "-", Right$(c_align(i), 1) = "-"
                    sAlign = UCase(Left(Replace$(c_align(i), "-", vbNullString), 1))
                    sFill = "-"
                Case Right$(c_align(i), 1) = "."
                    sAlign = UCase(Left$(c_align(i), 1))
                    sFill = "."
            End Select
            Select Case True
                Case sAlign = "L" And sFill = vbNullString
                Case sAlign = "L" And sFill = "."
                Case sAlign = "R"
                Case sAlign = "C" And sFill = vbNullString
                Case sAlign = "C" And sFill = "-"
                Case Else: Err.Raise AppErr(1), ErrSrc(PROC), "The specified alignment is invalid! Valid specs are:" & vbLf & _
                                                              "L, L., R, C, -C, C-, -C-"
            End Select
            ColAlignmentItem(iCol) = sAlign
            ColFill(iCol) = sFill
        End If
    Next i
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
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

Public Sub Headers(ParamArray c_headers() As Variant)
' -----------------------------------------------------------------------------------
' Spacifies the column headers for Entry items aligne in columns.
' -----------------------------------------------------------------------------------
    Const PROC = "Headers"
    
    On Error GoTo eh
    Dim i       As Long
    Dim j       As Long
    Dim iCol    As Long
    Dim s       As String
    Dim v       As Variant
    Dim lWidth  As Long
    
    aHeaders = vbNullString
    For i = LBound(c_headers) To UBound(c_headers)
        s = c_headers(i)
        If InStr(s, "|") <> 0 Then
            v = AsArray(s, "|")
            For j = LBound(v) To UBound(v)
                If v(j) <> vbNullString Then
                    iCol = j + 1
                    If ColAlignmentHeader(iCol) = vbNullString Then
                        '~~ The first specification of a column's header triggers the
                        '~~ implicit alignment specification - provided none has yet been specified explicitly
                        ColAlignmentHeader(iCol) = AlignedImplicitly(v(j))
                    End If
                    s = v(j)
                    If Len(s) < Len(Trim(s)) + (2 * Len(sColsMargin)) Then
                        '~~ The minimum width of a column header is the mere characters
                        '~~ string enclosed in a leading and trailing sColsMargin
                        s = WithMargins(s, sColsMargin)
                    End If
                    ColHeader(iCol) = Trim(s)
                    lWidth = ColWidth(iCol)
                    ColWidth(iCol) = Max(Len(s), lWidth)
                End If
            Next j
        Else
            iCol = i + 1
            If ColAlignmentHeader(iCol) = vbNullString Then
                '~~ Implicit alignment when none has been specified explicitly
                ColAlignmentHeader(iCol) = AlignedImplicitly(c_headers(i))
            End If
            ColHeader(iCol) = Trim(c_headers(i))
            lWidth = ColWidth(iCol)
            ColWidth(iCol) = Max(Len(WithMargins(c_headers(i), sColsMargin)), lWidth)
        End If
    Next i
    If IsAllocated(aHeaders) Then
        Select Case sColsDelimiter
            Case vbNullString:      sColsDelimiter = "|":   sColsMargin = " "
            Case "|":                                       sColsMargin = " "
            Case " ":                                       sColsMargin = vbNullString
        End Select
    End If
    
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub AlignmentHeaders(ParamArray c_align() As Variant)
' -----------------------------------------------------------------------------------
' This method explicitly specifies the columns alignment of the column headers.
' In case no explicit headers alignment is specified (or the specification's first
' character is neither L, nor C, nor R, the alignment is considered "implicit",
' indicated by the number of left and right spaces.
' -----------------------------------------------------------------------------------
    Dim i       As Long
    Dim iCol    As Long
    Dim s       As String
    
    For i = LBound(c_align) To UBound(c_align)
        iCol = i + 1
        s = UCase(Left(c_align(i), 1))
        Select Case s
            Case "L", "C", "R": ColAlignmentHeader(iCol) = s
        End Select
    Next i
    
End Sub

Private Sub Init()
    aAlignmentsHeader = vbNullString
    aAlignmentsItem = vbNullString
    aHeaders = vbNullString
    aItems = vbNullString
    aTitle = vbNullString
    bNewLog = True
    lKeepDays = 1
    sColsDelimiter = " "        ' Default when no Headers were specified
    sColsMargin = vbNullString  ' Default when no Headers were specified
End Sub

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

Public Sub NewLog(Optional ByVal n_items As Boolean = False, _
                  Optional ByVal n_line As String = vbNullString)
' -----------------------------------------------------------------------------------
' When not provided with any arguments this is an explicit indication that the next
' Entry is the first of a new series of log entries. I.e. with the next Entry a
' delimiter line (=====) is written - provided its not a new log file. In most cases
' there is no need to call this method explicitly because is is called implicitly
' with the next Entry method when a Title or Headers had been specified, or in case
' the Entry changes from a single string to column aligned items or vice versa.
' When called implicitly (n_line not a vbNullString) the method writes:
' - A delimiting line (=====) when the log-file already has entries, in the width of
'   the about to be written line (n_line)
' - The specified Title(s) if any, , in the width of the about to be written line
'   (n_line)
' - A column header, provided the about to be writen line (n_line) is for items
'   aligned in columns (n_items = True)
' -----------------------------------------------------------------------------------
    Const PROC = "NewLog"
    
    On Error GoTo eh
    Dim s                   As String
    Dim sDelimiter          As String
    Dim i                   As Long
    Dim sColDelim           As String
    Dim v                   As Variant
    Dim sHeader             As String
    Dim sHeaderUnderLine    As String
    Dim sLogDelim           As String
    Dim sTitle              As String
    Dim lWidthEntry         As Long
    Dim lLenHeader          As Long
    Dim sTitleFill          As String
    Dim sUnderLineDelim     As String
    
    If n_line = vbNullString Then
        '~~ Explicit call only indicates that the next Entry will write the first of
        '~~ a series of new log entries
        Init
        GoTo xt
    End If
    
    ProvideLogFile
    
    '~~ Assemble column header
    sHeader = vbNullString
    If IsAllocated(aHeaders) Then
        For i = LBound(aHeaders) To UBound(aHeaders)
            sHeader = sHeader & sColsDelimiter & ColHeader(i)
            lLenHeader = Max(lLenHeader, Len(sHeader))
        Next i
    End If
            
    lLenTitle = Max(lLenTitle, Len(n_line))
    lWidthEntry = Max(lLenTitle, lLenHeader, Len(n_line))
    '~~ New log delimiter line
    If fso.GetFile(sFileFullName).Size <> 0 Then
        '~~ Log entries delimiter
        WriteLine String(lWidthEntry, "=")
    End If
    
    If IsAllocated(aTitle) Then
        '~~ Write (optionally multiline) title
        If sAlignmentTitle = "C" Then sTitleFill = "-" Else sTitleFill = " "
        For i = LBound(aTitle) To UBound(aTitle)
            If aTitle(i) <> vbNullString Then
                Select Case True
                    Case lWidthEntry <= lLenTitle
                        sTitle = Align(aTitle(i), lWidthEntry, sAlignmentTitle, vbNullString, sTitleFill)
                    Case lWidthEntry > lLenTitle And sAlignmentTitle = "C"
                        sTitle = Align(aTitle(i), lWidthEntry, sAlignmentTitle, " ", "-")
                    Case lWidthEntry > lLenTitle And sAlignmentTitle <> "C"
                        sTitle = Align(aTitle(i), lWidthEntry, sAlignmentTitle, " ", " ")
                End Select
                WriteLine sTitle
            End If
        Next i
        '~~ Write title underline
        WriteLine String(lWidthEntry, "-")
    End If
      
    If sHeader <> vbNullString Then
        WriteLine sHeader
        '~~ Column header under-line based on the header line written
        v = Split(sHeader, sColsDelimiter)
        sHeaderUnderLine = vbNullString
        sColDelim = vbNullString
        If sColsDelimiter = "|" _
        Then sUnderLineDelim = "+" _
        Else sUnderLineDelim = " "
        
        For i = LBound(aHeaders) To UBound(aHeaders)
            sHeaderUnderLine = sHeaderUnderLine & sUnderLineDelim & String(Len(ColHeader(i)), "-")
        Next i
        If Len(sHeaderUnderLine) < lWidthEntry Then
            sHeaderUnderLine = sHeaderUnderLine & String(lWidthEntry - Len(sHeaderUnderLine), "-")
        End If
        WriteLine sHeaderUnderLine
    End If
        
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub ProvideLogFile()
    Dim lDaysAge As Long
    Dim ts       As TextStream
    With fso
        If Not .FileExists(sFileFullName) Then
            Set ts = .CreateTextFile(sFileFullName)
            ts.Close
            Debug.Print .GetFile(sFileFullName).DateLastAccessed
        Else
            lDaysAge = VBA.DateDiff("d", .GetFile(sFileFullName).DateLastAccessed, Now())
            If lDaysAge > lKeepDays Then
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
    TimeStamp = " "
    If bWithTimeStamp Then TimeStamp = Format(Now(), "yy-mm-dd-hh:mm:ss") & TimeStamp
End Function

Public Sub Title(ParamArray t() As Variant)
' ----------------------------------------------------------------------------
' Method for specifying an - optionally multi-lined - title for an
' (implicitly specified) new series of log entries.
' Note: The first title line implicitly specifies the title's alignment.
' ----------------------------------------------------------------------------
    Dim i       As Long
    Dim s       As String
    Dim iLine   As Long
    
    sAlignmentTitle = "C" ' default
    If Not bNewLog Then Init
    ReDim aTitle(UBound(t) + 1)
    lLenTitle = 0
    For i = LBound(t) To UBound(t)
        s = t(i)
        If i = LBound(t) Then sAlignmentTitle = AlignedImplicitly(s)
        iLine = i + 1
        Select Case sAlignmentTitle
            Case "L": s = RTrim$(s)
            Case "C": s = Trim$(s)
            Case "R": s = LTrim$(s)
        End Select
        aTitle(iLine) = s
        lLenTitle = Max(lLenTitle, Len(s))
    Next i
    aItems = vbNullString

End Sub

Public Sub MinColWidths(ParamArray c_widths() As Variant)
' -----------------------------------------------------------------------------------
' Specifies the minimum width of n columns. Each columns specified width may be
' expanded by a) an implicite header width specification, b) the width of the first
' _Entry_, i.e. finally it will be the maximum of the three.
' -----------------------------------------------------------------------------------
    Const PROC = "MinColWidths"
    
    On Error GoTo eh
    Dim i       As Long
    Dim iCol    As Long
    
    For i = LBound(c_widths) To UBound(c_widths)
        If Not IsError(c_widths(i)) Then
            If Not IsError(Abs(CInt(c_widths(i)))) Then
                iCol = i + 1
                ColWidth(iCol) = c_widths(i)
            End If
        End If
    Next i

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function WithMargins(ByVal w_string As String, _
                             ByVal w_margin As String) As String
' -----------------------------------------------------------------------------------
' Returns a string (w_string) with a leading and trailing margin (w_margin) in case
' the string (w_string) not already begins and/or ends with a margin (w_margin). A
' margin = vbNullString is ignored.
' -----------------------------------------------------------------------------------
    Dim s As String
    If w_margin <> vbNullString Then
        If Left(w_string, 1) <> w_margin _
        Then s = w_margin & w_string _
        Else s = w_string
        If Right(s, 1) <> w_margin _
        Then s = s & w_margin
        WithMargins = s
    Else
        WithMargins = w_string
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
    Dim sElement    As String
        
    For i = LBound(aItems) To UBound(aItems)
        s = s & sColsDelimiter & ColItem(i)
    Next i
    
    If bNewLog _
    Then NewLog True, s
    
    WriteLine s
    bNewLog = False ' may be set to True again when a new Title is provided

    
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

