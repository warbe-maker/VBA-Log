Attribute VB_Name = "mLogTest"
Option Explicit
Option Base 1
' ----------------------------------------------------------------------
' Standard Module mLogTest: Regression-Test for the clsLog Class Module
' =========================
'
' ----------------------------------------------------------------------
Private fso As New FileSystemObject

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

Private Property Get FileArry(Optional ByVal fa_file As String, _
                              Optional ByVal fa_excl_empty_lines As Boolean = False, _
                              Optional ByRef fa_split As String, _
                              Optional ByVal fa_append As Boolean = False) As Variant
' ----------------------------------------------------------------------------
' Returns the content of the file (fa_file) - a files full name - as array,
' with the used line break string returned in (fa_split).
' ----------------------------------------------------------------------------
    Const PROC  As String = "FileArry"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim a       As Variant
    Dim a1()    As String
    Dim sSplit  As String
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    Dim i       As Long
    Dim j       As Long
    Dim v       As Variant
    
    If Not fso.FileExists(fa_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "A file named '" & fa_file & "' does not exist!"
    
    '~~ Unload file to a string
    sFile = FileTxt(ft_file:=fa_file _
                    , ft_split:=sSplit _
                     )
    If sFile = vbNullString Then GoTo xt
    a = Split(sFile, sSplit)
    
    If Not fa_excl_empty_lines Then
        a1 = a
    Else
        '~~ Extract non-empty items
        For i = LBound(a) To UBound(a)
            If Len(Trim$(a(i))) <> 0 Then cll.Add a(i)
        Next i
        ReDim a1(cll.Count - 1)
        j = 0
        For Each v In cll
            a1(j) = v:  j = j + 1
        Next v
    End If
    
xt: FileArry = a1
    fa_split = sSplit
    Set cll = Nothing
    Set fso = Nothing
    Exit Property
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get FileTxt(Optional ByVal ft_file As Variant, _
                             Optional ByVal ft_append As Boolean = True, _
                             Optional ByRef ft_split As String) As String
' ----------------------------------------------------------------------------
' Returns the text file's (ft_file) content as string with VBA.Split() string
' in (ft_split). When the file doesn't exist a vbNullString is returned.
' Note: ft_append is not used but specified to comply with Property Let.
' ----------------------------------------------------------------------------
    Const PROC = "FileTxt-Get"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim ts      As TextStream
    Dim s       As String
    Dim sFl As String
   
    ft_split = ft_split  ' not used! for declaration compliance and dead code check only
    ft_append = ft_append ' not used! for declaration compliance and dead code check only
    
    With fso
        If TypeName(ft_file) = "File" Then
            sFl = ft_file.Path
        Else
            '~~ ft_file is regarded a file's full name, created if not existing
            sFl = ft_file
            If Not .FileExists(sFl) Then GoTo xt
        End If
        Set ts = .OpenTextFile(FileName:=sFl, IOMode:=ForReading)
    End With
    
    If Not ts.AtEndOfStream Then
        s = ts.ReadAll
        ft_split = SplitStr(s)
        If VBA.Right$(s, 2) = vbCrLf Then
            s = VBA.Left$(s, Len(s) - 2)
        End If
    Else
        FileTxt = vbNullString
    End If
    If FileTxt = vbCrLf Then FileTxt = vbNullString Else FileTxt = s

xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Property Get SplitStr(ByRef s As String)
' ------------------------------------------------------------------------------
' Returns the split string in string (s) used by VBA.Split() to turn the string
' into an array.
' ------------------------------------------------------------------------------
    If InStr(s, vbCrLf) <> 0 Then SplitStr = vbCrLf _
    Else If InStr(s, vbLf) <> 0 Then SplitStr = vbLf _
    Else If InStr(s, vbCr) <> 0 Then SplitStr = vbCr
End Property

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

Private Sub AssertResult(ByVal a_file As String, _
                         ByVal a_time_stamp As String, _
                         ParamArray a() As Variant)
    Dim vResult     As Variant
    Dim vExpected   As Variant
    Dim i           As Long
    
    vExpected = a
    vResult = FileArry(a_file)
    Debug.Assert UBound(vResult) = UBound(vExpected)
    For i = 0 To UBound(vResult)
        If Not vResult(i) Like "*" & vExpected(i) Then
            Debug.Print "Line  " & i + 1 & ":"
            Debug.Print "Result  : " & vResult(i)
            Debug.Print "Expected: " & vExpected(i)
            Stop
        End If
    Next i
    
End Sub

Public Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Public Sub EoP(ByVal e_proc As String, _
      Optional ByVal e_inf As String = vbNullString)
' ----------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional
' Compile Arguments are 0 or not set at all.
' ----------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
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
    ErrSrc = "mLogTest" & "." & sProc
End Function

Private Sub Test_00_Regression()
' ------------------------------------------------------------------------------
' Combines all individual test to one regression test which only stops when an
' expected result is not met. This test is obligatory after any code modifi-
' cation. When finished the test displays the Exec.trc file (written when the
' Conditional Compile Argument ExecTrace = 1.
'
' While the individual test display the result log-file when the test has one
' written, the regression test has the result assertion automated. Thus, when
' all results ar those expected, the test runs un-attended.
'
' Note: The regression test uses the Common Components:
'       - mErH  Error Handling when the Conditional Compile Argument ErHComp = 1
'       - fMsg  Error Message Display
'       - mMsg  Error Message Display
'       - mTrc  Execution trace
'
' W. Rauschenberger, Berlin May 2023
' ------------------------------------------------------------------------------
    Const PROC = "Test_00_Regression"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    mErH.Regression = True
    Test_01_Headers
    Test_02_ColsWidth
    Test_03_Property_Name
    Test_04_Property_Path_As_Workbook
    Test_05_WriteHeader
    Test_06_Log_Items

xt: EoP ErrSrc(PROC)
    mTrc.Dsply
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_01_Headers()
    Const PROC = "Test_01_Headers"
    Const HEADER_1 = "Column-01-Header"
    Const HEADER_2 = "-Column-02-Header-"
    Const HEADER_3 = "--Column-03-Header--"
    
    On Error GoTo eh
    Dim sColsMargin As String: sColsMargin = " "
    BoP ErrSrc(PROC)
    
    With New clsLog
        .ColsMargin = sColsMargin
        .Headers HEADER_1, HEADER_2, HEADER_3
        Debug.Assert .vColsWidth(1) = Len(HEADER_1): Debug.Assert .vHeaders(1) = HEADER_1
        Debug.Assert .vColsWidth(2) = Len(HEADER_2): Debug.Assert .vHeaders(2) = HEADER_2
        Debug.Assert .vColsWidth(3) = Len(HEADER_3): Debug.Assert .vHeaders(3) = HEADER_3
    End With

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_02_ColsWidth()
    Const PROC = "Test_02_ColsWidth"
    Const HEADER_1 = "Column-01-Header"
    Const HEADER_2 = "-Column-02-Header-"
    Const HEADER_3 = "--Column-03-Header--"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    With New clsLog
        .Headers HEADER_1, HEADER_2, HEADER_3
        .ColsWidth 20, 25, 30
        Debug.Assert .vColsWidth(1) = 20: Debug.Assert .vHeaders(1) = HEADER_1
        Debug.Assert .vColsWidth(2) = 25: Debug.Assert .vHeaders(2) = HEADER_2
        Debug.Assert .vColsWidth(3) = 30: Debug.Assert .vHeaders(3) = HEADER_3
    End With

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_03_Property_Name()
    Const PROC = "Test_03_Property_Name"
    
    On Error GoTo eh
    BoP ErrSrc(PROC)
    With New clsLog
        .FileName = "TestService.log"
        Debug.Assert .FileFullName = ThisWorkbook.Path & "\TestService.log"
    End With

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_04_Property_Path_As_Workbook()
    Const PROC = "Test_04_Property_Path_As_Workbook"
    
    On Error GoTo eh
    Dim fso As New FileSystemObject
    
    BoP ErrSrc(PROC)
    With New clsLog
        .Path = ActiveWorkbook.Path
        Debug.Assert .FileFullName = ActiveWorkbook.Path & "\" & fso.GetBaseName(ActiveWorkbook.Name) & ".log"
    End With

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_05_WriteHeader()
    Const PROC = "Test_05_WriteHeader"
    Const HEADER_1 = "Column-01-Header"
    Const HEADER_2 = "-Column-02-Header-"
    Const HEADER_3 = "--Column-03-Header--"
    
    Dim bTimeStamp As Boolean: bTimeStamp = True
    
    With New clsLog
        If fso.FileExists(.LogFile) Then fso.DeleteFile .LogFile
        .WithTimeStamp = bTimeStamp
        .Headers HEADER_1, HEADER_2, HEADER_3
        .ColsWidth 20, 25, 30
        .WriteHeader
        If Not mErH.Regression Then
            .Dsply
        Else
            AssertResult .LogFile _
                      , bTimeStamp _
                      , "|  Column-01-Header  |   -Column-02-Header-    |     --Column-03-Header--     " _
                      , "|--------------------+-------------------------+------------------------------"
        End If
    End With

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_06_Log_Items()
    Const PROC = "Test_06_Log_Items"
    Const HEADER_1 = "Column-01-Header"
    Const HEADER_2 = "-Column-02-Header-"
    Const HEADER_3 = "--Column-03-Header--"
    
    On Error GoTo eh
    Dim bTimeStamp As Boolean: bTimeStamp = True
    
    BoP ErrSrc(PROC)
    
    With New clsLog
        If fso.FileExists(.LogFile) Then fso.DeleteFile .LogFile
        .WithTimeStamp = bTimeStamp
        .Title = "Method 'Items' test:"
        .ColsMargin = vbNullString
        .ColsWidth 10, 25, 30
        .Headers HEADER_1, HEADER_2, HEADER_3
        .Items "xxx", "yyyyyy", "zzzzzz"
        .Items "xxx", "yyyyyy", "zzzzzz"
        .Items "xxx", "yyyyyy", "zzzzzz"
        .Title = "Method 'Items' test: New title, with marging"
        .ColsMargin = " "
        .Items "xxx", "yyyyyy", "zzzzzz"
        .Items "xxx", "yyyyyy", "zzzzzz"
        .Items "xxx", "yyyyyy", "zzzzzz"
        
        If Not mErH.Regression Then
            .Dsply
        Else
            AssertResult .LogFile _
                      , bTimeStamp _
                      , "|------------------------- Method 'Items' test: --------------------------" _
                      , "|Column-01-Header|   -Column-02-Header-    |     --Column-03-Header--     " _
                      , "|----------------+-------------------------+------------------------------" _
                      , "|xxx             |yyyyyy                   |zzzzzz                        " _
                      , "|xxx             |yyyyyy                   |zzzzzz                        " _
                      , "|xxx             |yyyyyy                   |zzzzzz                        " _
                      , "|===========================================================================" _
                      , "|-------------- Method 'Items' test: New title, with marging ---------------" _
                      , "| Column-01-Header |   -Column-02-Header-    |     --Column-03-Header--     " _
                      , "|------------------+-------------------------+------------------------------" _
                      , "| xxx              | yyyyyy                  | zzzzzz                       " _
                      , "| xxx              | yyyyyy                  | zzzzzz                       " _
                      , "| xxx              | yyyyyy                  | zzzzzz                       "
        End If
    End With

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

