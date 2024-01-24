Attribute VB_Name = "mLogTest"
Option Explicit
Option Compare Binary
' ----------------------------------------------------------------------
' Standard Module mLogTest: Individual tests plus a Regression-Test
' ========================= which combines them all.
'
' ----------------------------------------------------------------------
Private bRegTestFailed                  As Boolean
Private sRegTestResult                  As String
Private fso                             As New FileSystemObject
Private lLineExpected                   As Long
Private lLineResult                     As Long
Private sExpected                       As String
Private sResultExpected_FileFullName    As String
Private sResult_FileFullName            As String
Private sResult                         As String
Private sLineExpected                   As String
Private sLineResult                     As String

#If Not mMsg = 1 Then
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

Private Property Get FileString(Optional ByVal f_file_path As String) As String

    Open f_file_path For Input As #1
    FileString = Input$(LOF(1), 1)
    Close #1

End Function

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

Private Function ArrayIsAllocated(a_v As Variant) As Boolean
    On Error Resume Next
    ArrayIsAllocated = Not IsError(UBound(a_v))
End Function

Private Function ArrayTrimmed(ByVal t_v As Variant) As Variant
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Dim i   As Long
    Dim k   As Long
    Dim arr As Variant
    
    If Not ArrayIsAllocated(t_v) Then
        Exit Function
    End If
    
    '~~ Get first code line not empty
    For i = LBound(t_v) To UBound(t_v)
        If Len(Trim(t_v(i))) > 0 Then Exit For
    Next i
    
    If i < 0 Then
    Else
        '~~ Mode all items up
        For i = i To UBound(t_v)
            t_v(k) = t_v(i)
            k = k + 1
        Next i
        arr = t_v
        '~~ Eliminate trailing empty items
        Do While Trim(arr(UBound(arr))) = vbNullString And UBound(arr) > 0
            If UBound(arr) > 0 _
            Then ReDim Preserve arr(UBound(arr) - 1) _
            Else Exit Do
        Loop
        If Not Trim(arr(UBound(arr))) = vbNullString Then
            ArrayTrimmed = arr
        Else
        End If
    End If
    
End Function

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated). The services, when installed, are activated by the
' | Cond. Comp. Arg.             | Installed component |
' |------------------------------|---------------------|
' | ExecTraceBymTrc = 1          | mTrc                |
' | ExecTraceByclsTrc = 1        | clsTrc              |
' | ErHComp = 1                  | mErH                |
' I.e. both components are independant from each other!
' Note: This procedure is obligatory for any VB-Component using either the
'       the 'Common VBA Error Services' and/or the 'Common VBA Execution Trace
'       Service'.
' ------------------------------------------------------------------------------
    Dim s As String
    If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")

#If mErH = 1 Then
    '~~ The error handling also hands over to the mTrc/clsTrc component when
    '~~ either of the two is installed.
    mErH.BoP b_proc, s
#ElseIf clsTrc = 1 Then
    '~~ mErH is not installed but the mTrc is
    Trc.BoP b_proc, s
#ElseIf mTrc = 1 Then
    '~~ mErH neither mTrc is installed but clsTrc is
    mTrc.BoP b_proc, s
#End If

End Sub

Private Sub EoP(ByVal e_proc As String, Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'End of Procedure' interface for the 'Common VBA Error Services' and
' the 'Common VBA Execution Trace Service' (only in case the first one is not
' installed/activated).
' Note 1: The services, when installed, are activated by the
'         | Cond. Comp. Arg.             | Installed component |
'         |------------------------------|---------------------|
'         | ExecTraceBymTrc = 1          | mTrc                |
'         | ExecTraceByclsTrc = 1        | clsTrc              |
'         | ErHComp = 1                  | mErH                |
'         I.e. both components are independant from each other!
' Note 2: This procedure is obligatory for any VB-Component using either the
'         the 'Common VBA Error Services' and/or the 'Common VBA Execution
'         Trace Service'.
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ The error handling also hands over to the mTrc component when 'ExecTrace = 1'
    '~~ so the Else is only for the case the mTrc is installed but the merH is not.
    mErH.EoP e_proc
#ElseIf clsTrc = 1 Then
    Trc.EoP e_proc, e_inf
#ElseIf mTrc = 1 Then
    mTrc.EoP e_proc, e_inf
#End If

End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
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
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mLogTest" & "." & sProc
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

Private Sub ProvideNewTraceLogFile(ByVal p_name As String)
    Dim s As String
    With fso
        s = ThisWorkbook.Path & "\" & .GetBaseName(ThisWorkbook.Name) & "." & p_name
        If .FileExists(s) Then .DeleteFile s
    End With
    mTrc.FileFullName = s
End Sub

Private Sub RegressionTest()
' ------------------------------------------------------------------------------
' Please note: This test includes the result assertion which is the result from
'              Test_00_Regression - when ok - saved to the file
'              RegressionTestResultExpected.log in this projects parent folder.
' ------------------------------------------------------------------------------
    Const PROC = "RegressionTest"
    
    On Error GoTo eh

    sResultExpected_FileFullName = ThisWorkbook.Path & "\RegressionTestResultExpected.log"
    sResult_FileFullName = ThisWorkbook.Path & "\RegressionTestResult.log"
    
    ProvideNewTraceLogFile "RegressionTest.trc"
    BoP ErrSrc(PROC)
    mErH.Regression = True
    Test_00_Regression
    
    EoP ErrSrc(PROC)
    mErH.Regression = False
xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ResultAsserted(ByVal a_file As String, _
                                ByVal a_time_stamp As String, _
                                ByRef a_expected As Variant, _
                                ByRef a_result As Variant, _
                                ByRef a_line_expected As Long, _
                                ByRef a_line_result As Long, _
                                ByRef a_lines As Long) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the result in the log-file (a_result) is identical with the
' expected result (a_expected) whereby any line preceeding TimeStamp is ignored.
' ------------------------------------------------------------------------------
    Dim vResult         As Variant
    Dim vExpected       As Variant
    Dim i               As Long
    Dim sResult         As String
    Dim sExpected       As String
    Dim sResultFile     As String
    Dim sExpectedFile   As String
    
    ResultAsserted = True
    sExpectedFile = StringTrimmed(FileString(sResultExpected_FileFullName))
    sResultFile = StringTrimmed(FileString(a_file))
    
    vExpected = ArrayTrimmed(Split(sExpectedFile, vbCrLf))
    vResult = ArrayTrimmed(Split(sResultFile, vbCrLf))
    
    For i = LBound(vResult) To Min(UBound(vResult), UBound(vExpected))
        
        sResult = vResult(i)
        If sResult Like "*-*-*-*:*:*" _
        Then sResult = Right(sResult, Len(sResult) - Len("yy-mm-dd-hh:mm:ss"))
        
        sExpected = vExpected(i)
        If sExpected Like "*-*-*-*:*:*" _
        Then sExpected = Right(sExpected, Len(sExpected) - Len("yy-mm-dd-hh:mm:ss"))
        
        If Not sResult = sExpected Then
            ResultAsserted = False
            a_result = sResult
            a_expected = sExpected
            a_line_expected = i + 1
            a_line_result = i + 1
        End If
        a_lines = i + 1
    Next i
    
    Select Case True
        Case UBound(vResult) > UBound(vExpected)
            ResultAsserted = False
            a_result = vResult(UBound(vResult))
            a_expected = vbNullString
            a_line_result = UBound(vResult)
            a_line_expected = 0
            
        Case UBound(vResult) > UBound(vExpected)
            ResultAsserted = False
            a_result = vbNullString
            a_expected = vExpected(UBound(vExpected))
            a_line_result = 0
            a_line_expected = UBound(vExpected)
    End Select
    
End Function

Private Function StringTrimmed(ByVal s_s As String, _
                      Optional ByRef s_as_dict As Dictionary = Nothing) As String
' ----------------------------------------------------------------------------
' Returns the code (s_s) provided as a single string without leading and
' trailing empty lines. When a Dictionary is provided the string is returned
' as items with the line number as key.
' ----------------------------------------------------------------------------
    Dim s As String
    Dim i As Long
    Dim v As Variant
    
    s = s_s
    '~~ Eliminate leading empty code lines
    Do While Left(s, 2) = vbCrLf
        s = Right(s, Len(s) - 2)
    Loop
    '~~ Eliminate trailing eof
    If Right(s, 1) = VBA.Chr(26) _
    Then s = Left(s, Len(s) - 1)
    
    '~~ Eliminate trailing empty code lines
    Do While Right(s, 2) = vbCrLf
        s = Left(s, Len(s) - 2)
    Loop
    
    StringTrimmed = s
    If Not s_as_dict Is Nothing Then
        With s_as_dict
            For Each v In Split(s, vbCrLf)
                i = i + 1
                .Add i, v
            Next v
        End With
    End If
    
End Function

Private Sub Test_00_Regression()
' ------------------------------------------------------------------------------
' Regression-Test: The description of the individual test is either provided
'                  by a Title for a series of log-entries or - when no Title
'                  is specified - by the log-entries themselve.
' ------------------------------------------------------------------------------
    Const PROC = "Test_00_Regression"
    
    On Error GoTo eh
    Dim Log             As New clsLog
    Dim bTimeStamp      As Boolean: bTimeStamp = True
    Dim lLines          As Long
    Dim bAsserted       As Boolean
    
    If Not mErH.Regression Then ProvideNewTraceLogFile "RegressionTestResult.trc"
    BoP ErrSrc(PROC)
    
    '~~ ThisWorkbook because the ActiveWorkbook may be - by accident - another one
    sResult_FileFullName = ThisWorkbook.Path & "\RegressionTestResult.log"
    
    With Log
        .FileFullName = sResult_FileFullName
        .NewFile
        .WithTimeStamp bTimeStamp
        
        .Title "01-1 Title test: " _
             , "^^^^^^^^^^^^^^^^" _
             , "- 4 title lines as ParamAray (lines are comma delimited string)" _
             , "- 2 Single line log entries" _
             , "- Title left adjusted by means of a trailing space with the first title line"
        .Entry "01-1 1. Single string, new log, Single string, new log."
        .Entry "01-1 2. Single string, new log, no title. "
        
        .Title "01-2 Title test:"
        .Title "^^^^^^^^^^^^^^^^"
        .Title "(aligned centered (by no leading and no trailing space)"
        .Entry "01-2 1. Single string, new log, Single string, new log."
        .Entry "01-2 2. Single string without any width limit"
        
        .Title "- 01-3 Title: Regression test case 01-3: -"
        .Title "  ^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
        .Title "(centered, filled with - by a leainding and a trailing - with the first title line)"
        .Entry "01-3 1. Single string, new log. This is an extra long text to force all title lines with fill characters"
        .Entry "01-3 2. Single string without any width limit"
                
        '~~ --------------------------------------------------------------------
        '~~ Test 02-1: - NewLog explicitly idicateas a new series of log entries
        '~~            - The first line implicitly specifies the columnns
        '~~              alignment: R, C, R, L
        '~~            - The chage between column aligned and single string
        '~~              is possible. Column aligned entries are only correct
        '~~              implitly aligned when the first entry is an items entry
        '~~ --------------------------------------------------------------------
        .NewLog
        .AlignmentItems "|R|C|L|L|"
        .Entry "02 Test: Explicit items alignment "
        .Entry "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
        .Entry "- 'NewLog' explicitly called to indicate a new series of log entries without specifying a title, resulting in the delimiter line above"
        .Entry "- Explicit call of the 'AlignmentItems' method in order to have the below items adjusted as desired"
        .Entry "  Note: An implicit alignment spec is possible but only with the first of a series of log entries"
        .Entry " 02-1", "xxxx", " yyyyyy", "No Title! Alignments: R, C, R, L; Rightmost column without width limit"
        .Entry "02-1", "xxxx", " yyyyy", "... correct aligned in columns because the first entry indicated column alignment implicitly!"
        
        '~~ --------------------------------------------------------------------
        '~~ Test 02-2: - NewLog explicitly idicateas a new series of log entries
        '~~            - The first line implicitly specifies the columnns
        '~~              alignment: R, C, R, L
        '~~            - The chage between column aligned and single string
        '~~              is possible. Column aligned entries are only correct
        '~~              implitly aligned when the first entry is an items entry
        '~~ --------------------------------------------------------------------
        .NewLog
        .Entry "No Title, 'NewLog' explicitly called to indicate a new series of log entries"
        .Entry "AlignmentItems explicitly called to align the following two items correct"
        .Entry " 02-2", "xxxx", " yyyyyy", "Alignments: R, C, R, L; Rightmost column without width limit"
        .Entry "02-2", "xxxx", " yyyyy", "... correct aligned in columns because AlignmentItems explicitly specified them!"
        
        .Title "03 Test: Headers with implicit alignment " _
             , "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" _
             , "- 'Headers' method used to specify 2 column headers with implicitely specified alignment by means of leading and trailing spaces (R, C, C, L)" _
             , "- The maximum column width is the maximimum of the width implicitly specified by:" _
             , "  - the 'Headers' first line's specificateion" _
             , "  - the width of the first line's items width." _
             , "  - though less likely, the Entry-Items alignment is implicitly specified by the first line's items using leading and trailing spaces: R, L, L, L"
        .Headers " Nr", "Item", "Item", "Item 3 "
        .Headers "   ", " 1  ", " 2  ", "(no width limit) "
        .Entry " 03", "xxxx ", "yyyyyy ", " Rightmost column without width limit! (this first line implicitly indicated the columns width for being considered by the header) "
        .Entry "03", "xxxx", "yyyy", "         zzzzzz (note that leading spaces are preserved when/because the first line implicitly indicated left adjusted)"
        .Entry "03", "xxxx", "yyyyy", "zzzzzz "
        
        .NewLog ' has no effect since indicated by the below title
        .Title "04 Test: Explicit items length specification " _
             , "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" _
             , "- because no 'Headers' are specified the ColsDelimiter fefaults is a single space and the ColsMargin is a vbNullString" _
             , "- the implicit items alignment is: R, L, C, R"
        .MaxItemLengths 2, 10, 25, 30
        .Entry " 04", "xxx ", "yyyyyy", "     zzzzzz"
        .Entry "04", "xxx ", "yyyyyy ", "zzzzzz "
        .Entry "04", "xxx ", "yyyyyy ", "zzzzzz "
        
        .Title "05 Test: Explicit ColsDelimiter " _
             , "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" _
             , "- The 'ColsDelimiter' property explicitly specifies as a single space " _
             , "  (the specified headers would result in a | (vertical bar) as the columns delimiter" _
             , "- The Header alignments are implicitly: R, C, L, L" _
             , "- The item alignments are implicitly: R, L, C (filled with -), L" _
             , "  (leading spaces with left aligned items are preserved by default)"
        .ColsDelimiter = " "
        .Headers "| Nr| Item-1 |  Item-2  |Item-3 (no width limit) "
        .Entry " 05", "xxxx ", "- yyyyyy -", "Rightmost column without width limit! "
        .Entry " 05", "xxxx ", "yyyy       ", "         zzzzzz (note that leading spaces are preserved when/because the first line implicitly indicated left adjusted)"
        .Entry "05", "xxxx ", "yyyyy       ", "zzzzzz "
         
        .Title "06 Test: Explicit AlignmentHeaders, AlignmentItems, and MaxItemLengths "
        .Title "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^"
        .Title "- The ColsDelimiter explicitly specifies as a single space "
        .Title "- AlignmentHeaders explicitly specify: R, C, L, L"
        .Title "- AlignmentItems explicitly specifies: R, L, C (filled with -), L"
        .Title "- Leading spaces with left aligned items are preserved by default"
        .Title "- MaxItemsLengths explicitly specifies 3,10,25,30"
        .ColsDelimiter = " "
        .AlignmentHeaders "|R|C|L|L|"
        .Headers "|Nr|Item-1|Item-2|Item-3 (no width limit)"
        .AlignmentItems "R", "L", "- C -"
        .MaxItemLengths 3, 10, 25, 30
        .Entry "06", "xxxx", "yyyyyy", "Rightmost column without width limit! "
        .Entry "06", "xxxx", "yyyy  ", "         zzzzzz (note that leading spaces are preserved when/because the first line implicitly indicated left adjusted)   "
        .Entry "06", "xxxx", "yyyyy ", "zzzzzz "
         
         .Title "07 Test: AlignmentItems explicitly only specifies column 2 " _
              , "^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^^" _
              , "- Column 1: Alignment implicitly Right adjusted (because not explicitely specified" _
              , "- Column 2: - Length explicily specified = 25" _
              , "            - Alignment explicitly specified: : left adjusted, filled with . (dots), terminated with a : (colon)" _
              , "- Column 3: Alignment implicitly Left adjusted (because not explicitely specified"
        .MaxItemLengths , 25
        .AlignmentItems , "L.:"
        .Headers "| Nr | Item | Comment |"
        .ColsDelimiter = " "
        .Entry " 07", "xxxx ", "Rightmost column without width limit! "
        .Entry "07", "xxxxxxxxxxxxxxxxxxxx", "         zzzzzz (note that leading spaces are preserved when/because the first line implicitly indicated left adjusted)  "
        .Entry "07", "xxxxxxxxx", "zzzzzz"
    
         .Title "08 Test: Not provided items "
         .Title "^^^^^^^^^^^^^^^^^^^^^^^^^^^"
         .Title "- The second entry does not provide an item for the secind column"
         .Title "- Second column explicit specified width = 25"
         .Title "- Second columnn explicit alignment specified Left adjusted, filled with . terminated with :"
        .MaxItemLengths , 25
        .AlignmentItems , "L.:"
        .Headers "| Nr | Item | Comment |"
        .ColsDelimiter = " "
        .Entry " 08", "xxxx ", "Rightmost column without width limit! "
        .Entry "08", , "zzzzzz (note that leading spaces are preserved when/because the first line implicitly indicated left adjusted)  "
        .Entry "08", "xxxxxxxxx", "zzzzzz"
    
        .Reorg
    End With

xt: EoP ErrSrc(PROC)
    If mErH.Regression Then
        bAsserted = ResultAsserted(Log.LogFile _
                                 , bTimeStamp _
                                 , sExpected _
                                 , sResult _
                                 , lLineExpected _
                                 , lLineResult _
                                 , lLines)
        If Not bAsserted Then
            mTrc.LogInfo = "Test f a i l e d !"
            mTrc.LogInfo = "Line " & Format(lLineExpected, "00") & " Expected: " & sExpected
            mTrc.LogInfo = "Line " & Format(lLineResult, "00") & " Result: " & sResult
        Else
            mTrc.LogInfo = "Test p a s s e d !"
            mTrc.LogInfo = lLines & " Result lines match with " & lLines & " expected result lines!"
        End If
    End If
    If mErH.Regression Then
        If Not bAsserted Then
            mTrc.Dsply
            Log.Dsply sResultExpected_FileFullName
        Else
            mTrc.Dsply
        End If
    Else
        Log.Dsply
        mTrc.Dsply
    End If
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub Test_01_Align()
    Dim Log As New clsLog
    Dim sResult As String
    Dim lWidth  As Long
    
    With Log
        lWidth = 10
        sResult = "|" & .AlignString(" ", lWidth, "R", " ", " ") & "|"
        Debug.Assert sResult = "|          |"
        Debug.Assert Len(sResult) = lWidth + 2
        
        lWidth = 10
        sResult = "|" & .AlignString(" ", lWidth, "C", " ", " ") & "|"
        Debug.Assert sResult = "|          |"
        Debug.Assert Len(sResult) = lWidth + 2
        
        lWidth = 10
        sResult = "|" & .AlignString(" ", lWidth, "L", " ", " ") & "|"
        Debug.Assert sResult = "|          |"
        Debug.Assert Len(sResult) = lWidth + 2
        
        lWidth = 10
        sResult = .AlignString("  ", 10, "L", "", ".:")
        Debug.Assert sResult = "             "
        Debug.Assert Len(sResult) = lWidth + 3
            
        lWidth = 10
        sResult = .AlignString("  ", 10, "L", "", ".")
        Debug.Assert sResult = "            "
        Debug.Assert Len(sResult) = lWidth + 2
    
    End With
    
End Sub

