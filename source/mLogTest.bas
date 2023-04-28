Attribute VB_Name = "mLogTest"
Option Explicit
' ----------------------------------------------------------------------
' Standard Module mLogTest: Regression-Test for the clsLog Class Module
' =========================
'
' ----------------------------------------------------------------------

Private Sub Test_Regression()
    Test_Property_Path
    Test_Property_Name
    Test_ColsHeader
    Test_WriteHeader
End Sub

Private Sub Test_Property_Path()
    Dim Log As New clsLog
    Log.Path = ThisWorkbook.Path
    Debug.Assert Log.FileFullName = ThisWorkbook.Path & "\Service.log"
    Set Log = Nothing
End Sub

Private Sub Test_Property_Name()
    Dim Log As New clsLog
    Log.FileName = "TestService.log"
    Debug.Assert Log.FileFullName = ThisWorkbook.Path & "\TestService.log"
    Set Log = Nothing
End Sub

Private Sub Test_ColsHeader()
    Const HEADER_1 = "Column-01-Header"
    Const HEADER_2 = "-Column-02-Header-"
    Const HEADER_3 = "--Column-03-Header--"
    
    Dim Log As New clsLog
    
    Log.ColsHeader HEADER_1, HEADER_2, HEADER_3
    Debug.Assert Log.vColsWidth(0) = Len(HEADER_1) + 2: Debug.Assert Log.vColsHeader(0) = HEADER_1
    Debug.Assert Log.vColsWidth(1) = Len(HEADER_2) + 2: Debug.Assert Log.vColsHeader(1) = HEADER_2
    Debug.Assert Log.vColsWidth(2) = Len(HEADER_3) + 2: Debug.Assert Log.vColsHeader(2) = HEADER_3
    Set Log = Nothing
End Sub

Private Sub Test_ColsWidth()
    Const HEADER_1 = "Column-01-Header"
    Const HEADER_2 = "-Column-02-Header-"
    Const HEADER_3 = "--Column-03-Header--"
    
    Dim Log As New clsLog
    
    With Log
        .ColsHeader HEADER_1, HEADER_2, HEADER_3
        .ColsWidth 20, 25, 30
        Debug.Assert .vColsWidth(0) = 20: Debug.Assert .vColsHeader(0) = HEADER_1
        Debug.Assert .vColsWidth(1) = 25: Debug.Assert .vColsHeader(1) = HEADER_2
        Debug.Assert .vColsWidth(2) = 30: Debug.Assert .vColsHeader(2) = HEADER_3
    End With
    Set Log = Nothing

End Sub

Private Sub Test_WriteHeader()
    Const HEADER_1 = "Column-01-Header"
    Const HEADER_2 = "-Column-02-Header-"
    Const HEADER_3 = "--Column-03-Header--"
    
    Dim Log As New clsLog
    
    With Log
        .ColsHeader HEADER_1, HEADER_2, HEADER_3
        .ColsWidth 20, 25, 30
        .WriteHeader
        .Dsply
    End With
    Set Log = Nothing

End Sub

Private Sub Test_AddLog()
    Const HEADER_1 = "Column-01-Header"
    Const HEADER_2 = "-Column-02-Header-"
    Const HEADER_3 = "--Column-03-Header--"
    
    Dim Log As New clsLog
    
    With Log
        .ServiceHeader = "Test: AddLog"
'        .WithTimeStamp = True ' Default
        .ColsMargin = vbNullString
        .ColsWidth 10, 25, 30
        .ColsHeader " ", HEADER_1, HEADER_2, HEADER_3
        .WriteHeader
        .AddLog "xxx", "yyyyyy", "zzzzzz"
        .AddLog "xxx", "yyyyyy", "zzzzzz"
        .AddLog "xxx", "yyyyyy", "zzzzzz"
        .Dsply
    End With
    Set Log = Nothing

End Sub

