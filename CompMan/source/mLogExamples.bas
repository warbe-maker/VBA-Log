Attribute VB_Name = "mLogExamples"
Option Explicit

Private Sub Example_1()
' ----------------------------------------------------------------------
'
' ----------------------------------------------------------------------

    Dim Log As New clsLog
    
    '~~ Log Class-Module preparation
    With Log
        .WithTimeStamp                      ' leading timestamp with each log entry (defaults to False when omitted)
        .ColsSpecs = "C6, L15.:, L30"       ' explicit items alignment spec
        .ColsHeader = "Nr, Item, Comment"   ' aligned centered with the width spec above
        .ColsDelimiter = " "                ' defaults to | when headers are specified
    End With
    '~~ Any code
    
    Log.ColsItems "xxxx", "yyyyyy", "zzzzzzzz"
    Log.ColsItems "xxx", "yyyyyyyyyyyyyyy", "zzzzzzzzzzzzzzzzzzzzz"
    Log.Dsply

End Sub

Private Sub Example_2()
' ----------------------------------------------------------------------
'
' ----------------------------------------------------------------------

    Dim Log As New clsLog
    
    With Log
        '~~ Preparation
        .WithTimeStamp                              ' leading timestamp with each log entry (defaults to False when omitted)
        .ColsSpecs = "R6, C15, L30"                 ' explicit items alignment spec
        .ColsHeader = "Column , Column    , Column"      ' aligned centered with the width spec above
        .ColsHeader = "1      , 2         , 3"           ' aligned centered with the width spec above
        .ColsHeader = "(right), (centered), (left)"   ' any alignment implied is ignored
    End With
    
    '~~ Any code
    
    Log.ColsItems "xxxx", "yyyyyy", "zzzzzzzz"
    Log.Dsply

End Sub

