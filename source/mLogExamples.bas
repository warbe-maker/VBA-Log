Attribute VB_Name = "mLogExamples"
Option Explicit

Private Sub Example_1()
' ----------------------------------------------------------------------
'
' ----------------------------------------------------------------------

    Dim Log As New clsLog
    
    '~~ Log Class-Module preparation
    With Log
        .WithTimeStamp = True               ' defaults to False when ommited
        .AlignmentItems "|C|L.:|L|"         ' explicit items alignment spec
        .MaxItemLengths 6, 15, 30           ' explicit spec of the required column width
        .Headers "| Nr | Item | Comment |"  ' implicitly aligned centered
        .ColsDelimiter = " "                ' defaults to | when headers are specified
    End With
    '~~ Any code
    
    Log.Entry "xxxx", "yyyyyy", "zzzzzzzz"
    Log.Entry "xxx", "yyyyyyyyyyyyyyy", "zzzzzzzzzzzzzzzzzzzzz"
    Log.Dsply

End Sub

Private Sub Example_2()
' ----------------------------------------------------------------------
'
' ----------------------------------------------------------------------

    Dim Log As New clsLog
    
    With Log
        '~~ Preparation
        .WithTimeStamp = True                   ' defaults to False when ommited
        .AlignmentItems "|R|C|L|"               ' explicit items alignment spec
        .MaxItemLengths 6, 15, 30               ' explicit spec of the required column width
        .Headers "| Column|  Column  |Column |" ' this line (only) implicitly specifies the alignment
        .Headers "|    1  |   2      |   3   |" ' any alignment implied is ignored
        .Headers "|(right)|(centered)|(left) |" ' any alignment implied is ignored
    End With
    
    '~~ Any code
    
    Log.Entry "xxxx", "yyyyyy", "zzzzzzzz"
    Log.Dsply

End Sub

