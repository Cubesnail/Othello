Attribute VB_Name = "Module2"
'Option Explicit
'
'Function CheckDir(X As Integer, Y As Integer) As Integer
'    Dim Temp As Integer
'    Dim Total As Integer
'    Dim Other As Boolean
'    Dim Checked As Boolean
'    Dim Piece As Integer
'    Dim OriginalX As Integer
'    Dim OriginalY As Integer
'    OriginalX = X
'    OriginalY = Y
'    Other = False
'    Total = 0
'    Temp = 0
'
'    Check Up from the piece
'
'    Do
'        Y = Y - 1
'        If X = 0 Or Y = 0 Or X = 10 Or Y = 10 Or PieceMove(X, Y) = 0 Then
'            Other = True
'        End If
'        If Temp > 0 Then
'            If PieceMove(X, Y) = Piece Then Other
'        End If
'
'        If PieceMove(X, Y) < 0 And PieceMove(X, Y) <> Piece And Not Other Then
'            Temp = Temp + 1
'        End If
'
'    Loop While Not Other
'
'    Check up right from the piece
'
'    X = OriginalX
'    Y = OriginalY
'    Other = False
'    Total = Temp
'    Temp = 0
'    Do
'        If X = 0 Or Y = 0 Or X = 10 Or Y = 10 Or PieceMove(X, Y) = 0 Then
'            Other = True
'        End If
'        X = X + 1
'        Y = Y - 1
'        If PieceMove(X, Y) < 0 And PieceMove(X, Y) <> Piece Then
'            Temp = Temp + 1
'        End If
'    Loop While Not Other
'
'    Check right from the piece
'
'    X = OriginalX
'    Y = OriginalY
'    Other = False
'    Total = Temp
'    Temp = 0
'
'    Do
'        If X = 0 Or Y = 0 Or X = 10 Or Y = 10 Or PieceMove(X, Y) = 0 Then
'            Other = True
'        End If
'        X = X + 1
'        If PieceMove(X, Y) < 0 And PieceMove(X, Y) <> Piece Then
'            Temp = Temp + 1
'        End If
'    Loop While Not Other
'
'    Check down right from the piece
'
'    X = OriginalX
'    Y = OriginalY
'    Other = False
'    Total = Temp
'    Temp = 0
'
'    Do
'        If X = 0 Or Y = 0 Or X = 10 Or Y = 10 Or PieceMove(X, Y) = 0 Then
'            Other = True
'        End If
'        X = X + 1
'        Y = Y + 1
'        If PieceMove(X, Y) < 0 And PieceMove(X, Y) <> Piece Then
'            Temp = Temp + 1
'        End If
'    Loop While Not Other
'
'    Check down from the piece
'
'    X = OriginalX
'    Y = OriginalY
'    Other = False
'    Total = Temp
'    Temp = 0
'
'    Do
'        If X = 0 Or Y = 0 Or X = 10 Or Y = 10 Or PieceMove(X, Y) = 0 Then
'            Other = True
'        End If
'        X = X
'        Y = Y + 1
'        If PieceMove(X, Y) < 0 And PieceMove(X, Y) <> Piece Then
'            Temp = Temp + 1
'        End If
'    Loop While Not Other
'
'    Check down left from the piece
'
'    X = OriginalX
'    Y = OriginalY
'    Other = False
'    Total = Temp
'    Temp = 0
'
'    Do
'        If X = 0 Or Y = 0 Or X = 10 Or Y = 10 Or PieceMove(X, Y) = 0 Then
'            Other = True
'        End If
'        X = X - 1
'        Y = Y + 1
'        If PieceMove(X, Y) < 0 And PieceMove(X, Y) <> Piece Then
'            Temp = Temp + 1
'        End If
'    Loop While Not Other
'
'    Check left from the piece
'
'    X = OriginalX
'    Y = OriginalY
'    Other = False
'    Total = Temp
'    Temp = 0
'
'    Do
'        If X = 0 Or Y = 0 Or X = 10 Or Y = 10 Or PieceMove(X, Y) = 0 Then
'            Other = True
'        End If
'        X = X - 1
'        Y = Y
'        If PieceMove(X, Y) < 0 And PieceMove(X, Y) <> Piece Then
'            Temp = Temp + 1
'        End If
'    Loop While Not Other
'
'    Check up left from the piece
'
'    X = OriginalX
'    Y = OriginalY
'    Other = False
'    Total = Temp
'    Temp = 0
'
'    Do
'        If X = 0 Or Y = 0 Or X = 10 Or Y = 10 Or PieceMove(X, Y) = 0 Then
'            Other = True
'        End If
'        X = X - 1
'        Y = Y - 1
'        If PieceMove(X, Y) < 0 And PieceMove(X, Y) <> Piece Then
'            Temp = Temp + 1
'        End If
'    Loop While Not Other
'
'    CheckDir = Total
'
'
'End Function
'
'Determine if the user wishes to exit.
'Public Sub ConfirmExit()
'
'    Dim Reply As Integer
'    Dim Buttons As Integer
'
'    Buttons = vbQuestion + vbYesNo + vbDefaultButton2
'
'    Reply = MsgBox("Do you wish to exit?", Buttons, "Exit")
'
'    If Reply = vbYes Then
'        Beep
'        End
'    End If
'
'End Sub
'
'Obtain name and path of file.
'Public Function GetFile(Dialog As Control) As String
'
'    Dialog.Filename = ""
'    Dialog.InitDir = App.Path
'    Dialog.Filter = "Text Files|*.txt|All Files |*.*"
'
'    Dialog.ShowOpen
'
'    GetFile = Dialog.Filename
'
'End Function
'
'Read a text file.
'Public Sub ReadFileTXT(Cust() As CustRec, ByVal FN As String, RecLen As Integer, NumRec As Integer)
'
'    Dim X As Integer
'
'    X = 0
'
'    Open FN For Input As #1
'
'    Do While Not (EOF(1))
'
'        X = X + 1
'        Input #1, Cust(X).LastName, Cust(X).FirstName, Cust(X).HF, Cust(X).Mark
'
'    Loop
'
'    Close #1
'
'    NumRec = X
'    RecLen = Len(Cust(X))
'
'End Sub
'
'Determine if string S is too long and shortens it if it is.
'Public Function StringTrim(ByVal S As String) As String
'
'    Const TRIM_NUM = 15
'
'    Dim St As String
'
'    St = S
'
'    If Len(S) > TRIM_NUM Then
'
'        St = Left$(S, 12) & "..."
'
'    End If
'
'    StringTrim = St
'
'End Function
'
'Prompts the user for a number within the range.
'Public Function WithinRange() As Integer
'    Const HIGH = 10             'Not inclusive
'    Const LOW = 1               'Not inclusive
'
'    Dim Num As Integer
'    Dim Msg As String
'
'    Msg = "Number? (" & Trim$(Str$(LOW)) & "," & Trim$(Str$(HIGH)) & ")"
'
'    Do
'        Num = Val(InputBox$(Msg, "Num"))
'    Loop While Num > HIGH Or Num < LOW
'
'    WithinRange = Num
'
'End Function
'Checks file type.
'Public Function CheckFileType(ByVal Str As String)
'
'    Const LENFILETYPE = 3
'    Const FILETYPE1 = "txt"
'    Const FILETYPE2 = "rec"
'    Dim TempStr As String
'    Dim K As Integer
'    Dim PeriodCheck As Boolean
'
'    PeriodCheck = False
'
'    Retrieve file type extension.
'
'    For K = 1 To Len(Str)
'        If PeriodCheck = True Then
'            TempStr = TempStr & Mid$(Str, K, 1)
'        Else
'            If Mid$(Str, K, 1) = "." Then
'                PeriodCheck = True
'            End If
'        End If
'    Next K
'
'    Assign value of extension to function.
'
'    If TempStr = FILETYPE1 Then
'        CheckFileType = FILETYPE1
'    ElseIf TempStr = FILETYPE2 Then
'        CheckFileType = FILETYPE2
'    End If
'
'End Function
'
'Reads a record file.
'Public Sub ReadFileREC(Record() As CustRec, ByVal Filename As String, ByVal RecordLen As Integer, NumRecords As Integer)
'    Dim K As Integer
'
'    K = 0
'
'    Open Filename For Random As #1 Len = RecordLen
'    Do While Not EOF(1)
'        K = K + 1
'        Get #1, K, Record(K)
'    Loop
'    Close #1
'
'    NumRecords = K - 1
'
'End Sub
'
'Public Sub SaveFile(ByVal Filename As String, Record() As CustRec, ByVal NumRecords As Integer, ByVal RecLength As Integer)
'
'    Dim X As Integer
'
'    On Error GoTo ErrorSwagger
'
'    Kill Filename
'
'    Open Filename For Random As #1 Len = RecLength
'
'    For X = 1 To NumRecords
'        Put #1, X, Record(X)
'    Next X
'
'    Close #1
'
'    Exit Sub
'
'ErrorSwagger:
'    Resume Next
'
'End Sub
'
