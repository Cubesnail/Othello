Attribute VB_Name = "modOthello"
Option Explicit

'Declare the global constants, types and varaibles.

Const MAX = 10
Public FirstName As String
Public SecondName As String
Public Turn As Integer
Public FirstPicture As Integer
Public SecondPicture As Integer
Public CellWidth As Integer
Public Mode As Integer
Public PieceMove(1 To MAX, 1 To MAX) As Integer
Const RECORDLEN = 21

Type Score
    Name As String * 15
    Time As Integer
    FirstScore As Integer
    SecondScore As Integer
End Type

Public Scores(1 To 3, 1 To 6) As Score

Type Move ' Declare types specifically used for bots.
    X As Integer
    Y As Integer
    Amount As Integer
End Type

Public InvalidMove As Move
Public MoveList(0 To (MAX * MAX)) As Move

'Reads record files.
Public Sub ReadScore()
    Dim K As Integer

    K = 0

    Open App.Path & "\easyscore.rec" For Random As #1 Len = RECORDLEN
    Do While (Not EOF(1)) Or K < 6
        K = K + 1
        Get #1, K, Scores(1, K)
    Loop
    Close #1
    K = 0
    Open App.Path & "\mediumscore.rec" For Random As #1 Len = RECORDLEN
    Do While (Not EOF(1)) Or K < 6
        K = K + 1
        Get #1, K, Scores(2, K)
    Loop
    Close #1
    K = 0
    Open App.Path & "\hardscore.rec" For Random As #1 Len = RECORDLEN
    Do While (Not EOF(1)) Or K < 6
        K = K + 1
        Get #1, K, Scores(3, K)
    Loop
    Close #1
End Sub

'Saves scores to record files.

Public Sub SaveScore()

    Dim K As Integer
    
    On Error GoTo Error
    Kill App.Path & "\easyscore.rec"
    Open App.Path & "\easyscore.rec" For Random As #1 Len = RECORDLEN
    For K = 1 To 5
        Put #1, K, Scores(1, K)
    Next K
    Close #1
    
    Kill App.Path & "\mediumscore.rec"
    Open App.Path & "\mediumscore.rec" For Random As #1 Len = RECORDLEN
    For K = 1 To 5
        Put #1, K, Scores(2, K)
    Next K
    Close #1
    
    Kill App.Path & "\hardscore.rec"
    Open App.Path & "\hardscore.rec" For Random As #1 Len = RECORDLEN
    For K = 1 To 5
        Put #1, K, Scores(3, K)
    Next K
    Close #1
    Exit Sub
Error:
    Resume Next

End Sub

Function Hardbot() As Move
    
    Dim K As Integer
    Dim Temp As Move
    Dim Amount As Integer
    Dim Priority(1 To (MAX * MAX)) As Integer
    Dim Hard As Integer
    Dim EqualPriority(1 To (MAX * MAX)) As Integer
    Dim EqualCount As Integer
    
    Amount = GetMove()
    
    Hard = 1
    If Amount > 0 Then
        For K = 1 To Amount
            With MoveList(K)
                Priority(K) = 0
                If (.X = 1 Or .X = MAX) And (.Y = 1 Or .Y = MAX) Then
                    Priority(K) = Priority(K) + 100     'Check for corners.
                ElseIf .X = 1 Or .X = MAX Then
                    If .Y = 2 Or .Y = MAX - 1 Then    'Check for moves beside corners.
                        Priority(K) = Priority(K) - 100
                    ElseIf .Y = 3 Or .Y = MAX - 2 Then   'Check for moves 3 away from corners.
                        Priority(K) = Priority(K) + 15
                    Else
                        Priority(K) = Priority(K) + 15  'Prioritize edges.
                    End If
                ElseIf .Y = 1 Or .Y = MAX Then
                    If .X = 2 Or .X = MAX - 1 Then    'Check for moves beside corners.
                        Priority(K) = Priority(K) - 100
                    ElseIf .X = 3 Or .X = MAX - 2 Then  'Check for moves 3 away from corners.
                        Priority(K) = Priority(K) + 15
                    Else
                        Priority(K) = Priority(K) + 15 'Prioritize edges.
                    End If
                ElseIf .X = 2 Or .X = MAX - 1 Then      'Check for moves beside edges (horizontal).
                    If .Y = 2 Or .Y = MAX - 1 Then      'Check for moves beside corners.
                        Priority(K) = Priority(K) - 50
                    ElseIf .Y = 3 Or .Y = MAX - 2 Then   'Check for moves 3 away from corners.
                        Priority(K) = Priority(K) + 5
                    Else
                        Priority(K) = Priority(K) - 10
                    End If
                ElseIf .Y = 2 Or .Y = MAX - 1 Then      'Check for moves beside edges (vertical).
                    If .X = 2 Or .X = MAX - 1 Then      'Check for moves beside corners.
                        Priority(K) = Priority(K) - 50
                    ElseIf .X = 3 Or .X = MAX - 2 Then   'Check for moves 3 away from corners.
                        Priority(K) = Priority(K) + 5
                    Else
                        Priority(K) = Priority(K) - 10
                    End If
                ElseIf (.X = 3 Or .X = MAX - 2) And (.Y = 3 Or .Y = MAX - 2) Then
                    Priority(K) = Priority(K) + 15      'Prioritize 3 away from corners.
                ElseIf .X = 3 Or .X = MAX - 2 Then      'Check for moves 3 away from edges (horizontal).
                    Priority(K) = Priority(K) + 5
                ElseIf .Y = 3 Or .Y = MAX - 2 Then      'Check for moves 3 away from edges(vertical).
                    Priority(K) = Priority(K) + 5
                End If
                Priority(K) = Priority(K) + .Amount     'Check for the amount of pieces it takes.
                If Priority(K) > Priority(Hard) Then    'Store moves with equal priority.
                    Hard = K
                    EqualCount = 0
                    EqualPriority(1) = K
                ElseIf Priority(K) = Priority(Hard) Then
                    EqualCount = EqualCount + 1
                    EqualPriority(EqualCount) = K
                End If
            End With
        Next K
    Else
        Hardbot = MoveList(0)
        Exit Function
    End If
    Hardbot = MoveList(EqualPriority(Int(Rnd() * EqualCount) + 1))  'Randomize to choose one with most priority.
    
End Function

Function MediumBot() As Move
    
    Dim Temp As Move
    Dim Amount As Integer
    Dim Priority(1 To (MAX * MAX)) As Integer
    Dim Medium As Integer
    Dim K As Integer
    Dim EqualPriority(1 To (MAX * MAX)) As Integer
    Dim EqualCount As Integer
    
    Amount = GetMove()
    EqualCount = 0
    Medium = 1
    
    If Amount > 0 Then
        For K = 1 To Amount
            With MoveList(K)
                Priority(K) = 0
                If (.X = 1 Or .X = MAX) And (.Y = 1 Or .Y = MAX) Then   'Prioritize corners.
                    Priority(K) = Priority(K) + 100
                End If
                Priority(K) = Priority(K) + .Amount     'Prioritize moves with more pieces taken.
                If Priority(K) > Priority(Medium) Then  'Store moves with the highest priority.
                    Medium = K
                    EqualCount = 0
                    EqualPriority(1) = K
                ElseIf Priority(K) = Priority(Medium) Then
                    EqualCount = EqualCount + 1
                    EqualPriority(EqualCount) = K
                End If
            End With
        Next K
    Else
        MediumBot = MoveList(0)
        Exit Function
    End If
    MediumBot = MoveList(EqualPriority(Int(Rnd() * EqualCount) + 1))    'Randomize to find a move with the highest priority.
End Function

Function EasyBot() As Move
    
    Dim Temp As Move
    Dim Easy As Integer
    Dim K As Integer
    Dim Amount As Integer
    Amount = GetMove()
    Easy = Int((Rnd() * Amount)) + 1
    Temp = MoveList(Easy)
    EasyBot = Temp
End Function

'Stores all possible moves and amount of pieces it converts in an array.

Function GetMove() As Integer
    
    Dim Temp As Integer
    Dim Counter As Integer
    Dim K As Integer
    Dim M As Integer
    Dim Dir As Integer
    Dim Check As Boolean
    
    Counter = 1
    
    For K = 1 To MAX
        For M = 1 To MAX
            Check = False
            MoveList(Counter).Amount = 0
            For Dir = 1 To 8
                If PieceMove(M, K) = 0 Then
                    If Valid(Turn, M, K, Dir, False, , , MoveList(Counter).Amount) Then
                        Check = True
                    End If
                End If
            Next Dir
            If Check = True Then
                MoveList(Counter).X = M
                MoveList(Counter).Y = K
                Counter = Counter + 1
            End If
        Next M
    Next K
    
    GetMove = Counter - 1
End Function

'Converts x and y cooridnates to index.

Function XYtoNum(ByVal X As Integer, ByVal Y As Integer) As Integer
    Dim Temp As Integer
    Temp = 0
    Temp = Temp + (Y * 10)
    Temp = Temp + X
    XYtoNum = Temp
End Function

'Converts index to x coordinate.

Function NumtoX(ByVal Num As Integer) As Integer
    Dim Temp As Integer
    Temp = Num Mod MAX
    If Temp = 0 Then
        Temp = MAX
    End If
    NumtoX = Temp
End Function

'Converts index to y coordinate.

Function NumtoY(ByVal Num As Integer) As Integer
    Dim Temp As Integer
    Temp = Num \ MAX
    If NumtoX(Num) = MAX Then
        Temp = Temp - 1
    End If
    NumtoY = Temp
End Function
    
'Checks for any possible moves and if none can be made, returns false.
    
Public Function CheckEnd() As Boolean
    Dim K As Integer
    Dim M As Integer
    Dim Temp As Boolean
    Dim X As Integer
    
    Temp = False
    K = 0
    Do While Not Temp And K < MAX
        K = K + 1
        M = 0
        Do While Not Temp And M < MAX
            M = M + 1
            If PieceMove(K, M) = 0 Then
                Temp = True
            End If
        Loop
    Loop
    
    K = 0
    If Temp Then
        Temp = False
        Do While Not Temp And K < MAX
            K = K + 1
            M = 0
            Do While Not Temp And M < MAX
                M = M + 1
                X = 0
                Do While Not Temp And X < 8
                    X = X + 1
                    Temp = Valid(Turn, K, M, X, False)
                Loop
            Loop
        Loop
    End If
    If Temp = True Then
        CheckEnd = False
    Else
        CheckEnd = True
    End If
    
End Function

'Checks for a valid move and places pieces.

Public Function Valid(Player As Integer, ByVal X As Integer, _
    ByVal Y As Integer, ByVal Dir As Integer, Change As Boolean, Optional Start As Boolean = True, _
        Optional ByRef Ended As Boolean = False, Optional ByRef Count As Integer = 0) As Boolean
        
    Dim Temp As Boolean
    
    Select Case Dir
        Case 1, 2, 8
            X = X - 1
        Case 4, 5, 6
            X = X + 1
    End Select
    
    Select Case Dir
        Case 2, 3, 4
            Y = Y - 1
        Case 8, 7, 6
            Y = Y + 1
    End Select
    
    If X > 0 And Y > 0 And X <= MAX And Y <= MAX Then
        If Ended = False Then
        
            If PieceMove(X, Y) = Player Mod 2 + 1 Then
                    Start = False
                    Temp = Valid(Player, X, Y, Dir, Change, Start, Ended, Count)
            ElseIf PieceMove(X, Y) = Player Then
                If Start = False Then
                    Ended = True
                End If
            End If
        End If
        If Ended = True Then
            If PieceMove(X, Y) <> Player Then
            Count = Count + 1
            End If
            If Change Then
                PieceMove(X, Y) = Player
            End If
            
            Valid = True
        End If
    End If
    If Temp = True Then
        Valid = True
    End If
End Function
