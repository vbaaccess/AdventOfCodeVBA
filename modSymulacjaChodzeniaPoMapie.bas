Attribute VB_Name = "modSymulacjaChodzeniaPoMapie"
Option Compare Database
Option Explicit

'modSymulacjaChodzeniaPoMapie

Sub TestWalk()
    Call WalkFromFile("C:\Project\AdventOfCodeVBA\Input.txt")
End Sub

Sub WalkFromFile(filePath As String)
    Dim f As Integer
    Dim fileContent As String
    Dim i As Long
    
    ' macierz dynamiczna
    Dim grid() As Long
    Dim rows As Long, cols As Long
    Dim x As Long, y As Long  ' pozycja
    
    ' start: 1x1
    rows = 1: cols = 1
    ReDim grid(0 To rows - 1, 0 To cols - 1)
    
    x = 0: y = 0
    grid(y, x) = grid(y, x) + 1  ' punkt startowy
    
    ' wczytaj plik
    f = FreeFile
    Open filePath For Input As #f
    fileContent = Input$(LOF(f), f)
    Close #f
    
    ' petla po znakach
    For i = 1 To Len(fileContent)
        Select Case Mid$(fileContent, i, 1)
            Case "^": y = y - 1
            Case "v": y = y + 1
            Case "<": x = x - 1
            Case ">": x = x + 1
        End Select
        
        ' sprawdz czy trzeba powiekszyc tablice
        Call EnsureInBounds(grid, rows, cols, x, y)
        
        ' zwieksz licznik w nowej pozycji
        grid(y, x) = grid(y, x) + 1
    Next i
    
    ' drukuj tablice
    Call PrintGrid(grid, rows, cols)
    
    Dim liczbaPol As Long
    liczbaPol = CountVisited(grid, rows, cols)
    Debug.Print "Liczba p�l odwiedzonych (warto�� > 0): "; liczbaPol
End Sub

' --- pomocnicze ---

Private Sub EnsureInBounds(ByRef grid() As Long, _
                           ByRef rows As Long, ByRef cols As Long, _
                           ByRef x As Long, ByRef y As Long)
    Dim newRows As Long, newCols As Long
    Dim dx As Long, dy As Long
    Dim tmp() As Long
    Dim i As Long, j As Long
    
    newRows = rows
    newCols = cols
    dx = 0: dy = 0
    
    If x < 0 Then
        newCols = newCols + 1
        dx = 1
        x = x + 1
    ElseIf x >= cols Then
        newCols = newCols + 1
    End If
    
    If y < 0 Then
        newRows = newRows + 1
        dy = 1
        y = y + 1
    ElseIf y >= rows Then
        newRows = newRows + 1
    End If
    
    If newRows <> rows Or newCols <> cols Then
        ReDim tmp(0 To newRows - 1, 0 To newCols - 1)
        
        For i = 0 To rows - 1
            For j = 0 To cols - 1
                tmp(i + dy, j + dx) = grid(i, j)
            Next j
        Next i
        
        grid = tmp
        rows = newRows
        cols = newCols
    End If
End Sub

Private Sub PrintGrid(grid() As Long, rows As Long, cols As Long)
    Dim i As Long, j As Long
    Debug.Print "Mapa " & rows & "x" & cols
    For i = 0 To rows - 1
        Dim line As String
        line = ""
        For j = 0 To cols - 1
            line = line & grid(i, j) & vbTab
        Next j
        Debug.Print line
    Next i
End Sub

Private Function CountVisited(grid() As Long, rows As Long, cols As Long) As Long
    Dim i As Long, j As Long, cnt As Long
    cnt = 0
    For i = 0 To rows - 1
        For j = 0 To cols - 1
            If grid(i, j) > 0 Then cnt = cnt + 1
        Next j
    Next i
    CountVisited = cnt
End Function
