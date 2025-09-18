Option Compare Database
Option Explicit

Private mFilePath As String
Private mPostacie As Collection
Private mDataFromFile As Variant

' grid(row, col) -> row = indeks Y, col = indeks X
Private grid() As Variant
Private rows As Long, cols As Long

' offsety: index = coord + offset
Private offsetX As Long, offsetY As Long

Public Property Let filePath(ByVal s As String)
    mFilePath = s
End Property

Public Property Get filePath() As String
    filePath = mFilePath
End Property

Private Sub Class_Initialize()
    Set mPostacie = New Collection
    rows = 0: cols = 0
    offsetX = 0: offsetY = 0
End Sub

Public Sub DodajPostac(p As clsPostac)
    If mPostacie Is Nothing Then Set mPostacie = New Collection
    mPostacie.Add p
End Sub

Public Sub WyswietlPostacie()
    Dim p As clsPostac
    Debug.Print "Postacie (" & IIf(mPostacie Is Nothing, 0, mPostacie.Count) & "):"
    If Not mPostacie Is Nothing Then
        For Each p In mPostacie
            Debug.Print " - " & p.Nazwa & " (" & p.Inicjal & "), Inicjatywa=" & p.Inicjatywa & _
                        "  Pos=(" & p.x & "," & p.y & ")"
        Next p
    End If
End Sub

Private Sub InitGridIfNeeded()
    If rows = 0 Or cols = 0 Then
        rows = 1: cols = 1
        ReDim grid(0 To rows - 1, 0 To cols - 1)
        offsetX = 0: offsetY = 0
        ' utwórz lokację (0,0)
        Dim loc As clsLokacja
        Set loc = New clsLokacja
        loc.Init 0, 0
        Set grid(0, 0) = loc
    End If
End Sub

Private Function GetLocationByCoord(ByVal x As Long, ByVal y As Long) As clsLokacja
    InitGridIfNeeded
    
    ' sprawdz granice, ewentualnie rozszerz siatke
    EnsureInBoundsForCoord x, y
    
    Dim ix As Long, iy As Long
    ix = x + offsetX
    iy = y + offsetY
    
    Dim tmp As clsLokacja
    On Error Resume Next
    Set tmp = grid(iy, ix)
    On Error GoTo 0
    If tmp Is Nothing Then
        Set tmp = New clsLokacja
        tmp.Init x, y
        Set grid(iy, ix) = tmp
    End If
    Set GetLocationByCoord = tmp
End Function

' --- rozszerza grid tak by wsp, (x,y) miescily sie w tablicy ---
Private Sub EnsureInBoundsForCoord(ByVal x As Long, ByVal y As Long)
    ' oblicz indeksy przy obecnych offsetach
    Dim ix As Long, iy As Long
    ix = x + offsetX
    iy = y + offsetY
    
    Dim needLeft As Long, needRight As Long, needTop As Long, needBottom As Long
    needLeft = 0: needRight = 0: needTop = 0: needBottom = 0
    
    If rows = 0 Or cols = 0 Then
        ' init
        rows = 1: cols = 1
        ReDim grid(0 To 0, 0 To 0)
        offsetX = 0: offsetY = 0
        ix = x + offsetX: iy = y + offsetY
    End If
    
    If ix < 0 Then needLeft = -ix
    If ix > cols - 1 Then needRight = ix - (cols - 1)
    If iy < 0 Then needTop = -iy
    If iy > rows - 1 Then needBottom = iy - (rows - 1)
    
    
    If needLeft + needRight + needTop + needBottom = 0 Then Exit Sub ' mieszczymy się
    
    Dim newCols As Long, newRows As Long
    newCols = cols + needLeft + needRight
    newRows = rows + needTop + needBottom
    
    Dim tmp() As Variant
    ReDim tmp(0 To newRows - 1, 0 To newCols - 1)
    
    Dim r As Long, c As Long
    ' skopiuj istniejace elementy do przesunietej pozycji
    For r = 0 To rows - 1
        For c = 0 To cols - 1
            On Error Resume Next
            If Not (IsEmpty(grid(r, c))) Then
                Set tmp(r + needTop, c + needLeft) = grid(r, c)
            End If
            On Error GoTo 0
        Next c
    Next r
    
    ' ustaw nowa siatke i offsety
    grid = tmp
    offsetX = offsetX + needLeft
    offsetY = offsetY + needTop
    cols = newCols
    rows = newRows
    
    ' utowrz obiekty clsLokacja dla pustych pol (domyslnie Visits = 0)
    Dim loc As clsLokacja
    For r = 0 To rows - 1
        For c = 0 To cols - 1
            On Error Resume Next
            Set loc = grid(r, c)
            On Error GoTo 0
            If loc Is Nothing Then
                Set loc = New clsLokacja
                loc.Init (c - offsetX), (r - offsetY)
                Set grid(r, c) = loc
            End If
        Next c
    Next r
End Sub

Public Sub LoadFile()
    Dim data As Variant
    Dim i As Long
    
    data = LoadFileToArray(mFilePath)
    
    If IsArray(data) Then
        mDataFromFile = data
    Else
        MsgBox "Brak danych do wczytania."
    End If
End Sub

Public Sub PrintDataFromFile()
    Dim i&
    If IsArray(mDataFromFile) Then
        For i = LBound(mDataFromFile) To UBound(mDataFromFile)
            Debug.Print mDataFromFile(i)
            DoEvents
        Next i
    Else
        MsgBox "Brak danych do wyswietlenia."
    End If
End Sub

Public Function LoadFileToArray(filePath As String) As Variant
    Dim f As Integer
    Dim content As String
    Dim arr() As String
    Dim i As Long
    Dim fileLen As Long
    
    ' --- Sprawdz czy plik istnieje ---
    If Dir(filePath) = "" Then
        Exit Function   ' plik nie istnieje › funkcja zwroci Empty
    End If
    
    On Error GoTo ErrHandler
    
    f = FreeFile
    Open filePath For Binary As #f
    fileLen = LOF(f)   ' len pliku w bajtach
    
    If fileLen > 0 Then
        content = String(fileLen, vbNullChar)
        Get #f, , content
        Close #f
        
        ' Utworz tablice znakow
        ReDim arr(1 To Len(content))
        For i = 1 To Len(content)
            arr(i) = Mid$(content, i, 1)
        Next i
        
        LoadFileToArray = arr
    Else
        Close #f
        LoadFileToArray = Empty
    End If
    
    Exit Function
    
ErrHandler:
    On Error Resume Next
    Close #f
    LoadFileToArray = Empty
End Function

Private Function LoadAllFileToArray(filePath As String) As Variant
    Dim f As Integer
    Dim line As String
    Dim arr() As String
    Dim cnt As Long
    
    ' --- czy plik istnieje ---
    If Dir(filePath) = "" Then
        MsgBox "Plik nie istnieje: " & filePath, vbExclamation, "Błąd"
        Exit Function
    End If
    
    On Error GoTo ErrHandler
    
    f = FreeFile
    Open filePath For Input As #f
    
    cnt = 0
    Do Until EOF(f)
        Line Input #f, line
        cnt = cnt + 1
        ReDim Preserve arr(1 To cnt)
        arr(cnt) = line
    Loop
    
    Close #f
    
    ' Jesli plik pusty › arr nie zostanie zainicjalizowana
    If cnt > 0 Then
        LoadAllFileToArray = arr
    Else
        LoadAllFileToArray = Empty
    End If
    
    Exit Function
    
ErrHandler:
    MsgBox "Błąd przy odczycie pliku: " & Err.Description, vbCritical, "Błąd"
    On Error Resume Next
    Close #f
    LoadAllFileToArray = Empty
End Function

Public Sub RozpocznijSpacer()
Const MOV_DESC = " (oldX,oldY) => (newX,newY)"
    
    Dim f As Integer
    Dim fileContent As String
    Dim i As Long, moveIdx As Long, tura As Long
    Dim p As clsPostac
    Dim num As Long
    
    If mPostacie Is Nothing Or mPostacie.Count = 0 Then
        Err.Raise vbObjectError + 1, , "Brak dodanych postaci. Użyj DodajPostac."
    End If
    
    ' przygotuj siatke, zainicjuj startowa pozycje (0,0)
    InitGridIfNeeded

    ' reset pozycji postaci i wprowadz odwiedziny startowe
    For Each p In mPostacie
        p.ResetPozycja
        Dim startLoc As clsLokacja
        Set startLoc = GetLocationByCoord(0, 0)
        startLoc.IncVisits 1
    Next p

    'czytam wskazowki
    Dim znak$, v
    Dim oldX As Long, oldY As Long
    Dim newX As Long, newY As Long
    
    num = mPostacie.Count
    moveIdx = 0
    
    For Each v In mDataFromFile
        znak = CStr(v)
        If InStr("^v<>", znak) = 0 Then GoTo NextChar
        '--- START ---
            moveIdx = moveIdx + 1
            tura = ((moveIdx - 1) Mod num) + 1
            Set p = mPostacie(tura)
            
            oldX = p.x: newX = p.x
            oldY = p.y: newY = p.y
            
            Select Case znak
                Case "^": newY = oldY + 1     ' góra -> y zwiększa się
                Case "v": newY = oldY - 1     ' dół -> y zmniejsza się
                Case "<": newX = oldX - 1
                Case ">": newX = oldX + 1
            End Select
            
            ' pobierz/utworz lokacje, zwieksz ilosc odwiedzin (Visits) i ustaw nowa pozycje postaci
            Dim movDes$
            Dim loc As clsLokacja
            Set loc = GetLocationByCoord(newX, newY)
            
            loc.IncVisits 1
            p.x = newX
            p.y = newY
            
            movDes = MOV_DESC
            movDes = Replace(movDes, "oldX", oldX)
            movDes = Replace(movDes, "oldY", oldY)
            movDes = Replace(movDes, "newX", newX)
            movDes = Replace(movDes, "newY", newY)
            movDes = movDes & "(" & loc.Visits & ")"
            Debug.Print moveIdx & " : " & p.Nazwa & " znak: " & znak & movDes
        '--- END ---
NextChar:
    Next v

End Sub

Public Sub WyswietlMape()
    Dim r As Long, c As Long
    Debug.Print "Mapa " & rows & "x" & cols & "  (offsetX=" & offsetX & ", offsetY=" & offsetY & ")"
    For r = rows - 1 To 0 Step -1
        Dim line As String
        line = ""
        For c = 0 To cols - 1
            Dim loc As clsLokacja
                Set loc = Nothing
            On Error Resume Next
                Set loc = grid(r, c)
            On Error GoTo 0
            If Not loc Is Nothing Then
                line = line & loc.Visits & vbTab
            Else
                line = line & "0" & vbTab
            End If
        Next c
        Debug.Print line
    Next r
End Sub

Public Function PoliczOdwiedzone() As Long
    Dim r As Long, c As Long, cnt As Long
    cnt = 0
    If rows = 0 Or cols = 0 Then
        PoliczOdwiedzone = 0
        Exit Function
    End If
    Dim loc As clsLokacja
    For r = 0 To rows - 1
        For c = 0 To cols - 1
                Set loc = Nothing
            On Error Resume Next
                Set loc = grid(r, c)
            On Error GoTo 0
            If Not loc Is Nothing Then
                If loc.Visits > 0 Then cnt = cnt + 1
            End If
        Next c
    Next r
    PoliczOdwiedzone = cnt
End Function

Public Property Get PostaciCount() As Long
    If mPostacie Is Nothing Then PostaciCount = 0 Else PostaciCount = mPostacie.Count
End Property
