Option Compare Database
Option Explicit

'modAnalizaTekstu

Public Sub AnalizaTekstu()
    Dim objF As New clsFiles
    Dim sFilePath$
    Dim arrDataFromFile As Variant

    sFilePath = "C:\Project\AnalizowanyPlik.txt"
    Call objF.LoadFileToArray(sFilePath, Wiersze)
    arrDataFromFile = objF.DataArr
   'objF.PrintArrayData
    
    'Analiza
    
    Dim iNice&, bNice As Boolean
    Dim v, sWer$, iWer&
    
    iNice = 0
    
    Dim chars() As String
    
    For Each v In arrDataFromFile
        
        sWer = ""
        bNice = True
        
        Debug.Print CStr(v)
        
        '--- rule 1 - vowels
        'If ZasadaNr01(CStr(v)) = False Then bNice = False
        '--- rule 2
        'If ZasadaNr02(CStr(v)) = False Then bNice = False
        '--- rule 3
        'If ZasadaNr03(CStr(v)) = False Then bNice = False
        
        '--- rule 4
        If ZasadaNr04(CStr(v)) = False Then bNice = False
        
        '--- rule 5
        If ZasadaNr05(CStr(v)) = False Then bNice = False
        
        If bNice Then iNice = iNice + 1
        
        DoEvents
    Next v
    
    Debug.Print "Nice count = " & iNice
   'arrDataFromFile = objF.DataArr
    
End Sub

Private Function ZasadaNr01(sWeryfikowanyCiag As String) As Boolean
Const VOLVES = "aeiou"
    Dim bBrak As Boolean
    Dim sWer$, iWer&, vo
    Dim arrVowels
    
    arrVowels = Split(StrConv(VOLVES, vbUnicode), Chr(0))
    sWer = sWeryfikowanyCiag
    
    For Each vo In arrVowels
        'Debug.Print vo
        sWer = Replace(sWer, CStr(vo), "")
    Next vo
    iWer = Len(sWeryfikowanyCiag) - Len(sWer)
    If iWer < 3 Then bBrak = True
    
    ZasadaNr01 = Not (bBrak)
End Function

Private Function ZasadaNr02(sWeryfikowanyCiag As String) As Boolean
    Dim bBrak As Boolean
    Dim sWer$, iWer&
    Dim ln!, zLp!
    
    sWer = sWeryfikowanyCiag
    ln = Len(sWeryfikowanyCiag)
    
    For zLp = 1 To ln
        sWer = Mid$(sWeryfikowanyCiag, zLp, 2)
        
        If Mid$(sWer, 1, 1) = Mid$(sWer, 2, 1) Then iWer = iWer + 1
        
       'Debug.Print "1:" & Mid$(sWer, 1, 1) & " =?= " & "2:" & Mid$(sWer, 2, 1)
        
        'Debug.Print sWer
        'Debug.Print "1:" & Mid$(sWer, 1, 1)
        'Debug.Print "2:" & Mid$(sWer, 2, 1)
    Next zLp
    If iWer < 1 Then bBrak = False
    
    ZasadaNr02 = Not (bBrak)
End Function

Private Function ZasadaNr03(sWeryfikowanyCiag As String) As Boolean
Const DUBLE_CHR = "abcdpqxy"
    Dim arrDubleChar, vo
    Dim bBrak As Boolean
    
    arrDubleChar = SplitPairs(DUBLE_CHR)
        
    For Each vo In arrDubleChar
        'Debug.Print vo
        If InStr(1, sWeryfikowanyCiag, CStr(vo)) > 0 Then
            bBrak = True
            Exit For
        End If
    Next vo
    
    ZasadaNr03 = Not (bBrak)
End Function

Private Function ZasadaNr04(sWeryfikowanyCiag As String) As Boolean
    Dim zLp!
    Dim sWer(1 To 3) As String
    Dim bOK As Boolean
    
    For zLp = 1 To Len(sWeryfikowanyCiag)
        sWer(1) = Mid$(sWeryfikowanyCiag, 1, zLp - 1)
        sWer(2) = Mid$(sWeryfikowanyCiag, zLp, 2)
        sWer(3) = Mid$(sWeryfikowanyCiag, zLp + 2)
        
        If Len(sWer(2)) = 2 Then
            If InStr(sWer(1), sWer(2)) > 0 Then bOK = True
            If InStr(sWer(3), sWer(2)) > 0 Then bOK = True
        End If
        
        If bOK Then Exit For
        'Debug.Print "1 prev: " & sWer(1)
        'Debug.Print "2 s:    " & sWer(2)
        'Debug.Print "3 after:" & sWer(3)
        
    Next zLp
    ZasadaNr04 = bOK
End Function

Private Function ZasadaNr05(sWeryfikowanyCiag As String) As Boolean
    Dim bOK As Boolean
    Dim sWer$
    Dim zLp!
    
    sWer = sWeryfikowanyCiag
    
    For zLp = 1 To Len(sWeryfikowanyCiag) - 2
        sWer = Mid$(sWeryfikowanyCiag, zLp, 3)
        
        If Mid$(sWer, 1, 1) = Mid$(sWer, 3, 1) Then
            'iWer = iWer + 1
            bOK = True
            Exit For
        End If

    Next zLp
    
    ZasadaNr05 = bOK
End Function


Function SplitPairs(s As String) As Variant
    Dim arr() As String
    Dim i As Long, n As Long
    
    n = Len(s) \ 2   ' liczba par (dzielenie calkowite)
    ReDim arr(1 To n)
    
    For i = 1 To n
        arr(i) = Mid$(s, (i - 1) * 2 + 1, 2)
    Next i
    
    SplitPairs = arr
End Function

