Option Compare Database
Option Explicit

Const WlascicielPinu = "pablo"
Const PoziomSzyfrowania = "e2ea7b" '6 zer na poczatku md5

'--- modMD5: tworzeni i szukanie uktytego PINu ---
'Szukamy md5 dla pinu wlasciciel ktore polega na utworzeniu opisu pinu np.: pablo123456
'pod warunkiem ze md5 dla ciagu pablo123456 ma na poczatku taki ciag jak w zmiennej PozimSyzfrowani
'np. "000000" 6 zer
'
'manipulujac     iStart ora iMax mozey wyswietlic jakis zakres danych szyfrowych dla bardzo duzeo pinu (dowolna ilosc cyf
'                i zapisac PoziomSzyfrowania ktory bedzie okreslal nasz pin dla danego wlasciciel
'--- ------------------------------ ---

' Funkcja obliczająca hash MD5 dla ciągu tekstowego
Public Function MD5(ByVal sText As String) As String
    Dim enc As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Long
    
    ' Tworzymy provider MD5 z .NET
    Set enc = CreateObject("System.Security.Cryptography.MD5CryptoServiceProvider")
    
    ' Zamiana stringa na tablicę bajtów (ANSI / UTF-8 zależnie od potrzeby)
    bytes = StrConv(sText, vbFromUnicode)
    
    ' Obliczenie hash
    hash = enc.ComputeHash_2((bytes))
    
    ' Konwersja bajtów na hex
    MD5 = ""
    For i = 0 To UBound(hash)
        MD5 = MD5 & LCase(Right$("0" & Hex$(hash(i)), 2))
    Next i
End Function

' Funkcja obliczająca hash SHA-1 dla ciągu tekstowego
Public Function SHA1(ByVal sText As String) As String
    Dim enc As Object
    Dim bytes() As Byte
    Dim hash() As Byte
    Dim i As Long
    
    ' Tworzymy provider MD5 z .NET
    Set enc = CreateObject("System.Security.Cryptography.SHA1CryptoServiceProvider")
    
    ' Zamiana stringa na tablicę bajtów (ANSI / UTF-8 zależnie od potrzeby)
    bytes = StrConv(sText, vbFromUnicode)
    
    ' Obliczenie hash
    hash = enc.ComputeHash_2((bytes))
    
    ' Konwersja bajtów na hex
    SHA1 = ""
    For i = 0 To UBound(hash)
        SHA1 = SHA1 & LCase(Right$("0" & Hex$(hash(i)), 2))
    Next i
End Function

Public Sub Md5Finder()
    Dim iLicznik&, iStart&, iMax&, iTotalSteps&
    Dim sKlucz$, sPrefix$, sMD5$
    Dim searchedHexStart$
    
    Dim stepSize As Double
    Dim nextPercent As Long
    Dim currentPercent As Long
    
    Dim startTime As Date, endTime As Date
    
    startTime = Now()
    
    sPrefix =  WlascicielPinu
    searchedHexStart = PoziomSzyfrowania
   
    iStart = 1
    iMax = 1250000
    
    iStart = 2050123
    iMax = 5200500
    
    iStart = 2050123
    iMax = 5200500
    
    iTotalSteps = iMax - iStart + 1
    
    Debug.Print "Sprawdzam od " & iStart & " do " & iMax & " (" & iTotalSteps & " sprawdzeń)"
    Debug.Print "Start: " & Format(startTime, "yyyy-mm-dd hh:nn:ss")
    
    iLicznik = iStart
    Do
        
        ' oblicz aktualny procent (zaokraglam w dol)
        currentPercent = Int((iLicznik - iStart + 1) / iTotalSteps * 100)
        
        sKlucz = sPrefix & CStr(iLicznik)
        sMD5 = MD5(sKlucz)
       'Debug.Print "sKlucz: " & sKlucz & " =MD5> " & sMD5
        
        '--- Wyswietlanie % postepu ---
            If currentPercent >= nextPercent Then
                Debug.Print currentPercent & "% ukończono (" & (iLicznik - iStart + 1) & " sprawdzeń)";
                Debug.Print "sKlucz: " & sKlucz & " =MD5> " & sMD5
                nextPercent = nextPercent + 1
            End If
        '--- --- --- --- --- --- --- ---
        
        If Mid(sMD5, 1, Len(searchedHexStart)) = searchedHexStart Then
            Debug.Print "!) Secret key => " & iLicznik
            Exit Do
        End If
        If iLicznik = iMax Then
            Debug.Print ":( Secret key not found"
            Exit Do
        End If
        
        iLicznik = iLicznik + 1
        DoEvents
    Loop
    
    endTime = Now()
    Debug.Print "Start:   " & Format(startTime, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Koniec:  " & Format(endTime, "yyyy-mm-dd hh:nn:ss")
    Debug.Print "Trwało:  " & Format(endTime - startTime, "hh:nn:ss")
    
End Sub
