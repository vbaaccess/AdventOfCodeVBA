Option Compare Database
Option Explicit

'modSymulacja

Public gra As New clsSymulacja

Public Sub Symulacja_DodajPostac1()
    Dim objPostac As New clsPostac
    
    objPostac.Nazwa = "Santa"
    objPostac.Inicjal = "S"
    objPostac.Inicjatywa = 5
    
    gra.DodajPostac objPostac
End Sub

Public Sub Symulacja_DodajPostac2()
    Dim objPostac As New clsPostac
    
    objPostac.Nazwa = "Robot"
    objPostac.Inicjal = "R"
    objPostac.Inicjatywa = 5
    
    gra.DodajPostac objPostac
End Sub

Public Sub Symulacja_WgrajTrase()
    gra.filePath = "C:\Project\AdventOfCodeVBA\Input.txt"
    gra.LoadFile
End Sub

Sub TestSymulacja()
    gra.WyswietlPostacie
    gra.RozpocznijSpacer
    gra.WyswietlMape
    Debug.Print "Odwiedzone pola (>0): "; gra.PoliczOdwiedzone
End Sub
