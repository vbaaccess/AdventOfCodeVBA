''========================================================================================
'' Module : clsPixelsMatrix  | Type Class   : Class Module
'' Author : pablo           | Created: 2025-09-25
'' Purpose: Zaplanie macierzy pixeli
''----------------------------------------------------------------------------------------
'' Description : Procedura analityczna
''----------------------------------------------------------------------------------------
'' Dependencies:  (Obiekty, od ktorych modul zalezy)
''----------------------------------------------------------------------------------------
'' Change Log
''  2025-09-25   | Autor: pablo| Utworzenie
''
''========================================================================================

Option Compare Database
Option Explicit

Private mFilePath As String
Private mDataFromFile As Variant

' --- property ---
Public Property Let filePath(ByVal s As String)
    mFilePath = s
End Property

Public Property Get filePath() As String
    filePath = mFilePath
End Property

Public Sub LoadFile()
    Dim objF As New clsFiles
    Dim data As Variant
    Dim i As Long
    
    data = objF.LoadFileToArray(mFilePath, Wiersze)
    
    mDataFromFile = objF.DataArr
    objF.PrintArrayData
    
    If Not IsArray(mDataFromFile) Then
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
