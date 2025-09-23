Option Compare Database
Option Explicit

'clsFiles
Private m_eTypDanych As Database.EnumFiles

Private m_sDataArr() As String


Public Property Get DataArr() As String()

    DataArr = m_sDataArr

End Property

Public Property Let DataArr(sNewValue() As String)

    m_sDataArr = sNewValue

End Property

Public Property Get TypDanych() As Database.EnumFiles

    TypDanych = m_eTypDanych

End Property

Public Property Let TypDanych(ByVal eNewValue As Database.EnumFiles)

    m_eTypDanych = eNewValue

End Property


Public Function LoadFileToArray(filePath As String, TypDanych As Database.EnumFiles) As Variant
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
        
        If TypDanych = Znaki Then
            ' Utworz tablice znakow
            ReDim arr(1 To Len(content))
            For i = 1 To Len(content)
                arr(i) = Mid$(content, i, 1)
            Next i
        End If
        
        If TypDanych = Wiersze Then
            ' rozdzielenie na linie
            'Dim lines() As String
            arr() = Split(content, vbCrLf)
        End If
        
        If IsArray(arr) Then
            m_sDataArr = arr
        End If
        
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

Public Sub PrintArrayData()
    
    Dim v
    
    For Each v In m_sDataArr
        Debug.Print CStr(v)
        DoEvents
    Next v
End Sub
