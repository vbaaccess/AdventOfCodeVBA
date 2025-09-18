' Klasa: clsPostac
Option Compare Database
Option Explicit


Private mNazwa As String
Private mInicjal As String
Private mInicjatywa As Integer
Private mX As Long
Private mY As Long

Public Property Let Nazwa(ByVal s As String): mNazwa = s: End Property
Public Property Get Nazwa() As String: Nazwa = mNazwa: End Property

Public Property Let Inicjal(ByVal s As String): mInicjal = s: End Property
Public Property Get Inicjal() As String: Inicjal = mInicjal: End Property

Public Property Let Inicjatywa(ByVal v As Integer): mInicjatywa = v: End Property
Public Property Get Inicjatywa() As Integer: Inicjatywa = mInicjatywa: End Property

Public Property Let x(ByVal v As Long): mX = v: End Property
Public Property Get x() As Long: x = mX: End Property

Public Property Let y(ByVal v As Long): mY = v: End Property
Public Property Get y() As Long: y = mY: End Property

Public Sub ResetPozycja()
    mX = 0
    mY = 0
End Sub
