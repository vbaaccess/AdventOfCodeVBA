' Klasa: clsLokacja
Option Compare Database
Option Explicit


Private mX As Long
Private mY As Long
Private mVisits As Long

Public Sub Init(ByVal x As Long, ByVal y As Long)
    mX = x
    mY = y
    mVisits = 0
End Sub

Public Property Let x(ByVal v As Long)
    mX = v
End Property
Public Property Get x() As Long
    x = mX
End Property

Public Property Let y(ByVal v As Long)
    mY = v
End Property
Public Property Get y() As Long
    y = mY
End Property

Public Property Get Visits() As Long
    Visits = mVisits
End Property

Public Property Let Visits(ByVal v As Long)
    mVisits = v
End Property

Public Sub IncVisits(Optional ByVal n As Long = 1)
    mVisits = mVisits + n
End Sub
