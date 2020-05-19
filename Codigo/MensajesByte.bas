Attribute VB_Name = "MensajesByte"
Option Explicit
Sub IntAByte(Inte As Integer, ByRef Bait() As Byte)

ReDim Bait(1 To 2) As Byte

Bait(1) = Inte \ 256
Bait(2) = Inte - (Inte \ 256)

End Sub
Sub LongAByte(Lon As Long, ByRef Bait() As Byte)

ReDim Bait(1 To 4) As Byte

Bait(1) = Inte \ 256
Bait(2) = Inte - (Inte \ 256)
Bait(3) = Inte \ 256
Bait(4) = Inte - (Inte \ 256)

End Sub
