Attribute VB_Name = "ModFunction"
Function getBaris(Sht As Worksheet, Kl As String, Optional Baru As Boolean = False) As Long
Rem Auth: andi Setiadi
getBaris = Sht.Range(Kl & Rows.Count).End(3).Row
If Baru Then IncL getBaris
End Function

Function IncL(ByRef X As Long) As Long
Rem Increment Long
IncL = X + 1
'tambah update mas imam
'tambah lagi 13/7/2020
End Function

Function IncI(ByRef X As Integer) As Integer
Rem Increment untuk Integer
IncI = X + 1
End Function

