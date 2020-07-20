Attribute VB_Name = "Modandi"
Sub Initandi()
Debug.Print "Saya andi"
End Sub

Public  Sub testDariVS()
    MsgBox "hai, saya Andi",,"VS Editor" 
End Sub

Public  Function Hallo() as string
    hallo = "Hari ini adalah hari senin"
End Function

Public  Function Hitung( x as range, y as range) as double
    hitung = x.value * y.value
End Function