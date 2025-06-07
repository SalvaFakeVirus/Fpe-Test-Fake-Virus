Dim respuesta
Dim WshShell
Dim i
Dim urlImagen
Dim urlMeme
Dim cmdError

Set WshShell = CreateObject("WScript.Shell")

respuesta = MsgBox("Quieres ir a FPE o no?", vbYesNo + vbCritical, "FPE Decision")

If respuesta = vbNo Then
    urlImagen = "https://encrypted-tbn0.gstatic.com/images?q=tbn:ANd9GcT4gqu3bdawtMTlhrC39KEZr0ZkijDU098SNw&s"
    
    For i = 1 To 25
        WshShell.Run urlImagen, 1, False
    Next
    
    For i = 1 To 50
        ' Muestra ERROR 205 y se cierra en 1 segundo (sin mostrar timeout)
        cmdError = "cmd /c echo ERROR 205 & timeout /t 1 /nobreak >nul"
        WshShell.Run cmdError, 1, False
        ' Sin pausa entre lanzamientos para saturar rápido
    Next
    
    For i = 1 To 100
        MsgBox "Твой компьютер будет заражён", vbCritical + vbSystemModal, "Вирус!"
        ' Sin Sleep para saturar rápido
    Next

ElseIf respuesta = vbYes Then
    urlMeme = "https://media.tenor.com/Yn0DXNS8iKoAAAAe/fpe-oliver.png"
    WshShell.Run urlMeme, 1, False
End If
