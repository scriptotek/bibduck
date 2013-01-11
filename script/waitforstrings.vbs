' Bugs i SNetTerms JScript-støtte:
' - WaitForString støttet ikke timeout-parameteren, ble alltid 30 s
'  - WaitForStrings funket ikke i det hele tatt, taklet antakelig ikke javascript-arrays
' Derfor en liten VBSCript snutt

Function VBWaitForStrings(snt, str)
    str = Split(str, "|")
    VBWaitForStrings = snt.WaitForStrings(str, 10)
End Function
