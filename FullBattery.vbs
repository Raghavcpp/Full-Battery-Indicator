Set oLocator = CreateObject("WbemScripting.SWbemLocator")
Set oServices = oLocator.ConnectServer(".", "root\wmi")
Set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")

For Each oResult In oResults
    iFull = oResult.FullChargedCapacity
Next

soundFile = "C:\Users\H\Downloads\notification.wav"

Function PlaySound(file)
    Set soundPlayer = CreateObject("WMPlayer.OCX.7")
    soundPlayer.URL = file
    soundPlayer.settings.volume = 100 ' Set volume to maximum
    soundPlayer.Controls.play
    ' Wait for the sound to complete
    Do While soundPlayer.PlayState <> 1 ' 1 = Stopped
        WScript.Sleep 100
    Loop
    Set soundPlayer = Nothing
End Function

While (1)
    Set oResults = oServices.ExecQuery("select * from batterystatus")
    For Each oResult In oResults
        iRemaining = oResult.RemainingCapacity
        bCharging = oResult.Charging
    Next
    iPercent = ((iRemaining / iFull) * 100) Mod 100
    If bCharging And (iPercent > 98) Then
        PlaySound soundFile
        MsgBox "Battery is fully charged"
    End If
    WScript.Sleep 30000 ' 30 seconds
Wend
