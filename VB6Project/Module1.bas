Attribute VB_Name = "FIDEMainModule"


Public fMainForm As frmMain

' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' SUB: Main
' PURPOSE: Gets the ball rolling
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Sub Main()

    frmSplash.Show
    frmSplash.Refresh
    
    Set fMainForm = New frmMain
    Load fMainForm
    
    While Not frmSplash.TimeUp
        DoEvents
    Wend
    
    Unload frmSplash
    
    
    fMainForm.Show
End Sub




