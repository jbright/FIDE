Attribute VB_Name = "Settings"

'******************************************************************************
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' The Compiler location
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get CompilerPath() As String
    CompilerPath = GetSetting(App.Title, "Settings", "Compiler", App.Path & "\Internal\9lc.exe")
End Property
Public Property Let CompilerPath(strNewValue As String)
    SaveSetting App.Title, "Settings", "Compiler", strNewValue
End Property


' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
' The Trash folder location
' * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Property Get ProjectTrashPath() As String
    ProjectTrashPath = GetSetting(App.Title, "Settings", "Trash", App.Path & "\Trash")
End Property
Public Property Let ProjectTrashPath(strNewValue As String)
    SaveSetting App.Title, "Settings", "ProjectPath", strNewValue
End Property

