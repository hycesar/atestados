Attribute VB_Name = "mdSecurity"
Option Compare Database

Public Sub SetStartupOptions(propname As String, propdb As Variant, prop As Variant)
  'Set passed startup property.
  'Call SetStartupOptions("AllowBypassKey", dbBoolean, False)
  'Call SetStartupOptions("StartupShowDBWindow", dbBoolean, False)
  Dim dbs As Object, prp As Object
  Set dbs = CurrentDb
  On Error Resume Next
  dbs.Properties(propname) = prop
  If Err.Number = 3270 Then
    Set prp = dbs.CreateProperty(propname, propdb, prop)
    dbs.Properties.Append prp
  End If
  Set dbs = Nothing
  Set prp = Nothing
End Sub

' Go to Tools -> References... and check "Microsoft Scripting Runtime" to be able to use
' the FileSystemObject which has many useful features for handling files and folders
Public Function SaveTextToFile(msg As String, filePath As String) As Boolean
    ' The advantage of correctly typing fso as FileSystemObject is to make autocompletion
    ' (Intellisense) work, which helps you avoid typos and lets you discover other useful methods
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileStream As TextStream

    'get the directory path
    Folder = Left(filePath, InStrRev(filePath, "\"))
    
    'Is there a folder? So, create one
    If Len(Dir(Folder, vbDirectory)) = 0 Then
        MkDir Folder
    End If
    
    'Here the actual file is created and opened for write access
    Set fileStream = fso.CreateTextFile(filePath, overwrite:=True)
    
    ' Write something to the file
    If Len(msg) = 0 Then
        fileStream.WriteLine "null"
    Else
        fileStream.WriteLine msg
    End If

    ' Close it, so it's not locked anymore
    fileStream.Close

    'Checks if a file exists: sucess
    If fso.FileExists(filePath) Then
        SaveTextToFile = True
    Else
        SaveTextToFile = False
    End If
    
    ' Explicitly setting objects to Nothing should not be necessary in most cases, but if
    ' you're writing macros for Microsoft Access, you may want to uncomment the following
    ' two lines (see https://stackoverflow.com/a/517202/2822719 for details):
    Set fileStream = Nothing
    Set fso = Nothing
End Function

' Go to Tools -> References... and check "Microsoft Scripting Runtime" to be able to use
' the FileSystemObject which has many useful features for handling files and folders
Public Function ReadFileToText(filePath As String) As String
    On Error GoTo ReadFileToText_Err
    
    Dim fso As FileSystemObject
    Set fso = New FileSystemObject
    Dim fileStream As TextStream

    If fso.FileExists(filePath) Then
        ' Here the actual file is created and opened for read access
        Set fileStream = fso.OpenTextFile(filePath, IOMode:=ForReading)
        
        ' Write something to the file
        ReadFileToText = fileStream.ReadLine
    
        ' Close it, so it is not locked anymore
        fileStream.Close
        

    End If
    'Garbage colector
    Set fileStream = Nothing
    Set fso = Nothing

ReadFileToText_Exit:
    Exit Function
    
ReadFileToText_Err:
    MsgBox Error$, vbCritical, GENERALTTL & FATALTTL
    Cancel = False
    Resume ReadFileToText_Exit
End Function

Function ap_DisableShift()
    'This function disable the shift at startup. This action causes
    'the Autoexec macro and Startup properties to always be executed.
    
    On Error GoTo errDisableShift
    
    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270
    
    Set db = CurrentDb()
    
    'This next line disables the shift key on startup.
    db.Properties("AllowByPassKey") = False
    
    'The function is successful.
    Exit Function
    
errDisableShift:
    'The first part of this error routine creates the "AllowByPassKey
    'property if it does not exist.
    If Err = conPropNotFound Then
        Set prop = db.CreateProperty("AllowByPassKey", _
        dbBoolean, False)
        db.Properties.Append prop
        Resume Next
    Else
        MsgBox "Function 'ap_DisableShift' did not complete successfully."
        Exit Function
    End If
End Function

Function ap_EnableShift()
    'This function enables the SHIFT key at startup. This action causes
    'the Autoexec macro and the Startup properties to be bypassed
    'if the user holds down the SHIFT key when the user opens the database.
    
    On Error GoTo errEnableShift
    
    Dim db As DAO.Database
    Dim prop As DAO.Property
    Const conPropNotFound = 3270
    
    Set db = CurrentDb()
    
    'This next line of code disables the SHIFT key on startup.
    db.Properties("AllowByPassKey") = True
    
    'function successful
    Exit Function
    
errEnableShift:
    'The first part of this error routine creates the "AllowByPassKey
    'property if it does not exist.
    If Err = conPropNotFound Then
        Set prop = db.CreateProperty("AllowByPassKey", _
        dbBoolean, True)
        db.Properties.Append prop
        Resume Next
    Else
        MsgBox "Function 'ap_DisableShift' did not complete successfully."
        Exit Function
    End If
End Function

Sub relink()
    Dim strDbFile As String
    Dim strConnect As String
    
    strDbFile = mdSecurity.ReadFileToText("C:\temp\dba.txt")
    If Len(strDbFile) = 0 Then
        strDbFile = cmdFileDialog_Click
    End If
    
    strConnect = "MS Access;PWD=" & "3,14159265358979323" & ";DATABASE=" & strDbFile
    
    Dim tdf As DAO.TableDef
    Dim db As DAO.Database

    Set db = CurrentDb

    For Each tdf In db.TableDefs
        ' ignore system and temp tables
        If Not (tdf.Name Like "MSys*" Or tdf.Name Like "~*" Or tdf.Name Like "exl*" Or tdf.Name Like "USys*") Then
            tdf.Connect = strConnect
            tdf.RefreshLink
        End If
    Next
End Sub


Private Function cmdFileDialog_Click() As String
   ' Requires reference to Microsoft Office 11.0 Object Library.
   Dim fDialog As Office.FileDialog
   Dim objFSO As New FileSystemObject
 
   ' Set up the File Dialog.
   Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
 
   With fDialog
 
      ' Allow user to make multiple selections in dialog box
      .AllowMultiSelect = False
             
      ' Set the title of the dialog box.
      .title = "Por favor, corrigir o caminho da sua instalação!"
 
      ' Clear out the current filters, and add our own.
      .Filters.Clear
      .Filters.Add "Access Databases", "*.accdb"
 
      ' Show the dialog box. If the .Show method returns True, the
      ' user picked at least one file. If the .Show method returns
      ' False, the user clicked Cancel.
      If .Show = True Then
        mdSecurity.SaveTextToFile .SelectedItems(1), "C:\temp\dba.txt"
        cmdFileDialog_Click = .SelectedItems(1)
      Else
         MsgBox "Impossível continuar sem corrigir sua instalação!"
      End If
   End With
End Function
