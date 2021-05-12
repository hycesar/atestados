Attribute VB_Name = "mdRibbon"
Option Compare Database
Public Function mdRibbon()
Dim rsRib As DAO.Recordset
On Error GoTo TrataErro
'-----------------------------------------------------------------
'Esta função carrega as ribbons armazenadas na tabela tblRibbons,
'que deve ser chamada pela macro autoexec
'
'Crie a macro autoexec, selecione a ação EXECUTARCÓDIGO
'e escreva o nome da função no argumento: fncCarregaRibbon()
'------------------------------------------------------------------
Set rsRib = CurrentDb.OpenRecordset("tbRibbon", dbOpenDynaset)
Do While Not rsRib.EOF
  Application.LoadCustomUI rsRib!RibbonName, rsRib!RibbonXml
  rsRib.MoveNext
Loop
rsRib.Close
Set rsRib = Nothing

Sair:
  Exit Function
TrataErro:
  Select Case Err.Number
    Case 3078
      MsgBox "Tabela não encontrada...", vbInformation, "Aviso"
    Case Else
      MsgBox "Erro: " & Err.Number & vbCrLf & Err.description, _
      vbCritical, "Aviso", Err.HelpFile, Err.HelpContext
  End Select
  Resume Sair:
End Function

