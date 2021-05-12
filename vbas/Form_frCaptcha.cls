VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frCaptcha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Ciente_Click()
    If Me.out_desafio.Caption = Me.in_desafio.Value Then
        DoCmd.Close
    Else
        MsgBox "Atenção: Os números devem ser iguais"
    End If
End Sub

Private Sub Form_load()
    If IsNull(Me.OpenArgs) Then
        Me.info.Caption = "Você está ciente?"
    Else
        Me.info.Caption = Me.OpenArgs
    End If
    
    'Inicializar gerador de números aleatórios
    Randomize
    
    'Inserir desafio
    Me.out_desafio.Caption = CStr(Int(((99 - 10 + 1) * Rnd) + 10))
    
    'Limpar campo
    Me.in_desafio.Value = ""
End Sub
