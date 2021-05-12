VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frUserSignIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_Unload(Cancel As Integer)
    If mdSecurity.ReadFileToText("C:\temp\log.txt") = "start" Then
        If MsgBox("Finalizar programa?", vbQuestion + vbYesNo + vbDefaultButton2, GENERALTTL & QUERYTTL) = vbYes Then
            mdSecurity.SaveTextToFile "end", "C:\temp\log.txt"
            DoCmd.Quit
        Else
            Cancel = True
        End If
    End If
End Sub

Private Sub Form_open(Cancel As Integer)
    On Error GoTo Form_load_Err
    msg = mdSecurity.ReadFileToText("C:\temp\log.txt")
    
    mdSecurity.relink
    
    'check if it's in a current session
    If Not IsNumeric(msg) Then
        If Not mdSecurity.SaveTextToFile("start", "C:\temp\log.txt") Then
            MsgBox "Falha em acessar sistema de arquivos. Verifique instalação", vbExclamation, GENERALTTL & WARNTTL
        End If
    End If

Form_load_Exit:
    Exit Sub
    
Form_load_Err:
    MsgBox Error$, vbCritical, GENERALTTL & FATALTTL
    Cancel = False
    Resume Form_load_Exit
End Sub

'------------------------------------------------------------
' Form_frUserLogin: Verificar login
' Sumary: Valida campos, hash de senha e permitir acesso
'------------------------------------------------------------
Private Sub userAcessar_Click()
    On Error GoTo userAcessar_Click_Err

userAcessar_Click_Restart:
    If Not IsNull(Me.pswd.Value) And Not IsNull(Me.matr.Value) Then
        'Guarda valor do sal de hash, objeto da classe HashMD5 e uma string
        Dim salt As Variant, objMD5 As New clHashMD5, hashPswd As String
        
        'procurar na coluna 'salt', na tabela 'tbUser', o valor de sal
        salt = DLookup("[sal]", "tbUsuario", "[matricula] = " & Me.matr.Value)
        
        If Not IsNull(salt) Then
            'Calcular hash e guardar na string
            hashPswd = objMD5.DigestStrToHexStr(pswd + salt)

            'Verificar senha
            If (hashPswd = DLookup("[senha]", "tbUsuario", "[matricula] =" & Me.matr.Value)) Then
                'Grava arquivo
                retorno = mdSecurity.SaveTextToFile(Me.matr.Value, "C:\temp\log.txt")
                
                'Fecha formulario
                DoCmd.Close

                'Mensagem de confirmação
                MsgBox "Acesso efetuado com sucesso.", vbInformation, GENERALTTL & INFOTTL
                
                'Abre Formulário
                DoCmd.OpenForm "LAFEPE"
            Else
                'Mensagem de Erro dependente do usuário
                MsgBox "Senha incorreta, tente novamente.", vbExclamation, GENERALTTL & WARNTTL
            End If
        Else
            'Mensagem de Erro dependente do usuário
            MsgBox "Usuário incorreto, tente novamente.", vbExclamation, GENERALTTL & WARNTTL
        End If
    Else
        'Mensagem de Erro dependente do usuário
        MsgBox "Favor, digitar usuário e senha.", vbExclamation, GENERALTTL & WARNTTL
    End If
    
userAcessar_Click_Exit:
    Exit Sub

userAcessar_Click_Err:
    If Err.Number = 3044 Or Err.Number = 3024 Then
        RunCommand acCmdLinkedTableManager
        Resume userAcessar_Click_Restart
    Else
        MsgBox Error$, vbCritical, GENERALTTL & FATALTTL
        Resume userAcessar_Click_Exit
    End If
End Sub
