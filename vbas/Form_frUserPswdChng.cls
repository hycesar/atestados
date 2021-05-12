VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frUserPswdChng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
'------------------------------------------------------------
' Form_frUser: Cadastrar Usuário do Sistema
' Sumary: Valida campos, hash de senha e guardar em tabela
'------------------------------------------------------------

Private Sub Form_load()
    Dim args As Variant
    args = Me.OpenArgs
    If IsNull(args) Then
        Me.matr.Value = mdSecurity.ReadFileToText("C:\temp\log.txt")
    Else
        Me.matr.Value = Me.OpenArgs
    End If
End Sub

Private Sub userAlterarSenha_Click()
    On Error GoTo userCadastrar_Click_Err

    'Declarando um objeto da classe HashMD5
    Dim objMD5 As New clHashMD5
    
    'Inicializar gerador de números aleatórios
    Randomize
    
    'Preencher formulário com salt
    Me.salt.Value = CStr(Int(((9999 - 1000 + 1) * Rnd) + 1000))
    
    If Not IsNull(Me.pswd.Value) Then
        If Me.pswd.Value = Me.pswdConf.Value Then
            'Calcular hash
            Me.pswd.Value = objMD5.DigestStrToHexStr(Me.pswdConf.Value + Me.salt.Value)
            
            'Preencher formulário
            Me.pswdConf.Value = Me.pswd.Value
            
            'Procurar entrada duplicada
            If IsNull(DLookup("[matricula]", "tbUsuario", "[matricula] = " & Me.matr.Value)) Then
                
                'SQL INSERT
                CurrentDb.Execute "INSERT tbUsuario SET matr = '" & Me.matr.Value & "', pswd = '" & Me.pswd.Value & _
                "', salt = '" & Me.salt.Value & "' WHERE matr = " & Me.matr.Value & "'"
                
                'Fechar Formulario
                DoCmd.Close
                
                'Mensagem de sucesso
                MsgBox "Usuário e senha cadastrados com sucesso.", vbInformation, GENERALTTL & INFOTTL
            
            Else
                'Mensagem de dúvida
                If (vbYes = MsgBox("Realmente deseja alterar sua senha?", _
                vbQuestion + vbYesNo + vbDefaultButton2, GENERALTTL & QUERYTTL)) Then
                    'Verificar se tem permissão
                    'If userMatr = Me.matr.Value Then
                        'SQL UPDATE
                        CurrentDb.Execute "UPDATE tbUsuario" _
                        & " SET senha = '" & Me.pswd.Value _
                        & "', sal = '" & Me.salt.Value _
                        & "' WHERE matricula = " & Me.matr.Value
                                                                        
                        'Fechar Formulario
                        DoCmd.Close
                        
                        'Mensagem de sucesso
                        MsgBox "Senha atualizada com sucesso.", vbInformation, GENERALTTL & INFOTTL
                    'Else
                        'Mensagem de Erro dependente do usuário
                    '    MsgBox "Usuário não tem permissão para alterar senha de outro.", vbExclamation, GENERALTTL & WARNTTL
                    'End If
                Else
                    
                    'Alertar usuário
                    MsgBox "Nenhuma operação foi executada.", vbCritical, GENERALTTL & ERRORTTL
                End If
            End If
        Else
            'Mensagem de Erro dependente do usuário
            MsgBox "Favor, digitar senhas iguais.", vbExclamation, GENERALTTL & WARNTTL
        End If
    Else
        MsgBox "Favor, digitar uma senha.", vbExclamation, GENERALTTL & WARNTTL
    End If

userCadastrar_Click_Exit:
    Exit Sub

userCadastrar_Click_Err:
    MsgBox Error$, vbCritical, GENERALTTL & FATALTTL
    'Limpar formulários
    Me.pswd.Value = ""
    Me.pswdConf.Value = ""
    Resume userCadastrar_Click_Exit
End Sub
