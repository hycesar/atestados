VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frUserSignUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub Form_open(Cancel As Integer)
On Error GoTo Form_open_error:
    Dim actual_user As String
    actual_user = mdSecurity.ReadFileToText("C:\temp\log.txt")
    
    'Check if can be converted to a id
    If Not IsNumeric(actual_user) Then
        MsgBox "Não há usuário logado.", vbExclamation, GENERALTTL & ERRORTTL
        Cancel = True
    Else
        If DLookup("[tbGrupo_id]", "tbUsuario", "[matricula]= " & actual_user) <> 1 Then
            MsgBox "Somente administradores tem permissão de cadastrar usuários.", vbExclamation, GENERALTTL & WARNTTL
            Cancel = True
        End If
    End If

Form_open_exit:
    Exit Sub

Form_open_error:
    MsgBox Error$, vbCritical, GENERALTTL & FATALTTL
    Resume Form_open_exit
End Sub

Private Sub Form_Error(DataErr As Integer, Response As Integer)
    'Err: OpenForm was canceled = 2501
    Resume Next
End Sub

'------------------------------------------------------------
' Form_frUser: Cadastrar Usuário do Sistema
' Sumary: Valida campos, hash de senha e guardar em tabela
'------------------------------------------------------------
Private Sub userCadastrar_Click()
    On Error GoTo userCadastrar_Click_Err

    'Declarando um objeto da classe HashMD5
    Dim objMD5 As New clHashMD5
    
    'Inicializar gerador de números aleatórios
    Randomize
    
    'Preencher formulário com salt
    Me.salt.Value = CStr(Int(((9999 - 1000 + 1) * Rnd) + 1000))
    
    If Not IsNull(Me.matr.Value + Me.pswd.Value + Me.pswdConf.Value + Me.cbGrupo.Value) Then
        If Me.pswd.Value = Me.pswdConf.Value Then
            'Calcular hash
            Me.pswd.Value = objMD5.DigestStrToHexStr(Me.pswdConf.Value + Me.salt.Value)
            
            'Preencher formulário
            Me.pswdConf.Value = Me.pswd.Value
            
            'Procurar entrada duplicada
            If IsNull(DLookup("[matricula]", "tbUsuario", "[matricula] = " & Me.matr.Value)) Then
                
                'Iniciar transação
                'DBEngine.Workspaces(0).BeginTrans
            
                'SQL INSERT
                CurrentDb.Execute "INSERT INTO tbUsuario VALUES (" & Me.matr.Value & ", '" _
                                                                    & Me.pswd.Value & "', '" _
                                                                    & Me.salt.Value & "', " _
                                                                    & Me.cbGrupo.Value & ")", dbFailOnError
                'Concluir transação
                'DBEngine.Workspaces(0).CommitTrans

                'Limpar formulários
                Me.matr.Value = Null
                Me.pswd.Value = Null
                Me.pswdConf.Value = Null
                Me.cbGrupo.Value = Null
            
                'Mensagem de sucesso
                MsgBox "Usuário e senha cadastrados com sucesso.", vbInformation, GENERALTTL & INFOTTL
            Else
                If (vbYes = MsgBox("Usuário já estava cadastrado!" & vbNewLine & "Deseja é alterar sua senha?", vbQuestion + vbYesNo + vbDefaultButton2, GENERALTTL & QUERYTTL)) Then
                    'keep the value after close
                    Dim args As Variant
                    args = Me.matr.Value
                    
                    'Fechar atual
                    DoCmd.Close
                    
                    'Abrir formulario de alteração
                    DoCmd.OpenForm "frUserPswdChng", OpenArgs:=args
                Else
                    'Limpar formulários
                    Me.matr.Value = Null
                    Me.pswd.Value = Null
                    Me.pswdConf.Value = Null
                    Me.cbGrupo.Value = Null
                    
                    'Alertar usuário
                    MsgBox "Nenhuma operação foi executada. Tente novamente.", vbCritical, GENERALTTL & ERRORTTL
                End If
            End If
        Else
            'Mensagem de Erro dependente do usuário
            MsgBox "Favor, digitar senhas iguais.", vbExclamation, GENERALTTL & WARNTTL
        End If
    Else
        'Mensagem de Erro dependente do usuário
        MsgBox "Favor, preencher todos os campos.", vbExclamation, GENERALTTL & WARNTTL
    End If

userCadastrar_Click_Exit:
    Exit Sub

userCadastrar_Click_Err:
    If Err.Number = 3022 Then
        If (vbYes = MsgBox("Usuário já estava cadastrado!" & vbNewLine & "Deseja é alterar sua senha?", vbQuestion + vbYesNo + vbDefaultButton2, GENERALTTL & QUERYTTL)) Then
            'keep the value after close
            Dim arg As Variant
            arg = Me.matr.Value
            
            'Fechar atual
            DoCmd.Close
            
            'Abrir formulario de alteração
            DoCmd.OpenForm "frUserPswdChng", OpenArgs:=arg
            
            'Sair da rotina
            Resume userCadastrar_Click_Exit
        Else
            'Limpar formulários
            Me.matr.Value = Null
            Me.pswd.Value = Null
            Me.pswdConf.Value = Null
            Me.cbGrupo.Value = Null
            
            'Alertar usuário
            MsgBox "Nenhuma operação foi executada. Tente novamente.", vbCritical, GENERALTTL & ERRORTTL
        End If
    ElseIf Err.Number = 3003 Then
        'Limpar BD
        DBEngine.Workspaces(0).CommitTrans
        DBEngine.Workspaces(0).Close
        CurrentDb.Close
    Else
        'Alertar usuário
        MsgBox Error$, vbCritical, GENERALTTL & FATALTTL
    End If
    'Limpar formulários
    Me.matr.Value = Null
    Me.pswd.Value = Null
    Me.pswdConf.Value = Null
    Me.cbGrupo.Value = Null
    
    'Sair da rotina
    Resume userCadastrar_Click_Exit
End Sub



