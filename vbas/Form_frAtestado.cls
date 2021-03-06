VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frAtestado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private Sub dataHora_entrada_AfterUpdate()
    'Ocultar horas de abono
    Me.hora_abono.visible = False
    
    'Validations
    If DateDiff("h", Me.dataHora_entrada.Value, Me.dataHora_saida.Value) > 12 Then
        MsgBox "Uma declara��o n�o pode ter maior que 12 horas!", vbExclamation, GENERALTTL & WARNTTL
        Me.dataHora_entrada = Now()
    ElseIf DateDiff("n", Me.dataHora_entrada.Value, Me.dataHora_saida.Value) < 30 Then
        MsgBox "A diferen�a entre a hora de entrada e de sa�da deve ser maior que 30min!", vbExclamation, GENERALTTL & WARNTTL
        Me.dataHora_entrada = Now()
    Else
        Me.hora_abono.visible = True
        If Me.dias.Value > 0 Then
            MsgBox "Declara��es n�o podem ter dias de afastamento!", vbExclamation, GENERALTTL & WARNTTL
            Me.dias.Value = 0
        End If
    End If
End Sub

Private Sub dataHora_saida_AfterUpdate()
    'Ocultar horas de abono
    Me.hora_abono.visible = False
    
    'Validations
    If DateDiff("h", Me.dataHora_entrada.Value, Me.dataHora_saida.Value) > 12 Then
        MsgBox "Uma declara��o n�o pode ter maior que 12 horas!", vbExclamation, GENERALTTL & WARNTTL
        Me.dataHora_saida = Now()
    ElseIf DateDiff("n", Me.dataHora_entrada.Value, Me.dataHora_saida.Value) < 30 Then
        MsgBox "A diferen�a entre a hora de entrada e de sa�da deve ser maior que 30min!", vbExclamation, GENERALTTL & WARNTTL
        Me.dataHora_saida = Now()
    Else
        Me.hora_abono.visible = True
        If Me.dias.Value > 0 Then
            MsgBox "Declara��es n�o podem ter dias de afastamento!", vbExclamation, GENERALTTL & WARNTTL
            Me.dias.Value = 0
        End If
    End If
End Sub

Private Sub dias_AfterUpdate()
    If Me.dias.Value > 0 And DateDiff("h", Me.dataHora_entrada.Value, Me.dataHora_saida.Value) > 12 Then
        MsgBox "Atestado n�o pode ter hora de entrada e de sa�da!", vbExclamation, GENERALTTL & WARNTTL
        Me.dataHora_entrada = Now()
    End If
End Sub

Private Sub Form_AfterUpdate()
    If Me.dias.Value > 15 Then
        MsgBox "Funcion�rio com mais de 15 dias deve ser encaminhado ao INSS.", vbExclamation, GENERALTTL & WARNTTL
        DoCmd.OpenForm "frCaptcha", OpenArgs:="Funcion�rio deve ser encaminhado ao INSS."
    End If
End Sub
