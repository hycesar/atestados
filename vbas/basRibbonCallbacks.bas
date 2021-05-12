Attribute VB_Name = "basRibbonCallbacks"
Option Compare Database

'################################################################
'#                                                              #
'#      Created with / Erstellt mit:                            #
'#      IDBE RibbonCreator 2019                                 #
'#      Version: 1.1003                                         #
'#                                                              #
'#      (c) 2007-2021 IDBE Avenius                              #
'#                                                              #
'#      http://www.ribboncreator2019.com                        #
'#      http://www.ribboncreator2016.com                        #
'#      http://www.ribboncreator2010.com                        #
'#      http://www.ribboncreator.com                            #
'#      http://www.accessribon.com                              #
'#      http://www.avenius.com                                  #
'#                                                              #
'#      You may send change requests or report errors to:       #
'#      Aenderungswuensche oder Fehler bitte an:                #
'#                                                              #
'#      mailto://info@ribboncreator2019.de                      #
'#                                                              #
'################################################################


' Globals:


Global Const strAppPicturePath As String = "Pics"             ' The pictures/icons are available in the following directory
                                                              ' below the database directory
                                                              ' %Databasepath%\Pics
                                                              ' Die Bilder / Icons liegen in folgendem Verzeichnis
                                                              ' unterhalb des Datenbankverzeichnises
                                                              ' %Datenbankpfad%\Pics
                                                              
Global Const bolUseDynamicPicturePath As Boolean = False      ' The images should be loaded from this directory
                                                              ' and not from the directory in the Ribbon XML.
                                                              ' Die Bilder sollen aus diesem Verzeichnis geladen werden
                                                              ' und nicht ueber die Verzeichnisangabe im Ribbon XML.

                                                              ' These values are used in the function "GetImages"
                                                              ' Diese Werte werden in der Funktion "GetImages" verwendet.

                                                              
Public gobjRibbon As IRibbonUI

Public bolEnabled As Boolean    ' Used in Callback "getEnabled"
                                ' Further informations in Callback "getEnabled"
                                ' Fuer Callback "getEnabled"
                                ' Genauere Informationen in Callback "getEnabled".
                               
Public bolVisible As Boolean    ' Used in Callback "getVisible"
                                ' More information in Callback "getVisible
                                ' Fuer Callback "getVisible"
                                ' Further informations in Callback "getVisible

' For Sample Callback "GetContent"
' Fuer Beispiel Callback "GetContent"
Public Type ItemsVal
    id As String
    label As String
    imageMso As String
End Type


' Callbacks:

Sub OnRibbonLoad(ribbon As IRibbonUI)
    ' Callbackname in XML File "onLoad"

    Set gobjRibbon = ribbon
End Sub

Sub LoadImages(control, ByRef image)
    ' Callbackname in XML File "loadImage"

    ' Loads an image with transparency to the ribbon
    ' Modul basGDIPlus is required
    ' Laed ein Bild mit Transparenz in das Ribbon
    ' Modul basGDIPlus wird dafuer benoetigt
    
    Dim strImage        As String
    Dim strPicture      As String
    
    strImage = CStr(control)
    strPicture = getPic(strImage)
    
    If strImage <> "" Then
        If bolUsePicturesFromTable = True Then
            If strPicture <> "" Then
                Set image = getIconFromTable(strPicture)
            Else
                Set image = Nothing
            End If
        Else
            Set image = LoadPictureGDIP(strImage)
        End If
    Else
        Set image = Nothing
    End If
    
End Sub

Sub GetImages(control As IRibbonControl, ByRef image)
    ' Callbackname in XML File "getImages"
    
    ' Loads an image with transparency to the ribbon
    ' Modul basGDIPlus is required
    ' Laed ein Bild mit Transparenz in das Ribbon
    ' Modul basGDIPlus wird dafuer benoetigt
    
    Dim strPicturePath  As String
    Dim strPicture      As String
    
    strPicture = getTheValue(control.Tag, "CustomPicture")
    
    If bolUsePicturesFromTable = True Then
        Set image = getIconFromTable(strPicture)
    Else
        If bolUseDynamicPicturePath = True Then
            strPicturePath = getAppPath & strAppPicturePath & "\"
        Else
            strPicturePath = getTheValue(control.Tag, "CustomPicturePath")
        End If
        Set image = LoadPictureGDIP(strPicturePath & strPicture)
    End If

End Sub

Sub GetEnabled(control As IRibbonControl, ByRef Enabled)
    ' Callbackname in XML File "getEnabled"
    
    ' To set the property "enabled" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
    ' Setzen der Enabled Eigenschaft eines Ribbon Controls
    ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12

    Select Case control.id
        
        Case Else
            Enabled = True

    End Select

End Sub

Sub GetVisible(control As IRibbonControl, ByRef visible)
    ' Callbackname in XML File "getVisible"
    
    ' To set the property "visible" to a Ribbon Control
    ' For further information see: http://www.accessribbon.de/en/index.php?Downloads:12
    ' Setzen der Visible Eigenschaft eines Ribbon Controls
    ' Weitere Informationen: http://www.accessribbon.de/index.php?Downloads:12

    Select Case control.id
        
        Case Else
            visible = True

    End Select

End Sub

Sub GetLabel(control As IRibbonControl, ByRef label)
    ' Callbackname in XML File "getLabel"
    ' To set the property "label" to a Ribbon Control

    Select Case control.id
        ''GetLabel''
        Case Else
            label = "*getLabel*"

    End Select

End Sub

Sub GetScreentip(control As IRibbonControl, ByRef screentip)
    ' Callbackname in XML File "getScreentip"
    ' To set the property "screentip" to a Ribbon Control

    Select Case control.id
        
        Case Else
            screentip = "*getScreentip*"

    End Select

End Sub

Sub GetSupertip(control As IRibbonControl, ByRef supertip)
    ' Callbackname in XML File "getSupertip"
    ' To set the property "supertip" to a Ribbon Control

    Select Case control.id
        
        Case Else
            supertip = "*getSupertip*"

    End Select

End Sub

Sub GetDescription(control As IRibbonControl, ByRef description)
    ' Callbackname in XML File "getDescription"
    ' To set the property "description" to a Ribbon Control

    Select Case control.id
        
        Case Else
            description = "*getDescription*"

    End Select

End Sub

Sub GetTitle(control As IRibbonControl, ByRef title)
    ' Callbackname in XML File "getTitle"
    ' To set the property "title" to a Ribbon MenuSeparator Control

    Select Case control.id
        
        Case Else
            title = "*getTitle*"

    End Select

End Sub

'Button

Sub OnActionButton(control As IRibbonControl)
    ' Callbackname in XML File "onAction"

    ' Callback for event button click
    ' Callback fuer Button Click
    
    Select Case control.id
              
        Case Else
            MsgBox "EU Button """ & control.id & """ clicked" & vbCrLf & _
                           "Es wurde auf Button """ & control.id & """ im Ribbon geklickt", _
                           vbInformation
    End Select

End Sub

'Command Button

Sub OnActionButtonHelp(control As IRibbonControl, ByRef CancelDefault)
    ' Callbackname in XML File Command "onAction"

    ' Callback for command event button click
    ' Callback fuer Command Button Click

    MsgBox "Button ""Help"" clicked" & vbCrLf & _
                           "Es wurde auf Button ""Hilfe"" geklickt", _
                           vbInformation
    CancelDefault = True

End Sub

'CheckBox

Sub OnActionCheckBox(control As IRibbonControl, _
                     pressed As Boolean)
    ' Callbackname in XML File "OnActionCheckBox"
    
    ' Callback for event checkbox click
    ' Callback fuer Checkbox Click

    Select Case control.id
                
        Case Else
            MsgBox "The Value of the Checkbox """ & control.id & """ is: " & pressed & vbCrLf & _
                   "Der Wert der Checkbox """ & control.id & """ ist: " & pressed, _
                   vbInformation

    End Select

End Sub

Sub GetPressedCheckBox(control As IRibbonControl, _
                       ByRef bolReturn)
    ' Callbackname in XML File "GetPressedCheckBox"
    
    ' Callback for checkbox
    ' indicates how the control is displayed
    ' Callback fuer Checkbox wie das Control
    ' angezeigt werden soll

    Select Case control.id
        
        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                bolReturn = True
            Else
                bolReturn = False
            End If

    End Select

End Sub

'ToggleButton

Sub OnActionTglButton(control As IRibbonControl, _
                      pressed As Boolean)
                              
    ' Callbackname in XML File "onAction"
    
    ' Callback fuer einen Toggle Button Klick
    ' Callback for a Toggle Buttons click event

    Select Case control.id
        
        Case Else
            MsgBox "The Value of the Toggle Button """ & control.id & """ is: " & pressed & vbCrLf & _
                   "Der Wert der Toggle Button """ & control.id & """ ist: " & pressed, _
                   vbInformation

    End Select

End Sub

Sub GetPressedTglButton(control As IRibbonControl, _
                        ByRef pressed)
    ' Callbackname in XML File "getPressed"

    ' Callback fuer ein Access ToogleButton Control wie dieser Angezeigt werden soll
    ' Callback for an Access ToogleButton Control. Indicates how the control is displayed

    Select Case control.id
        
        Case Else
            If getTheValue(control.Tag, "DefaultValue") = "1" Then
                pressed = True
            Else
                pressed = False
            End If

    End Select

End Sub

'EditBox

Sub GetTextEditBox(control As IRibbonControl, _
                   ByRef strText)
    ' Callbackname in XML File "GetTextEditBox"
    
    ' Callback fuer EditBox welcher Wert in der
    ' EditBox eingetragen werden soll.
    ' Callback for an EditBox Control
    ' Indicates which value is to set to the control

    Select Case control.id
        
        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    
    End Select
    
End Sub

Sub OnChangeEditBox(control As IRibbonControl, _
                    strText As String)
    ' Callbackname in XML File "OnChangeEditBox"
    
    ' Callback Editbox: Rueckgabewert der Editbox
    ' Callback Editbox: Return value of the Editbox

    Select Case control.id
        
        Case Else
            MsgBox "The Value of the EditBox """ & control.id & """ is: " & strText & vbCrLf & _
                   "Der Wert der EditBox """ & control.id & """ ist: " & strText, _
                   vbInformation

    End Select

End Sub

'DropDown

Sub OnActionDropDown(control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionDropDown"
    
    ' Callback onAction (DropDown)
    
    Select Case control.id
        
        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of DropDown-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                           "Die selektierte ItemID des DropDown-Control """ & control.id & """ ist : """ & selectedId & """", _
                           vbInformation
            End Select
    End Select

End Sub


Sub GetSelectedItemIndexDropDown(control As IRibbonControl, _
                                 ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexDropDown"
    
    ' Callback getSelectedItemIndex
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.id
            
            Case Else
                index = getTheValue(control.Tag, "DefaultValue")

        End Select

    End If

End Sub

'Gallery

Sub GetSelectedItemIndexGallery(control As IRibbonControl, _
                                   ByRef index)
    ' Callbackname in XML File "GetSelectedItemIndexGallery"
    
    ' Callback GetSelectedItemIndexGallery
    
    Dim varIndex As Variant
    varIndex = getTheValue(control.Tag, "DefaultValue")
    
    If IsNumeric(varIndex) Then
        Select Case control.id
            
            Case Else
                index = varIndex

        End Select

    End If

End Sub

Sub OnActionGallery(control As IRibbonControl, _
                     selectedId As String, _
                     selectedIndex As Integer)
    ' Callbackname in XML File "OnActionGallery"
    
    ' Callback onAction (Gallery)
    
    Select Case control.id
        
        Case Else
            Select Case selectedId
                Case Else
                    MsgBox "The selected ItemID of Gallery-Control """ & control.id & """ is : """ & selectedId & """" & vbCrLf & _
                           "Die selektierte ItemID des Gallery-Control """ & control.id & """ ist : """ & selectedId & """", _
                           vbInformation
            End Select
    End Select

End Sub

'Combobox

Sub GetTextComboBox(control As IRibbonControl, _
                      ByRef strText)

    ' Callbackname im XML File "GetTextComboBox"
    
    ' Callback getText (Combobox)
                           
    Select Case control.id
        
        Case Else
            strText = getTheValue(control.Tag, "DefaultValue")
    End Select

End Sub


Sub OnChangeComboBox(control As IRibbonControl, _
                               strText As String)
                           
    ' Callbackname im XML File "OnChangeCombobox"
    
    ' Callback onChange (Combobox)
   
    Select Case control.id
        
        Case Else
            MsgBox "The selected Item of Combobox-Control """ & control.id & """ is : """ & strText & """" & vbCrLf & _
                   "Das selektierte Item des Combobox-Control """ & control.id & """ ist : """ & strText & """", _
                   vbInformation
    End Select

End Sub

' DynamicMenu

Sub GetContent(control As IRibbonControl, _
               ByRef XMLString)

    ' Sample for a Ribbon XML "getContent" Callback
    ' See also http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '     and: http://www.accessribbon.de/en/index.php?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    ' Beispiel fuer einen Ribbon XML - "getContent" Callback
    ' Siehe auch: http://www.accessribbon.de/index.php?Access_-_Ribbons:Callbacks:dynamicMenu_-_getContent
    '       und : http://www.accessribbon.de/?Access_-_Ribbons:Ribbon_XML___Controls:Dynamic_Menu

    Select Case control.id
        
        Case Else
            XMLString = getXMLForDynamicMenu()
    End Select
 
End Sub


' Helper Function
' Hilfsfunktionen

Public Function getXMLForDynamicMenu() As String
    
    ' Creates a XML String for DynamicMenu CallBack - getContent
    
    ' Erstellt den Inhalt fuer das DynamicMenu im Callback getContent
    
    Dim lngDummy    As Long
    Dim strDummy    As String
    Dim strContent  As String
    
    Dim Items(4) As ItemsVal
    Items(0).id = "btnDy1"
    Items(0).label = "Item 1"
    Items(0).imageMso = "_1"
    Items(1).id = "btnDy2"
    Items(1).label = "Item 2"
    Items(1).imageMso = "_2"
    Items(2).id = "btnDy3"
    Items(2).label = "Item 3"
    Items(2).imageMso = "_3"
    Items(3).id = "btnDy4"
    Items(3).label = "Item 4"
    Items(3).imageMso = "_4"
    Items(4).id = "btnDy5"
    Items(4).label = "Item 5"
    Items(4).imageMso = "_5"
    
    strDummy = "<menu xmlns=""http://schemas.microsoft.com/office/2009/07/customui"">" & vbCrLf
    
        For lngDummy = LBound(Items) To UBound(Items)
            strContent = strContent & _
                "<button id=""" & Items(lngDummy).id & """" & _
                " label=""" & Items(lngDummy).label & """" & _
                " imageMso=""" & Items(lngDummy).imageMso & """" & _
                " onAction=""OnActionButton""/>" & vbCrLf
        Next
 

    strDummy = strDummy & strContent & "</menu>"
    getXMLForDynamicMenu = strDummy

End Function

Public Function getTheValue(strTag As String, strValue As String) As String
    ' *************************************************************
    ' Created from     : Avenius
    ' Parameter        : Input String, SuchValue String
    ' Date created     : 05.01.2008
    '
    ' Sample:
    ' getTheValue("DefaultValue:=Test;Enabled:=0;Visible:=1", "DefaultValue")
    ' Return           : "Test"
    ' *************************************************************
      
   On Error Resume Next
      
   Dim workTb()     As String
   Dim Ele()        As String
   Dim myVariabs()  As String
   Dim i            As Integer

      workTb = Split(strTag, ";")
      
      ReDim myVariabs(LBound(workTb) To UBound(workTb), 0 To 1)
      For i = LBound(workTb) To UBound(workTb)
         Ele = Split(workTb(i), ":=")
         myVariabs(i, 0) = Ele(0)
         If UBound(Ele) = 1 Then
            myVariabs(i, 1) = Ele(1)
         End If
      Next
      
      For i = LBound(myVariabs) To UBound(myVariabs)
         If strValue = myVariabs(i, 0) Then
            getTheValue = myVariabs(i, 1)
         End If
      Next
      
End Function

Public Function getAppPath() As String
    Dim strDummy As String
    strDummy = CurrentProject.Path
    If Right(strDummy, 1) <> "\" Then strDummy = strDummy & "\"
    getAppPath = strDummy
End Function

Public Function getIconFromTable(strFilename As String) As Picture
'*****************************************************************************
'Funktion 'getIconFromTable' holt ein Bild aus der Binaer-Tabelle und gibt
' ein Picture Objekt zurueck
' strFilename ist Bildobjekt welches in der Datei 'tblBinary" vorhanden ist.
'*****************************************************************************

Dim lSize As Long
Dim arrBin() As Byte
Dim Rs As DAO.Recordset
 
    On Error GoTo Errr
 
    Set Rs = DBEngine(0)(0).OpenRecordset("tblBinary", dbOpenDynaset)
    Rs.FindFirst "[FileName]='" & strFilename & "'"
    If Rs.NoMatch Then
        Set getIconFromTable = Nothing
    Else
        lSize = Rs.Fields("binary").FieldSize
        ReDim arrBin(lSize)
        arrBin = Rs.Fields("binary").GetChunk(0, lSize)
        Set getIconFromTable = ArrayToPicture(arrBin)
    End If
    Rs.Close
 
fExit:
    Reset
    Erase arrBin
    Set Rs = Nothing
    Exit Function
Errr:
    Resume fExit
End Function

Public Function getPic(strFullPath As String) As String
    Dim strResult As String
    
    If InStrRev(strFullPath, "\") > 0 Then
        strResult = Mid(strFullPath, InStrRev(strFullPath, "\") + 1)
    Else
        strResult = ""
    End If
   
    getPic = strResult
End Function

Private Sub gerarMensagemAleatoriaEspera()
    Dim num As Integer
    num = Int(((10 - 1 + 1) * Rnd) + 1)
    If (num = 1) Then
        MsgBox "Essa funcionalidade está sendo implementada", vbInformation, GENERALTTL & INFOTTL
    ElseIf (num = 2) Then
        MsgBox "Aguarde novidades...", vbInformation, GENERALTTL & INFOTTL
    ElseIf (num = 3) Then
        MsgBox "Mal podemos esperar!", vbInformation, GENERALTTL & INFOTTL
    ElseIf (num = 4) Then
        MsgBox "Estamos trabalhando nisso!", vbInformation, GENERALTTL & INFOTTL
    ElseIf (num = 5) Then
        MsgBox "Logo, logo, aguarde!", vbInformation, GENERALTTL & INFOTTL
    ElseIf (num = 6) Then
        MsgBox "Tenha calma, já vai!", vbInformation, GENERALTTL & INFOTTL
    ElseIf (num = 7) Then
        MsgBox "Eita, tá chegando. Aguarde mais um pouco.", vbInformation, GENERALTTL & INFOTTL
    ElseIf (num = 8) Then
        MsgBox "Demora só um pouco, mas vamos entregar essa funcionalidade.", vbInformation, GENERALTTL & INFOTTL
    ElseIf (num = 9) Then
        MsgBox "Aguarde só mais um pouquinho.", vbInformation, GENERALTTL & INFOTTL
    Else
        MsgBox "Estamos trabalhando nessa funcionalidade, só mais um pouco.", vbInformation, GENERALTTL & INFOTTL
    End If
End Sub

Sub btn_sbt_documentos_OnActionButton(control As IRibbonControl)
    DoCmd.OpenForm ("frAtestado")
End Sub

Sub btn_docEditar_OnActionButton(control As IRibbonControl)
    DoCmd.OpenForm ("frAtestadoEditor")
End Sub

Sub btn_docConsultar_OnActionButton(control As IRibbonControl)
    DoCmd.Close acQuery, "csAtestado"
    DoCmd.OpenQuery ("csAtestado")
End Sub

Sub btn_sbt_colaborador_OnActionButton(control As IRibbonControl)
    gerarMensagemAleatoriaEspera
End Sub

Sub btn_adicionar_OnActionButton(control As IRibbonControl)
    gerarMensagemAleatoriaEspera
End Sub

Sub btn_col_editar_OnActionButton(control As IRibbonControl)
    gerarMensagemAleatoriaEspera
End Sub

Sub btn_divisao_OnActionButton(control As IRibbonControl)
    gerarMensagemAleatoriaEspera
End Sub

Sub btn_doencas_OnActionButton(control As IRibbonControl)
    gerarMensagemAleatoriaEspera
End Sub

Sub btn_medicos_OnActionButton(control As IRibbonControl)
    gerarMensagemAleatoriaEspera
End Sub

Sub btn_afastamentos_OnActionButton(control As IRibbonControl)
    MsgBox "Lista dos funcionários com mais de 15 dias de atestados nos últimos 2 meses", vbExclamation, GENERALTTL & WARNTTL
    DoCmd.Close acQuery, "csEmpregadoINSS"
    DoCmd.OpenQuery "csEmpregadoINSS"
End Sub

Sub btn_declara_OnActionButton(control As IRibbonControl)
    MsgBox "Lista dos funcionários para abonar horas", vbExclamation, GENERALTTL & WARNTTL
    DoCmd.Close acQuery, "csEmpregadoAbono"
    DoCmd.OpenQuery "csEmpregadoAbono"
End Sub

Sub btn_sbt_usuario_OnActionButton(control As IRibbonControl)
    'gerarMensagemAleatoriaEspera
End Sub

Sub btn_cadastra_OnActionButton(control As IRibbonControl)
    DoCmd.OpenForm "frUserSignUp"
End Sub

Sub btn_user_editar_OnActionButton(control As IRibbonControl)
    DoCmd.OpenForm "frUserPswdChng"
End Sub

Sub btn_login_OnActionButton(control As IRibbonControl)
    'Abrir tela de login
    DoCmd.OpenForm "frUserSignIn"
End Sub

Sub btn_repair_OnActionButton(control As IRibbonControl)
    mdSecurity.relink
End Sub

Sub btn_informacao_OnActionButton(control As IRibbonControl)
    MsgBox "Em progressivo desenvolvimento por: Hywre ;)" & vbNewLine & " Dúvidas? Ramal: 831213", vbInformation, GENERALTTL & INFOTTL
End Sub
