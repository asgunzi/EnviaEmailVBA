Attribute VB_Name = "M�dulo1"


Sub enviaEmail()
Dim i As Long
Dim e_dados As Variant, nl As Long
Dim boolConfirmar As Boolean

Application.ScreenUpdating = False

nl = Application.WorksheetFunction.CountA(Range("b5:b5000"))
If nl = 0 Then
    MsgBox "N�o h� e-mail a enviar"
    Exit Sub
End If

boolConfirmar = Range("aa1")

e_dados = Range("b5:d" & 4 + nl)

For i = 1 To nl
    E_Mail e_dados(i, 3), e_dados(i, 2), e_dados(i, 1), False, boolConfirmar
    
Next i

End Sub


Sub E_Mail(sHTMLBody As Variant, sSubject As Variant, sTo As Variant, bHasAttachment As Boolean, boolconf As Boolean, _
                Optional sPathAttachment As String, Optional sCC As String, Optional sBC As String)
'#######################################################
'#                                                                                                          #
'# Monta um e-mail automaticamente de acordo com os par�metros de entrada.  #
'#                                                                                                          #
'#######################################################

'\\ Declara��o de vari�veis
Dim objOutlook As Object, objOutlookMail As Object

'\\ Desativa atualiza��o de tela e alertas do Excel
With Application
  .ScreenUpdating = False
  .DisplayAlerts = False
End With

'\\ Em caso de erro, solicita o envio do e-mail manualmente
On Error GoTo fim

'\\ Define os objetos do outlook: Aplica��o e novo e-mail
Set objOutlook = Interaction.CreateObject("Outlook.Application") 'Applica��o do Outlook
Set objOutlookMail = objOutlook.CreateItem(0) 'Novo E-mail

'\\ Insere as informa��es no e-mail para envio
With objOutlookMail
  .to = sTo
  .CC = sCC
  .BCC = sBC
  .Subject = sSubject
  .BodyFormat = 2 'olFormatHTML
  If bHasAttachment Then .Attachments.Add sPathAttachment 'Insere o anexo caso seja necess�rio
  .Display 'exibe o e-mail para copiar a assinatura
  sHTMLBody = sHTMLBody & "<br>" & .HTMLBody 'Copia a assinatura de e-mail padr�o
  .HTMLBody = sHTMLBody
  
  If Not boolconf Then
      .Send 'para enviar automaticamente
    End If
End With

fim:
If Err.Number > 0 Then
  Interaction.MsgBox "N�o foi poss�vel o envio autom�tico do e-mail, favor enviar o relat�rio manualmente.", vbCritical
End If
On Error GoTo 0

'\\ Esvazia os objetos
Set objOutlookMail = Nothing
Set objOutlook = Nothing

'\\ Ativa novamente a atualiza��o de tela e alertas do excel
With Application
  .DisplayAlerts = True
  .ScreenUpdating = True
End With

End Sub





