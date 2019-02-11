Attribute VB_Name = "DOWNLOAD_SPOF"
Sub Downloadd()

''' Krzysztof Zi�kowski
''' Makro kopiuje tylko PIERWSZY za��cznik z maila do folderu SPOF_PDF K:\SPM\Key Materials Stock Management\Projekt SPOF

Dim myInspector As Outlook.Inspector
Dim myItem As Outlook.MailItem
Dim myAttachments As Outlook.Attachments

Set myInspector = Application.ActiveInspector

If Not TypeName(myInspector) = "Nothing" Then

         If TypeName(myInspector.CurrentItem) = "MailItem" Then
         Set myItem = myInspector.CurrentItem
         Set myAttachments = myItem.Attachments
         
    On Error GoTo ErrorHandle
        myAttachments.Item(1).SaveAsFile "K:\SPM\Key Materials Stock Management\Projekt SPOF\SPOF_PDF\" & myAttachments.Item(1).DisplayName
        MsgBox ("Za��cznik o nazwie " & myAttachments.Item(1).DisplayName & " zosta� pomy�lnie zapisany" & vbNewLine & vbNewLine & "Nast�pi zamkni�cie okna")

        myItem.Close olSave
        
ErrorHandle:
     Exit Sub
     
End If


Else
        MsgBox "Nie masz �adnej otwartego okna wiadomo�ci"
  
 End If

End Sub




