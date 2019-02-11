Attribute VB_Name = "Module1"
Option Explicit


Sub dodaj_Wpis()

Dim wiersz1, place, kolumna, wiersz, wiersz2 As Range

Worksheets("z³om-SM").Activate
Worksheets("z³om-SM").Range("C1").End(xlDown).Select

Set wiersz1 = Cells(ActiveCell.Row, ActiveCell.Column)    'zmienna do poruszania sie wdluz wiersza gdzie beda wpisywane dane
Set place = Cells(ActiveCell.Row, ActiveCell.Column + 2)


Worksheets("GRP SC & SAF").Activate
Range("C64").End(xlDown).Offset(1, 0).Select



Do
ActiveCell.Offset(-1, 0).Select
Loop Until ActiveCell = wiersz1 And Cells(ActiveCell.Row, ActiveCell.Column + 6) = "AKTUALNA" And Cells(ActiveCell.Row, ActiveCell.Column + 3) = place

Set wiersz2 = Cells(ActiveCell.Row, ActiveCell.Column)

Worksheets("z³om-SM").Activate
wiersz1.Offset(0, 5) = wiersz2.Offset(0, 4)                          'przypisuje odbiorce
wiersz1.Offset(0, 6) = wiersz2.Offset(0, 5)                          'przypisuje cenê za tonê
wiersz1.Offset(0, 7) = wiersz2.Offset(0, 5) * wiersz1.Offset(0, 4)   'oblicza ca³¹ kwotê jaka bêdzie na fakturze
wiersz1.Offset(0, 8) = wiersz2.Offset(0, -2)                         'przypisuje numer formularza SC
'wiersz1.Offset(0, 9) = wiersz2.Offset(0, 2)                          'przypisuje dostêpn¹ do zu¿ycia iloœæ formularza sc
wiersz1.Offset(0, 10) = wiersz2.Offset(0, -1)                        'przypisuje numer aukcji


End Sub



Sub otworz_zlomowe_zestawy()
Workbooks.Open ("F:\GRE PROJECTS\SPM\Scrapping process\Wagon scrapping\GRP\2016\Z£OMOWANIE CZÊŒCI\ZESTAWY" & "\WS scrapped_2016.xlsx")
End Sub

Sub otworz_zlomowe_ramy()
Workbooks.Open ("F:\GRE PROJECTS\SPM\Scrapping process\Wagon scrapping\GRP\2016\Z£OMOWANIE CZÊŒCI\INNY Z£OM\ramy zez³omowane_2016.xlsx")
End Sub


Sub otworz_SC()

   Dim objWord
   Dim objDoc
   Set objWord = CreateObject("Word.Application")
   Set objDoc = objWord.Documents.Open("K:\SPM\Scrapping process\Material Scrapping\Formularze SC\F 7-40-10.01 Material Scrapping Application szablon.docx")
   objWord.Visible = True

End Sub


Sub Mail_Selection_Range_Outlook_Body()


    Dim rng As Range
    Dim OutApp As Object
    Dim OutMail As Object
    Dim emailRng As Range, cl As Range
    Dim sTo As String

'tu wprowadz zakres komorek z adresami
Set emailRng = Worksheets("address list").Range("D3:D50")
For Each cl In emailRng
        sTo = sTo & ";" & cl.Value
    Next
sTo = Mid(sTo, 2)

    Set rng = Nothing
    On Error Resume Next
    'Only the visible cells in the selection
    Set rng = Selection.SpecialCells(xlCellTypeVisible)
    'You can also use a fixed range if you want
    'Set rng = Sheets("YourSheet").Range("D4:D12").SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If rng Is Nothing Then
        MsgBox "The selection is not a range or the sheet is protected" & _
               vbNewLine & "please correct and try again.", vbOKOnly
        Exit Sub
        
    
    
    
    End If

    With Application
        .EnableEvents = False
        .ScreenUpdating = False
    End With

    Set OutApp = CreateObject("Outlook.Application")
    Set OutMail = OutApp.CreateItem(0)

    On Error Resume Next
    With OutMail
        .To = "Krzysztof.Ziolkowski@gatx.eu"
        .CC = "Bartosz.Porzucek@gatx.eu"
        .BCC = sTo 'sTo lista komorek w zakresie wprowadzona wyzej
        .Subject = "GATX RAIL GERMANY/ WSO - konkurs na sprzedaz zlomu: " & ActiveSheet.Name 'edytuj TEMAT WIADOMOSCI
        .HTMLBody = RangetoHTML(rng)
        .Display   'or use .Send
    End With
    On Error GoTo 0

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
    End With

    Set OutMail = Nothing
    Set OutApp = Nothing
End Sub


Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2016
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         Filename:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function




