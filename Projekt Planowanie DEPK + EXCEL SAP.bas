Attribute VB_Name = "Kopiuj_SPOF"
Option Explicit
Sub Kopiuj_SPOF()
''' Krzysztof ZiÛ≥kowski
''' Program zajmuje siÍ konwertowaniem pojedynczego wiersza przeniesionego z TWIST do excela, i przeniesieniem istotnych dla SPM wartosci w okreslonym porzπdku do zak≥adki SPOF.
''' Program rozpoczyna siÍ od zaznaczenia wiersza ktÛry chcemy przekonwertowaÊ - w zak≥adce Twist Convert a nastÍpnie wciúniÍciu odpwiedniego przycisku
    
' DEKLARACJA ZMIENNYCH
Dim nr_zam, nr_wiersza As Integer
Dim cell, rng, zaznaczenie_komorki_spof As Range

'PRZYPISZ DO ZMIENNEJ KOM”RK  Z ADRESEM ZAM”WIENIA
'PRZYPISZ DO ZMIENNEJ NUMER ZAZNACZONEGO WIERSZA
nr_zam = Cells(Selection.Row, 2)
nr_wiersza = Selection.Row

'PRZYPISZ DO ZMIENNEJ NUMER SPOFA KT”RY B DZIESZ WYSZUKIWA∆ W PRZEDZIALE B:B
Set zaznaczenie_komorki_spof = Worksheets("Twist convert").Cells(Selection.Row, 2)
'Worksheets("SPOF").Activate
Set rng = Worksheets("SPOF").Range("A:A")




'DLA KAØDEJ ODWIEDZONEJ KOM”RKI W ZAKRESIE A:a ( SPOF-WYSY£KI )WYKONAJ:
For Each cell In rng
    'cell.Select
    
'JEåLI SZUKANY NR SPOFA ZNAJDUJE SI  JUØ W ZAKRESIE B:B, AKTUALIZUJ REKORD WG PONIØSZEGO:
    If cell = nr_zam Then
        
        cell.Offset(0, 1).Value = zaznaczenie_komorki_spof.Offset(0, 8).Value  'Rodzaj SPOF , offset 8
        cell.Offset(0, 2).Value = zaznaczenie_komorki_spof.Offset(0, 2).Value  'Status zamÛwienia, offset 2
        cell.Offset(0, 3).Value = zaznaczenie_komorki_spof.Offset(0, 53).Value 'zamawiajπcy 53
        cell.Offset(0, 4).Value = zaznaczenie_komorki_spof.Offset(0, 66).Value 'SM owner 1 66
        cell.Offset(0, 5).Value = zaznaczenie_komorki_spof.Offset(0, 67).Value 'SM owner 2 67
        cell.Offset(0, 6).Value = zaznaczenie_komorki_spof.Offset(0, 19).Value 'Przekazano do 19
        cell.Offset(0, 7).Value = zaznaczenie_komorki_spof.Offset(0, 3).Value 'PrzejÍcie przez SM 1 3
        cell.Offset(0, 8).Value = zaznaczenie_komorki_spof.Offset(0, 58).Value 'PrzejÍcie przez SM 2 58
        cell.Offset(0, 9).Value = zaznaczenie_komorki_spof.Offset(0, 18).Value 'Przekazano do 18
        
    
        
            If Worksheets("Twist convert").Range("CE3") = Empty Then
                cell.Offset(0, 15).Value = zaznaczenie_komorki_spof.Offset(0, 22).Value ' MiejscowoúÊ 22
        Else:
                cell.Offset(0, 15).Value = zaznaczenie_komorki_spof.Offset(0, 82).Value + " " _
                + zaznaczenie_komorki_spof.Offset(0, 83).Value + " " + zaznaczenie_komorki_spof.Offset(0, 84).Value + " " + zaznaczenie_komorki_spof.Offset(0, 85).Value + " " + _
                zaznaczenie_komorki_spof.Offset(0, 86).Value
            
        End If
        
        cell.Offset(0, 16).Value = zaznaczenie_komorki_spof.Offset(0, 24).Value 'Osoba kontaktowa 24
        cell.Offset(0, 17).Value = zaznaczenie_komorki_spof.Offset(0, 14).Value 'nr zlecenia 14
        cell.Offset(0, 18).Value = zaznaczenie_komorki_spof.Offset(0, 69).Value 'Wagon 69
        cell.Offset(0, 19).Value = zaznaczenie_komorki_spof.Offset(0, 74).Value 'Indeks twist 74
        cell.Offset(0, 20).Value = zaznaczenie_komorki_spof.Offset(0, 16).Value 'Nazwa materia≥u twist 16
        cell.Offset(0, 23).Value = zaznaczenie_komorki_spof.Offset(0, 33).Value 'Komentarz 33
        cell.Offset(0, 27).Value = zaznaczenie_komorki_spof.Offset(0, 17).Value 'ZamÛwiona iloúÊ 17
        cell.Offset(0, 28).Value = zaznaczenie_komorki_spof.Offset(0, 11).Value '"Na Koszt 11
        cell.Offset(0, 36).Value = zaznaczenie_komorki_spof.Offset(0, 9).Value ' nr zapotrzebowania
        zaznaczenie_komorki_spof.Offset(0, -1).Value = "ZAKTUALIZOWANY"
        
        Exit For
    
    
    'JEØELI NIE ZNAJDZIESZ W ZAKRESIE TAKIEGO NR SPOF, STW”RZ NOWY REKORD WG PONIØSZEGO
    ElseIf cell = Empty Then
        
        cell.Offset(0, 0).Value = zaznaczenie_komorki_spof.Offset(0, 0).Value
        cell.Offset(0, 1).Value = zaznaczenie_komorki_spof.Offset(0, 8).Value  'Rodzaj SPOF , offset 8
        cell.Offset(0, 2).Value = zaznaczenie_komorki_spof.Offset(0, 2).Value  'Status zamÛwienia, offset 2
        cell.Offset(0, 3).Value = zaznaczenie_komorki_spof.Offset(0, 53).Value 'zamawiajπcy 53
        cell.Offset(0, 4).Value = zaznaczenie_komorki_spof.Offset(0, 66).Value 'SM owner 1 66
        cell.Offset(0, 5).Value = zaznaczenie_komorki_spof.Offset(0, 67).Value 'SM owner 2 67
        cell.Offset(0, 6).Value = zaznaczenie_komorki_spof.Offset(0, 19).Value 'Przekazano do DATA 18
        cell.Offset(0, 7).Value = zaznaczenie_komorki_spof.Offset(0, 3).Value 'PrzejÍcie przez SM 1 3
        cell.Offset(0, 8).Value = zaznaczenie_komorki_spof.Offset(0, 58).Value 'PrzejÍcie przez SM 2 58
        cell.Offset(0, 9).Value = zaznaczenie_komorki_spof.Offset(0, 18).Value 'Przekazano do 19

            If Worksheets("Twist convert").Range("CE3") = Empty Then
                cell.Offset(0, 15).Value = zaznaczenie_komorki_spof.Offset(0, 22).Value ' MiejscowoúÊ 22
        Else:
                cell.Offset(0, 15).Value = zaznaczenie_komorki_spof.Offset(0, 82).Value + " " _
                + zaznaczenie_komorki_spof.Offset(0, 83).Value + " " + zaznaczenie_komorki_spof.Offset(0, 84).Value + " " + zaznaczenie_komorki_spof.Offset(0, 85).Value + " " + _
                zaznaczenie_komorki_spof.Offset(0, 86).Value
            
        End If
        
        cell.Offset(0, 16).Value = zaznaczenie_komorki_spof.Offset(0, 24).Value 'Osoba kontaktowa 24
        cell.Offset(0, 17).Value = zaznaczenie_komorki_spof.Offset(0, 14).Value 'nr zlecenia 14
        cell.Offset(0, 18).Value = zaznaczenie_komorki_spof.Offset(0, 69).Value 'Wagon 69
        cell.Offset(0, 19).Value = zaznaczenie_komorki_spof.Offset(0, 74).Value 'Indeks twist 74
        cell.Offset(0, 20).Value = zaznaczenie_komorki_spof.Offset(0, 16).Value 'Nazwa materia≥u twist 16
        cell.Offset(0, 23).Value = zaznaczenie_komorki_spof.Offset(0, 33).Value 'Komentarz 33
        cell.Offset(0, 27).Value = zaznaczenie_komorki_spof.Offset(0, 17).Value 'ZamÛwiona iloúÊ 17
        cell.Offset(0, 28).Value = zaznaczenie_komorki_spof.Offset(0, 11).Value '"Na Koszt 11
        cell.Offset(0, 36).Value = zaznaczenie_komorki_spof.Offset(0, 9).Value ' nr zapotrzebowania
        
        zaznaczenie_komorki_spof.Offset(0, -1).Value = "WPISANY"
        'zaznaczenie_komorki_spof.Offset(0, -1).Interior.ColorIndex = 37
        
        Exit For
    
    Else: End If
    
Next cell




End Sub

Option Explicit

Sub KopiujWgZaznaczenia_stock_mgmt()
''' Krzysztof ZiÛ≥kowski
''' Program kopiuje zaznaczony zakres (indeksÛw ) do kolumny BA, pomija wartoúci puste ( oznaczone jako "-" ). Indeksy sπ wstawiane jeden pod drugim, a na koniec kopiowane.

'DEKLARACJA ZMIENNYCH
    Dim rng As Range
    Dim cell As Range
    Dim nr As Integer
    
'ZMIENNA OKREåLAJ•CA NR WIERSZA, OD KT”REGO NALEØY WKLEJA∆ INDEKSY DO KOLUMNY BA
    nr = 2
    
'DLA KAØDEJ KOM”RKI W ZAZNACZENIU
    Set rng = Selection
    For Each cell In rng
        
'JEØELI AKTUALNA KOM”RKA MA WARTOå∆ "-", PRZEJDè DO NAST PNEJ, W PRZCIWNYM WYPADKU SKOPIUJ WARTOå∆ DO KOLUMNY BA
        If cell.Value <> "-" Then
            cell.Copy Cells(nr, 39)
            nr = nr + 1
            
        Else: End If
        
    Next cell
    Cells(2, 39).Select
    
    
'ZAZNACZA KOM”RKI W KOLUMNIE BA DO SKOPIOWANIA
    Range("AM2", Range("AM2").End(xlDown)).Copy
    
    

End Sub



Option Explicit
Sub Matchuj_indeks_SAP()
'''Krzysztof ZiÛ≥kowski
''' Program PrÛbuje zidentyfikowaÊ indeks SAP na podstawie indeksu TWIST, oraz wypisuje Indeks, nazwÍ i iloúÊ SAP w komÛrcÍ pt KOMENTARZ

' indeks_twist          - zawiera referencjÍ do komorki z indeksem twist w arkuszu TWIST convert
' wartosc_wyszukiwania  - s≥uøy tylko do moøliwoúci sprawdzenia, czy wyszukiwanie docelowego indeksu w Macierz iloúci da≥o rezultat pozytywny , czy negatywny
' nazwa_SAP             - przechowuje nazwÍ indeksu TWIST z arkusza Macierz iloúci i dodaje go do pola KOMENTARZ na poczπtku komÛrki
' finalne_pole_reg      - przechowuje referencjÍ do komÛrki w ktÛrej bÍdπ dostÍpne indeksy i iloúci czÍúci po reg
' finalne_pole_do_reg   - przechowuje referencjÍ do komÛrki w ktÛrej bÍdπ dostÍpne indeksy i iloúci czÍúci DO REG
' znaleziona            - s≥uøy do zaznaczenia znalezionego indeksu TWIST w Macierz iloúci

Dim indeks_twist, wartosc_wyszukiwania, znaleziona, finalne_pole_reg, finalne_pole_doreg, finalne_pole_nazwa_SAP As Range

Worksheets("SPOF").Activate
Set indeks_twist = Selection
Set finalne_pole_doreg = Selection.Offset(0, 2)
Set finalne_pole_reg = Selection.Offset(0, 3)
Set finalne_pole_nazwa_SAP = Selection.Offset(0, 4)

Dim indeks_odniesienie As Range
Set indeks_odniesienie = Selection

If Selection.Column <> 20 Then
    MsgBox "Zaznaczy≥eú komÛrkÍ w niew≥aúcej kolumnie, wybierz komÛrkÍ z indeksem TWIST"
    Exit Sub
ElseIf Selection = "" Then
    MsgBox "Wybrane pole jest puste, wybierz pole z indeksem TWIST lub wpisz indeks"
    Exit Sub
End If
    

Dim licznik_komunikatu As Range

        
Cells(indeks_twist.Row, 22).ClearContents
Cells(indeks_twist.Row, 23).ClearContents
Cells(indeks_twist.Row, 24).ClearContents

Worksheets("Main").Activate

'PRZYPISUJE ZMIENN• DO WYNIKU WYSZUKIWANIA NUMERU INDEKSU
Set wartosc_wyszukiwania = Cells.Find(What:=indeks_twist, After:=ActiveCell, LookIn:=xlFormulas, _
                           LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                           MatchCase:=False, SearchFormat:=False)


                      
                                    If wartosc_wyszukiwania Is Nothing Then  'JEåLI NIE ZNAJDZIESZ TAKIEGO INDEKSU TWIST TO ZR”B FILTR NA ARKUSZU SPOF-Wysy≥ki DLA TEGO INDEKSU
                                        
                                        Worksheets("SPOF").Select
                                        ActiveSheet.Range("$A$2:$X$252").AutoFilter Field:=20, Criteria1:=indeks_twist
                                    
                                    
                                
                                    
                                    Else                                    'JEåLI ZNAJDIESZ TAKI INDEKS W "MACIERZ ILOåCI: TO:
                                            Set znaleziona = Cells.Find(What:=indeks_twist, After:=ActiveCell, LookIn:=xlFormulas, _
                                                             LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                                             MatchCase:=False, SearchFormat:=False)
                                                             znaleziona.Select
                                                             
                                                             Dim nazwa_SAP As Range
                                                             Dim zakres_ilosci_reg As Range
                                                             Dim zakres_ilosci_doreg As Range
                                                             
                                                             Set nazwa_SAP = Cells(znaleziona.Row, 1)
                                                             finalne_pole_nazwa_SAP = nazwa_SAP
                                                             
                                                             Set zakres_ilosci_reg = Range((Cells(znaleziona.Row, 18).Address()), (Cells(znaleziona.Row, 27).Address()))
                                                             Set zakres_ilosci_doreg = Range((Cells(znaleziona.Row, 28).Address()), (Cells(znaleziona.Row, 29).Address()))
                                                            
                            
                                           
                                            Dim cell, cell2 As Range
                                            
                                            For Each cell In zakres_ilosci_reg       'PRZESZUKAJ PUL  INDEKS”W, I WYBIERZ DO SKOPIOWANIA TE, KT”RYCH ILOåCI S• NA STANIE
                                                    cell.Select
                                                    If cell.Value <> 0 Then
                                                    finalne_pole_reg.Value = finalne_pole_reg.Value & " " & cell.Offset(0, -13).Value & " - " & cell.Value & " szt"
                                                            
                                                    Else: End If
                                                            
                                            Next cell
                                        
                                        
                                            For Each cell2 In zakres_ilosci_doreg       'PRZESZUKAJ PUL  INDEKS”W, I WYBIERZ DO SKOPIOWANIA TE, KT”RYCH ILOåCI S• NA STANIE
                                                    cell2.Select
                                                    If cell2.Value <> 0 Then
                                                    finalne_pole_doreg.Value = finalne_pole_doreg.Value & " " & cell2.Offset(0, -13).Value & " - " & cell2.Value & " szt"
                                                            
                                                    Else: End If
                                                            
                                            Next cell2
                                        
                                            
                                End If
                                
                                If finalne_pole_reg = Empty Then
                                finalne_pole_reg.Value = "Brak indeksÛw lub materia≥u"
                                Else: End If
                                
                                If finalne_pole_doreg = Empty Then
                                finalne_pole_doreg.Value = "Brak indeksÛw lub materia≥u"
                                Else: End If
                                
                                Worksheets("SPOF").Activate
                                indeks_twist.Select
                            
             
Set licznik_komunikatu = indeks_twist.Offset(0, 5)

If licznik_komunikatu >= indeks_twist.Offset(0, 6) Then
    licznik_komunikatu = 1

Else
    licznik_komunikatu = licznik_komunikatu + 1
End If
                                                
End Sub


Sub Outlook_emails()

Dim initalizuj_outlook As Outlook.Application
Dim nowy_email As Outlook.MailItem
Dim olInsp As Outlook.Inspector
Dim wdDoc As Word.Document
Dim nr_Spof As String

If Selection = "" Then
GoTo komunikat
komunikat:
MsgBox "Wybierz niepustπ komÛrkÍ! "
Exit Sub

Else

nr_Spof = Cells(Selection.Row, 1).Value
 '
Dim strUszanowanko As String
Dim dupa As Range


strUszanowanko = "Witam, proszÍ o realizacjÍ SPOF."

Set initalizuj_outlook = New Outlook.Application
Set nowy_email = initalizuj_outlook.CreateItem(outlookObiekt)

With nowy_email
    
    .BodyFormat = olFormatHTML
    .Display
    .To = "material@wsostroda.eu"
    .CC = "GRPDLTOSparePartsMgmt@gatx.eu;mag@wsostroda.eu;Dariusz.Jelen@wsostroda.eu"
    .Subject = "SPOF-" & Cells(Selection.Row, 37) & " (" & nr_Spof & "/" & Cells(Selection.Row, 19) & ")"
    .Attachments.Add "F:\GRE PROJECTS\SPM\Key Materials Stock Management\Projekt SPOF\SPOF_PDF\" & "SPOF-Liste " & nr_Spof & ".pdf"
 
         
    Set olInsp = .GetInspector
    Set wdDoc = olInsp.WordEditor
    
    wdDoc.Range.InsertBefore strUszanowanko
    
    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=1
    Arkusz3.Range("A3", Cells(Selection.Row, 29)).Copy
    wdDoc.Range(Len(strUszanowanko), Len(strUszanowanko)).Paste

    ActiveSheet.Outline.ShowLevels RowLevels:=0, ColumnLevels:=2
    Application.CutCopyMode = False
    
  

End With

End If

End Sub


Public Sub Get_attachment_adress()

path_file As String
path_file = "K:\SPM\Key Materials Stock Management\Projekt SPOF\SPOF_PDF\" & "SPOF-Liste " & nr_Spof & ".pdf"


End Sub




Sub Rozmiar_arkuszÛw()
'Update 20140526
Dim xWs As Worksheet
Dim rng As Range
Dim xOutWs As Worksheet
Dim xOutFile As String
Dim xOutName As String
xOutName = "KutoolsforExcel"
xOutFile = ThisWorkbook.Path & "\TempWb.xls"
On Error Resume Next
Application.DisplayAlerts = False
Err = 0
Set xOutWs = Application.Worksheets(xOutName)
If Err = 0 Then
    xOutWs.Delete
    Err = 0
End If
With Application.ActiveWorkbook.Worksheets.Add(Before:=Application.Worksheets(1))
    .Name = xOutName
    .Range("A1").Resize(1, 2).Value = Array("Worksheet Name", "Size")
End With
Set xOutWs = Application.Worksheets(xOutName)
Application.ScreenUpdating = False
xIndex = 1
For Each xWs In Application.ActiveWorkbook.Worksheets
    If xWs.Name <> xOutName Then
        xWs.Copy
        Application.ActiveWorkbook.SaveAs xOutFile
        Application.ActiveWorkbook.Close savechanges:=False
        Set rng = xOutWs.Range("A1").Offset(xIndex, 0)
        rng.Resize(1, 2).Value = Array(xWs.Name, VBA.FileLen(xOutFile))
        Kill xOutFile
        xIndex = xIndex + 1
    End If
Next
Application.ScreenUpdating = True
Application.Application.DisplayAlerts = True
End Sub



Option Explicit

Sub Wyczysc_kolumne_stock_mgmt()
''' Krzysztof ZiÛ≥kowski
''' Program czyúci kolumnÍ BA ze skopiowanych tam indeksÛw, zaczynajπc od komÛrki BA3 i koÒczπc na ostatniej ,w ktÛrej zawarty jest jakiú ≥aÒcuch znakÛw.

    Range("AJ2", Range("AJ2").End(xlDown)).Select
    Selection.Clear
    Range("AJ2").Select
    
End Sub


Sub Usun_entery()


    Dim MyRange As Range
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
 
    For Each MyRange In Selection
        If 0 < InStr(MyRange, Chr(10)) Then
            MyRange = Replace(MyRange, Chr(10), "")
        End If
    Next
 
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


