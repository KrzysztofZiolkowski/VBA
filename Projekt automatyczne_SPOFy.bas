Attribute VB_Name = "Nowy_projekt"
Option Explicit
Sub Kopiuj_SPOF()
''' Krzysztof Zió³kowski
''' Program zajmuje siê konwertowaniem pojedynczego wiersza przeniesionego z TWIST do excela, i przeniesieniem istotnych dla SPM wartosci w okreslonym porz¹dku do zak³adki SPOF.
''' Program rozpoczyna siê od zaznaczenia wiersza który chcemy przekonwertowaæ - w zak³adce Twist Convert a nastêpnie wciœniêciu odpwiedniego przycisku
 
' DEKLARACJA ZMIENNYCH
Dim nr_zam, nr_wiersza As Integer
Dim cell, rng, zaznaczenie_komorki_spof As Range



nr_zam = Cells(Selection.Row, 2) 'PRZYPISZ DO ZMIENNEJ KOMÓRKÊ Z ADRESEM ZAMÓWIENIA
nr_wiersza = Selection.Row 'PRZYPISZ DO ZMIENNEJ NUMER ZAZNACZONEGO WIERSZA

Set zaznaczenie_komorki_spof = Worksheets("TWIST-SPOF").Cells(Selection.Row, 2) 'PRZYPISZ DO ZMIENNEJ NUMER SPOFA KTÓRY BÊDZIESZ WYSZUKIWAÆ W PRZEDZIALE F-F

'Worksheets("SPOF").Activate
Set rng = Worksheets("OrCe").Range("F:F")



'DLA KA¯DEJ ODWIEDZONEJ KOMÓRKI W ZAKRESIE F:F ( Or-Ce)WYKONAJ:
For Each cell In rng
    'cell.Select
    
'JEŒLI SZUKANY NR SPOFA ZNAJDUJE SIÊ JU¯ W ZAKRESIE F-F, AKTUALIZUJ REKORD WG PONI¯SZEGO:
    If cell = nr_zam Then
        
        cell.Offset(0, 1).Value = zaznaczenie_komorki_spof.Offset(0, 8).Value  'Rodzaj SPOF , offset 8
        cell.Offset(0, 2).Value = zaznaczenie_komorki_spof.Offset(0, 2).Value  'Status zamówienia, offset 2
        cell.Offset(0, 3).Value = zaznaczenie_komorki_spof.Offset(0, 53).Value 'zamawiaj¹cy 53
        cell.Offset(0, 4).Value = zaznaczenie_komorki_spof.Offset(0, 66).Value 'SM owner 1 66
        cell.Offset(0, 5).Value = zaznaczenie_komorki_spof.Offset(0, 67).Value 'SM owner 2 67
        cell.Offset(0, 6).Value = zaznaczenie_komorki_spof.Offset(0, 19).Value 'Przekazano do 19
        cell.Offset(0, 7).Value = zaznaczenie_komorki_spof.Offset(0, 3).Value 'Przejêcie przez SM 1 3
        cell.Offset(0, 8).Value = zaznaczenie_komorki_spof.Offset(0, 58).Value 'Przejêcie przez SM 2 58
        cell.Offset(0, 9).Value = zaznaczenie_komorki_spof.Offset(0, 18).Value 'Przekazano do 18
        
    
        
            If Worksheets("Twist convert").Range("CE3") = Empty Then
                cell.Offset(0, 15).Value = zaznaczenie_komorki_spof.Offset(0, 22).Value ' Miejscowoœæ 22
        Else:
                cell.Offset(0, 15).Value = zaznaczenie_komorki_spof.Offset(0, 82).Value + " " _
                + zaznaczenie_komorki_spof.Offset(0, 83).Value + " " + zaznaczenie_komorki_spof.Offset(0, 84).Value + " " + zaznaczenie_komorki_spof.Offset(0, 85).Value + " " + _
                zaznaczenie_komorki_spof.Offset(0, 86).Value
            
        End If
        
        cell.Offset(0, 16).Value = zaznaczenie_komorki_spof.Offset(0, 24).Value 'Osoba kontaktowa 24
        cell.Offset(0, 17).Value = zaznaczenie_komorki_spof.Offset(0, 14).Value 'nr zlecenia 14
        cell.Offset(0, 18).Value = zaznaczenie_komorki_spof.Offset(0, 69).Value 'Wagon 69
        cell.Offset(0, 19).Value = zaznaczenie_komorki_spof.Offset(0, 74).Value 'Indeks twist 74
        cell.Offset(0, 20).Value = zaznaczenie_komorki_spof.Offset(0, 16).Value 'Nazwa materia³u twist 16
        cell.Offset(0, 23).Value = zaznaczenie_komorki_spof.Offset(0, 33).Value 'Komentarz 33
        cell.Offset(0, 27).Value = zaznaczenie_komorki_spof.Offset(0, 17).Value 'Zamówiona iloœæ 17
        cell.Offset(0, 28).Value = zaznaczenie_komorki_spof.Offset(0, 11).Value '"Na Koszt 11
        cell.Offset(0, 36).Value = zaznaczenie_komorki_spof.Offset(0, 9).Value ' nr zapotrzebowania
        zaznaczenie_komorki_spof.Offset(0, -1).Value = "ZAKTUALIZOWANY"
        
        Exit For
    
    
    'JE¯ELI NIE ZNAJDZIESZ W ZAKRESIE TAKIEGO NR SPOF, STWÓRZ NOWY REKORD WG PONI¯SZEGO
    ElseIf cell = Empty Then
        
        cell.Offset(0, 0).Value = zaznaczenie_komorki_spof.Offset(0, 0).Value
        cell.Offset(0, 1).Value = zaznaczenie_komorki_spof.Offset(0, 8).Value  'Rodzaj SPOF , offset 8
        cell.Offset(0, 2).Value = zaznaczenie_komorki_spof.Offset(0, 2).Value  'Status zamówienia, offset 2
        cell.Offset(0, 3).Value = zaznaczenie_komorki_spof.Offset(0, 53).Value 'zamawiaj¹cy 53
        cell.Offset(0, 4).Value = zaznaczenie_komorki_spof.Offset(0, 66).Value 'SM owner 1 66
        cell.Offset(0, 5).Value = zaznaczenie_komorki_spof.Offset(0, 67).Value 'SM owner 2 67
        cell.Offset(0, 6).Value = zaznaczenie_komorki_spof.Offset(0, 19).Value 'Przekazano do DATA 18
        cell.Offset(0, 7).Value = zaznaczenie_komorki_spof.Offset(0, 3).Value 'Przejêcie przez SM 1 3
        cell.Offset(0, 8).Value = zaznaczenie_komorki_spof.Offset(0, 58).Value 'Przejêcie przez SM 2 58
        cell.Offset(0, 9).Value = zaznaczenie_komorki_spof.Offset(0, 18).Value 'Przekazano do 19

            If Worksheets("Twist convert").Range("CE3") = Empty Then
                cell.Offset(0, 15).Value = zaznaczenie_komorki_spof.Offset(0, 22).Value ' Miejscowoœæ 22
        Else:
                cell.Offset(0, 15).Value = zaznaczenie_komorki_spof.Offset(0, 82).Value + " " _
                + zaznaczenie_komorki_spof.Offset(0, 83).Value + " " + zaznaczenie_komorki_spof.Offset(0, 84).Value + " " + zaznaczenie_komorki_spof.Offset(0, 85).Value + " " + _
                zaznaczenie_komorki_spof.Offset(0, 86).Value
            
        End If
        
        cell.Offset(0, 16).Value = zaznaczenie_komorki_spof.Offset(0, 24).Value 'Osoba kontaktowa 24
        cell.Offset(0, 17).Value = zaznaczenie_komorki_spof.Offset(0, 14).Value 'nr zlecenia 14
        cell.Offset(0, 18).Value = zaznaczenie_komorki_spof.Offset(0, 69).Value 'Wagon 69
        cell.Offset(0, 19).Value = zaznaczenie_komorki_spof.Offset(0, 74).Value 'Indeks twist 74
        cell.Offset(0, 20).Value = zaznaczenie_komorki_spof.Offset(0, 16).Value 'Nazwa materia³u twist 16
        cell.Offset(0, 23).Value = zaznaczenie_komorki_spof.Offset(0, 33).Value 'Komentarz 33
        cell.Offset(0, 27).Value = zaznaczenie_komorki_spof.Offset(0, 17).Value 'Zamówiona iloœæ 17
        cell.Offset(0, 28).Value = zaznaczenie_komorki_spof.Offset(0, 11).Value '"Na Koszt 11
        cell.Offset(0, 36).Value = zaznaczenie_komorki_spof.Offset(0, 9).Value ' nr zapotrzebowania
        
        zaznaczenie_komorki_spof.Offset(0, -1).Value = "WPISANY"
        'zaznaczenie_komorki_spof.Offset(0, -1).Interior.ColorIndex = 37
        
        Exit For
    
    Else: End If
    
Next cell




End Sub



Sub Kopiuj_SPOFY()
''' Krzysztof Zió³kowski
''' Makro kopiuje brakuj¹ce zamówienia GATX do pliku OC

Dim zakres_spof As Range
Dim ostatni_wiersz_oc As Range
Dim cell As Range
Dim gatx_cell As Range

If Worksheets("Arkusz1").Range("A2") = "" Then
    MsgBox "Arkusz jest kompletny ! "
    Exit Sub
Else

Dim licznik As Integer
licznik = 1

Dim XL As Excel.Application
Dim WBK As Excel.Workbook

Set XL = CreateObject("Excel.application")
Set WBK = XL.Workbooks.Open("F:\GRE PROJECTS\SPM\Key Materials Stock Management\Projekt SPOF\FINAL PLIK.xlsm")
'Set WBK = XL.Workbooks.Open("file:///\\OSTFS01\VOL1\USER\wso-kziolkow\Desktop\FINAL%20PLIK.xlsm")
    
Set ostatni_wiersz_oc = Worksheets("OrCe").Range("A4").End(xlDown) 'wartoœæ ostatniej komórki w oc


' przeszukuj ka¿d¹ komórkê w kolumnie z numerami spofów niewpisanych
For Each cell In Worksheets("Arkusz1").Range("A2", Range("A2").End(xlDown))
            
            If cell = "" Then Exit For
            'przeszukuj ka¿d¹ komórkê w pliku GATX w celu znalezenia szukanego spofa i skopiowania danych z tego wiersza
            For Each gatx_cell In WBK.Worksheets("SPOF").Range("A4", "A1048576")
                If gatx_cell.Value = cell.Value Then
                    
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 2) = gatx_cell.Offset(0, 28).Value 'p³atnik
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 5) = gatx_cell.Value               'nr zamówienia SPOF
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 4) = gatx_cell.Offset(0, 9).Value   'data zamówienia
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 6) = gatx_cell.Offset(0, 20).Value  'nr i nazwa indeksu klienta
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 8) = "Dostêpne do reg:  " & gatx_cell.Offset(0, 21).Value & "      Dostêpne po reg/kwal/nowe:  " & gatx_cell.Offset(0, 22).Value    'nr indeksu WSO
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 9) = gatx_cell.Offset(0, 27).Value  'zamówiona iloœc
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 14) = gatx_cell.Offset(0, 15).Value  'Adres dostawy
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 15) = gatx_cell.Offset(0, 18).Value  'Uwagi do listu przewozowego '' wagon ?
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 16) = gatx_cell.Offset(0, 28).Value  'P³atnik frachtu
                        Worksheets("OrCe").Range("A4").End(xlDown).Offset(licznik, 17) = gatx_cell.Offset(0, 26).Value  'Uwagi
                    
                    
                    licznik = licznik + 1
                    Exit For
                    
                 ElseIf gatx_cell = "" Then Exit For
                 Else: End If
                 Next gatx_cell
        
        Next cell
        Opis_NOWY
        
    MsgBox "SPOFY skopiowane pomyœlnie !"
End If

End Sub




Sub Opis_NOWY()

Dim obecnie_przeszukiwana As Range
Set obecnie_przeszukiwana = Worksheets(1).Range("A3").End(xlDown).Offset(1, 0)

Do While obecnie_przeszukiwana.Offset(0, 2).Value <> ""
    
obecnie_przeszukiwana.Value = "NOWY"
Set obecnie_przeszukiwana = obecnie_przeszukiwana.Offset(1, 0)

Loop


End Sub





Sub Generuj_Listê_Brakuj¹cych_SPOF()
'''Krzysztof Zió³kowski
'''Makro sprawdza który wiersz ma opis "BRAK WPISU", i na tej podstawie pobiera wszystkie brakuj¹ce SPOFY, przekleja je do Kolumny A


Range("A2:A1048576").Clear
Dim Obecna_komorka_A As Range

Set Obecna_komorka_A = Range("A2")

For Each cell In Range("E2:E1048576")

    If cell.Value = "BRAK WPISU" Then
        Obecna_komorka_A.Value = cell.Offset(0, -3).Value
        Set Obecna_komorka_A = Obecna_komorka_A.Offset(1, 0)
    
    ElseIf cell.Value = "" Then Exit For
    ElseIf cell.Value = "WPISANY" Then End If


Next


MsgBox "Lista zaktualizowana pomyœlnie !"
End Sub























