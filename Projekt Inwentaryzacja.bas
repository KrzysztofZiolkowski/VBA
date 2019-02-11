Attribute VB_Name = "Module1"
Sub Stwórz_interfejs_TWIST_inwntaryzacji()

'tworzy nowy arkusz i na nim pracuje, zmienia nazwe oryginalnego oraz nowego arkusza
Dim ws As Worksheet
    Set wh = Worksheets(ActiveSheet.Name)
        ActiveSheet.Name = "Interfejs TWIST"
        ActiveSheet.Copy before:=Worksheets(Sheets.Count)
        ActiveSheet.Name = "Orygina³ TWIST"
    wh.Activate


Dim cell, cell2 As Range

'Przygotuj odpowiedni format dla liczb, wprowadŸ format wartoœci ujemnych i dodatnych zgodny z SAP.
For Each cell In Range("X:X")
    If cell.Value = "-1" Then
        cell.Offset(0, 1).Value = 0 - cell.Offset(0, 1).Value
    End If
Next cell

'usuñ zbêdne kolumny
Range("A:A,D:H,J:K,P:T,W:X,Z:AC,AF:BG,BI:BX").Delete

'zmieñ nazwy pozosta³ym nag³ówkom
Range("A1").Value = "Rodzaj ruchu"
Range("B1").Value = "Nr ruchu"
Range("C1").Value = "os. ksiêguj¹ca"
Range("D1").Value = "Indeks TWIST"
Range("E1").Value = "Nr sk³adu"
Range("F1").Value = "Nazwa sk³adu"
Range("G1").Value = "Materia³"
Range("H1").Value = "Nr. zam"
Range("I1").Value = "Nr. listu przewozowego"
Range("J1").Value = "Iloœæ"
Range("K1").Value = "Data ksiêgowania"
Range("L1").Value = "Wagon"
Range("M1").Value = "Pole 'Komentarz' w Twist"

'wstaw pust¹ kolumnê- sluzy do zatrzymania wykraczania tekstu z kolumny G poza zakres ( estetyka )
Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("H:H").Value = " "
Range("O:O").Value = " "

'Wyœrodkuj tekst w kolumnach i zmieñ rozmiar tekstu
With Range("A:N").EntireColumn
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Size = 8
End With

'Reszta formatowania
Range("G:G").HorizontalAlignment = xlLeft
Range("N:N").HorizontalAlignment = xlLeft
Range("L:L").NumberFormat = "yyyy-mm-dd"

'Zawijaj tekst komórek nag³ówkowych
With Range("A1:N1")
    .WrapText = True
    .Font.Bold = True
End With

'Ustaw szerokoœæ kolumn
Range("A:A").ColumnWidth = 5.14
Range("B:B").ColumnWidth = 5.86
Range("C:C").ColumnWidth = 7.29
Range("D:D").ColumnWidth = 9.43
Range("E:E").ColumnWidth = 5.29
Range("F:F").ColumnWidth = 7
Range("G:G").ColumnWidth = 31.43
Range("H:H").ColumnWidth = 0.5
Range("I:I").ColumnWidth = 7
Range("J:J").ColumnWidth = 12.57
Range("K:K").ColumnWidth = 4.3
Range("L:L").ColumnWidth = 11.71
Range("M:M").ColumnWidth = 14.86
Range("N:N").ColumnWidth = 19.29

'usuñ wszystkie wiersze z typem ruchu: "RG"
For Each cell2 In Range("A2", Range("A2").End(xlDown))
    If cell2.Value = "Rg" Then
        cell2.Value = Null
    End If
        
Next cell2

 Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete


'wyœwietl komunikat o pomyœlnym ukoñczeniu pracy
Cells(1, 1).Select
MsgBox ("Interfejs utworzony pomyœlnie" & vbNewLine & "Oryginalny arkusz zosta³ równie¿ zachowany")

End Sub
