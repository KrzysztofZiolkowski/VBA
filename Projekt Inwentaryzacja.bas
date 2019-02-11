Attribute VB_Name = "Module1"
Sub Stw�rz_interfejs_TWIST_inwntaryzacji()

'tworzy nowy arkusz i na nim pracuje, zmienia nazwe oryginalnego oraz nowego arkusza
Dim ws As Worksheet
    Set wh = Worksheets(ActiveSheet.Name)
        ActiveSheet.Name = "Interfejs TWIST"
        ActiveSheet.Copy before:=Worksheets(Sheets.Count)
        ActiveSheet.Name = "Orygina� TWIST"
    wh.Activate


Dim cell, cell2 As Range

'Przygotuj odpowiedni format dla liczb, wprowad� format warto�ci ujemnych i dodatnych zgodny z SAP.
For Each cell In Range("X:X")
    If cell.Value = "-1" Then
        cell.Offset(0, 1).Value = 0 - cell.Offset(0, 1).Value
    End If
Next cell

'usu� zb�dne kolumny
Range("A:A,D:H,J:K,P:T,W:X,Z:AC,AF:BG,BI:BX").Delete

'zmie� nazwy pozosta�ym nag��wkom
Range("A1").Value = "Rodzaj ruchu"
Range("B1").Value = "Nr ruchu"
Range("C1").Value = "os. ksi�guj�ca"
Range("D1").Value = "Indeks TWIST"
Range("E1").Value = "Nr sk�adu"
Range("F1").Value = "Nazwa sk�adu"
Range("G1").Value = "Materia�"
Range("H1").Value = "Nr. zam"
Range("I1").Value = "Nr. listu przewozowego"
Range("J1").Value = "Ilo��"
Range("K1").Value = "Data ksi�gowania"
Range("L1").Value = "Wagon"
Range("M1").Value = "Pole 'Komentarz' w Twist"

'wstaw pust� kolumn�- sluzy do zatrzymania wykraczania tekstu z kolumny G poza zakres ( estetyka )
Columns("H:H").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
Columns("H:H").Value = " "
Range("O:O").Value = " "

'Wy�rodkuj tekst w kolumnach i zmie� rozmiar tekstu
With Range("A:N").EntireColumn
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Font.Size = 8
End With

'Reszta formatowania
Range("G:G").HorizontalAlignment = xlLeft
Range("N:N").HorizontalAlignment = xlLeft
Range("L:L").NumberFormat = "yyyy-mm-dd"

'Zawijaj tekst kom�rek nag��wkowych
With Range("A1:N1")
    .WrapText = True
    .Font.Bold = True
End With

'Ustaw szeroko�� kolumn
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

'usu� wszystkie wiersze z typem ruchu: "RG"
For Each cell2 In Range("A2", Range("A2").End(xlDown))
    If cell2.Value = "Rg" Then
        cell2.Value = Null
    End If
        
Next cell2

 Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete


'wy�wietl komunikat o pomy�lnym uko�czeniu pracy
Cells(1, 1).Select
MsgBox ("Interfejs utworzony pomy�lnie" & vbNewLine & "Oryginalny arkusz zosta� r�wnie� zachowany")

End Sub
