Attribute VB_Name = "Generuj_Arkusze"
Option Explicit
Public Sub Kopiuj_arkusze()

'''Krzysztof Zió³kowski
'Metoda generuje arkusze bêd¹ce odpowiednikiem wagonów wykazanych w arkuszu "Fc zbiorówka"

    Dim cell, cell2, rng As Range
    Dim WS As Excel.Worksheet
    Dim WBK As Excel.Workbook
    
    
    Set WBK = ActiveWorkbook

    For Each cell In Selection

    cell.NumberFormat = "@"
    cell.Value = cell.FormulaR1C1

           With WBK
                Worksheets("baza formularz").Copy After:=Worksheets(Sheets.Count)
                On Error Resume Next
                ActiveSheet.Name = cell.Value
                Range("B2").Value = cell.Value 'wagon
                Range("F2").Value = cell.Offset(0, 3)  'typ
                Range("E2").Value = cell.Offset(0, 1) / 1000 'tara
                Range("I2").Value = cell.Offset(0, 2)  'wlasciciel
            End With
        
    With cell.Offset(0, 7)
        .Value = Date
        .NumberFormat = "dd/mm/yyyy"
    End With
 
        
    Next cell
    
    Worksheets("baza lista").Activate
    
    
    
    
    
End Sub



Sub Szukaj_Arkusza()

Dim xName As String
Dim xFound As Boolean

xName = InputBox("Wpisz nazwê arkusza któr¹ chcesz znaleŸæ:", "Szukanie arkusza")

If xName = "" Then Exit Sub
On Error Resume Next
ActiveWorkbook.Sheets(xName).Select
xFound = (Err = 0)

On Error GoTo 0


If xFound Then
    MsgBox "Arkusz '" & xName & "' Arkusz zosta³ odnaleziony i zaznaczony!"
Else
    MsgBox "Podany arkusz '" & xName & "' nie istnieje!"
    
End If


End Sub


Sub Powrot_do_bazy()
    Worksheets("Fc zbiorówka").Activate
End Sub



Sub Format_WagonNr()

    Dim cell As Range
    
    For Each cell In Selection
    
    
    cell.NumberFormat = "@"

    cell.Replace What:=" ", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
   
    cell.Replace What:="-", Replacement:="", LookAt:=xlPart, _
    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    ReplaceFormat:=False
    cell.FormulaR1C1 = cell.Value
    cell.Value = cell.FormulaR1C1
    
    cell.NumberFormat = "@"
     
     Next cell
   
End Sub



Sub Restore_WagonNr()

Dim cell As Range
Dim a, b, c, d, e As String


For Each cell In Selection

a = Mid(cell.Text, 1, 2)
b = Mid(cell.Text, 3, 2)
c = Mid(cell.Text, 5, 4)
d = Mid(cell.Text, 9, 3)
e = Mid(cell.Text, 12, 1)

cell.Value = a & " " & b & " " & c & " " & d & "-" & e

Cells(1, 1) = cell.Value

Next cell



End Sub



Sub GET_WAGA()

Dim cell, rngCAR As Range
Set rngCAR = Worksheets("baza lista").Range("A2", Range("A2").End(xlDown))

On Error GoTo errorhandle

For Each cell In rngCAR
    If cell.Value = "" Then Exit Sub
    cell.Offset(0, 5).Value = ActiveWorkbook.Sheets(cell.Text).Range("I100").End(xlUp).Value
    

errorhandle:
Resume Next

Next cell
End Sub




Sub CHECK_DATE()

Dim xName As String
Dim cell As Range



For Each cell In Range("h2", "h50")
    xName = cell.Offset(0, -7).Text
 
    On Error GoTo xd:
    If ActiveWorkbook.Sheets(xName).Range("B2").Text = xName Then
    
    End If
    
xd:
    cell.Select
    cell.Value = ""
    Next cell

    


End Sub













 
