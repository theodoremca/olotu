Sub Creeate_Sheet()

    Dim sheetNames33 As Variant
    sheetNames33 = Array("FUTURE45", "FUTURE85", "HISTORY_DATA", "CRUD")
    Dim sheetName33 As Variant
    For Each sheetName33 In sheetNames33
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName33
    Next sheetName33
    
    Dim sheetNames As Variant
    sheetNames = Array("CCCMA_45_2040_2069", "CCCMA_45_2070_2099", "CCCMA_85_2040_2069", "CCCMA_85_2070_2099", "CCCMA_HIS_1981_2010")
    Dim sheetName As Variant
    For Each sheetName In sheetNames
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName
    Next sheetName
    
     Dim sheetNames1 As Variant
    sheetNames1 = Array("MIROC_45_2040_2069", "MIROC_45_2070_2099", "MIROC_85_2040_2069", "MIROC_85_2070_2099", "MIROC_HIS_1981_2010")
    Dim sheetName1 As Variant
    For Each sheetName1 In sheetNames1
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName1
    Next sheetName1
    
    
     Dim sheetNames2 As Variant
    sheetNames2 = Array("MOHC_45_2040_2069", "MOHC_45_2070_2099", "MOHC_85_2040_2069", "MOHC_85_2070_2099", "MOHC_HIS_1981_2010")
    Dim sheetName2 As Variant
    For Each sheetName2 In sheetNames2
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName2
    Next sheetName2
    
         Dim sheetNames3 As Variant
    sheetNames3 = Array("MPI_45_2040_2069", "MPI_45_2070_2099", "MPI_85_2040_2069", "MPI_85_2070_2099", "MPI_HIS_1981_2010", "CRUD_1981_2010")
    Dim sheetName3 As Variant
    For Each sheetName3 In sheetNames3
            Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName3
    Next sheetName3
End Sub


Sub Delete_Future()
   DeleteRows 2040, 2099, 1
End Sub


Function DeleteRows(ByVal firstYear As Integer, ByVal lastYear As Integer, ByVal firstRow As Long)
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim yearNumber As Integer
    Dim cellString As String
    Dim value As Variant
    Dim condition1 As Boolean
    Dim condition2 As Boolean
    
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    
    For i = lastRow To firstRow Step -1
        cellString = "A" & i
        value = ActiveSheet.range(cellString).value
        If IsDate(value) Then
        yearNumber = CInt(Format(value, "YYYY"))

         condition1 = yearNumber < firstYear
        condition2 = yearNumber > lastYear
              
        If condition1 Or condition2 Then
            Rows(i).Delete
        End If
        End If
    Next i
End Function




Function GetLastRow() As Long
    Dim lastRow As Long
    lastRow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
    GetLastRow = lastRow
End Function


Function CopyRangeBetweenWorkbooks(ByVal source As String, ByVal sourceSheet As String, ByVal destination As String, ByVal destinationSheet As String, ByVal range As String)
    Dim sourceWorkbook As Workbook
    Dim destinationWorkbook As Workbook
    Dim sourceWorksheet As Worksheet
    Dim destinationWorksheet As Worksheet
    Dim sourceRange As range
    Dim destinationRange As range
    
    ' Set the source workbook, sheet, and range
    Set sourceWorkbook = Workbooks(source) ' Update with the source workbook name
    Set sourceWorksheet = sourceWorkbook.Sheets(sourceSheet) ' Update with the source sheet name
    Set sourceRange = sourceWorksheet.range(range) ' Update with the source range
    
    ' Open the destination workbook
    Set destinationWorkbook = Workbooks(destination) ' Update with the destination workbook path
    
    ' Set the destination sheet and range
    Set destinationWorksheet = destinationWorkbook.Sheets(destinationSheet) ' Update with the destination sheet name
    Set destinationRange = destinationWorksheet.range(range)
    
    ' Copy the range from source to destination
    sourceRange.Copy destinationRange
    
    ' Save and close the destination workbook
    destinationWorkbook.Save
    sourceWorkbook.Close
End Function


Sub Future()
    Dim source As String
    Dim sourceSheet As String
    Dim destination As String
    Dim destinationSheet As String
    Dim range As String
    
    destinationSheet = "CCCMA"
    source = destinationSheet & ".xlsx"
    sourceSheet = "Lagos"
    destination = sourceSheet & ".xlsx"
    
    range = "A1:H1147"
    
    CopyRangeBetweenWorkbooks source, sourceSheet, destination, destinationSheet, range
End Sub

Sub Create_Sheet_1()

    Dim sheetNames33 As Variant
    sheetNames33 = Array("CCCMA", "MIROC", "MPI", "MOHC")
    Dim sheetName33 As Variant
    For Each sheetName33 In sheetNames33
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName33
    Next sheetName33
    

End Sub

Sub Amodel()
    Dim source As String
    Dim sourceSheet As String
    Dim destination As String
    Dim destinationSheet As String
    Dim range As String
    
    Dim sheetNames33 As Variant
    sheetNames33 = Array("CCCMA", "MIROC", "MPI", "MOHC")
    Dim sheetName33 As Variant
    For Each sheetName33 In sheetNames33
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = sheetName33
    destinationSheet = sheetName33
    source = destinationSheet & ".xlsx"
    sourceSheet = "Lagos"
    destination = sourceSheet & ".xlsx"
    
    range = "A1:H1147"
    
    CopyRangeBetweenWorkbooks source, sourceSheet, destination, destinationSheet, range
    Next sheetName33
    

End Sub

