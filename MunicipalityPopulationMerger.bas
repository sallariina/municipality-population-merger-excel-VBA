Attribute VB_Name = "MunicipalityPopulationMerger"
Sub MainApplication()
    Dim ws As Worksheet
    Dim cityPopulationData As Collection
    Set cityPopulationData = New Collection

    ' Read data from CSVs
    OpenSheet ws, "2022"
    ProcessCSVData ws, cityPopulationData, "2022" ' Process data for 2022
    CloseSheet

    OpenSheet ws, "2023"
    ProcessCSVData ws, cityPopulationData, "2023" ' Process data for 2023
    CloseSheet
    
    ' Call WriteVäkiluvutData to create the sheet and write data
    WriteVäkiluvutData cityPopulationData
    
    ' Call SaveVäkiluvutSheet to save Väkiluvut sheet to the same folder as the workbook
    SaveVäkiluvutSheet
End Sub

Sub OpenSheet(ByRef ws As Worksheet, ByRef year As String)
    Dim folderpath As String, fileName As String
    Dim wb As Workbook
    
    folderpath = Application.ActiveWorkbook.Path
    If year = "2022" Then
        fileName = "2022.csv"
    ElseIf year = "2023" Then
        fileName = "2023.csv"
    End If
    
    ' Open the workbook and set the active sheet
    Set wb = Workbooks.Open(folderpath & "\" & fileName)
    Set ws = wb.Sheets(1)
End Sub

Sub CloseSheet()
    ActiveWorkbook.Close SaveChanges:=False
End Sub

Sub ProcessCSVData(ws As Worksheet, cityPopulationData As Collection, year As String)
    Dim city As String
    Dim population As Long
    Dim rowData() As String
    Dim rowIndex As Long, lastRow As Long
    Dim cityData As Collection
    Dim cityEntry As Variant
    Dim cityFound As Boolean

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Loop through populated rows
    For rowIndex = 2 To lastRow ' Start from row 2 to skip the header
        ' If row isn't empty and it contains value in the first column
        If Len(Trim(ws.Cells(rowIndex, 1).Value)) > 0 Then
            ' Split the row into parts by ";"
            rowData = Split(ws.Cells(rowIndex, 1).Value, ";")
            
            ' Ensure row contains valid data (city and population)
            If UBound(rowData) = 2 Then
                city = Trim(rowData(0))
                population = CLng(Trim(rowData(2)))
                
                ' Check if city already exists in collection
                cityFound = False
                
                For Each cityEntry In cityPopulationData
                    ' cityEntry(0) is the city name and cityEntry(1) is the population data collection
                    If cityEntry(0) = city Then
                        ' City found, add population to it
                        cityEntry(1).Add Array(year, population)
                        cityFound = True
                        Exit For
                    End If
                Next cityEntry
                
                ' If city doesn't exist in collection, create a new entry
                If Not cityFound Then
                    Set cityData = New Collection
                    cityData.Add Array(year, population) ' Add population data for this year
                    cityPopulationData.Add Array(city, cityData) ' Add the city and its population data collection
                End If
            End If
        End If
    Next rowIndex
End Sub

Sub WriteVäkiluvutData(cityPopulationData As Collection)
    Dim newWs As Worksheet
    Dim rowNum As Long
    Dim city As String
    Dim cityData As Variant
    Dim populationData As Variant
    
    ' Add new sheet "Väkiluvut" or use existing sheet with that name
    On Error Resume Next
    Set newWs = ThisWorkbook.Sheets("Väkiluvut")
    On Error GoTo 0
    
    If newWs Is Nothing Then
        Set newWs = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        newWs.Name = "Väkiluvut"
    End If
    
    ' Move data from collection to the new sheet
    newWs.Cells.Clear ' Clear sheet from existing data
    newWs.Cells(1, 1).Value = "Paikkakunta"
    newWs.Cells(1, 2).Value = "2022"
    newWs.Cells(1, 3).Value = "2023"
    newWs.Columns(1).ColumnWidth = 30  ' Paikkakunta width
    newWs.Columns(2).ColumnWidth = 10   ' 2022 width
    newWs.Columns(3).ColumnWidth = 10   ' 2023 width
    rowNum = 2 ' Start from row 2 for data
    
    ' Write data from collection to sheet
    For Each cityData In cityPopulationData
        city = cityData(0)
        newWs.Cells(rowNum, 1).Value = city
        
        ' Loop through the population data for each city and add the populations to columns
        For Each populationData In cityData(1) ' cityData(1) is the collection of population data
            If populationData(0) = "2022" Then
                newWs.Cells(rowNum, 2).Value = populationData(1)
            ElseIf populationData(0) = "2023" Then
                newWs.Cells(rowNum, 3).Value = populationData(1)
            End If
        Next populationData
        
        rowNum = rowNum + 1
    Next cityData
End Sub

Sub SaveVäkiluvutSheet()
    Dim newWs As Worksheet
    Dim folderpath As String, saveName As String
    Dim newWorkbook As Workbook

    folderpath = Application.ActiveWorkbook.Path
    ' Get the "Väkiluvut" worksheet
    On Error Resume Next
    Set newWs = ThisWorkbook.Sheets("Väkiluvut")
    On Error GoTo 0

    ' If the sheet exists, save it as a new file
    If Not newWs Is Nothing Then
        saveName = "Väkiluvut_" & Format(Now, "yyyy-mm-dd_hh-mm-ss") & ".xlsx"

        ' Copy the "Väkiluvut" sheet to a new workbook and save the new workbook
        newWs.Copy
        Set newWorkbook = ActiveWorkbook
        newWorkbook.SaveAs folderpath & "\" & saveName, FileFormat:=xlOpenXMLWorkbook ' Save as .xlsx
        newWorkbook.Close SaveChanges:=False
        
        ' Delete the "Väkiluvut" sheet from the current workbook
        Application.DisplayAlerts = False ' Disable confirmation alert for deleting sheet
        newWs.Delete
        Application.DisplayAlerts = True
    End If
End Sub

