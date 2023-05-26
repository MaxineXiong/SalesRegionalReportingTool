Attribute VB_Name = "SalesRegionalReporting"
Option Explicit

Sub ClearWorkbook()
'SETP 0a: Initialize Workbook for new data analysis
    Dim TableTemp As ListObject

    'Clear NewData tab
    shNewDat.Cells.Clear
    
    'Clear workings tab
    Set TableTemp = shRegion.ListObjects("TableTemp")
    If TableTemp.ListRows.Count > 1 Then
        shRegion.Range(TableTemp.DataBodyRange(2, 1), _
                       TableTemp.DataBodyRange(TableTemp.ListRows.Count, TableTemp.ListColumns.Count)).Delete
    End If
                   
    'Clear task tracking table
    shStart.Range("B72", shStart.Range("B72").End(xlDown).Address).Clear
    
    MsgBox "Workbook cleared!", vbInformation, "Workbook Initialized"
    
End Sub



Sub Hide_Unhide_Sheets()
'STEP 0b: Admin access to hide/unhide other worksheets

    Dim InputPassword As String
    Dim myPassword As String
    Dim sh As Worksheet
    
    myPassword = "maxine"
    
    'Prompt the user to enter a password to toggle access
    InputPassword = InputBox("Enter password to toggle access")
    
    If InputPassword = "" Then
        Exit Sub
    ElseIf InputPassword <> myPassword Then
        MsgBox "Incorrect password entered. Please try again.", vbCritical, "Incorrect Password"
        Exit Sub
    Else
        Application.ScreenUpdating = False
        
        'If the shIntro sheet is visible, hide all other sheets except for Start and Workings
        If shIntro.Visible = xlSheetVisible Then
            For Each sh In ThisWorkbook.Sheets
                If sh.Name <> "Start" And sh.Name <> "Workings" Then
                    sh.Visible = xlSheetVeryHidden
                End If
            Next sh
        Else
            'If the shIntro sheet is not visible, unhide all sheets except for Start and Workings
            For Each sh In ThisWorkbook.Sheets
                If sh.Name <> "Start" And sh.Name <> "Workings" Then
                    sh.Visible = xlSheetVisible
                End If
            Next sh
        End If
            
        Application.ScreenUpdating = True
    End If
    
    shStart.Select

End Sub




Sub ImportData()
'STEP 1: Import data from multiple files into NewData tab

    Dim InputFiles As Variant
    Dim i As Byte
    Dim wb As Workbook
    Dim r As Byte
    Dim CompID As String
    Dim NewDataTable As ListObject
     
    On Error GoTo ErrorHandler
    
    'Prompt the user to select data files to import
    InputFiles = Excel.Application.GetOpenFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx", _
            Title:="Select Data Files to Import", MultiSelect:=True)
              
    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    If IsArray(InputFiles) Then
        'Open the first selected file
        Set wb = Workbooks.Open(InputFiles(1))
        
        'Set the first header as CompanyID
        shNewDat.Range("A1").Value = "CompanyID"
        
        'Paste headers from the first selected file into shNewDat
        wb.Sheets("Sales").Range("B4", wb.Sheets("Sales").Range("B4").End(xlToRight).Address).Copy
        shNewDat.Range("B1").PasteSpecial xlPasteValuesAndNumberFormats
        shNewDat.Range("A1", shNewDat.Range("A1").End(xlToRight).Address).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False
        wb.Close
    
        For i = 1 To UBound(InputFiles)
            'Open each of the selected workbook
            Set wb = Workbooks.Open(InputFiles(i))
            With wb.Sheets("Sales")
                'Copy data from each workbook to shNewDat
                .Range("B5", .Range("I5").End(xlDown).Address).Copy
                shNewDat.Range("B" & (WorksheetFunction.CountA(shNewDat.Range("B:B")) + 1)).PasteSpecial xlPasteAll
                CompID = .Range("C2").Value
            End With
            Application.CutCopyMode = False
            wb.Close
            
            'Populate the CompanyID column
            For r = (WorksheetFunction.CountA(shNewDat.Range("A:A")) + 1) To shNewDat.Range("B1").End(xlDown).row
                shNewDat.Range("A" & r).Value = CompID
            Next r
        Next i
    End If
    
    'Convert the new data range into table
    With shNewDat
        Set NewDataTable = .ListObjects.Add(SourceType:=xlSrcRange, Source:=.Range("A1").CurrentRegion, _
                                    xllistobjecthasheaders:=xlYes, tablestylename:="TableStyleLight1")
        NewDataTable.Name = "NewDataTable"
    End With
    
    'Delete Internal code column
    NewDataTable.ListColumns("Internal code").Delete
    
    MsgBox "Data imported successfully!", vbInformation, "Input Data Imported"
    
    'Register Task Tracking table
    shStart.Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = "New data imported on " & Time & " " & Date
    
    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
    
    shStart.Select
    
Exit Sub
ErrorHandler:
    Select Case Err.Number
        Case 9
            Exit Sub
        Case 1004
            MsgBox "Please clear the workbook before importing the new data!", vbExclamation, "Data Already Exists"
        Case Else
            MsgBox "An error has occured! Please try clear the workbook and restart the workflow.", vbExclamation, "ERROR"
    End Select
End Sub



Sub UpdateWorkings()
'STEP 2: Update all tables and charts on Workings tab
    Dim NewDataTable As ListObject
    Dim TableTemp As ListObject
    
    Application.ScreenUpdating = False
    
    Set NewDataTable = shNewDat.ListObjects("NewDataTable")
    Set TableTemp = shRegion.ListObjects("TableTemp")
    
    'Filter NewDataTable according selected region
    If shStart.AxRegion.Value = "America" Then
        NewDataTable.Range.AutoFilter Field:=1, Criteria1:="=*US", Operator:=xlAnd
    Else
        NewDataTable.Range.AutoFilter Field:=1, Criteria1:="<>*US", Operator:=xlAnd
    End If
    
    'Copy filtered data and paste to TableTemp
    NewDataTable.DataBodyRange.SpecialCells(xlCellTypeVisible).Copy
    shRegion.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats
    
    'Remove filter from NewDataTable
    NewDataTable.Range.AutoFilter Field:=1
    
    'Refresh all pivot tables in the workbook
    ThisWorkbook.RefreshAll
    
    MsgBox "Reporting tables and charts successfully updated!", vbInformation, "Reports Updated"
    
    'Register Task Tracking table
    shStart.Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = "Reports updated on " & Time & " " & Date
    
    Application.ScreenUpdating = True
    
End Sub


Sub RM_Report()
'STEP 3a: Save as Regional Manager report
    Dim RMpath As String
    Dim RMwb As Workbook
    Dim PT As PivotTable
    Dim PT1 As PivotTable
    Dim RMPivotCache As PivotCache
    Dim slicer As SlicerCache
    Dim i As Byte

    Application.ScreenUpdating = False
    
    'File path to Regional Management Report
    RMpath = ThisWorkbook.Path & "\RM_" & shRegion.Range("period").Value & shStart.AxRegion.Value & "_" _
             & Format(Date, "ddmmyyyy") & Format(Time, "hhmmss") & ".xlsx"
    
    'Copy shRegion sheet into RMwb
    Set RMwb = Workbooks.Add
    shRegion.Copy RMwb.Sheets(1)
    
    With RMwb
        'Exclude top management reporting part
        .Sheets(1).Range("AJ:AU").Delete
        
        'Disconnect the slicer control to the two pivot tables
        .SlicerCaches("Slicer_Company_Name").PivotTables.RemovePivotTable _
            (.Sheets(1).PivotTables("PTArticle"))
        .SlicerCaches("Slicer_Company_Name").PivotTables.RemovePivotTable _
            (.Sheets(1).PivotTables("PTReject"))
        
        'Change each pivot table's data source to the local TableTemp
        Set PT1 = .Sheets(1).PivotTables(1)
        PT1.ChangePivotCache .PivotCaches.Create(xlDatabase, "TableTemp")
        
        For Each PT In .Sheets(1).PivotTables
            If PT.Name <> PT1.Name Then
                PT.ChangePivotCache PT1.Name
            End If
        Next PT
        
        'Reconnect the slicer control to the two pivot tables
        .SlicerCaches("Slicer_Company_Name").PivotTables.AddPivotTable _
            (.Sheets(1).PivotTables("PTArticle"))
        .SlicerCaches("Slicer_Company_Name").PivotTables.AddPivotTable _
            (.Sheets(1).PivotTables("PTReject"))
        
        'Group and collapse columns A to Q to hide local TableTemp
        .Sheets(1).Range("A:Q").Group
        .Sheets(1).Outline.ShowLevels rowlevels:=0, columnlevels:=1
        
        'Rename the sheet by period
        .Sheets(1).Name = shRegion.Range("period").Value
        
        'Save and close the output workbook
        .SaveAs RMpath
        .Close
    End With
    
    MsgBox "A Regional Management Report successfully genderated!", vbInformation, "RM Report Created"
    
    'Register Task Tracking table
    shStart.Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = "Regional Report for " & shStart.AxRegion.Value & _
        " " & shRegion.Range("period").Value & " created on " & Time & " " & Date
        
    Application.ScreenUpdating = True
    
End Sub



Sub TM_Report()
'STEP 3b: Save as Top Manager report
    Dim TMPath As String
    Dim TMwb As Workbook
    
    Application.ScreenUpdating = False
    
    'File path to Top Management Report
    TMPath = ThisWorkbook.Path & "\TM_" & shRegion.Range("period").Value & shStart.AxRegion.Value & "_" _
             & Format(Date, "ddmmyyyy") & Format(Time, "hhmmss") & ".xlsx"
    
    Set TMwb = Workbooks.Add
    'Disable gridlines
    ActiveWindow.DisplayGridlines = False
    'Format header in TM report
    With TMwb.Sheets(1).Range("A1:C1")
        .Cells(1, 1).Value = "Regional Overview"
        .Font.ColorIndex = 2
        .Font.Bold = True
        .Font.Size = 20
        .RowHeight = 36
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 55
    End With
    
    'Copy PTManager into the new workbook as a hardcoded table
    shRegion.PivotTables("PTManager").TableRange1.Copy
    With TMwb.Sheets(1).Range("A3")
        .PasteSpecial xlPasteValuesAndNumberFormats
        .PasteSpecial xlPasteFormats
        .PasteSpecial xlPasteColumnWidths
    End With
    
    'Rename the sheet by period
    TMwb.Sheets(1).Name = shRegion.Range("period").Value
    'Save and close the output workbook
    TMwb.SaveAs TMPath
    TMwb.Close
    
    MsgBox "A Regional Overview Report for Top Management successfully genderated!", vbInformation, "TM Report Created"
    
    'Register Task Tracking table
    shStart.Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = "Regional Overview Report for Top Management - " & shStart.AxRegion.Value & _
        " " & shRegion.Range("period").Value & " created on " & Time & " " & Date
    
    Application.ScreenUpdating = True
End Sub




Sub ExportAsCSV()
'STEP 4: Export all data as CSV file with custom delimiter
    Dim sep As String
    Dim NewDataTable As ListObject
    Dim row As String
    Dim r As Long
    Dim c As Byte
    Dim CSVpath As String
    
    'Prompt the user to specify the delimiter for separating data in the CSV output
    sep = InputBox("Please specify the delimiter to separate data in the CSV output:", "Specify Delimiter", ",")
    
    'Create the file path for the CSV file based on the region, period, and timestamp
    CSVpath = ThisWorkbook.Path & "\" & shRegion.Range("period").Value & "_" & shStart.AxRegion.Value & "_" _
             & Format(Date, "ddmmyyyy") & Format(Time, "hhmmss") & ".csv"
    
    If sep <> "" Then
        Set NewDataTable = shNewDat.ListObjects("NewDataTable")
        'create and open a csv file
        Open CSVpath For Output As #1
            For r = 1 To NewDataTable.ListRows.Count
                row = ""
                For c = 1 To NewDataTable.ListColumns.Count
                    'Construct each row with the specified delimiter
                    row = row & sep & NewDataTable.DataBodyRange(r, c)
                Next c
                row = Right(row, Len(row) - 1)
                'write each row into the csv file
                Print #1, row
            Next r
        Close #1
    
    Else
        MsgBox "Invalid delimiter provided. Please try again!", vbExclamation, "Invalid Delimiter"
    End If
    
    MsgBox "A copy of CSV file successfully created!", vbInformation, "CSV saved"
    
    'Register Task Tracking table
    shStart.Range("B" & Rows.Count).End(xlUp).Offset(1, 0).Value = "CSV for the period " & shRegion.Range("period").Value & _
                  " created on " & Time & " " & Date

End Sub




