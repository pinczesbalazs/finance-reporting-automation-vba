Option Explicit

Sub GenerateFinanceReport()

    Dim wsData As Worksheet, wsReport As Worksheet, wsLog As Worksheet
    Dim lastRow As Long, reportRow As Long, logRow As Long, i As Long
    Dim department As String, costCenter As String, status As String
    Dim amount As Double
    Dim foundCell As Range

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Set wsData = ThisWorkbook.Sheets("FinanceData")
    Set wsReport = ThisWorkbook.Sheets("Report")
    Set wsLog = ThisWorkbook.Sheets("ErrorLog")

    wsReport.Cells.Clear
    wsLog.Cells.Clear

    wsReport.Range("A1:D1").Value = Array("Department", "Cost Center", "Total Amount", "Transaction Count")
    wsLog.Range("A1:C1").Value = Array("Row", "Issue", "Details")

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    logRow = 2

    For i = 2 To lastRow

        department = Trim(wsData.Cells(i, "A").Value)
        costCenter = Trim(wsData.Cells(i, "B").Value)
        status = Trim(wsData.Cells(i, "D").Value)

        If department = "" Then
            wsLog.Cells(logRow, "A").Value = i
            wsLog.Cells(logRow, "B").Value = "Missing department"
            wsLog.Cells(logRow, "C").Value = "Column A is empty"
            logRow = logRow + 1
            GoTo NextRow
        End If

        If costCenter = "" Then
            wsLog.Cells(logRow, "A").Value = i
            wsLog.Cells(logRow, "B").Value = "Missing cost center"
            wsLog.Cells(logRow, "C").Value = "Column B is empty"
            logRow = logRow + 1
            GoTo NextRow
        End If

        If Not IsNumeric(wsData.Cells(i, "C").Value) Then
            wsLog.Cells(logRow, "A").Value = i
            wsLog.Cells(logRow, "B").Value = "Invalid amount"
            wsLog.Cells(logRow, "C").Value = "Column C is not numeric"
            logRow = logRow + 1
            GoTo NextRow
        End If

        amount = CDbl(wsData.Cells(i, "C").Value)

        If amount <= 0 Then
            wsLog.Cells(logRow, "A").Value = i
            wsLog.Cells(logRow, "B").Value = "Invalid amount"
            wsLog.Cells(logRow, "C").Value = "Amount must be greater than 0"
            logRow = logRow + 1
            GoTo NextRow
        End If

        If status <> "Approved" And status <> "Pending" And status <> "Rejected" Then
            wsLog.Cells(logRow, "A").Value = i
            wsLog.Cells(logRow, "B").Value = "Invalid status"
            wsLog.Cells(logRow, "C").Value = "Status must be Approved, Pending or Rejected"
            logRow = logRow + 1
            GoTo NextRow
        End If

        Set foundCell = wsReport.Range("A:A").Find(department, LookAt:=xlWhole)

        If foundCell Is Nothing Then
            reportRow = wsReport.Cells(wsReport.Rows.Count, "A").End(xlUp).Row + 1

            wsReport.Cells(reportRow, "A").Value = department
            wsReport.Cells(reportRow, "B").Value = costCenter
            wsReport.Cells(reportRow, "C").Value = amount
            wsReport.Cells(reportRow, "D").Value = 1
        Else
            foundCell.Offset(0, 2).Value = foundCell.Offset(0, 2).Value + amount
            foundCell.Offset(0, 3).Value = foundCell.Offset(0, 3).Value + 1
        End If

NextRow:

    Next i

    wsReport.Columns("A:D").AutoFit
    wsLog.Columns("A:C").AutoFit

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    MsgBox "Finance report generated successfully!", vbInformation

End Sub
