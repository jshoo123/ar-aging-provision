Attribute VB_Name = "Module_AR_Aging"
' ============================================================
' Accounts Receivable Aging - Bad Debt Provision Automation Macro
'
' How to use:
'   1. In Excel, press Alt+F11 (Mac: Option+F11) to open VBA editor
'   2. File > Import File > select this .bas file
'   3. Run in a workbook that has AR_Data and Aging_Policy sheets
'   4. Run macro (F5) or call Run_AR_Aging_Provision
'
' Required sheet/column layout:
'   [AR_Data]      A:CustomerName  B:IssueDate  C:Amount   (data starts row 2)
'   [Aging_Policy]  A:MinDays  B:BucketLabel  C:MaxDays(ref)  D:Rate  (data starts row 2, ascending by MinDays)
'
' NOTE: last-row detection for Aging_Policy uses column B, not column A,
' because column A may contain a stray note/comment below the table.
' ============================================================

Option Explicit

Sub Run_AR_Aging_Provision()

    Dim wsData As Worksheet
    Dim wsPolicy As Worksheet
    Dim wsResult As Worksheet

    Dim lastDataRow As Long
    Dim lastPolicyRow As Long
    Dim i As Long, j As Long

    Dim custName As String
    Dim issueDate As Date
    Dim amount As Double
    Dim agingDays As Long
    Dim provisionRate As Double
    Dim provisionAmt As Double

    Dim totalAmount As Double
    Dim totalProvision As Double

    On Error GoTo ErrHandler

    Set wsData = ThisWorkbook.Sheets("AR_Data")
    Set wsPolicy = ThisWorkbook.Sheets("Aging_Policy")

    ' Reset result sheet if it already exists
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("AR_Aging_Summary_VBA").Delete
    Application.DisplayAlerts = True
    On Error GoTo ErrHandler

    Set wsResult = ThisWorkbook.Sheets.Add(After:=wsPolicy)
    wsResult.Name = "AR_Aging_Summary_VBA"

    ' Header row
    With wsResult
        .Range("A1").Value = ChrW(44144) & ChrW(47000) & ChrW(52376) & ChrW(47749)
        .Range("B1").Value = ChrW(52292) & ChrW(44428) & ChrW(48156) & ChrW(49373) & ChrW(51068)
        .Range("C1").Value = ChrW(52292) & ChrW(44428) & ChrW(44552) & ChrW(50529)
        .Range("D1").Value = ChrW(44221) & ChrW(44284) & ChrW(51068) & ChrW(49688)
        .Range("E1").Value = ChrW(45824) & ChrW(49552) & ChrW(50984)
        .Range("F1").Value = ChrW(45824) & ChrW(49552) & ChrW(52649) & ChrW(45817) & ChrW(44552)
        .Range("A1:F1").Font.Bold = True
        .Range("A1:F1").Interior.Color = RGB(112, 173, 71)
        .Range("A1:F1").Font.Color = RGB(255, 255, 255)
    End With

    lastDataRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    ' Use column B (not A) to find the last policy row: column A may hold
    ' a stray note below the table that would otherwise be misread as a row.
    lastPolicyRow = wsPolicy.Cells(wsPolicy.Rows.Count, "B").End(xlUp).Row

    Dim outRow As Long
    outRow = 2
    totalAmount = 0
    totalProvision = 0

    For i = 2 To lastDataRow
        custName = wsData.Cells(i, 1).Value
        issueDate = wsData.Cells(i, 2).Value
        amount = wsData.Cells(i, 3).Value

        agingDays = Date - issueDate

        provisionRate = 0
        For j = 2 To lastPolicyRow
            If IsNumeric(wsPolicy.Cells(j, 1).Value) Then
                If agingDays >= wsPolicy.Cells(j, 1).Value Then
                    provisionRate = wsPolicy.Cells(j, 4).Value
                Else
                    Exit For
                End If
            End If
        Next j

        provisionAmt = amount * provisionRate

        wsResult.Cells(outRow, 1).Value = custName
        wsResult.Cells(outRow, 2).Value = issueDate
        wsResult.Cells(outRow, 2).NumberFormat = "yyyy-mm-dd"
        wsResult.Cells(outRow, 3).Value = amount
        wsResult.Cells(outRow, 3).NumberFormat = "#,##0"
        wsResult.Cells(outRow, 4).Value = agingDays
        wsResult.Cells(outRow, 5).Value = provisionRate
        wsResult.Cells(outRow, 5).NumberFormat = "0.0%"
        wsResult.Cells(outRow, 6).Value = provisionAmt
        wsResult.Cells(outRow, 6).NumberFormat = "#,##0"

        totalAmount = totalAmount + amount
        totalProvision = totalProvision + provisionAmt

        outRow = outRow + 1
    Next i

    wsResult.Cells(outRow, 1).Value = ChrW(54633) & ChrW(44228)
    wsResult.Cells(outRow, 1).Font.Bold = True
    wsResult.Cells(outRow, 3).Value = totalAmount
    wsResult.Cells(outRow, 3).NumberFormat = "#,##0"
    wsResult.Cells(outRow, 3).Font.Bold = True
    wsResult.Cells(outRow, 6).Value = totalProvision
    wsResult.Cells(outRow, 6).NumberFormat = "#,##0"
    wsResult.Cells(outRow, 6).Font.Bold = True

    wsResult.Columns("A:F").AutoFit

    MsgBox ChrW(45824) & ChrW(49552) & ChrW(52649) & ChrW(45817) & ChrW(44552) & " " & ChrW(51088) & ChrW(46041) & ChrW(49444) & ChrW(51221) & ChrW(51060) & " " & ChrW(50756) & ChrW(47308) & ChrW(46104) & ChrW(50632) & ChrW(49845) & ChrW(45768) & ChrW(45796) & "." & vbCrLf & _
           ChrW(52509) & " " & ChrW(45824) & ChrW(49552) & ChrW(52649) & ChrW(45817) & ChrW(44552) & ": " & Format(totalProvision, "#,##0") & ChrW(50896), vbInformation

    Exit Sub

ErrHandler:
    MsgBox ChrW(47588) & ChrW(53356) & ChrW(47196) & " " & ChrW(49892) & ChrW(54665) & " " & ChrW(51473) & " " & ChrW(50724) & ChrW(47448) & ChrW(44032) & " " & ChrW(48156) & ChrW(49373) & ChrW(54664) & ChrW(49845) & ChrW(45768) & ChrW(45796) & ": " & Err.Description, vbCritical
End Sub
