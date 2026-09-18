Attribute VB_Name = "ReportGenerator"
'==========================================================
' 帳票自動作成マクロ
' 元データ（1行1件）をもとに、決まったテンプレートへ自動転記し、
' 1件ごとにシート（または別ブック）として帳票を出力します。
' 想定用途：請求書・納品書・発注書などの定型帳票作成
'==========================================================
Option Explicit
Sub GenerateReports()
Dim wsData As Worksheet
Dim wsTemplate As Worksheet
Dim lastRow As Long
Dim i As Long
Dim wsNew As Worksheet
Set wsData = ThisWorkbook.Worksheets("Data")
Set wsTemplate = ThisWorkbook.Worksheets("Template")
lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
Application.ScreenUpdating = False
Application.DisplayAlerts = False
For i = 2 To lastRow
wsTemplate.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
Set wsNew = ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
wsNew.Name = "帳票_" & i - 1
wsNew.Range("B2").Value = wsData.Cells(i, 1).Value
wsNew.Range("B3").Value = wsData.Cells(i, 2).Value
wsNew.Range("B4").Value = wsData.Cells(i, 3).Value
wsNew.Range("B5").Value = wsData.Cells(i, 4).Value
wsNew.Range("B6").Value = wsData.Cells(i, 3).Value * wsData.Cells(i, 4).Value
Next i
Application.DisplayAlerts = True
Application.ScreenUpdating = True
MsgBox "帳票作成が完了しました。作成件数: " & lastRow - 1, vbInformation
End Sub
