Attribute VB_Name = "MonthlySummary"
'==========================================================
' 月次集計マクロ
' 日次データ（日付・カテゴリ・金額）から、月次×カテゴリのサマリー表を自動作成します。
' 想定用途：売上集計、経費集計、作業時間集計など
'==========================================================
Option Explicit
Sub CreateMonthlySummary()
Dim wsData As Worksheet
Dim wsSummary As Worksheet
Dim lastRow As Long
Dim i As Long
Dim dt As Date
Dim category As String
Dim amount As Double
Dim monthKey As String
Dim colIndex As Long
Dim rowIndex As Long
Set wsData = ThisWorkbook.Worksheets("Data")
Set wsSummary = ThisWorkbook.Worksheets("Summary")
wsSummary.Cells.Clear
lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
Dim dict As Object
Set dict = CreateObject("Scripting.Dictionary")
Dim categories As Object
Set categories = CreateObject("Scripting.Dictionary")
For i = 2 To lastRow
dt = wsData.Cells(i, 1).Value
category = wsData.Cells(i, 2).Value
amount = wsData.Cells(i, 3).Value
monthKey = Format(dt, "yyyy-mm")
Dim key As String
key = monthKey & "|" & category
If dict.Exists(key) Then
dict(key) = dict(key) + amount
Else
dict(key) = amount
End If
If Not categories.Exists(category) Then
categories.Add category, categories.Count + 2
End If
Next i
wsSummary.Cells(1, 1).Value = "月"
Dim catKey As Variant
For Each catKey In categories.Keys
wsSummary.Cells(1, categories(catKey)).Value = catKey
Next catKey
Dim months As Object
Set months = CreateObject("Scripting.Dictionary")
Dim k As Variant
For Each k In dict.Keys
Dim parts() As String
parts = Split(k, "|")
If Not months.Exists(parts(0)) Then
months.Add parts(0), months.Count + 2
wsSummary.Cells(months(parts(0)), 1).Value = parts(0)
End If
wsSummary.Cells(months(parts(0)), categories(parts(1))).Value = dict(k)
Next k
wsSummary.Columns.AutoFit
MsgBox "月次集計が完了しました。", vbInformation
End Sub
