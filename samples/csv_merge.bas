Attribute VB_Name = "CSVMerge"
'==========================================================
' CSV結合マクロ
' 指定フォルダ内の複数CSVファイルを1つのシートへ自動統合します。
' 想定用途：複数店舗・複数期間のデータをまとめて集計前処理する場合など
'==========================================================
Option Explicit
Sub MergeCSVFiles()
Dim targetFolder As String
Dim fileName As String
Dim wsOutput As Worksheet
Dim outputRow As Long
Dim wbSource As Workbook
Dim isFirstFile As Boolean
Set wsOutput = ThisWorkbook.Worksheets("Merged")
wsOutput.Cells.Clear
outputRow = 1
isFirstFile = True
With Application.FileDialog(msoFileDialogFolderPicker)
If .Show <> -1 Then Exit Sub
targetFolder = .SelectedItems(1) & "\"
End With
Application.ScreenUpdating = False
fileName = Dir(targetFolder & "*.csv")
Do While fileName <> ""
Set wbSource = Workbooks.Open(targetFolder & fileName)
Dim srcRange As Range
Dim startRow As Long
startRow = IIf(isFirstFile, 1, 2)
With wbSource.Worksheets(1)
Set srcRange = .Range(.Cells(startRow, 1), .Cells(.Cells(.Rows.Count, 1).End(xlUp).Row, .Cells(1, .Columns.Count).End(xlToLeft).Column))
End With
srcRange.Copy wsOutput.Cells(outputRow, 1)
outputRow = outputRow + srcRange.Rows.Count
wbSource.Close SaveChanges:=False
isFirstFile = False
fileName = Dir
Loop
Application.ScreenUpdating = True
MsgBox "CSV結合が完了しました。" & vbCrLf & "出力行数: " & outputRow - 1, vbInformation
End Sub
