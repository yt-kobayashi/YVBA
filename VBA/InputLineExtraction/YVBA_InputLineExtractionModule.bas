Attribute VB_Name = "YVBA_InputLineExtractionModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_InputLineExtraction
'|
'| 指定した列で入力されている行だけを抽出し，別のワークシートにコピーするマクロです．
'| 一つのワークブック内でのコピペに対応しています．
'| 別のワークブックへのコピペは改造が必要です．
'|
'| それぞれの定数を希望のものに指定してください．
'+------------------------------------------------------------------------------------------------------------------------
'| MIT License
'|
'| Copyright (c) 2020 yt-kobayashi
'|
'| Permission is hereby granted, free of charge, to any person obtaining a copy
'| of this software and associated documentation files (the "Software"), to deal
'| in the Software without restriction, including without limitation the rights
'| to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'| copies of the Software, and to permit persons to whom the Software is
'| furnished to do so, subject to the following conditions:
'|
'| The above copyright notice and this permission notice shall be included in all
'| copies or substantial portions of the Software.
'|
'| THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'| IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'| FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'| AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'| LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'| OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'| SOFTWARE.
'+------------------------------------------------------------------------------------------------------------------------

Private Const originalWorksheetName As String = "リスト"            ' コピー元ワークシート名を指定する
Private Const pasteWorksheetName As String = "使用済みIPアドレス"   ' 貼り付け先のワークシート名を指定する
Private Const keyColumn As Long = 2                                 ' 入力されているか判別したい列を指定する(A=1, B=2, ... , Z=26)
Private Const pasteCellPosition As String = "A1"                    ' 貼り付け先のワークシートの座標を指定する

Sub YVBA_InputLineExtraction()
    Dim pasteWorksheet As Worksheet
    Dim originalWorksheet As Worksheet
    Dim rowLimit As Long
    
    Set originalWorksheet = ThisWorkbook.Worksheets(originalWorksheetName)
    Set pasteWorksheet = ThisWorkbook.Worksheets(pasteWorksheetName)
    pasteWorksheet.Cells.Clear
    

'   使用済みセル範囲を取得する
    With originalWorksheet.UsedRange
        rowLimit = .Rows(.Rows.Count).Row
    End With
    
'   空行でない行を抽出する
    With originalWorksheet
        .Range(.Cells(1, keyColumn), .Cells(rowLimit, keyColumn)).AutoFilter Field:=1, Criteria1:="<>"
        .Range(.Cells(1, keyColumn), .Cells(rowLimit, keyColumn)).CurrentRegion.Copy pasteWorksheet.Range(pasteCellPosition)
        .Range(.Cells(1, keyColumn), .Cells(rowLimit, keyColumn)).AutoFilter
    End With
End Sub

