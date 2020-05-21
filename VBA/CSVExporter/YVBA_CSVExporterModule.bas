Attribute VB_Name = "YVBA_CSVExporterModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA YVBA_CSVExporter
'|
'| 指定されたワークシートのセルをCSV形式でテキストファイルへ出力します．
'| ファイルはデフォルトでExcelファイルと同階層のディレクトリに作成されます．
'|
'| 注意事項
'|  YVBA_CommonModuleを読み込んでください．
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

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA DataSeparateType
'|
'| ファイル形式をEnum形式列挙する．
'+------------------------------------------------------------------------------------------------------------------------
Enum DataSeparateType
    CSV = 1
End Enum

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA YVBA_RunCSVExporter
'|
'| YVBA_CSVExporterの実行関数です．
'+------------------------------------------------------------------------------------------------------------------------
Sub YVBA_RunCSVExporter()
    YVBA_CSVExporter ThisWorkbook.Worksheets("規則的")
End Sub

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_GetFinalCellPosition
'| [概要]
'|  指定されたワークシートをCSV形式でテキストファイルへ出力します．
'|
'| [引数]
'|  targetSheet                     :   テキストファイルへ出力したいシート
'|  separator  [省略可能]           :   ファイル形式
'|  folderPath [省略可能]           :   フォルダパス
'|                                      省略時  :   マクロが含まれるExcelファイルのディレクトリを使用します．
'|                                      設定時  :   指定したディレクトリを使用します．
'+------------------------------------------------------------------------------------------------------------------------
Sub YVBA_CSVExporter(targetSheet As Worksheet, Optional separator As DataSeparateType = CSV, Optional folderPath As String = "")
    Dim value() As Variant
    Dim csvText As String
    Dim exportStream As Object
    Dim filePath As String
    Dim fileType As Variant
    Dim separatorChar As String
    
    fileType = YVBA_SelectSeparator(separator)
    separatorChar = fileType(0)
    Set exportStream = CreateObject("ADODB.Stream")
    
    If folderPath = "" Then
        folderPath = ThisWorkbook.Path
    End If
    filePath = folderPath & "\" & targetSheet.Name & fileType(1)
    
    YVBA_GetCellValues targetSheet, value
    YVBA_Join2D value, separatorChar, csvText
    
    With exportStream
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        .WriteText csvText
        .SaveToFile filePath, adSaveCreateOverWrite
        .Close
    End With
End Sub

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_GetCellValues
'| [概要]
'|  指定されたワークシートのセルに入力されている値を2次元配列に格納する．
'|
'| [引数]
'|  targetSheet                     :   2次元配列に格納したいシート
'|  result  [省略可能]              :   格納した2次元配列
'|
'| [戻値]
'|  格納した2次元配列
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_GetCellValues(targetSheet As Worksheet, Optional result As Variant) As Variant
    Dim finalCellPos As Variant
    
    finalCellPos = YVBA_GetFinalCellPosition(targetSheet)
    result = targetSheet.Range("A1", finalCellPos)
    
    YVBA_GetCellValues = result
End Function

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_SelectSeparator
'| [概要]
'|  区切り文字と拡張子を判別します．
'|  DataSeparateTypeにあるものだけが対象です．
'|
'| [引数]
'|  separator                       :   ファイル形式
'|
'| [戻値]
'|  区切り文字と拡張子を格納した配列 0:区切り文字 1:拡張子
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_SelectSeparator(separator As DataSeparateType) As Variant
    Dim separatorChar As Variant
    
    Select Case separator
    Case CSV
        separatorChar = Array(",", ".csv")
    End Select
    
    YVBA_SelectSeparator = separatorChar
End Function
