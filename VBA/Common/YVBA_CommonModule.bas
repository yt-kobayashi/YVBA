Attribute VB_Name = "YVBA_CommonModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA YVBA_Common
'|
'| 個人的に共通で使えたら便利な関数をまとめています．
'|
'| 注意事項
'|  自作関数しかまとめていません．
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
'| Auto_Open
'| [概要]
'|  Excelファイルを開いたときに自動で実行される関数です．
'+------------------------------------------------------------------------------------------------------------------------
Public FOLDER_PATH As String
Sub Auto_Open()
    FOLDER_PATH = ThisWorkbook.Path
End Sub

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_GetFinalCellPosition
'| [概要]
'|  使用済み最終セル座標をRange型で取得します．
'|
'| [引数]
'|  targetSheet                     :   セル座標を取得したいシート
'|  positionVariable [省略可能]     :   座標を格納したい変数
'|  cellsType        [省略可能]     :   Cells型の表示形式の有効/無効を設定する変数
'|                                      True  ---- Cells型でセル座標を返す
'|                                      False ---- Range型でセル座標を返す [デフォルト]
'|
'| [戻値]
'|  セル座標
'|  Cells型 Long型配列 0:行番号(row) 1:列番号(col)
'|  Range型 String型
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_GetFinalCellPosition(targetSheet As Worksheet, Optional positionVariable As Variant, Optional cellsType As Boolean = False) As Variant
    Dim cellPosition() As Variant
    Dim row As Long
    Dim col As Long
    
    ' 使用済み最終セル座標を取得する
    With targetSheet.UsedRange
        row = .Rows(.Rows.Count).row
        col = .Columns(.Columns.Count).Column
    End With
    
    ' 表示形式選択
    If cellsType Then
        positionVariable = Array(row, col)
    Else
        positionVariable = YVBA_ConvertCells2Range(row, col)
    End If
    
    YVBA_GetFinalCellPosition = positionVariable
End Function

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_ConvertCells2Range
'| [概要]
'|  セル座標をCells型からRange型へ変換します．
'|  ( 1,  1) ---> "A1"
'|  (13, 14) ---> "M14"
'|
'| [引数]
'|  rowValue                        :   行番号
'|  colvalue                        :   列番号
'|  positionVariable  [省略可能]    :   座標を格納したい変数
'|
'| [戻値]
'|  セル座標
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_ConvertCells2Range(rowValue As Long, colValue As Long, Optional positionVariable As String) As String
    Dim inspectDigit As Double
    Dim col As Long
    Dim digitNumber As Integer
    Dim baseNumber As Long
    Dim digitLoop As Integer
    Dim a1Pos() As String
    Dim alphabetNumber As Long
    
    col = colValue
    digitNumber = 0
    
    ' 列番号からA1参照式でのアルファベット部の桁数を算出する．
    Do
        inspectDigit = (col / 26)
        col = CLng(inspectDigit)
        digitNumber = digitNumber + 1
    Loop While 1 < inspectDigit
    
    ReDim a1Pos(0 To digitNumber)
    col = colValue
    
    ' 行番号から各桁毎のアルファベット番号を算出する．
    For digitLoop = digitNumber To 1 Step -1
        baseNumber = 26 ^ (digitLoop - 1)
        a1Pos(digitNumber - digitLoop) = Chr(64 + Fix(col / baseNumber))
        col = ((col - 1) Mod baseNumber) + 1
    Next digitLoop
    
    a1Pos(digitNumber) = rowValue
    positionVariable = Join(a1Pos, "")
    
    YVBA_ConvertCells2Range = positionVariable
End Function

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_Join2D
'| [概要]
'|  2次元配列を文字列へ変換します．
'|
'| [引数]
'|  targetArray()           :   変換対象の2次元配列
'|  separator [省略可能]    :   区切り文字
'|  result    [省略可能]    :   変換後の文字列を格納する変数
'|
'| [戻値]
'|  変換後の文字列
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_Join2D(targetArray() As Variant, Optional separator As String = ",", Optional result As String) As String
    Dim suffix(1) As Long
    Dim joinLoop As Long
    Dim rowArray() As Variant
    Dim colArray() As Variant
    
    suffix(0) = UBound(targetArray, 1) - 1
    suffix(1) = UBound(targetArray, 2) - 1
    
    ReDim rowArray(suffix(0))
    ReDim colArray(suffix(1))
    
    For joinLoop = 0 To suffix(0)
        colArray = WorksheetFunction.Index(targetArray, joinLoop + 1)
        rowArray(joinLoop) = Join(colArray, separator)
    Next joinLoop
    
    result = Join(rowArray, vbCrLf)
    
    YVBA_Join2D = result
End Function
