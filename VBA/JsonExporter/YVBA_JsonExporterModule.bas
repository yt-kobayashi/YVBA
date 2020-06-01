Attribute VB_Name = "YVBA_JsonExporterModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_JsonExporter
'|
'| 指定したシートの値をJson形式へ変換します．
'|
'| 注意事項
'|  出力できるのは入力されている値のみです．
'|  イメージとしてはCSVファイルのフォーマットがJsonへ変わっただけです．
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
'| YVBA_JsonExporter
'| [概要]
'|  指定された対象のワークブックもしくはワークシートをJson形式のテキストに変換します．
'|
'| [引数]
'|  target                         :   対象のワークブック，もしくはワークシート
'|  folderPath [省略可能]          :   出力先のフォルダ―パス
'+------------------------------------------------------------------------------------------------------------------------
Sub YVBA_JsonExporter(target As Object, Optional folderPath As String = "")
    Dim result As String
    Dim filePath As String
    Dim exportStream As Object
    Set exportStream = CreateObject("ADODB.Stream")
    
    
    If TypeName(target) = "Worksheet" Then
        result = "{" & YVBA_Worksheet2Json(target, True) & "}"
    ElseIf TypeName(target) = "Workbook" Then
        result = "{" & YVBA_Workbook2Json(target, True) & "}"
    Else
        MsgBox "targetにはワークブックもしくは，ワークシートを指定してください．"
        Exit Sub
    End If
    
    If folderPath = "" Then
        folderPath = ThisWorkbook.Path
    End If
    filePath = folderPath & "\" & target.Name & ".json"
    
    With exportStream
        .Charset = "UTF-8"
        .LineSeparator = adLF
        .Open
        .WriteText result
        .SaveToFile filePath, adSaveCreateOverWrite
        .Close
    End With
End Sub

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_Workbook2Json
'| [概要]
'|  指定された対象のワークブックをJson形式のテキストに変換します．
'|
'| [引数]
'|  targetBook                     :   対象のワークブック
'|  indention [省略可能]           :   改行の有効化/無効化
'|                                     True  ---- 有効
'|                                     False ---- 無効
'|
'| [戻値]
'|  変換後のJson形式の文字列
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_Workbook2Json(targetBook As Workbook, Optional indention As Boolean = False) As String
    Dim sheet As Worksheet
    Dim result As String
    Dim newLine As String
    
    result = ""
        
    newLine = ""
    If indention = True Then
        newLine = vbCrLf
    End If
    
    For Each sheet In targetBook.Sheets
        result = result & YVBA_Worksheet2Json(sheet, indention)
        
        If Not sheet.Name = targetBook.Worksheets(targetBook.Worksheets.count).Name Then
            result = Left(result, Len(result) - 2) & ","
        End If
        
        result = result
    Next sheet
    
    YVBA_Workbook2Json = result
End Function

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_Worksheet2Json
'| [概要]
'|  指定された対象のワークシートをJson形式のテキストに変換します．
'|
'| [引数]
'|  targetSheet                    :   対象のシート
'|  indention [省略可能]           :   改行の有効化/無効化
'|                                     True  ---- 有効
'|                                     False ---- 無効
'|
'| [戻値]
'|  変換後のJson形式の文字列
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_Worksheet2Json(targetSheet As Worksheet, Optional indention As Boolean = False) As String
    Dim targetArray As Variant
    Dim result As String
    Dim newLine As String
    result = ""
    newLine = ""
    
    
    If indention = True Then
        newLine = vbCrLf
    End If
    
    targetArray = targetSheet.Range("A1", YVBA_GetFinalCellPosition(targetSheet))
    
    result = newLine & """" & targetSheet.Name & """:" & YVBA_ConvertArray2Json(targetArray, indention) & newLine
    
    YVBA_Worksheet2Json = result
End Function

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_ConvertArray2Json
'| [概要]
'|  対象の2次元配列をJson形式に変換します．
'|
'| [引数]
'|  target                    :   変換対象の2次元配列
'|
'| [戻値]
'|  置換後のJson形式の文字列
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_ConvertArray2Json(target As Variant, Optional indention As Boolean = False) As String
    Dim singleArray As Variant
    Dim index As Variant
    Dim arrayIndex As Variant
    Dim tempResult As String
    Dim result As String
    Dim addTarget As Variant
    Dim newLine As String
    
    newLine = ""
    singleArray = target
    
    If indention = True Then
        newLine = vbCrLf
    End If
    
    If 1 < YVBA_GetArrayDimention(target) Then
        ReDim singleArray(LBound(target) To UBound(target))
        
        result = "[" & newLine
            
        For arrayIndex = LBound(target) To UBound(target)
            singleArray = WorksheetFunction.index(target, arrayIndex)
            
            tempResult = "["
            
            For index = LBound(singleArray) To UBound(singleArray)
                addTarget = singleArray(index)
                tempResult = tempResult & """" & YVBA_EscapeCharForJson(CStr(addTarget)) & """"
                
                If Not index = UBound(singleArray) Then
                    tempResult = tempResult & ","
                End If
            Next index
            
            tempResult = tempResult & "]"
            
            result = result & tempResult
            If Not arrayIndex = UBound(target) Then
                    result = result & "," & newLine
            End If
        Next arrayIndex
        
        result = result & newLine & "]"
    End If
    
    YVBA_ConvertArray2Json = result
End Function

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_EscapeCharForJson
'| [概要]
'|  対象文字列に含まれるJsonファイルで特殊文字となる文字を置換します．
'|
'| [引数]
'|  targetString                    :   変換対象の文字列
'|
'| [戻値]
'|  置換後の文字列
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_EscapeCharForJson(targetString As String) As String
    Static escapeList As Object
    Dim result As String
    Dim escapeChar As Variant
    
    If escapeList Is Nothing Then
        Set escapeList = CreateObject("Scripting.Dictionary")
    
        With escapeList
            .Add "\", "\\"
            .Add """", "\"""
            .Add "/", "\/"
            .Add vbBack, "\b"
            .Add vbFormFeed, "\f"
            .Add vbLf, "\n"
            .Add vbCr, "\r"
            .Add vbTab, "\t"
        End With
    End If
    
    result = targetString
    
    For Each escapeChar In escapeList
        result = Replace(result, escapeChar, escapeList.item(escapeChar))
    Next escapeChar
    
    YVBA_EscapeCharForJson = result
End Function
