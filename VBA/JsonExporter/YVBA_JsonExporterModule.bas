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
        result = Replace(result, escapeChar, escapeList.Item(escapeChar))
    Next escapeChar
    
    YVBA_EscapeCharForJson = result
End Function
