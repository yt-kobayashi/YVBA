Attribute VB_Name = "YVBA_SetArrayElementModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_SetArrayElement
'|
'| Array関数を使いましょう．
'| 引数で渡した値を1次元配列に代入します．
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

Function YVBA_SetArrayElement(ParamArray arrayElement() As Variant) As Variant
    Dim resultArray() As Variant
    Dim element As Variant
    Dim count As Long
    
    ReDim resultArray(UBound(arrayElement) - LBound(arrayElement))
    count = 0
    
    For Each element In arrayElement
        resultArray(count) = element
        count = count + 1
    Next element
    
    YVBA_SetArrayElement = resultArray
End Function
