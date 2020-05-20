Attribute VB_Name = "YVBA_AddRemoveReferenceModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA YVBA_AddReference
'|      YVBA_RemoveReference
'|
'| 参照設定で追加/削除したいライブラリをそれぞれ指定することで，一括処理するマクロです．
'| 追加マクロと削除マクロがあります．
'|
'| 注意事項
'|  個人で作成したライブラリ等は非対応です．
'|  標準のライブラリしか対応していません．
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

Sub YVBA_AddReference()
    Dim referenceList As Object
    Dim referenceGUIDList As Object
    Set referenceList = CreateObject("System.Collections.ArrayList")
    Set referenceGUIDList = CreateObject("System.Collections.ArrayList")
    
'   ここに参照設定に追加したいライブラリ名を入れる
    With referenceList
        .Add "Microsoft Scripting Runtime"
        '.Add "Microsoft Visual Studio Project Model (10.0)"
        '.Add "Microsoft ActiveX Data Objects 6.1 Library"
    End With
    
    Set referenceGUIDList = YVBA_GetReferenceGUID(referenceList)
    
    On Error GoTo ErrorCheck:
    Dim referenceData As Variant
    Dim reference As Variant
    
    For Each referenceData In referenceGUIDList
        Set reference = ThisWorkbook.VBProject.References.AddFromGuid(referenceData(0), referenceData(1), referenceData(2))
    Next referenceData
    
SubExit:
    Set reference = Nothing
    Exit Sub
    
ErrorCheck:
    If Err.Number = 32813 Then
        Resume Next
    Else
        MsgBox "Error Number : " & Err.Number & vbCrLf & Err.Description
        GoTo SubExit:
    End If
End Sub

Sub YVBA_RemoveReference()
    Dim referenceList As Object
    Set referenceList = CreateObject("System.Collections.ArrayList")
    
'   ここに参照設定から削除したいライブラリ名を入れる
    With referenceList
        .Add "Microsoft Scripting Runtime"
        '.Add "Microsoft Visual Studio Project Model (10.0)"
        '.Add "Microsoft ActiveX Data Objects 6.1 Library"
    End With
    
    
    Dim reference As Variant
    Dim count As Long
    Dim filterRef As Variant

    With ThisWorkbook.VBProject
        For Each reference In .References
            If Not reference.IsBroken And referenceList.Contains(reference.Description) Then
                .References.Remove reference
            End If
        Next reference
    End With
End Sub
