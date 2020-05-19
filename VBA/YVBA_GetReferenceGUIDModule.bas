Attribute VB_Name = "YVBA_GetReferenceGUIDModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA YVBA_GetReferenceGUID
'|
'| 参照設定で追加/削除したいライブラリをそれぞれ指定することで，一括処理するマクロ．
'| 追加マクロと削除マクロがある．
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

Function YVBA_GetReferenceGUID(referenceList As Object) As Object
    Dim locator As Object
    Dim service As Object
    Dim referenceGUIDList As Object
    Dim registry As Variant
    Dim searchKey As Variant
    Dim typeLibKeys As Variant
    Dim typeLibSubKeys As Variant
    Dim referenceName As Variant
    Dim guid As String
    Dim version As Variant
    Dim count As Long
    Dim subCount As Long
    
    Set referenceGUIDList = CreateObject("System.Collections.ArrayList")
    Set locator = CreateObject("WbemScripting.SWbemLocator")
    Set service = locator.ConnectServer(vbNullString, "root\default")
    Set registry = service.get("StdRegProv")
    
    Const HKEY_CLASSES_ROOT = &H80000000
    Const REGISTRY_KEY As String = "TypeLib"
    
    searchKey = REGISTRY_KEY
    registry.EnumKey HKEY_CLASSES_ROOT, searchKey, typeLibKeys
    
    For count = LBound(typeLibKeys) To UBound(typeLibKeys)
        guid = typeLibKeys(count)
        searchKey = REGISTRY_KEY & "\" & guid
        registry.EnumKey HKEY_CLASSES_ROOT, searchKey, typeLibSubKeys
        
        If IsArray(typeLibSubKeys) Then
            For subCount = LBound(typeLibSubKeys) To UBound(typeLibSubKeys)
                version = typeLibSubKeys(subCount)
                searchKey = REGISTRY_KEY & "\" & guid & "\" & version
                registry.GetStringValue HKEY_CLASSES_ROOT, searchKey, "", referenceName
                
                If IsNull(referenceName) Then
                    GoTo CONTINUE
                End If
                
                version = Split(version, ".")
                
                If referenceList.Contains(referenceName) Then
                    If IsNumeric(version(0)) And IsNumeric(version(1)) Then
                        referenceGUIDList.Add Array(guid, CLng(version(0)), CLng(version(1)), referenceName)
                    End If
                End If
CONTINUE:
            Next subCount
        End If
    Next count
    
    Set locator = Nothing
    Set service = Nothing
    Set registry = Nothing
    
    Set YVBA_GetReferenceGUID = referenceGUIDList
End Function

Sub YVBA_PrintReferenceGUID()
    Dim locator As Object
    Dim service As Object
    Dim referenceGUIDList As Object
    Dim referenceList As Variant
    Dim registry As Variant
    Dim searchKey As Variant
    Dim typeLibKeys As Variant
    Dim typeLibSubKeys As Variant
    Dim referenceName As Variant
    Dim guid As String
    Dim version As Variant
    Dim count As Long
    Dim subCount As Long
    
    Set referenceGUIDList = CreateObject("System.Collections.ArrayList")
    Set locator = CreateObject("WbemScripting.SWbemLocator")
    Set service = locator.ConnectServer(vbNullString, "root\default")
    Set registry = service.get("StdRegProv")
    
    Const HKEY_CLASSES_ROOT = &H80000000
    Const REGISTRY_KEY As String = "TypeLib"
    
    searchKey = REGISTRY_KEY
    registry.EnumKey HKEY_CLASSES_ROOT, searchKey, typeLibKeys
    ReDim referenceList(UBound(typeLibKeys), 3)
    
    For count = LBound(typeLibKeys) To UBound(typeLibKeys)
        guid = typeLibKeys(count)
        searchKey = REGISTRY_KEY & "\" & guid
        registry.EnumKey HKEY_CLASSES_ROOT, searchKey, typeLibSubKeys
        
        
        If IsArray(typeLibSubKeys) Then
            For subCount = LBound(typeLibSubKeys) To UBound(typeLibSubKeys)
                version = typeLibSubKeys(subCount)
                searchKey = REGISTRY_KEY & "\" & guid & "\" & version
                registry.GetStringValue HKEY_CLASSES_ROOT, searchKey, "", referenceName
                
                If IsNull(referenceName) Then
                    GoTo CONTINUE
                End If
                
                version = Split(version, ".")
                referenceList(count, 0) = referenceName
                referenceList(count, 1) = guid
                referenceList(count, 2) = version(0)
                referenceList(count, 3) = version(1)

CONTINUE:
            Next subCount
        End If
    Next count
    
    Set locator = Nothing
    Set service = Nothing
    Set registry = Nothing
    
    ThisWorkbook.Worksheets.Add after:=Worksheets(Worksheets.count)
    ActiveSheet.Name = "Reference List"
    With ThisWorkbook.Worksheets("Reference List")
        .Range(.Cells(1, 1), .Cells(UBound(typeLibKeys), 4)) = referenceList
        Call .Range(.Cells(1, 1), .Cells(UBound(typeLibKeys), 4)).Sort(Range("A1"))
        .Columns("A:E").AutoFit
    End With
End Sub
