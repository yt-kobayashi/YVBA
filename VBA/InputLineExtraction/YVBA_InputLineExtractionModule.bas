Attribute VB_Name = "YVBA_InputLineExtractionModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_InputLineExtraction
'|
'| �w�肵����œ��͂���Ă���s�����𒊏o���C�ʂ̃��[�N�V�[�g�ɃR�s�[����}�N���ł��D
'| ��̃��[�N�u�b�N���ł̃R�s�y�ɑΉ����Ă��܂��D
'| �ʂ̃��[�N�u�b�N�ւ̃R�s�y�͉������K�v�ł��D
'|
'| ���ꂼ��̒萔����]�̂��̂Ɏw�肵�Ă��������D
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

Private Const originalWorksheetName As String = "���X�g"            ' �R�s�[�����[�N�V�[�g�����w�肷��
Private Const pasteWorksheetName As String = "�g�p�ς�IP�A�h���X"   ' �\��t����̃��[�N�V�[�g�����w�肷��
Private Const keyColumn As Long = 2                                 ' ���͂���Ă��邩���ʂ���������w�肷��(A=1, B=2, ... , Z=26)
Private Const pasteCellPosition As String = "A1"                    ' �\��t����̃��[�N�V�[�g�̍��W���w�肷��

Sub YVBA_InputLineExtraction()
    Dim pasteWorksheet As Worksheet
    Dim originalWorksheet As Worksheet
    Dim rowLimit As Long
    
    Set originalWorksheet = ThisWorkbook.Worksheets(originalWorksheetName)
    Set pasteWorksheet = ThisWorkbook.Worksheets(pasteWorksheetName)
    pasteWorksheet.Cells.Clear
    

'   �g�p�ς݃Z���͈͂��擾����
    With originalWorksheet.UsedRange
        rowLimit = .Rows(.Rows.Count).Row
    End With
    
'   ��s�łȂ��s�𒊏o����
    With originalWorksheet
        .Range(.Cells(1, keyColumn), .Cells(rowLimit, keyColumn)).AutoFilter Field:=1, Criteria1:="<>"
        .Range(.Cells(1, keyColumn), .Cells(rowLimit, keyColumn)).CurrentRegion.Copy pasteWorksheet.Range(pasteCellPosition)
        .Range(.Cells(1, keyColumn), .Cells(rowLimit, keyColumn)).AutoFilter
    End With
End Sub

