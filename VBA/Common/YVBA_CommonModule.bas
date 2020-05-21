Attribute VB_Name = "YVBA_CommonModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA YVBA_Common
'|
'| �l�I�ɋ��ʂŎg������֗��Ȋ֐����܂Ƃ߂Ă��܂��D
'|
'| ���ӎ���
'|  ����֐������܂Ƃ߂Ă��܂���D
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
'| YVBA_GetFinalCellPosition
'| [�T�v]
'|  �g�p�ςݍŏI�Z�����W��Range�^�Ŏ擾���܂��D
'|
'| [����]
'|  targetSheet                     :   �Z�����W���擾�������V�[�g
'|  positionVariable [�ȗ��\]     :   ���W���i�[�������ϐ�
'|  cellsType        [�ȗ��\]     :   Cells�^�̕\���`���̗L��/������ݒ肷��ϐ�
'|                                      True  ---- Cells�^�ŃZ�����W��Ԃ�
'|                                      False ---- Range�^�ŃZ�����W��Ԃ� [�f�t�H���g]
'|
'| [�ߒl]
'|  �Z�����W
'|  Cells�^ Long�^�z�� 0:�s�ԍ�(row) 1:��ԍ�(col)
'|  Range�^ String�^
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_GetFinalCellPosition(targetSheet As Worksheet, Optional positionVariable As Variant, Optional cellsType As Boolean = False) As Variant
    Dim cellPosition() As Variant
    Dim row As Long
    Dim col As Long
    
    ' �g�p�ςݍŏI�Z�����W���擾����
    With targetSheet.UsedRange
        row = .Rows(.Rows.Count).row
        col = .Columns(.Columns.Count).Column
    End With
    
    ' �\���`���I��
    If cellsType Then
        positionVariable = Array(row, col)
    Else
        positionVariable = YVBA_ConvertCells2Range(row, col)
    End If
    
    YVBA_GetFinalCellPosition = positionVariable
End Function

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_ConvertCells2Range
'| [�T�v]
'|  �Z�����W��Cells�^����Range�^�֕ϊ����܂��D
'|  ( 1,  1) ---> "A1"
'|  (13, 14) ---> "M14"
'|
'| [����]
'|  rowValue                        :   �s�ԍ�
'|  colvalue                        :   ��ԍ�
'|  positionVariable  [�ȗ��\]    :   ���W���i�[�������ϐ�
'|
'| [�ߒl]
'|  �Z�����W
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
    
    ' ��ԍ�����A1�Q�Ǝ��ł̃A���t�@�x�b�g���̌������Z�o����D
    Do
        inspectDigit = (col / 26)
        col = CLng(inspectDigit)
        digitNumber = digitNumber + 1
    Loop While 1 < inspectDigit
    
    ReDim a1Pos(0 To digitNumber)
    col = colValue
    
    ' �s�ԍ�����e�����̃A���t�@�x�b�g�ԍ����Z�o����D
    For digitLoop = digitNumber To 1 Step -1
        baseNumber = 26 ^ (digitLoop - 1)
        a1Pos(digitNumber - digitLoop) = Chr(64 + Fix(col / baseNumber))
        col = ((col - 1) Mod baseNumber) + 1
    Next digitLoop
    
    a1Pos(digitNumber) = rowValue
    positionVariable = Join(a1Pos, "")
    
    YVBA_ConvertCells2Range = positionVariable
End Function

