Attribute VB_Name = "YVBA_CSVExporterModule"
Option Explicit

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA YVBA_CSVExporter
'|
'| �w�肳�ꂽ���[�N�V�[�g�̃Z����CSV�`���Ńe�L�X�g�t�@�C���֏o�͂��܂��D
'| �t�@�C���̓f�t�H���g��Excel�t�@�C���Ɠ��K�w�̃f�B���N�g���ɍ쐬����܂��D
'|
'| ���ӎ���
'|  YVBA_CommonModule��ǂݍ���ł��������D
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
'| �t�@�C���`����Enum�`���񋓂���D
'+------------------------------------------------------------------------------------------------------------------------
Enum DataSeparateType
    CSV = 1
End Enum

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA YVBA_RunCSVExporter
'|
'| YVBA_CSVExporter�̎��s�֐��ł��D
'+------------------------------------------------------------------------------------------------------------------------
Sub YVBA_RunCSVExporter()
    YVBA_CSVExporter ThisWorkbook.Worksheets("�K���I")
End Sub

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_GetFinalCellPosition
'| [�T�v]
'|  �w�肳�ꂽ���[�N�V�[�g��CSV�`���Ńe�L�X�g�t�@�C���֏o�͂��܂��D
'|
'| [����]
'|  targetSheet                     :   �e�L�X�g�t�@�C���֏o�͂������V�[�g
'|  separator  [�ȗ��\]           :   �t�@�C���`��
'|  folderPath [�ȗ��\]           :   �t�H���_�p�X
'|                                      �ȗ���  :   �}�N�����܂܂��Excel�t�@�C���̃f�B���N�g�����g�p���܂��D
'|                                      �ݒ莞  :   �w�肵���f�B���N�g�����g�p���܂��D
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
'| [�T�v]
'|  �w�肳�ꂽ���[�N�V�[�g�̃Z���ɓ��͂���Ă���l��2�����z��Ɋi�[����D
'|
'| [����]
'|  targetSheet                     :   2�����z��Ɋi�[�������V�[�g
'|  result  [�ȗ��\]              :   �i�[����2�����z��
'|
'| [�ߒl]
'|  �i�[����2�����z��
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_GetCellValues(targetSheet As Worksheet, Optional result As Variant) As Variant
    Dim finalCellPos As Variant
    
    finalCellPos = YVBA_GetFinalCellPosition(targetSheet)
    result = targetSheet.Range("A1", finalCellPos)
    
    YVBA_GetCellValues = result
End Function

'+------------------------------------------------------------------------------------------------------------------------
'| YVBA_SelectSeparator
'| [�T�v]
'|  ��؂蕶���Ɗg���q�𔻕ʂ��܂��D
'|  DataSeparateType�ɂ�����̂������Ώۂł��D
'|
'| [����]
'|  separator                       :   �t�@�C���`��
'|
'| [�ߒl]
'|  ��؂蕶���Ɗg���q���i�[�����z�� 0:��؂蕶�� 1:�g���q
'+------------------------------------------------------------------------------------------------------------------------
Function YVBA_SelectSeparator(separator As DataSeparateType) As Variant
    Dim separatorChar As Variant
    
    Select Case separator
    Case CSV
        separatorChar = Array(",", ".csv")
    End Select
    
    YVBA_SelectSeparator = separatorChar
End Function
