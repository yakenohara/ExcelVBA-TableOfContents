Attribute VB_Name = "TableOfContents"
'<License>------------------------------------------------------------
'
' Copyright (c) 2018 Shinnosuke Yakenohara
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <http://www.gnu.org/licenses/>.
'
'-----------------------------------------------------------</License>

'
'�ڎ������
'�t�H�[�J�X���������Ă���Z�����������݊J�n�Z���Ƃ݂Ȃ��A
'�S�V�[�g���̃����N�t�����X�g�����܂�
'
Sub TableOfContents()
    
    '�ϐ��錾
    Dim writePlace As Range
    Dim numOfWorkSheets As Long
    Dim cout As Long
    
    Dim cautionMessage As String: cautionMessage = "����Sub�v���V�[�W���́A" & vbLf & _
                                                   "���݂̑I��͈͂ɑ΂��Ēl�̏������݂��s���܂��B" & vbLf & vbLf & _
                                                   "���s���܂���?"
    
    '���s�m�F
    retVal = MsgBox(cautionMessage, vbOKCancel + vbExclamation)
    If retVal <> vbOK Then
        Exit Sub
        
    End If
    
    '�V�[�g�I����ԃ`�F�b�N
    If ActiveWindow.SelectedSheets.count > 1 Then
        MsgBox "�����V�[�g���I������Ă��܂�" & vbLf & _
               "�s�v�ȃV�[�g�I�����������Ă�������"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Set writePlace = Cells(Selection.Row, Selection.Column)
    numOfWorkSheets = ActiveWorkbook.Worksheets.count
    
    '�㏑���m�F
    If WorksheetFunction.CountA(Range(writePlace, Cells(writePlace.Row + numOfWorkSheets - 1, writePlace.Column))) > 0 Then
        yn = MsgBox("�쐬��̃Z���ɒl�������Ă��܂�" & vbLf & vbLf & _
                    "�㏑�����܂����H", _
                    vbOKCancel)
        
        If yn = vbCancel Then
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
    End If
    
    cout = 0
    For Each sh In ActiveWorkbook.Worksheets
        '�i���\��
        Application.StatusBar = "Progress:" & cout & "/" & numOfWorkSheets
    
        '�����𕶎���^�ɕύX
        writePlace.Clear
        writePlace.NumberFormatLocal = "@"
        
        '�n�C�p�[�����N�̍쐬
        ActiveSheet.Hyperlinks.Add _
                                Anchor:=writePlace, _
                                Address:="", _
                                SubAddress:="'" & sh.Name & "'!A1", _
                                TextToDisplay:="'" & sh.Name
                                
        '�������ݐ�Z���ʒu�̈ړ�
        Set writePlace = Cells(writePlace.Row + 1, writePlace.Column)
        
        cout = cout + 1
    Next sh
    
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox ("Done!")
    
End Sub

