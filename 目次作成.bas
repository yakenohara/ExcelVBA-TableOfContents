Attribute VB_Name = "�ڎ��쐬"
'
'�ڎ������
'�t�H�[�J�X���������Ă���Z�����������݊J�n�Z���Ƃ݂Ȃ��A
'�S�V�[�g���̃����N�t�����X�g�����܂�
'
Sub �ڎ��쐬()
    
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
    If ActiveWindow.SelectedSheets.Count > 1 Then
        MsgBox "�����V�[�g���I������Ă��܂�" & vbLf & _
               "�s�v�ȃV�[�g�I�����������Ă�������"
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    Set writePlace = Cells(Selection.Row, Selection.Column)
    numOfWorkSheets = ActiveWorkbook.Worksheets.Count
    
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

