Attribute VB_Name = "ModOutputPDF"
Option Explicit

'OutputPDF�E�E�E���ꏊ�FFukamiAddins3.ModFile

'------------------------------

'------------------------------


Sub OutputPDF(TargetSheet As Worksheet, Optional FolderPath$, Optional FileName$, _
              Optional MessageIrunaraTrue As Boolean = True)
'�w��V�[�g��PDF������
'20210721

'TargetSheet�E�E�EPDF������Ώۂ̃V�[�g
'FolderPath �E�E�E�o�͐�t�H���_ �w�肵�Ȃ��ꍇ�̓u�b�N�Ɠ����t�H���_
'FileName   �E�E�E�o��PDF�̃t�@�C���� �w�肵�Ȃ��ꍇ�̓V�[�g�̖��O
    
    '�����`�F�b�N
    If FolderPath = "" Then
        FolderPath = TargetSheet.Parent.Path '�w�肪�Ȃ��ꍇ�͎��u�b�N�̃t�H���_�p�X
    End If
    
    If FileName = "" Then
        FileName = TargetSheet.Name '�w�肪�Ȃ��ꍇ�̓V�[�g��
    End If
    
    '�o�͐�t�H���_���Ȃ��ꍇ�͍쐬����B
    If Dir(FolderPath, vbDirectory) = "" Then
        MkDir FolderPath
    End If
    
    '�o�͂���PDF�̃t�@�C�������쐬����
    Dim OutputFileName$
    OutputFileName = FolderPath & "\" & FileName & ".pdf"
    
    'PDF�ŏo�͂���
    TargetSheet.ExportAsFixedFormat Type:=xlTypePDF, FileName:=OutputFileName
    
    '�o�͌��ʂ̊m�F���b�Z�[�W
    If MessageIrunaraTrue Then
        If MsgBox("�u" & FileName & ".pdf" & "�v" & vbLf & "���쐬���܂���" & vbLf & _
            "�o�͐�t�H���_���N�����܂���?", vbYesNo + vbQuestion) = vbYes Then
            Shell "C:\Windows\explorer.exe " & FolderPath, vbNormalFocus
        End If
    End If
    
End Sub


