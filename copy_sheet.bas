Attribute VB_Name = "Module1"
Sub CopySheet()
    Dim ws As Worksheet, ws1 As Worksheet
    ' �V�[�g�u�R�s�[���v�����݂��邩�m�F
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("�R�s�[��")
    Set ws1 = ThisWorkbook.Sheets("�R�s�[��")
    On Error GoTo 0
    ' �V�[�g�u�R�s�[���v�����݂���ꍇ�A�R�s�[���āu�R�s�[��v�Ƃ������O�ŕۑ�
    If Not ws Is Nothing Then
        Application.DisplayAlerts = False ' �x�����\���ɂ���
        ws.Copy Before:=ThisWorkbook.Sheets(1)
        Application.DisplayAlerts = True ' �x�����\���ɂ���
        ' �V�����쐬���ꂽ�V�[�g�ɖ��O��ݒ�
        If Not ws1 Is Nothing Then
            copyDate = Format(Now, "yyyymmdd_hhmmss")
            ActiveSheet.Name = "�R�s�[��_" & copyDate
            MsgBox "�V�[�g�u�R�s�[���v���u�R�s�[��_" & copyDate & "�v�Ƃ��ăR�s�[����܂����B"
        Else
            ActiveSheet.Name = "�R�s�[��"
            MsgBox "�V�[�g�u�R�s�[���v���u�R�s�[��v�Ƃ��ăR�s�[����܂����B"
        End If
    Else
        MsgBox "�V�[�g�u�R�s�[���v�����݂��܂���B"
    End If
End Sub
