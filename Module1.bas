Attribute VB_Name = "Module1"
'================================
'
'file:  test
'Date:  2021/06/09
'�p�r:  ���V�[�g��������ɂ�������܂Ƃ߂�
'
'ver. 2.0
'================================

'�萔
Const SheetName As String = "Sheet4"            '<<�����ɁA�o�͂������V�[�g�̖��O������>>
Const require As Date = "2021/3/1  0:00:00"     '<<�����́A�����ADate�ɂȂ��Ă�̂ŏ����ɍ��킹�ĕϐ��錾���Ă��������B>>

'***********************
'���C��
'***********************
Sub SampleMain_onClick()
On Error GoTo ERROR_
    Dim A() As String   '�z��
    Dim B() As String   '�z��
    Dim C() As Date     '�z��
    Dim D() As String   '�z��
    
    Call syokika_onClick                        '�������A���[�U�[���\�b�h
    
    Application.ScreenUpdating = False          '��ʏ���STOP
    
    '�V�[�g1����X�^�[�g�A�A�z�z��
    For Each mySheet In Worksheets
        If Not mySheet.Name = SheetName Then                 '�����A�o�͎w��V�[�g�Ȃ��΂�
        
            Max_i = Worksheets(mySheet.Name).Range("A" & rows.count).End(xlUp).Row              '�I�����Ă�V�[�g�̍ő� �s ���擾
            Worksheets(mySheet.Name).Activate
            
            '�}��
            For i = 1 To (Max_i - 1)
            
                If require <= Cells(1, IndexUpdate(mySheet.Name, "date")).Value Then    '���� �G���[�̏ꍇ ""�̒��g����v���Ă邩�m�F���Ă݂�
                
                    If isArrayEx(A) = -1 Then               '�������̔���A���[�U�[���\�b�h
                        
                        ReDim Preserve A(0)                 '����������Ă��Ȃ��̂ŏ�����
                        ReDim Preserve B(0)                 '����������Ă��Ȃ��̂ŏ�����
                        ReDim Preserve C(0)                 '����������Ă��Ȃ��̂ŏ�����
                        ReDim Preserve D(0)                 '����������Ă��Ȃ��̂ŏ�����
                        
                    Else
                        
                        ReDim Preserve A(UBound(A) + 1)     '�z��𑝂₷
                        ReDim Preserve B(UBound(B) + 1)     '�z��𑝂₷
                        ReDim Preserve C(UBound(C) + 1)     '�z��𑝂₷
                        ReDim Preserve D(UBound(D) + 1)     '�z��𑝂₷
                        
                    End If
                    
                    A(UBound(A)) = Cells(i + 1, 1).Value    '�l��}�����Ă���
                    B(UBound(B)) = Cells(i + 1, 2).Value    '�l��}�����Ă���
                    C(UBound(C)) = Cells(i + 1, 3).Value    '�l��}�����Ă���
                    D(UBound(D)) = Cells(i + 1, 4).Value    '�l��}�����Ă���
                    
                End If
            Next i
        End If
    Next
    
    '�\��
    Application.ScreenUpdating = True       '��ʏ����J�n
    Worksheets(SheetName).Activate
            
    For i = 1 To (UBound(A))
        index = i
    
        Cells(index, 1).Value = A(i) '�l��Cell�ɑ}�����Ă���
        Cells(index, 2).Value = B(i) '�l��Cell�ɑ}�����Ă���
        Cells(index, 3).Value = C(i) '�l��Cell�ɑ}�����Ă���
        Cells(index, 4).Value = D(i) '�l��Cell�ɑ}�����Ă���
        
    Next i
    
ERROR_:
    If Err.Number = 1004 Then
        MsgBox "38�s��:IF�̎w�荀�ڂ����݂��Ă��Ȃ��\��������܂��B�ēx�m�F���Ă��������B", vbOKOnly
    Else
    'any error's
    End If
End Sub
'--------------------------------------------------------------
'
'syokika_onClick
'
'@param     void
'@return    void
'
'�ŏ��̎��s���Ɉ�x�������s�����B�Z�����܂�����ȏ�Ԃɖ߂��B
'
'--------------------------------------------------------------
Sub syokika_onClick()
    Application.ScreenUpdating = False          '��ʏ���STOP
    
    '������
    Max_i = Worksheets(SheetName).Range("A" & rows.count).End(xlUp).Row              '�I�����Ă�V�[�g�̍ő� �s ���擾
    Worksheets(SheetName).Activate
    
    For i = 1 To Max_i
        
        Cells(i, 1).Value = ""      '���������Ă���
        Cells(i, 2).Value = ""
        Cells(i, 3).Value = ""
        Cells(i, 4).Value = ""
        
    Next i
    
    Application.ScreenUpdating = True          '��ʏ���STOP
End Sub
'--------------------------------------------------------------
'
'IndexUpdate
'
'@param     String      text    //����
'@return    int         i       //���t�̈ʒu���擾
'
'������text�Ɠ������ڂ�����΁A�����̈ʒu�̒l��Ԃ��B�����łȂ���΃G���[���o�͂���B
'
'--------------------------------------------------------------
Function IndexUpdate(mySheet As String, text As String) As Integer
    Max_c = Worksheets(mySheet).Cells(1, Columns.count).End(xlToLeft).Column               '�I�����Ă�V�[�g�̍ő� �� ���擾
    
    For C = 1 To Max_c
        If text = Cells(1, C).Value Then
            IndexUpdate = C
            Exit Function
        End If
    Next C
    
    IndexUpdate = 0
End Function

'--------------------------------------------------------------
'WEB�T�C�g������p
'�ڂ����́Fhttps://zukucode.com/2019/08/vba-array-loop.html
'
'�@�\�F�������z�񂩔��肵�A�z��̏ꍇ�͋󂩂ǂ��������肷��
'�߂�l�F���茋�ʁi1:�z�� / 0:��̔z�� / -1:�z��ł͂Ȃ�
'--------------------------------------------------------------
Public Function isArrayEx(varArray As Variant) As Long
On Error GoTo ERROR_
    If IsArray(varArray) Then
        isArrayEx = IIf(UBound(varArray) >= 0, 1, 0)
    Else
        isArrayEx = -1
    End If

    Exit Function

ERROR_:
    If Err.Number = 9 Then
        isArrayEx = -1
    Else
        '�z��O�G���[
    End If
End Function
