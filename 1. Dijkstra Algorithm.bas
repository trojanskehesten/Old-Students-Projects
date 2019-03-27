Attribute VB_Name = "Module1"
Option Base 1

Dim i As Integer, j As Integer, k As Integer, x As Integer '��������
Dim intSzGr As Integer, intSzVer As Integer '������ ������ ����� � ������
Dim blnBlockGr() As Boolean '������� �� ����� ���������� �������
Dim intPr As Integer '������ ������ ����������������� �������
Dim intRowPr() As Integer '������� ������������������ �������, ��������: ������ ������� �����

Const cInf As Double = 1E+308 '�������� ������������������, ���������

Sub main() '������� ���������. ��������� ��������� ��������� � ���������� �������

Call Clean '������� ������� ������ � ���������� �� ��������� ������ �� ���

Call AddVertex   '������� ������� ������

Call Inf_Weight  '����������� ����������� ���� ���� ��������, ����� ��������

Call AddGraphStart '������� ���� ��������� �������

Call AddGraph  '������� ���� ��������� ������

Call Design '��������� ���� Excel

End Sub

Sub Clean() '������� ������� ������ � ���������� �� ��������� ������ �� ���

i = 2

Do Until Cells(i, 8) = ""

    i = i + 1

Loop

Range(Cells(2, 8), Cells(i, 10)).ClearContents

'������� ��������� ������� � ��������� ������� � �����������
i = 2

Do Until Cells(i, 1) = ""

    Cells(i, 1) = Trim(Cells(i, 1))
    Cells(i, 2) = Trim(Cells(i, 2))
        
     i = i + 1

Loop

End Sub

Sub AddVertex() '������� ������� ������

Dim blnBlock As Boolean
'blnBlock - ��������� ������ ���������� ��������� � ������� ������

'intSzVer - ������� ������ ������� ������
'intSzGr - ������� ������ ������� �����
'i � j - ����� ������ � ������� �������� ������� �����
'k - ����� ������ �������� ������� ������

intSzVer = 2
blnBlock = False

For j = 1 To 2
i = 2 '��������� ������ ��������� � ������ ������

    Do Until Cells(i, j) = ""
    
        '�������� �� ��������� � �������� �������
        If Cells(2, 5) = Cells(i, j) Then GoTo Block1
                
        '�������� �� ������� ������ � ������� ������
        For k = 2 To intSzVer - 1 '��������� ������ ��������� � ������ ������

        If Cells(k, 8) = Cells(i, j) Then
        blnBlock = True
        Exit For
        End If
 
        Next k

    If blnBlock = False Then
    Cells(intSzVer, 8) = Cells(i, j)
    intSzVer = intSzVer + 1
    End If

    blnBlock = False
    
Block1:
    i = i + 1

    Loop

Next j

intSzVer = intSzVer - 2
i = i - 2 '�������� �� ������ �������� �������� �������� � �� 1�� ������ ���������
intSzGr = i

'���������� ��������� ������� ������
Range(Cells(2, 8), Cells(intSzVer + 1, 8)).Sort _
key1:=Range(Cells(2, 8), Cells(intSzVer + 1, 8)), _
order1:=xlAscending, Header:=xlNo

End Sub

Sub Inf_Weight() '����������� ����������� ���� ���� ��������, ����� ��������

For i = 2 To intSzVer + 1
    Cells(i, 9) = cInf
Next i

End Sub

Sub AddGraphStart() '������� ���� ��������� �������

intPr = 0

ReDim blnBlockGr(intSzGr) As Boolean
ReDim intRowPr(intSzVer) As Integer

'����� ��������� ������� � ������� ��������� �������
For i = 2 To intSzGr + 1
    If Cells(i, 1) = Cells(2, 5) Then
        For k = 2 To intSzVer + 1
            If Cells(i, 2) = Cells(k, 8) Then
                If Cells(i, 3) < Cells(k, 9) Then '��������� �������: ����� ������� < �������� ���������� �� ��������� ������� �� ������
                    
                    Cells(k, 9) = Cells(i, 3) '����������� ��� ������ ����� �� ���������� ������ �� �������
                    
                    If Len(Cells(2, 5)) > 4 Then '��������� ����, �������� �������� ������ �� ������������� �� ���� ������ ����
                        If Len(Cells(i, 2)) > 4 Then
                            Cells(k, 10) = Left(Cells(2, 5), 3) & ". - " & Left(Cells(i, 2), 3) & "."
                            Else
                            Cells(k, 10) = Left(Cells(2, 5), 3) & ". - " & Cells(i, 2)
                        End If
                    Else
                        If Len(Cells(i, 2)) > 4 Then
                            Cells(k, 10) = Cells(2, 5) & " - " & Left(Cells(i, 2), 3) & "."
                            Else
                            Cells(k, 10) = Cells(2, 5) & " - " & Cells(i, 2)
                        End If
                    End If
                    
                    For x = 1 To intSzVer
                        If intRowPr(x) = k - 1 Then GoTo Block2:
                    Next x
                    
                    intPr = intPr + 1
                    '����������� ������� �������������� �������
                    intRowPr(intPr) = k - 1
                    
Block2:
                    blnBlockGr(i - 1) = True
                    Exit For
                 End If
            blnBlockGr(i - 1) = True
            Exit For
            End If
        Next k
    End If
Next i

'����� ��������� ������� � ������� �������� �������
For i = 2 To intSzGr + 1
    If Cells(i, 2) = Cells(2, 5) Then
        For k = 2 To intSzVer + 1
            If Cells(i, 1) = Cells(k, 8) Then
                If Cells(i, 3) < Cells(k, 9) Then '��������� �������: ����� ������� < �������� ���������� �� ��������� ������� �� ������
                    
                    Cells(k, 9) = Cells(i, 3) '����������� ��� ������ ����� �� ���������� ������ �� �������
                    
                     If Len(Cells(2, 5)) > 4 Then '��������� ����, �������� �������� ������ �� ������������� �� ���� ������ ����
                        If Len(Cells(i, 1)) > 4 Then
                            Cells(k, 10) = Left(Cells(2, 5), 3) & ". - " & Left(Cells(i, 1), 3) & "."
                            Else
                            Cells(k, 10) = Left(Cells(2, 5), 3) & ". - " & Cells(i, 1)
                        End If
                    Else
                        If Len(Cells(i, 1)) > 4 Then
                            Cells(k, 10) = Cells(2, 5) & " - " & Left(Cells(i, 1), 3) & "."
                            Else
                            Cells(k, 10) = Cells(2, 5) & " - " & Cells(i, 1)
                        End If
                    End If
                    
                    For x = 1 To intSzVer
                        If intRowPr(x) = k - 1 Then GoTo Block3:
                    Next x
                    intPr = intPr + 1
                    '����������� ������� �������������� �������
                    intRowPr(intPr) = k - 1
                    
Block3:
                    blnBlockGr(i - 1) = True
                    Exit For
                End If
                blnBlockGr(i - 1) = True
                Exit For
            End If
        Next k
    End If
Next i
End Sub

Sub AddGraph()  '������� ���� ��������� ������

Dim qO As Integer, pO As Double, l As Integer

For l = 1 To intSzVer
qO = intRowPr(l)
pO = Cells(intRowPr(l) + 1, 9)

    For i = 2 To intSzGr + 1
        If blnBlockGr(i - 1) = False Then
            If Cells(i, 1) = Cells(qO + 1, 8) Then
                For k = 2 To intSzVer + 1
                    If Cells(i, 2) = Cells(k, 8) Then
                        If Cells(i, 3) + pO < Cells(k, 9) Then
                            Cells(k, 9) = Cells(i, 3) + pO
                            
                            If Len(Cells(i, 2)) > 4 Then
                                Cells(k, 10) = Cells(qO + 1, 10) & " - " & Left(Cells(i, 2), 3) & "."
                                Else
                                Cells(k, 10) = Cells(qO + 1, 10) & " - " & Cells(i, 2)
                            End If
                            
                            For x = 1 To intSzVer
                                If intRowPr(x) = k - 1 Then GoTo Block4:
                            Next x
                            intPr = intPr + 1
                            '����������� ������� �������������� �������
                            intRowPr(intPr) = k - 1
                    
Block4:
                            blnBlockGr(i - 1) = True
                            Exit For
                        End If
                        blnBlockGr(i - 1) = True
                        Exit For
                    End If
                Next k
            End If
            If Cells(i, 2) = Cells(qO + 1, 8) Then
                For k = 2 To intSzVer + 1
                    If Cells(i, 1) = Cells(k, 8) Then
                        If Cells(i, 3) + pO < Cells(k, 9) Then
                            Cells(k, 9) = Cells(i, 3) + pO
                            
                            If Len(Cells(i, 1)) > 4 Then
                                Cells(k, 10) = Cells(qO + 1, 10) & " - " & Left(Cells(i, 1), 3) & "."
                                Else
                                Cells(k, 10) = Cells(qO + 1, 10) & " - " & Cells(i, 1)
                            End If
                                
                            For x = 1 To intSzVer
                                If intRowPr(x) = k - 1 Then GoTo Block5:
                            Next x
                            intPr = intPr + 1
                            '����������� ������� �������������� �������
                            intRowPr(intPr) = k - 1
                    
Block5:
                            blnBlockGr(i - 1) = True
                            Exit For
                        End If
                        blnBlockGr(i - 1) = True
                        Exit For
                    End If
                Next k
            End If
        End If
    Next i
Next l

End Sub

Sub Design() '��������� ���� Excel

For i = 2 To intSzVer + 1 '���� ��� ��������� ������ ������������������, ����� "���������� �� ����������"
    If Cells(i, 9) = cInf Then
        Cells(i, 9) = "���������� �� ����������"
    End If
Next i

End Sub

