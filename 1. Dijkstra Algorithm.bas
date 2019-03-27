Attribute VB_Name = "Module1"
Option Base 1

Dim i As Integer, j As Integer, k As Integer, x As Integer 'Счетчики
Dim intSzGr As Integer, intSzVer As Integer 'Размер матриц графа и вершин
Dim blnBlockGr() As Boolean 'Удаляет из графа пройденную вершину
Dim intPr As Integer 'Размер матриц последоватльности звеньев
Dim intRowPr() As Integer 'Матрица последовательности звеньев, элементы: строки матрицы графа

Const cInf As Double = 1E+308 'Значение квазибесконечности, константа

Sub main() 'Главная процедура. Запускает остальные процедуры в правильном порядке

Call Clean 'Очищает матрицу вершин и расстояний от исходного пункта до них

Call AddVertex   'Создает матрицу вершин

Call Inf_Weight  'Присваивает бесконечные веса всем вершинам, кроме исходной

Call AddGraphStart 'Создает граф начальной вершины

Call AddGraph  'Создает граф остальных вершин

Call Design 'Оформляет лист Excel

End Sub

Sub Clean() 'Очищает матрицу вершин и расстояний от исходного пункта до них

i = 2

Do Until Cells(i, 8) = ""

    i = i + 1

Loop

Range(Cells(2, 8), Cells(i, 10)).ClearContents

'Убирает случайные пробелы в названиях пунктов и расстояниях
i = 2

Do Until Cells(i, 1) = ""

    Cells(i, 1) = Trim(Cells(i, 1))
    Cells(i, 2) = Trim(Cells(i, 2))
        
     i = i + 1

Loop

End Sub

Sub AddVertex() 'Создает матрицу вершин

Dim blnBlock As Boolean
'blnBlock - запрещает повтор одинаковых элементов в матрице вершин

'intSzVer - текущий размер матрицы вершин
'intSzGr - текущий размер матрицы графа
'i и j - номер строки и столбца элемента матрицы графа
'k - номер строки элемента матрицы вершин

intSzVer = 2
blnBlock = False

For j = 1 To 2
i = 2 'Заголовки таблиц находятся в первой строке

    Do Until Cells(i, j) = ""
    
        'Проверка на овпадение с исходным пунктом
        If Cells(2, 5) = Cells(i, j) Then GoTo Block1
                
        'Проверка на наличие пункта в матрице вершин
        For k = 2 To intSzVer - 1 'Заголовки таблиц находятся в первой строке

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
i = i - 2 'Поправка за лишнюю конечную итерацию счетчика и за 1ую строку заголовка
intSzGr = i

'сортировка элементов матрицы вершин
Range(Cells(2, 8), Cells(intSzVer + 1, 8)).Sort _
key1:=Range(Cells(2, 8), Cells(intSzVer + 1, 8)), _
order1:=xlAscending, Header:=xlNo

End Sub

Sub Inf_Weight() 'Присваивает бесконечные веса всем вершинам, кроме исходной

For i = 2 To intSzVer + 1
    Cells(i, 9) = cInf
Next i

End Sub

Sub AddGraphStart() 'Создает граф начальной вершины

intPr = 0

ReDim blnBlockGr(intSzGr) As Boolean
ReDim intRowPr(intSzVer) As Integer

'Поиск начальной вершины в колонке начальных пунктов
For i = 2 To intSzGr + 1
    If Cells(i, 1) = Cells(2, 5) Then
        For k = 2 To intSzVer + 1
            If Cells(i, 2) = Cells(k, 8) Then
                If Cells(i, 3) < Cells(k, 9) Then 'Проверяет условие: длина стороны < текущего расстояния от начальной вершины до данной
                    
                    Cells(k, 9) = Cells(i, 3) 'Присваивает вес равный длине от начального пункта до вершины
                    
                    If Len(Cells(2, 5)) > 4 Then 'Указывает путь, сокращая названия пункта по необходимости до трех первых букв
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
                    'Присваивает позицию использованной стороны
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

'Поиск начальной вершины в колонке конечных пунктов
For i = 2 To intSzGr + 1
    If Cells(i, 2) = Cells(2, 5) Then
        For k = 2 To intSzVer + 1
            If Cells(i, 1) = Cells(k, 8) Then
                If Cells(i, 3) < Cells(k, 9) Then 'Проверяет условие: длина стороны < текущего расстояния от начальной вершины до данной
                    
                    Cells(k, 9) = Cells(i, 3) 'Присваивает вес равный длине от начального пункта до вершины
                    
                     If Len(Cells(2, 5)) > 4 Then 'Указывает путь, сокращая названия пункта по необходимости до трех первых букв
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
                    'Присваивает позицию использованной стороны
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

Sub AddGraph()  'Создает граф остальных вершин

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
                            'Присваивает позицию использованной стороны
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
                            'Присваивает позицию использованной стороны
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

Sub Design() 'Оформляет лист Excel

For i = 2 To intSzVer + 1 'Если все получился равным квазибесконечности, пишет "Расстояние не определено"
    If Cells(i, 9) = cInf Then
        Cells(i, 9) = "Расстояние не определено"
    End If
Next i

End Sub

