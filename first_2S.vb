Public Class Form1


    'Дан список группы с ФИ и с k оценками. Найти Студента с макс. сред. баллом.
    'Составить список двойшников. Надо ввести кол-во оценок = b!!!!!
    'Проверка на пустоту массива! Или сделать k - глобальным и проверять его!!!

    '!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!ВНИМАТЕЛЬНЕЕ!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    'Я изменил код. Он отличается от того что мы писали на доске. 
    'Замена заключается в том, что мы сразу задаем кол-во оценок, а после постоянно отсылаемся на него.
    'В коде который был на доске, мы вводили количество оценок для каждлго студента по отдельности!
    'Название Функций, Процедур и Переменных могут не совпадать с тем что было написано на доске.




    Private Structure Stud 'Составляется для одного! После будет вызываться для каждого: Stud(i).Parameters
        Public First_name, Second_name As String
        Dim Group As String
        Public i, Points() As Integer 'Только динамический массив
        Dim Sr As Single

        Public Sub Vvod_S(ByVal k As Integer)
            Dim i As Integer
            First_name = InputBox("Введите Фамилию")
            Second_name = InputBox("Введите Имя")
            Group = InputBox("Введите группу")
            'k = Val(InputBox("Введите кол-во оценок"))
            ReDim Points(k - 1)
            For i = 0 To k - 1
                Points(i) = InputBox(Str(i + 1) + "-я оценка")
            Next
        End Sub

        Public Function Srb(ByVal k As Integer) As Single
            Dim i As Integer

            Srb = 0
            For i = 0 To k - 1
                Srb = Srb + Points(i)
            Next
            Srb = Srb / k
        End Function

        Public Function Search_Losers(ByVal k As Integer) As Boolean
            Dim i As Integer
            Search_Losers = False
            For i = 0 To k - 1
                If Points(i) = 2 Then
                    Search_Losers = True
                End If
            Next
        End Function

        Public Function Strr(ByVal k As Integer) As String
            Dim i As Integer
            Dim s As String
            s = LSet(First_name, 20) + LSet(Second_name, 15) + LSet(Group, 10) 'Формирует строку заданной длины
            For i = 0 To k - 1
                s += Str(Points(i)) + vbTab
            Next
            Return s
        End Function
    End Structure

    Private Sub Vvod(ByRef Mas() As Stud, ByRef n As Integer, ByRef k As Integer)
        Dim i As Integer
        Do
            n = Val(InputBox("Введите кол-во студентов"))
            k = Val(InputBox("Введите кол-во оценок"))
        Loop Until ((n > 0) And (k > 0))
        n -= 1
        ReDim Mas(n)
        For i = 0 To n
            Mas(i).Vvod_S(k)
        Next
    End Sub

    Private Sub Vivod(ByVal Mas() As Stud, ByVal n As Integer, ByVal k As Integer, ByRef lst As ListBox)
        Dim i As Integer
        lst.Items.Clear()
        For i = 0 To n
            lst.Items.Add(Mas(i).Strr(k))
        Next
    End Sub

    Private Function Best_stud(ByVal Mas() As Stud, ByVal n As Integer, ByVal k As Integer) As Stud
        Dim i As Integer
        Dim max As Stud

        max = Mas(0)
        For i = 0 To n
            If Mas(i).Srb(k) > max.Srb(k) Then
                max = Mas(i)
            End If
        Next
        'Поиск максимум обычный по среднему баллу
        Return max
    End Function


    Private Sub List_Losers(ByVal Mas() As Stud, ByVal n As Integer, ByVal k As Integer, ByRef List() As Stud)
        Dim i, kol As Integer
        kol = -1
        For i = 0 To n
            If Mas(i).Search_Losers(k) = True Then
                kol += 1
                ReDim Preserve List(kol)
                List(kol) = Mas(i)
            End If
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Mas(), List() As Stud
        Dim n, L, k As Integer
        Vvod(Mas, n, k)
        Vivod(Mas, n, k, ListBox1)
        TextBox1.Text = Best_stud(Mas, n, k).Strr(k)
        List_Losers(Mas, n, k, List)
        If List Is Nothing Then 'Список пуст(массив)
            MsgBox("Отсталых нету!")
        Else
            L = UBound(List)
            Vivod(List, L, k, ListBox2)
        End If

    End Sub
End Class
