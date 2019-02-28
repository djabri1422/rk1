Public Class Form1

    'Дан массив треугольников. Каждый треугольник задан 3 сторонами.
    'Отсортировать треугольники по периметрам(по возрастанию).
    'В отдельный массив перенести равносторонние треугольники. Вывести список таких треугольников во 2 форму.
    'Найти кол-во треугольников периметр которых больше среднего. 

    Dim i As Integer 'Я ленивый не хочу везде объявлять! Правда так не правильно делать!



    Public Structure Trio
        Public a, b, c As Integer
        Dim Sr As Single

        Public Sub Vvod_trio()
            a = Val(InputBox("Введите сторону а"))
            b = Val(InputBox("Введите сторону b"))
            c = Val(InputBox("Введите сторону c"))
        End Sub

        Public Function Stroka() As String
            Stroka = Str(a) + vbTab + Str(b) + vbTab + Str(c) + vbTab + Str(P)
        End Function

        Public Function P() As Integer
            P = a + b + c
        End Function

        Public Function Equal() As Boolean
            Equal = False
            If (a = b) And (a = c) Then
                Equal = True
            End If

        End Function

    End Structure

    'Проверка читают комментарии или нет, а то сладко живется! 
    'В программе специально допущены ошибки! Ищем и исправляем!  

    Public Sub Vvod(ByRef M() As Trio, ByVal n As Integer)
        Do
            n = Val(InputBox("Введите кол-во треугольников"))
        Loop Until n > 0
        n -= 1
        ReDim M(n)
        For i = 0 To n
            M(i).Vvod_trio()
        Next
    End Sub

    Public Sub Vivod(ByVal M() As Trio, ByVal n As Integer, ByRef Lst As ListBox)
        Lst.Items.Clear()
        For i = 0 To n
            Lst.Items.Add(M(i).Stroka())
        Next
    End Sub

    Public Sub Sort(ByVal M() As Trio, ByVal n As Integer)
        Dim Help_sort As Boolean
        Dim Tr As Trio
        Do
            Help_sort = True
            For i = 0 To n - 1
                If M(i).P() > M(i + 1).P() Then
                    Tr = M(i)
                    M(i) = M(i + 1)
                    M(i + 1) = Tr
                    Help_sort = False
                End If
            Next
        Loop Until Help_sort = True
    End Sub

    Public Sub New_Mas(ByVal M() As Trio, ByVal n As Integer, ByRef New_M() As Trio, ByRef k As Integer)
        k = -1
        For i = 0 To n
            If M(i).Equal = True Then
                k += 1
                ReDim Preserve New_M(k)
                New_M(k) = M(i)
            End If
        Next
    End Sub


    Public Function Srb(ByVal M() As Trio, ByVal n As Integer) As Integer
        Srb = 0
        For i = 0 To n
            Srb += M(i).P()
        Next
        Srb = Srb / (n + 1)
    End Function

    Public Function Kol_Up_Srb(ByVal M() As Trio, ByVal n As Integer) As Integer
        Dim sr As Single
        Kol_Up_Srb = 0
        sr = Srb(M, n)
        For i = 0 To n
            If M(i).P() > sr Then
                Kol_Up_Srb += 1
            End If
        Next
    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim M(), New_M() As Trio
        Dim n, k As Integer

        Vvod(M, n)
        Vivod(M, n, ListBox1)
        Sort(M, n)
        Vivod(M, n, ListBox2)
        Form2.Show()
        New_Mas(M, n, New_M, k)
        Vivod(New_M, k, Form2.ListBox1)
        TextBox1.Text = Str(Kol_Up_Srb(M, n))
    End Sub
End Class
