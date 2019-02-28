Public Class Form1

    Dim i As Integer
    Public Structure tovar
        Dim country, name As String
        Dim cost As Single
        Dim kol As Integer
        Public Sub vvod()
            name = InputBox("vvedite name")
            country = InputBox("vvedite country")
            cost = InputBox("vvedite cost")
            kol = InputBox("vvedite kol")
        End Sub
        Public Function vivod() As String
            Dim sum As Integer
            Dim s As String
            sum = dlisum()
            s = LSet(name, 30) + LSet(country, 15) + LSet(cost, 10) + LSet(kol, 3) + LSet(sum, 10)
            Return s
        End Function
        Public Function dlisum() As Single
            dlisum = cost * kol
        End Function
    End Structure

    Public Sub vvod1(ByRef M() As tovar, ByRef n As Integer)
        Do
            n = Val(InputBox("vvedite kol-vo tovarov"))
        Loop Until n > 0
        n -= 1
        ReDim M(n)
        For i = 0 To n
            M(i).vvod()
        Next
    End Sub
    Public Sub vivod1(ByVal M() As tovar, ByVal n As Integer, ByRef lst As ListBox)
        lst.Items.Clear()
        For i = 0 To n
            lst.Items.Add(M(i).vivod())
        Next
    End Sub
    Private Function stoim(ByVal M() As tovar, ByVal n As Integer) As Single
        stoim = 0
        For i = 0 To n
            stoim += M(i).dlisum()

        Next
    End Function
    Public Sub sort(ByVal M() As tovar, ByVal n As Integer)
        Dim P As tovar
        Dim help_sort As Boolean
        Do
            help_sort = True
            For i = 0 To n - 1
                If M(i).country > M(i + 1).country Then
                    P = M(i)
                    M(i) = M(i + 1)
                    M(i + 1) = P
                    help_sort = False
                End If
            Next
        Loop Until help_sort = True
    End Sub
    Public Sub naim(ByVal M() As tovar, ByVal n As Integer, ByRef B() As tovar, ByRef k As Integer, ByVal imya As String)
        k = -1
        For i = 0 To n
            If M(i).name = imya Then
                k += 1
                ReDim Preserve B(k)
                B(k) = M(i)
            End If
        Next
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim M(), B() As tovar
        Dim n, k As Integer
        Dim imya As String
        imya = InputBox("vvedite imya tovara")
        vvod1(M, n)
        vivod1(M, n, ListBox1)
        sort(M, n)
        textbox1.text = Str(stoim(M, n))
        vivod1(M, n, ListBox2)
        Form2.Show()
        naim(M, n, B, k, imya)
        If B Is Nothing Then
            MsgBox("massive is poost")
        Else
            vivod1(B, k, Form2.ListBox1)
        End If

    End Sub
End Class
