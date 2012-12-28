Imports System.Windows.Forms

Public Class TestForm

    Dim instrument As miSerialProtocol.miSerialProtocolClass
    Dim mini As MiniATClass
    Dim max As MiniMaxClass
    Dim ec As ECATClass
    Dim tci As TCIClass




    Private Sub TestForm_Dispose(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Disposed
        If Not instrument Is Nothing Then
            instrument.Dispose()
        End If

    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miniAtButton.Click

        Try
            If mini Is Nothing Then
                mini = New MiniATClass(6, MiniATClass.BaudRateEnum.b9600)
            End If

            AddHandler miSerialProtocolClass.CurrentItemProgress, AddressOf increaseProgress

            instrument = mini
            mini.Connect()
        Catch ex As CommInUseException
            mini.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)

            mini.Dispose()
            mini = Nothing
        End Try
    End Sub

    Private Sub miniMaxButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles miniMaxButton.Click

        Try
            If max Is Nothing Then
                max = New MiniMaxClass(6, MiniMaxClass.BaudRateEnum.b9600)
            End If
            AddHandler miSerialProtocolClass.CurrentItemProgress, AddressOf increaseProgress
            instrument = max
            max.Connect()
        Catch ex As CommInUseException
            MessageBox.Show(ex.Message)
            max.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            instrument.Disconnect()
            instrument = Nothing
        Catch ex As Exception
            Console.Write(ex.Message)
        End Try

    End Sub

    Private Sub WriteButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles WriteButton.Click
        Try
            instrument.WD(itemnumTextBox.Text, valueTextBox.Text)
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Private Sub ReadButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ReadButton.Click
        Try

            valueTextBox.Text = instrument.RD(itemnumTextBox.Text)

        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ECATButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ECATButton.Click
        Try
            If ec Is Nothing Then
                ec = New ECATClass(1, ECATClass.BaudRateEnum.b9600)
            End If
            instrument = ec
            ec.Connect()
        Catch ex As CommInUseException
            ec.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)

            ec.Dispose()
            ec = Nothing
        End Try
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Try
            If Not mini Is Nothing Then
                mini.ResetAlarms()
            ElseIf Not max Is Nothing Then
                max.ResetAlarms()
            ElseIf Not ec Is Nothing Then
                ec.ResetAlarms()
            ElseIf Not tci Is Nothing Then
                tci.ResetAlarms()
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Try
            If tci Is Nothing Then
                tci = New TCIClass(7, TCIClass.BaudRateEnum.b9600)
            End If
            instrument = tci
            tci.Connect()
        Catch ex As CommInUseException
            tci.Dispose()
        Catch ex As Exception
            MessageBox.Show(ex.Message)

            tci.Dispose()
            tci = Nothing
        End Try
    End Sub

    Private Sub increaseProgress(ByVal itemNum As Integer)
        Me.ItemsProgressBar.Increment(itemNum)
    End Sub

    Private Sub ItemCounter(ByVal items As Integer)
        ItemsProgressBar.Maximum = items
    End Sub


    Private Sub TestForm_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        AddHandler miSerialProtocolClass.ItemCount, AddressOf ItemCounter
    End Sub

    Private Sub ItemsProgressBar_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ItemsProgressBar.Click

    End Sub
End Class