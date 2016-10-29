Imports Microsoft.VisualBasic.CompilerServices
Imports System.IO
Imports System.Text
Imports System.Runtime.CompilerServices


'Coded By X-SLAYER updated by slogutis 
'Feel Free To Contact Me :
'https://web.facebook.com/ihebbriki1
'https://twitter.com/IHEB_X_SLAYER

'Feel Free To Support Me With a Donation :
'https://pledgie.com/campaigns/32450
'BTC:
'1PqKXtkpfejYe2tdsMTL1KopxRZECtT7qx


Public Class Form1

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.ListBox1.Items.AddRange(Clipboard.GetText.ToString.Split(New String() {ChrW(13) & ChrW(10)}, StringSplitOptions.RemoveEmptyEntries))
        Me.Label1.Text = (Conversions.ToString(Me.ListBox1.Items.Count))
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Timer1.Start()
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Try
            ListBox1.SelectedIndex += 1

            If ListBox1.SelectedItem.ToString.Contains("gmail") Then 'gmail
                Lgmail.Items.Add(ListBox1.SelectedItem.ToString)


            ElseIf ListBox1.SelectedItem.ToString.Contains("hotmail") Then 'Hotmail

                Lhotmail.Items.Add(ListBox1.SelectedItem.ToString)
            ElseIf ListBox1.SelectedItem.ToString.Contains("yahoo") Then 'Yahoo

                Lyahoo.Items.Add(ListBox1.SelectedItem.ToString)
            Else
                lother.Items.Add(ListBox1.SelectedItem.ToString)
            End If

        Catch ex As Exception
            Timer1.Stop()
        End Try
        gg.Text = Lgmail.Items.Count
        y.Text = Lyahoo.Items.Count
        other.Text = lother.Items.Count
        lm.Text = Lhotmail.Items.Count
        Label9.Text = ListBox1.SelectedIndex
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If (Me.Lgmail.Items.Count = 0) Then
            Interaction.MsgBox("There is Nothing to Copy.", MsgBoxStyle.ApplicationModal, Nothing)
        Else
            Dim enumerator As IEnumerator
            Dim builder As New StringBuilder
            Try
                enumerator = Me.Lgmail.Items.GetEnumerator
                Do While enumerator.MoveNext
                    builder.AppendLine(RuntimeHelpers.GetObjectValue(enumerator.Current).ToString)
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try

            Dim pp As New SaveFileDialog
            pp.Filter = "|*.txt"
            pp.FileName = "gmail"
            If pp.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim k As New StreamWriter(pp.FileName)
                k.Write(builder.ToString)
                k.Close()
            End If

            Interaction.MsgBox((Conversions.ToString(Me.Lgmail.Items.Count) & " Items Saved."), MsgBoxStyle.ApplicationModal, Nothing)
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If (Me.Lyahoo.Items.Count = 0) Then
            Interaction.MsgBox("There is nothing to copy.", MsgBoxStyle.ApplicationModal, Nothing)
        Else
            Dim enumerator As IEnumerator
            Dim builder As New StringBuilder
            Try
                enumerator = Me.Lyahoo.Items.GetEnumerator
                Do While enumerator.MoveNext
                    builder.AppendLine(RuntimeHelpers.GetObjectValue(enumerator.Current).ToString)
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try
            Dim pp As New SaveFileDialog
            pp.Filter = "|*.txt"
            pp.FileName = "yahoo"
            If pp.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim k As New StreamWriter(pp.FileName)
                k.Write(builder.ToString)
                k.Close()
            End If
            Interaction.MsgBox((Conversions.ToString(Me.Lyahoo.Items.Count) & " Items Saved."), MsgBoxStyle.ApplicationModal, Nothing)
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If (Me.Lhotmail.Items.Count = 0) Then
            Interaction.MsgBox("There is nothing to copy.", MsgBoxStyle.ApplicationModal, Nothing)
        Else
            Dim enumerator As IEnumerator
            Dim builder As New StringBuilder
            Try
                enumerator = Me.Lhotmail.Items.GetEnumerator
                Do While enumerator.MoveNext
                    builder.AppendLine(RuntimeHelpers.GetObjectValue(enumerator.Current).ToString)
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try

            Dim pp As New SaveFileDialog
            pp.Filter = "|*.txt"
            pp.FileName = "hotmail"
            If pp.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim k As New StreamWriter(pp.FileName)
                k.Write(builder.ToString)
                k.Close()
            End If

            Interaction.MsgBox((Conversions.ToString(Me.Lhotmail.Items.Count) & " Items Saved."), MsgBoxStyle.ApplicationModal, Nothing)
        End If

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        If (Me.lother.Items.Count = 0) Then
            Interaction.MsgBox("nothing to copy.", MsgBoxStyle.Exclamation, Nothing)
        Else
            Dim enumerator As IEnumerator
            Dim builder As New StringBuilder
            Try
                enumerator = Me.lother.Items.GetEnumerator
                Do While enumerator.MoveNext
                    builder.AppendLine(RuntimeHelpers.GetObjectValue(enumerator.Current).ToString)
                Loop
            Finally
                If TypeOf enumerator Is IDisposable Then
                    TryCast(enumerator, IDisposable).Dispose()
                End If
            End Try

            Dim pp As New SaveFileDialog
            pp.Filter = "|*.txt"
            pp.FileName = "Other"
            If pp.ShowDialog = Windows.Forms.DialogResult.OK Then
                Dim k As New StreamWriter(pp.FileName)
                k.Write(builder.ToString)
                k.Close()
            End If

            Interaction.MsgBox((Conversions.ToString(Me.lother.Items.Count) & " Items Saved."), MsgBoxStyle.ApplicationModal, Nothing)
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Timer1.Stop()
        ListBox1.Items.Clear()
        Lgmail.Items.Clear()
        Lyahoo.Items.Clear()
        Lhotmail.Items.Clear()
        lother.Items.Clear()
        gg.Text = "0"
        y.Text = "0"
        lm.Text = "0"
        Label9.Text = "0"
        Label1.Text = "0"
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        MsgBox("Coded By X-SLAYER", MsgBoxStyle.Information, "OK")
    End Sub
End Class
