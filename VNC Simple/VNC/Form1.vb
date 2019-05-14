Imports System.IO
Public Class Form1

    Public Links
    Public ComboNames
    Public VNCpath
    Public AutoScaling = ""

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(CurDir() & "\VNC.CSV")
        Dim linha = reader.ReadLine
        linha = reader.ReadLine
        Dim n = 0
        Do While linha <> "Configuration"
            ReDim Preserve Links(n)
            ReDim Preserve ComboNames(n)
            Links(n) = linha
            Dim temp = Split(linha, ";")
            ComboNames(n) = temp(1) & " - " & temp(2) & " - " & temp(3)
            n = n + 1
            linha = reader.ReadLine
        Loop
       
        VNCpath = reader.ReadLine

        reader.Close()


        ComboBox1.DataSource = ComboNames
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Try
            Dim temp1 = Split(Links(ComboBox1.SelectedIndex), ";")
            Process.Start(VNCpath & " ", Trim(temp1(3)) & AutoScaling & " -password " & temp1(5))
        Catch ex As Exception
            MsgBox("problemas ao lançar o VNC")
        End Try
    End Sub

    Private Sub CheckBox1_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked = True Then
            AutoScaling = " -autoscaling"
        Else
            AutoScaling = ""
        End If
    End Sub

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Dim temp2 = Split(Links(ComboBox1.SelectedIndex), ";")


        If My.Computer.Network.Ping(Trim(temp2(3))) Then
            ListBox1.Items.Insert(0, "success - " & Trim(temp2(3)))
        Else
            ListBox1.Items.Insert(0, "no reply - " & Trim(temp2(3)))
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        ListBox1.Items.Clear()
        If Timer1.Enabled = False Then
            Timer1.Enabled = True
            Button2.Text = "Stop"
        Else
            Timer1.Enabled = False
            Button2.Text = "Start"
        End If
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        ListBox1.Items.Clear()
        Dim temp3 = Split(Links(ComboBox1.SelectedIndex), ";")
        ListBox1.Items.Add(Trim(temp3(0)))
        ListBox1.Items.Add(Trim(temp3(1)))
        ListBox1.Items.Add(Trim(temp3(2)))
        ListBox1.Items.Add(Trim(temp3(3)))
        ListBox1.Items.Add(Trim(temp3(4)))
        ListBox1.Items.Add(Trim(temp3(5)))

    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Shell("notepad.exe " & CurDir() & "\VNC.CSV", AppWinStyle.MaximizedFocus)
    End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1_Click(sender, e)
        End If
    End Sub
End Class
