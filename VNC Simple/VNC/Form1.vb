Imports System.IO

Public Class Form1

    Public Links
    Public ComboNames
    Public MenuItemsList
    Public FoundList
    Public VNCpath
    Public Puttypath
    Public Chromepath
    Public WinSCPpath
    Public Notepadpath
    Public AutoScaling = ""

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.Text = "VNC 1.5"
        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(CurDir() & "\VNC.CSV")
        Dim linha = reader.ReadLine
        linha = reader.ReadLine
        Dim n = 0
        Do While linha <> "Configuration"
            ReDim Preserve Links(n)
            ReDim Preserve ComboNames(n)
            Links(n) = linha
            Dim temp = Split(linha, ";")
            ComboNames(n) = temp(1) & " - " & temp(2) & " - " & temp(3) & " - " & temp(4)
            n = n + 1
            linha = reader.ReadLine
        Loop
       
        VNCpath = reader.ReadLine
        Puttypath = reader.ReadLine
        Chromepath = reader.ReadLine
        WinSCPpath = reader.ReadLine
        Notepadpath = reader.ReadLine

        reader.Close()
        reader.Dispose()

        If WinSCPpath <> Nothing Then Button6.Visible = True
        If Notepadpath <> Nothing Then Button7.Visible = True

        ComboBox1.DataSource = ComboNames

        ReadMenu()
    End Sub

    Sub ReadMenu()
        Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(CurDir() & "\MENU.CSV")
        Dim linha = reader.ReadLine
        linha = reader.ReadLine
        Dim n = 0
        Do While linha <> "End"
            ReDim Preserve MenuItemsList(n)
            MenuItemsList(n) = linha

            n = n + 1
            linha = reader.ReadLine
        Loop


        reader.Close()
        reader.Dispose()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim encontrado = Nothing
        'MsgBox(FoundList(0))
        Try

            'Dim temp3 = Split(ListBox1.SelectedItem, "-")
            ''MsgBox(Trim(temp3(0)))
            ''MsgBox(Trim(temp3(2)))

            'For Each str As String In Links
            '    If str.ToLower.Contains(Trim(temp3(0)).ToLower) Then
            '        encontrado = str
            '    End If
            'Next

            Dim temp1 = Split(FoundList(ListBox1.SelectedIndex), ";")

            'MsgBox(Trim(temp1(3)))
            'MsgBox(Trim(temp1(5)))

            Process.Start(VNCpath & " ", Trim(temp1(3)) & AutoScaling & " -password " & temp1(5))
        Catch ex As Exception
            MsgBox("problemas ao lançar o VNC! - " & encontrado)
        End Try

        TextBox1.Select()

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
        'ListBox1.Items.Clear()
        'Dim temp3 = Split(Links(ComboBox1.SelectedIndex), ";")
        'ListBox1.Items.Add(Trim(temp3(0)))
        'ListBox1.Items.Add(Trim(temp3(1)))
        'ListBox1.Items.Add(Trim(temp3(2)))
        'ListBox1.Items.Add(Trim(temp3(3)))
        'ListBox1.Items.Add(Trim(temp3(4)))
        'ListBox1.Items.Add(Trim(temp3(5)))

    End Sub

    Private Sub ComboBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles ComboBox1.KeyPress

    End Sub

    'Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
    '    Shell("notepad.exe " & CurDir() & "\VNC.CSV", AppWinStyle.MaximizedFocus)
    '    Shell("notepad.exe " & CurDir() & "\MENU.CSV", AppWinStyle.MaximizedFocus)
    '    TextBox1.Select()

    'End Sub

    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button1_Click(sender, e)
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Trim(TextBox1.Text) = "" Then
            ListBox1.Items.Clear()
            ReDim FoundList(-1)
        Else
            ListBox1.Items.Clear()

            Dim parts(-1)
            parts = Split(Trim(TextBox1.Text), " ")


            Dim i = 0
            For Each str As String In ComboNames
                If parts.Length = 1 Then                                            'pesquisa com uma parte
                    If str.ToLower.Contains(TextBox1.Text.ToLower) Then
                        ListBox1.Items.Add(str)
                        ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                        FoundList(ListBox1.Items.Count - 1) = Links(i)
                    End If
                ElseIf parts.Length = 2 Then                                        'pesquisa com duas partes
                    If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) Then
                        ListBox1.Items.Add(str)
                        ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                        FoundList(ListBox1.Items.Count - 1) = Links(i)
                    End If
                Else                                                                'pesquisa com três partes
                    If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) And str.ToLower.Contains(parts(2).ToLower) Then 'pesquisa com três partes
                        ListBox1.Items.Add(str)
                        ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                        FoundList(ListBox1.Items.Count - 1) = Links(i)
                    End If
                End If
                i = i + 1
            Next
        End If
        If ListBox1.Items.Count <> 0 Then ListBox1.SelectedIndex = 0



    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Try
            Dim temp3 = Split(ListBox1.SelectedItem, "-")
            TextBox2.Text = Trim(temp3(2))
        Catch ex As Exception

        End Try

        TextBox1.Select()
    End Sub

    Private Sub ListBox1_DoubleClick(sender As Object, e As EventArgs) Handles ListBox1.DoubleClick
        Button1_Click(sender, e)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

        Try
            Process.Start(Puttypath, " root@" & TextBox2.Text & " -pw pinguim")
        Catch ex As Exception
            MsgBox("problemas ao lançar o putty!")
        End Try

        TextBox1.Select()
    End Sub

    Private Sub CasiBancadaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CasiBancadaToolStripMenuItem.Click
        Process.Start(Chromepath & " ", Trim(MenuItemsList(0)))
    End Sub

    Private Sub UserParametroToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserParametroToolStripMenuItem.Click
        Try
            My.Computer.Clipboard.SetText("ger#siac@2010")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PassParametroToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PassParametroToolStripMenuItem.Click
        Try
            My.Computer.Clipboard.SetText("$param@ger2010")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub UserPromoçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserPromoçãoToolStripMenuItem.Click
        Try
            My.Computer.Clipboard.SetText("ADMBANCTEC")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PassPromoçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PassPromoçãoToolStripMenuItem.Click
        Try
            My.Computer.Clipboard.SetText("MANAGER PROMO")
        Catch ex As Exception
        End Try
    End Sub

    Private Sub PassConsiacToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PassConsiacToolStripMenuItem.Click
        My.Computer.Clipboard.SetText("sup" & Mid(Now, 4, 2) & "_" & Mid(Now, 12, 2))
    End Sub

    Private Sub SiacBancadaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SiacBancadaToolStripMenuItem.Click
        Process.Start(Chromepath & " ", Trim(MenuItemsList(1)))
    End Sub

    Private Sub CasiProduçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CasiProduçãoToolStripMenuItem.Click
        Process.Start(Chromepath & " ", Trim(MenuItemsList(2)))
    End Sub

    Private Sub SiacProduçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SiacProduçãoToolStripMenuItem.Click
        Process.Start(Chromepath & " ", Trim(MenuItemsList(3)))
    End Sub

    Private Sub CasiPréproduçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CasiPréproduçãoToolStripMenuItem.Click
        Process.Start(Chromepath & " ", Trim(MenuItemsList(4)))
    End Sub

    Private Sub SiacPréproduçãoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SiacPréproduçãoToolStripMenuItem.Click
        Process.Start(Chromepath & " ", Trim(MenuItemsList(5)))
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Dim encontrado = Nothing
        Try
            Process.Start(VNCpath & " ", TextBox2.Text & AutoScaling)
        Catch ex As Exception
            MsgBox("problemas ao lançar o VNC! - " & encontrado)
        End Try

        TextBox1.Select()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            Dim p As New ProcessStartInfo
            p.FileName = WinSCPpath
            p.Arguments = "sftp://root:pinguim@" & TextBox2.Text
            Process.Start(p)
        Catch ex As Exception

        End Try
        Padding = Nothing

        TextBox1.Select()

    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            Dim p As New ProcessStartInfo
            p.FileName = Notepadpath
            Process.Start(p)
        Catch ex As Exception

        End Try
        Padding = Nothing

        TextBox1.Select()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            Button5_Click(sender, e)
        End If
    End Sub

    Private Sub AbrirPastaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AbrirPastaToolStripMenuItem.Click
        Process.Start("explorer.exe", CurDir())
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click

        OpenFileDialog1.FileName = "Dadospdv.xml"
        OpenFileDialog1.DefaultExt = ".xml"
        OpenFileDialog1.InitialDirectory = "K:\"
        Dim result As DialogResult = OpenFileDialog1.ShowDialog()
        If result = Windows.Forms.DialogResult.OK Then
            Dadosxmlfile = OpenFileDialog1.FileName
            Try
                Dadospdv.Show()
            Catch ex As Exception
                Me.Text = "Error"
            End Try
        End If
    End Sub


End Class
