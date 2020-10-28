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
    'para não deixar a ligação cair
    Public Lastpointer As Point = MousePosition()
    Public MPx As Point = MousePosition()
    Public DireitaOuEsquerda = 0
    Public TimerEnable = 0
    Public TimerTime = 0
    Public MouseEnable = 0
    Public ScrollEnable = 0
    'para não deixar a ligação cair

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.AllowDrop = True
        Me.Text = "VNC 1.7"
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
        'para não deixar a ligação cair
        TimerEnable = reader.ReadLine
        TimerTime = reader.ReadLine
        MouseEnable = reader.ReadLine
        ScrollEnable = reader.ReadLine

        If TimerEnable = 1 Then
            Timer1.Enabled = True
            Timer1.Interval = TimerTime
        End If
        'para não deixar a ligação cair

        reader.Close()
        reader.Dispose()

        If WinSCPpath <> Nothing Then Button6.Visible = True
        If Notepadpath <> Nothing Then Button7.Visible = True

        ComboBox1.DataSource = ComboNames

        ReadMenu()
        For f = 0 To 255
            DomainUpDown1.Items.Add(f)
        Next
        RadioButton1.Checked = True
        CheckBox3.Checked = True
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
        'para Ping
        'Dim temp2 = Split(Links(ComboBox1.SelectedIndex), ";")
        'If My.Computer.Network.Ping(Trim(temp2(3))) Then
        '    ListBox1.Items.Insert(0, "success - " & Trim(temp2(3)))
        'Else
        '    ListBox1.Items.Insert(0, "no reply - " & Trim(temp2(3)))
        'End If
        'Windows.Forms.Cursor.Position = New Point(x, y)



        ' para não deixar cair a ligação

        If ScrollEnable = 1 Then
            SendKeys.Send("{SCROLLLOCK}")
        End If


        If MouseEnable = 1 Then

            MPx = MousePosition()
            If MPx = Lastpointer Then
                If DireitaOuEsquerda = 0 Then
                    Windows.Forms.Cursor.Position = New Point((Lastpointer.X.ToString + 10), MPx.Y.ToString)
                    DireitaOuEsquerda = 1
                Else
                    Windows.Forms.Cursor.Position = New Point((Lastpointer.X.ToString - 10), MPx.Y.ToString)
                    DireitaOuEsquerda = 0
                End If

            End If
            Lastpointer = MPx
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'ListBox1.Items.Clear()
        'If Timer1.Enabled = False Then
        '    Timer1.Enabled = True
        '    Button2.Text = "Stop"
        'Else
        '    Timer1.Enabled = False
        '    Button2.Text = "Start"
        'End If


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
                    If RadioButton1.Checked = True Then
                        If str.ToLower.Contains(TextBox1.Text.ToLower) Then
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton2.Checked = True Then
                        If str.ToLower.Contains(TextBox1.Text.ToLower) And str.ToLower.Contains("srv") Then
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton3.Checked = True Then
                        If str.ToLower.Contains(TextBox1.Text.ToLower) And str.ToLower.Contains("bombas") Then
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton4.Checked = True Then
                        If str.ToLower.Contains(TextBox1.Text.ToLower) And str.ToLower.Contains("myauchan") Then
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                ElseIf parts.Length = 2 Then                                        'pesquisa com duas partes
                    If RadioButton1.Checked = True Then
                        If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) Then
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton2.Checked = True Then
                        If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) And str.ToLower.Contains("srv") Then
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton3.Checked = True Then
                        If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) And str.ToLower.Contains("bombas") Then
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton4.Checked = True Then
                        If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) And str.ToLower.Contains("myauchan") Then
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                Else                                                              'pesquisa com três partes
                    If RadioButton1.Checked = True Then
                        If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) And str.ToLower.Contains(parts(2).ToLower) Then 'pesquisa com três partes
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton2.Checked = True Then
                        If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) And str.ToLower.Contains(parts(2).ToLower) And str.ToLower.Contains("srv") Then 'pesquisa com três partes
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton3.Checked = True Then
                        If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) And str.ToLower.Contains(parts(2).ToLower) And str.ToLower.Contains("bombas") Then 'pesquisa com três partes
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
                    End If
                    If RadioButton4.Checked = True Then
                        If str.ToLower.Contains(parts(0).ToLower) And str.ToLower.Contains(parts(1).ToLower) And str.ToLower.Contains(parts(2).ToLower) And str.ToLower.Contains("myauchan") Then 'pesquisa com três partes
                            ListBox1.Items.Add(str)
                            ReDim Preserve FoundList(ListBox1.Items.Count - 1)
                            FoundList(ListBox1.Items.Count - 1) = Links(i)
                        End If
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
            MsgBox("problemas ao lançar o putty!" & vbCrLf & vbCrLf & ex.ToString)
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

    Private Sub SiacFranquiasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SiacFranquiasToolStripMenuItem.Click
        Process.Start(Chromepath & " ", Trim(MenuItemsList(7))) '172.24.251.130 e casi igual
    End Sub

    Private Sub CasiFranqiasToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CasiFranqiasToolStripMenuItem.Click
        Process.Start(Chromepath & " ", Trim(MenuItemsList(8)))
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click

        Try
            Process.Start(VNCpath & " ", TextBox2.Text & AutoScaling)
        Catch ex As Exception
            MsgBox("problemas ao lançar o VNC!" & vbCrLf & vbCrLf & ex.ToString)
        End Try

        TextBox1.Select()

    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Dim encontrado = Nothing
        Try
            Dim p As New ProcessStartInfo
            p.FileName = WinSCPpath
            p.Arguments = "sftp://root:pinguim@" & TextBox2.Text
            Process.Start(p)
        Catch ex As Exception
            MsgBox("problemas ao lançar o WinSCP!" & vbCrLf & vbCrLf & ex.ToString)
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
            MsgBox("problemas ao lançar o Notepad++!" & vbCrLf & vbCrLf & ex.ToString)
        End Try
        Padding = Nothing

        TextBox1.Select()
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Try
            Dim tempx = Split(TextBox2.Text, ".")
            DomainUpDown1.SelectedIndex = tempx(3)
        Catch ex As Exception

        End Try

    End Sub

    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            If CheckBox2.Checked = True Then
                Button5_Click(sender, e)
            End If
            If CheckBox3.Checked = True Then
                Button4_Click(sender, e)
            End If
            If CheckBox4.Checked = True Then
                Button6_Click(sender, e)
            End If
        End If
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

    Private Sub Form1_DragDrop(sender As Object, e As DragEventArgs) Handles Me.DragDrop
        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)


        For Each path In files
            Dim v = path.Split("\")
            If v(v.Length - 1) = "DADOSPDV.XML" Then
                Dim result As DialogResult = MessageBox.Show("O ficheiro dadospdv.xml pode ser aberto como tabela, pretente abrir este ficheiro como uma tabela?", "Abrir como tabela?", MessageBoxButtons.YesNoCancel)
                If result = DialogResult.Cancel Then
                    Exit Sub
                ElseIf result = DialogResult.No Then
                    'fazer nada e seguir caminho
                ElseIf result = DialogResult.Yes Then
                    Dadosxmlfile = path
                    Try
                        Dadospdv.Show()
                    Catch ex As Exception
                        Me.Text = "Error"
                    End Try
                    Exit Sub
                End If
            End If

            Try
                System.Diagnostics.Process.Start(Notepadpath, """" & path & """")
            Catch ex As Exception

            End Try
        Next
    End Sub

    Private Sub Form1_DragEnter(sender As Object, e As DragEventArgs) Handles Me.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub DomainUpDown1_SelectedItemChanged(sender As Object, e As EventArgs) Handles DomainUpDown1.SelectedItemChanged
        Try
            Dim tempx = Split(TextBox2.Text, ".")
            TextBox2.Text = tempx(0) & "." & tempx(1) & "." & tempx(2) & "." & DomainUpDown1.Text
            TextBox2.Focus()
            My.Computer.Keyboard.SendKeys("{END}", True)
        Catch ex As Exception

        End Try
    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        TextBox1_TextChanged(sender, e)
        TextBox1.Focus()
        My.Computer.Keyboard.SendKeys("{END}", True)
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        TextBox1_TextChanged(sender, e)
        TextBox1.Focus()
        My.Computer.Keyboard.SendKeys("{END}", True)
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        TextBox1_TextChanged(sender, e)
        TextBox1.Focus()
        My.Computer.Keyboard.SendKeys("{END}", True)
    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        TextBox1_TextChanged(sender, e)
        TextBox1.Focus()
        My.Computer.Keyboard.SendKeys("{END}", True)
    End Sub

    Private Sub PastaDoProgramaToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PastaDoProgramaToolStripMenuItem.Click
        Process.Start("explorer.exe", CurDir())
    End Sub

    Private Sub KtransferToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles KtransferToolStripMenuItem.Click
        Process.Start("explorer.exe", "k:\transfer\Q&A")
    End Sub

    Private Sub QAToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QAToolStripMenuItem.Click
        Process.Start("explorer.exe", "k:\OKI\Q&A (não apagar)")
    End Sub

    Private Sub OutraToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OutraToolStripMenuItem.Click
        Process.Start("explorer.exe", "k:\OKI\Q&A (não apagar)\_Connector\RDP")
    End Sub

    Private Sub OutraToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles OutraToolStripMenuItem1.Click
        Process.Start("explorer.exe", MenuItemsList(9))
    End Sub
End Class
