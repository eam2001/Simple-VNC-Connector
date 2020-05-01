Imports System.Xml
Imports System.IO

Public Class Dadospdv


    Private Sub Dadospdv_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ToolStripComboBox1.Items.Add("Apenas a ultima entrada de cada caixa")
        ToolStripComboBox1.Items.Add("Todas as entrada das caixas")
        ToolStripComboBox1.SelectedIndex = 0



    End Sub



    Private Sub ToolStripComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ToolStripComboBox1.SelectedIndexChanged
        DataGridView1.Rows.Clear()
        Dim Dados(999, 4)

        Dim xmldoc As New XmlDataDocument()
        Dim xmlnode As XmlNodeList
        Dim i As Integer
        Dim fs As New FileStream(Dadosxmlfile, FileMode.Open, FileAccess.Read)
        xmldoc.Load(fs)
        xmlnode = xmldoc.GetElementsByTagName("PDV")

        If ToolStripComboBox1.SelectedIndex = 0 Then
            ' apenas a ultima de cada
            For i = 0 To xmlnode.Count - 1
                Dim numero = xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
                Dados(numero, 0) = xmlnode(i).ChildNodes.Item(0).InnerText.Trim()
                Dados(numero, 1) = xmlnode(i).ChildNodes.Item(1).InnerText.Trim()
                Dados(numero, 2) = xmlnode(i).ChildNodes.Item(2).InnerText.Trim()
                Dados(numero, 3) = xmlnode(i).ChildNodes.Item(3).InnerText.Trim()
                Dados(numero, 4) = xmlnode(i).ChildNodes.Item(4).InnerText.Trim()

            Next
            For o = 0 To 999
                If Dados(o, 0) <> Nothing Then
                    DataGridView1.Rows.Add(Dados(o, 0), Dados(o, 1), Dados(o, 2), Dados(o, 3), Dados(o, 4))
                End If
            Next
        Else
            'todas as entradas
            For i = 0 To xmlnode.Count - 1
                DataGridView1.Rows.Add(xmlnode(i).ChildNodes.Item(0).InnerText.Trim(), xmlnode(i).ChildNodes.Item(1).InnerText.Trim(), xmlnode(i).ChildNodes.Item(2).InnerText.Trim(), xmlnode(i).ChildNodes.Item(3).InnerText.Trim(), xmlnode(i).ChildNodes.Item(4).InnerText.Trim())
            Next
        End If

        Me.Text = Me.Text & " - " & i & " Entradas"

    End Sub
End Class