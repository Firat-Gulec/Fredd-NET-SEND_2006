
Imports System.DirectoryServices
Imports System.Net

Public Class NetSend

    Private v_opac As String
    Private v_x, v_y, x, y, count, count1 As Integer
    Dim t As Boolean
    
    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        Select Case v_opac
            Case "Close"
                If Me.Opacity > 0 Then
                    Me.Opacity -= 0.04
                Else
                    Me.Close()
                End If
            Case "Minimize"
                If Me.Opacity > 0 Then
                    Me.Opacity -= 0.05
                Else
                    Me.WindowState = FormWindowState.Minimized
                    v_opac = "Normal"
                End If
            Case "Normal"
                If Me.WindowState <> FormWindowState.Minimized Then
                    If Me.Opacity < 1 Then
                        Me.Opacity += 0.05
                    Else
                        If t = True Then
                            Liste()
                        End If
                        Timer1.Enabled = False
                        t = False
                        End If
                End If
        End Select
    End Sub

    Private Sub pctClose_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pctClose.MouseMove
        pctClose.BackgroundImage = My.Resources.CloseOver
    End Sub

    Private Sub pctClose_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles pctClose.MouseLeave
        On Error Resume Next
        pctClose.BackgroundImage = My.Resources.Close
    End Sub

    Private Sub pctClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles pctClose.Click
        v_opac = "Close"
        Timer1.Enabled = True
    End Sub

    Private Sub Mouse_Leave(ByVal sender As Object, ByVal e As System.EventArgs)
        sender.Font = New Font("Calibri", 10, FontStyle.Bold + FontStyle.Italic, GraphicsUnit.Point)
    End Sub

    Private Sub Mouse_Move(ByVal Sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs)
        Sender.Cursor = Cursors.Hand
        Sender.Font = New Font("Calibri", 11, FontStyle.Bold + FontStyle.Underline + FontStyle.Italic, GraphicsUnit.Point)
    End Sub

    Private Sub FormMoveMouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblTitle.MouseDown
        v_x = e.X
        v_y = e.Y
    End Sub

    Private Sub FormMove_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lblTitle.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Me.Left += e.X - v_x
            Me.Top += e.Y - v_y
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        t = True
        Me.Label2.Text = My.Computer.Name
        Me.v_opac = "Normal"
        Timer1.Enabled = True
        SplitContainer3.Panel1Collapsed = True
        ContextMenuStrip1.AllowTransparency = True
        ContextMenuStrip1.Opacity = 0.9

    End Sub

    Private Sub pctMinimize_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles pctMinimize.MouseLeave
        Me.pctMinimize.BackgroundImage = My.Resources.Minimize
    End Sub

    Private Sub pctMinimize_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pctMinimize.MouseMove
        Me.pctMinimize.BackgroundImage = My.Resources.MinimizeOver
    End Sub

    Private Sub MenuStrip1_MenuActivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.MenuActivate
        SplitContainer3.Panel1Collapsed = False
        SplitContainer3.SplitterDistance = 125

    End Sub

    Private Sub MenuStrip1_MenuDeactivate(ByVal sender As Object, ByVal e As System.EventArgs) Handles MenuStrip1.MenuDeactivate
        SplitContainer3.Panel1Collapsed = True

    End Sub

    Private Sub FormMove_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles _
        pctTitle02.MouseDown, pctTitle03.MouseDown, PictureBox6.MouseDown, lblTitle.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Left Then
            If sender.name = "PictureBox6" Then
                ContextMenuStrip1.Show(Me.Left + 1, Me.Top + 72)
            End If
            x = e.X
            y = e.Y
        ElseIf e.Button = Windows.Forms.MouseButtons.Right Then
            Select Case sender.name
                Case "pctTitle02"
                    ContextMenuStrip1.Show(Me.Left + +91 + e.X, Me.Top + e.Y)
                Case "lblTitle"
                    ContextMenuStrip1.Show(Me.Left + 155 + e.X, Me.Top + e.Y)
                Case "pctTitle03"
                    ContextMenuStrip1.Show(Me.Left + 155 + Me.lblTitle.Width + e.X, Me.Top + e.Y)
            End Select
        End If

    End Sub


    Private Sub pctMinimize_MouseClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles pctMinimize.MouseClick
        If e.Button = Windows.Forms.MouseButtons.Left Then
            v_opac = "Minimize"
            Timer1.Enabled = True
        End If
    End Sub

    Private Sub TreeView1_AfterCheck(ByVal sender As Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles TreeView1.AfterCheck
        say()

    End Sub

    Private Sub say()
        count = 0
        For a As Integer = 0 To Me.TreeView1.Nodes.Count - 1
            For b As Integer = 0 To Me.TreeView1.Nodes(a).Nodes.Count - 1
                If Me.TreeView1.Nodes(a).Nodes(b).Checked = True Then
                    count += 1
                    Me.Label1.Text = "Kiþi Sayýsý : " & count
                End If
                If count = 0 Then
                    Me.Label1.Text = "Kiþi Sayýsý : " & 0
                End If
            Next
        Next
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, ToolStripButton2.Click, LjblToolStripMenuItem.Click
        Me.ProgressBar1.Visible = True
        For a As Integer = 0 To Me.TreeView1.Nodes.Count - 1
            For b As Integer = 0 To Me.TreeView1.Nodes(a).Nodes.Count - 1
                If Me.TreeView1.Nodes(a).Nodes(b).Checked = True Then
                    If txtMessage.Text <> "" Then
                        If Me.CheckBox1.Checked = True Then
                            Shell("CMD /C NET SEND /DOMAIN:" & Me.TreeView1.Nodes(a).Nodes(b).Text.Remove(Me.TreeView1.Nodes(a).Nodes(b).Text.IndexOf("(") - 1) & " " & txtMessage.Text)
                        Else
                            Shell("CMD /C NET SEND " & Me.TreeView1.Nodes(a).Nodes(b).Text.Remove(Me.TreeView1.Nodes(a).Nodes(b).Text.IndexOf("(") - 1) & " " & txtMessage.Text)
                        End If

                        count1 += 1
                    End If
                    Me.ProgressBar1.Value += 100 / count
                End If
            Next
        Next
        If Me.TextBox2.Text.Length <> 0 Then
            If Me.txtMessage.Text <> "" Then
                If Me.CheckBox1.Checked = True Then
                    Shell("CMD /C NET SEND /DOMAIN:" & Me.TextBox2.Text & " " & txtMessage.Text)
                Else
                    Shell("CMD /C NET SEND " & Me.TextBox2.Text & " " & txtMessage.Text)
                End If

                count1 += 1
            End If
        End If
        If count1 = 0 Then
            MessageBox.Show("Mesaj Gönderilemedi!" & vbCrLf & vbCrLf & "Mesaj Göndermek için :" & vbCrLf & vbCrLf & "1. En az bir bilgisayar seçmelisiniz" & vbCrLf & vbCrLf & "   veya bilgisayar adý, IP adresi tanýmlamalýsýnýz." & vbCrLf & vbCrLf & "2. Mesaj kutusunu boþ býrakmamalýsýnýz", "Fredd NET SEND", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Else
            MessageBox.Show("Mesajýnýz " & count1 & " bilgisayara baþarýyla gönderildi.", "Fredd NET SEND", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
        count1 = 0
        Me.ProgressBar1.Value = 0
        Me.ProgressBar1.Visible = False
    End Sub

    Function GetIPAddress(ByVal CompName As String) As String
        Dim oAddr As System.Net.IPAddress
        Dim sAddr As String
        Try
            With System.Net.Dns.GetHostByName(CompName)
                oAddr = New System.Net.IPAddress(.AddressList(0).Address)
                sAddr = oAddr.ToString
            End With
            GetIPAddress = sAddr
        Catch Excep As Exception
            MessageBox.Show(Excep.Message, "Fredd NET SEND", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Function

    Private Sub Liste()
        Me.TreeView1.Nodes.Clear()

        Dim childEntry As DirectoryEntry
        Dim ParentEntry As New DirectoryEntry()
        Try
            ParentEntry.Path = "WinNT:"
            For Each childEntry In ParentEntry.Children
                Dim newNode As New TreeNode(childEntry.Name)
                Select Case childEntry.SchemaClassName
                    Case "Domain"
                        Dim ParentDomain As New TreeNode(childEntry.Name & "    (Çalýþma Grubu)")
                        Me.TreeView1.Nodes.AddRange(New TreeNode() {ParentDomain})
                        Dim SubChildEntry As DirectoryEntry
                        Dim SubParentEntry As New DirectoryEntry()
                        SubParentEntry.Path = "WinNT://" & childEntry.Name
                        For Each SubChildEntry In SubParentEntry.Children
                            Dim newNode1 As New TreeNode(SubChildEntry.Name)
                            Select Case SubChildEntry.SchemaClassName
                                Case "Computer"
                                    ParentDomain.Nodes.Add(newNode1.Text & " (" & GetIPAddress(newNode1.Text) & ")")
                            End Select
                        Next
                End Select
            Next
        Catch Excep As Exception
            MessageBox.Show(Excep.Message, "Fredd NET SEND", MessageBoxButtons.OK)
        Finally
            ParentEntry = Nothing
            Me.TreeView1.Sort()
        End Try
        Me.TreeView1.ExpandAll()
        Label6.Visible = False
    End Sub

    Private Sub Button2_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click, KhbToolStripMenuItem.Click, ToolStripButton1.Click
        Me.Label6.Visible = True
        Liste()
    End Sub

    Private Sub ToolStripButton5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton5.Click
        Me.txtMessage.Clear()
    End Sub

    Private Sub ToolStripButton3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton3.Click
        For a As Integer = 0 To Me.TreeView1.Nodes.Count - 1
            For b As Integer = 0 To Me.TreeView1.Nodes(a).Nodes.Count - 1
                Me.TreeView1.Nodes(a).Nodes(b).Checked = True
            Next
        Next

    End Sub

    Private Sub ToolStripButton4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton4.Click
        For a As Integer = 0 To Me.TreeView1.Nodes.Count - 1
            For b As Integer = 0 To Me.TreeView1.Nodes(a).Nodes.Count - 1
                Me.TreeView1.Nodes(a).Nodes(b).Checked = False
            Next
        Next

    End Sub

    Private Sub CloseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CloseToolStripMenuItem.Click
        Me.v_opac = "Close"
        Me.Timer1.Enabled = True
    End Sub

    Private Sub MinimizeToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MinimizeToolStripMenuItem.Click
        Me.v_opac = "Minimize"
        Me.Timer1.Enabled = True
    End Sub
End Class
