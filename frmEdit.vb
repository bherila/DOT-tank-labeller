Imports System.Windows.Forms
Imports System.Drawing

Friend Class frmEdit
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents Panel3 As System.Windows.Forms.Panel
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox4 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox5 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox6 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox7 As System.Windows.Forms.TextBox
    Friend WithEvents TextBox8 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(frmEdit))
        Me.Panel3 = New System.Windows.Forms.Panel
        Me.Label4 = New System.Windows.Forms.Label
        Me.TextBox8 = New System.Windows.Forms.TextBox
        Me.TextBox7 = New System.Windows.Forms.TextBox
        Me.TextBox6 = New System.Windows.Forms.TextBox
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.TextBox5 = New System.Windows.Forms.TextBox
        Me.TextBox4 = New System.Windows.Forms.TextBox
        Me.TextBox3 = New System.Windows.Forms.TextBox
        Me.TextBox2 = New System.Windows.Forms.TextBox
        Me.TextBox1 = New System.Windows.Forms.TextBox
        Me.PictureBox2 = New System.Windows.Forms.PictureBox
        Me.Button1 = New System.Windows.Forms.Button
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel3
        '
        Me.Panel3.AutoScroll = True
        Me.Panel3.BackColor = System.Drawing.Color.Gray
        Me.Panel3.BackgroundImage = CType(resources.GetObject("Panel3.BackgroundImage"), System.Drawing.Image)
        Me.Panel3.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.TextBox8)
        Me.Panel3.Controls.Add(Me.TextBox7)
        Me.Panel3.Controls.Add(Me.TextBox6)
        Me.Panel3.Controls.Add(Me.PictureBox1)
        Me.Panel3.Controls.Add(Me.TextBox5)
        Me.Panel3.Controls.Add(Me.TextBox4)
        Me.Panel3.Controls.Add(Me.TextBox3)
        Me.Panel3.Controls.Add(Me.TextBox2)
        Me.Panel3.Controls.Add(Me.TextBox1)
        Me.Panel3.Controls.Add(Me.PictureBox2)
        Me.Panel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel3.Location = New System.Drawing.Point(0, 0)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(888, 685)
        Me.Panel3.TabIndex = 1
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.ForeColor = System.Drawing.Color.Gray
        Me.Label4.Location = New System.Drawing.Point(16, 576)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(368, 16)
        Me.Label4.TabIndex = 8
        Me.Label4.Text = "h"
        '
        'TextBox8
        '
        Me.TextBox8.BackColor = System.Drawing.Color.GreenYellow
        Me.TextBox8.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox8.Font = New System.Drawing.Font("Franklin Gothic Medium", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox8.Location = New System.Drawing.Point(816, 624)
        Me.TextBox8.MaxLength = 1
        Me.TextBox8.Name = "TextBox8"
        Me.TextBox8.Size = New System.Drawing.Size(24, 29)
        Me.TextBox8.TabIndex = 1
        Me.TextBox8.Text = "0"
        Me.TextBox8.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox7
        '
        Me.TextBox7.BackColor = System.Drawing.Color.GreenYellow
        Me.TextBox7.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox7.Font = New System.Drawing.Font("Franklin Gothic Medium", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox7.Location = New System.Drawing.Point(816, 568)
        Me.TextBox7.MaxLength = 1
        Me.TextBox7.Name = "TextBox7"
        Me.TextBox7.Size = New System.Drawing.Size(24, 29)
        Me.TextBox7.TabIndex = 1
        Me.TextBox7.Text = "0"
        Me.TextBox7.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'TextBox6
        '
        Me.TextBox6.BackColor = System.Drawing.Color.GreenYellow
        Me.TextBox6.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox6.Font = New System.Drawing.Font("Franklin Gothic Medium", 14.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox6.Location = New System.Drawing.Point(816, 520)
        Me.TextBox6.MaxLength = 1
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(24, 29)
        Me.TextBox6.TabIndex = 1
        Me.TextBox6.Text = "0"
        Me.TextBox6.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox1.Location = New System.Drawing.Point(432, 16)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(432, 440)
        Me.PictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.PictureBox1.TabIndex = 9
        Me.PictureBox1.TabStop = False
        Me.ToolTip1.SetToolTip(Me.PictureBox1, "Click here to select image")
        '
        'TextBox5
        '
        Me.TextBox5.BackColor = System.Drawing.Color.Honeydew
        Me.TextBox5.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox5.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox5.Location = New System.Drawing.Point(16, 216)
        Me.TextBox5.MaxLength = 3000
        Me.TextBox5.Multiline = True
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox5.Size = New System.Drawing.Size(368, 344)
        Me.TextBox5.TabIndex = 4
        Me.TextBox5.Tag = "300"
        Me.TextBox5.Text = ""
        '
        'TextBox4
        '
        Me.TextBox4.BackColor = System.Drawing.Color.Honeydew
        Me.TextBox4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox4.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox4.Location = New System.Drawing.Point(16, 144)
        Me.TextBox4.MaxLength = 500
        Me.TextBox4.Multiline = True
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox4.Size = New System.Drawing.Size(368, 64)
        Me.TextBox4.TabIndex = 3
        Me.TextBox4.Text = "Contains: "
        '
        'TextBox3
        '
        Me.TextBox3.BackColor = System.Drawing.Color.Honeydew
        Me.TextBox3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox3.Font = New System.Drawing.Font("Arial", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox3.Location = New System.Drawing.Point(16, 94)
        Me.TextBox3.Multiline = True
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.TextBox3.Size = New System.Drawing.Size(368, 40)
        Me.TextBox3.TabIndex = 2
        Me.TextBox3.Text = ""
        '
        'TextBox2
        '
        Me.TextBox2.BackColor = System.Drawing.Color.Honeydew
        Me.TextBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox2.Enabled = False
        Me.TextBox2.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox2.Location = New System.Drawing.Point(16, 57)
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(368, 26)
        Me.TextBox2.TabIndex = 1
        Me.TextBox2.Text = "Part Number"
        '
        'TextBox1
        '
        Me.TextBox1.BackColor = System.Drawing.Color.Honeydew
        Me.TextBox1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.TextBox1.Enabled = False
        Me.TextBox1.Font = New System.Drawing.Font("Arial", 16.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TextBox1.Location = New System.Drawing.Point(16, 16)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(368, 32)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "Name"
        '
        'PictureBox2
        '
        Me.PictureBox2.BackColor = System.Drawing.Color.Transparent
        Me.PictureBox2.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox2.Name = "PictureBox2"
        Me.PictureBox2.Size = New System.Drawing.Size(880, 679)
        Me.PictureBox2.TabIndex = 10
        Me.PictureBox2.TabStop = False
        '
        'Button1
        '
        Me.Button1.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(768, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(88, 23)
        Me.Button1.TabIndex = 12
        Me.Button1.Text = "&Done"
        '
        'ToolTip1
        '
        Me.ToolTip1.AutomaticDelay = 0
        Me.ToolTip1.ShowAlways = True
        '
        'frmEdit
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(888, 685)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Panel3)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximumSize = New System.Drawing.Size(896, 712)
        Me.MinimizeBox = False
        Me.Name = "frmEdit"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Tag = "closeNo"
        Me.Text = "Label Designer"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region

    Public MF As Form1

    Private Sub Label1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub TextBox5_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TextBox5.TextChanged
        'Dim img As New Bitmap(300, 300)
        'Dim g As Graphics = Graphics.FromImage(img)
        Label4.Text = (TextBox5.MaxLength - TextBox5.Text.Length) & " characters remaining."
    End Sub

    Private Sub PictureBox1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox1.Click
        Dim g As New img
        g.Tag = PictureBox1.Tag
        g.ShowDialog()
        PictureBox1.Tag = CStr(g.Tag).Replace(".png", "")
        If (Not g.Tag Is Nothing) AndAlso (g.Tag <> "") Then
            PictureBox1.Image = App.GetEmbeddedResource(g.Tag & ".png")
        Else
            PictureBox1.Image = New Bitmap(100, 100)
        End If
        'Commit()
    End Sub

    Private Sub TextBox1_LostFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles TextBox1.LostFocus, TextBox2.LostFocus, TextBox3.LostFocus, TextBox4.LostFocus, TextBox5.LostFocus, TextBox6.LostFocus, TextBox7.LostFocus, TextBox8.LostFocus
        'commit()
    End Sub

    Function Save() As Boolean

        Try
            Application.DoEvents()
            Dim cmd As New SqlClient.SqlCommand
            Dim sb As New System.Text.StringBuilder
            sb.Append("update [labels] set [dot_shipping_name] = '")
            sb.Append(TextBox3.Text.Replace("'", "''"))
            sb.Append("', [contents] = '")
            sb.Append(TextBox4.Text.Replace("'", "''"))
            sb.Append("', [boiler_plate] = '")
            sb.Append(TextBox5.Text.Replace("'", "''"))
            sb.Append("', [health] = '")
            sb.Append(TextBox6.Text.Replace("'", "''"))
            sb.Append("', [flammabilty] = '")
            sb.Append(TextBox7.Text.Replace("'", "''"))
            sb.Append("', [reactivity] = '")
            sb.Append(TextBox8.Text.Replace("'", "''"))
            sb.Append("', [icon] = '")
            sb.Append(PictureBox1.Tag)
            sb.Append("' where [product_number] = '")
            sb.Append(TextBox2.Text.Replace("'", "''"))
            sb.Append("'")
            cmd.CommandText = sb.ToString
            cmd.Connection = MF.conn
            cmd.ExecuteNonQuery()
            cmd.Dispose()
            Return True
        Catch
            MsgBox("Could not save. Please check your values, then try again.")
            Return False
        End Try
        'With CType(MF.dss.Labels.Select("[product_number] = '" & TextBox2.Text & "'")(0), Data.LabelsRow)
        '    .Product_Name = TextBox1.Text
        '    .DOT_Shipping_Name = TextBox3.Text
        '    .Contents = TextBox4.Text
        '    .Boiler_Plate = TextBox5.Text
        '    .Health = TextBox6.Text
        '    .Flammabilty = TextBox7.Text
        '    .Reactivity = TextBox8.Text
        '    .Icon = PictureBox1.Tag
        'End With
        'MF.da.Update(MF.dss)

    End Function

    'Sub LoadItems()
    '    ListBox1.Items.Clear()
    '    Dim cmd As New SqlClient.SqlCommand("select [product_number] from [labels]", MF.conn)
    '    Dim dr As SqlClient.SqlDataReader
    '    dr = cmd.ExecuteReader
    '    While dr.Read
    '        ListBox1.Items.Add(dr(0))
    '    End While
    '    dr.Close()
    '    'Dim i As Integer
    '    'For i = 0 To MF.dss.Labels.Count - 1
    '    '    ListBox1.Items.Add(MF.dss.Labels.Rows(i)("product_number"))
    '    'Next
    '    ListBox1.SelectedIndex = si
    'End Sub

    Private Sub frmEdit_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs)
        'Add Label

        'Catch'
        'End Try
    End Sub

    Public si As Integer = -1

    'Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    On Error Resume Next
    '    si = ListBox1.SelectedIndex
    '    Dim cmd As New SqlClient.SqlCommand("select * from [labels] where [product_number] = '" & ListBox1.Items(ListBox1.SelectedIndex) & "'", MF.conn)
    '    Dim dr As SqlClient.SqlDataReader = cmd.ExecuteReader()
    '    dr.Read()
    '    TextBox1.Text = ""
    '    TextBox2.Text = ""
    '    TextBox3.Text = ""
    '    TextBox4.Text = ""
    '    TextBox5.Text = ""
    '    TextBox6.Text = "0"
    '    TextBox7.Text = "0"
    '    TextBox8.Text = "0"
    '    TextBox1.Text = dr(0)
    '    TextBox2.Text = dr(1)
    '    TextBox3.Text = dr(2)
    '    TextBox4.Text = dr(3)
    '    TextBox5.Text = dr(4)
    '    TextBox6.Text = dr(5)
    '    TextBox7.Text = dr(6)
    '    TextBox8.Text = dr(7)
    '    PictureBox1.Tag = dr(8)
    '    If (Not PictureBox1.Tag Is Nothing) AndAlso (Not PictureBox1.Tag = "") Then
    '        PictureBox1.Image = App.GetEmbeddedResource(PictureBox1.Tag & ".png")
    '    End If
    '    dr.Close()
    '    cmd.Dispose()
    'End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Close()
    End Sub

    Private Sub frmEdit_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        If Me.Tag = "closeOk" Then
            e.Cancel = False
        Else
            Select Case MsgBox("Do you want to save any changes before closing this label?", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Exclamation, "Save Changes")
                Case MsgBoxResult.Yes
                    If Save() Then
                        Me.Tag = "closeOk"
                        e.Cancel = False
                    Else
                        Me.Tag = "closeNotOk"
                        e.Cancel = True
                    End If
                Case MsgBoxResult.No
                    Me.Tag = "closeOk"
                    e.Cancel = False
            End Select
        End If
    End Sub
End Class
