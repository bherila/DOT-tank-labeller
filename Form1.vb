Imports System.Drawing
Imports System.Windows.Forms
Imports Marisol

Friend Class Form1
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
    Friend WithEvents Button3 As System.Windows.Forms.Button
    Friend WithEvents ppd As System.Windows.Forms.PrintPreviewDialog
    Friend WithEvents ofd As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ps As System.Windows.Forms.PageSetupDialog
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents pd As System.Windows.Forms.PrintDialog
    Friend WithEvents Button5 As System.Windows.Forms.Button
    Friend WithEvents Timer1 As System.Windows.Forms.Timer
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents Button2 As System.Windows.Forms.Button
    Friend WithEvents zplSave As System.Windows.Forms.SaveFileDialog
    Friend WithEvents Splitter1 As System.Windows.Forms.Splitter
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader
    Friend WithEvents clb As System.Windows.Forms.ListView
    Friend WithEvents Button6 As System.Windows.Forms.Button
    Friend WithEvents pp As System.Windows.Forms.PictureBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Dim resources As System.Resources.ResourceManager = New System.Resources.ResourceManager(GetType(Form1))
        Me.Button3 = New System.Windows.Forms.Button
        Me.ppd = New System.Windows.Forms.PrintPreviewDialog
        Me.ofd = New System.Windows.Forms.OpenFileDialog
        Me.ps = New System.Windows.Forms.PageSetupDialog
        Me.Label2 = New System.Windows.Forms.Label
        Me.pd = New System.Windows.Forms.PrintDialog
        Me.Button5 = New System.Windows.Forms.Button
        Me.Timer1 = New System.Windows.Forms.Timer(Me.components)
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Button6 = New System.Windows.Forms.Button
        Me.Button2 = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.zplSave = New System.Windows.Forms.SaveFileDialog
        Me.clb = New System.Windows.Forms.ListView
        Me.ColumnHeader1 = New System.Windows.Forms.ColumnHeader
        Me.ColumnHeader2 = New System.Windows.Forms.ColumnHeader
        Me.Splitter1 = New System.Windows.Forms.Splitter
        Me.pp = New System.Windows.Forms.PictureBox
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Button3
        '
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button3.Enabled = False
        Me.Button3.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button3.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button3.Location = New System.Drawing.Point(112, 16)
        Me.Button3.Name = "Button3"
        Me.Button3.TabIndex = 9
        Me.Button3.Text = "Print"
        '
        'ppd
        '
        Me.ppd.AutoScrollMargin = New System.Drawing.Size(0, 0)
        Me.ppd.AutoScrollMinSize = New System.Drawing.Size(0, 0)
        Me.ppd.ClientSize = New System.Drawing.Size(400, 300)
        Me.ppd.Enabled = True
        Me.ppd.Icon = CType(resources.GetObject("ppd.Icon"), System.Drawing.Icon)
        Me.ppd.Location = New System.Drawing.Point(17, 17)
        Me.ppd.MinimumSize = New System.Drawing.Size(375, 250)
        Me.ppd.Name = "ppd"
        Me.ppd.TransparencyKey = System.Drawing.Color.Empty
        Me.ppd.UseAntiAlias = True
        Me.ppd.Visible = False
        '
        'ofd
        '
        Me.ofd.DefaultExt = "mdb"
        Me.ofd.Filter = "Microsoft Access Databases (*.mdb)|*.mdb"
        Me.ofd.Title = "Select Database"
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 12)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(96, 16)
        Me.Label2.TabIndex = 14
        Me.Label2.Text = "Actions:"
        '
        'Button5
        '
        Me.Button5.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button5.Enabled = False
        Me.Button5.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button5.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button5.Location = New System.Drawing.Point(280, 16)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(80, 23)
        Me.Button5.TabIndex = 17
        Me.Button5.Text = "Add Label"
        '
        'Timer1
        '
        Me.Timer1.Enabled = True
        Me.Timer1.Interval = 500
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Button6)
        Me.Panel1.Controls.Add(Me.Button2)
        Me.Panel1.Controls.Add(Me.Button1)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Button5)
        Me.Panel1.Controls.Add(Me.Button3)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Panel1.Location = New System.Drawing.Point(0, 305)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(675, 56)
        Me.Panel1.TabIndex = 21
        '
        'Button6
        '
        Me.Button6.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button6.Enabled = False
        Me.Button6.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button6.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button6.Location = New System.Drawing.Point(200, 16)
        Me.Button6.Name = "Button6"
        Me.Button6.TabIndex = 23
        Me.Button6.Text = "Page Setup"
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Enabled = False
        Me.Button2.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button2.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button2.Location = New System.Drawing.Point(464, 16)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(80, 23)
        Me.Button2.TabIndex = 21
        Me.Button2.Text = "Delete Label"
        '
        'Button1
        '
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.Enabled = False
        Me.Button1.FlatStyle = System.Windows.Forms.FlatStyle.Popup
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 7.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(368, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(80, 23)
        Me.Button1.TabIndex = 20
        Me.Button1.Text = "Edit Label"
        '
        'zplSave
        '
        Me.zplSave.Filter = "ZPL Files (*.zpl)|*.zpl"
        '
        'clb
        '
        Me.clb.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.clb.Dock = System.Windows.Forms.DockStyle.Left
        Me.clb.FullRowSelect = True
        Me.clb.HideSelection = False
        Me.clb.Location = New System.Drawing.Point(0, 0)
        Me.clb.Name = "clb"
        Me.clb.Size = New System.Drawing.Size(224, 305)
        Me.clb.TabIndex = 25
        Me.clb.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "M Code"
        Me.ColumnHeader1.Width = 88
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Name"
        Me.ColumnHeader2.Width = 132
        '
        'Splitter1
        '
        Me.Splitter1.Location = New System.Drawing.Point(224, 0)
        Me.Splitter1.Name = "Splitter1"
        Me.Splitter1.Size = New System.Drawing.Size(3, 305)
        Me.Splitter1.TabIndex = 26
        Me.Splitter1.TabStop = False
        '
        'pp
        '
        Me.pp.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.pp.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pp.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pp.Location = New System.Drawing.Point(227, 0)
        Me.pp.Name = "pp"
        Me.pp.Size = New System.Drawing.Size(448, 305)
        Me.pp.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage
        Me.pp.TabIndex = 27
        Me.pp.TabStop = False
        '
        'Form1
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 14)
        Me.ClientSize = New System.Drawing.Size(675, 361)
        Me.Controls.Add(Me.pp)
        Me.Controls.Add(Me.Splitter1)
        Me.Controls.Add(Me.clb)
        Me.Controls.Add(Me.Panel1)
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MinimumSize = New System.Drawing.Size(683, 388)
        Me.Name = "Form1"
        Me.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Show
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Print Labels"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region



    Dim CurrentPage As Integer = 1
    Dim TotalPages As Integer = 5
    Dim Font9 As New Font("Arial", 9, FontStyle.Bold, GraphicsUnit.Point, 0)
    Dim Font12 As New Font("Arial", 12, FontStyle.Bold, GraphicsUnit.Point, 0)
    Dim Font14 As New Font("Arial", 14, FontStyle.Bold, GraphicsUnit.Point, 0)
    Dim Font16 As New Font("Arial", 16, FontStyle.Bold, GraphicsUnit.Point, 0)
    Dim Font18 As New Font("Arial", 18, FontStyle.Bold, GraphicsUnit.Point, 0)
    Dim Font20 As New Font("Arial", 20, FontStyle.Bold, GraphicsUnit.Point, 0)
    Dim Font30 As New Font("Arial", 25, FontStyle.Bold, GraphicsUnit.Point, 0)
    Dim Font36 As New Font("Arial", 36, FontStyle.Bold, GraphicsUnit.Point, 0)

    Public MyDal As DAL.DAL_General
    Dim da As SqlClient.SqlDataAdapter
    Dim dss As New DAL.LabelsData
    Public conn As SqlClient.SqlConnection

    Const PhoneNumber As String = "732-469-5100"    ' ALSO SET INSIDE THE DRUMLABEL REPORT rptDrumLabel.cs



    'Const StrMdbLocation = _
    '"V:\Access File for Labels\Marisol_Product_Drum_Label_Database_08132004.mdb"

    'Const strConnectionString = _
    ' "Jet Sql:Global Partial Bulk Ops=2;Jet Sql:Registry Path=;Jet Sql:Database Locking Mode=1;Jet Sql:Database Password=;Data Source=""" & StrMdbLocation & """;Password=;Jet Sql:Engine Type=5;Jet Sql:Global Bulk Transactions=1;Provider=""Microsoft.Jet.SqlClient.4.0"";Jet Sql:System database=;Jet Sql:SFP=False;Extended Properties=;Mode=Share Deny None;Jet Sql:New Database Password=;Jet Sql:Create System Database=False;Jet Sql:Don't Copy Locale on Compact=False;Jet Sql:Compact Without Replica Repair=False;User ID=Admin;Jet Sql:Encrypt Database=False"

    Private cr As DAL.LabelsData.LabelsRow

    Sub PrintPage(ByVal sender As Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs)
        On Error Resume Next
        e.Graphics.PageUnit = GraphicsUnit.Inch

        ' Dim cr As DAL.LabelsData.LabelsRow = CType(dss.Labels.Select("[Product_Number] = '" & clb.SelectedItems(0).Text & "'")(0), DAL.LabelsData.LabelsRow)
        cr = CType(dss.Labels.Select("[Product_Number] = '" & clb.SelectedItems(0).Text & "'")(0), DAL.LabelsData.LabelsRow)
        RenderPage(e.Graphics, False, cr)

        ' e.Graphics.RotateTransform(270)

        CurrentPage += 1
        e.HasMorePages = (CurrentPage <= TotalPages)

        isprev = False
    End Sub

    Sub RenderPreview(ByVal ProductNumber As String)
        Dim bmp As New System.Drawing.Bitmap(1100, 800)
        Dim g As System.Drawing.Graphics = System.Drawing.Graphics.FromImage(bmp)
        g.PageUnit = GraphicsUnit.Inch
        Dim cr As DAL.LabelsData.LabelsRow = CType(dss.Labels.Select("[Product_Number] = '" & ProductNumber & "'")(0), DAL.LabelsData.LabelsRow)
        RenderPage(g, True, cr)
        pp.Image = bmp
    End Sub

    Sub RenderPage(ByRef GraphicsObject As System.Drawing.Graphics, ByRef isPrev As Boolean, ByRef cr As DAL.LabelsData.LabelsRow)
        If isPrev = True Then
            [GraphicsObject].DrawImage(App.GetEmbeddedResource("template.png"), CSng(0.0), CSng(0.0), CSng(11.0), CSng(8.5))
        End If

        Dim tm As New Drawing2D.Matrix


        tm.Translate(CSng(ppd.Document.DefaultPageSettings.Margins.Left / 100), CSng(ppd.Document.DefaultPageSettings.Margins.Top / 100))
        tm.Translate(0.125, 0)
        'tm.Rotate(270)


        [GraphicsObject].Transform = tm

        'Dim ProductNameField As New RectangleF(0, 0.0417, 4.4584, 0.4167)
        'Dim ProductNumberField As New RectangleF(0, 0.4583, 2.9167, 0.25)
        'Dim DOT_Shipping_NameField As New RectangleF(0, 0.7917, 4.5, 0.5417)
        'Dim ContentsField As New RectangleF(0, 1.375, 4.5, 0.875)
        'Dim BoilerPlateField As New RectangleF(0, 2.375, 4.4583, 4.9167)

        'Dim ProductNameField As New RectangleF(0, 0.0417, 5.2, 0.4167)
        'Dim ProductNumberField As New RectangleF(0, 0.4583, 5.2, 0.25)
        'Dim DOT_Shipping_NameField As New RectangleF(0, 0.7917, 5.2, 0.5417)
        'Dim ContentsField As New RectangleF(0, 1.375, 5.2, 0.875)
        'Dim BoilerPlateField As New RectangleF(0, 2.375, 5.1, 4.9167)

        Dim ProductNameField As New RectangleF(0, 0.0417, 6.5, 1.0167)
        Dim ProductNumberField As New RectangleF(0, 1.23, 5.2, 0.25)
        Dim DOT_Shipping_NameField As New RectangleF(0, 1.4917, 5.2, 0.5417)
        Dim ContentsField As New RectangleF(0, 1.8, 5.2, 0.875)
        Dim BoilerPlateField As New RectangleF(0, 2.375, 5.1, 4.9167)

        ' Dim ProductName2Field As New RectangleF(5, 5.4583, 6.75, 2.2083)
        ' Dim ProductNumber2Field As New RectangleF(5, 5.6667, 6.0417, 2.2083)

        Dim ProductName2Field As New RectangleF(5.25, 5.74, 6.75, 2.2083)
        Dim ProductNumber2Field As New RectangleF(5.25, 5.9667, 6.0417, 2.2083)

        Dim HealthField As New RectangleF(10.0042, 6.425, 1.3333, 1.2917)
        Dim FlammabilityField As New RectangleF(10.0042, 7.05, 1.3333, 1.2917)
        Dim ReactivityField As New RectangleF(10.0042, 7.7167, 1.3333, 1.2917)


        Dim PhoneNumberField As New RectangleF(4.455, 8.1375, 1.3125, 0.2188)
        'Dim PhoneNumberField As New RectangleF(4.455, 8.2, 1.3125, 0.2188)
        'Dim ImageField As New RectangleF(4.9, -0.1, 5.5, 5.5)
        Dim ImageField As New RectangleF(5.03 - 0.0625, -0.08, 5.9 + 0.0625, 5.9 + 0.0625)


        Dim CenteredTextStyle As New System.Drawing.StringFormat
        CenteredTextStyle.Alignment = StringAlignment.Center
        CenteredTextStyle.LineAlignment = StringAlignment.Center

        Dim Icon As Bitmap = App.GetEmbeddedResource(cr.Icon.ToLower & ".png")
        If Not Icon Is Nothing Then _
            [GraphicsObject].DrawImage(Icon, ImageField)

        '   [GraphicsObject].DrawString(cr.Product_Name, Font20, Brushes.Black, ProductNameField)
        [GraphicsObject].DrawString(clb.SelectedItems(0).SubItems(1).Text, Font36, Brushes.Black, ProductNameField)
        [GraphicsObject].DrawString(cr.Product_Number, Font14, Brushes.Black, ProductNumberField)
        [GraphicsObject].DrawString(cr.DOT_Shipping_Name, Font9, Brushes.Black, DOT_Shipping_NameField)
        [GraphicsObject].DrawString(cr.Contents, Font9, Brushes.Black, ContentsField)
        [GraphicsObject].DrawString(cr.Boiler_Plate, Font9, Brushes.Black, BoilerPlateField)
        ' [GraphicsObject].DrawString(cr.Product_Name, Font14, Brushes.Black, ProductName2Field)
        ' [GraphicsObject].DrawString(cr.Product_Number, Font14, Brushes.Black, ProductNumber2Field)
        '[GraphicsObject].DrawString(cr.Product_Name, Font12, Brushes.Black, ProductName2Field)
        [GraphicsObject].DrawString(clb.SelectedItems(0).SubItems(1).Text, Font12, Brushes.Black, ProductName2Field)
        [GraphicsObject].DrawString(cr.Product_Number, Font12, Brushes.Black, ProductNumber2Field)

        [GraphicsObject].DrawString(cr.Health, Font30, Brushes.Black, HealthField)
        [GraphicsObject].DrawString(cr.Flammabilty, Font30, Brushes.Black, FlammabilityField)
        [GraphicsObject].DrawString(cr.Reactivity, Font30, Brushes.Black, ReactivityField)
        [GraphicsObject].DrawString(PhoneNumber, Font9, Brushes.Black, PhoneNumberField, CenteredTextStyle)
        '[GraphicsObject].DrawString(cr.Icon, Font16, Brushes.Gray, ImageField, CenteredTextStyle)
    End Sub

    Sub Reset(ByVal Sender As Object, ByVal e As System.Drawing.Printing.PrintEventArgs)
        CurrentPage = 1
        TotalPages = clb.SelectedItems.Count
        pns.Clear()

        Dim i As Integer
        For i = 0 To TotalPages - 1
            pns.Add((clb.SelectedItems.Item(i).Text))
        Next
    End Sub

    'Function XPix(ByVal e As System.Drawing.Graphics, ByVal Inches As Double)
    '    Return Inches * e.PageUnit
    'End Function
    'Function YPix(ByVal e As System.Drawing.Graphics, ByVal Inches As Double)
    '    Return Inches * e.DpiY
    'End Function

    Public CurrentPermissionSet As PermissionSet
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        CurrentPermissionSet = New PermissionSet(True, True, True)

        clb.Items.Clear()
        With CurrentPermissionSet
            '    If .IsValid = False Then
            '        MsgBox("Your user identity has invalid authentication data. Please contact your system administrator.", MsgBoxStyle.Critical)
            '        Application.Exit()
            '        Exit Sub
            '    End If
            Button3.Enabled = .CanPrint
        End With

        Me.Text = "Marisol Tank Label Editor - " & Environment.UserName

        ppd.Document = New System.Drawing.Printing.PrintDocument
        ppd.Document.DefaultPageSettings.Landscape = True
        ppd.Document.DocumentName = "Marisol Labels"
        ppd.Document.DefaultPageSettings.Margins.Left = 0
        ppd.Document.DefaultPageSettings.Margins.Right = 0
        ppd.Document.DefaultPageSettings.Margins.Top = 0
        ppd.Document.DefaultPageSettings.Margins.Bottom = 0
        AddHandler ppd.Document.PrintPage, AddressOf PrintPage
        AddHandler ppd.Document.BeginPrint, AddressOf Reset
        ps.Document = ppd.Document
        pd.Document = ppd.Document
        'pp.Document = ppd.Document

        'CFName = GetSetting("MarisolLabels", "Settings", "LastFile", "")
        'Try
        '    If CFName.Length > 0 Then
        '        If conn.State = ConnectionState.Open Then conn.Close()
        '        '    conn.ConnectionString = strConnectionString
        '        'conn.ConnectionString = "Jet Sql:Global Partial Bulk Ops=2;Jet Sql:Registry Path=;Jet Sql:Database Locking Mode=1;Jet Sql:Database Password=;Data Source=""" & CFName & """;Password=;Jet Sql:Engine Type=5;Jet Sql:Global Bulk Transactions=1;Provider=""Microsoft.Jet.SqlClient.4.0"";Jet Sql:System database=;Jet Sql:SFP=False;Extended Properties=;Mode=Share Deny None;Jet Sql:New Database Password=;Jet Sql:Create System Database=False;Jet Sql:Don't Copy Locale on Compact=False;Jet Sql:Compact Without Replica Repair=False;User ID=Admin;Jet Sql:Encrypt Database=False"
        '        conn.ConnectionString = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
        '        conn.Open()
        '    End If
        'Catch ex As Exception
        '    'Dim el As New EventLog
        '    'el.WriteEntry(ex.Message, EventLogEntryType.Error)
        '    'el.Dispose()
        '    MsgBox(ex.Message & Environment.NewLine & Environment.NewLine & ex.StackTrace)
        'End Try

        conn = MyDal.GetConnection()
        da = MyDal.Labels_CreateDataAdapter(conn)

        LoadItems()
    End Sub

    Sub LoadItems()
        Dim ff As New frmLoading
        Me.AddOwnedForm(ff)
        ff.Show()
        Application.DoEvents()
        clb.Items.Clear()
        MyDal.Labels_GetLabelList(conn, clb, CInt(currentsort))
        dss.Clear()
        MyDal.Labels_FillTypedDataSet(da, dss)
        Button3.Enabled = CurrentPermissionSet.CanPrint
        Button6.Enabled = Button3.Enabled
        Button5.Enabled = CurrentPermissionSet.CanEditLabels
        Button1.Enabled = Button5.Enabled
        Button2.Enabled = Button5.Enabled
        ff.Close()
    End Sub

    Private Sub conn_StateChange(ByVal sender As System.Object, ByVal e As System.Data.StateChangeEventArgs)
        If e.CurrentState = ConnectionState.Open Then
            LoadItems()
        Else
            'Button2.Enabled = False
            Button3.Enabled = False
            Button5.Enabled = False
        End If
        Me.Focus()
    End Sub

    Dim pns As New ArrayList

    ' Print
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If clb.SelectedItems.Count > 0 Then

            ' Ben's code for printing....:

            If pd.ShowDialog() = Windows.Forms.DialogResult.OK Then
                pd.Document.Print()
                ss("PrinterName", pd.Document.PrinterSettings.PrinterName)
                'ss("", pd.Document.PrinterSettings.FromPage
            End If

            ' End of Ben's code


            '' John Woll's code for calling Active Report:
            'Me.Cursor = Cursors.WaitCursor
            'cr = CType(dss.Labels.Select("[Product_Number] = '" & clb.SelectedItems(0).Text & "'")(0), DAL.LabelsData.LabelsRow)
            'Dim rpt As Reports.rptDrumLabel = New Reports.rptDrumLabel(cr, _
            '        clb.SelectedItems(0).SubItems(1).Text)
            'Reports.PrintPreview.Display(rpt)
            'Me.Cursor = Cursors.Default
            ''End of John Woll's code

        Else
            MsgBox("There is no label selected. Please pick a label, then try again.", MsgBoxStyle.Critical, "No Label Selected")
        End If
    End Sub

    Sub ss(ByVal setting As String, ByVal value As String)

    End Sub

    ' ADD LABEL
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Try
            If CurrentPermissionSet.CanEditLabels Then
                Dim pn As New frmNewLabel
                Dim ll As New frmLoading
                ll.Label1.Text = "One moment, please..."
                ll.Show()
                Application.DoEvents()
                Dim cn As New SqlClient.SqlCommand
                cn.Connection = conn
                Dim dr As SqlClient.SqlDataReader
                Dim descs As New ArrayList
                cn.CommandText = "select code,product,description from MAR.dbo.LU_MCode order by code asc"
                dr = cn.ExecuteReader
                While dr.Read
                    Dim it As New ListViewItem(CStr(dr(0)))
                    it.SubItems.Add(dr(1).ToString())  ' John Woll added the ToString() to avoid a runtime error "Requires narrowing conversion"
                    it.SubItems.Add(dr(2).ToString())  ' John Woll added the ToString() to avoid a runtime error "Requires narrowing conversion"
                    pn.lv.Items.Add(it)
                    descs.Add(dr(2))
                End While
                dr.Close()
                pn.Tag = descs
                ll.Close()
                If pn.ShowDialog = Windows.Forms.DialogResult.OK Then
                    Dim where As String = "where product_number = '" & pn.lv.SelectedItems(0).Text.Replace("'", "''") & "'"
                    cn.CommandText = "select count(*) from labels " & where & " and deleted = 1"
                    Dim hres As Integer = CInt(cn.ExecuteScalar)
                    Dim finished As Boolean = False
                    If hres > 0 Then
                        Dim fc As New frmDataConflict
                        fc.ShowDialog()
                        If fc.RadioButton1.Checked Then
                            'Restore
                            cn.CommandText = "update labels set deleted = 0 " & where
                            cn.ExecuteNonQuery()
                            finished = True
                        ElseIf fc.RadioButton2.Checked Then
                            'Overwrite
                            cn.CommandText = "delete from [labels] " & where
                            cn.ExecuteNonQuery()
                            hres = 0
                        Else
                            cn.Dispose()
                            Button5_Click(sender, e)
                        End If
                    End If
                    If hres = 0 And finished = False Then
                        Try
                            cn.CommandText = "insert into [labels] ([product_name],[product_number], [dot_shipping_name], [contents], [boiler_plate], [health], [flammabilty], [reactivity], [icon], [deleted]) values ('" & pn.lv.SelectedItems(0).SubItems(1).Text.Replace("'", "''") & "', '" & pn.lv.SelectedItems(0).SubItems(0).Text.Replace("'", "''") & "', '" & pn.lv.SelectedItems(0).SubItems(2).Text.Replace("'", "''") & "', '-', '-', '0', '0', '0', '',0)"
                            cn.ExecuteNonQuery()
                            finished = True

                        Catch sqlex As SqlClient.SqlException
                            MsgBox("The Mcode selected already has a label.", MsgBoxStyle.Exclamation)
                        Catch ex As Exception
                            MsgBox("The label cannot be created because you entered invalid data or it already exists.", MsgBoxStyle.Critical)
                        End Try
                    End If
                    If finished Then
                        Edit(pn.lv.SelectedItems(0).Text)
                        LoadItems()
                    End If
                End If
                cn.Dispose()
            Else
                MsgBox("You do not have permission to do this.", MsgBoxStyle.Critical, "Access Denied")
            End If
        Catch ex As Exception
            MsgBox(ex.Message.ToString(), MsgBoxStyle.Critical, "Access Denied")
        End Try
    End Sub

    Dim isprev As Boolean = False

    Private Sub Timer1_Tick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Timer1.Tick
        isprev = True
        'pp.InvalidatePreview()
        'clb.SelectedItems(0).SubItems(0).Text
        If clb.SelectedItems.Count > 0 Then
            RenderPreview(clb.SelectedItems(0).SubItems(0).Text)
        End If
        Timer1.Stop()
        Timer1.Enabled = False
    End Sub

    Dim CFName As String = ""

    'Private Sub MenuItem2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItem2.Click
    '    ' MessageBox.Show("The application now uses an Access database, located at " & StrMdbLocation)
    '    'Me.Focus()
    '    ' Exit Sub

    '    If ofd.ShowDialog = DialogResult.OK Then
    '        CFName = ofd.FileName

    '        Try
    '            SaveSetting("MarisolLabels", "Settings", "LastFile", CFName)
    '            If conn.State = ConnectionState.Open Then conn.Close()
    '            'conn.ConnectionString = strConnectionString
    '            'conn.ConnectionString = "Jet Sql:Global Partial Bulk Ops=2;Jet Sql:Registry Path=;Jet Sql:Database Locking Mode=1;Jet Sql:Database Password=;Data Source=""" & CFName & """;Password=;Jet Sql:Engine Type=5;Jet Sql:Global Bulk Transactions=1;Provider=""Microsoft.Jet.SqlClient.4.0"";Jet Sql:System database=;Jet Sql:SFP=False;Extended Properties=;Mode=Share Deny None;Jet Sql:New Database Password=;Jet Sql:Create System Database=False;Jet Sql:Don't Copy Locale on Compact=False;Jet Sql:Compact Without Replica Repair=False;User ID=Admin;Jet Sql:Encrypt Database=False"
    '            'da.UpdateCommand.Connection = conn
    '            'da.InsertCommand.Connection = conn
    '            'da.DeleteCommand.Connection = conn
    '            conn.Open()
    '        Catch ex As Exception
    '            Dim el As New EventLog
    '            el.WriteEntry(ex.Message, EventLogEntryType.Error)
    '            el.Dispose()
    '            MsgBox(ex.Message & Environment.NewLine & Environment.NewLine & ex.StackTrace)
    '        End Try
    '    End If
    '    Me.Focus()
    'End Sub

    Private Sub MenuItem3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Application.Exit()
    End Sub

    ' DELETE LABEL
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Try
            If CurrentPermissionSet.CanDeleteLabels = True Then
                If MsgBox("Are you sure you want to permanently remove this item?", MsgBoxStyle.Exclamation + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Confirm Item Delete") = MsgBoxResult.Yes Then
                    '   Dim cmd As New SqlClient.SqlCommand("update [labels] set Deleted = 1 where [product_number] = '" & clb.Items(clb.SelectedItems(0).Index).Text & "'", conn)
                    Dim cmd As New SqlClient.SqlCommand("delete [labels] where [product_number] = '" & clb.Items(clb.SelectedItems(0).Index).Text & "'", conn)
                    cmd.ExecuteNonQuery()
                    cmd.Dispose()
                    LoadItems()
                End If
            Else
                MsgBox("You do not have permission to delete labels.", MsgBoxStyle.Critical, "Permission Denied")
            End If
        Catch ex As ArgumentOutOfRangeException
            MsgBox("Select an item to delete.", MsgBoxStyle.Critical, "No Label Selected")
        Catch ex As Exception

        End Try
    End Sub

    ' EDIT LABEL
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If clb.SelectedItems.Count > 0 Then
            Try
                Dim iii As Integer = clb.SelectedItems(0).Index
                Edit(clb.Items(clb.SelectedItems(0).Index).Text)
                clb.Items(iii).Selected = True
                LoadItems()
            Catch ex As ArgumentOutOfRangeException
                MsgBox("Select a label to edit.", MsgBoxStyle.Critical)
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical)
            End Try
        Else
            MsgBox("There is no label selected. In order to edit a label, you must first pick one from the list on the left.", MsgBoxStyle.Critical)
        End If
    End Sub

    Sub Edit(ByVal id As String)
        On Error Resume Next
        Dim cx As New frmEdit
        With cx
            .si = clb.SelectedItems(0).Index
            Dim cmd As New SqlClient.SqlCommand("select * from [mes].dbo.[labels] where [product_number] = '" & id & "'", conn)
            Dim dr As SqlClient.SqlDataReader = cmd.ExecuteReader()
            dr.Read()
            .TextBox1.Text = ""
            .TextBox2.Text = ""
            .TextBox3.Text = ""
            .TextBox4.Text = ""
            .TextBox5.Text = ""
            .TextBox6.Text = "0"
            .TextBox7.Text = "0"
            .TextBox8.Text = "0"
            .TextBox1.Text = dr(0)
            .TextBox2.Text = dr(1)
            .TextBox3.Text = dr(2)
            .TextBox4.Text = dr(3)
            .TextBox5.Text = dr(4)
            .TextBox6.Text = dr(5)
            .TextBox7.Text = dr(6)
            .TextBox8.Text = dr(7)
            .PictureBox1.Tag = dr(8)
            If (Not .PictureBox1.Tag Is Nothing) AndAlso (Not .PictureBox1.Tag = "") Then
                .PictureBox1.Image = App.GetEmbeddedResource(.PictureBox1.Tag & ".png")
            End If
            dr.Close()
            cmd.Dispose()
            .MF = Me
            .ShowDialog()
            LoadItems()
        End With
    End Sub

    'Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
    '    'FUNCTIONALITY DISABLED
    '    'Dim s As New System.Text.StringBuilder
    '    's.Append("^XA")
    '    's.Append(Environment.NewLine)
    '    'Dim cr As Data.LabelsRow = CType(dss.Labels.Select("[Product_Number] = '" & clb.Items(clb.SelectedIndex) & "'")(0), Data.LabelsRow)
    '    's.Append("^LH30,30")
    '    's.Append(Environment.NewLine)
    '    's.Append("^FO20,10")
    '    's.Append(Environment.NewLine)
    '    's.Append("^AD")
    '    's.Append(Environment.NewLine)
    '    's.Append("^FD")
    '    's.Append(cr.Product_Name)
    '    's.Append("^FS")
    '    's.Append(Environment.NewLine)
    '    's.Append("^FO20,60")
    '    's.Append("^FD")
    '    's.Append(cr.Product_Number)
    '    's.Append("^FS")
    '    's.Append(Environment.NewLine)
    '    's.Append("^XZ")
    '    ''If zplSave.ShowDialog = DialogResult.OK Then
    '    'Dim str As New IO.FileStream("\\mardcfs1\product label printer\lbl.zpl", IO.FileMode.Create)
    '    'Dim sw As New IO.StreamWriter(str)
    '    'sw.Write(s.ToString)
    '    'sw.Close()
    '    'str.Close()
    '    'MsgBox("ZPL File copied to printer!", MsgBoxStyle.Information)
    '    ''End If
    'End Sub

    Private Sub ListView1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles clb.SelectedIndexChanged
        If CurrentPermissionSet.CanEditLabels Then
            Button1.Enabled = (clb.SelectedIndices.Count > 0) And (CurrentPermissionSet.CanEditLabels)
        Else
            Button1.Enabled = False
        End If
        If CurrentPermissionSet.CanPrint Then
            Button3.Enabled = (clb.SelectedIndices.Count > 0) And (CurrentPermissionSet.CanPrint)
        Else
            Button3.Enabled = False
        End If
        If CurrentPermissionSet.CanDeleteLabels Then
            Button2.Enabled = (clb.SelectedIndices.Count > 0) And (CurrentPermissionSet.CanDeleteLabels)
        Else
            Button2.Enabled = False
        End If
        Button6.Enabled = Button3.Enabled
        Button1.Enabled = Button5.Enabled
        Button2.Enabled = Button5.Enabled
        Timer1.Stop()
        Timer1.Enabled = True
        Timer1.Start()
    End Sub

    Public currentsort As Sorts = Sorts.MCodeID
    Enum Sorts
        MCodeID = 0
        Name = 1
    End Enum
    Private Sub clb_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles clb.ColumnClick
        currentsort = e.Column
        LoadItems()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        ps.ShowDialog()
    End Sub

    Private Sub Form1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles MyBase.Closing
        conn.Close()
    End Sub

End Class
