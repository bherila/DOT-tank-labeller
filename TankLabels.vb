Public Class TankLabels
    Dim f As Form1

    Public Sub New(ByVal MDIParent As System.Windows.Forms.Form, ByVal CanAccess As Boolean, ByVal CanEdit As Boolean, _
        ByVal CanDelete As Boolean, ByVal dalObject As DAL.DAL_General)
        f = New Form1
        f.CurrentPermissionSet = New PermissionSet(CanAccess, CanEdit, CanDelete)
        f.MyDal = dalObject
        '  Me.MDIParent = MDIParent
    End Sub

    Public ReadOnly Property CurrentPermissionSet() As PermissionSet
        Get
            Return f.CurrentPermissionSet
        End Get
    End Property

    Public Property MDIParent() As System.Windows.Forms.Form
        Get
            Return f.MdiParent
        End Get
        Set(ByVal Value As System.Windows.Forms.Form)
            f.MdiParent = Value
        End Set
    End Property

    Public Sub Show()
        f.Show()
    End Sub

    Public Sub ShowDialog()
        f.ShowDialog()
    End Sub

End Class
