Imports System.Windows.Forms

Public Module Crypto
    Public Class PermissionSet
        Private _CanPrint As Boolean
        Private _CanEditLabels As Boolean
        Private _CanDeleteLabels As Boolean
        Private _IsValid As Boolean
        Public ReadOnly Property CanPrint() As Boolean
            Get
                Return _CanPrint
            End Get
        End Property
        Public ReadOnly Property CanEditLabels() As Boolean
            Get
                Return _CanEditLabels
            End Get
        End Property
        Public ReadOnly Property CanDeleteLabels() As Boolean
            Get
                Return _CanDeleteLabels
            End Get
        End Property
        Public ReadOnly Property IsValid() As Boolean
            Get
                Return _IsValid
            End Get
        End Property
        'Sub New(ByVal PermissionString As String)
        '    Dim sp() As String = PermissionString.Split("A"c)
        '    Dim chk As String = Hex(Environment.UserName.GetHashCode()).Replace("A", "X")
        '    If sp(0).ToUpper = chk.ToUpper Then
        '        IsValid = True
        '        Dim r As String = sp(2)
        '        Me.CanDeleteLabels = r.IndexOf("MDZ") >= 0
        '        Me.CanEditLabels = r.IndexOf("PQB") >= 0
        '        Me.CanManageUsers = r.IndexOf("YCE") >= 0
        '        Me.CanPrint = r.IndexOf("SGT") >= 0
        '    Else
        '        Me.CanPrint = False
        '        Me.CanDeleteLabels = False
        '        Me.CanEditLabels = False
        '        Me.CanManageUsers = False
        '    End If
        'End Sub
        '

        Sub New(ByVal Access As Boolean, ByVal Edit As Boolean, ByVal Delete As Boolean)
            Me._CanPrint = Access
            Me._CanEditLabels = Edit
            Me._CanDeleteLabels = Delete
            Me._IsValid = True
        End Sub

        'Overrides Function ToString() As String
        '    Dim uid As String = Hex(Environment.UserName.GetHashCode()).Replace("A", "X")
        '    uid &= "A"
        '    uid &= RandomString()
        '    uid &= "A"
        '    If Me.CanDeleteLabels = True Then uid &= "MDZ"
        '    If Me.CanEditLabels = True Then uid &= "PQB"
        '    If Me.CanPrint = True Then uid &= "SGT"
        '    uid &= "A"
        '    uid &= RandomString()
        '    Return uid.ToUpper
        'End Function
    End Class


    Friend Function RandomString() As String
        Randomize()
        Dim len As Integer = CInt(Rnd() * 65 + 10)
        Dim i As Integer
        RandomString = ""
        For i = 1 To len
            Randomize()
            RandomString &= UCase(Chr(CInt(Rnd() * 26 + 54)))
        Next
        RandomString = RandomString.Replace("A", "X")
    End Function

    Friend Function Hash(ByVal Input As String) As String
        Return Hex(Input.GetHashCode * 3)
    End Function

    'Friend Sub LogonUser()
    '    Try
    '        Dim u As New Users
    '        u.ReadXml(New IO.FileInfo(Application.ExecutablePath).Directory.FullName.TrimEnd("\") & "\config.xml")
    '        Dim auth As String = ""
    '        If u.users.Select("[name] = '" & (Environment.UserName.Replace("'", "''")) & "'").Length < 1 Then
    '            MsgBox("You do not have permission to run this application in the current context.", MsgBoxStyle.Critical, "Security Warning")
    '            Application.Exit()

    '        Else
    '            Dim r As Users.usersRow = u.users.Select("[name] = '" & (Environment.UserName.Replace("'", "''")) & "'")(0)
    '            CurrentPermissionSet = New PermissionSet(r.data)
    '        End If
    '        u.Dispose()
    '    Catch
    '        MsgBox("User authentication data is not available. Please obtain the authentication key and try again.", MsgBoxStyle.Critical, "Security Warning")
    '        Application.Exit()
    '    End Try
    'End Sub

    Friend Function cs(ByVal username As String) As String
        Dim chk As String = Hex(username.GetHashCode()).Replace("A", "X")
        Return chk.ToUpper & "A" & RandomString() & "ASGTA" & RandomString()
    End Function


End Module
