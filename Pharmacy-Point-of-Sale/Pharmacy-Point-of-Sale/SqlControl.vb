Imports System.Data.SqlClient

Public Class SqlControl

  'Create ADO.net object
  'Protected Connection As New SqlConnection With {.ConnectionString = "Server=WORKSYSTEM;Database=cbis_icphDB;User ID=" & Environment.UserName & ";Password=;Trusted_Connection=True;"}

  'Protected Connection As New SqlConnection With {.ConnectionString = "Server=KIZURASHI\SQLEXPRESS;Database=cbis_icphDB;USER=Kizurashi;Password=12345"}

  'Data Source=sage2\MSSQLSERVER;Database=MATERIEL;Integrated Security=False;User Instance=False;user='" + TxtUser.Text + "';password='" + TxtPassword.Text + "'"

  'Protected Connection As New SqlConnection With {.ConnectionString = "Server=SERVER\SQLEXPRESS;Database=cbis_icphDB;USER=admin;Password=ICPH2017"}

  Protected Connection As New SqlConnection With {.ConnectionString = "Server=RANIEL-GARCIA\SQLEXPRESS;Database=cbis_icphDB;Trusted_Connection=True;TrustServerCertificate=True"}

  'Protected Connection As New SqlConnection With {.ConnectionString = "Data Source=192.168.1.57,1433;Network Library=DBMSSOCN;Initial Catalog=cbis_icphDB;User ID=" & Environment.UserName & ";Password=SWORDfish2596;"}
  Protected security As New Security

  ' For Update, Delete, Insert query only
  Protected Function ExecuteCommand(ByVal commandQuery As String) As Boolean

    Dim sqlCommand As New SqlCommand(commandQuery, Me.Connection)

    Try
      If Me.Connection.State = False Then
        Me.Connection.Open()
      End If

      sqlCommand.ExecuteReader()
      Return True
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error: 0x101 Execute command", MessageBoxButtons.OK, MessageBoxIcon.Warning)
    Finally
      sqlCommand.Dispose()
      Me.Connection.Close()
    End Try

    Return False
  End Function

End Class
