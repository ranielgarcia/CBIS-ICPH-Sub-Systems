Imports System.Data.SqlClient

Public Class SqlControl

  'Protected Connection As New SqlConnection With {.ConnectionString = "Server=WORKSYSTEM;Database=cbis_icphDB;User ID=" & Environment.UserName & ";Password=;Trusted_Connection=True;"}
  ' Here
  Protected Connection As New SqlConnection With {.ConnectionString = "Server=RANIEL-GARCIA\SQLEXPRESS;Database=cbis_icphDB;Trusted_Connection=True;TrustServerCertificate=True"}

  'Protected Connection As New SqlConnection With {.ConnectionString = "Data Source=192.168.11.14,1434;Database=cbis_icphDB;USER=admin;Password=ICPH2017"}

  'Protected Connection As New SqlConnection With {.ConnectionString = "Server=WORKSYSTEM;Database=cbis_icphDB;User ID=" & Environment.UserName & ";Password=;Trusted_Connection=True;"}
  ' For Update, Delete, Insert query only

  'Dim conStr As String = Me.sqlConnString
  'Protected Connection As New SqlConnection With {.ConnectionString = conStr}


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

