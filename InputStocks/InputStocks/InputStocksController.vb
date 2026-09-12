Imports System.Data.SqlClient

Public Class InputStocksController

  Protected Connection As New SqlConnection With {.ConnectionString = "Server=RANIEL-GARCIA\SQLEXPRESS;Database=cbis_icphDB;Trusted_Connection=True;TrustServerCertificate=True"}

  'Protected Connection As New SqlConnection With {.ConnectionString = "Server=WORKSYSTEM;Database=cbis_icphDB;User ID=" & Environment.UserName & ";Password=;Trusted_Connection=True;"}

  Delegate Sub DisplayAllSuppliers(ByVal id As Integer, ByVal name As String)
  Delegate Sub DisplayAllMedicines(ByVal id As Integer, ByVal brandName As String, ByVal genName As String)
  Delegate Sub DisplayAllBatch(ByRef id As Integer, ByVal batch As String)

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

  Public Function CheckMedCategoryExistence(ByVal medCategory As String) As Boolean

    Dim getMedCategoryQuery As String = "SELECT id FROM MedicineCategory WHERE category='" & medCategory & "'"
    Dim sqlCommand As New SqlCommand(getMedCategoryQuery, Me.Connection)

    If Me.Connection.State = False Then
      Me.Connection.Open()
    End If

    Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

    Try
      If (reader.Read()) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message)
    Finally
      reader.Close()
      sqlCommand.Dispose()
      Me.Connection.Close()
    End Try

    Return False
  End Function

  Public Function InsertNewMedCategory(ByVal catName As String) As Boolean
    Try
      Dim insertNewMedCatQuery As String = "INSERT INTO MedicineCategory(category) VALUES('" & catName & "')"

      If Me.ExecuteCommand(insertNewMedCatQuery) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Insert new Category", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Return False
  End Function

  Public Function CheckMedClassificationExistence(ByVal medClassification As String) As Boolean

    Dim getMedClassQuery As String = "SELECT id FROM MedicineClassification WHERE classification= '" & medClassification & "'"
    Dim sqlCommand As New SqlCommand(getMedClassQuery, Me.Connection)

    If Me.Connection.State = False Then
      Me.Connection.Open()
    End If

    Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

    Try
      If (reader.Read()) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Check Med Classification Existence", MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      reader.Close()
      sqlCommand.Dispose()
      Me.Connection.Close()
    End Try

    Return False
  End Function

  Public Function InsertNewMedClassification(ByVal className As String) As Boolean
    Try
      Dim insertNewMedClassQuery As String = "INSERT INTO MedicineClassification(classification) VALUES('" & className & "')"

      If Me.ExecuteCommand(insertNewMedClassQuery) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Insert new medecine classification", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Return False
  End Function

  Public Function GetMedCategoryId(ByVal catName As String) As Integer

    Dim getMedCatIdQuery As String = "SELECT id FROM MedicineCategory WHERE category='" & catName & "'"
    Dim sqlCommand As New SqlCommand(getMedCatIdQuery, Me.Connection)

    If Me.Connection.State = False Then
      Me.Connection.Open()
    End If

    Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

    Try
      If (reader.Read()) Then
        Return reader("id")
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Get medecine category id", MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      reader.Close()
      sqlCommand.Dispose()
      Me.Connection.Close()
    End Try

    Return -1
  End Function

  Public Function GetMedClassificationId(ByVal className As String) As Integer

    Dim getMedCatIdQuery As String = "SELECT id FROM MedicineClassification WHERE classification='" & className & "'"
    Dim sqlCommand As New SqlCommand(getMedCatIdQuery, Me.Connection)

    If Me.Connection.State = False Then
      Me.Connection.Open()
    End If

    Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

    Try
      If (reader.Read()) Then
        Return reader("id")
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Get medecine classification id", MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      reader.Close()
      sqlCommand.Dispose()
      Me.Connection.Close()
    End Try

    Return -1
  End Function

  Public Function CheckMedicineExistence(ByVal brandName As String, ByVal genName As String) As Boolean

    Dim getMedQuery As String = "SELECT * FROM Medicines WHERE brandName='" & brandName & "' AND genericName='" & genName & "'"
    Dim sqlCommand As New SqlCommand(getMedQuery, Me.Connection)

    If Me.Connection.State = False Then
      Me.Connection.Open()
    End If

    Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

    Try
      If (reader.Read()) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Check Medicine Existence", MessageBoxButtons.OK, MessageBoxIcon.Error)
    Finally
      reader.Close()
      sqlCommand.Dispose()
      Me.Connection.Close()
    End Try

    Return False
  End Function

  Public Function InsertNewMed(ByVal brandName As String, ByVal genName As String, ByVal unitCost As Decimal, ByVal catId As Integer, ByVal classId As String) As Boolean
    Try
      Dim insertNewMedQuery As String = "INSERT INTO Medicines(brandName, genericName, unitCost, categoryId, medClassificationId) VALUES('" & brandName & "','" & genName & "', " & unitCost & ", " & catId & ", " & classId & ")"

      If Me.ExecuteCommand(insertNewMedQuery) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Insert new medecine classification", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Return False
  End Function

  Public Function InsertNewSupplier(ByVal compName As String, ByVal name As String, ByVal contactNo As String, ByVal email As String, ByVal address As String) As Boolean
    Try
      Dim insertNewSupplierQuery As String = "INSERT INTO Suppliers(companyName, name, phoneNum, email,address) VALUES('" & compName & "','" & name & "','" & contactNo & "','" & email & "','" & address & "')"

      If Me.ExecuteCommand(insertNewSupplierQuery) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Insert new medecine classification", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Return False
  End Function

  Public Sub GetAllSuppliers(ByRef displayAllSuppliers As DisplayAllSuppliers, ByVal supName As String)
    Try
      Dim getAllSuppliersQuery As String = "SELECT * FROM Suppliers WHERE name LIKE '%" & supName & "%'"

      Dim sqlCommand As New SqlCommand(getAllSuppliersQuery, Me.Connection)

      If Me.Connection.State = False Then
        Me.Connection.Open()
      End If

      Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

      Try
        While (reader.Read())
          displayAllSuppliers(reader("id"), reader("name"))
        End While
      Catch ex As Exception
        MessageBox.Show(ex.Message, "Get All Suppliers (2)", MessageBoxButtons.OK, MessageBoxIcon.Error)
      Finally
        ' Clossing connections
        reader.Close()
        sqlCommand.Dispose()
        Me.Connection.Close()
      End Try

    Catch ex As Exception
      MessageBox.Show(ex.Message, "Get All Suppliers (1)", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
  End Sub

  Public Sub GetAllMedicines(ByRef displayAllMedicines As DisplayAllMedicines, ByVal brandOrGenName As String)
    Try
      Dim getAllMedicinesQuery As String = "SELECT id, brandName, genericName FROM Medicines WHERE brandName LIKE '%" & brandOrGenName & "%' OR genericName LIKE '%" & brandOrGenName & "%'"

      Dim sqlCommand As New SqlCommand(getAllMedicinesQuery, Me.Connection)

      If Me.Connection.State = False Then
        Me.Connection.Open()
      End If

      Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

      Try
        While (reader.Read())
          displayAllMedicines(reader("id"), reader("brandName"), reader("genericName"))
        End While
      Catch ex As Exception
        MessageBox.Show(ex.Message, "Get All Medicines (2)", MessageBoxButtons.OK, MessageBoxIcon.Error)
      Finally
        ' Clossing connections
        reader.Close()
        sqlCommand.Dispose()
        Me.Connection.Close()
      End Try

    Catch ex As Exception
      MessageBox.Show(ex.Message, "Get All Medicines (1)", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
  End Sub

  Public Function GetTheLastBatchID() As Integer

    Dim id As Integer = 0

    Try
      Dim getTheLastBatchIdQuery As String = "SELECT TOP 1 id FROM MedicinesStockBatch ORDER BY id DESC"

      Dim sqlCommand As New SqlCommand(getTheLastBatchIdQuery, Me.Connection)

      If Me.Connection.State = False Then
        Me.Connection.Open()
      End If

      Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

      Try
        If reader.Read() Then
          id = reader("id")
          Return id
        End If
      Catch ex As Exception
        MessageBox.Show(ex.Message, "Get the last batch id (2)", MessageBoxButtons.OK, MessageBoxIcon.Error)
      Finally
        ' Clossing connections
        reader.Close()
        sqlCommand.Dispose()
        Me.Connection.Close()
      End Try

    Catch ex As Exception
      MessageBox.Show(ex.Message, "Get the last batch id (1)", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Return id
  End Function

  Public Function GetBatch(ByRef displayAllBatch As DisplayAllBatch, ByVal lotNum As String) As Integer

    Dim id As Integer = 0

    Try
      Dim getAllBatchQuery As String = "SELECT id, lotNumber FROM MedicinesStockBatch WHERE lotNumber LIKE '%" & lotNum & "%'"

      Dim sqlCommand As New SqlCommand(getAllBatchQuery, Me.Connection)

      If Me.Connection.State = False Then
        Me.Connection.Open()
      End If

      Dim reader As SqlDataReader = sqlCommand.ExecuteReader()

      Try
        While (reader.Read())
          displayAllBatch(reader("id"), reader("lotNumber"))
        End While
      Catch ex As Exception
        MessageBox.Show(ex.Message, "Get all batch (2)", MessageBoxButtons.OK, MessageBoxIcon.Error)
      Finally
        ' Clossing connections
        reader.Close()
        sqlCommand.Dispose()
        Me.Connection.Close()
      End Try

    Catch ex As Exception
      MessageBox.Show(ex.Message, "Get all batch (1)", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Return id
  End Function

  Public Function InsertNewBatch(ByVal batchNo As String, ByVal suppId As Integer) As Boolean
    Try
      Dim insertNewBatchQuery As String = "INSERT INTO MedicinesStockBatch(lotNumber, supplierId) VALUES('" & batchNo & "', " & suppId & ")"

      If Me.ExecuteCommand(insertNewBatchQuery) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Insert New batch", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Return False
  End Function

  Public Function InsertBatchItems(ByVal medId As Integer, ByVal quantity As Integer, ByVal currStocks As Integer, ByVal expDate As Date, ByVal batchId As Integer) As Boolean
    Try
      Dim insertBatchItemQuery As String = "INSERT INTO MedicinesStockBatchItems(medicineId, quantity, currentStocks, expirationDate, batchId)" &
                                          "VALUES(" & medId & ", " & quantity & ", " & currStocks & ", '" & expDate & "', " & batchId & ")"

      If Me.ExecuteCommand(insertBatchItemQuery) Then
        Return True
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Insert New batch", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Return False
  End Function
End Class
