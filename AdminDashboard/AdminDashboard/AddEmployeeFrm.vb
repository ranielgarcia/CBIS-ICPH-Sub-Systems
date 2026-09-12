Public Class AddEmployeeFrm

  Public EmployeeId As Integer = -1

  Private FirstName As String
  Private MiddleName As String
  Private LastName As String
  Private Gender As String
  Private Address As String
  Private BirthDate As String
  Private EmployeeTypeId As Integer
  Private DepartmentId As Integer

  Private EmpMangement As New EmployeesManagement

  Private Sub AddEmployeeFrm_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
    ' Display all employee types
    Me.EmpMangement.GetAllEmployeeTypes(Me.EmpTypeCBox)
    ' Display all departments
    Me.EmpMangement.GetAllDepartment(Me.EmpDepartmentCBox)
  End Sub

  Private Function SetEmployeeInfo() As Boolean


    Try
      ' Employee birth date ex: 12/12/2016
      Dim bdateMonth As String = EmpBrithDateDPkr.Value.Month.ToString
      Dim bdateDay As String = EmpBrithDateDPkr.Value.Day.ToString
      Dim bdateYear As String = EmpBrithDateDPkr.Value.Year.ToString
      Dim bdate As String = bdateDay & "/" & bdateMonth & "/" & bdateYear

      Dim format() = {"dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy"}
      Dim birthDate As Date
      Date.TryParseExact(bdate, format,
        System.Globalization.DateTimeFormatInfo.InvariantInfo,
          Globalization.DateTimeStyles.None, birthDate)


      ' Get the employee Gender
      Dim genderIndex As Integer = EmpGenderCbox.SelectedIndex
      Dim genderValue As String = ""
      If genderIndex > -1 Then
        genderValue = EmpGenderCbox.Items(genderIndex)
      End If

      Dim empDeptIndex As Integer = EmpDepartmentCBox.SelectedIndex
      Dim empDeptValue As String = EmpDepartmentCBox.Items(empDeptIndex)

      Dim empTypeIndex As Integer = EmpTypeCBox.SelectedIndex
      Dim empTypeValue As String = EmpTypeCBox.Items(empTypeIndex)

      '' Get the employee type selected item index
      'Dim empTypeIndex As Integer = EmpTypeCBox.SelectedIndex
      'Dim empTypeValue As String = ""
      'Dim empTypeId As String = ""
      'If empTypeIndex > -1 Then
      '    empTypeValue = EmpTypeCBox.Items(empTypeIndex)
      '    Dim empTypeSplitVal() As String = empTypeValue.Split("-")
      '    empTypeId = Integer.Parse(empTypeSplitVal(0))
      'End If

      '' Get the Department selected item index
      'Dim empDeptIndex As Integer = EmpDepartmentCBox.SelectedIndex
      'Dim empDeptValue As String = ""
      'Dim empDeptId As String = ""
      'If empDeptIndex > -1 Then
      '    empDeptValue = EmpDepartmentCBox.Items(empDeptIndex)
      '    Dim empDeptSplitVal() As String = empDeptValue.Split("-")
      '    empDeptId = Integer.Parse(empDeptSplitVal(0))
      'End If


      If Me.CheckValue(EmpFirstNameTxt.Text, Me.fnameError) And
          Me.CheckValue(EmpMiddleNameTxt.Text, Me.mnameError) And
          Me.CheckValue(EmpLastNameTxt.Text, Me.lnameError) And
          Me.CheckValue(genderValue, Me.genderError) And
          Me.CheckValue(EmpAddressTxt.Text, Me.addressError) And
          Me.CheckValue(empTypeValue, Me.empTypeError) And
          Me.CheckValue(empDeptValue, Me.depError) Then

        'Me.fnameError.Visible = True

        If birthDate <> Today Then
          Me.FirstName = EmpFirstNameTxt.Text
          Me.MiddleName = EmpMiddleNameTxt.Text
          Me.LastName = EmpLastNameTxt.Text
          Me.Gender = genderValue
          Me.Address = EmpAddressTxt.Text
          Me.BirthDate = bdate
          Me.EmployeeTypeId = Me.EmpMangement.GetEmployeeTypeID(empTypeValue)
          Me.DepartmentId = Me.EmpMangement.GetDepartmentID(empDeptValue)

          Return True
        Else
          MessageBox.Show("Birth date is not valid", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Error)
          Me.bdateError.Visible = True
          Return False
        End If
      Else
        MessageBox.Show("Please complete all require attribute", "Invalid input", MessageBoxButtons.OK, MessageBoxIcon.Error)
        Return False
      End If
    Catch ex As Exception
      MessageBox.Show(ex.Message, "Error: 0x101 - setting up employee info", MessageBoxButtons.OK, MessageBoxIcon.Error)
      Return False
    End Try

    Return False
  End Function

  Private Function CheckValue(ByVal text As String, ByRef input As Object) As Boolean
    If text <> "" Then
      input.Visible = False
      Return True
    Else
      input.Visible = True
    End If
    Return False
  End Function

  Private Sub AddNewEmployeeBtn_Click(sender As System.Object, e As System.EventArgs) Handles AddNewEmployeeBtn.Click

    If MenusFrm.adminPermission = "wrx" Or MenusFrm.adminPermission = "rw" Then
      If Me.SetEmployeeInfo() Then
        If Me.EmpMangement.AddNewEmployee(Me.FirstName, Me.MiddleName, Me.LastName, Me.Gender, Me.Address, Me.BirthDate, Me.EmployeeTypeId, Me.DepartmentId) Then
          MessageBox.Show("New Employee Successfully added", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
          Me.ResetAllInputBox()
          Me.UpdateEmployeesDataGridView()
        End If
      End If
    Else
      MessageBox.Show("You don't have permission to add new Employee", "Permission denied", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End If
  End Sub

  Private Sub UpdateEmployeesDataGridView()
    AddRemoveEmployeeMainFrm.AllEmployeesDGV.Rows.Clear()
    Me.EmpMangement.GetAllEmployees(AddressOf AddRemoveEmployeeMainFrm.DisplayEmployeeDataGrid)
  End Sub

  Private Sub ResetAllInputBox()
    EmpFirstNameTxt.Text = ""
    EmpMiddleNameTxt.Text = ""
    EmpLastNameTxt.Text = ""
    EmpGenderCbox.ResetText()
    EmpAddressTxt.Text = ""
    EmpTypeCBox.ResetText()
    EmpDepartmentCBox.ResetText()
  End Sub


  ' For updating Employee information

  Private Sub UpdateEmployeeInfoBtn_Click(sender As System.Object, e As System.EventArgs) Handles UpdateEmployeeInfoBtn.Click

    If MenusFrm.adminPermission = "wrx" Or MenusFrm.adminPermission = "rw" Then
      If Me.SetEmployeeInfo() Then

        If Me.EmployeeId <> -1 Then
          'UpdateEmployeeInfo
          If Me.EmpMangement.UpdateEmployeeInfo(Me.EmployeeId, Me.FirstName, Me.MiddleName, Me.LastName, Me.Gender, Me.Address, Me.BirthDate, Me.EmployeeTypeId, Me.DepartmentId) Then
            MessageBox.Show("Updated", "Successful", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Me.ResetAllInputBox()
            Me.UpdateEmployeesDataGridView()
            Me.Close()
          End If
        End If

      End If
    Else
      MessageBox.Show("You don't have permission to add new Employee", "Permission denied", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End If

  End Sub

  'Private Sub AddEmployeeFrm_FormClosing(sender As System.Object, e As System.Windows.Forms.FormClosingEventArgs) Handles MyBase.FormClosing
  '    Me.UpdateEmployeesDataGridView()
  'End Sub
End Class