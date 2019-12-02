Imports System.Data.OleDb
Imports System.IO
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form3

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Call GetThefile()
    End Sub
    Public Sub GetThefile()
        Dim sFileArray As Object
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = "Open File Dialog"
        ' fd.InitialDirectory = "D:\my dbs"
        fd.Filter = "Excel files (*.xlsm*)|*.xlsm*"
        fd.FilterIndex = 9
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then
            sFileArray = fd.FileName
            Label13.Text = "جاري تحميل اسماء الصفحات المتوفرة"
        Else
            Label13.Text = ""
            Exit Sub
        End If
        Label10.Text = fd.FileName.ToString
        Dim oXL As Excel.Application = New Excel.Application

        Dim oWB As Excel.Workbook = oXL.Workbooks.Open(fd.FileName)

        oXL.Visible = False


        For Each oSheet As Excel.Worksheet In oWB.Sheets
            ComboBoxSheet.Items.Add(oSheet.Name)
            '  ComboBoxName.Items.Add(oSheet.Rows(1).cells)

            Dim objRng As Excel.Range
            ' Dim strCol, strCell As String
            Dim maxCol, maxRow As Integer
            Dim iRow, iCol As Integer
            maxRow = 1
            maxCol = 50
            For iCol = 1 To maxCol
                For iRow = 1 To 1
                    objRng = oSheet.Cells(iRow, iCol)
                    If objRng.Value IsNot Nothing Then
                        ComboBoxName.Items.Add(objRng.Value)
                        ComboBoxPhone1.Items.Add(objRng.Value)
                        ComboBoxPhone2.Items.Add(objRng.Value)
                        ComboBoxPhone3.Items.Add(objRng.Value)
                        ComboBoxPhone4.Items.Add(objRng.Value)
                        ComboBoxAddress.Items.Add(objRng.Value)
                        ComboBoxOrg.Items.Add(objRng.Value)
                        ComboBoxEmail.Items.Add(objRng.Value)
                        ComboBoxNotes.Items.Add(objRng.Value)

                        Form2.DataGridView1.Rows().Add("", objRng.Value)
                        'MsgBox(iRow)
                        Form2.DataGridView1.Rows(iCol).Cells("Fname").Value = objRng.Value

                    End If
                Next
            Next
        Next
        ComboBoxSheet.SelectedIndex = 0

        Label13.Text = ""




    End Sub
    Sub AddExcelData()
        Dim oXL As Excel.Application = New Excel.Application
        Dim oWB As Excel.Workbook = oXL.Workbooks.Open(Label10.Text)
        oXL.Visible = False
        For Each oSheet As Excel.Worksheet In oWB.Sheets
            Dim objRng As Excel.Range
            Dim maxCol, maxRow As Integer
            Dim iRow, iCol As Integer
            maxRow = 50
            maxCol = ComboBoxName.SelectedIndex + 1
            For iCol = maxCol To maxCol
                For iRow = 2 To maxRow
                    objRng = oSheet.Cells(iRow, iCol)
                    If objRng.Value <> Nothing Then
                        Form2.DataGridView1.Rows().Add("", objRng.Value)
                    End If
                Next
            Next


        Next

        oWB.Close()

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call AddExcelData()
    End Sub
End Class