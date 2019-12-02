Imports System.IO
'Imports System.Linq
Imports System.Runtime.InteropServices

Public Class Form1
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call Import_VCF_Files()
    End Sub

    Public Sub Import_VCF_Files()


        Dim iPtr As Long
        Dim sFileArray As Object
        Dim sFileName As Object
        Dim dtStart As Date
        'classa
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "D:\my dbs"
        fd.Filter = "Vcf files (*.vcf*)|*.vcf*|vcf files (*.vcf*)|*.*"
        fd.FilterIndex = 9
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then sFileArray = fd.FileName


        'Dim ActiveWorkbook As Object = "D:\1.vcf"
        'iPtr = InStrRev(ActiveWorkbook.FullName, ".")
        'If iPtr = 0 Then
        '    sFileName = ActiveWorkbook.FullName & ".vcf"
        'Else
        '    ' sFileName = Left(ActiveWorkbook.FullName, iPtr - 1) & ".vcf"
        '    sFileName = ActiveWorkbook.FullName & ".vcf"
        'End If
        'sFileArray = fd.ShowDialog(FileFilter:="VCF files (*.vcf), *.vcf", MultiSelect:=True)

        Dim iFileCount, iRowPointer As Integer
        iFileCount = 0
        iRowPointer = 0

        ' sFileName = "D:\my dbs\1.vcf"

        'For Each sFileName In sFileArray
        '    ' MsgBox(fd.FileName)
        Call Ahmed(fd.FileName)

        '    iFileCount = iFileCount + 1
        'Next sFileName


    End Sub

    Private Sub Import_Contacts(ByVal argFilename As String)
        '  Range("f1").Value = 0
        Dim intFH As Integer
        Dim sRecord, lastrow As String
        Dim curstr, pcurstr As Integer

        curstr = 0
        Dim objfso, strTextFile, strData, arrLines, LineCount
        'Const ForReading = 1

        'name of the text file
        strTextFile = argFilename

        'Create a File System Object
        objfso = CreateObject("Scripting.FileSystemObject")

        'Open the text file - strData now contains the whole file
        '  strData = objfso.OpenTextFile(strTextFile, ForReading).ReadAll
        strData = File.ReadAllText(argFilename)
        'Split by lines, put into an array
        arrLines = Split(strData, vbCrLf)

        'Use UBound to count the lines
        LineCount = UBound(arrLines) + 1
        'MsgBox LineCount




        intFH = FreeFile()
        '  argFilename.
        Using StreamReader As System.IO.StreamReader = System.IO.File.OpenText(argFilename)
            While Not StreamReader.EndOfStream
                argFilename = StreamReader.ReadLine()

            End While
        End Using

        Dim orgrow, destrow, destcel, vc, position, lenstr, restname, prestname As Integer
        'lastrow = Cells(Rows.Count, 1).End(xlUp).Row + 1
        '  If lastrow = 5 Then lastrow = 6 Else lastrow = Cells(Rows.Count, 1).End(xlUp).Row + 1
        ' destrow = lastrow
        Dim iFileCount, iRowPointer As Integer
        iFileCount = 0
        iRowPointer = 0

        Do Until EOF(intFH)
            sRecord = LineInput(intFH)
            iRowPointer = iRowPointer + 1
            position = InStr(sRecord, ":")
            lenstr = Len(sRecord)
            vc = lenstr - position

            sRecord.Substring(0, 2)
            If sRecord.Substring(0, 2) = "FN" Then
                MsgBox(sRecord)
            End If
        Loop


    End Sub
    Sub Ahmed(ByVal argFilename As String)


        Dim TextLine As String = ""


        If System.IO.File.Exists(argFilename) = True Then
            Dim objReader As New System.IO.StreamReader(argFilename)
            Try
                Do While objReader.Peek() <> -1
                    DataGridView1.Rows.Add(TextLine)
                    'Do Until objReader.EndOfStream
                    TextLine = objReader.ReadLine()
                    If TextLine.Substring(0, 2) = "FN" Then
                        DataGridView1.Rows.Add(TextLine)
                        ' MsgBox(TextLine)
                    End If
                    'MsgBox(objReader.EndOfStream.ToString())
                    '  objReader.
                    For i = 1 To 2
                        ' MsgBox("d")

                        '   If TextLine.Substring(0, 2) = "FN" Then DataGridView1.Rows.Add(TextLine)
                        If objReader.EndOfStream = False Then Exit For
                        i = i + 1
                    Next

                    'If TextLine.Substring(0, 2) = "FN" Then
                    '        ' Me.DataGridView1.Rows.Add(TextLine)

                    '        If TextLine.Substring(0, 2) = "FN" Then DataGridView1.CurrentRow.Cells(1).Value = TextLine

                    '        If TextLine.Substring(0, 2) = "TE" Then DataGridView1.CurrentRow.Cells(2).Value = TextLine
                    '    End If

                Loop

            Catch ex As Exception

            End Try
        Else
            MsgBox("File Does Not Exist")
        End If

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

        Dim lines = IO.File.ReadAllLines("D:\my dbs\1.vcf")
        Dim position, lenstr, vc As Integer
        For Each line In lines
            position = InStr(line, ":")
            lenstr = Len(line)
            vc = lenstr - position

            If line.Substring(0, 2) = "FN" Then DataGridView1.Rows.Add(Strings.Right(line, vc))
            If line.Substring(0, 2) = "TE" Then DataGridView1.CurrentRow.Cells("Phone").Value = Strings.Right(line, vc)

        Next
        'DataGridView1.DataSource = dt


        'Dim TextLine = objReader.ReadLine()
        'Dim textint As Integer = 0
        'Do Until objReader.Peek() <> -1
        '    textint = textint + 1
        '    If objReader.EndOfStream = False Then Exit Do
        'Loop
        'MsgBox(textint)
        'DataGridView1.Rows.Add(30)
        'Dim RowNum As Integer = 0
        'For RowNum = 0 To 10
        '    DataGridView1.Rows(RowNum).Cells("phone").Value = "test value"
        'Next
    End Sub
    Sub DOG(ByVal argFilename As String)
        'Dim lines = IO.File.ReadAllLines(argFilename)
        'Dim intFH As Integer

        'Dim objfso = CreateObject("Scripting.FileSystemObject")
        'Dim strData = objfso.OpenTextFile(argFilename).ReadAll
        'Dim arrLines = Split(strData, vbCrLf)
        'Dim LineCount = UBound(arrLines) + 1
        'intFH = FreeFile()
        'IO.File.Open(argFilename, FileMode.Open)


        Dim fileReader As String
        fileReader = My.Computer.FileSystem.ReadAllText(argFilename)
        MsgBox(fileReader)







        'For Each line In lines
        '    Dim mystr As String = line
        'Next

        'Do While mystr <> ""
        'Dim position, lenstr, vc As Integer
        'position = InStr(line, ":")
        '    lenstr = Len(line)
        '    vc = lenstr - position
        '    Try


        '        If line.Substring(0, 2) = "FN" Then DataGridView1.Rows.Add(Strings.Right(line, vc))

        '        'If line.Substring(0, 2) = "TE" Then DataGridView1.Rows(RowNum).Cells("Phone").Value = (Strings.Right(line, vc))

        '    Catch ex As Exception
        '        MsgBox(ex.Message)

        '    End Try
        'Loop


    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim fileReader As String
        fileReader = My.Computer.FileSystem.ReadAllText("D:\my dbs\1.vcf")

        Dim arrLines = Split(fileReader, vbCrLf)
        Dim LineCount = UBound(arrLines) + 1

        MsgBox(fileReader)

    End Sub
    Sub nbnb()
        Dim fn As String = "D:\my dbs\1.vcf"
        Dim sr As New StreamReader(fn)
        Dim LineFromFile As String = Nothing
        Dim textRowNo As Integer = 0
        Dim arrival As String = Nothing
        Dim status As String = Nothing

        '...
        While Not sr.EndOfStream
            textRowNo += 1
            LineFromFile = sr.ReadLine()
            If textRowNo > 0 Then
                LineFromFile = sr.ReadLine()
                DataGridView1.Rows.Add()
                Dim position, lenstr, vc As Integer
                position = InStr(LineFromFile, ":")
                lenstr = Len(LineFromFile)
                vc = lenstr - position
                Dim rindex As Integer
                'If LineFromFile.Substring(0, 2) = "FN" Then
                '    DataGridView1.Rows.Add(Strings.Right(LineFromFile, vc))
                '    rindex = DataGridView1.CurrentRow.Index

                'End If
                'If LineFromFile.Substring(0, 2) = "TE" Then DataGridView1.Rows(rindex).Cells("Phone").Value = (Strings.Right(LineFromFile, vc))

                '...
            End If
        End While
        sr.Close()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call nbnb()
    End Sub
End Class
