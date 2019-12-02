Imports System.ComponentModel
Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Text
Imports System.Threading

Public Class Form2
    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Dim sFileArray As Object
        Dim fd As OpenFileDialog = New OpenFileDialog()


        fd.Title = "اختر ملف الاسماء"
        fd.InitialDirectory = "D:\my dbs"
        fd.Filter = "Vcf files (*.vcf*)|*.vcf*"
        fd.FilterIndex = 9
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then

            sFileArray = fd.FileName

        Else
            Exit Sub
        End If
        Dim Myfilename As String = fd.FileName
        System.IO.Directory.CreateDirectory(Application.StartupPath & "\Recovery")
        Call exporttovcf(Application.StartupPath & "\Recovery\recovery at " & Format(Now, "yyyy-MM-dd  HH_mm-ss"))
        DataGridView1.Rows.Clear()
        Call Filterd(Myfilename)

    End Sub
    Sub Filterd(ByVal Filename As String)

        Dim TextLine As String = ""
        'Filename = ""
        If System.IO.File.Exists(Filename) = True Then
            Dim objReader As New System.IO.StreamReader(Filename)
            Try
                Do While objReader.Peek() <> -1
                    TextLine = objReader.ReadLine()
                    Dim ConstactsSum As Integer = InStr(TextLine, "FN")
                    DataGridView2.Rows.Add(TextLine)
                Loop
                DataGridView2.Rows.Add(1)
            Catch
            End Try
        End If
        '   MsgBox("تم الاستيراد الى شبكة البيانات الاولى  ")

        Dim Rowindexa, intarow, PAded, Imagejoin As Integer
        Dim position, lenstr, vc As Integer
        ' Rowindexa = DataGridView1.Rows.Count - 1
        Rowindexa = 0
        '  DataGridView1.Rows.Add(1)
        '  MsgBox(Rowindexa)
        intarow = -1
        PAded = 0
        'Rowindexa = TextBox1.Text

        Try
            Do Until DataGridView2.Rows(Rowindexa).Cells("OrgStr").RowIndex = DataGridView2.RowCount - 1
                Dim mysta = DataGridView2.Rows(Rowindexa).Cells("OrgStr").Value.ToString
                Dim MYNSTA = DataGridView2.Rows(Rowindexa + 1).Cells("OrgStr").Value.ToString

                position = InStr(mysta, ":")
                lenstr = Len(mysta)
                vc = lenstr - position
                Dim Namejoin As Integer
                Try

                    If mysta.Substring(0, 2) = "FN" Then

                        DataGridView1.Rows(intarow).Cells("Fname").Value = (Strings.Right(mysta, vc))
                        Namejoin = 1
                        Try
                            If InStr(MYNSTA, ":") = 0 And Namejoin = 1 Then DataGridView1.Rows(intarow).Cells("Fname").Value = DataGridView1.Rows(intarow).Cells("Fname").Value & (Strings.Right(MYNSTA, vc))
                            DataGridView1.Rows(intarow).Cells("Fname").Value = HexToAlfa(DataGridView1.Rows(intarow).Cells("Fname").Value)
                            ' ''' Label1.Text = HexToAlfa(DataGridView1.Rows(intarow).Cells("Fname").Value)
                        Catch
                        End Try
                    End If

                    If mysta.Substring(0, 2) = "TE" Then
                        Namejoin = 0
                        If DataGridView1.Rows(intarow).Cells("Phone").Value = "" Then
                            DataGridView1.Rows(intarow).Cells("Phone").Value = (Strings.Right(mysta, vc))
                        Else

                            If DataGridView1.Rows(intarow).Cells("Phone2").Value = "" Then
                                DataGridView1.Rows(intarow).Cells("Phone2").Value = (Strings.Right(mysta, vc))
                            Else
                                If DataGridView1.Rows(intarow).Cells("Phone3").Value = "" Then
                                    DataGridView1.Rows(intarow).Cells("Phone3").Value = (Strings.Right(mysta, vc))
                                Else
                                    If DataGridView1.Rows(intarow).Cells("Phone4").Value = "" Then
                                        DataGridView1.Rows(intarow).Cells("Phone4").Value = (Strings.Right(mysta, vc))
                                    Else
                                        If DataGridView1.Rows(intarow).Cells("Phone5").Value = "" Then
                                            DataGridView1.Rows(intarow).Cells("Phone5").Value = (Strings.Right(mysta, vc))
                                            ' DataGridView1.Rows(intarow).
                                        Else
                                            If DataGridView1.Rows(intarow).Cells("Phone5").Value <> "" Then
                                                DataGridView1.Rows(intarow).Cells("notes").Value = DataGridView1.Rows(intarow).Cells("notes").Value & " | " & (Strings.Right(mysta, vc))
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If mysta.Substring(0, 2) = "AD" Then DataGridView1.Rows(intarow).Cells("Address").Value = (Strings.Right(mysta, vc))
                    If mysta.Substring(0, 2) = "EM" Then DataGridView1.Rows(intarow).Cells("email").Value = (Strings.Right(mysta, vc))
                    If mysta.Substring(0, 2) = "Ti" Then DataGridView1.Rows(intarow).Cells("Title").Value = (Strings.Right(mysta, vc))
                    If mysta.Substring(0, 2) = "OR" Then DataGridView1.Rows(intarow).Cells("ORG").Value = (Strings.Right(mysta, vc))
                    If mysta.Substring(0, 2) = "PH" Then
                        DataGridView1.Rows(intarow).Cells("Photo").Value = (Strings.Right(mysta, vc))

                        Imagejoin = 1
                    End If
                    If mysta.Substring(0, 2) = "Ti" Then DataGridView1.Rows(intarow).Cells("Title").Value = (Strings.Right(mysta, vc))
                    If InStr(mysta, ":") = 0 And Imagejoin = 1 Then DataGridView1.Rows(intarow).Cells("Photo").Value = DataGridView1.Rows(intarow).Cells("Photo").Value & Chr(13) & (Strings.Right(mysta, vc))


                    If mysta.Substring(0, 2) = "BE" Then
                        DataGridView1.Rows.Add(1)
                        Imagejoin = 0
                        intarow = intarow + 1
                    End If
                Catch ex As Exception
                    '  MsgBox(ex.Message)
                End Try
                Dim irow As Integer
                For irow = 0 To DataGridView1.Rows.Count - 1
                    If DataGridView1.Rows(irow).Cells("email").Value IsNot Nothing And DataGridView1.Rows(irow).Cells("fname").Value Is Nothing Then DataGridView1.Rows(irow).Cells("fname").Value = DataGridView1.Rows(irow).Cells("email").Value

                Next
                If DataGridView1.Rows(intarow).Cells("Fname").Value <> "" Then DataGridView1.Rows(intarow).Cells("Serial").Value = DataGridView1.Rows(intarow).Index + 1
                Rowindexa = Rowindexa + 1
                Dim C As DataGridViewButtonCell = DataGridView1.Rows(intarow).Cells("AddImg")
                Dim M As DataGridViewButtonColumn = DataGridView1.Columns("AddImg")
                ' M.UseColumnTextForButtonValue = True
                M.Text = "change"
                If DataGridView1.Rows(intarow).Cells("Photo").Value <> "" Then
                    C.Style.BackColor = Color.Green
                    C.ToolTipText = "Edit"
                    DataGridView1.Rows(intarow).Cells("AddIMG").Value = "Change"
                    '   C.text.UserColumnTextForButtonValue = True
                    ' M.Text = "change"
                    C.UseColumnTextForButtonValue = True
                    C.Value = "Ahmed"

                Else
                    C.Style.BackColor = Color.White
                End If
                ' Label1.Text = intarow
            Loop

        Catch ex As NullReferenceException

            '  MsgBox(ex.Message & "الاول")
        Catch ex1 As Exception
            MsgBox(ex1.Message & "  خطأ اثناء الاستيراد  ")
        End Try
        'Label1.Text = ""

        Beep()
        Form4.Hide()

        MsgBox("تم الانتهاء من الاستيراد والتصنيف")
        DataGridView2.Rows.Clear()

    End Sub
    Public Function DecodeBase64(ByVal strData As String) As Byte()
        Dim objXML As Object
        Dim objNode As Object
        objXML = CreateObject("MSXML2.DOMDocument")
        objNode = objXML.createElement("b64")
        objNode.DataType = "bin.base64"
        objNode.Text = strData
        DecodeBase64 = objNode.nodeTypedValue
        System.IO.File.WriteAllBytes(Application.UserAppDataPath() & "currentimg.tmp", DecodeBase64)
        objNode.Text = "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"
        DecodeBase64 = objNode.nodeTypedValue
        System.IO.File.WriteAllBytes(Application.UserAppDataPath() & "empty.tmp", DecodeBase64)
        objNode = Nothing
        objXML = Nothing
        ' PictureBox1.ImageLocation = Application.UserAppDataPath() & "\Images\" & DataGridView1.Rows(irow).Cells("fname").Value
    End Function

    Sub getimag()
        Try
            Dim C As DataGridViewButtonCell = DataGridView1.CurrentRow.Cells("AddImg")
            Dim M As DataGridViewButtonColumn = DataGridView1.Columns("AddImg")

            If DataGridView1.CurrentRow.Cells("Photo").Value <> "" Then
                DecodeBase64(DataGridView1.CurrentRow.Cells("Photo").Value)
                PictureBox1.ImageLocation = Application.UserAppDataPath() & "currentimg.tmp"

                C.Style.BackColor = Color.Green
                C.ToolTipText = "Edit"
                DataGridView1.CurrentRow.Cells("AddIMG").Value = "Change"
                C.UseColumnTextForButtonValue = True
                C.Value = "Ahmed"

            Else
                C.Style.BackColor = Color.White
                PictureBox1.ImageLocation = Application.UserAppDataPath() & "Empty.tmp"
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            ' MsgBox("غير قادر على جلب الصورة")
        End Try
        '  Call MakeItRound()
    End Sub
    Sub SaveAllImages()
        Dim istherefiles As Integer
        For irow = 1 To DataGridView1.Rows.Count - 1

            Try
                Dim C As DataGridViewButtonCell = DataGridView1.CurrentRow.Cells("AddImg")
                Dim M As DataGridViewButtonColumn = DataGridView1.Columns("AddImg")

                If DataGridView1.Rows(irow).Cells("Photo").Value <> "" Then
                    DecodeBase64All(DataGridView1.Rows(irow).Cells("Photo").Value, DataGridView1.Rows(irow).Cells("fname").Value)
                    istherefiles = 1
                End If
            Catch ex As Exception
                MsgBox(ex.Message)
                ' MsgBox("غير قادر على جلب الصورة")
            End Try
            '  Call MakeItRound()
        Next
        If istherefiles = 1 Then
            Dim imagespath As String = Application.StartupPath() & "\Images_" & Format(Now, "yyyy-MM-dd") & "\"
            '  MsgBox(Format((System.IO.Directory.GetLastWriteTime(imagespath)), "yyyy-MM-dd"))
            MsgBox("تم حفظ الصور   " & imagespath)
        Else
            MsgBox("لا يوجد صور لحفظها")
        End If
    End Sub

    Private Sub DataGridView1_RowEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            Call getimag()
        Catch
        End Try
        If e.ColumnIndex = 14 Then
            ' MsgBox(("Row : " + e.RowIndex.ToString & "  Col : ") + e.ColumnIndex.ToString)
            Call AddImage()
        End If

    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Call getimag()

    End Sub


    Private Sub DataGridView1_CellEnter(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellEnter
        Call getimag()
    End Sub
    Public Function HexToAlfa(result) As String
        result = Replace(result, ";", "")
        'sRecord = Replace(sRecord, "=", "")

        'result = Replace(result, False, "0")
        'result = Replace(result, "=DB=B0", "0")
        'result = Replace(result, "=D9=A1", "1")
        'result = Replace(result, "=DB=B1", "1")
        'result = Replace(result, "=D9=A2", "2")
        'result = Replace(result, "=DB=B2", "2")
        'result = Replace(result, "=D9=A3", "3")
        'result = Replace(result, "=DB=B3", "3")
        'result = Replace(result, "=D9=A4", "4")
        'result = Replace(result, "=DB=B4", "4")
        'result = Replace(result, "=D9=A5", "5")
        'result = Replace(result, "=DB=B5", "5")
        'result = Replace(result, "=D9=A6", "6")
        'result = Replace(result, "=DB=B6", "6")
        'result = Replace(result, "=D9=A7", "7")
        'result = Replace(result, "=DB=B7", "7")
        'result = Replace(result, "=D9=A8", "8")
        'result = Replace(result, "=DB=B8", "8")
        'result = Replace(result, "=D9=A9", "9")
        'result = Replace(result, "=DB=B9", "9")
        result = Replace(result, "=DB=A5", "ۥ")
        result = Replace(result, "=D9=80", "ـ")
        result = Replace(result, "=DB=A6", "ۦ")
        result = Replace(result, "=D8=82", "؂")
        result = Replace(result, "=D8=9E", "؞")
        result = Replace(result, "=D8=8C", "،")
        result = Replace(result, "=D8=9B", "؛")
        result = Replace(result, "=D8=9F", "؟")
        result = Replace(result, "=D9=AA", "٪")
        result = Replace(result, "=D9=AB", "٫")
        result = Replace(result, "=D9=AC", "٬")
        result = Replace(result, "=DB=94", "۔")
        result = Replace(result, "=D8=8B", "؋")
        result = Replace(result, "=E2=97=8C=DB=AC", "◌۬")
        result = Replace(result, "=E2=97=8C=D8=91", "◌ؑ")
        result = Replace(result, "=E2=97=8C=D8=92", "◌ؒ")
        result = Replace(result, "=E2=97=8C=D8=93", "◌ؓ")
        result = Replace(result, "=E2=97=8C=D8=94", "◌ؔ")
        result = Replace(result, "=E2=97=8C=DB=9B", "◌ۛ")
        result = Replace(result, "=E2=97=8C=D9=96", "◌ٖ")
        result = Replace(result, "=E2=97=8C=D9=97", "◌ٗ")
        result = Replace(result, "=E2=97=8C=D8=90", "◌ؐ")
        result = Replace(result, "=E2=97=8C=DB=AB", "◌۫")
        result = Replace(result, "=E2=97=8C=D8=95", "◌ؕ")
        result = Replace(result, "=E2=97=8C=DB=99", "◌ۙ")
        result = Replace(result, "=E2=97=8C=DB=A4", "◌ۤ")
        result = Replace(result, "=E2=97=8C=DB=A1", "◌ۡ")
        result = Replace(result, "=E2=97=8C=DB=9A", "◌ۚ")
        result = Replace(result, "=E2=97=8C=DB=9C", "◌ۜ")
        result = Replace(result, "=E2=97=8C=DB=AA", "◌۪")
        result = Replace(result, "=E2=97=8C=DB=A3", "◌ۣ")
        result = Replace(result, "=E2=97=8C=DB=98", "◌ۘ")
        result = Replace(result, "=E2=97=8C=D9=95", "◌ٕ")
        result = Replace(result, "=E2=97=8C=D9=9F", "◌ٟ")
        result = Replace(result, "=E2=97=8C=D9=94", "◌ٔ")
        result = Replace(result, "=E2=97=8C=DB=A7", "◌ۧ")
        result = Replace(result, "=E2=97=8C=D9=93", "◌ٓ")
        result = Replace(result, "=E2=97=8C=DB=A8", "◌ۨ")
        result = Replace(result, "=E2=97=8C=DB=AD", "◌ۭ")
        result = Replace(result, "=E2=97=8C=DB=A2", "◌ۢ")
        result = Replace(result, "=E2=97=8C=D9=98", "◌٘")
        result = Replace(result, "=E2=97=8C=D9=8B", "◌ً")
        result = Replace(result, "=E2=97=8C=D9=8C", "◌ٌ")
        result = Replace(result, "=E2=97=8C=D9=8D", "◌ٍ")
        result = Replace(result, "=E2=97=8C=D9=8E", "◌َ")
        result = Replace(result, "=E2=97=8C=D9=8F", "◌ُ")
        result = Replace(result, "=E2=97=8C=D9=90", "◌ِ")
        result = Replace(result, "=E2=97=8C=D9=92", "◌ْ")
        result = Replace(result, "=E2=97=8C=D9=91", "◌ّ")
        result = Replace(result, "=E2=97=8C=D9=B0", "◌ٰ")
        result = Replace(result, "=E2=97=8C=D9=99", "◌ٙ")
        result = Replace(result, "=E2=97=8C=D9=9A", "◌ٚ")
        result = Replace(result, "=E2=97=8C=D9=9B", "◌ٛ")
        result = Replace(result, "=E2=97=8C=D9=9C", "◌ٜ")
        result = Replace(result, "=E2=97=8C=D9=9D", "◌ٝ")
        result = Replace(result, "=E2=97=8C=D9=9E", "◌ٞ")
        result = Replace(result, "=E2=97=8C=D8=96", "◌ؖ")
        result = Replace(result, "=E2=97=8C=D8=97", "◌ؗ")
        result = Replace(result, "=E2=97=8C=D8=98", "◌ؘ")
        result = Replace(result, "=E2=97=8C=D8=99", "◌ؙ")
        result = Replace(result, "=E2=97=8C=D8=9A", "◌ؚ")
        result = Replace(result, "=E2=97=8C=DB=A0", "◌۠")
        result = Replace(result, "=E2=97=8C=DB=9F", "◌۟")
        result = Replace(result, "=E2=97=8C=D9=B4", "◌ٴ")
        result = Replace(result, "=E2=97=8C=DB=96", "◌ۖ")
        result = Replace(result, "=E2=97=8C=DB=97", "◌ۗ")
        result = Replace(result, "=DB=BD", "۽")
        result = Replace(result, "=DB=9D", "۝")
        result = Replace(result, "=D8=8E", "؎")
        result = Replace(result, "=DB=BE", "۾")
        result = Replace(result, "=D8=80", "؀")
        result = Replace(result, "=DB=A9", "۩")
        result = Replace(result, "=D9=AD", "٭")
        result = Replace(result, "=DB=9E", "۞")
        result = Replace(result, "=D8=8F", "؏")
        result = Replace(result, "=D8=8D", "؍")
        result = Replace(result, "=D8=81", "؁")
        result = Replace(result, "=D8=83", "؃")
        result = Replace(result, "=D8=86", "؆")
        result = Replace(result, "=D8=87", "؇")
        result = Replace(result, "=D8=88", "؈")
        result = Replace(result, "=D8=89", "؉")
        result = Replace(result, "=D8=8A", "؊")
        result = Replace(result, "=D8=A1", "ء")
        result = Replace(result, "=D9=B1", "ٱ")
        result = Replace(result, "=D9=B3", "ٳ")
        result = Replace(result, "=D9=B2", "ٲ")
        result = Replace(result, "=D8=A7", "ا")
        result = Replace(result, "=D8=A5", "إ")
        result = Replace(result, "=D8=A3", "أ")
        result = Replace(result, "=D8=A2", "آ")
        result = Replace(result, "=D9=B5", "ٵ")
        result = Replace(result, "=D8=A8", "ب")
        result = Replace(result, "=D9=BB", "ٻ")
        result = Replace(result, "=D9=AE", "ٮ")
        result = Replace(result, "=DA=80", "ڀ")
        result = Replace(result, "=D9=BE", "پ")
        result = Replace(result, "=D8=A9", "ة")
        result = Replace(result, "=DB=83", "ۃ")
        result = Replace(result, "=D8=AA", "ت")
        result = Replace(result, "=D9=BA", "ٺ")
        result = Replace(result, "=D9=BF", "ٿ")
        result = Replace(result, "=D9=BC", "ټ")
        result = Replace(result, "=D9=BD", "ٽ")
        result = Replace(result, "=D9=B9", "ٹ")
        result = Replace(result, "=D8=AB", "ث")
        result = Replace(result, "=D8=AC", "ج")
        result = Replace(result, "=DA=83", "ڃ")
        result = Replace(result, "=DA=84", "ڄ")
        result = Replace(result, "=DA=86", "چ")
        result = Replace(result, "=DA=BF", "ڿ")
        result = Replace(result, "=DA=87", "ڇ")
        result = Replace(result, "=D8=AD", "ح")
        result = Replace(result, "=DA=81", "ځ")
        result = Replace(result, "=DA=82", "ڂ")
        result = Replace(result, "=DA=85", "څ")
        result = Replace(result, "=D8=AE", "خ")
        result = Replace(result, "=D8=AF", "د")
        result = Replace(result, "=DB=AE", "ۮ")
        result = Replace(result, "=DA=8B", "ڋ")
        result = Replace(result, "=DA=88", "ڈ")
        result = Replace(result, "=DA=89", "ډ")
        result = Replace(result, "=DA=8A", "ڊ")
        result = Replace(result, "=DA=8D", "ڍ")
        result = Replace(result, "=DA=8E", "ڎ")
        result = Replace(result, "=DA=8F", "ڏ")
        result = Replace(result, "=DA=90", "ڐ")
        result = Replace(result, "=D8=B0", "ذ")
        result = Replace(result, "=DA=8C", "ڌ")
        result = Replace(result, "=D8=B1", "ر")
        result = Replace(result, "=DA=95", "ڕ")
        result = Replace(result, "=DA=92", "ڒ")
        result = Replace(result, "=DB=AF", "ۯ")
        result = Replace(result, "=DA=94", "ڔ")
        result = Replace(result, "=DA=96", "ږ")
        result = Replace(result, "=DA=97", "ڗ")
        result = Replace(result, "=DA=91", "ڑ")
        result = Replace(result, "=DA=93", "ړ")
        result = Replace(result, "=D8=B2", "ز")
        result = Replace(result, "=DA=99", "ڙ")
        result = Replace(result, "=DA=98", "ژ")
        result = Replace(result, "=D8=B3", "س")
        result = Replace(result, "=DA=9B", "ڛ")
        result = Replace(result, "=DA=9A", "ښ")
        result = Replace(result, "=DA=9C", "ڜ")
        result = Replace(result, "=D8=B4", "ش")
        result = Replace(result, "=DB=BA", "ۺ")
        result = Replace(result, "=D8=B5", "ص")
        result = Replace(result, "=DA=9D", "ڝ")
        result = Replace(result, "=DA=9E", "ڞ")
        result = Replace(result, "=D8=B6", "ض")
        result = Replace(result, "=DB=BB", "ۻ")
        result = Replace(result, "=D8=B7", "ط")
        result = Replace(result, "=DA=9F", "ڟ")
        result = Replace(result, "=D8=B8", "ظ")
        result = Replace(result, "=D8=B9", "ع")
        result = Replace(result, "=DA=A0", "ڠ")
        result = Replace(result, "=D8=BA", "غ")
        result = Replace(result, "=DB=BC", "ۼ")
        result = Replace(result, "=DA=A1", "ڡ")
        result = Replace(result, "=D9=81", "ف")
        result = Replace(result, "=DA=A2", "ڢ")
        result = Replace(result, "=DA=A3", "ڣ")
        result = Replace(result, "=DA=A4", "ڤ")
        result = Replace(result, "=DA=A5", "ڥ")
        result = Replace(result, "=DA=A6", "ڦ")
        result = Replace(result, "=D9=AF", "ٯ")
        result = Replace(result, "=D9=82", "ق")
        result = Replace(result, "=DA=A7", "ڧ")
        result = Replace(result, "=DA=A8", "ڨ")
        result = Replace(result, "=D9=83", "ك")
        result = Replace(result, "=DA=AB", "ګ")
        result = Replace(result, "=DA=AE", "ڮ")
        result = Replace(result, "=DA=AC", "ڬ")
        result = Replace(result, "=DA=AD", "ڭ")
        result = Replace(result, "=DA=A9", "ک")
        result = Replace(result, "=D8=BB", "ػ")
        result = Replace(result, "=D8=BC", "ؼ")
        result = Replace(result, "=DA=AA", "ڪ")
        result = Replace(result, "=DA=AF", "گ")
        result = Replace(result, "=DA=B0", "ڰ")
        result = Replace(result, "=DA=B1", "ڱ")
        result = Replace(result, "=DA=B3", "ڳ")
        result = Replace(result, "=DA=B2", "ڲ")
        result = Replace(result, "=DA=B4", "ڴ")
        result = Replace(result, "=D9=84", "ل")
        result = Replace(result, "=DA=B5", "ڵ")
        result = Replace(result, "=DA=B8", "ڸ")
        result = Replace(result, "=DA=B6", "ڶ")
        result = Replace(result, "=DA=B7", "ڷ")
        result = Replace(result, "=D9=85", "م")
        result = Replace(result, "=DA=BA", "ں")
        result = Replace(result, "=D9=86", "ن")
        result = Replace(result, "=DA=BC", "ڼ")
        result = Replace(result, "=DA=BB", "ڻ")
        result = Replace(result, "=DA=B9", "ڹ")
        result = Replace(result, "=DA=BD", "ڽ")
        result = Replace(result, "=D9=87", "ه")
        result = Replace(result, "=DB=BF", "ۿ")
        result = Replace(result, "=DB=81", "ہ")
        result = Replace(result, "=DA=BE", "ھ")
        result = Replace(result, "=DB=82", "ۂ")
        result = Replace(result, "=D9=88", "و")
        result = Replace(result, "=DB=84", "ۄ")
        result = Replace(result, "=DB=8F", "ۏ")
        result = Replace(result, "=DB=8A", "ۊ")
        result = Replace(result, "=D8=A4", "ؤ")
        result = Replace(result, "=D9=B6", "ٶ")
        result = Replace(result, "=DB=86", "ۆ")
        result = Replace(result, "=DB=85", "ۅ")
        result = Replace(result, "=DB=87", "ۇ")
        result = Replace(result, "=D9=B7", "ٷ")
        result = Replace(result, "=DB=88", "ۈ")
        result = Replace(result, "=DB=89", "ۉ")
        result = Replace(result, "=DB=8B", "ۋ")
        result = Replace(result, "=DB=90", "ې")
        result = Replace(result, "=DB=8D", "ۍ")
        result = Replace(result, "=D9=89", "ى")
        result = Replace(result, "=D9=8A", "ي")
        result = Replace(result, "=DB=8E", "ێ")
        result = Replace(result, "=DB=91", "ۑ")
        result = Replace(result, "=DB=92", "ے")
        result = Replace(result, "=DB=8C", "ی")
        result = Replace(result, "=D8=BD", "ؽ")
        result = Replace(result, "=D8=BE", "ؾ")
        result = Replace(result, "=D8=BF", "ؿ")
        result = Replace(result, "=D8=A6", "ئ")
        result = Replace(result, "=DB=93", "ۓ")
        result = Replace(result, "=D9=B8", "ٸ")
        result = Replace(result, "=DB=95", "ە")
        result = Replace(result, "=DB=80", "ۀ")
        result = Replace(result, "=D8=A0", "ؠ")
        result = Replace(result, "=", "")
        result = Replace(result, "E299A1", "♥")
        result = Replace(result, "21", "!")
        result = Replace(result, "23", "#")
        result = Replace(result, "24", "$")
        result = Replace(result, "25", "%")
        result = Replace(result, "26", "&")
        result = Replace(result, "27", "'")
        result = Replace(result, "28", "(")
        result = Replace(result, "29", ")")
        result = Replace(result, "2A", "*")
        result = Replace(result, "2B", "+")
        result = Replace(result, "2C", ",")
        result = Replace(result, "2D", "-")
        result = Replace(result, "2E", ".")
        result = Replace(result, "2F", "/")
        result = Replace(result, "30", "0")
        result = Replace(result, "31", "1")
        result = Replace(result, "32", "2")
        result = Replace(result, "33", "3")
        result = Replace(result, "34", "4")
        result = Replace(result, "35", "5")
        result = Replace(result, "36", "6")
        result = Replace(result, "37", "7")
        result = Replace(result, "38", "8")
        result = Replace(result, "39", "9")
        result = Replace(result, "3A", ":")
        result = Replace(result, "3B", ";")
        result = Replace(result, "3C", "<")
        result = Replace(result, "3D", "=")
        result = Replace(result, "3E", ">")
        result = Replace(result, "3F", "?")
        result = Replace(result, "40", "@")
        result = Replace(result, "41", "A")
        result = Replace(result, "42", "B")
        result = Replace(result, "43", "C")
        result = Replace(result, "44", "D")
        result = Replace(result, "45", "E")
        result = Replace(result, "46", "F")
        result = Replace(result, "47", "G")
        result = Replace(result, "48", "H")
        result = Replace(result, "49", "I")
        result = Replace(result, "4A", "J")
        result = Replace(result, "4B", "K")
        result = Replace(result, "4C", "L")
        result = Replace(result, "4D", "M")
        result = Replace(result, "4E", "N")
        result = Replace(result, "4F", "O")
        result = Replace(result, "50", "P")
        result = Replace(result, "51", "Q")
        result = Replace(result, "52", "R")
        result = Replace(result, "53", "S")
        result = Replace(result, "54", "T")
        result = Replace(result, "55", "U")
        result = Replace(result, "56", "V")
        result = Replace(result, "57", "W")
        result = Replace(result, "58", "X")
        result = Replace(result, "59", "Y")
        result = Replace(result, "5A", "Z")
        result = Replace(result, "5B", "[")
        result = Replace(result, "5C", "\")
        result = Replace(result, "5D", "]")
        result = Replace(result, "5E", "^")
        result = Replace(result, "5F", "_")
        result = Replace(result, "60", "`")
        result = Replace(result, "61", "a")
        result = Replace(result, "62", "b")
        result = Replace(result, "63", "c")
        result = Replace(result, "64", "d")
        result = Replace(result, "65", "e")
        result = Replace(result, "66", "f")
        result = Replace(result, "67", "g")
        result = Replace(result, "68", "h")
        result = Replace(result, "69", "i")
        result = Replace(result, "6A", "j")
        result = Replace(result, "6B", "k")
        result = Replace(result, "6C", "l")
        result = Replace(result, "6D", "m")
        result = Replace(result, "6E", "n")
        result = Replace(result, "6F", "o")
        result = Replace(result, "70", "p")
        result = Replace(result, "71", "q")
        result = Replace(result, "72", "r")
        result = Replace(result, "73", "s")
        result = Replace(result, "74", "t")
        result = Replace(result, "75", "u")
        result = Replace(result, "76", "v")
        result = Replace(result, "77", "w")
        result = Replace(result, "78", "x")
        result = Replace(result, "79", "y")
        result = Replace(result, "7A", "z")
        result = Replace(result, "7B", "{")
        result = Replace(result, "7C", "|")
        result = Replace(result, "7D", "}")
        result = Replace(result, "7E", "~")
        result = Replace(result, "20", " ")



        'result = Replace(result, "=20", " ")
        HexToAlfa = result
    End Function
    Private Sub AddImage()
        Dim sFileArray As Object
        Dim fd As OpenFileDialog = New OpenFileDialog()
        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "D:\my dbs"
        fd.Filter = "Images files (*.jpg*)|*.jpg*"
        fd.FilterIndex = 9
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then sFileArray = fd.FileName Else Exit Sub
        Const adTypeBinary = 1
        Dim objXML
        Dim objDocElem
        Dim objStream
        objStream = CreateObject("ADODB.Stream")
        objStream.Type = adTypeBinary
        objStream.Open
        objStream.LoadFromFile(fd.FileName)
        objXML = CreateObject("MSXml2.DOMDocument")
        objDocElem = objXML.createElement("Base64Data")
        objDocElem.DataType = "bin.base64"
        objDocElem.nodeTypedValue = objStream.Read()
        DataGridView1.CurrentRow.Cells("Photo").Value = objDocElem.Text
        objXML = Nothing
        objDocElem = Nothing
        objStream = Nothing
        Call getimag()
    End Sub

    Public Function EncodeImage(strPicPath As Object) As String

        Const adTypeBinary = 1
        Dim objXML
        Dim objDocElem
        Dim objStream
        objStream = CreateObject("ADODB.Stream")
        objStream.Type = adTypeBinary
        objStream.Open
        objStream.LoadFromFile(strPicPath)
        objXML = CreateObject("MSXml2.DOMDocument")
        objDocElem = objXML.createElement("Base64Data")
        objDocElem.DataType = "bin.base64"
        objDocElem.nodeTypedValue = objStream.Read()
        EncodeImage = objDocElem.Text
        DataGridView1.CurrentRow.Cells("Photo").Value = objDocElem.Text
        objXML = Nothing
        objDocElem = Nothing
        objStream = Nothing

    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs)
        Call AddImage()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs)
        Form3.Show()
    End Sub

    Private Sub datagrid_KeyDown(ByVal sender As System.Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles DataGridView1.KeyDown
        Try
            If e.Control And (e.KeyCode = Keys.C) Then
                Dim d As DataObject = DataGridView1.GetClipboardContent()
                Clipboard.SetDataObject(d)
                e.Handled = True
            ElseIf (e.Control And e.KeyCode = Keys.V) Then
                PasteUnboundRecords()
            End If
        Catch ex As Exception
            'Log Exception
        End Try
    End Sub

    Private Sub PasteUnboundRecords()

        Try
            Dim rowLines As String() = Clipboard.GetText(TextDataFormat.Text).Split(New String(0) {vbCr & vbLf}, StringSplitOptions.None)
            Dim currentRowIndex As Integer = (If(DataGridView1.CurrentRow IsNot Nothing, DataGridView1.CurrentRow.Index, 0))
            Dim currentColumnIndex As Integer = (If(DataGridView1.CurrentCell IsNot Nothing, DataGridView1.CurrentCell.ColumnIndex, 0))
            Dim currentColumnCount As Integer = DataGridView1.Columns.Count

            DataGridView1.AllowUserToAddRows = True
            For rowLine As Integer = 0 To rowLines.Length - 1

                If rowLine = rowLines.Length - 1 AndAlso String.IsNullOrEmpty(rowLines(rowLine)) Then
                    Exit For
                Else
                    ' DataGridView1.Rows.Add()
                End If

                Dim columnsData As String() = rowLines(rowLine).Split(New String(0) {vbTab}, StringSplitOptions.None)
                If (currentColumnIndex + columnsData.Length) > DataGridView1.Columns.Count Then
                    For columnCreationCounter As Integer = 0 To ((currentColumnIndex + columnsData.Length) - currentColumnCount) - 1
                        If columnCreationCounter = rowLines.Length - 1 Then
                            Exit For
                        End If
                    Next
                End If
                DataGridView1.Rows.Add()  ' works prefectly
                If DataGridView1.Rows.Count > (currentRowIndex + rowLine) Then
                    For columnsDataIndex As Integer = 0 To columnsData.Length - 1
                        '   DataGridView1.Rows.Add()  works but doubled
                        If currentColumnIndex + columnsDataIndex <= DataGridView1.Columns.Count - 1 Then
                            ' DataGridView1.Rows.Add()
                            DataGridView1.Rows(currentRowIndex + rowLine).Cells(currentColumnIndex + columnsDataIndex).Value = columnsData(columnsDataIndex)

                        End If
                    Next

                Else
                    Dim pasteCells As String() = New String(DataGridView1.Columns.Count - 1) {}
                    For cellStartCounter As Integer = currentColumnIndex To DataGridView1.Columns.Count - 1
                        If columnsData.Length > (cellStartCounter - currentColumnIndex) Then
                            pasteCells(cellStartCounter) = columnsData(cellStartCounter - currentColumnIndex)
                        End If

                    Next
                End If
            Next
            Call autoserial()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub
    Sub autoserial()
        Dim irow As Integer
        For irow = 0 To DataGridView1.Rows.Count - 1
            Try
                If DataGridView1.Rows(irow).Cells("Fname").Value <> "" Then DataGridView1.Rows(irow).Cells("Serial").Value = DataGridView1.Rows(irow).Index + 1

                ' If DataGridView1.Rows(irow).Cells("Fname").Value <> "" Then DataGridView1.Rows.Remove(DataGridView1.Rows(irow))
            Catch
            End Try
        Next
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs)

    End Sub

    Private Sub ToolStripMenuItemPaste_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemPaste.Click
        Call PasteUnboundRecords()
    End Sub
    Private Sub ToolStripMenuItemcopy_Click(sender As Object, e As EventArgs) Handles ToolStripMenuItemCopy.Click
        Dim d As DataObject = DataGridView1.GetClipboardContent()
        Clipboard.SetDataObject(d)
    End Sub
    Sub exporttovcf(myfile As String)
        Dim irow As Integer
        Dim dafname, damob1, damob2, damob3, damob4, damob5, daaddress, danotes, daorg, daemail, datitle, DAPHOTO, da13, da14 As String
        Dim filePath As String = myfile & ".vcf"
        Dim delimeter As String = ","
        Dim sb As New StringBuilder
        For irow = 0 To DataGridView1.Rows.Count - 1
            ' Do Until DataGridView1.Rows(irow).Cells("Fname").Value = Nothing      'Cell(irow, 1) <> "" And 'irow < threehun  'And onlineyear < validyear
            dafname = DataGridView1.Rows(irow).Cells("fname").Value
            damob1 = DataGridView1.Rows(irow).Cells("phone").Value
            damob2 = DataGridView1.Rows(irow).Cells("phone2").Value
            damob3 = DataGridView1.Rows(irow).Cells("phone3").Value
            damob4 = DataGridView1.Rows(irow).Cells("phone4").Value
            damob5 = DataGridView1.Rows(irow).Cells("phone5").Value
            daaddress = DataGridView1.Rows(irow).Cells("address").Value
            danotes = DataGridView1.Rows(irow).Cells("notes").Value
            daorg = DataGridView1.Rows(irow).Cells("org").Value
            datitle = DataGridView1.Rows(irow).Cells("title").Value
            daemail = DataGridView1.Rows(irow).Cells("email").Value
            DAPHOTO = DataGridView1.Rows(irow).Cells("PHOTO").Value
            sb.AppendLine("BEGIN:VCARD" & vbNewLine & "VERSION:2.1")
            If dafname <> "" Then sb.AppendLine("FN;LANGUAGE=en-us;CHARSET=windows-1256:" & dafname)
            If damob1 <> "" Then sb.AppendLine("TEL;VOICE:" & damob1)
            If damob2 <> "" Then sb.AppendLine("TEL;VOICE:" & damob2)
            If damob3 <> "" Then sb.AppendLine("TEL;VOICE:" & damob3)
            If damob4 <> "" Then sb.AppendLine("TEL;VOICE:" & damob4)
            If damob5 <> "" Then sb.AppendLine("TEL;VOICE:" & damob5)
            If daaddress <> "" Then sb.AppendLine("Address;LANGUAGE=en-us;CHARSET=windows-1256:" & daaddress)
            If danotes <> "" Then sb.AppendLine("Note;LANGUAGE=en-us;CHARSET=windows-1256:" & danotes)
            If daorg <> "" Then sb.AppendLine("Org;LANGUAGE=en-us;CHARSET=windows-1256:" & daorg)
            If datitle <> "" Then sb.AppendLine("Title;LANGUAGE=en-us;CHARSET=windows-1256:" & datitle)
            If daemail <> "" Then sb.AppendLine("Email;LANGUAGE=en-us;CHARSET=windows-1256:" & daemail)
            If DAPHOTO <> "" Then sb.AppendLine("FN;LANGUAGE=en-us;CHARSET=windows-1256:" & DAPHOTO)
            sb.AppendLine("URL:http://Notes-ar.blogspot.com")
            sb.AppendLine("END:VCARD")
            irow = irow + 1
        Next
        File.WriteAllText(filePath, sb.ToString)
        'Process.Start(filePath)
    End Sub

    Private Sub Button6_Click_1(sender As Object, e As EventArgs) Handles Button6.Click
        Dim sFileArray As Object
        Dim fd As SaveFileDialog = New SaveFileDialog()
        fd.Title = "Export Vcf File"
        fd.InitialDirectory = "D:1.vcf"
        fd.Filter = "Bussiness Card FIles  (*.Vcf*)|*.Vcf*"
        fd.FileName = "Mycontacts - -" & Format(Now, "yyyy-mm-dd_HH-mm-ss")

        fd.FilterIndex = 9
        fd.RestoreDirectory = True
        If fd.ShowDialog() = DialogResult.OK Then sFileArray = fd.FileName Else Exit Sub
        Call exporttovcf(sFileArray)
        MsgBox("تم التصدير بنجاح")
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs)
        Call autoserial()
    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles DataGridView1.CellMouseClick
        '  DataGridView1.CurrentCell.KeyEntersEditMode(sender)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Call SearchSub()
    End Sub
    Sub SearchSub()
        If TextBox1.Text = "" Then Exit Sub
        Dim irow As Integer

        DataGridView1.Sort(DataGridView1.Columns("serial"), ListSortDirection.Ascending)
        For irow = 1 To DataGridView1.Rows.Count - 1
            Dim Searchstr As String = TextBox1.Text
            '  Dim orgstr As String = DataGridView1.Rows(irow).Cells("fname").Value.ToString

            Try

                DataGridView1.Rows(irow).Cells("fname").Selected = False
                ' MsgBox(InStr(DataGridView1.Rows(irow).Cells("fname").Value.ToString, Searchstr))
                If InStr(DataGridView1.Rows(irow).Cells("fname").Value.ToString, Searchstr) > 0 Or InStr(DataGridView1.Rows(irow).Cells("phone").Value.ToString, Searchstr) > 0 Then

                    'ComboBox1.Items.Add(DataGridView1.Rows(irow).Cells("fname").Value.ToString)
                    DataGridView1.Rows(irow).Cells("fname").Selected = True
                    DataGridView1.Rows(irow).Cells("SSort").Value = 1
                    DataGridView1.Sort(DataGridView1.Columns("SSort"), ListSortDirection.Descending)

                    'DataGridView1.Rows(irow).Cells("SSort").Value = ""

                End If
            Catch ex As Exception
                '  MsgBox(ex.Message)

            End Try
        Next


        For irow = 1 To DataGridView1.Rows.Count - 1
            DataGridView1.Rows(irow).Cells("SSort").Value = Nothing

        Next
        DataGridView1.FirstDisplayedScrollingRowIndex = 0

    End Sub






    ' This is a class you can use to set properties that you may otherwise send a parameters to your sub that does the data fetching. 
    Public Class Worker
        Public Sub GetData()

            'Do the Work Here 
            Dim filename As String = "D:my dbs\1.vcf"
            Dim TextLine As String = ""
            If System.IO.File.Exists(filename) = True Then
                Dim objReader As New System.IO.StreamReader(filename)
                Try
                    Do While objReader.Peek() <> -1
                        TextLine = objReader.ReadLine()
                        Dim ConstactsSum As Integer = InStr(TextLine, "FN")
                        '  form2.DataGridView2.Rows.Add(TextLine)
                    Loop
                    '  DataGridView2.Rows.Add(1)
                Catch
                End Try
            End If
            MsgBox("تم الاستيراد الجزء الاول")




        End Sub
    End Class

    'Here the simplest example method I can come up with 
    Private Sub Main()

        'Create the worker objects 
        Dim worker1 As New Worker
        Dim worker2 As New Worker

        'Create the threads pointing to the worker GetData() 
        Dim thread1 As New Thread(AddressOf worker1.GetData)
        Dim thread2 As New Thread(AddressOf worker2.GetData)


        'Start the threads 
        thread1.Start()
        thread2.Start()

        'Call Join on the threads to wait for them to be done. 
        thread1.Join()
        thread2.Join()
    End Sub
    Sub Mysaa()

        'If ProgressBar1.Value <= ProgressBar1.Maximum - 1 Then
        '    ProgressBar1.Value += 1

        ''End If
        'ProgressBar1.Value = 0
        'Timer1.Enabled = True
        'MsgBox("d")
        TT.Show()

    End Sub

    Private Sub TextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            MessageBox.Show("Enter key pressed")
            SearchSub()
        End If
    End Sub
    Private Sub BackgroundWorker1_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        '   Call Filterd()


    End Sub



    Private Sub Button2_Click_1(sender As Object, e As EventArgs) Handles Button2.Click
        SaveAllImages()
    End Sub
    Public Function DecodeBase64All(ByVal strData As String, filename As String) As Byte()
        Dim objXML As Object
        Dim objNode As Object
        objXML = CreateObject("MSXML2.DOMDocument")
        objNode = objXML.createElement("b64")
        objNode.DataType = "bin.base64"
        objNode.Text = strData
        DecodeBase64All = objNode.nodeTypedValue
        Dim imagespath As String = Application.StartupPath() & "\Images_" & Format(Now, "yyyy-MM-dd") & "\"
        System.IO.Directory.CreateDirectory(imagespath)
        System.IO.File.WriteAllBytes(imagespath & filename & ".jpg", DecodeBase64All)
        objNode = Nothing
        objXML = Nothing
    End Function

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Form3.Show()
    End Sub
End Class
