Option Explicit On
Imports System.ComponentModel
Imports System.Data.OleDb
Imports System.Linq
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Public Class Form1

    ReadOnly app As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application

    ReadOnly wApp As Microsoft.Office.Interop.Word.Application = New Microsoft.Office.Interop.Word.Application

    Private ReadOnly wf = app.WorksheetFunction
    Function FlipNames(MyString As Object)
        '
        '
        '
        Dim lastname As String

        Dim firstname As String

        Dim name As String


        If InStr(1, MyString, ", ") Then


            lastname = Microsoft.VisualBasic.Left(MyString, wf.Find(", ", MyString) - 1)

            firstname = Microsoft.VisualBasic.Right(MyString, Len(MyString) - wf.Find(", ", MyString) - 1)

            name = firstname & " " & lastname

            FlipNames = Trim(name)


        Else


            firstname = Microsoft.VisualBasic.Left(MyString, wf.Find(" ", MyString))

            lastname = Microsoft.VisualBasic.Right(MyString, Len(MyString) - wf.Find(" ", MyString))

            name = lastname & ", " & firstname

            FlipNames = Trim(name)


        End If


    End Function
    Private Function ReadExcelFile(sheetname As String, path As String) As Data.DataTable


        Using conn As New OleDb.OleDbConnection()

            Dim dt As New Data.DataTable

            Dim Import_FileName As String = path

            Dim fileExtension As String = IO.Path.GetExtension(Import_FileName)

            conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'"

            Using comm As New OleDbCommand()

                comm.CommandText = "Select * from [" + sheetname + "$]"

                comm.Connection = conn

                Using da As New OleDbDataAdapter()

                    da.SelectCommand = comm

                    da.Fill(dt)

                    Return dt

                End Using
            End Using
        End Using
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '
        '
        '
        Dim errln As Integer = 0
        Dim y As Integer = 0
        Dim jc As String = Nothing
        Dim store_name, fromDate, thruDate As String
        Dim objWord, objDoc, objExport As Document
        Dim ww As Integer = 0
        Dim xx As Integer = 0
        Dim yy As Integer = 0
        Dim zz As Integer = 0
        Dim daystotal As Integer = 0
        Dim progPcnt As Double = 0
        Dim i As Integer = 0
        Dim ii As Integer = 0
        Dim iii As Integer = 0
        Dim iiii As Integer = 0
        Dim iiiii As Integer = 0
        Dim parm, pst As Worksheet
        Dim d_jcode() As String = Nothing
        Dim n_jcode() As String = Nothing
        Dim d_station() As String = Nothing
        Dim n_station() As String = Nothing
        Dim d_start() As String = Nothing
        Dim n_start() As String = Nothing
        Dim d_end() As String = Nothing
        Dim n_end() As String = Nothing
        Dim xDoc As String = Nothing
        Dim iDoc As String = Nothing
        Dim oDoc As String = Nothing
        Dim d_emp() As String = Nothing
        Dim n_emp() As String = Nothing
        Dim thisWb As Workbook = Nothing
        Dim copyWb As Workbook = Nothing
        Dim r As Integer = 0
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim date_row() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Dim shift_row() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Dim unit, businessDate, sales, shift, jobtitles, employees, startTime, endTime, station As Integer
        Dim countAM As Integer = 0
        Dim countPM As Integer = 0
        Dim max_days As Integer = 0

        Dim watch As New Stopwatch

        watch.Start()


        'On Error GoTo errors


        Me.UseWaitCursor = True

        Me.Cursor = Cursors.WaitCursor

        Button1.Text = "Downloading..."



        Dim sPath As String = My.Computer.FileSystem.GetTempFileName

        My.Computer.FileSystem.WriteAllBytes(sPath, My.Resources.CRIB_SHEET_API, False)

        thisWb = app.Workbooks.Open(sPath)

        pst = thisWb.Worksheets("DATA")

        parm = thisWb.Worksheets("PARAMETERS")



        app.Visible = False

        app.DisplayAlerts = False

        app.ScreenUpdating = False

        app.EnableEvents = False


        wApp.ScreenUpdating = False

        wApp.Options.CheckGrammarAsYouType = False

        wApp.Options.CheckGrammarWithSpelling = False

        wApp.Options.CheckSpellingAsYouType = False

        wApp.Options.AnimateScreenMovements = False

        wApp.Options.BackgroundSave = False

        wApp.Options.CheckHangulEndings = False

        wApp.Options.DisableFeaturesbyDefault = True

        wApp.Visible = False



        'Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"

        'y = pst.Cells(pst.Rows.Count, 1).End(XlDirection.xlUp).Row

        parm.Range("name_FromDate").Value = Me.DateTimePicker1.Value

        parm.Range("name_ThruDate").Value = Me.DateTimePicker2.Value

        parm.Range("STORENAME").Value = Me.ComboBox1.Text

        thisWb.RefreshAll()

        thisWb.Save()

        thisWb.Close()

        app.Quit()



        Dim dt = ReadExcelFile("DATA", sPath)

        unit = dt.Columns(0).Ordinal

        businessDate = dt.Columns(1).Ordinal

        sales = dt.Columns(2).Ordinal

        shift = dt.Columns(3).Ordinal

        jobtitles = dt.Columns(4).Ordinal

        employees = dt.Columns(5).Ordinal

        startTime = dt.Columns(6).Ordinal

        endTime = dt.Columns(7).Ordinal

        station = dt.Columns(8).Ordinal


        Dim pt = ReadExcelFile("PARAMETERS", sPath)

        store_name = pt.Rows(0).ItemArray(3).ToString

        daystotal = pt.Rows(0).ItemArray(4).ToString

        Dim dayStart As Date = Me.DateTimePicker1.Value

        Dim dayLast As Date = Me.DateTimePicker2.Value

        y = dt.Rows.Count - 1


        watch.Stop()

        Dim timerresult = Watch.ElapsedMilliseconds / 1000

        'MsgBox(y & " " & store_name & " " & daystotal & " " & timerresult)



        '        On Error Resume Next
        '       For Each cell As Microsoft.Office.Interop.Excel.Range In pst.Range(pst.Cells(2, 4), pst.Cells(y, 4))
        '
        '     jobcodes.Add(cell.Value, cell.Value)
        '
        ' Next
        'On Error GoTo errors




        'Me.Close()

        'GoTo exitsub


        progPcnt = 6

        Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"



        sPath = My.Computer.FileSystem.GetTempFileName()

        My.Computer.FileSystem.WriteAllBytes(sPath, My.Resources.template, False)

        objWord = wApp.Documents.Open(sPath)

        objWord.Application.Visible = False


        progPcnt = 8

        Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"


        iDoc = My.Computer.FileSystem.GetTempFileName()

        objWord.SaveAs(iDoc, WdSaveFormat.wdFormatDocumentDefault)

        xDoc = My.Computer.FileSystem.GetTempFileName()

        objWord.SaveAs(xDoc, WdSaveFormat.wdFormatDocumentDefault)

        oDoc = My.Computer.FileSystem.GetTempFileName()

        objWord.SaveAs(oDoc, WdSaveFormat.wdFormatDocumentDefault)

        objWord.Close(SaveChanges:=False)


        objExport = wApp.Documents.Open(xDoc)

        objExport.Application.Visible = False


        progPcnt = 9

        Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"


        '''''''

        'iii = date

        'i is looping each job code for each shift

        'xx is looping AM job codes that have at least one person working

        'yy is looping PM job codes that have at least one person working

        'ii is looping each row in the paste tab

        '''''''



        ' populate an array containing row numbers which indicate the starting row number for each day in the report
        ' where i is the value for the day index in the report
        '

        ' populate an array containing row numbers which indicate the starting row number for each PM shift in the report
        ' where ii is the value for the day index in the report

        max_days = daystotal - 1
        i = 0 ' iterate dates
        ii = 0 ' iterate AM/PM splits
        iii = 0
        iiii = 0
        Dim max_row As Integer = 0
        Dim e_name As String = Nothing
        Dim next_e_name As String = Nothing
        Dim row_num As Integer = 1
        Dim next_row As Integer = 0
        date_row(0) = 0
        ReDim d_jcode(y), n_jcode(y), d_emp(y), n_emp(y), d_station(y), n_station(y), d_start(y), n_start(y), d_end(y), n_end(y)

        For row_num = 1 To y

            next_row = row_num + 1

            If next_row <= y Then

                If Convert.ToDateTime(dt.Rows(next_row).Item(businessDate)) > Convert.ToDateTime(dt.Rows(row_num).Item(businessDate)) Then

                    date_row(i) = next_row

                    i += 1



                End If

                If dt.Rows(next_row).Item(shift).ToString = "PM" And dt.Rows(row_num).Item(shift).ToString = "AM" Then

                    shift_row(ii) = next_row

                    ii += 1

                End If

            End If

        Next
        '
        '
        '

        'max_row = date_row(iii + 1)

        'max_row -= 1


        errln = 372



        For iii = 0 To max_days ' for each day of the week

            max_row = date_row(iii + 1) - 1

            i = 0

            ii = 0

            iiiii = 1

            Dim jobcodes As Dictionary(Of String, Integer) = New Dictionary(Of String, Integer)


            If iii.Equals(max_days) Then

                max_row = y

            End If



            r = date_row(iii)

            For row_num = Int(date_row(iii)) To max_row

                If r > max_row Then

                    Continue For

                Else

                    ' for each row in AM shift

                    If r < shift_row(iii) Then
                        'populate AM job codes list
                        jc = dt.Rows(r).Item(jobtitles).ToString
                        On Error Resume Next
                        If Not jobcodes.TryGetValue("PM" & jc, iiiii) Then
                            jobcodes("AM" & jc) = iiiii
                            iiiii += 1
                            On Error GoTo 0
                        End If
                        '
                        '
                        'populate AM shift info
                        d_jcode(i) = jc

                        d_emp(i) = dt.Rows(r).Item(employees).ToString

                        d_station(i) = dt.Rows(r).Item(station).ToString

                        d_start(i) = Convert.ToDateTime(dt.Rows(r).Item(startTime)).ToString("h:mmtt")

                        d_end(i) = Convert.ToDateTime(dt.Rows(r).Item(endTime)).ToString("h:mmtt")


                        If dt.Rows(r).Item(employees).ToString = dt.Rows(r + 1).Item(employees).ToString And dt.Rows(r).Item(endTime).ToString = dt.Rows(r + 1).Item(startTime).ToString Then

                            d_end(i) = Convert.ToDateTime(dt.Rows(r + 1).Item(endTime)).ToString("h:mmtt")

                            d_station(i) = d_station(i) & "/" & dt.Rows(r + 1).Item(station).ToString

                            r += 1

                        End If

                        r += 1

                        i += 1

                    End If
                    errln = 435



                    If r >= shift_row(iii) Then

                        'populate PM job codes list
                        jc = dt.Rows(r).Item(jobtitles).ToString

                        On Error Resume Next
                        If Not jobcodes.TryGetValue("PM" & jc, iiiii) Then
                            jobcodes("PM" & jc) = iiiii
                            iiiii += 1
                            On Error GoTo 0
                        End If
                        ' 
                        '
                        'populate PM shift info
                        n_jcode(ii) = jc

                        n_emp(ii) = dt.Rows(r).Item(employees).ToString

                        n_station(ii) = dt.Rows(r).Item(station).ToString

                        n_start(ii) = Convert.ToDateTime(dt.Rows(r).Item(startTime)).ToString("h:mmtt")

                        n_end(ii) = Convert.ToDateTime(dt.Rows(r).Item(endTime)).ToString("h:mmtt")


                        If r < y Then


                            If dt.Rows(r).Item(employees).ToString = dt.Rows(r + 1).Item(employees).ToString And dt.Rows(r).Item(endTime).ToString = dt.Rows(r + 1).Item(startTime).ToString Then

                                n_end(ii) = Convert.ToDateTime(dt.Rows(r + 1).Item(endTime)).ToString("h:mmtt")

                                n_station(ii) = n_station(ii) & "/" & dt.Rows(r + 1).Item(station).ToString

                                r += 1

                            End If


                        End If


                        r += 1

                        ii += 1


                    End If

                    errln = 474

                End If

            Next




            objDoc = wApp.Documents.Open(iDoc)


            objDoc.Application.Visible = False


            With objDoc
                errln = 488
                .Application.Selection.Find.Text = "<<STORE_NAME>>"
                .Application.Selection.Find.Execute()
                .Application.Selection.Text = store_name
                .Application.Selection.EndOf()

                .Application.Selection.Find.Text = "<<DAY>>"
                .Application.Selection.Find.Execute()
                .Application.Selection.Text = DateAdd("d", Convert.ToDouble(iii), dayStart).DayOfWeek.ToString
                .Application.Selection.EndOf()

                .Application.Selection.Find.Text = "<<DATE>>"
                .Application.Selection.Find.Execute()
                .Application.Selection.Text = DateAdd("d", Convert.ToDouble(iii), dayStart).ToShortDateString
                .Application.Selection.EndOf()

                .Application.Selection.Find.Text = "<<SALES_FORECAST>>"
                .Application.Selection.Find.Execute()
                .Application.Selection.Text = Format(dt.Rows(date_row(iii)).Item(sales), "0,000")
                .Application.Selection.EndOf()
                errln = 508

                countAM = 0

                countPM = 0


                For Each kvp As KeyValuePair(Of String, Integer) In jobcodes

                    If Microsoft.VisualBasic.Left(kvp.Key, 2) = "AM" Then

                        countAM += 1
                    Else

                        countPM += 1

                    End If

                Next
                On Error GoTo errors
                errln = 527

                Dim emp As String = ""


                For iiii = 1 To countAM


                    If iiii <= countAM Then


                        emp = Nothing

                        iiiii = 0

                        For r = 0 To i


                            If d_jcode(r) = Microsoft.VisualBasic.Right(jobcodes.Keys(iiii).ToString, Len(jobcodes.Keys(iiii).ToString) - 2) Then

                                emp = emp & FlipNames(d_emp(r)) & " - " & d_start(r).ToString & "-" & d_end(r).ToString & " (" & d_station(r) & ")" & vbCrLf

                                iiiii += 1

                            End If


                        Next


                        .Application.Selection.Find.Execute(FindText:="<<D_JOB_CODE_" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                        .Application.Selection.Text = Microsoft.VisualBasic.Right(jobcodes.Keys(iiii).ToString, Len(jobcodes.Keys(iiii).ToString) - 2) & " (" & iiiii & ")"

                        .Application.Selection.EndOf(WdUnits.wdLine)


                        .Application.Selection.Find.Execute(FindText:="<<D_EMP" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                        .Application.Selection.Text = emp

                        .Application.Selection.EndOf(WdUnits.wdLine)


                    End If
                Next

                errln = 582

                For iiii = countAM To countPM

                    If iiii <= countPM Then


                        emp = Nothing

                        iiiii = 0

                        For r = 0 To i

                            If n_jcode(r) = Microsoft.VisualBasic.Right(jobcodes.Keys(iiii).ToString, Len(jobcodes.Keys(iiii).ToString) - 2) Then

                                emp = emp & FlipNames(n_emp(r)) & " - " & n_start(r).ToString & "-" & n_end(r).ToString & " (" & n_station(r) & ")" & vbCrLf

                                iiiii += 1

                            End If

                        Next

                        .Application.Selection.Find.Execute(FindText:="<<N_JOB_CODE_" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                        .Application.Selection.Text = Microsoft.VisualBasic.Right(jobcodes.Keys(iiii).ToString, Len(jobcodes.Keys(iiii).ToString) - 2) & " (" & iiiii & ")"

                        .Application.Selection.EndOf(WdUnits.wdLine)


                        .Application.Selection.Find.Execute(FindText:="<<N_EMP" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                        .Application.Selection.Text = emp

                        .Application.Selection.EndOf(WdUnits.wdLine)




                    End If
                    .Application.Selection.Find.Execute(FindText:="<<D_JOB_CODE_" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                    .Application.Selection.Delete()

                    .Application.Selection.EndOf(WdUnits.wdLine)


                    .Application.Selection.Find.Execute(FindText:="<<D_EMP" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                    .Application.Selection.Delete()

                    .Application.Selection.EndOf(WdUnits.wdLine)

                    .Application.Selection.Find.Execute(FindText:="<<N_JOB_CODE_" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                    .Application.Selection.Delete()

                    .Application.Selection.EndOf(WdUnits.wdLine)


                    .Application.Selection.Find.Execute(FindText:="<<N_EMP" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                    .Application.Selection.Delete()

                    .Application.Selection.EndOf(WdUnits.wdLine)

                    errln = 631
                Next

                .SaveAs(oDoc, WdSaveFormat.wdFormatDocumentDefault)

                .Close(SaveChanges:=False)


            End With

            errln = 642

            With objExport


                If iii = 0 Then

                    .Content.Select()

                End If


                .Application.Selection.InsertFile(oDoc)

                .Application.Selection.EndKey(WdUnits.wdStory)

                My.Computer.FileSystem.DeleteFile(oDoc)


                If iii < daystotal Then


                    .Application.Selection.InsertBreak(WdBreakType.wdPageBreak)

                    .Application.Selection.EndKey(WdUnits.wdStory)


                End If


            End With

            errln = 673

            progPcnt += ((30) / daystotal)

            Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"


        Next iii

        progPcnt = 99

        Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"


        objExport.ExportAsFixedFormat(IO.Directory.GetCurrentDirectory, WdExportFormat.wdExportFormatPDF, OpenAfterExport:=True)

        objExport.Close(False)

        My.Computer.FileSystem.DeleteFile(xDoc)

        My.Computer.FileSystem.DeleteFile(iDoc)

        'Me.Activate()


        Stop
exitsub:

        Button1.Text = "GENERATE CRIB SHEET"

        Me.UseWaitCursor = False

        Me.Cursor = Cursors.Default


        wApp.ScreenUpdating = True

        wApp.Options.CheckGrammarAsYouType = True

        wApp.Options.CheckGrammarWithSpelling = True

        wApp.Options.CheckSpellingAsYouType = True

        wApp.Options.AnimateScreenMovements = True

        wApp.Options.BackgroundSave = True

        wApp.Options.CheckHangulEndings = True

        wApp.Options.DisableFeaturesbyDefault = False

        Exit Sub



errors:

        'On Error Resume Next

        Me.UseWaitCursor = False

        Me.Cursor = Cursors.Default

        Button1.Text = "GENERATE CRIB SHEET"



        wApp.ScreenUpdating = True

        wApp.Options.CheckGrammarAsYouType = True

        wApp.Options.CheckGrammarWithSpelling = True

        wApp.Options.CheckSpellingAsYouType = True

        wApp.Options.AnimateScreenMovements = True

        wApp.Options.BackgroundSave = True

        wApp.Options.CheckHangulEndings = True

        wApp.Options.DisableFeaturesbyDefault = False


        MsgBox("Yikes dude, looks like something went hecka wrong. Please email mitch@rocksolidrestaurants.com right away with any details, and make sure to attach this excel file in your message. " & Err.Description & " Line #" & errln & " Err #" & Err.Number)

    End Sub
    Private Sub Form1_Closed(sender As Object, e As EventArgs) Handles Me.Closing

        wApp.Quit(WdSaveOptions.wdDoNotSaveChanges)

        Marshal.FinalReleaseComObject(wApp)

        Marshal.FinalReleaseComObject(app)

        GC.Collect()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.DateTimePicker1.Value = DateAdd("d", 1, Today())

        Me.DateTimePicker2.Value = DateAdd("d", 1, Today())

        Me.ComboBox1.Text = "Bonney Lake"

    End Sub

    Private Sub ComboBox1_SelectedValueChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedValueChanged

        If Me.ComboBox1.Text <> "" Then


            Me.Button1.Enabled = True

            Me.Button1.Visible = True


        Else


            Me.Button1.Enabled = False

            Me.Button1.Visible = False


        End If

    End Sub
End Class
