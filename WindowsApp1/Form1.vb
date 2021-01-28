Option Explicit On
Imports Newtonsoft.Json
Imports RestSharp
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Public Class Form1

    Dim app As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application

    Dim wApp As Microsoft.Office.Interop.Word.Application = New Microsoft.Office.Interop.Word.Application



    'Private ReadOnly wf = app.WorksheetFunction
    Function GetRequest(unitID As String, FromDate As String, ThruDate As String)

        Dim dat As Data.DataTable

        Dim client As New RestClient("https://dc01.rmdatacentral.com/Portal1242/api/api/v2/Reports/b05c3f07-ecd7-455a-a460-e96d05376ac5/Result?@UnitID=" & unitID & "&@FromDate=" & Convert.ToDateTime(FromDate).ToString("MM-dd-yyyy") & "&@ThruDate=" & Convert.ToDateTime(ThruDate).ToString("MM-dd-yyyy"))

        client.Timeout = -1

        Dim request As RestRequest = New RestRequest(Method.GET)

        request.AddHeader("Accept", "application/json")

        request.AddHeader("DCKey", "v3q8upr9eNm0TsQHJz3AcXnG1lhCpNSJW8EnXNSGR343-yJcKB8eUKYJeLLwr5YeZ50")

        Dim response As IRestResponse = client.Execute(request)

        Dim _ds As DataSet

        _ds = JsonConvert.DeserializeObject(Of DataSet)(response.Content)

        Dim _dt As Data.DataTable = _ds.Tables.Item(0)

        Dim __dt As Data.DataSet = _dt.Rows(5).Table.DataSet

        Dim ___dt As Data.DataTable = __dt.Tables.Item(0)

        dat = ___dt.Rows(5).Item(2)

        'Dim _dat As New Data.DataTable

        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("UnitName")))
        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("BusinessDate")))
        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("Forecast")))
        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("Shift")))
        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("JobName")))
        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("EmployeeName")))
        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("StartTime")))
        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("EndTime")))
        '_dat.Columns.Add(dat.Columns.Item(dat.Columns.IndexOf("StationName")))


        Return dat

    End Function
    Function FlipNames(MyString As Object)
        '
        '
        '
        Dim lastname As String

        Dim firstname As String

        Dim name As String


        'If InStr(1, MyString, ", ") Then


        lastname = Microsoft.VisualBasic.Left(MyString, InStr(MyString, ",") - 1)

        firstname = Microsoft.VisualBasic.Right(MyString, Len(MyString) - InStr(MyString, ",") - 1)

        name = firstname & " " & lastname

        FlipNames = Trim(name)


        'Else


        'firstname = Microsoft.VisualBasic.Left(MyString, InStr(" ", MyString))

        '    lastname = Microsoft.VisualBasic.Right(MyString, Len(MyString) - InStr(" ", MyString))

        '    name = lastname & ", " & firstname

        '    FlipNames = Trim(name)


        ' End If

        Return FlipNames

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


        On Error GoTo errors


        Me.UseWaitCursor = True

        Me.Cursor = Cursors.WaitCursor

        Button1.Text = "Downloading..."



        Dim sPath As String = My.Computer.FileSystem.GetTempFileName

        My.Computer.FileSystem.WriteAllBytes(sPath, My.Resources.CRIB_SHEET_API, False)

        thisWb = app.Workbooks.Open(sPath)

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



        parm.Range("name_FromDate").Value = Me.DateTimePicker1.Value

        parm.Range("name_ThruDate").Value = Me.DateTimePicker2.Value

        parm.Range("STORENAME").Value = Me.ComboBox1.Text

        daystotal = parm.Range("DaysTotal").Value


        Dim dtab As Data.DataTable = GetRequest(parm.Range("name_UnitID").Value, Me.DateTimePicker1.Value, Me.DateTimePicker2.Value)


        thisWb.Close(SaveChanges:=False)

        app.DisplayAlerts = True

        app.ScreenUpdating = True

        app.EnableEvents = True

        app.Quit()

        Dim dv As New DataView(dtab)

        dv.Sort = "BusinessDate ASC, Shift ASC, JobName ASC, EmployeeName ASC, StartTime ASC"

        Dim dt As Data.DataTable = dv.ToTable

        'DataGridView1.DataSource = dt

        unit = dt.Columns.IndexOf("UnitName")

        businessDate = dt.Columns.IndexOf("businessDate")

        sales = dt.Columns.IndexOf("Forecast")

        shift = dt.Columns.IndexOf("Shift")

        jobtitles = dt.Columns.IndexOf("JobName")

        employees = dt.Columns.IndexOf("EmployeeName")

        startTime = dt.Columns.IndexOf("StartTime")

        endTime = dt.Columns.IndexOf("EndTime")

        station = dt.Columns.IndexOf("StationName")


        store_name = Me.ComboBox1.Text



        Dim dayStart As Date = Me.DateTimePicker1.Value

        Dim dayLast As Date = Me.DateTimePicker2.Value

        y = dt.Rows.Count - 1


        progPcnt = 6

        Button1.Text = $"Working... {Format(progPcnt, "0")}%"



        sPath = My.Computer.FileSystem.GetTempFileName()

        My.Computer.FileSystem.WriteAllBytes(sPath, My.Resources.template, False)

        objWord = wApp.Documents.Open(sPath)

        objWord.Application.Visible = False


        progPcnt = 8

        Button1.Text = $"Working... {Format(progPcnt, "0")}%"


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

        Button1.Text = $"Working... {Format(progPcnt, "0")}%"


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
        i = 1 ' iterate dates
        ii = 0 ' iterate AM/PM splits
        iii = 0
        iiii = 0
        Dim max_row As Integer = 0
        Dim e_name As String = Nothing
        Dim next_e_name As String = Nothing
        Dim row_num As Integer = 1
        Dim next_row As Integer = 0
        date_row(0) = 0


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



        Dim jobcodesAM As Dictionary(Of String, Integer)

        Dim jobcodesPM As Dictionary(Of String, Integer)


        For iii = 0 To max_days ' for each day of the week

            ReDim d_jcode(y), n_jcode(y), d_emp(y), n_emp(y), d_station(y), n_station(y), d_start(y), n_start(y), d_end(y), n_end(y)

            max_row = date_row(iii + 1) - 1

            i = 0

            ii = 0

            iiiii = 1

            jobcodesAM = New Dictionary(Of String, Integer)

            jobcodesPM = New Dictionary(Of String, Integer)


            If iii.Equals(max_days) Then

                max_row = y

            End If



            r = date_row(iii)

            progPcnt += ((30) / daystotal)

            Button1.Text = $"Working... {Format(progPcnt, "0")}%"

            For row_num = date_row(iii) To max_row

                If r > max_row Then

                    Continue For

                Else

                    ' for each row in AM shift

                    If r < shift_row(iii) Then
                        'populate AM job codes list
                        jc = dt.Rows(r).Item(jobtitles).ToString
                        On Error Resume Next
                        If Not jobcodesAM.TryGetValue(jc, iiiii) Then
                            jobcodesAM(jc) = iiiii
                            iiiii += 1
                            On Error GoTo errors
                        End If
                        '
                        '
                        'populate AM shift info
                        d_jcode(i) = jc

                        d_emp(i) = dt.Rows(r).Item(employees).ToString

                        d_station(i) = dt.Rows(r).Item(station).ToString

                        d_start(i) = Convert.ToDateTime(dt.Rows(r).Item(startTime)).ToString("h:mmtt")

                        d_end(i) = Convert.ToDateTime(dt.Rows(r).Item(endTime)).ToString("h:mmtt")


                        If dt.Rows(r).Item(employees).ToString = dt.Rows(r + 1).Item(employees).ToString And dt.Rows(r).Item(endTime).ToString = dt.Rows(r + 1).Item(startTime).ToString And dt.Rows(r).Item(jobtitles).ToString = dt.Rows(r + 1).Item(jobtitles).ToString Then

                            Do Until Not (dt.Rows(r).Item(employees).ToString = dt.Rows(r + 1).Item(employees).ToString And dt.Rows(r).Item(endTime).ToString = dt.Rows(r + 1).Item(startTime).ToString And dt.Rows(r).Item(jobtitles).ToString = dt.Rows(r + 1).Item(jobtitles).ToString)

                                d_end(i) = Convert.ToDateTime(dt.Rows(r + 1).Item(endTime)).ToString("h:mmtt")


                                If dt.Rows(r + 1).Item(station).ToString <> "" Then

                                    If d_station(i) <> "" Then

                                        d_station(i) = d_station(i) & "/" & dt.Rows(r + 1).Item(station).ToString

                                    Else

                                        d_station(i) = dt.Rows(r + 1).Item(station).ToString

                                    End If

                                End If


                                r += 1

                            Loop

                        End If

                        r += 1

                        i += 1

                    End If



                    If r >= shift_row(iii) Then

                        'populate PM job codes list
                        jc = dt.Rows(r).Item(jobtitles).ToString

                        On Error Resume Next
                        If Not jobcodesPM.TryGetValue(jc, iiiii) Then
                            jobcodesPM(jc) = iiiii
                            iiiii += 1
                            On Error GoTo errors
                        End If
                        ' 
                        '
                        'populate PM shift info
                        n_jcode(ii) = jc

                        n_emp(ii) = dt.Rows(r).Item(employees).ToString

                        n_station(ii) = dt.Rows(r).Item(station).ToString

                        n_start(ii) = Convert.ToDateTime(dt.Rows(r).Item(startTime)).ToString("h:mmtt")

                        n_end(ii) = Convert.ToDateTime(dt.Rows(r).Item(endTime)).ToString("h:mmtt")

                        Do While r < y

                            If (dt.Rows(r).Item(employees).ToString = dt.Rows(r + 1).Item(employees).ToString And dt.Rows(r).Item(endTime).ToString = dt.Rows(r + 1).Item(startTime).ToString And dt.Rows(r).Item(jobtitles).ToString = dt.Rows(r + 1).Item(jobtitles).ToString) Then


                                n_end(ii) = Convert.ToDateTime(dt.Rows(r + 1).Item(endTime)).ToString("h:mmtt")


                                If dt.Rows(r + 1).Item(station).ToString <> "" Then


                                    If n_station(ii) <> "" Then


                                        n_station(ii) = n_station(ii) & "/" & dt.Rows(r + 1).Item(station).ToString

                                    Else

                                        n_station(ii) = dt.Rows(r + 1).Item(station).ToString


                                    End If


                                End If

                                r += 1

                            Else

                                Exit Do

                            End If


                        Loop





                        End If





                        r += 1

                        ii += 1


                    End If




            Next

            progPcnt += ((30) / daystotal)

            Button1.Text = $"Working... {Format(progPcnt, "0")}%"


            objDoc = wApp.Documents.Open(iDoc)


            objDoc.Application.Visible = False


            With objDoc

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
                .Application.Selection.Text = Format(dt.Rows(date_row(iii)).Item(sales), "N")
                .Application.Selection.EndOf()

                countAM = jobcodesAM.Count - 1

                countPM = jobcodesPM.Count - 1



                Dim emp As String = ""


                For iiii = 0 To 7


                    If iiii <= countAM Then


                        emp = ""

                        iiiii = 0


                        For r = 0 To i - 1


                            If d_jcode(r) = jobcodesAM.Keys(iiii).ToString Then

                                emp = emp & FlipNames(d_emp(r)) & " - " & d_start(r).ToString & "-" & d_end(r).ToString

                                If d_station(r) <> "" Then

                                    emp = emp & " (" & d_station(r) & ")" & vbCrLf

                                Else

                                    emp &= vbCrLf

                                End If

                                iiiii += 1

                            End If


                        Next


                        .Application.Selection.Find.Execute(FindText:="<<D_JOB_CODE_" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                        .Application.Selection.Text = jobcodesAM.Keys(iiii).ToString & " (" & iiiii & ")"

                        .Application.Selection.EndOf(WdUnits.wdLine)


                        .Application.Selection.Find.Execute(FindText:="<<D_EMP" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                        .Application.Selection.Text = emp

                        .Application.Selection.EndOf(WdUnits.wdLine)


                    End If



                    .Application.Selection.Find.Execute(FindText:="<<D_JOB_CODE_" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                    .Application.Selection.Text = vbCrLf

                    .Application.Selection.EndOf(WdUnits.wdLine)


                    .Application.Selection.Find.Execute(FindText:="<<D_EMP" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                    .Application.Selection.Delete()

                    .Application.Selection.EndOf(WdUnits.wdLine)

                Next


                For iiii = 0 To 7

                    If iiii <= countPM Then

                        emp = ""

                        iiiii = 0


                        For r = 0 To ii - 1

                            If n_jcode(r) = jobcodesPM.Keys(iiii).ToString Then

                                emp = emp & FlipNames(n_emp(r)) & " - " & n_start(r).ToString & "-" & n_end(r).ToString

                                If n_station(r) <> "" Then

                                    emp &= " (" & n_station(r) & ")" & vbCrLf

                                Else

                                    emp &= vbCrLf

                                End If

                                iiiii += 1

                            End If

                        Next


                        .Application.Selection.Find.Execute(FindText:="<<N_JOB_CODE_" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                        .Application.Selection.Text = jobcodesPM.Keys(iiii).ToString & " (" & iiiii & ")"

                        .Application.Selection.EndOf(WdUnits.wdLine)


                        .Application.Selection.Find.Execute(FindText:="<<N_EMP" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                        .Application.Selection.Text = emp

                        .Application.Selection.EndOf(WdUnits.wdLine)



                    End If


                    .Application.Selection.Find.Execute(FindText:="<<N_JOB_CODE_" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                    .Application.Selection.Text = vbCrLf

                    .Application.Selection.EndOf(WdUnits.wdLine)


                    .Application.Selection.Find.Execute(FindText:="<<N_EMP" & CStr(iiii) & ">>", Wrap:=WdFindWrap.wdFindContinue)

                    .Application.Selection.Delete()

                    .Application.Selection.EndOf(WdUnits.wdLine)


                Next

                .SaveAs(oDoc, WdSaveFormat.wdFormatDocumentDefault)

                .Close(SaveChanges:=False)


            End With

            With objExport


                If iii = 0 Then

                    .Content.Select()

                End If


                .Application.Selection.InsertFile(oDoc)

                .Application.Selection.EndKey(WdUnits.wdStory)

                My.Computer.FileSystem.DeleteFile(oDoc)


                If iii < max_days Then


                    .Application.Selection.InsertBreak(WdBreakType.wdPageBreak)

                    .Application.Selection.EndKey(WdUnits.wdStory)


                End If


            End With


            progPcnt += ((30) / daystotal)

            Button1.Text = $"Working... {Format(progPcnt, "0")}%"


        Next iii

        progPcnt = 99

        Button1.Text = $"Working... {Format(progPcnt, "0")}%"

        If daystotal = 1 Then

            objExport.ExportAsFixedFormat(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "/Crib_Sheet_" & Convert.ToDateTime(Me.DateTimePicker1.Value).ToString("M.dd") & ".pdf", WdExportFormat.wdExportFormatPDF, OpenAfterExport:=True)

        Else

            objExport.ExportAsFixedFormat(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) & "/Crib_Sheet_" & Convert.ToDateTime(Me.DateTimePicker1.Value).ToString("M.dd") & "_" & Convert.ToDateTime(Me.DateTimePicker2.Value).ToString("M.dd") & ".pdf", WdExportFormat.wdExportFormatPDF, OpenAfterExport:=True)

        End If

        objExport.Close(False)

        My.Computer.FileSystem.DeleteFile(xDoc)

        My.Computer.FileSystem.DeleteFile(iDoc)

        'Me.Activate()



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

        On Error Resume Next

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


        MsgBox("Yikes dude, looks like something went hecka wrong. Please email mitch@rocksolidrestaurants.com right away with any details, and make sure to attach this excel file in your message. " & Err.Description & " Line #" & Err.Erl & " Err #" & Err.Number)

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

        'Dim dtab As Data.DataTable = GetRequest(11, Me.DateTimePicker1.Value, Me.DateTimePicker2.Value)

        'dtab.DefaultView.Sort = "BusinessDate ASC, Shift ASC, EmployeeName ASC, StartTime ASC"

        'DataGridView1.DataSource = dtab.DefaultView.ToTable

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
