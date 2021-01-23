Option Explicit On
Imports System.ComponentModel
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Word
Public Class Form1

    Dim app As Microsoft.Office.Interop.Excel.Application = New Microsoft.Office.Interop.Excel.Application

    Dim wApp As Microsoft.Office.Interop.Word.Application = New Microsoft.Office.Interop.Word.Application

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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '
        '
        '

        Dim x, y As Long
        Dim pasteFileName As String
        Dim store_name, fromDate, thruDate As String
        Dim objWord, objDoc, objExport As Document
        Dim ww As Integer = 0
        Dim xx As Integer = 0
        Dim yy As Integer = 0
        Dim zz As Integer = 0
        Dim daystotal As Integer = 0
        Dim progPcnt As Double = 0
        Dim i As Integer = Nothing
        Dim ii As Integer = Nothing
        Dim iii As Integer = Nothing
        Dim parm, pst As Worksheet
        Dim d_jcode() As String = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}
        Dim n_jcode() As String = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}
        Dim xDoc As String = Nothing
        Dim iDoc As String = Nothing
        Dim oDoc As String = Nothing
        Dim d_emp() As String = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}
        Dim n_emp() As String = {"", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}
        Dim thisWb As Workbook = Nothing
        Dim copyWb As Workbook = Nothing
        Dim jobcodes As New Collection
        Dim r As Long
        Dim rng As Microsoft.Office.Interop.Excel.Range
        Dim date_row() As Integer = {0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0}
        Dim shift_row(,) As VariantType = {"", "";"",""}


        'On Error GoTo errors


        Me.UseWaitCursor = True

        Me.Cursor = Cursors.WaitCursor

        Button1.Text = "Working... 0%"



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



        Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"

        'x = pst.Cells(1, pst.Columns.Count).End(XlDirection.xlToLeft).Column

        y = pst.Cells(pst.Rows.Count, 1).End(XlDirection.xlUp).Row

        parm.Range("name_FromDate").Value = Me.DateTimePicker1.Value

        parm.Range("name_ThruDate").Value = Me.DateTimePicker2.Value

        parm.Range("STORENAME").Value = Me.ComboBox1.Text

        thisWb.RefreshAll()

        store_name = parm.Range("STORENAME").Value

        daystotal = parm.Range("DAYSTOTAL").Value

        Dim dayStart As Date = Me.DateTimePicker1.Value

        Dim dayLast As Date = Me.DateTimePicker2.Value


        '        On Error Resume Next
        '       For Each cell As Microsoft.Office.Interop.Excel.Range In pst.Range(pst.Cells(2, 4), pst.Cells(y, 4))
        '
        '     jobcodes.Add(cell.Value, cell.Value)
        '
        ' Next
        'On Error GoTo errors





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
        '''


        ' populate an array containing row numbers which indicate the starting row number for each day in the report
        '
        '
        i = 0

        date_row(0) = 2


        For Each cell As Microsoft.Office.Interop.Excel.Range In pst.Range(pst.Cells(2, 1), pst.Cells(y, 1))

            If cell.Value > pst.Cells(date_row(i - 1), 1).Value Then

                date_row(i) = cell.Row

                i += 1

            End If

        Next
        '
        '
        '


        For iii = 0 To (daystotal - 1)


            If ttls.Cells(2 + (iii * 9), 7).Value > 0 Then


                xx = 1

                For i = 1 To 6

                    If ttls.Cells((3 + i + (iii * 9)), 3).Value <> 0 Then

                        d_jcode(xx) = ttls.Cells(3 + i + (iii * 9), 2).Value & " (" & ttls.Cells(3 + i + (iii * 9), 3).Value & ")"

                        xx += 1

                    End If

                Next i



                yy = 1

                For i = 1 To 6

                    If ttls.Cells(3 + i + (iii * 9), 6).Value <> 0 Then

                        n_jcode(yy) = ttls.Cells(3 + i + (iii * 9), 5).Value & " (" & ttls.Cells(3 + i + (iii * 9), 6).Value & ")"

                        yy += 1

                    End If

                Next i


                '''''''''''
                '''''''''''
                '''''''''''
                progPcnt += (30 / daystotal)

                Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"

                zz = 0

                ww = 0

                For i = 1 To 6


                    If ttls.Cells(3 + i + (iii * 9), 3).Value <> 0 Then

                        zz += 1

                        d_emp(zz) = ""

                        For ii = 2 To pst.Cells(pst.Rows.Count, 1).End(XlDirection.xlUp).Row

                            If pst.Cells(ii, 4).Value = ttls.Cells(3 + i + (iii * 9), 2).Value _
                                And pst.Cells(ii, 2).Value = "AM" _
                                And pst.Cells(ii, 3).Value = ttls.Cells(2 + (iii * 9), 2).Value Then


                                d_emp(zz) = d_emp(zz) & FlipNames(pst.Cells(ii, 5).Value) & " - " _
                                    & wf.text(pst.Cells(ii, 6).Value, "[$-en-US]h:mmAM/PM;@") & "-" _
                                    & wf.text(pst.Cells(ii, 7).Value, "[$-en-US]h:mmAM/PM;@") & vbCrLf




                            End If


                        Next ii


                    End If

                    If ttls.Cells(3 + i + (iii * 9), 6).Value <> 0 Then

                        ww += 1

                        n_emp(ww) = ""

                        For ii = 2 To pst.Cells(pst.Rows.Count, 1).End(XlDirection.xlUp).Row

                            If pst.Cells(ii, 4).Value = ttls.Cells(3 + i + (iii * 9), 2).Value _
                                And pst.Cells(ii, 2).Value = "PM" _
                                And pst.Cells(ii, 3).Value = ttls.Cells(2 + (iii * 9), 2).Value Then


                                n_emp(ww) = n_emp(ww) & FlipNames(pst.Cells(ii, 5).Value) & " - " _
                                    & wf.text(pst.Cells(ii, 6).Value, "[$-en-US]h:mmAM/PM;@") & "-" _
                                    & wf.Text(pst.Cells(ii, 7).Value, "[$-en-US]h:mmAM/PM;@") & vbCrLf



                            End If


                        Next ii


                    End If

                    progPcnt += ((5) / daystotal)

                    Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"



                Next i

                '''''''''''
                '''''''''''
                '''''''''''




                objDoc = wApp.Documents.Open(iDoc)


                objDoc.Application.Visible = False


                With objDoc

                    .Application.Selection.Find.Text = "<<STORE_NAME>>"
                    .Application.Selection.Find.Execute()
                    .Application.Selection.Text = store_name
                    .Application.Selection.EndOf()

                    .Application.Selection.Find.Text = "<<DAY>>"
                    .Application.Selection.Find.Execute()
                    .Application.Selection.Text = ttls.Cells(3 + (iii * 9), 2).value
                    .Application.Selection.EndOf()

                    .Application.Selection.Find.Text = "<<DATE>>"
                    .Application.Selection.Find.Execute()
                    .Application.Selection.Text = ttls.Cells(2 + (iii * 9), 2).value
                    .Application.Selection.EndOf()

                    For i = 1 To 6

                        If i <= xx - 1 Then

                            '.Application.Selection.Find.Text = ("<<D_JOB_CODE_" & CStr(i) & ">>")
                            .Application.Selection.Find.Execute(FindText:="<<D_JOB_CODE_" & CStr(i) & ">>", Wrap:=WdFindWrap.wdFindContinue)
                            .Application.Selection.Text = d_jcode(i)
                            .Application.Selection.EndOf(WdUnits.wdLine)

                            '.Application.Selection.Find.Text = ("<<D_EMP" & CStr(i) & ">>")
                            .Application.Selection.Find.Execute(FindText:="<<D_EMP" & CStr(i) & ">>", Wrap:=WdFindWrap.wdFindContinue)
                            .Application.Selection.Text = d_emp(i)
                            .Application.Selection.EndOf(WdUnits.wdLine)

                        Else

                            '.Application.Selection.Find.Text = ("<<D_JOB_CODE_" & CStr(i) & ">>")
                            .Application.Selection.Find.Execute(FindText:="<<D_JOB_CODE_" & CStr(i) & ">>", Wrap:=WdFindWrap.wdFindContinue)
                            .Application.Selection.Delete()
                            .Application.Selection.EndOf(WdUnits.wdLine)

                            '.Application.Selection.Find.Text = ("<<D_EMP" & CStr(i) & ">>")
                            .Application.Selection.Find.Execute(FindText:="<<D_EMP" & CStr(i) & ">>", Wrap:=WdFindWrap.wdFindContinue)
                            .Application.Selection.Delete()
                            .Application.Selection.EndOf(WdUnits.wdLine)


                        End If

                        If i <= yy - 1 Then

                            '.Application.Selection.Find.Text = ("<<N_JOB_CODE_" & CStr(i) & ">>")
                            .Application.Selection.Find.Execute(FindText:="<<N_JOB_CODE_" & CStr(i) & ">>", Wrap:=WdFindWrap.wdFindContinue)
                            .Application.Selection.Text = n_jcode(i)
                            .Application.Selection.EndOf(WdUnits.wdLine)


                            '.Application.Selection.Find.Text = ("<<N_EMP" & CStr(i) & ">>")
                            .Application.Selection.Find.Execute(FindText:="<<N_EMP" & CStr(i) & ">>", Wrap:=WdFindWrap.wdFindContinue)
                            .Application.Selection.Text = n_emp(i)
                            .Application.Selection.EndOf(WdUnits.wdLine)


                        Else

                            '.Application.Selection.Find.Text = ("<<N_JOB_CODE_" & CStr(i) & ">>")
                            .Application.Selection.Find.Execute(FindText:="<<N_JOB_CODE_" & CStr(i) & ">>", Wrap:=WdFindWrap.wdFindContinue)
                            .Application.Selection.Delete()
                            .Application.Selection.EndOf(WdUnits.wdLine)


                            '.Application.Selection.Find.Text = ("<<N_EMP" & CStr(i) & ">>")
                            .Application.Selection.Find.Execute(FindText:="<<N_EMP" & CStr(i) & ">>", Wrap:=WdFindWrap.wdFindContinue)
                            .Application.Selection.Delete()
                            .Application.Selection.EndOf(WdUnits.wdLine)


                        End If

                    Next i


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


                    If ttls.Cells(2 + ((iii + 1) * 9), 7).value > 0 Then


                        .Application.Selection.InsertBreak(WdBreakType.wdPageBreak)

                        .Application.Selection.EndKey(WdUnits.wdStory)


                    End If


                End With


            End If

            progPcnt += ((30) / daystotal)

            Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"


        Next iii

        progPcnt = 99

        Button1.Text = $"Working... {wf.Text(progPcnt, "0")}%"


        objExport.ExportAsFixedFormat(TextBox2.Text, WdExportFormat.wdExportFormatPDF, OpenAfterExport:=True)

        objExport.Close(False)

        My.Computer.FileSystem.DeleteFile(xDoc)

        My.Computer.FileSystem.DeleteFile(iDoc)

        Me.Activate()



exitsub:

        Button1.Text = "GENERATE CRIB SHEET"

        Me.UseWaitCursor = False

        Me.Cursor = Cursors.Default



        app.DisplayAlerts = True

        app.ScreenUpdating = True

        app.EnableEvents = True

        app.Calculation = XlCalculation.xlCalculationAutomatic

        thisWb.Close(SaveChanges:=False)


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

        app.DisplayAlerts = True

        app.ScreenUpdating = True

        app.EnableEvents = True

        app.Calculation = XlCalculation.xlCalculationAutomatic

        thisWb.Close(SaveChanges:=False)






        wApp.ScreenUpdating = True

        wApp.Options.CheckGrammarAsYouType = True

        wApp.Options.CheckGrammarWithSpelling = True

        wApp.Options.CheckSpellingAsYouType = True

        wApp.Options.AnimateScreenMovements = True

        wApp.Options.BackgroundSave = True

        wApp.Options.CheckHangulEndings = True

        wApp.Options.DisableFeaturesbyDefault = False


        MsgBox("Yikes dude, looks like something went hecka wrong. Please email mitch@rocksolidrestaurants.com right away with any details, and make sure to attach this excel file in your message. " & Err.Description)

    End Sub
    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        app.Quit()

        wApp.Quit(WdSaveOptions.wdDoNotSaveChanges)

        Marshal.FinalReleaseComObject(wApp)

        Marshal.FinalReleaseComObject(app)

        GC.Collect()

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load

        Me.DateTimePicker1.Value = DateAdd("d", 1, Today())

        Me.DateTimePicker2.Value = DateAdd("d", 1, Today())

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
