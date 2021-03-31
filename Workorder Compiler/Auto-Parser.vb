Public Class Form1
    Dim currententry As Integer = 1
    Dim counter As Integer = 1
    Dim crrnt As String
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub
    Private Sub dby_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles dby.CheckedChanged
        'hides data destination if N/A
        If dby.Checked = True Then
            t12.Enabled = True
            t12.Text = ""
            t12.Visible = True
            l13.Visible = True
            t12.Select()
        ElseIf dby.Checked = False Then
            t12.Enabled = False
            t12.Visible = False
            t12.Text = "N/A"
            l13.Visible = False
        End If

    End Sub
    Private Sub Clear_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Clear.Click
        'clears all inputs except Days to Contact and CSR name
        t1.Text = ""
        t2.Text = ""
        t3.Text = ""
        t4.Text = ""
        't5.Text = ""
        t6.Text = ""
        t7.Text = ""
        t8.Text = ""
        t9.Text = ""
        t10.Text = ""
        t12.Text = ""
        t13.Text = ""
        acy.Checked = False
        acn.Checked = False
        wm.Checked = False
        wa.Checked = False
        wr.Checked = False
        wn.Checked = False
        dby.Checked = False
        dbn.Checked = False
        al.Checked = False
        ad.Checked = False
        at.Checked = False
        pd.Checked = False
        ps.Checked = False
        pt.Checked = False
        hd.Checked = False
        ot.Checked = False
        pl.Checked = False
    End Sub
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        MessageBox.Show("All code written by Nicholas Fry, Version 3.2 by Nicholas Fry and Callie Gary", "About")
    End Sub
    Private Sub ButtonCopy_Click(sender As Object, e As EventArgs) Handles ButtonCopy.Click

    End Sub

    Private Sub copy_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles copy.Click

        'If t11.Text = "" Then
        'MessageBox.Show("Please enter days to contact by clicking" & Environment.NewLine & "'Days to Contact' in the Lower Left.", "Error: Days to contact not set")
        'Removing the "days to contact" requirement for the program as we are consistantly 24 hours to contact, can add back in later if that changes
        'Button has been disabled 
        If t5.Text = "" Then
            MessageBox.Show("Please enter POS ID in the field in the upper right.", "Error: CSR name not set")
        Else
            'dim variables, t() is text fields, c for counting, final is the output string, rb() is the raw radio button inputs, psw is powerspec warranty warning
            Dim t() As String = {t1.Text.ToString, t2.Text.ToString, t3.Text.ToString, t4.Text.ToString, t5.Text.ToString, t6.Text.ToString, t7.Text.ToString, t8.Text.ToString, t9.Text.ToString, t10.Text.ToString, t11.Text.ToString, t12.Text.ToString, t13.Text.ToString}
            Dim c As Integer = 0
            Dim final As String = ""
            Dim rb(5) As String
            'add dashes to phone number
            If Len(t(1)) = 10 Then
                t(1) = t(1).Insert(3, "-")
                t(1) = t(1).Insert(7, "-")
            End If
            If Len(t(2)) = 10 Then
                t(2) = t(2).Insert(3, "-")
                t(2) = t(2).Insert(7, "-")
            End If
            'Add grammar to phone numbers
            If t(2) <> "" Then
                t(2) = (" or " & t(2))
            End If
            'fixes grammar for data destination
            If dby.Checked = True Then
                t(11) = (" to: " & t(11))
            End If
            'Translate radio button input
            If acy.Checked = True Then
                rb(1) = "Yes"
            End If
            If acn.Checked = True Then
                rb(1) = "No"
            End If
            If wm.Checked = True Then
                rb(2) = "Manufacturer's"
            End If
            If wa.Checked = True Then
                rb(2) = "ADH Service Plan"
            End If
            If wr.Checked = True Then
                rb(2) = "Carry-in Extension"
            End If
            If wn.Checked = True Then
                rb(2) = "No Warranty"
            End If
            If dby.Checked = True Then
                rb(3) = "Yes"
            End If
            If dbn.Checked = True Then
                rb(3) = "No, Contact if Necessary"
            End If
            If al.Checked = True Then
                rb(4) = "Apple Laptop"
            End If
            If ad.Checked = True Then
                rb(4) = "Apple Desktop"
            End If
            If at.Checked = True Then
                rb(4) = "Apple Tablet"
            End If
            If pl.Checked = True Then
                rb(4) = "Laptop"
            End If
            If pd.Checked = True Then
                rb(4) = "Desktop"
            End If
            If pt.Checked = True Then
                rb(4) = "Tablet"
            End If
            If ps.Checked = True Then
                rb(4) = "PowerSpec Desktop"
            End If
            If hd.Checked = True Then
                rb(4) = "Hard Drive"
            End If
            If ot.Checked = True Then
                rb(4) = "Other"
            End If
            'Replaces blank fields with N/A
            For c = 0 To 12
                If t(c) = "" Then
                    t(c) = "N/A"
                End If
            Next
            'fix redundant N/A's
            If t(2) = "N/A" Then
                t(2) = ""
            End If
            If t(11) = "N/A" Then
                t(11) = ""
            End If
            'fix data backup decline if unchecked
            If rb(3) = "" Then
                rb(3) = "No, Contact if Necessary"
            End If
            'fix Email N/A if not provided
            If t(3) = "N/A" Then
                t(3) = "No Email"
            End If
            'Replaces blank or unchecked radio button inputs with N/A
            c = 0
            For c = 0 To 4
                If rb(c) = "" Then
                    rb(c) = "N/A"
                End If
            Next
            'Check for PowerSpec under Warranty
            If ps.Checked = True Then
                If wm.Checked = True Then
                    rb(4) = "POWERSPEC UNDER WARRANTY 48 HOURS"
                    t(10) = "48 HOURS"
                ElseIf wa.Checked = True Then
                    rb(4) = "POWERSPEC UNDER WARRANTY 48 HOURS"
                    t(10) = "48 HOURS"
                ElseIf wr.Checked = True Then
                    rb(4) = "POWERSPEC UNDER WARRANTY 48 HOURS"
                    t(10) = "48 HOURS"
                End If
            End If
            'Check for Apple, and fix Days to contact if so
            If al.Checked = True Or ad.Checked = True Or at.Checked = True Then
                t(10) = t14.Text.ToString
            End If

            t(4) = (Environment.NewLine & "Checked-in by: " & t(4))
            final = (t(0) & " :: " & t(1) & t(2) & Environment.NewLine & t(3) & " :: " & " Device: " & rb(4) & Environment.NewLine & "Problem: " & t(5) & Environment.NewLine & "Password: " & t(6) & Environment.NewLine & "Approved Services: " & t(7) & Environment.NewLine & "Date of Last Backup: " & t(8) & Environment.NewLine & "Current AV: " & t(9) & Environment.NewLine & "Days to Contact: " & t(10) & Environment.NewLine & "AC Adapter: " & rb(1) & Environment.NewLine & "Coverage: " & rb(2) & Environment.NewLine & "Data Backup: " & rb(3) & t(11) & Environment.NewLine & "Condition: " & t(12) & t(4))
            'copy final output
            My.Computer.Clipboard.SetText(final)
        End If


    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        t11.Text = InputBox("Current days to contact for Non-Apple Units:", "Days to Contact")
        t14.Text = InputBox("Current days to contact for Apple Units:", "Days to Contact")
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        MessageBox.Show("All code written by Nicholas Fry", "About")
    End Sub

    Private Sub ButtonClear_Click(sender As Object, e As EventArgs) Handles ButtonClear.Click
        ' Clears all given fields in the form
        RadioNod32.Checked = False
        RadioIs.Checked = False
        RadioNoAv.Checked = False
        RadioDbu1Tb.Checked = False
        RadioDbu2Tb.Checked = False
        RadioNoDbu.Checked = False
        CheckExpedite.Checked = False
        CheckMca.Checked = False
        TextName.Text = ""
        TextPrimaryPhone.Text = ""
        TextSecondaryPhone.Text = ""
        TextEmail.Text = ""
        CheckPhone.Checked = False
        CheckEmail.Checked = False
        CheckText.Checked = False
        ComboBoxTextNumber.Text = ""
        ComboBrand.SelectedIndex = 0
        TextPassword.Text = ""
        ComboDevice.SelectedIndex = 0
        TextSerial.Text = ""
        TextCondition = ""
        RadioACYes.Checked = False
        RadioACNo.Checked = False
        TextTransaction.Text = ""
        TextProblem.Text = ""
        RadioWarMfr.Checked = False
        RadioWarAcc.Checked = False
        RadioWarRep.Checked = False
        RadioWarMCA.Checked = False
        RadioWarNone.Checked = False
    End Sub


    '    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        MessageBox.Show("Welcome to the Store Stock Quick Parser or SSQP [su-Skwip]" _
    '& Environment.NewLine & Environment.NewLine &
    '"-Scan all the necessary clearance lables into the Input" & Environment.NewLine &
    '"-Click Parse to append them with '-45' " & Environment.NewLine &
    '"-Clicking Previous will let you manually cycle backward through the list." & Environment.NewLine &
    '"-Clicking Next will move to the next entry, copy it, and rename Next to Append." & Environment.NewLine &
    '"-Clicking Append will add description text around the clearance tag and copy it." & Environment.NewLine &
    '"-Repeat as needed", "Help")
    '    End Sub


    '    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        tss.Text = ""
    '        tss.Enabled = True
    '        currententry = 1
    '        bnxt.Text = "Next"
    '    End Sub

    '    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Dim orig As String = tss.Text
    '        Dim entrys() As String = orig.Split(New String() {Environment.NewLine}, StringSplitOptions.None)
    '        Dim ln As Integer = entrys.Length
    '        Dim x As Integer = 0
    '        For x = 0 To (ln - 1)
    '            If entrys(x) <> "" Then
    '                entrys(x) = String.Concat(entrys(x), "-45")
    '            End If

    '        Next
    '        tss.Text = ""
    '        x = 0
    '        For Each s In entrys
    '            If x = 0 Then
    '                tss.Text = s
    '                x += 1
    '            Else
    '                tss.Text = (tss.Text & Environment.NewLine & s)
    '                x += 1
    '            End If
    '        Next
    '        tss.Enabled = False
    '    End Sub

    '    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        If counter = 1 Then
    '            Dim list As String = tss.Text
    '            Dim term As Integer = InStr(currententry, list, "-45", CompareMethod.Text)
    '            Dim len As Integer
    '            Dim start As Integer = currententry
    '            If term <> 0 Then
    '                term += 3
    '                tss.SelectionStart = currententry - 1
    '                tss.SelectionLength = (term - currententry)
    '                len = (term - currententry)
    '                currententry = term
    '            Else
    '                MessageBox.Show("End of List", "Error")
    '                tss.DeselectAll()
    '                currententry = 1
    '                GoTo endof
    '            End If

    '            If currententry < 0 Then
    '                currententry = 1
    '            End If

    '            crrnt = Mid(list, start, len)
    '            crrnt = crrnt.Trim(Environment.NewLine)
    '            crrnt = crrnt.Trim(vbCr)
    '            crrnt = crrnt.Trim(vbLf)
    '            My.Computer.Clipboard.SetText(crrnt)
    '            counter = 2
    '            bnxt.Text = "Append"
    '        ElseIf counter = 2 Then
    '            crrnt = crrnt.Trim(Environment.NewLine)
    '            crrnt = ("Store Stock restore     " & crrnt & Environment.NewLine & csr.Text)
    '            My.Computer.Clipboard.SetText(crrnt)
    '            counter = 1
    '            bnxt.Text = "Next"
    '        End If

    'endof:

    '    End Sub

    '    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '        Dim list As String = tss.Text
    '        Dim term As Integer = InStr(currententry, list, "-45", CompareMethod.Text)
    '        If currententry <> 1 Then
    '            term += 3
    '            tss.SelectionStart = currententry - (term - currententry - 1)
    '            tss.SelectionLength = (term - currententry)
    '            currententry = currententry - (term - currententry)
    '        Else
    '            MessageBox.Show("Top of List", "Error")
    '            tss.DeselectAll()
    '            currententry = 1
    '        End If
    '        If currententry < 0 Then
    '            currententry = 1
    '        End If
    '        counter = 1
    '        bnxt.Text = "Next"
    '    End Sub


End Class
