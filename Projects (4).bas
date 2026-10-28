Attribute VB_Name = "Projects"
Option Explicit

' =====================================================================
'  Project follow-up for the Human Capital agency
'  Creates the biweekly project-update drafts in the agency mailbox.
'  All Arabic wording lives in the texts sheet - nothing is hard-coded here.
' =====================================================================

Private Const C_NAVY As Long = 4661525      ' #152446
Private Const C_TEAL As Long = 12892747     ' #4BB8C4
Private Const C_GREEN As Long = 6530953     ' #89A563
Private Const C_YELLOW As Long = 14810623   ' #FFF8E1

' columns of the departments sheet
Private Const D_NAME As Long = 2
Private Const D_OWNER As Long = 3
Private Const D_STATUS As Long = 4
Private Const D_REP As Long = 5
Private Const D_MAIL As Long = 6
Private Const D_CALL As Long = 7
Private Const D_ALT As Long = 8
Private Const D_ALTMAIL As Long = 9
Private Const D_ALTCALL As Long = 10
Private Const D_UNTIL As Long = 11
Private Const D_SUBJ As Long = 12
Private Const D_XTO As Long = 13
Private Const D_XCC As Long = 14
Private Const D_FILE As Long = 15
Private Const D_LINK As Long = 16
Private Const D_EN As Long = 17
Private Const D_ENTRY As Long = 19
Private Const D_STORE As Long = 20
Private Const D_CONV As Long = 21
Private Const D_SENT As Long = 22
Private Const D_REPLY As Long = 23
Private Const D_FIRST As Long = 6           ' first data row

Private gText As Object
Private gReady As Boolean

' =====================================================================
'  setup
' =====================================================================
Private Sub Init()
    Dim ws As Worksheet, r As Long, last As Long
    If gReady Then Exit Sub
    Set gText = CreateObject("Scripting.Dictionary")
    Set ws = SH("shText")
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 6 To last
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" Then
            gText(Trim$(CStr(ws.Cells(r, 2).Value))) = CStr(ws.Cells(r, 3).Value)
        End If
    Next r
    gReady = True
End Sub

Private Function SH(id As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If CStr(ws.Range("AZ1").Value) = id Then
            Set SH = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 11, , "Sheet not found: " & id
End Function

Private Function TX(k As String) As String
    If gText Is Nothing Then Init
    If gText.Exists(k) Then TX = gText(k) Else TX = k
End Function

Private Function SetVal(r As Long) As Variant
    SetVal = SH("shSet").Cells(r, 3).Value
End Function

Private Function FontName() As String
    FontName = Trim$(CStr(SetVal(10)))
    If FontName = "" Then FontName = "Somar"
End Function

Private Function DeptEN(r As Long) As String
    ' pop-up windows cannot show Arabic on this machine, so they use the English label
    Dim ws As Worksheet
    Set ws = SH("shDept")
    DeptEN = Trim$(CStr(ws.Cells(r, D_EN).Value))
    If DeptEN = "" Then DeptEN = "Dept " & (r - D_FIRST + 1)
End Function

Private Function TeamEN(r As Long) As String
    Dim ws As Worksheet
    Set ws = SH("shTeam")
    TeamEN = Trim$(CStr(ws.Cells(r, 4).Value))
    If TeamEN = "" Then TeamEN = "User " & (r - 5)
End Function

Private Function CleanPath(v As Variant) As String
    ' tolerates the quotes that Windows "Copy as path" adds, and stray spaces
    Dim p As String
    p = Trim$(CStr(Nz(v, "")))
    Do While Left$(p, 1) = """" Or Left$(p, 1) = "'"
        p = Mid$(p, 2)
    Loop
    Do While Right$(p, 1) = """" Or Right$(p, 1) = "'"
        p = Left$(p, Len(p) - 1)
    Loop
    CleanPath = Trim$(p)
End Function

Private Function FileThere(p As String) As Boolean
    ' never raises: an unusable path simply counts as "not found"
    On Error Resume Next
    If p = "" Then Exit Function
    FileThere = (Dir$(p) <> "")
    If Err.Number <> 0 Then
        Err.Clear
        FileThere = False
    End If
    On Error GoTo 0
End Function

Private Function FileMB(p As String) As Double
    On Error Resume Next
    FileMB = FileLen(p) / 1048576
    If Err.Number <> 0 Then
        Err.Clear
        FileMB = 0
    End If
    On Error GoTo 0
End Function

Private Function Blank(v As Variant) As Boolean
    Blank = (IsEmpty(v) Or Trim$(CStr(v)) = "")
End Function

Private Function LastDeptRow() As Long
    Dim ws As Worksheet
    Set ws = SH("shDept")
    LastDeptRow = ws.Cells(ws.Rows.Count, D_NAME).End(xlUp).Row
    If LastDeptRow < D_FIRST Then LastDeptRow = D_FIRST - 1
End Function

' =====================================================================
'  buttons and the brand gradient
' =====================================================================
Public Sub SetupButtons()
    Dim ws As Worksheet, s As Shape, i As Long
    Init
    Application.ScreenUpdating = False
    Set ws = SH("shMain")
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    AddButton ws, TX("b_start"), "Btn_Start", 1, True
    AddButton ws, TX("b_link"), "Btn_Link", 2, False
    AddButton ws, TX("b_replies"), "Btn_Replies", 3, False
    AddButton ws, TX("b_remind"), "Btn_Remind", 4, False
    AddButton ws, TX("b_refresh"), "Btn_Refresh", 5, False
    DrawGradient ws
    Refresh "setup"
    Application.ScreenUpdating = True
    MsgBox TX("dlg_setup"), vbInformation, TX("app_title")
End Sub

Private Sub AddButton(ws As Worksheet, label As String, macro As String, idx As Long, primary As Boolean)
    Dim s As Shape, t As Double, l As Double, w As Double, h As Double, gap As Double
    h = 34: gap = 10
    If primary Then
        w = 190
        l = ws.Range("B11").Left
        t = ws.Range("B11").Top
    Else
        w = 150
        l = ws.Range("B11").Left + 190 + gap + (idx - 2) * (w + gap)
        t = ws.Range("B11").Top
    End If
    Set s = ws.Shapes.AddShape(5, l, t, w, h)          ' 5 = rounded rectangle
    s.Name = "btn_" & macro
    s.Adjustments(1) = 0.16
    s.Line.Visible = msoFalse
    With s.TextFrame2
        .TextRange.Text = label
        .TextRange.Font.Name = FontName()
        .TextRange.Font.Size = IIf(primary, 12, 10.5)
        .TextRange.Font.Bold = primary
        .TextRange.ParagraphFormat.Alignment = 2        ' centre
        .VerticalAnchor = 3
        .MarginLeft = 4: .MarginRight = 4
    End With
    If primary Then
        s.Fill.ForeColor.RGB = C_NAVY
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        s.Shadow.Visible = msoFalse
    Else
        s.Fill.ForeColor.RGB = RGB(255, 255, 255)
        s.Line.Visible = msoTrue
        s.Line.ForeColor.RGB = C_NAVY
        s.Line.Weight = 1
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = C_NAVY
    End If
    s.OnAction = macro
End Sub

Private Sub DrawGradient(ws As Worksheet)
    ' the brand mark: stepped teal-to-green light lines on the navy band
    Dim bars As Variant, i As Long, s As Shape, l As Double, t As Double, w As Double
    bars = Array(Array(0.08, 0.52), Array(0.02, 0.34), Array(0.1, 0.62), Array(0.04, 0.44))
    t = ws.Range("B2").Top + 10
    For i = 0 To UBound(bars)
        l = ws.Range("B2").Left + ws.Range("B2:L2").Width * bars(i)(0)
        w = ws.Range("B2:L2").Width * bars(i)(1)
        Set s = ws.Shapes.AddShape(1, l, t + i * 9, w, 3)
        s.Name = "brand_" & i
        s.Line.Visible = msoFalse
        With s.Fill
            .Visible = msoTrue
            .ForeColor.RGB = C_TEAL
            .BackColor.RGB = C_GREEN
            .TwoColorGradient 1, 1                       ' horizontal
        End With
        s.Shadow.Visible = msoFalse
    Next i
End Sub

' =====================================================================
'  who is using the file
' =====================================================================
Private Function TeamRowOfUser() As Long
    Dim ws As Worksheet, r As Long, u As String
    Set ws = SH("shTeam")
    u = LCase$(Trim$(Environ$("USERNAME")))
    If u = "" Then Exit Function
    For r = 6 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If LCase$(Trim$(CStr(ws.Cells(r, 3).Value))) = u And u <> "" Then
            TeamRowOfUser = r
            Exit Function
        End If
    Next r
End Function

Private Function DeptsOf(person As String) As String
    Dim r As Long, ws As Worksheet, out As String
    Set ws = SH("shDept")
    For r = D_FIRST To LastDeptRow()
        If Norm(CStr(ws.Cells(r, D_OWNER).Value)) = Norm(person) Then
            If out <> "" Then out = out & TX("sep")
            out = out & CStr(ws.Cells(r, D_NAME).Value)
        End If
    Next r
    DeptsOf = out
End Function

Private Function DeptsOfEN(person As String) As String
    Dim r As Long, ws As Worksheet, out As String
    Set ws = SH("shDept")
    For r = D_FIRST To LastDeptRow()
        If Norm(CStr(ws.Cells(r, D_OWNER).Value)) = Norm(person) Then
            If out <> "" Then out = out & ", "
            out = out & DeptEN(r)
        End If
    Next r
    DeptsOfEN = out
End Function

Private Function AskWho() As String
    ' returns the chosen person's name, or "" if cancelled
    Dim ws As Worksheet, r As Long, n As Long, list As String, names() As String, ans As String
    Dim guess As String, i As Long
    Set ws = SH("shTeam")
    guess = ""
    r = TeamRowOfUser()
    If r > 0 Then guess = Trim$(CStr(ws.Cells(r, 2).Value))
    If guess <> "" Then
        If MsgBox(Replace(TX("dlg_who"), "{name}", TeamEN(r)) & vbCrLf & vbCrLf & _
                  Replace(TX("dlg_who_sub"), "{depts}", DeptsOfEN(guess)), _
                  vbYesNo + vbQuestion, TX("app_title")) = vbYes Then
            AskWho = guess
            Exit Function
        End If
    End If
    n = 0
    ReDim names(1 To 50)
    For r = 6 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If Trim$(CStr(ws.Cells(r, 2).Value)) <> "" And n < 50 Then
            n = n + 1
            names(n) = Trim$(CStr(ws.Cells(r, 2).Value))
            list = list & n & ") " & TeamEN(r) & "  -  " & DeptsOfEN(names(n)) & vbCrLf
        End If
    Next r
    If n = 0 Then Exit Function
    ans = InputBox(TX("dlg_pick") & vbCrLf & vbCrLf & list, TX("app_title"), "1")
    If Trim$(ans) = "" Then Exit Function
    If Not IsNumeric(ans) Then Exit Function
    i = CLng(ans)
    If i < 1 Or i > n Then Exit Function
    AskWho = names(i)
End Function

' =====================================================================
'  dates
' =====================================================================
Private Function ArabicDate(d As Date) As String
    ArabicDate = Day(d) & " " & TX("m_" & Month(d)) & " " & Year(d)
End Function

Private Function NextMonday(afterD As Date) As Date
    Dim d As Date
    d = afterD + 1
    Do While Weekday(d, vbSunday) <> vbMonday
        d = d + 1
    Loop
    NextMonday = d
End Function

Private Function DueThursday() As Date
    ' the next sending Thursday, counting in steps from the first round
    Dim start As Date, step_ As Long, d As Date
    step_ = CLng(Nz(SetVal(12), 2))
    If step_ < 1 Then step_ = 1
    If Blank(SetVal(13)) Then
        d = Date
        Do While Weekday(d, vbSunday) <> vbThursday
            d = d + 1
        Loop
        DueThursday = d
        Exit Function
    End If
    start = CDate(SetVal(13))
    d = start
    Do While d < Date
        d = d + 7 * step_
    Loop
    DueThursday = d
End Function

Private Function Nz(v As Variant, dflt As Variant) As Variant
    If Blank(v) Then Nz = dflt Else Nz = v
End Function

Private Function Norm(s As String) As String
    Dim i As Long, ch As String, c As Long, out As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        Select Case c
            Case 32, 160, &H640
            Case &H622, &H623, &H625: out = out & ChrW(&H627)
            Case &H629: out = out & ChrW(&H647)
            Case &H649: out = out & ChrW(&H64A)
            Case &H64B To &H652
            Case Else: out = out & ch
        End Select
    Next i
    Norm = LCase$(out)
End Function

' =====================================================================
'  Outlook helpers
' =====================================================================
Private Function OutlookApp() As Object
    On Error Resume Next
    Set OutlookApp = GetObject(, "Outlook.Application")
    If OutlookApp Is Nothing Then Set OutlookApp = CreateObject("Outlook.Application")
    On Error GoTo 0
End Function

Private Function AgencyStore(app As Object) As Object
    Dim st As Object, want As String
    want = Norm(CStr(Nz(SetVal(6), "")))
    If want = "" Then Exit Function
    For Each st In app.GetNamespace("MAPI").Stores
        On Error Resume Next
        If InStr(Norm(CStr(st.DisplayName)), want) > 0 Then
            Set AgencyStore = st
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next st
End Function

Private Function AgencyDrafts(app As Object) As Object
    Dim st As Object
    Set st = AgencyStore(app)
    If st Is Nothing Then Exit Function
    On Error Resume Next
    Set AgencyDrafts = st.GetDefaultFolder(16)          ' olFolderDrafts
    On Error GoTo 0
End Function

Private Function AgencyAccount(app As Object) As Object
    Dim acc As Object, want As String
    want = Norm(CStr(Nz(SetVal(6), "")))
    For Each acc In app.Session.Accounts
        On Error Resume Next
        If InStr(Norm(CStr(acc.DisplayName)), want) > 0 Or _
           InStr(Norm(CStr(acc.SmtpAddress)), want) > 0 Then
            Set AgencyAccount = acc
            On Error GoTo 0
            Exit Function
        End If
        On Error GoTo 0
    Next acc
End Function

' =====================================================================
'  recipients
' =====================================================================
Private Function ListFor(dept As String, kind As String) As String
    ' walks the blocks top to bottom, so the order in the sheet is the order in the e-mail
    Dim ws As Worksheet, r As Long, last As Long, cur As String, out As String
    Dim k As String, v As String, want As Boolean
    Set ws = SH("shList")
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 5 To last
        k = CStr(ws.Cells(r, 1).Value)
        If Left$(k, 5) = "LIST|" Then
            cur = Mid$(k, 6)
            want = (Norm(cur) = Norm(dept) Or Norm(cur) = Norm(TX("w_all")))
        ElseIf k <> "HEAD" And want Then
            v = Trim$(CStr(ws.Cells(r, 2).Value))
            If v <> "" Then
                If Norm(CStr(ws.Cells(r, 3).Value)) = Norm(kind) Then out = AddAddr(out, v)
            End If
        End If
    Next r
    ListFor = out
End Function

Private Function AddAddr(cur As String, addr As String) As String
    Dim parts As Variant, i As Long, a As String
    AddAddr = cur
    a = Trim$(addr)
    If a = "" Then Exit Function
    parts = Split(Replace(a, ",", ";"), ";")
    For i = 0 To UBound(parts)
        a = Trim$(parts(i))
        If a <> "" Then
            If InStr(1, ";" & Norm(AddAddr) & ";", ";" & Norm(a) & ";") = 0 Then
                If AddAddr <> "" Then AddAddr = AddAddr & "; "
                AddAddr = AddAddr & a
            End If
        End If
    Next i
End Function

Private Function Delegated(r As Long) As Boolean
    Dim ws As Worksheet, untilD As Variant
    Set ws = SH("shDept")
    If Blank(ws.Cells(r, D_ALTMAIL).Value) And Blank(ws.Cells(r, D_ALT).Value) Then Exit Function
    untilD = ws.Cells(r, D_UNTIL).Value
    If Blank(untilD) Then
        Delegated = True
    ElseIf IsDate(untilD) Then
        Delegated = (Date <= CDate(untilD))
    End If
End Function

Private Function ToLine(r As Long) As String
    Dim ws As Worksheet, out As String
    Set ws = SH("shDept")
    If Delegated(r) Then
        out = AddAddr("", CStr(Nz(ws.Cells(r, D_ALTMAIL).Value, ws.Cells(r, D_ALT).Value)))
    Else
        out = AddAddr("", CStr(Nz(ws.Cells(r, D_MAIL).Value, ws.Cells(r, D_REP).Value)))
    End If
    out = AddAddr(out, ListFor(CStr(ws.Cells(r, D_NAME).Value), "To"))
    out = AddAddr(out, CStr(ws.Cells(r, D_XTO).Value))
    ToLine = out
End Function

Private Function CcLine(r As Long) As String
    Dim ws As Worksheet, out As String
    Set ws = SH("shDept")
    If Delegated(r) Then
        out = AddAddr("", CStr(Nz(ws.Cells(r, D_MAIL).Value, ws.Cells(r, D_REP).Value)))
    End If
    out = AddAddr(out, ListFor(CStr(ws.Cells(r, D_NAME).Value), "CC"))
    out = AddAddr(out, CStr(ws.Cells(r, D_XCC).Value))
    CcLine = out
End Function

Private Function CallName(r As Long) As String
    ' the preferred form of address; falls back to first + last name
    Dim ws As Worksheet, nm As String, full As String, parts As Variant
    Set ws = SH("shDept")
    If Delegated(r) Then
        nm = Trim$(CStr(ws.Cells(r, D_ALTCALL).Value))
        full = Trim$(CStr(ws.Cells(r, D_ALT).Value))
    Else
        nm = Trim$(CStr(ws.Cells(r, D_CALL).Value))
        full = Trim$(CStr(ws.Cells(r, D_REP).Value))
    End If
    If nm <> "" Then
        CallName = nm
        Exit Function
    End If
    If full = "" Then Exit Function
    parts = Split(full, " ")
    If UBound(parts) >= 1 Then
        CallName = parts(0) & " " & parts(UBound(parts))
    Else
        CallName = full
    End If
End Function

' =====================================================================
'  the e-mail body
' =====================================================================
Private Function HtmlEsc(s As String) As String
    HtmlEsc = Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

Private Function GradientBand() As String
    Dim k As Long, t As Double, f As Double, i As Long, col As String, cells As String
    Dim r1 As Long, g1 As Long, b1 As Long, r2 As Long, g2 As Long, b2 As Long
    Dim stops_ As Variant
    stops_ = Array(Array(75, 184, 196), Array(21, 36, 70), Array(137, 165, 99))
    For k = 0 To 47
        t = k / 47 * 2
        i = Int(t)
        If i > 1 Then i = 1
        f = t - i
        r1 = stops_(i)(0): g1 = stops_(i)(1): b1 = stops_(i)(2)
        r2 = stops_(i + 1)(0): g2 = stops_(i + 1)(1): b2 = stops_(i + 1)(2)
        col = Right$("0" & Hex$(CLng(r1 + (r2 - r1) * f)), 2) & _
              Right$("0" & Hex$(CLng(g1 + (g2 - g1) * f)), 2) & _
              Right$("0" & Hex$(CLng(b1 + (b2 - b1) * f)), 2)
        cells = cells & "<td height='2' bgcolor='#" & col & "' style='mso-line-height-rule:exactly;" & _
                "line-height:2px;font-size:1px;height:2px;background:#" & col & ";padding:0'>&#8203;</td>"
    Next k
    GradientBand = "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' border='0' " & _
        "style='border-collapse:collapse;width:100%'><tr>" & cells & "</tr></table>" & _
        "<div style='height:14px;line-height:14px;font-size:1px'>&nbsp;</div>"
End Function

Private Function MailHtml(r As Long, dueD As Date) As String
    Dim sF As String, sP As String, sB As String, nm As String, out As String, due As String
    sF = "font-family:" & FontName() & ",Arial,sans-serif;"
    sP = "margin:0 0 12px 0;" & sF & "font-size:11pt;color:#000000;"
    sB = "margin:0 0 8px 0;" & sF & "font-size:11pt;color:#152446;"
    nm = CallName(r)
    due = "<span style='font-weight:bold'>" & TX("d_mon") & " " & ArabicDate(dueD) & "</span>"
    out = "<div dir='rtl' style='text-align:right;" & sF & "font-size:11pt;color:#000000'>" & GradientBand()
    out = out & "<p style='" & sP & "font-weight:bold'>" & TX("mail_greet") & "</p>"
    If nm <> "" Then
        out = out & "<p style='" & sP & "'>" & Replace(TX("mail_name"), "{name}", HtmlEsc(nm)) & "</p>"
    End If
    out = out & "<p style='" & sP & "'>" & TX("mail_1") & "</p>"
    out = out & "<ul style='margin:0 0 14px 0;padding-right:24px;" & sF & "'>" & _
        "<li style='" & sB & "'>" & Replace(TX("mail_b1"), "{due}", due) & "</li>" & _
        "<li style='" & sB & "'>" & TX("mail_b2") & "</li>" & _
        "<li style='" & sB & "'>" & TX("mail_b3") & "</li></ul>"
    out = out & "<p style='" & sP & "'>" & TX("mail_note") & "</p>"
    out = out & "<p style='" & sP & "'>" & TX("mail_end") & "</p></div>"
    MailHtml = out
End Function

' =====================================================================
'  main action: create the drafts
' =====================================================================
Public Sub Btn_Start()
    Dim who As String, ws As Worksheet, r As Long, rowsSel() As Long, n As Long
    Dim files() As String, i As Long, j As Long, f As String, msg As String, app As Object
    Dim drafts As Object, made As Long, dueD As Date, warn As String, st As String
    On Error GoTo Failed
    Init
    Set ws = SH("shDept")
    who = AskWho()
    If who = "" Then Exit Sub
    dueD = NextMonday(Date)

    ' 1) which rows belong to this person
    n = 0
    ReDim rowsSel(1 To 60)
    ReDim files(1 To 60)
    For r = D_FIRST To LastDeptRow()
        If Norm(CStr(ws.Cells(r, D_OWNER).Value)) = Norm(who) And n < 60 Then
            st = Norm(CStr(ws.Cells(r, D_STATUS).Value))
            If st <> Norm(TX("w_hold")) And Trim$(CStr(ws.Cells(r, D_NAME).Value)) <> "" Then
                n = n + 1
                rowsSel(n) = r
            End If
        End If
    Next r
    If n = 0 Then
        MsgBox TX("dlg_nodept"), vbInformation, TX("app_title")
        Exit Sub
    End If

    ' 2) pick a file per department
    For i = 1 To n
        f = PickFile(DeptEN(rowsSel(i)))
        files(i) = f
        If f <> "" Then
            If FileMB(f) > CDbl(Nz(SetVal(9), 25)) Then
                warn = warn & vbCrLf & TX("dlg_big") & ": " & Mid$(f, InStrRev(f, "\") + 1)
            End If
        End If
    Next i
    For i = 1 To n
        For j = i + 1 To n
            If files(i) <> "" And files(i) = files(j) Then warn = warn & vbCrLf & TX("dlg_same")
        Next j
    Next i

    ' 3) review screen
    msg = TX("dlg_confirm") & vbCrLf
    For i = 1 To n
        r = rowsSel(i)
        msg = msg & vbCrLf & "- " & DeptEN(r)
        If Norm(CStr(ws.Cells(r, D_STATUS).Value)) = Norm(TX("w_test")) Then msg = msg & "  [" & TX("w_test") & "]"
        msg = msg & vbCrLf & "   " & TX("dlg_to") & ": " & Left$(ToLine(r), 90)
        If Delegated(r) Then msg = msg & "   (" & TX("w_delegate") & ")"
        msg = msg & vbCrLf & "   " & TX("dlg_cc") & ": " & Left$(CcLine(r), 90)
        If files(i) = "" Then
            msg = msg & vbCrLf & "   " & TX("dlg_skip") & " - " & TX("dlg_skip")
        Else
            msg = msg & vbCrLf & "   " & Mid$(files(i), InStrRev(files(i), "\") + 1)
        End If
        If Blank(ws.Cells(r, D_ENTRY).Value) And Blank(ws.Cells(r, D_CONV).Value) Then _
            msg = msg & vbCrLf & "   " & TX("dlg_nothread")
        If Trim$(ToLine(r)) = "" Then msg = msg & vbCrLf & "   " & TX("dlg_nomail")
        msg = msg & vbCrLf
    Next i
    If CleanPath(SetVal(7)) = "" Then warn = warn & vbCrLf & TX("dlg_no_pdf")
    If Not IsEmpty(DeliverAt()) Then
        If CDate(DeliverAt()) > Now Then
            msg = msg & vbCrLf & TX("dlg_delay") & ": " & Format$(CDate(DeliverAt()), "hh:nn AM/PM") & vbCrLf
        Else
            warn = warn & vbCrLf & TX("dlg_delay_past")
        End If
    End If
    If warn <> "" Then msg = msg & vbCrLf & "---" & warn & vbCrLf
    If MsgBox(msg, vbOKCancel + vbQuestion, TX("app_title")) <> vbOK Then Exit Sub

    ' 4) create the drafts
    Set app = OutlookApp()
    If app Is Nothing Then
        MsgBox TX("dlg_no_outlook"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    Set drafts = app.GetNamespace("MAPI").GetDefaultFolder(16)      ' your own Drafts
    Application.ScreenUpdating = False
    For i = 1 To n
        r = rowsSel(i)
        If files(i) <> "" And Trim$(ToLine(r)) <> "" Then
            If MakeDraft(app, drafts, r, files(i), dueD, who) Then made = made + 1
        Else
            LogRow r, who, "", TX("dlg_skip")
        End If
    Next i
    Refresh "drafts"
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    If made = 0 Then
        MsgBox TX("dlg_none"), vbInformation, TX("app_title")
    Else
        MsgBox Replace(Replace(TX("dlg_done"), "{n}", CStr(made)), _
               "{box}", CStr(Nz(SetVal(6), ""))), vbInformation, TX("app_title")
    End If
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox TX("err_general") & ": " & Err.Description, vbExclamation, TX("app_title")
End Sub

Private Function FixedPdf() As String
    ' the fixed attachment: a full path, or just a file name next to this workbook
    Dim p As String
    p = CleanPath(SetVal(7))
    If p = "" Then Exit Function
    If FileThere(p) Then
        FixedPdf = p
        Exit Function
    End If
    If InStr(p, "\") = 0 Then
        p = ThisWorkbook.Path & "\" & p
        If FileThere(p) Then FixedPdf = p
    End If
End Function

Private Function PickFile(dept As String) As String
    Dim fd As Object, p As String
    Set fd = Application.FileDialog(3)                 ' msoFileDialogFilePicker
    fd.Title = Replace(TX("dlg_file"), "{dept}", dept)
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "PowerPoint", "*.pptx; *.ppt"
    fd.Filters.Add "All files", "*.*"
    p = CleanPath(SetVal(8))
    If p <> "" Then
        If Right$(p, 1) <> "\" Then p = p & "\"
        On Error Resume Next
        fd.InitialFileName = p
        On Error GoTo 0
    End If
    If fd.Show = -1 Then PickFile = fd.SelectedItems(1)
End Function

Private Function FindThreadItem(app As Object, r As Long) As Object
    ' The conversation id is the same for everyone; the EntryID/StoreID pair is not,
    ' so it is only used as a shortcut for whoever linked the thread on this machine.
    Dim ws As Worksheet, ns As Object, it As Object, conv As String, st As Object
    Set ws = SH("shDept")
    Set ns = app.GetNamespace("MAPI")
    conv = Trim$(CStr(ws.Cells(r, D_CONV).Value))

    If Not Blank(ws.Cells(r, D_ENTRY).Value) Then
        On Error Resume Next
        If Not Blank(ws.Cells(r, D_STORE).Value) Then
            Set it = ns.GetItemFromID(CStr(ws.Cells(r, D_ENTRY).Value), CStr(ws.Cells(r, D_STORE).Value))
        End If
        If it Is Nothing Then Set it = ns.GetItemFromID(CStr(ws.Cells(r, D_ENTRY).Value))
        Err.Clear
        On Error GoTo 0
    End If
    If Not it Is Nothing Then
        Set FindThreadItem = it
        Exit Function
    End If

    If conv = "" Then Exit Function
    Set st = AgencyStore(app)
    If st Is Nothing Then Exit Function
    Set it = NewestInConv(st, 5, conv)                    ' Sent Items of the agency mailbox
    If it Is Nothing Then Set it = NewestInConv(st, 6, conv)   ' then its Inbox
    Set FindThreadItem = it
End Function

Private Function NewestInConv(st As Object, folderId As Long, conv As String) As Object
    Dim f As Object, items As Object, it As Object, best As Object, bestD As Date, d As Date
    On Error Resume Next
    Set f = st.GetDefaultFolder(folderId)
    If f Is Nothing Then Exit Function
    Set items = f.Items
    If items Is Nothing Then Exit Function
    For Each it In items
        If it.Class = 43 Then
            If CStr(it.ConversationID) = conv Then
                d = CDate(0)
                If folderId = 5 Then d = it.SentOn Else d = it.ReceivedTime
                If best Is Nothing Or d > bestD Then
                    Set best = it
                    bestD = d
                End If
            End If
        End If
    Next it
    On Error GoTo 0
    Set NewestInConv = best
End Function

Private Function DeliverAt() As Variant
    ' the shared delivery moment: today at the time set in the settings sheet
    Dim t As String, d As Date
    t = Trim$(CStr(Nz(SetVal(15), "")))
    If t = "" Then Exit Function
    On Error Resume Next
    d = CDate(Format$(Date, "yyyy/mm/dd") & " " & t)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If
    On Error GoTo 0
    DeliverAt = d
End Function

Private Function MakeDraft(app As Object, drafts As Object, r As Long, file As String, _
                           dueD As Date, who As String) As Boolean
    ' Each step is guarded on its own, so one refusal by Outlook does not lose the draft.
    Dim ws As Worksheet, ns As Object, src As Object, m As Object, pos As Long, p2 As Long
    Dim html As String, bodyHtml As String, pdf As String, subj As String
    Dim isTest As Boolean, step_ As String, note As String
    Set ws = SH("shDept")
    Set ns = app.GetNamespace("MAPI")
    isTest = (Norm(CStr(ws.Cells(r, D_STATUS).Value)) = Norm(TX("w_test")))

    step_ = "thread"
    On Error Resume Next
    Set src = FindThreadItem(app, r)
    Err.Clear

    step_ = "create"
    If Not src Is Nothing Then Set m = src.ReplyAll
    If m Is Nothing Then
        Err.Clear
        Set m = app.CreateItem(0)
        note = note & "new "
    End If
    If m Is Nothing Then
        On Error GoTo 0
        LogRow r, who, file, TX("err_row") & ": create"
        Exit Function
    End If

    step_ = "subject"
    subj = Trim$(CStr(ws.Cells(r, D_SUBJ).Value))
    If subj <> "" Then
        If src Is Nothing Then m.Subject = subj Else m.Subject = "RE: " & subj
    End If
    Err.Clear

    step_ = "recipients"
    SetRecipients m, ToLine(r), IIf(isTest, "", CcLine(r))
    Err.Clear

    step_ = "defer"
    If Not IsEmpty(DeliverAt()) Then
        If CDate(DeliverAt()) > Now Then m.DeferredDeliveryTime = CDate(DeliverAt())
    End If
    Err.Clear

    step_ = "body"
    bodyHtml = MailHtml(r, dueD)
    html = m.HTMLBody
    If html = "" Then html = "<html><body></body></html>"
    pos = InStr(1, LCase$(html), "<body")
    If pos > 0 Then
        p2 = InStr(pos, html, ">") + 1
        m.HTMLBody = Left$(html, p2 - 1) & bodyHtml & "<br>" & Mid$(html, p2)
    Else
        m.HTMLBody = bodyHtml & html
    End If
    Err.Clear

    step_ = "attach"
    pdf = CleanPath(SetVal(7))
    If pdf <> "" Then
        m.Attachments.Add pdf
        If Err.Number <> 0 Then                       ' try it next to this workbook
            Err.Clear
            If InStr(pdf, "\") = 0 Then
                m.Attachments.Add ThisWorkbook.Path & "\" & pdf
            Else
                m.Attachments.Add ThisWorkbook.Path & "\" & Mid$(pdf, InStrRev(pdf, "\") + 1)
            End If
            If Err.Number <> 0 Then
                note = note & "no-pdf(" & Err.Number & ") "
                Err.Clear
            End If
        End If
    Else
        note = note & "no-pdf-path "
    End If
    If file <> "" Then
        m.Attachments.Add file
        If Err.Number <> 0 Then
            note = note & "no-file(" & Err.Number & ") "
            Err.Clear
        End If
    End If

    step_ = "save"
    m.Save
    If Err.Number <> 0 Then
        LogRow r, who, file, TX("err_row") & ": save - " & Err.Description
        Err.Clear
        On Error GoTo 0
        Exit Function
    End If

    step_ = "move"
    If Not drafts Is Nothing Then
        If Not m.Parent Is Nothing Then
            If m.Parent.EntryID <> drafts.EntryID Then Set m = m.Move(drafts)
        End If
        If Err.Number <> 0 Then
            note = note & "kept-in-mailbox "
            Err.Clear
        End If
    End If
    On Error GoTo 0

    ws.Cells(r, D_FILE).Value = Mid$(file, InStrRev(file, "\") + 1)
    ws.Cells(r, D_SENT).Value = Date
    LogRow r, who, file, TX("w_draft") & " " & Trim$(note)
    MakeDraft = True
End Function

Private Sub SetRecipients(m As Object, toList As String, ccList As String)
    ' adds the addresses one by one so a single bad entry cannot stop the draft
    Dim parts As Variant, i As Long, a As String, rcp As Object
    On Error Resume Next
    m.To = ""
    m.CC = ""
    parts = Split(Replace(toList, ",", ";"), ";")
    For i = 0 To UBound(parts)
        a = Trim$(parts(i))
        If a <> "" Then
            Set rcp = m.Recipients.Add(a)
            If Not rcp Is Nothing Then rcp.Type = 1          ' olTo
        End If
    Next i
    parts = Split(Replace(ccList, ",", ";"), ";")
    For i = 0 To UBound(parts)
        a = Trim$(parts(i))
        If a <> "" Then
            Set rcp = m.Recipients.Add(a)
            If Not rcp Is Nothing Then rcp.Type = 2          ' olCC
        End If
    Next i
    m.Recipients.ResolveAll
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub LogRow(r As Long, who As String, file As String, state As String)
    Dim ws As Worksheet, dept As Worksheet, row As Long
    Set ws = SH("shLog")
    Set dept = SH("shDept")
    ws.Rows(6).Insert
    ws.Range("B6:G6").Value = Array(Now, dept.Cells(r, D_NAME).Value, who, _
        Left$(ToLine(r), 200), IIf(file = "", "", Mid$(file, InStrRev(file, "\") + 1)), state)
    ws.Range("B6").NumberFormat = "yyyy/mm/dd hh:mm"
    StyleRow ws, 6, 2, 7
End Sub

Private Sub StyleRow(ws As Worksheet, r As Long, c1 As Long, c2 As Long)
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Font.Name = FontName()
        .Font.Size = 10.5
        .Borders.LineStyle = 1
        .Borders.Color = RGB(184, 195, 209)
        .VerticalAlignment = -4108
        .HorizontalAlignment = -4108
        .WrapText = True
    End With
    ws.Rows(r).RowHeight = 22
End Sub

' =====================================================================
'  link a thread
' =====================================================================
Public Sub Btn_Link()
    Dim app As Object, item As Object, ws As Worksheet, r As Long, n As Long
    Dim list As String, rowsN() As Long, ans As String, i As Long
    On Error GoTo Failed
    Init
    Set ws = SH("shDept")
    Set app = OutlookApp()
    If app Is Nothing Then
        MsgBox TX("dlg_no_outlook"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    MsgBox TX("dlg_link_open"), vbInformation, TX("app_title")
    On Error Resume Next
    Set item = app.ActiveInspector.CurrentItem
    If item Is Nothing Then Set item = app.ActiveExplorer.Selection(1)
    On Error GoTo Failed
    If item Is Nothing Then
        MsgBox TX("dlg_link_none"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    n = 0
    ReDim rowsN(1 To 60)
    For r = D_FIRST To LastDeptRow()
        If Trim$(CStr(ws.Cells(r, D_NAME).Value)) <> "" And n < 60 Then
            n = n + 1
            rowsN(n) = r
            list = list & n & ") " & DeptEN(r) & vbCrLf
        End If
    Next r
    ans = InputBox(TX("dlg_link_pick") & vbCrLf & vbCrLf & list, TX("app_title"))
    If Trim$(ans) = "" Then Exit Sub
    If Not IsNumeric(ans) Then Exit Sub
    i = CLng(ans)
    If i < 1 Or i > n Then Exit Sub
    r = rowsN(i)
    ws.Cells(r, D_ENTRY).Value = item.EntryID
    On Error Resume Next
    ws.Cells(r, D_STORE).Value = item.Parent.StoreID
    ws.Cells(r, D_CONV).Value = item.ConversationID
    If Blank(ws.Cells(r, D_SUBJ).Value) Then
        ws.Cells(r, D_SUBJ).Value = CleanSubject(CStr(item.Subject))
    End If
    Err.Clear
    On Error GoTo Failed
    If Blank(ws.Cells(r, D_CONV).Value) Then
        MsgBox TX("dlg_link_noconv"), vbExclamation, TX("app_title")
    End If
    Refresh "link"
    ThisWorkbook.Save
    MsgBox Replace(TX("dlg_link_done"), "{dept}", DeptEN(r)), vbInformation, TX("app_title")
    Exit Sub
Failed:
    MsgBox TX("err_link") & ": " & Err.Description, vbExclamation, TX("app_title")
End Sub

Private Function CleanSubject(s As String) As String
    Dim out As String
    out = Trim$(s)
    Do While LCase$(Left$(out, 3)) = "re:" Or LCase$(Left$(out, 3)) = "fw:"
        out = Trim$(Mid$(out, 4))
    Loop
    CleanSubject = out
End Function

' =====================================================================
'  replies
' =====================================================================
Public Sub Btn_Replies()
    Dim app As Object, st As Object, inbox As Object, items As Object, it As Object
    Dim ws As Worksheet, r As Long, found As Long, conv As String, sentD As Variant
    On Error GoTo Failed
    Init
    Set ws = SH("shDept")
    Set app = OutlookApp()
    If app Is Nothing Then
        MsgBox TX("dlg_no_outlook"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    Set st = AgencyStore(app)
    If st Is Nothing Then
        MsgBox TX("dlg_no_store"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    Set inbox = st.GetDefaultFolder(6)                  ' olFolderInbox
    Set items = inbox.Items.Restrict("[ReceivedTime] >= '" & _
                Format$(Date - 120, "ddddd h:nn AMPM") & "'")
    items.Sort "[ReceivedTime]", True
    Application.ScreenUpdating = False
    For Each it In items
        On Error Resume Next
        If it.Class = 43 Then
            conv = CStr(it.ConversationID)
            For r = D_FIRST To LastDeptRow()
                If conv <> "" And conv = CStr(ws.Cells(r, D_CONV).Value) Then
                    sentD = ws.Cells(r, D_SENT).Value
                    If Blank(sentD) Or CDate(it.ReceivedTime) >= CDate(sentD) Then
                        If Blank(ws.Cells(r, D_REPLY).Value) Or _
                           CDate(it.ReceivedTime) > CDate(ws.Cells(r, D_REPLY).Value) Then
                            ws.Cells(r, D_REPLY).Value = CDate(it.ReceivedTime)
                            found = found + 1
                        End If
                    End If
                End If
            Next r
        End If
        On Error GoTo Failed
    Next it
    Refresh "replies"
    Application.ScreenUpdating = True
    ThisWorkbook.Save
    MsgBox Replace(TX("dlg_rep_done"), "{n}", CStr(found)), vbInformation, TX("app_title")
    Exit Sub
Failed:
    Application.ScreenUpdating = True
    MsgBox TX("err_replies") & ": " & Err.Description, vbExclamation, TX("app_title")
End Sub

' =====================================================================
'  reminder
' =====================================================================
Public Sub Btn_Remind()
    Dim app As Object, ap As Object, pat As Object, d As Date, t As String, step_ As Long
    On Error GoTo Failed
    Init
    Set app = OutlookApp()
    If app Is Nothing Then
        MsgBox TX("dlg_no_outlook"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    step_ = CLng(Nz(SetVal(12), 2))
    d = DueThursday()
    t = CStr(Nz(SetVal(14), "09:00"))
    If ReminderExists(app) Then
        MsgBox TX("dlg_rem_exists"), vbInformation, TX("app_title")
        Exit Sub
    End If
    Set ap = app.CreateItem(1)                          ' olAppointmentItem
    ap.Subject = TX("rem_title")
    ap.Body = TX("rem_body")
    ap.Start = CDate(Format$(d, "yyyy/mm/dd") & " " & t)
    ap.Duration = 30
    ap.ReminderSet = True
    ap.ReminderMinutesBeforeStart = 30
    ap.BusyStatus = 0
    Set pat = ap.GetRecurrencePattern
    pat.RecurrenceType = 1                              ' olRecursWeekly
    pat.Interval = step_
    pat.DayOfWeekMask = 16                              ' Thursday
    pat.PatternStartDate = d
    pat.NoEndDate = True
    ap.Save
    MsgBox Replace(TX("dlg_rem_done"), "{weeks}", CStr(step_)), vbInformation, TX("app_title")
    Exit Sub
Failed:
    MsgBox TX("err_remind") & ": " & Err.Description, vbExclamation, TX("app_title")
End Sub

Private Function ReminderExists(app As Object) As Boolean
    Dim items As Object, it As Object, want As String
    want = Norm(TX("rem_title"))
    On Error Resume Next
    Set items = app.GetNamespace("MAPI").GetDefaultFolder(9).Items     ' olFolderCalendar
    If items Is Nothing Then Exit Function
    For Each it In items
        If Norm(CStr(it.Subject)) = want Then
            ReminderExists = True
            Exit For
        End If
    Next it
    On Error GoTo 0
End Function

' =====================================================================
'  refresh the two overview sheets
' =====================================================================
Public Sub Btn_Refresh()
    Init
    Application.ScreenUpdating = False
    Refresh "manual"
    Application.ScreenUpdating = True
    ThisWorkbook.Save
End Sub

Private Sub Refresh(src As String)
    Dim ws As Worksheet, main As Worksheet, tr As Worksheet, r As Long, out As Long
    Dim who As String, due As Date, linked As String, rounds As Long, reps As Long
    Set ws = SH("shDept")
    Set main = SH("shMain")
    Set tr = SH("shTrack")

    who = ""
    r = TeamRowOfUser()
    If r > 0 Then who = Trim$(CStr(SH("shTeam").Cells(r, 2).Value))
    If who = "" Then
        main.Range("C7").Value = TX("lbl_unknown")
        main.Range("C8").Value = ""
    Else
        main.Range("C7").Value = who
        If DeptsOf(who) <> "" Then
            main.Range("C8").Value = DeptsOf(who)
        Else
            main.Range("C8").Value = TX("lbl_none")
        End If
    End If
    due = DueThursday()
    main.Range("C9").Value = ArabicDate(due)
    If due < Date Then
        main.Range("C9").Font.Color = RGB(192, 57, 43)
    Else
        main.Range("C9").Font.Color = C_NAVY
    End If
    main.Range("B42").Value = Replace(TX("lbl_updated"), "{date}", Format$(Now, "yyyy/mm/dd hh:mm"))

    main.Range("B16:G41").Clear
    tr.Range("B6:G200").Clear
    EnsureBlocks
    out = 15
    rounds = 0
    For r = D_FIRST To LastDeptRow()
        If Trim$(CStr(ws.Cells(r, D_NAME).Value)) <> "" And out < 41 Then
            out = out + 1
            rounds = rounds + 1
            If Blank(ws.Cells(r, D_CONV).Value) And Blank(ws.Cells(r, D_ENTRY).Value) Then
                linked = TX("w_notlinked")
            Else
                linked = TX("w_linked")
            End If
            main.Range(main.Cells(out, 2), main.Cells(out, 7)).Value = Array( _
                ws.Cells(r, D_NAME).Value, ws.Cells(r, D_OWNER).Value, ws.Cells(r, D_STATUS).Value, _
                linked, ws.Cells(r, D_SENT).Value, Replied(r))
            StyleRow main, out, 2, 7
            main.Cells(out, 6).NumberFormat = "yyyy/mm/dd"
            If linked = TX("w_notlinked") Then main.Cells(out, 5).Font.Color = RGB(192, 57, 43)
            If Replied(r) = TX("st_replied") Then
                main.Cells(out, 7).Interior.Color = RGB(228, 242, 211)
                reps = reps + 1
            ElseIf Replied(r) = TX("st_waiting") Then
                main.Cells(out, 7).Interior.Color = RGB(250, 212, 208)
            End If

            tr.Range(tr.Cells(5 + rounds, 2), tr.Cells(5 + rounds, 7)).Value = Array( _
                ws.Cells(r, D_NAME).Value, RepShown(r), ws.Cells(r, D_SENT).Value, _
                ws.Cells(r, D_REPLY).Value, Replied(r), _
                IIf(Delegated(r), TX("w_delegate"), ""))
            StyleRow tr, 5 + rounds, 2, 7
            tr.Cells(5 + rounds, 4).NumberFormat = "yyyy/mm/dd"
            tr.Cells(5 + rounds, 5).NumberFormat = "yyyy/mm/dd"
            ws.Cells(r, D_LINK).Value = linked
        End If
    Next r
End Sub

Private Function RepShown(r As Long) As String
    Dim ws As Worksheet
    Set ws = SH("shDept")
    If Delegated(r) Then
        RepShown = CStr(ws.Cells(r, D_ALT).Value) & " (" & TX("w_delegate") & ")"
    Else
        RepShown = CStr(ws.Cells(r, D_REP).Value)
    End If
End Function

Private Function Replied(r As Long) As String
    Dim ws As Worksheet
    Set ws = SH("shDept")
    If Blank(ws.Cells(r, D_SENT).Value) Then
        Replied = TX("st_notsent")
    ElseIf Blank(ws.Cells(r, D_REPLY).Value) Then
        Replied = TX("st_waiting")
    ElseIf CDate(ws.Cells(r, D_REPLY).Value) >= CDate(ws.Cells(r, D_SENT).Value) Then
        Replied = TX("st_replied")
    Else
        Replied = TX("st_waiting")
    End If
End Function

Private Sub EnsureBlocks()
    ' every department gets its own To/Cc block; new departments are appended
    Dim ls As Worksheet, ds As Worksheet, r As Long, last As Long, have As Object, nm As String, k As String
    Set ls = SH("shList")
    Set ds = SH("shDept")
    Set have = CreateObject("Scripting.Dictionary")
    last = ls.Cells(ls.Rows.Count, 1).End(xlUp).Row
    For r = 5 To last
        k = CStr(ls.Cells(r, 1).Value)
        If Left$(k, 5) = "LIST|" Then have(Norm(Mid$(k, 6))) = True
    Next r
    For r = D_FIRST To LastDeptRow()
        nm = Trim$(CStr(ds.Cells(r, D_NAME).Value))
        If nm <> "" Then
            If Not have.Exists(Norm(nm)) Then
                AddBlock ls, nm
                have(Norm(nm)) = True
            End If
        End If
    Next r
End Sub

Private Sub AddBlock(ls As Worksheet, nm As String)
    Dim r As Long, j As Long
    r = ls.Cells(ls.Rows.Count, 1).End(xlUp).Row + 14
    If r < 6 Then r = 5
    ls.Cells(r, 1).Value = "LIST|" & nm
    ls.Range(ls.Cells(r, 2), ls.Cells(r, 4)).Merge
    With ls.Cells(r, 2)
        .Value = nm
        .Font.Name = FontName()
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = C_NAVY
        .HorizontalAlignment = -4152
        .IndentLevel = 1
    End With
    ls.Range(ls.Cells(r, 2), ls.Cells(r, 4)).Interior.Color = RGB(221, 240, 244)
    ls.Rows(r).RowHeight = 26
    ls.Cells(r + 1, 1).Value = "HEAD"
    ls.Range(ls.Cells(r + 1, 2), ls.Cells(r + 1, 4)).Value = Array(TX("hdr_mail"), TX("hdr_kind"), TX("hdr_note"))
    StyleRow ls, r + 1, 2, 4
    ls.Range(ls.Cells(r + 1, 2), ls.Cells(r + 1, 4)).Interior.Color = RGB(217, 223, 234)
    ls.Range(ls.Cells(r + 1, 2), ls.Cells(r + 1, 4)).Font.Bold = True
    For j = r + 2 To r + 13
        StyleRow ls, j, 2, 4
        ls.Range(ls.Cells(j, 2), ls.Cells(j, 4)).Interior.Color = C_YELLOW
    Next j
    AddKindDV ls, ls.Range(ls.Cells(r + 2, 3), ls.Cells(r + 13, 3))
End Sub

Private Sub AddKindDV(ws As Worksheet, rng As Range)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add 3, 1, 1, "CC,To"
    rng.Validation.ShowError = False
    On Error GoTo 0
End Sub
