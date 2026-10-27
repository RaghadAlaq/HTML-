Attribute VB_Name = "Murasalat"
'==============================================================================
' External correspondence follow-up - VBA version (Excel + Outlook only)
' All Arabic text is read from the sheet with id shText, so this module is ASCII.
' Buttons: Btn_Import, Btn_Update, Btn_Drafts, Btn_Replies, Btn_Move
' Run SetupButtons once after importing this module.
'==============================================================================
Option Explicit

' ---- fields of a transaction row (hidden sheet shRows)
Private Const R_TXN = 1, R_PERSON = 2, R_DEPT = 3, R_RECV = 4, R_FIRST = 5, R_SUBJ = 6
Private Const R_CREATED = 7, R_TYPE = 8, R_SENDER = 9, R_STATUS = 10, R_ACTIVE = 11
Private Const R_SOURCE = 12, R_CLOSE = 13, R_CLOSEDBY = 14, R_LC = 15, R_LR = 16
Private Const R_PENDING = 17, R_FROMH = 18, R_TOH = 19, R_MOVED = 20, R_INREP = 21
Private Const R_FWD = 22, R_ARCH = 23, R_N = 23
' ---- fields of an e-mail thread (hidden sheet shThreads)
Private Const T_CODE = 1, T_PEOPLE = 2, T_TXNS = 3, T_CREATED = 4, T_FIRST = 5, T_SENDS = 6
Private Const T_LASTID = 7, T_DRAFT = 8, T_LASTDRAFT = 9, T_ENDED = 10, T_CONV = 11, T_SUBJ = 12, T_N = 12

Private Const LOOKBACK_DAYS = 45
Private Const C_YELLOW = 14809343   ' RGB(255,248,225)
Private Const C_MULTI = 12121855    ' RGB(255,242,184)
Private Const C_HEAD = 15392729     ' RGB(217,223,234)
Private Const C_BAR = 15986909      ' RGB(221,240,244)
Private Const C_RED = 8555506       ' RGB(242,139,130)
Private Const C_REDL = 13686010     ' RGB(250,212,208)
Private Const C_GREENL = 13890276   ' RGB(228,242,211)

Private gSheets As Object, gText As Object
Private gRows As Object, gThreads As Object, gReps As Collection, gSteps As Object, gProc As Object
Private gForce As Object, gLog As Collection, gMeta As Object, gAlias As Object
Private gTh As Long, gDl As Long

'============================================================================== buttons
Public Sub Btn_Import(): RunAction "import": End Sub
Public Sub Btn_Update(): RunAction "update": End Sub
Public Sub Btn_Drafts(): RunAction "drafts": End Sub
Public Sub Btn_Replies(): RunAction "replies": End Sub
Public Sub Btn_Move(): RunAction "move": End Sub

Private Sub RunAction(act As String)
    Dim reportPath As String
    If act = "import" Then
        reportPath = PickReport()
        If reportPath = "" Then Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Working..."
    On Error GoTo Failed
    Init
    LoadState
    ReadEdits
    Select Case act
        Case "import": ImportReport reportPath
        Case "move": MoveClosed
        Case "replies": SyncOutlook True
        Case "drafts": CreateDrafts
    End Select
    ApplyReopen
    Compute
    SafetyNet
    BuildViews act
    SaveState
    WriteLog act
    ThisWorkbook.Save
    GoTo Done
Failed:
    MsgBox "Something went wrong: " & Err.Description & vbCrLf & "Nothing was saved.", vbExclamation, "Murasalat"
Done:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Function PickReport() As String
    Dim fd As Object
    Set fd = Application.FileDialog(3)        ' msoFileDialogFilePicker
    fd.AllowMultiSelect = False
    fd.Filters.Clear
    fd.Filters.Add "Report", "*.xlsx;*.xls;*.csv"
    If fd.Show = -1 Then PickReport = fd.SelectedItems(1)
End Function

'============================================================================== basics
Private Sub Init()
    Dim ws As Worksheet, r As Long, k As String
    Set gSheets = CreateObject("Scripting.Dictionary")
    For Each ws In ThisWorkbook.Worksheets
        k = CStr(ws.Range("AZ1").Value)
        If k <> "" Then Set gSheets(k) = ws
    Next ws
    Set gText = CreateObject("Scripting.Dictionary")
    With SH("shText")
        For r = 4 To .Cells(.Rows.Count, 1).End(xlUp).Row
            If .Cells(r, 1).Value <> "" Then gText(CStr(.Cells(r, 1).Value)) = CStr(.Cells(r, 2).Value)
        Next r
    End With
    gTh = Val(SH("shSet").Range("B4").Value): If gTh < 1 Then gTh = 5
    gDl = Val(SH("shSet").Range("B5").Value): If gDl < 1 Then gDl = 10
    Set gForce = CreateObject("Scripting.Dictionary")
    Set gLog = New Collection
    Set gAlias = CreateObject("Scripting.Dictionary")
    With SH("shSet")
        For r = 4 To 60
            If Trim(.Cells(r, 9).Value) <> "" And Trim(.Cells(r, 10).Value) <> "" Then
                gAlias(NormAr(.Cells(r, 9).Value)) = Trim(.Cells(r, 10).Value)
            End If
        Next r
    End With
    LoadMeta
End Sub

Private Function SH(id As String) As Worksheet
    Set SH = gSheets(id)
End Function

Private Function TX(k As String) As String
    If gText.Exists(k) Then
        TX = gText(k)
    Else
        TX = DefaultText(k)
    End If
End Function

Private Function AW(ParamArray codes() As Variant) As String
    ' builds an Arabic word from character codes (keeps this file plain ASCII)
    Dim i As Long, s As String
    For i = 0 To UBound(codes): s = s & ChrW(codes(i)): Next i
    AW = s
End Function

Private Function DefaultText(k As String) As String
    Select Case k
        Case "path_col": DefaultText = AW(&H627, &H644, &H645, &H633, &H627, &H631)                       ' path
        Case "view_word": DefaultText = AW(&H639, &H631, &H636)                                         ' view
        Case "pa_4": DefaultText = AW(&H645, &H62F, &H629, 32, &H627, &H644, &H628, &H642, &H627, &H621, 32, _
                                     &H639, &H646, &H62F, &H647)                                         ' stay
        Case "no": DefaultText = AW(&H644, &H627)
        Case "yes": DefaultText = AW(&H646, &H639, &H645)
        Case Else: DefaultText = k
    End Select
End Function

Private Function NormAr(ByVal s As Variant) As String
    ' removes spaces/diacritics and unifies alef, ta marbuta, ya - for matching only
    Dim t As String, i As Long, ch As Long, out As String
    t = Trim(CStr(s))
    For i = 1 To Len(t)
        ch = AscW(Mid$(t, i, 1))
        Select Case ch
            Case 32, 160, &H640, &H64B To &H652
            Case &H622, &H623, &H625: out = out & ChrW(&H627)
            Case &H629: out = out & ChrW(&H647)
            Case &H649: out = out & ChrW(&H64A)
            Case Else: out = out & ChrW(ch)
        End Select
    Next i
    NormAr = out
End Function

Private Function CleanName(ByVal s As Variant) As String
    Dim t As String
    t = Trim(Replace(CStr(s), vbLf, " "))
    Do While InStr(t, "  ") > 0: t = Replace(t, "  ", " "): Loop
    If gAlias.Exists(NormAr(t)) Then t = gAlias(NormAr(t))
    CleanName = t
End Function

Private Function TxnId(ByVal v As Variant) As String
    Dim t As String, i As Long, ch As Long, out As String
    If IsNumeric(v) And Not IsEmpty(v) Then
        t = CStr(CDbl(v))
    Else
        t = CStr(v)
    End If
    For i = 1 To Len(t)
        ch = AscW(Mid$(t, i, 1))
        If ch >= &H660 And ch <= &H669 Then
            out = out & ChrW(48 + ch - &H660)
        ElseIf ch <> 32 Then
            out = out & ChrW(ch)
        End If
    Next i
    If Right$(out, 2) = ".0" Then out = Left$(out, Len(out) - 2)
    TxnId = out
End Function

Private Function ToDate(ByVal v As Variant) As Variant
    ' returns a Date or Empty
    Dim res As Variant, t As String
    res = Empty
    If IsEmpty(v) Or IsNull(v) Then ToDate = res: Exit Function
    If VarType(v) = vbDate Then ToDate = v: Exit Function
    If IsNumeric(v) Then
        If CDbl(v) > 30000 And CDbl(v) < 80000 Then res = CDate(CDbl(v))
        ToDate = res: Exit Function
    End If
    t = Trim$(Replace(CStr(v), "T", " "))
    If t = "" Then ToDate = res: Exit Function
    On Error Resume Next
    If Len(t) >= 10 And Mid$(t, 5, 1) = "-" Then        ' yyyy-mm-dd [hh:mm[:ss]]
        res = DateSerial(CInt(Left$(t, 4)), CInt(Mid$(t, 6, 2)), CInt(Mid$(t, 9, 2)))
        If Len(t) >= 16 Then res = res + TimeSerial(CInt(Mid$(t, 12, 2)), CInt(Mid$(t, 15, 2)), 0)
    ElseIf Len(t) >= 10 And Mid$(t, 5, 1) = "/" Then    ' yyyy/mm/dd
        res = DateSerial(CInt(Left$(t, 4)), CInt(Mid$(t, 6, 2)), CInt(Mid$(t, 9, 2)))
    ElseIf IsDate(t) Then
        res = CDate(t)
    End If
    On Error GoTo 0
    ToDate = res
End Function

Private Function Blank(ByVal v As Variant) As Boolean
    If IsEmpty(v) Or IsNull(v) Then
        Blank = True
    ElseIf VarType(v) = vbString Then
        Blank = (Trim$(v) = "")
    End If
End Function

Private Function DayOnly(ByVal v As Variant) As Variant
    If Blank(v) Then
        DayOnly = Empty
    ElseIf VarType(v) = vbDate Or IsNumeric(v) Then
        DayOnly = DateSerial(Year(v), Month(v), Day(v))
    Else
        Dim t As Variant
        t = ToDate(v)
        If IsEmpty(t) Then DayOnly = Empty Else DayOnly = DateSerial(Year(t), Month(t), Day(t))
    End If
End Function

Private Function WorkDays(ByVal d1 As Variant, ByVal d2 As Variant) As Long
    ' working days after d1 up to d2 (Fri/Sat off + holidays from the settings sheet)
    If IsEmpty(d1) Or IsEmpty(d2) Then Exit Function
    If DayOnly(d2) <= DayOnly(d1) Then Exit Function
    WorkDays = Application.WorksheetFunction.NetworkDays_Intl(DayOnly(d1) + 1, DayOnly(d2), 7, _
               SH("shSet").Range("E4:E300"))
End Function

Private Function IsWorkDay(d As Date) As Boolean
    IsWorkDay = (WorkDays(d - 1, d) = 1)
End Function

Private Function YMD(d As Date) As String
    YMD = Right$("0" & (Year(d) Mod 100), 2) & Right$("0" & Month(d), 2) & Right$("0" & Day(d), 2)
End Function

Private Function FmtDate(ByVal v As Variant) As String
    If Blank(v) Then Exit Function
    If VarType(v) <> vbDate And Not IsNumeric(v) Then v = ToDate(v)
    If IsEmpty(v) Then Exit Function
    FmtDate = Year(v) & "/" & Right$("0" & Month(v), 2) & "/" & Right$("0" & Day(v), 2)
End Function

Private Function A1(ParamArray xs() As Variant) As Variant
    Dim r() As Variant, i As Long
    ReDim r(1 To UBound(xs) + 1)
    For i = 0 To UBound(xs): r(i + 1) = xs(i): Next i
    A1 = r
End Function

Private Function FindCode(ByVal s As String) As String
    ' finds MR-yymmdd-nn anywhere in a text
    Dim p As Long, c As String
    p = InStr(1, s, "MR-", vbTextCompare)
    Do While p > 0
        c = Mid$(s, p, 12)
        If Len(c) = 12 And c Like "MR-######-##" Then FindCode = UCase$(c): Exit Function
        p = InStr(p + 3, s, "MR-", vbTextCompare)
    Loop
End Function

'============================================================================== state
Private Sub LoadState()
    Dim v As Variant, i As Long, j As Long, a() As Variant, k As String, n As Long
    Set gRows = CreateObject("Scripting.Dictionary")
    v = ReadTable(SH("shRows"), R_N)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If v(i, 1) <> "" Then
                ReDim a(1 To R_N)
                For j = 1 To R_N: a(j) = v(i, j): Next j
                gRows(a(R_TXN) & "|" & a(R_PERSON)) = a
            End If
        Next i
    End If
    Set gThreads = CreateObject("Scripting.Dictionary")
    v = ReadTable(SH("shThreads"), T_N)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If v(i, 1) <> "" Then
                ReDim a(1 To T_N)
                For j = 1 To T_N: a(j) = v(i, j): Next j
                gThreads(CStr(a(T_CODE))) = a
            End If
        Next i
    End If
    Set gReps = New Collection
    v = ReadTable(SH("shReps"), 5)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If v(i, 1) <> "" Then
                ReDim a(1 To 5)
                For j = 1 To 5: a(j) = v(i, j): Next j
                gReps.Add a
            End If
        Next i
    End If
    Set gSteps = CreateObject("Scripting.Dictionary")
    v = ReadTable(SH("shSteps"), 6)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If v(i, 1) <> "" Then
                k = CStr(v(i, 1))
                If Not gSteps.Exists(k) Then gSteps.Add k, New Collection
                ReDim a(1 To 6)
                For j = 1 To 6: a(j) = v(i, j): Next j
                gSteps(k).Add a
            End If
        Next i
    End If
    Set gProc = CreateObject("Scripting.Dictionary")
    v = ReadTable(SH("shProc"), 1)
    If IsArray(v) Then
        For i = 1 To UBound(v, 1)
            If v(i, 1) <> "" Then gProc(CStr(v(i, 1))) = True
        Next i
    End If
End Sub

Private Function ReadTable(ws As Worksheet, nCols As Long) As Variant
    Dim last As Long, v As Variant
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If last < 2 Then ReadTable = Empty: Exit Function
    v = ws.Range(ws.Cells(2, 1), ws.Cells(last, nCols)).Value
    If Not IsArray(v) Then
        Dim one(1 To 1, 1 To 1) As Variant
        one(1, 1) = v
        ReadTable = one
    Else
        ReadTable = v
    End If
End Function

Private Sub SaveState()
    WriteDict SH("shRows"), gRows, R_N
    WriteDict SH("shThreads"), gThreads, T_N
    Dim ws As Worksheet, arr() As Variant, i As Long, j As Long, it As Variant, k As Variant, n As Long
    ' replies
    Set ws = SH("shReps")
    ws.Range("A2:E" & ws.Rows.Count).ClearContents
    If gReps.Count > 0 Then
        ReDim arr(1 To gReps.Count, 1 To 5)
        For i = 1 To gReps.Count
            For j = 1 To 5: arr(i, j) = gReps(i)(j): Next j
        Next i
        ws.Range("A2").Resize(gReps.Count, 5).Value = arr
    End If
    ' steps
    Set ws = SH("shSteps")
    ws.Range("A2:F" & ws.Rows.Count).ClearContents
    n = 0
    For Each k In gSteps.Keys: n = n + gSteps(k).Count: Next k
    If n > 0 Then
        ReDim arr(1 To n, 1 To 6)
        i = 0
        For Each k In gSteps.Keys
            For Each it In gSteps(k)
                i = i + 1
                For j = 1 To 6: arr(i, j) = it(j): Next j
            Next it
        Next k
        ws.Range("A2").Resize(n, 6).Value = arr
    End If
    ' processed mail ids (keep the latest 4000)
    Set ws = SH("shProc")
    ws.Range("A2:A" & ws.Rows.Count).ClearContents
    If gProc.Count > 0 Then
        Dim keys As Variant, start As Long
        keys = gProc.Keys
        start = 0
        If gProc.Count > 4000 Then start = gProc.Count - 4000
        ReDim arr(1 To gProc.Count - start, 1 To 1)
        For i = start To gProc.Count - 1: arr(i - start + 1, 1) = keys(i): Next i
        ws.Range("A2").Resize(gProc.Count - start, 1).Value = arr
    End If
End Sub

Private Sub WriteDict(ws As Worksheet, d As Object, nCols As Long)
    Dim arr() As Variant, i As Long, j As Long, k As Variant, a As Variant
    ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, nCols)).ClearContents
    If d.Count = 0 Then Exit Sub
    ReDim arr(1 To d.Count, 1 To nCols)
    i = 0
    For Each k In d.Keys
        i = i + 1
        a = d(k)
        For j = 1 To nCols: arr(i, j) = a(j): Next j
    Next k
    ws.Range("A2").Resize(d.Count, nCols).Value = arr
End Sub

Private Sub LoadMeta()
    ' sheet: name -> Array(dept, email, preferred name)
    Dim ws As Worksheet, r As Long, nm As String
    Set gMeta = CreateObject("Scripting.Dictionary")
    Set ws = SH("shMeta")
    For r = 5 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        nm = CleanName(ws.Cells(r, 2).Value)
        If nm <> "" Then gMeta(nm) = Array(Trim(ws.Cells(r, 3).Value), Trim(ws.Cells(r, 4).Value), Trim(ws.Cells(r, 5).Value), r)
    Next r
End Sub

Private Function MetaDept(nm As String) As String
    If gMeta.Exists(nm) Then MetaDept = gMeta(nm)(0)
End Function

Private Function MetaEmail(nm As String) As String
    If gMeta.Exists(nm) Then MetaEmail = gMeta(nm)(1)
End Function

Private Sub Note(s As String)
    gLog.Add s
End Sub

'============================================================================== reading your edits
Private Sub ReadEdits()
    Dim ws As Worksheet, r As Long, last As Long, k As String, parts() As String, txn As String
    Dim person As String, v As Variant, w As Variant, toName As String, code As String
    ' sheet: send date in column G
    Set ws = SH("shOpen")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 6 To last
        txn = CStr(ws.Cells(r, 1).Value)
        If txn <> "" Then SetPending txn, ToDate(ws.Cells(r, 7).Value)
    Next r
    '  : G send date, H transfer, J include today, K link e-mail code, L written send date
    Set ws = SH("shPeople")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 7 To last
        k = CStr(ws.Cells(r, 1).Value)
        If Left$(k, 2) = "P|" Then
            parts = Split(k, "|")
            txn = parts(1): person = parts(2)
            v = ToDate(ws.Cells(r, 7).Value)
            w = ToDate(ws.Cells(r, 12).Value)
            If FmtDate(v) <> FmtDate(w) Then SetPending txn, v
            toName = CleanName(ws.Cells(r, 8).Value)
            If toName <> "" And NormAr(toName) <> NormAr(person) Then Transfer txn & "|" & person, toName
            If Trim(ws.Cells(r, 10).Value) = TX("yes") Then gForce(txn & "|" & person) = True
            code = FindCode(CStr(ws.Cells(r, 11).Value))
            If code <> "" Then LinkCode txn, person, code
        End If
    Next r
    ' sheet: a name chosen in the yellow row = the transaction moves to that person
    Set ws = SH("shPaths")
    last = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For r = 5 To last
        k = CStr(ws.Cells(r, 1).Value)
        If Left$(k, 5) = "TADD|" Then
            toName = CleanName(ws.Cells(r, 3).Value)
            If toName <> "" Then MoveAllHolders Mid$(k, 6), toName
        End If
    Next r
End Sub

Private Sub SetPending(txn As String, v As Variant)
    Dim k As Variant, a As Variant
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_TXN) = txn And a(R_ACTIVE) = 1 And a(R_STATUS) <> "CLOSED" And a(R_ARCH) <> 1 Then
            a(R_PENDING) = v
            gRows(k) = a
        End If
    Next k
End Sub

Private Sub Transfer(key As String, toName As String)
    ' one holder's part moves to another person
    Dim a As Variant, b As Variant, nk As String
    If Not gRows.Exists(key) Then Exit Sub
    a = gRows(key)
    If a(R_ACTIVE) <> 1 Or a(R_STATUS) = "CLOSED" Or a(R_STATUS) = "FWD" Then Exit Sub
    a(R_STATUS) = "FWD": a(R_FWD) = toName: a(R_ACTIVE) = 0
    gRows(key) = a
    nk = a(R_TXN) & "|" & toName
    If gRows.Exists(nk) Then
        b = gRows(nk)
        If b(R_ACTIVE) = 1 And b(R_STATUS) <> "CLOSED" Then Exit Sub
    End If
    b = a
    b(R_PERSON) = toName: b(R_DEPT) = MetaDept(toName): b(R_STATUS) = "NEW": b(R_ACTIVE) = 1
    b(R_SOURCE) = "forward": b(R_MOVED) = Now: b(R_LC) = Empty: b(R_LR) = Empty
    b(R_FWD) = Empty: b(R_CLOSE) = Empty: b(R_CLOSEDBY) = Empty: b(R_ARCH) = 0
    gRows(nk) = b
    Note TX("ph_7") & ": " & a(R_TXN) & " - " & a(R_PERSON) & " > " & toName
End Sub

Private Sub MoveAllHolders(txn As String, toName As String)
    Dim k As Variant, a As Variant, ks As Collection
    Set ks = New Collection
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_TXN) = txn And a(R_ACTIVE) = 1 And a(R_STATUS) <> "CLOSED" And a(R_STATUS) <> "FWD" Then
            If NormAr(a(R_PERSON)) <> NormAr(toName) Then ks.Add CStr(k)
        End If
    Next k
    For Each k In ks: Transfer CStr(k), toName: Next k
End Sub

Private Sub LinkCode(txn As String, person As String, code As String)
    Dim t As Variant
    If gThreads.Exists(code) Then
        t = gThreads(code)
        If InList(t(T_PEOPLE), person) And InList(t(T_TXNS), txn) And t(T_ENDED) <> 1 Then Exit Sub
    Else
        ReDim t(1 To T_N)
        t(T_CODE) = code
        On Error Resume Next
        t(T_CREATED) = DateSerial(2000 + CInt(Mid$(code, 4, 2)), CInt(Mid$(code, 6, 2)), CInt(Mid$(code, 8, 2)))
        On Error GoTo 0
        t(T_PEOPLE) = "": t(T_TXNS) = ""
    End If
    If Not InList(t(T_PEOPLE), person) Then t(T_PEOPLE) = AddToList(t(T_PEOPLE), person)
    If Not InList(t(T_TXNS), txn) Then t(T_TXNS) = AddToList(t(T_TXNS), txn)
    t(T_ENDED) = 0
    gThreads(code) = t
    Note TX("lg_linked") & " " & txn & " > " & code
End Sub

Private Function InList(ByVal lst As Variant, ByVal item As String) As Boolean
    Dim p As Variant
    If CStr(lst) = "" Then Exit Function
    For Each p In Split(CStr(lst), "|")
        If p = item Then InList = True: Exit Function
    Next p
End Function

Private Function AddToList(ByVal lst As Variant, ByVal item As String) As String
    If CStr(lst) = "" Then AddToList = item Else AddToList = CStr(lst) & "|" & item
End Function

Private Sub ApplyReopen()
    Dim ws As Worksheet, r As Long, txn As String, k As Variant, a As Variant, best As String, bestRecv As Variant
    Set ws = SH("shClosed")
    For r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row To 6 Step -1
        If Trim(ws.Cells(r, 17).Value) = TX("yes") Then
            txn = CStr(ws.Cells(r, 1).Value)
            best = "": bestRecv = Empty
            For Each k In gRows.Keys
                a = gRows(k)
                If a(R_TXN) = txn Then
                    a(R_ARCH) = 0
                    If a(R_STATUS) = "CLOSED" Then
                        If Blank(a(R_LC)) Then a(R_STATUS) = "NEW" Else a(R_STATUS) = "SENT"
                        a(R_CLOSE) = Empty: a(R_CLOSEDBY) = Empty: a(R_PENDING) = Empty
                        If IsEmpty(bestRecv) Or a(R_RECV) >= bestRecv Then best = CStr(k): bestRecv = a(R_RECV)
                    End If
                    gRows(k) = a
                End If
            Next k
            If best <> "" Then a = gRows(best): a(R_ACTIVE) = 1: gRows(best) = a
            ws.Rows(r).Delete
            Note TX("lg_reopen") & " " & txn
        End If
    Next r
End Sub

'============================================================================== the report
Private Sub ImportReport(path As String)
    Dim wb As Workbook, v As Variant, hdr As Long, i As Long, j As Long, c As Object
    Dim cTxn As Long, cSubj As Long, cType As Long, cRecv As Long, cSent As Long, cDept As Long
    Dim cSender As Long, cRecip As Long, cCreated As Long, cFrom As Long, cTo As Long
    Dim nRows As Long, nCols As Long, txn As String, recv As Variant, rec As Variant, groups As Object
    Dim bad As String, nm As String
    Application.DisplayAlerts = False
    Set wb = Workbooks.Open(path, ReadOnly:=True, UpdateLinks:=False)
    v = wb.Worksheets(1).UsedRange.Value
    wb.Close SaveChanges:=False
    Application.DisplayAlerts = True
    nRows = UBound(v, 1): nCols = UBound(v, 2)
    ' header row = the row containing
    For i = 1 To Application.Min(25, nRows)
        For j = 1 To nCols
            If NormAr(v(i, j)) = NormAr(TX("rep_txn")) Then hdr = i: Exit For
        Next j
        If hdr > 0 Then Exit For
    Next i
    If hdr = 0 Then Note TX("lg_badrep"): Exit Sub
    For j = 1 To nCols
        Select Case NormAr(v(hdr, j))
            Case NormAr(TX("rep_txn")): cTxn = j
            Case NormAr(TX("rep_subject")): cSubj = j
            Case NormAr(TX("rep_type")): cType = j
            Case NormAr(TX("rep_recv")): cRecv = j
            Case NormAr(TX("rep_sent")): cSent = j
            Case NormAr(TX("rep_dept")): cDept = j
            Case NormAr(TX("rep_sender")): cSender = j
            Case NormAr(TX("rep_recipient")): cRecip = j
            Case NormAr(TX("rep_created")): cCreated = j
            Case NormAr(TX("rep_from")): cFrom = j
            Case NormAr(TX("rep_to")): cTo = j
        End Select
    Next j
    If cTxn = 0 Or cRecv = 0 Or cRecip = 0 Then Note TX("lg_badrep"): Exit Sub

    ' sheet: external records only, identical lines once
    Dim orig As Collection, seen As Object, sig As String, outArr() As Variant, ws As Worksheet
    Set orig = New Collection
    Set seen = CreateObject("Scripting.Dictionary")
    Set groups = CreateObject("Scripting.Dictionary")
    For i = hdr + 1 To nRows
        txn = TxnId(v(i, cTxn))
        If txn = "" Then GoTo NextLine
        If cType > 0 Then
            If InStr(1, NormAr(v(i, cType)), NormAr(TX("rep_ext"))) = 0 Then GoTo NextLine
        End If
        sig = ""
        For j = 1 To nCols: sig = sig & "|" & CStr(v(i, j)): Next j
        If seen.Exists(sig) Then GoTo NextLine
        seen(sig) = True
        orig.Add i
        recv = ToDate(v(i, cRecv))
        If IsEmpty(recv) Then bad = bad & " " & txn: GoTo NextLine
        nm = CleanName(v(i, cRecip))
        rec = Array(recv, ToDate(GV(v, i, cSent)), nm, Trim$(CStr(GV(v, i, cDept))), CleanName(GV(v, i, cSender)), _
                    Trim$(CStr(GV(v, i, cSubj))), ToDate(GV(v, i, cCreated)), Trim$(CStr(GV(v, i, cType))), _
                    Trim$(CStr(GV(v, i, cFrom))), Trim$(CStr(GV(v, i, cTo))))
        If Not groups.Exists(txn) Then groups.Add txn, New Collection
        groups(txn).Add rec
NextLine:
    Next i
    ' write
    Set ws = SH("shOrig")
    ws.Range("A4:" & ws.Cells(ws.Rows.Count, 60).Address).Clear
    ReDim outArr(1 To orig.Count + 1, 1 To nCols)
    For j = 1 To nCols: outArr(1, j) = v(hdr, j): Next j
    For i = 1 To orig.Count
        For j = 1 To nCols
            outArr(i + 1, j) = v(orig(i), j)
            If j = cRecip Or (cSender > 0 And j = cSender) Then outArr(i + 1, j) = CleanName(v(orig(i), j))
        Next j
    Next i
    ws.Range("A4").Resize(orig.Count + 1, nCols).Value = outArr
    With ws.Range("A4").Resize(1, nCols)
        .Font.Bold = True: .Interior.Color = C_HEAD
    End With
    ws.Range("A4").Resize(orig.Count + 1, nCols).Font.Name = FontName()
    ws.Range("A4").Resize(orig.Count + 1, nCols).Columns.AutoFit

    MergeReport groups
    ' report period, as written in the report itself
    Dim fh As String, th As String, gMin As Variant, gMax As Variant, g As Variant, it As Variant
    For Each g In groups.Keys
        For Each it In groups(g)
            If it(8) <> "" And (fh = "" Or it(8) < fh) Then fh = it(8)
            If it(9) <> "" And (th = "" Or it(9) > th) Then th = it(9)
            If IsEmpty(gMin) Or it(0) < gMin Then gMin = it(0)
            If IsEmpty(gMax) Or it(0) > gMax Then gMax = it(0)
        Next it
    Next g
    Dim per As String
    per = TX("period_txt")
    per = Replace(per, "{fh}", fh): per = Replace(per, "{th}", th)
    per = Replace(per, "{fg}", FmtDate(gMin)): per = Replace(per, "{tg}", FmtDate(gMax))
    per = Replace(per, "{at}", FmtDate(Date) & " " & Format(Time, "hh:nn"))
    SH("shMain").Range("F14").Value = per
    Note TX("lg_report") & " " & Dir(path) & " - " & TX("lg_txns") & " " & groups.Count
    If bad <> "" Then Note TX("lg_bad_dates") & bad
End Sub

Private Function GV(v As Variant, i As Long, c As Long) As Variant
    If c > 0 Then GV = v(i, c) Else GV = ""
End Function

Private Sub MergeReport(groups As Object)
    Dim txn As Variant, recs As Collection, rec As Variant, lastDay As Variant, firstD As Variant
    Dim holders As Object, h As Variant, k As String, a As Variant, kk As Variant, nNew As Long, nClosed As Long
    Dim archived As Object, steps As Collection, i As Long, sorted() As Variant, j As Long, tmp As Variant
    Set archived = CreateObject("Scripting.Dictionary")
    For Each kk In gRows.Keys
        If gRows(kk)(R_ARCH) = 1 Then archived(CStr(gRows(kk)(R_TXN))) = True
    Next kk
    For Each kk In gRows.Keys                  ' everything not in this report is marked
        a = gRows(kk): a(R_INREP) = 0: gRows(kk) = a
    Next kk
    For Each txn In groups.Keys
        If archived.Exists(CStr(txn)) Then GoTo NextTxn      ' closed is final
        Set recs = groups(txn)
        lastDay = Empty: firstD = Empty
        For Each rec In recs
            If IsEmpty(lastDay) Or DayOnly(rec(0)) > lastDay Then lastDay = DayOnly(rec(0))
            If IsEmpty(firstD) Or rec(0) < firstD Then firstD = rec(0)
        Next rec
        ' path steps (sorted by receipt time)
        ReDim sorted(1 To recs.Count)
        For i = 1 To recs.Count: sorted(i) = recs(i): Next i
        For i = 2 To recs.Count
            tmp = sorted(i): j = i - 1
            Do While j >= 1
                If sorted(j)(0) <= tmp(0) Then Exit Do
                sorted(j + 1) = sorted(j): j = j - 1
            Loop
            sorted(j + 1) = tmp
        Next i
        If gSteps.Exists(CStr(txn)) Then gSteps.Remove CStr(txn)
        gSteps.Add CStr(txn), New Collection
        For i = 1 To recs.Count
            gSteps(CStr(txn)).Add A1(CStr(txn), i, sorted(i)(4), sorted(i)(2), sorted(i)(3), sorted(i)(0))
        Next i
        ' holders = everyone who received it on the last day
        Set holders = CreateObject("Scripting.Dictionary")
        For Each rec In recs
            If DayOnly(rec(0)) = lastDay And rec(2) <> "" Then
                If holders.Exists(rec(2)) Then
                    h = holders(rec(2))
                    If rec(0) > h(0) Then h(0) = rec(0)
                    If Not IsEmpty(rec(1)) Then h(1) = rec(1)
                    holders(rec(2)) = h
                Else
                    holders(rec(2)) = rec
                End If
            End If
        Next rec
        If holders.Count = 0 Then GoTo NextTxn               ' no usable name: keep what we have
        For Each kk In holders.Keys
            h = holders(kk)
            k = CStr(txn) & "|" & kk
            If gRows.Exists(k) Then
                a = gRows(k)
            Else
                ReDim a(1 To R_N)
                a(R_TXN) = CStr(txn): a(R_PERSON) = kk: a(R_STATUS) = "NEW": a(R_SOURCE) = "system"
                a(R_ARCH) = 0: a(R_ACTIVE) = 1
                nNew = nNew + 1
            End If
            a(R_DEPT) = h(3): a(R_RECV) = h(0): a(R_FIRST) = firstD: a(R_SUBJ) = h(5)
            a(R_CREATED) = h(6): a(R_TYPE) = h(7): a(R_SENDER) = h(4): a(R_FROMH) = h(8): a(R_TOH) = h(9)
            a(R_INREP) = 1: a(R_SOURCE) = "system"
            If a(R_STATUS) <> "FWD" Then a(R_ACTIVE) = 1
            If Not IsEmpty(h(1)) And a(R_STATUS) <> "CLOSED" Then
                a(R_STATUS) = "CLOSED": a(R_CLOSE) = DayOnly(h(1)): a(R_CLOSEDBY) = "system"
                nClosed = nClosed + 1
            End If
            gRows(k) = a
        Next kk
        ' people who held it before but no longer do
        For Each kk In gRows.Keys
            a = gRows(kk)
            If a(R_TXN) = CStr(txn) And a(R_ACTIVE) = 1 And Not holders.Exists(a(R_PERSON)) Then
                If a(R_SOURCE) = "system" Then
                    a(R_ACTIVE) = 0: gRows(kk) = a
                ElseIf a(R_SOURCE) = "forward" Then
                    If DayOnly(a(R_MOVED)) < lastDay Then a(R_ACTIVE) = 0: gRows(kk) = a
                End If
            End If
        Next kk
NextTxn:
    Next txn
    If nNew > 0 Then Note TX("lg_new") & " " & nNew
    If nClosed > 0 Then Note TX("lg_closed") & " " & nClosed
End Sub

'============================================================================== closing / computing
Private Sub MoveClosed()
    Dim k As Variant, a As Variant, n As Long
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_ACTIVE) = 1 And a(R_STATUS) <> "CLOSED" And Not Blank(a(R_PENDING)) Then
            a(R_STATUS) = "CLOSED": a(R_CLOSE) = a(R_PENDING): a(R_CLOSEDBY) = "me": a(R_PENDING) = Empty
            gRows(k) = a
            n = n + 1
        End If
    Next k
    Note TX("lg_moved") & " " & n
End Sub

Private Function OpenTxns() As Object
    ' txn -> Collection of active row keys, for transactions not archived
    Dim d As Object, k As Variant, a As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_ACTIVE) = 1 And a(R_ARCH) <> 1 Then
            If Not d.Exists(CStr(a(R_TXN))) Then d.Add CStr(a(R_TXN)), New Collection
            d(CStr(a(R_TXN))).Add CStr(k)
        End If
    Next k
    Set OpenTxns = d
End Function

Private Sub Compute()
    ' a transaction whose every current holder is closed goes to the permanent archive
    Dim d As Object, txn As Variant, k As Variant, allClosed As Boolean, a As Variant
    Set d = OpenTxns()
    For Each txn In d.Keys
        allClosed = True
        For Each k In d(txn)
            If gRows(k)(R_STATUS) <> "CLOSED" Then allClosed = False: Exit For
        Next k
        If allClosed Then ArchiveTxn CStr(txn), d(txn)
    Next txn
End Sub

Private Sub ArchiveTxn(txn As String, keys As Collection)
    Dim ws As Worksheet, r As Long, k As Variant, a As Variant, first As Variant, closeD As Variant
    Dim people As String, depts As String, senders As String, byMe As Boolean, x As Variant, days As Long
    For Each k In keys
        a = gRows(k)
        If IsEmpty(first) Or (Not IsEmpty(a(R_FIRST)) And a(R_FIRST) < first) Then first = a(R_FIRST)
        If IsEmpty(closeD) Or a(R_CLOSE) > closeD Then closeD = a(R_CLOSE)
        people = AddUnique(people, a(R_PERSON))
        depts = AddUnique(depts, a(R_DEPT))
        senders = AddUnique(senders, a(R_SENDER))
        If a(R_CLOSEDBY) = "me" Then byMe = True
        x = a
    Next k
    days = WorkDays(first, closeD)
    Set ws = SH("shClosed")
    r = Application.Max(6, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1)
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 17)).Value = Array(txn, txn, x(R_SUBJ), x(R_TYPE), DayOnly(x(R_CREATED)), _
        DayOnly(first), DayOnly(closeD), depts, senders, people, days, IIf(days < gDl, TX("yes"), TX("no")), _
        x(R_FROMH), x(R_TOH), IIf(byMe, TX("manual"), TX("report")), Format(closeD, "yyyy-mm"), "")
    StyleRow ws, r, 2, 17
    ws.Range(ws.Cells(r, 5), ws.Cells(r, 7)).NumberFormat = "yyyy/mm/dd"
    ws.Cells(r, 12).Interior.Color = IIf(days < gDl, C_GREENL, C_REDL)
    ws.Cells(r, 17).Interior.Color = C_YELLOW
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_TXN) = txn Then a(R_ARCH) = 1: gRows(k) = a
    Next k
End Sub

Private Function AddUnique(ByVal lst As String, ByVal item As Variant) As String
    Dim s As String
    s = Trim(CStr(item))
    If s = "" Then AddUnique = lst: Exit Function
    If InStr(1, ChrW(&H60C) & " " & lst & ChrW(&H60C) & " ", ChrW(&H60C) & " " & s & ChrW(&H60C) & " ") > 0 Then
        AddUnique = lst
    ElseIf lst = "" Then
        AddUnique = s
    Else
        AddUnique = lst & ChrW(&H60C) & " " & s
    End If
End Function

Private Sub SafetyNet()
    ' an open transaction may only leave the open list by being closed (archived)
    Dim seenOpen As Object, k As Variant, a As Variant, has As Object, latest As Object, restored As String
    Set has = CreateObject("Scripting.Dictionary")
    Set latest = CreateObject("Scripting.Dictionary")
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_ARCH) <> 1 Then
            If a(R_ACTIVE) = 1 Then has(CStr(a(R_TXN))) = True
            If Not latest.Exists(CStr(a(R_TXN))) Then
                latest(CStr(a(R_TXN))) = CStr(k)
            ElseIf a(R_RECV) > gRows(latest(CStr(a(R_TXN))))(R_RECV) Then
                latest(CStr(a(R_TXN))) = CStr(k)
            End If
        End If
    Next k
    For Each k In latest.Keys
        If Not has.Exists(k) Then
            a = gRows(latest(k))
            a(R_ACTIVE) = 1
            If a(R_STATUS) = "FWD" Or a(R_STATUS) = "CLOSED" Then a(R_STATUS) = "NEW"
            gRows(latest(k)) = a
            restored = restored & " " & k
        End If
    Next k
    If restored <> "" Then Note TX("lg_restored") & restored
End Sub

Private Function RowDays(a As Variant) As Long
    Dim endD As Variant
    If Not Blank(a(R_PENDING)) Then endD = a(R_PENDING) Else endD = Date
    RowDays = WorkDays(a(R_FIRST), endD)
End Function

Private Function NeedsEmail(a As Variant, k As String) As Boolean
    If a(R_ACTIVE) <> 1 Or a(R_ARCH) = 1 Then Exit Function
    If a(R_STATUS) = "CLOSED" Or a(R_STATUS) = "FWD" Then Exit Function
    If Not Blank(a(R_PENDING)) Then Exit Function
    If gForce.Exists(k) Then NeedsEmail = True: Exit Function
    If Not IsWorkDay(Date) Then Exit Function
    If RowDays(a) < gTh Then Exit Function
    If DayOnly(a(R_LC)) = Date Then Exit Function
    If DayOnly(a(R_LR)) = Date Then Exit Function
    NeedsEmail = True
End Function

Private Function StatusText(keys As Collection) As String
    Dim k As Variant, a As Variant, pending As Boolean, maxD As Long, lc As Variant, replied As Boolean
    Dim anyContact As Boolean
    replied = True
    For Each k In keys
        a = gRows(k)
        If Not Blank(a(R_PENDING)) Then pending = True
        If RowDays(a) > maxD Then maxD = RowDays(a)
        If Not Blank(a(R_LC)) Then
            anyContact = True
            If IsEmpty(lc) Then
                lc = a(R_LC)
            ElseIf a(R_LC) > lc Then
                lc = a(R_LC)
            End If
            If Blank(a(R_LR)) Then
                replied = False
            ElseIf a(R_LR) < a(R_LC) Then
                replied = False
            End If
        End If
    Next k
    If pending Then
        StatusText = TX("st_closed")
    ElseIf maxD >= gDl - 2 Then
        StatusText = TX("st_must")
    ElseIf anyContact And Not replied Then
        If WorkDays(lc, Date) >= 1 Then StatusText = TX("st_noreply") Else StatusText = TX("st_wait")
    ElseIf anyContact And replied Then
        StatusText = TX("st_replied")
    Else
        StatusText = TX("st_under")
    End If
End Function

'============================================================================== paths
Private Function PathPeople(txn As String) As Collection
    ' Each entry: Array(name, dept, received, received from, step number, until)
    ' People who received it at the same time from the same person share one step number.
    Dim ents() As Variant, n As Long, st As Variant, k As Variant, a As Variant, man As Collection
    Dim i As Long, j As Long, tmp As Variant, arr() As Variant, seq As Long, e As Variant, fromP As String
    Dim holders As Object, endOpen As Variant, endD As Variant, lastIdx As Object, out As Collection
    ReDim ents(1 To 1)
    n = 0
    ' 1) the report's movements (the sender of the first one is only the source, not a holder)
    If gSteps.Exists(txn) Then
        For Each st In gSteps(txn)
            If CStr(st(4)) <> "" And CStr(st(4)) <> CStr(st(3)) Then
                n = n + 1: ReDim Preserve ents(1 To n)
                ents(n) = Array(CStr(st(4)), CStr(st(5)), st(6), CStr(st(3)), 0, Empty)
            End If
        Next st
    End If
    ' 2) manual moves, in the order they happened
    Set man = New Collection
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_TXN) = txn And a(R_SOURCE) = "forward" Then man.Add a
    Next k
    If man.Count > 0 Then
        ReDim arr(1 To man.Count)
        For i = 1 To man.Count: arr(i) = man(i): Next i
        For i = 2 To man.Count
            tmp = arr(i): j = i - 1
            Do While j >= 1
                If NzD(arr(j)(R_MOVED)) <= NzD(tmp(R_MOVED)) Then Exit Do
                arr(j + 1) = arr(j): j = j - 1
            Loop
            arr(j + 1) = tmp
        Next i
        For i = 1 To man.Count
            fromP = ""
            For Each k In gRows.Keys                ' who passed it to this person
                a = gRows(k)
                If a(R_TXN) = txn And a(R_STATUS) = "FWD" And a(R_FWD) = arr(i)(R_PERSON) Then fromP = CStr(a(R_PERSON))
            Next k
            n = n + 1: ReDim Preserve ents(1 To n)
            ents(n) = Array(CStr(arr(i)(R_PERSON)), CStr(arr(i)(R_DEPT)), arr(i)(R_MOVED), fromP, 0, Empty)
        Next i
    End If
    ' 3) current holders who are not in the list yet
    Set holders = CreateObject("Scripting.Dictionary")
    endOpen = Date
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_TXN) = txn And a(R_ACTIVE) = 1 And a(R_STATUS) <> "FWD" And a(R_STATUS) <> "CLOSED" Then
            holders(CStr(a(R_PERSON))) = True
            If Not Blank(a(R_PENDING)) Then endOpen = a(R_PENDING)
            tmp = False
            For i = 1 To n
                If ents(i)(0) = a(R_PERSON) Then tmp = True
            Next i
            If Not tmp Then
                n = n + 1: ReDim Preserve ents(1 To n)
                ents(n) = Array(CStr(a(R_PERSON)), CStr(a(R_DEPT)), a(R_RECV), "", 0, Empty)
            End If
        End If
    Next k
    ' 4) step numbers: same sender + same day = same step (received at the same time)
    seq = 0
    For i = 1 To n
        e = ents(i)
        If i = 1 Then
            seq = 1
        ElseIf e(3) <> "" And e(3) = ents(i - 1)(3) And DayOnly(e(2)) = DayOnly(ents(i - 1)(2)) Then
            ' same step as the previous one
        Else
            seq = seq + 1
        End If
        e(4) = seq
        ents(i) = e
    Next i
    ' 5) until when each person had it
    Set lastIdx = CreateObject("Scripting.Dictionary")
    For i = 1 To n: lastIdx(ents(i)(0)) = i: Next i
    For i = 1 To n
        e = ents(i): endD = Empty
        For j = i + 1 To n                           ' they passed it on
            If ents(j)(3) = e(0) Then endD = ents(j)(2): Exit For
        Next j
        If IsEmpty(endD) And holders.Exists(e(0)) And lastIdx(e(0)) = i Then endD = endOpen     ' still has it
        If IsEmpty(endD) Then
            For j = i + 1 To n                       ' otherwise: until the next step
                If ents(j)(4) > e(4) Then endD = ents(j)(2): Exit For
            Next j
        End If
        If IsEmpty(endD) Then endD = endOpen
        e(5) = endD
        ents(i) = e
    Next i
    Set out = New Collection
    For i = 1 To n: out.Add ents(i): Next i
    Set PathPeople = out
End Function

Private Function NzD(v As Variant) As Double
    If Blank(v) Then NzD = 0 Else NzD = CDbl(v)
End Function


'============================================================================== building the sheets
Private Function FontName() As String
    Dim f As String
    f = Trim(SH("shSet").Range("B6").Value)
    If f = "" Then f = "Somar"
    FontName = f
End Function

Private Sub StyleRow(ws As Worksheet, r As Long, c1 As Long, c2 As Long)
    StyleRange ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
End Sub

Private Sub StyleRange(rng As Range)
    With rng
        .Font.Name = FontName(): .Font.Size = 10.5: .Font.Color = 0
        .Borders.LineStyle = 1: .Borders.Color = RGB(184, 195, 209)
        .VerticalAlignment = -4108: .HorizontalAlignment = -4152: .WrapText = True
    End With
End Sub

Private Sub HeadRow(ws As Worksheet, r As Long, c1 As Long, titles As Variant)
    Dim i As Long
    For i = 0 To UBound(titles)
        ws.Cells(r, c1 + i).Value = titles(i)
    Next i
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c1 + UBound(titles)))
        .Font.Name = FontName(): .Font.Bold = True: .Font.Size = 10.5
        .Interior.Color = C_HEAD: .HorizontalAlignment = -4108: .VerticalAlignment = -4108: .WrapText = True
        .Borders.LineStyle = 1: .Borders.Color = RGB(184, 195, 209)
        .Borders(9).Weight = -4138: .Borders(9).Color = RGB(21, 36, 70)
    End With
    ws.Rows(r).RowHeight = 30
End Sub

Private Sub BarRow(ws As Worksheet, r As Long, c1 As Long, c2 As Long, txt As String)
    ws.Cells(r, c1).Value = txt
    With ws.Range(ws.Cells(r, c1), ws.Cells(r, c2))
        .Interior.Color = C_BAR
        .Font.Name = FontName(): .Font.Bold = True: .Font.Size = 12
        .VerticalAlignment = -4108
    End With
    ws.Cells(r, c1).HorizontalAlignment = -4152
    ws.Rows(r).RowHeight = 26
End Sub

Private Sub ClearFrom(ws As Worksheet, r As Long)
    ws.Range(ws.Cells(r, 1), ws.Cells(ws.Rows.Count, 30)).Clear
End Sub

Private Sub LinkPaths(order As Collection, pathRow As Object)
    Dim ws As Worksheet, r As Long, txn As Variant
    Set ws = SH("shOpen")
    HeadRow ws, 5, 15, Array(TX("path_col"))
    ws.Columns(15).ColumnWidth = 9
    ws.Range("O6:O" & ws.Rows.Count).Clear
    r = 5
    For Each txn In order
        r = r + 1
        If pathRow.Exists(CStr(txn)) Then
            ws.Hyperlinks.Add Anchor:=ws.Cells(r, 15), Address:="", _
                SubAddress:="'" & SH("shPaths").Name & "'!B" & pathRow(CStr(txn)), TextToDisplay:=TX("view_word")
            StyleRow ws, r, 15, 15
            ws.Cells(r, 15).HorizontalAlignment = -4108
            ws.Cells(r, 15).Font.Color = RGB(21, 36, 70)
            ws.Cells(r, 15).Font.Underline = 2
        End If
    Next txn
End Sub

Private Sub BuildViews(act As String)
    Dim d As Object, txn As Variant, k As Variant, a As Variant, order As Collection
    Set d = OpenTxns()
    Set order = SortByDays(d)
    AddNewNames
    BuildOpen d, order
    BuildPeople d, order
    Dim pathRow As Object
    Set pathRow = CreateObject("Scripting.Dictionary")
    BuildPaths d, order, pathRow
    LinkPaths order, pathRow
    BuildReplies d, order
    SH("shMain").Range("C5").Value = FmtDate(Date) & " " & Format(Time, "hh:nn") & " - " & TX("lg_" & act)
End Sub

Private Function SortByDays(d As Object) As Collection
    Dim arr() As Variant, i As Long, j As Long, n As Long, txn As Variant, tmp As Variant, mx As Long, k As Variant
    Dim out As Collection
    Set out = New Collection
    n = d.Count
    If n = 0 Then Set SortByDays = out: Exit Function
    ReDim arr(1 To n)
    i = 0
    For Each txn In d.Keys
        mx = 0
        For Each k In d(txn)
            If RowDays(gRows(k)) > mx Then mx = RowDays(gRows(k))
        Next k
        i = i + 1: arr(i) = Array(mx, CStr(txn))
    Next txn
    For i = 2 To n
        tmp = arr(i): j = i - 1
        Do While j >= 1
            If arr(j)(0) >= tmp(0) Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    For i = 1 To n: out.Add arr(i)(1): Next i
    Set SortByDays = out
End Function

Private Sub AddNewNames()
    ' new people from the data are added to  (never removed)
    Dim ws As Worksheet, r As Long, k As Variant, a As Variant
    Set ws = SH("shMeta")
    r = Application.Max(5, ws.Cells(ws.Rows.Count, 2).End(xlUp).Row + 1)
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_ACTIVE) = 1 And a(R_ARCH) <> 1 And Not gMeta.Exists(CStr(a(R_PERSON))) Then
            ws.Cells(r, 2).Value = a(R_PERSON): ws.Cells(r, 3).Value = a(R_DEPT)
            StyleRow ws, r, 2, 6
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Interior.Color = C_YELLOW
            gMeta(CStr(a(R_PERSON))) = Array(CStr(a(R_DEPT)), "", "", r)
            r = r + 1
        End If
    Next k
    ' missing e-mails in red
    For r = 5 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If Trim(ws.Cells(r, 2).Value) <> "" Then
            If Trim(ws.Cells(r, 4).Value) = "" Then ws.Cells(r, 4).Interior.Color = C_REDL Else ws.Cells(r, 4).Interior.Color = C_YELLOW
        End If
    Next r
End Sub

Private Sub BuildOpen(d As Object, order As Collection)
    Dim ws As Worksheet, r As Long, txn As Variant, k As Variant, a As Variant, x As Variant
    Dim people As String, depts As String, senders As String, pend As Variant, mx As Long, n As Long
    Dim arr() As Variant
    Set ws = SH("shOpen")
    ws.Range("A6:N" & ws.Rows.Count).ClearContents
    ws.Range("B6:N5000").Interior.Pattern = xlNone
    ws.Range("G6:G5000").Interior.Color = C_YELLOW
    n = order.Count
    If n = 0 Then Exit Sub
    ReDim arr(1 To n, 1 To 14)
    r = 0
    For Each txn In order
        people = "": depts = "": senders = "": pend = Empty: mx = 0
        For Each k In d(txn)
            a = gRows(k)
            people = AddUnique(people, a(R_PERSON))
            depts = AddUnique(depts, a(R_DEPT))
            senders = AddUnique(senders, a(R_SENDER))
            If Not Blank(a(R_PENDING)) Then pend = a(R_PENDING)
            If RowDays(a) > mx Then mx = RowDays(a)
            x = a
        Next k
        r = r + 1
        arr(r, 1) = txn: arr(r, 2) = txn: arr(r, 3) = x(R_SUBJ): arr(r, 4) = x(R_TYPE)
        arr(r, 5) = DayOnly(x(R_CREATED)): arr(r, 6) = DayOnly(x(R_FIRST)): arr(r, 7) = pend
        arr(r, 8) = depts: arr(r, 9) = senders: arr(r, 10) = people: arr(r, 11) = mx
        arr(r, 12) = StatusText(d(txn)): arr(r, 13) = x(R_FROMH): arr(r, 14) = x(R_TOH)
    Next txn
    ws.Range("A6").Resize(n, 14).Value = arr
    StyleRange ws.Range("B6").Resize(n, 13)
    ws.Range("E6").Resize(n, 3).NumberFormat = "yyyy/mm/dd"
    ws.Range("G6").Resize(n, 1).Interior.Color = C_YELLOW
    r = 5
    For Each txn In order                      ' yellow rows: held by more than one person
        r = r + 1
        If d(txn).Count > 1 Then
            ws.Range(ws.Cells(r, 2), ws.Cells(r, 6)).Interior.Color = C_MULTI
            ws.Range(ws.Cells(r, 8), ws.Cells(r, 10)).Interior.Color = C_MULTI
        End If
    Next txn
End Sub

Private Function DueToday() As Object
    ' key -> True for every holder due a reminder today
    Dim d As Object, k As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In gRows.Keys
        If NeedsEmail(gRows(k), CStr(k)) Then d(CStr(k)) = True
    Next k
    Set DueToday = d
End Function

Private Function TodayMailState(txn As String, person As String) As String
    Dim c As Variant, t As Variant
    For Each c In gThreads.Keys
        t = gThreads(c)
        If InList(t(T_TXNS), txn) And InList(t(T_PEOPLE), person) Then
            If InStr(1, CStr(t(T_SENDS)), FmtDate(Date)) > 0 Then TodayMailState = TX("rem_sent"): Exit Function
            If DayOnly(t(T_LASTDRAFT)) = Date Then TodayMailState = TX("rem_draft")
        End If
    Next c
End Function

Private Sub BuildPeople(d As Object, order As Collection)
    Dim ws As Worksheet, r As Long, txn As Variant, k As Variant, a As Variant, byP As Object, p As Variant
    Dim due As Object, remTxt As String, mx As Long, keys As Collection, missing As String, nDue As Long
    Dim heads As Variant
    Set ws = SH("shPeople")
    ClearFrom ws, 4
    ws.Columns(9).ColumnWidth = 40
    Set due = DueToday()
    Set byP = CreateObject("Scripting.Dictionary")
    For Each txn In order
        For Each k In d(txn)
            a = gRows(k)
            If a(R_STATUS) <> "CLOSED" Then
                If Not byP.Exists(CStr(a(R_PERSON))) Then byP.Add CStr(a(R_PERSON)), New Collection
                byP(CStr(a(R_PERSON))).Add CStr(k)
            End If
        Next k
    Next txn
    For Each k In due.Keys
        p = gRows(k)(R_PERSON)
        If MetaEmail(CStr(p)) = "" And InStr(1, missing, p) = 0 Then missing = AddUnique(missing, p)
        nDue = nDue + 1
    Next k
    ws.Range("B4").Value = TX("info_due") & " " & nDue
    ws.Range("B4").Font.Name = FontName()
    If missing <> "" Then
        ws.Range("B5").Value = TX("warn_missing") & " " & missing
        ws.Range("B5:K5").Interior.Color = C_REDL
        ws.Range("B5").Font.Bold = True: ws.Range("B5").Font.Name = FontName()
    End If
    heads = Array(TX("ph_1"), TX("ph_2"), TX("ph_3"), TX("ph_4"), TX("ph_5"), TX("ph_6"), TX("ph_7"), _
                  TX("ph_8"), TX("ph_9"), TX("ph_10"))
    r = 7
    For Each p In byP.Keys
        Set keys = byP(p)
        mx = 0
        For Each k In keys
            If RowDays(gRows(k)) > mx Then mx = RowDays(gRows(k))
        Next k
        BarRow ws, r, 2, 11, TX("employee") & ": " & p & "  " & ChrW(&H60C) & " " & TX("dept_word") & ": " & MetaDept(CStr(p)) & _
               "  " & ChrW(&H60C) & " " & TX("count_word") & ": " & keys.Count
        HeadRow ws, r + 1, 2, heads
        r = r + 2
        For Each k In keys
            a = gRows(k)
            If TodayMailState(CStr(a(R_TXN)), CStr(p)) <> "" Then
                remTxt = TodayMailState(CStr(a(R_TXN)), CStr(p))
            ElseIf due.Exists(CStr(k)) Then
                remTxt = TX("rem_due")
                If MetaEmail(CStr(p)) = "" Then remTxt = remTxt & " - " & TX("rem_noemail")
            ElseIf Not IsWorkDay(Date) Then
                remTxt = TX("rem_holiday")
            ElseIf RowDays(a) < gTh Then
                remTxt = TX("rem_notyet")
            ElseIf DayOnly(a(R_LC)) = Date Then
                remTxt = TX("rem_contacted")
            ElseIf DayOnly(a(R_LR)) = Date Then
                remTxt = TX("rem_replied")
            Else
                remTxt = ""
            End If
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)).Value = Array("P|" & a(R_TXN) & "|" & p, a(R_TXN), a(R_SUBJ), _
                DayOnly(a(R_FIRST)), TX("passed") & " " & RowDays(a) & " " & TX("days_word"), _
                OneStatus(a), a(R_PENDING), "", remTxt, "", LinkedCode(CStr(a(R_TXN)), CStr(p)), a(R_PENDING))
            StyleRow ws, r, 2, 11
            ws.Cells(r, 4).NumberFormat = "yyyy/mm/dd": ws.Cells(r, 7).NumberFormat = "yyyy/mm/dd"
            ws.Cells(r, 12).NumberFormat = "yyyy/mm/dd"
            ws.Cells(r, 7).Interior.Color = C_YELLOW: ws.Cells(r, 8).Interior.Color = C_YELLOW
            ws.Cells(r, 10).Interior.Color = C_YELLOW: ws.Cells(r, 11).Interior.Color = C_YELLOW
            If RowDays(a) >= gDl Then
                ws.Cells(r, 5).Interior.Color = C_RED
            ElseIf RowDays(a) >= gDl - 2 Then
                ws.Cells(r, 5).Interior.Color = C_REDL
            End If
            r = r + 1
        Next k
        r = r + 1
    Next p
    If r > 9 Then
        AddListDV ws.Range(ws.Cells(9, 8), ws.Cells(r, 8)), "='" & SH("shMeta").Name & "'!$B$5:$B$3000"
        AddListDV ws.Range(ws.Cells(9, 10), ws.Cells(r, 10)), TX("yes") & "," & TX("no")
    End If
End Sub

Private Function OneStatus(a As Variant) As String
    Dim c As New Collection
    c.Add CStr(a(R_TXN) & "|" & a(R_PERSON))
    OneStatus = StatusText(c)
End Function

Private Function LinkedCode(txn As String, person As String) As String
    Dim c As Variant, t As Variant, best As Variant
    For Each c In gThreads.Keys
        t = gThreads(c)
        If t(T_ENDED) <> 1 And InList(t(T_TXNS), txn) And InList(t(T_PEOPLE), person) Then LinkedCode = CStr(c)
    Next c
End Function

Private Function TextRow(k As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = SH("shText")
    For r = 4 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If ws.Cells(r, 1).Value = k Then TextRow = r: Exit Function
    Next r
End Function

Private Sub AddListDV(rng As Range, f As String)
    On Error Resume Next
    rng.Validation.Delete
    rng.Validation.Add Type:=3, AlertStyle:=2, Formula1:=f
    rng.Validation.ShowError = False
    On Error GoTo 0
End Sub

Private Sub BuildPaths(d As Object, order As Collection, pathRow As Object)
    Dim ws As Worksheet, r As Long, txn As Variant, lst As Collection, i As Long, firstAdd As Long, lastAdd As Long
    Dim stay As String
    Set ws = SH("shPaths")
    ws.Columns(5).ColumnWidth = 18
    ClearFrom ws, 5
    r = 5
    For Each txn In order
        pathRow(CStr(txn)) = r
        BarRow ws, r, 2, 5, TX("txn_word") & " " & txn & "  " & ChrW(&H60C) & " " & gRows(d(txn)(1))(R_SUBJ)
        HeadRow ws, r + 1, 2, Array(TX("pa_1"), TX("pa_2"), TX("pa_3"), TX("pa_4"))
        r = r + 2
        Set lst = PathPeople(CStr(txn))
        For i = 1 To lst.Count
            stay = ""
            If Not Blank(lst(i)(2)) And Not Blank(lst(i)(5)) Then
                stay = ChrW(&H200F) & WorkDays(DayOnly(lst(i)(2)), DayOnly(lst(i)(5))) & " " & TX("days_word")
            End If
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Value = Array("T|" & txn & "|" & i, lst(i)(4), lst(i)(0), _
                IIf(lst(i)(1) <> "", lst(i)(1), MetaDept(CStr(lst(i)(0)))), stay)
            StyleRow ws, r, 2, 5
            ws.Cells(r, 2).HorizontalAlignment = -4108
            ws.Cells(r, 5).HorizontalAlignment = -4108
            If lst(i)(5) = Date Or (i = lst.Count) Then ws.Cells(r, 5).Font.Bold = True
            r = r + 1
        Next i
        ws.Cells(r, 1).Value = "TADD|" & txn
        If lst.Count > 0 Then ws.Cells(r, 2).Value = lst(lst.Count)(4) + 1 Else ws.Cells(r, 2).Value = 1
        StyleRow ws, r, 2, 5
        ws.Cells(r, 2).HorizontalAlignment = -4108
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 4)).Interior.Color = C_YELLOW
        If firstAdd = 0 Then firstAdd = r
        lastAdd = r
        r = r + 2
    Next txn
    If lastAdd > 0 Then AddListDV ws.Range(ws.Cells(5, 3), ws.Cells(lastAdd, 3)), "='" & SH("shMeta").Name & "'!$B$5:$B$3000"
End Sub

Private Sub BuildReplies(d As Object, order As Collection)
    Dim ws As Worksheet, r As Long, txn As Variant, c As Variant, t As Variant, reps As Collection, rp As Variant
    Dim nRep As Long, lastRep As Variant, senders As String, anyMail As Boolean, st As String, wait As Variant
    Dim lastSend As Variant, sendDays As String, s As Variant, parts As Variant
    Set ws = SH("shReplies")
    ClearFrom ws, 4
    r = 4
    For Each txn In order
        BarRow ws, r, 2, 12, TX("txn_word") & " " & txn & "  " & ChrW(&H60C) & " " & gRows(d(txn)(1))(R_SUBJ)
        HeadRow ws, r + 1, 2, Array(TX("rp_1"), TX("rp_2"), TX("rp_3"), TX("rp_4"), TX("rp_5"), TX("rp_6"), _
                                    TX("rp_7"), TX("rp_8"), TX("rp_9"), TX("rp_10"), TX("rp_11"))
        r = r + 2
        anyMail = False
        For Each c In gThreads.Keys
            t = gThreads(c)
            If InList(t(T_TXNS), CStr(txn)) Then
                anyMail = True
                nRep = 0: lastRep = Empty: senders = ""
                For Each rp In gReps
                    If rp(1) = c Then
                        nRep = nRep + 1
                        senders = AddUnique(senders, rp(3))
                        If IsEmpty(lastRep) Then
                            lastRep = rp
                        ElseIf rp(2) >= lastRep(2) Then
                            lastRep = rp
                        End If
                    End If
                Next rp
                sendDays = "": lastSend = Empty
                For Each s In Split(CStr(t(T_SENDS)), ";")
                    If s <> "" Then
                        parts = Split(s, "#")
                        sendDays = AddUnique(sendDays, parts(0))
                        lastSend = parts(0)
                    End If
                Next s
                If Blank(t(T_FIRST)) Then
                    st = TX("th_draft"): wait = ""
                ElseIf Not IsEmpty(lastRep) Then
                    If FmtDate(lastRep(2)) >= CStr(lastSend) Then st = TX("th_replied"): wait = "" Else st = TX("th_wait"): wait = ""
                Else
                    st = TX("th_wait")
                    wait = ""
                End If
                If st = TX("th_wait") And CStr(lastSend) <> "" Then wait = WorkDays(ToDate(Replace(CStr(lastSend), "/", "-")), Date)
                Dim lrDate As Variant, lrText As Variant
                lrDate = "": lrText = ""
                If Not IsEmpty(lastRep) Then lrDate = lastRep(2): lrText = lastRep(4)
                ws.Range(ws.Cells(r, 2), ws.Cells(r, 12)).Value = Array( _
                    IIf(CStr(t(T_SUBJ)) <> "", t(T_SUBJ), c), Replace(CStr(t(T_PEOPLE)), "|", ChrW(&H60C) & " "), _
                    t(T_FIRST), sendDays, IIf(nRep > 0, TX("yes"), IIf(Blank(t(T_FIRST)), "", TX("no"))), _
                    IIf(nRep > 0, nRep, ""), senders, lrDate, st, wait, lrText)
                StyleRow ws, r, 2, 12
                ws.Cells(r, 4).NumberFormat = "yyyy/mm/dd": ws.Cells(r, 9).NumberFormat = "yyyy/mm/dd"
                If nRep > 0 Then
                    ws.Cells(r, 6).Interior.Color = C_GREENL
                ElseIf Not Blank(t(T_FIRST)) Then
                    ws.Cells(r, 6).Interior.Color = C_REDL
                End If
                r = r + 1
            End If
        Next c
        If Not anyMail Then
            ws.Cells(r, 2).Value = TX("no_mail_yet")
            ws.Cells(r, 2).Font.Name = FontName()
            r = r + 1
        End If
        r = r + 1
    Next txn
End Sub

Private Sub WriteLog(act As String)
    Dim ws As Worksheet, i As Long, txt As String
    Set ws = SH("shLog")
    For i = 1 To gLog.Count
        If txt <> "" Then txt = txt & vbLf
        txt = txt & gLog(i)
    Next i
    ws.Rows(5).Insert
    ws.Range("B5:D5").Value = Array(Now, TX("lg_" & act), txt)
    ws.Range("B5").NumberFormat = "yyyy/mm/dd hh:mm"
    StyleRow ws, 5, 2, 4
    ws.Range("B5:D5").Interior.Pattern = xlNone
End Sub

'============================================================================== Outlook
Private Function OL() As Object
    Dim o As Object
    On Error Resume Next
    Set o = GetObject(, "Outlook.Application")
    If o Is Nothing Then Set o = CreateObject("Outlook.Application")
    On Error GoTo 0
    Set OL = o
End Function

Private Sub SyncOutlook(Optional readReplies As Boolean = True)
    Dim app As Object, ns As Object, items As Object, it As Object, code As String, t As Variant
    Dim since As Date, k As Variant, a As Variant, nSends As Long, nReps As Long, conv As Object, c As Variant
    Set app = OL()
    If app Is Nothing Then Note TX("lg_noout"): Exit Sub
    Set ns = app.GetNamespace("MAPI")
    since = Date - LOOKBACK_DAYS
    ' conversation id -> code (to recognise replies even if the title was changed)
    Set conv = CreateObject("Scripting.Dictionary")
    For Each c In gThreads.Keys
        If CStr(gThreads(c)(T_CONV)) <> "" Then conv(CStr(gThreads(c)(T_CONV))) = CStr(c)
    Next c
    ' 1) what you actually sent
    Set items = ns.GetDefaultFolder(5).Items          ' Sent Items
    items.Sort "[SentOn]", True
    For Each it In items
        If it.Class = 43 Then
            If it.SentOn < since Then Exit For
            code = FindCode(it.Subject)
            If code <> "" And gThreads.Exists(code) Then
                t = gThreads(code)
                If InStr(1, CStr(t(T_SENDS)), it.EntryID) = 0 Then
                    t(T_SENDS) = CStr(t(T_SENDS)) & FmtDate(it.SentOn) & "#" & it.EntryID & ";"
                    If Blank(t(T_FIRST)) Then
                        t(T_FIRST) = DayOnly(it.SentOn)
                    ElseIf DayOnly(it.SentOn) < t(T_FIRST) Then
                        t(T_FIRST) = DayOnly(it.SentOn)
                    End If
                    If CStr(t(T_LASTID)) = "" Then
                        t(T_LASTID) = it.EntryID
                    Else
                        If LastSendDate(t) <= FmtDate(it.SentOn) Then t(T_LASTID) = it.EntryID
                    End If
                    On Error Resume Next
                    t(T_CONV) = it.ConversationID
                    On Error GoTo 0
                    If CStr(t(T_CONV)) <> "" Then conv(CStr(t(T_CONV))) = code
                    gThreads(code) = t
                    MarkThread code, "LC", DayOnly(it.SentOn)
                    nSends = nSends + 1
                End If
            End If
        End If
    Next it
    If Not readReplies Then GoTo Finish
    ' 2) replies: Inbox, and the folders of open transactions
    Dim folders As Collection, root As Object, subF As Object, d As Object, f As Variant
    Set folders = New Collection
    folders.Add ns.GetDefaultFolder(6)
    Set d = OpenTxns()
    On Error Resume Next
    Set root = ns.GetDefaultFolder(6).Folders(TX("ol_folder"))
    On Error GoTo 0
    If Not root Is Nothing Then
        For Each subF In root.Folders
            If d.Exists(Trim(Replace(subF.Name, TX("ol_sub"), ""))) Then folders.Add subF
        Next subF
    End If
    For Each f In folders
        Set items = f.Items
        items.Sort "[ReceivedTime]", True
        For Each it In items
            If it.Class = 43 Then
                If it.ReceivedTime < since Then Exit For
                If Not gProc.Exists(it.EntryID) Then
                    gProc(it.EntryID) = True
                    code = FindCode(it.Subject)
                    If code = "" Then
                        On Error Resume Next
                        If conv.Exists(CStr(it.ConversationID)) Then code = conv(CStr(it.ConversationID))
                        On Error GoTo 0
                    End If
                    If code = "" Then code = FindCode(Left$(it.Body, 20000))      ' reference line in the quoted e-mail
                    If code <> "" And gThreads.Exists(code) Then
                        gReps.Add A1(code, DayOnly(it.ReceivedTime), it.SenderName, ReplyText(it.Body), it.EntryID)
                        MarkThread code, "LR", DayOnly(it.ReceivedTime)
                        nReps = nReps + 1
                    End If
                End If
            End If
        Next it
    Next f
Finish:
    If nSends > 0 Then Note TX("lg_sends") & " " & nSends
    If nReps > 0 Then Note TX("lg_reps") & " " & nReps
End Sub

Private Function LastSendDate(t As Variant) As String
    Dim s As Variant, best As String
    For Each s In Split(CStr(t(T_SENDS)), ";")
        If s <> "" Then
            If Split(s, "#")(0) > best Then best = Split(s, "#")(0)
        End If
    Next s
    LastSendDate = best
End Function

Private Sub MarkThread(code As String, what As String, d As Variant)
    Dim t As Variant, k As Variant, a As Variant
    t = gThreads(code)
    For Each k In gRows.Keys
        a = gRows(k)
        If InList(t(T_TXNS), CStr(a(R_TXN))) And InList(t(T_PEOPLE), CStr(a(R_PERSON))) And a(R_ACTIVE) = 1 Then
            If what = "LC" Then
                If Blank(a(R_LC)) Then
                    a(R_LC) = d
                ElseIf d > a(R_LC) Then
                    a(R_LC) = d
                End If
                If a(R_STATUS) = "NEW" Then a(R_STATUS) = "SENT"
            Else
                If Blank(a(R_LR)) Then
                    a(R_LR) = d
                ElseIf d > a(R_LR) Then
                    a(R_LR) = d
                End If
            End If
            gRows(k) = a
        End If
    Next k
End Sub

Private Function ReplyText(ByVal b As String) As String
    ' keeps only the new part of a reply (cuts the quoted earlier e-mail)
    Dim marks As Variant, m As Variant, p As Long, cut As Long
    cut = Len(b) + 1
    marks = Array(vbCrLf & "From:", vbCrLf & ChrW(&H645) & ChrW(&H646) & ":", "-----Original", "________________")
    For Each m In marks
        p = InStr(1, b, m, vbTextCompare)
        If p > 0 And p < cut Then cut = p
    Next m
    b = Trim(Left$(b, cut - 1))
    Do While InStr(b, vbCrLf & vbCrLf & vbCrLf) > 0: b = Replace(b, vbCrLf & vbCrLf & vbCrLf, vbCrLf & vbCrLf): Loop
    ReplyText = Left$(b, 1500)
End Function

'============================================================================== reminders
Private Sub CreateDrafts()
    Dim due As Object, k As Variant, a As Variant, txnHolders As Object, txn As Variant, uf As Object
    Dim p As Variant, root As String, groups As Object, people As Object, missing As String, g As Variant
    Dim app As Object, nMade As Long, names As String
    SyncOutlook True                                   ' know what was sent / answered before deciding
    Compute
    Set due = DueToday()
    If due.Count = 0 Then Note TX("lg_due") & " 0": Exit Sub
    ' every holder of a due transaction is part of the e-mail
    Set txnHolders = CreateObject("Scripting.Dictionary")
    For Each k In due.Keys
        txnHolders(CStr(gRows(k)(R_TXN))) = True
    Next k
    Set uf = CreateObject("Scripting.Dictionary")
    Dim hold As Object
    Set hold = CreateObject("Scripting.Dictionary")
    For Each k In gRows.Keys
        a = gRows(k)
        If txnHolders.Exists(CStr(a(R_TXN))) And a(R_ACTIVE) = 1 And a(R_ARCH) <> 1 And _
           a(R_STATUS) <> "CLOSED" And a(R_STATUS) <> "FWD" Then
            If Not hold.Exists(CStr(a(R_TXN))) Then hold.Add CStr(a(R_TXN)), New Collection
            hold(CStr(a(R_TXN))).Add CStr(a(R_PERSON))
        End If
    Next k
    ' one e-mail per connected group: nobody gets two e-mails on the same day
    For Each txn In hold.Keys
        root = UF_Find(uf, CStr(hold(txn)(1)))
        For Each p In hold(txn)
            uf(UF_Find(uf, CStr(p))) = UF_Find(uf, root)
        Next p
    Next txn
    Set groups = CreateObject("Scripting.Dictionary")
    For Each txn In hold.Keys
        root = UF_Find(uf, CStr(hold(txn)(1)))
        If Not groups.Exists(root) Then groups.Add root, New Collection
        groups(root).Add CStr(txn)
    Next txn
    ' stop if any recipient has no e-mail address
    For Each txn In hold.Keys
        For Each p In hold(txn)
            If MetaEmail(CStr(p)) = "" Then missing = AddUnique(missing, p)
        Next p
    Next txn
    If missing <> "" Then
        Note TX("lg_missing") & " " & missing
        MsgBox "No drafts were created: some recipients have no e-mail address." & vbCrLf & _
               "See the red note at the top of the per-person sheet.", vbExclamation, "Murasalat"
        Exit Sub
    End If
    Set app = OL()
    If app Is Nothing Then Note TX("lg_noout"): Exit Sub
    For Each g In groups.Keys
        If MakeDraft(app, groups(g), hold) Then nMade = nMade + 1
    Next g
    Note TX("lg_due") & " " & groups.Count & " - " & TX("lg_drafts_n") & " " & nMade
End Sub

Private Function UF_Find(uf As Object, ByVal x As String) As String
    If Not uf.Exists(x) Then uf(x) = x
    Do While uf(x) <> x
        uf(x) = uf(uf(x))
        x = uf(x)
    Loop
    UF_Find = x
End Function

Private Function MakeDraft(app As Object, txns As Collection, hold As Object) As Boolean
    Dim people As String, peopleKey As String, txn As Variant, p As Variant, arr() As String, i As Long, j As Long
    Dim tmp As String, code As String, t As Variant, c As Variant, ns As Object, m As Object, old As Object
    Dim subj As String, html As String, toList As String, ccList As String, body As String, pos As Long, pos2 As Long
    ' sorted list of recipients = the thread identity
    For Each txn In txns
        For Each p In hold(txn)
            If Not InList(people, CStr(p)) Then people = AddToList(people, CStr(p))
        Next p
    Next txn
    arr = Split(people, "|")
    For i = 1 To UBound(arr)
        tmp = arr(i): j = i - 1
        Do While j >= 0
            If arr(j) <= tmp Then Exit Do
            arr(j + 1) = arr(j): j = j - 1
        Loop
        arr(j + 1) = tmp
    Next i
    peopleKey = Join(arr, "|")
    ' continue the open thread of these same people, or start a new one
    For Each c In gThreads.Keys
        t = gThreads(c)
        If CStr(t(T_PEOPLE)) = peopleKey And t(T_ENDED) <> 1 And ThreadOpen(t) Then code = CStr(c)
    Next c
    If code = "" Then
        i = 1
        Do While gThreads.Exists("MR-" & YMD(Date) & "-" & Right$("0" & i, 2)): i = i + 1: Loop
        code = "MR-" & YMD(Date) & "-" & Right$("0" & i, 2)
        ReDim t(1 To T_N)
        t(T_CODE) = code: t(T_PEOPLE) = peopleKey: t(T_TXNS) = "": t(T_CREATED) = Date: t(T_ENDED) = 0
    Else
        t = gThreads(code)
    End If
    For Each txn In txns
        If Not InList(t(T_TXNS), CStr(txn)) Then t(T_TXNS) = AddToList(t(T_TXNS), CStr(txn))
    Next txn
    subj = txns.Count & " " & IIf(txns.Count <= 2, TX("subj_one"), TX("subj_many")) & " " & ChrW(&H2014) & " " & _
           TX("subj_tail") & " (" & code & ")"
    html = MailHtml(txns, hold, arr, code)
    For i = 0 To UBound(arr)
        toList = toList & MetaEmail(arr(i)) & "; "
    Next i
    For i = 4 To 60
        If Trim(SH("shSet").Cells(i, 7).Value) <> "" Then ccList = ccList & Trim(SH("shSet").Cells(i, 7).Value) & "; "
    Next i
    Set ns = app.GetNamespace("MAPI")
    On Error Resume Next
    If CStr(t(T_DRAFT)) <> "" Then                       ' an unsent draft from before is replaced
        Set old = ns.GetItemFromID(t(T_DRAFT))
        If Not old Is Nothing Then
            If Not old.Sent Then old.Delete
        End If
    End If
    Set m = Nothing
    If CStr(t(T_LASTID)) <> "" Then                      ' Reply-All in the same thread
        Set m = ns.GetItemFromID(t(T_LASTID)).ReplyAll
        If Not m Is Nothing Then m.Subject = "RE: " & subj
    End If
    On Error GoTo 0
    If m Is Nothing Then
        Set m = app.CreateItem(0)
        m.Subject = subj
    End If
    Dim insp As Object
    On Error Resume Next
    Set insp = m.GetInspector                            ' adds your usual signature
    On Error GoTo 0
    m.To = toList
    m.CC = ccList
    body = m.HTMLBody
    pos = InStr(1, LCase$(body), "<body")
    If pos > 0 Then
        pos2 = InStr(pos, body, ">")
        m.HTMLBody = Left$(body, pos2) & html & "<br>" & Mid$(body, pos2 + 1)
    Else
        m.HTMLBody = html & body
    End If
    m.Save
    t(T_DRAFT) = m.EntryID
    t(T_LASTDRAFT) = Date
    t(T_SUBJ) = subj
    gThreads(code) = t
    MakeDraft = True
End Function

Private Function ThreadOpen(t As Variant) As Boolean
    Dim k As Variant, a As Variant
    For Each k In gRows.Keys
        a = gRows(k)
        If a(R_ACTIVE) = 1 And a(R_ARCH) <> 1 And a(R_STATUS) <> "CLOSED" And a(R_STATUS) <> "FWD" Then
            If InList(t(T_TXNS), CStr(a(R_TXN))) And InList(t(T_PEOPLE), CStr(a(R_PERSON))) Then ThreadOpen = True: Exit Function
        End If
    Next k
End Function

Private Function MailHtml(txns As Collection, hold As Object, people() As String, code As String) As String
    Dim sF As String, sP As String, sTD As String, sTH As String, h As String, txn As Variant, a As Variant
    Dim k As String, p As Variant, days As Long, remD As Long, nm As String, i As Long, out As String
    Dim cols As Variant, rows As String, x As Variant, depts As String, holders As String, lastDay As String
    sF = "font-family:Somar,Arial,sans-serif;"
    sP = "margin:0 0 10px 0;" & sF & "font-size:11pt;color:#000000;"
    sTD = "padding:6px 8px;border:1px solid #B8C3D1;vertical-align:top;" & sF & "font-size:10.5pt;color:#000000;"
    sTH = "padding:7px 8px;border:1px solid #152446;background:#152446;color:#FFFFFF;" & sF & "font-size:10.5pt;font-weight:normal;"
    cols = Array(TX("mc_1"), TX("mc_2"), TX("mc_3"), TX("mc_4"), TX("mc_5"), TX("mc_6") & "<br>" & TX("mail_fill"), _
                 TX("mc_7"), TX("mc_8"), TX("mc_9"), TX("mc_10"))
    For i = 0 To UBound(cols): h = h & "<th style='" & sTH & "'>" & cols(i) & "</th>": Next i
    For Each txn In txns
        holders = "": depts = "": days = 0
        For Each p In hold(txn)
            k = CStr(txn) & "|" & p
            a = gRows(k)
            holders = AddUnique(holders, p)
            depts = AddUnique(depts, a(R_DEPT))
            If RowDays(a) > days Then days = RowDays(a): x = a
        Next p
        If IsEmpty(x) Then x = gRows(CStr(txn) & "|" & hold(txn)(1))
        remD = gDl - days
        If remD = 1 Then lastDay = lastDay & " " & txn
        rows = rows & "<tr>" & _
            "<td style='" & sTD & "white-space:nowrap'>" & txn & "</td>" & _
            "<td style='" & sTD & "'>" & HtmlEsc(CStr(x(R_SUBJ))) & "</td>" & _
            "<td style='" & sTD & "white-space:nowrap'>" & TX("mail_type") & "</td>" & _
            "<td style='" & sTD & "white-space:nowrap'>" & FmtDate(DayOnly(x(R_CREATED))) & "</td>" & _
            "<td style='" & sTD & "white-space:nowrap'>" & FmtDate(DayOnly(x(R_FIRST))) & "</td>" & _
            "<td style='" & sTD & "background:#FFF8E1;min-width:90px'>&nbsp;</td>" & _
            "<td style='" & sTD & "'>" & holders & "</td>" & _
            "<td style='" & sTD & "'>" & depts & "</td>" & _
            "<td style='" & sTD & "white-space:nowrap'>" & TX("passed") & " " & days & " " & TX("days_word") & "</td>" & _
            "<td style='" & sTD & "white-space:nowrap;color:#C0392B'>" & _
                IIf(remD >= 0, TX("left") & " " & remD & " " & TX("days_word"), TX("over") & " " & (-remD) & " " & TX("days_word")) & _
            "</td></tr>"
        x = Empty
    Next txn
    If UBound(people) = 0 Then
        nm = Replace(TX("greet_single"), "{name}", PreferredName(people(0)))
    Else
        nm = TX("greet_group")
    End If
    out = "<div dir='rtl' style='text-align:right;" & sF & "font-size:11pt;color:#000000'>" & GradientBand() & _
        "<p style='" & sP & "'>" & nm & "<br>" & TX("mail_1") & "</p>" & _
        "<p style='" & sP & "'>" & Replace(TX("mail_2"), "{th}", CStr(gTh)) & "</p>" & _
        "<p style='" & sP & "'>" & TX("mail_3") & "</p>" & _
        "<p style='" & sP & "'>" & TX("mail_4") & "</p>"
    If lastDay <> "" Then
        out = out & "<p style='" & sP & "color:#C0392B;font-weight:bold;border:1px solid #C0392B;padding:6px 10px'>" & _
              TX("mail_alert") & IIf(txns.Count > 1, " (" & Trim$(lastDay) & ")", "") & "</p>"
    End If
    out = out & "<p style='" & sP & "font-weight:bold'>" & Replace(TX("mail_head"), "{th}", CStr(gTh)) & "</p>" & _
        "<table dir='rtl' style='border-collapse:collapse;margin:4px 0 14px 0'><tr>" & h & "</tr>" & rows & "</table>" & _
        "<p style='" & sP & "'>" & TX("mail_5") & "</p>" & _
        "<p style='" & sP & "'>" & TX("mail_6") & "<br>" & TX("mail_7") & "</p>" & _
        "<p style='margin:14px 0 0 0;" & sF & "font-size:8pt;color:#9AA3AF'>" & TX("mail_ref") & " " & code & "</p></div>"
    MailHtml = out
End Function

Private Function PreferredName(full As String) As String
    Dim pref As String, parts As Variant
    If gMeta.Exists(full) Then pref = gMeta(full)(2)
    If pref <> "" Then PreferredName = pref: Exit Function
    parts = Split(Application.WorksheetFunction.Trim(full), " ")
    If UBound(parts) >= 1 Then PreferredName = parts(0) & " " & parts(UBound(parts)) Else PreferredName = full
End Function

Private Function HtmlEsc(s As String) As String
    HtmlEsc = Replace(Replace(Replace(s, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
End Function

Private Function GradientBand() As String
    ' brand ribbon: Tiffany > navy > apple green, made of thin cells (Outlook can't draw CSS gradients)
    Dim stops As Variant, n As Long, k As Long, t As Double, i As Long, f As Double, col As String, s As String
    Dim r1 As Long, g1 As Long, b1 As Long, r2 As Long, g2 As Long, b2 As Long
    stops = Array("4BB8C4", "152446", "89A563")
    n = 48
    For k = 0 To n - 1
        t = k / (n - 1) * 2
        i = Int(t): If i > 1 Then i = 1
        f = t - i
        r1 = CLng("&H" & Mid$(stops(i), 1, 2)): g1 = CLng("&H" & Mid$(stops(i), 3, 2)): b1 = CLng("&H" & Mid$(stops(i), 5, 2))
        r2 = CLng("&H" & Mid$(stops(i + 1), 1, 2)): g2 = CLng("&H" & Mid$(stops(i + 1), 3, 2)): b2 = CLng("&H" & Mid$(stops(i + 1), 5, 2))
        col = Right$("0" & Hex(CLng(r1 + (r2 - r1) * f)), 2) & Right$("0" & Hex(CLng(g1 + (g2 - g1) * f)), 2) & _
              Right$("0" & Hex(CLng(b1 + (b2 - b1) * f)), 2)
        s = s & "<td height='2' bgcolor='#" & col & "' style='mso-line-height-rule:exactly;line-height:2px;font-size:1px;" & _
                "height:2px;background:#" & col & ";padding:0'>&#8203;</td>"
    Next k
    GradientBand = "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' border='0' " & _
        "style='border-collapse:collapse;width:100%'><tr>" & s & "</tr></table>" & _
        "<div style='height:12px;line-height:12px;font-size:1px'>&nbsp;</div>"
End Function

'============================================================================== one-time setup
Public Sub SetupButtons()
    ' run once after importing this module (Alt+F8 > SetupButtons > Run)
    Init
    PlaceButtons SH("shMain"), Array( _
        Array("b_import", "Btn_Import", "B7"), Array("b_drafts", "Btn_Drafts", "F7"), _
        Array("b_replies", "Btn_Replies", "B9"), Array("b_update", "Btn_Update", "F9"))
    PlaceButtons SH("shOpen"), Array(Array("b_move", "Btn_Move", "J3"))
    PlaceButtons SH("shPeople"), Array(Array("b_drafts", "Btn_Drafts", "C3"), Array("b_update", "Btn_Update", "H3"))
    PlaceButtons SH("shPaths"), Array(Array("b_update", "Btn_Update", "C3"))
    PlaceButtons SH("shClosed"), Array(Array("b_update", "Btn_Update", "J3"))
    MsgBox "Buttons are ready. You can start with the first button on the main sheet.", vbInformation, "Murasalat"
End Sub

Private Sub PlaceButtons(ws As Worksheet, defs As Variant)
    Dim sh As Shape, i As Long, d As Variant, c As Range, b As Shape
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 4) = "btn_" Then ws.Shapes(i).Delete
    Next i
    For i = 0 To UBound(defs)
        d = defs(i)
        Set c = ws.Range(d(2))
        Set b = ws.Shapes.AddShape(5, c.Left + 2, c.Top + 2, 230, 34)   ' rounded rectangle
        b.Name = "btn_" & d(1) & "_" & i
        b.Fill.ForeColor.RGB = RGB(205, 236, 242)
        b.Line.ForeColor.RGB = RGB(21, 36, 70)
        b.Line.Weight = 1.25
        With b.TextFrame2.TextRange
            .Text = TX(CStr(d(0)))
            .Font.Name = "Somar": .Font.Size = 12: .Font.Bold = True
            .Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
            .ParagraphFormat.Alignment = 2
        End With
        b.TextFrame2.VerticalAnchor = 3
        b.OnAction = d(1)
    Next i
End Sub
