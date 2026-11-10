Attribute VB_Name = "Projects"
Option Explicit

' =====================================================================
'  Project follow-up for the Human Capital agency
'  Creates the biweekly project-update drafts in the agency mailbox.
'  All Arabic wording lives in the texts sheet - nothing is hard-coded here.
' =====================================================================

' VBA stores a colour as &HBBGGRR, so the digits read backwards from the hex code
Private Const C_NAVY As Long = &H462415     ' #152446
Private Const C_TEAL As Long = &HC4B84B     ' #4BB8C4
Private Const C_GREEN As Long = &H63A589    ' #89A563
Private Const C_YELLOW As Long = &HE1F8FF   ' #FFF8E1

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
Private Const D_NUDGE As Long = 24      ' date of the last reminder sent
Private Const D_ATT As Long = 25        ' the reply that carried the PowerPoint
Private Const D_FIRST As Long = 6           ' first data row

Private gText As Collection
Private gReady As Boolean
Private gStep As String          ' last stage reached, shown if something goes wrong
Private gSheets As Collection    ' worksheets already located by their AZ1 tag
Private gFont As String          ' the message font, read once
Private gOurs As String          ' our own senders, so our copies never count
Private gBox As String           ' the agency mailbox, reduced to letters and digits

' =====================================================================
'  setup
' =====================================================================
Private Sub Fresh()
    ' Called at the top of every button so a text edited on the texts page shows up
    ' at once. It can never stop the run: every text has a value in DefaultText, so
    ' a failure here costs nothing and is swallowed.
    On Error Resume Next
    gReady = False
    gFont = ""
    Init
    Err.Clear
    On Error GoTo 0
End Sub

Private Sub Init()
    ' Reads the texts sheet into memory.
    '
    ' This used Scripting.Dictionary, which lives in a Windows library that locked
    ' down machines often block - and when it is blocked the whole module dies here
    ' with a bare "Automation error". Collection is part of VBA itself and cannot be
    ' taken away.
    Dim ws As Worksheet, r As Long, last As Long, k As String
    If gReady Then Exit Sub
    On Error GoTo Bail
    Set gText = New Collection
    Set ws = SH("shText")
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If last < 6 Then last = 6
    For r = 6 To last
        k = Trim$(CStr(ws.Cells(r, 2).Value))
        If k <> "" Then
            On Error Resume Next                 ' a repeated key: the first one wins
            gText.Add CStr(ws.Cells(r, 3).Value), k
            Err.Clear
            On Error GoTo Bail
        End If
    Next r
    gReady = True
    Exit Sub
Bail:
    ' Whatever went wrong, every text still has a value in DefaultText, so the module
    ' carries on reading from there rather than stopping.
    Err.Clear
    gReady = True
End Sub

Private Function SH(id As String) As Worksheet
    ' remembered per sheet: this is called from almost every routine, and walking
    ' every worksheet each time is a real cost inside the per-row loops
    Dim ws As Worksheet
    If gSheets Is Nothing Then Set gSheets = New Collection
    On Error Resume Next
    Set SH = gSheets(id)
    Err.Clear
    On Error GoTo 0
    If Not SH Is Nothing Then Exit Function
    For Each ws In ThisWorkbook.Worksheets
        If CStr(ws.Range("AZ1").Value) = id Then
            On Error Resume Next
            gSheets.Add ws, id
            Err.Clear
            On Error GoTo 0
            Set SH = ws
            Exit Function
        End If
    Next ws
    Err.Raise vbObjectError + 11, , "Sheet not found: " & id
End Function

Private Function TX(k As String) As String
    ' Every error handler calls this, so it must not be able to raise. If the texts
    ' could not be loaded it quietly falls back instead of failing inside a handler.
    Dim v As String
    On Error Resume Next
    If Not gReady Then Init
    Err.Clear
    If Not gText Is Nothing Then
        v = ""
        v = gText(k)                     ' raises if the key is not there; ignored
        Err.Clear
        If Trim$(v) <> "" Then
            On Error GoTo 0
            TX = v
            Exit Function
        End If
    End If
    Err.Clear
    On Error GoTo 0
    TX = DefaultText(k)
End Function

Private Function AW(ParamArray codes() As Variant) As String
    ' builds an Arabic string from character codes, so this file stays plain ASCII
    Dim i As Long
    For i = LBound(codes) To UBound(codes)
        AW = AW & ChrW(codes(i))
    Next i
End Function

Private Function DefaultText(k As String) As String
    ' Every text the module uses has a value here, so a newer module still works
    ' with an older workbook, and an edited or missing texts sheet cannot break it.
    '
    ' A long text is built by several statements rather than one long continued
    ' line: VBA allows only 24 continuations per statement, and the Arabic texts
    ' are far past that when written as character codes.
    Select Case k
        Case "app_title"
            DefaultText = "HC Projects"
        Case "dlg_who"
            DefaultText = "Is this you: {name}?"
        Case "dlg_who_sub"
            DefaultText = "Your departments: {depts}"
        Case "dlg_pick"
            DefaultText = "Who is using the file? Type the number next to your name."
        Case "dlg_confirm"
            DefaultText = "Drafts will be created for the departments below. Nothing is sent - " & "each one is saved in Drafts for you to review. Continue?"
        Case "dlg_to"
            DefaultText = "To"
        Case "dlg_cc"
            DefaultText = "Cc"
        Case "dlg_delay"
            DefaultText = "All of these will be delivered at the same moment"
        Case "dlg_delay_past"
            DefaultText = "NOTE: the shared delivery time has already passed today, so these " & "will go out as soon as you press Send."
        Case "dlg_done"
            DefaultText = "{n} draft(s) created in {box}. Review them before sending."
        Case "dlg_none"
            DefaultText = "No drafts were created. The log sheet says why."
        Case "dlg_nodept"
            DefaultText = "No department is assigned to you on the departments sheet."
        Case "dlg_no_outlook"
            DefaultText = "Outlook could not be reached. Open Outlook and try again."
        Case "dlg_no_store"
            DefaultText = "Your mailbox could not be opened."
        Case "dlg_rep_done"
            DefaultText = "{n} repl(ies) recorded."
        Case "dlg_rem_done"
            DefaultText = "A recurring reminder was added to your Outlook calendar, every " & "{weeks} week(s)."
        Case "dlg_rem_exists"
            DefaultText = "The recurring reminder is already in your calendar."
        Case "dlg_setup"
            DefaultText = "The buttons have been drawn on the main sheet."
        Case "dlg_link_pick"
            DefaultText = "Which department is this conversation for? Type the number."
        Case "dlg_link_none"
            DefaultText = "No message is open. Open the department thread in Outlook first, then " & "press this button."
        Case "dlg_link_noconv"
            DefaultText = "That message carries no conversation id, so it cannot be linked."
        Case "dlg_link_open"
            DefaultText = "Open the department thread in Outlook, select the newest message, " & "then press this button again."
        Case "dlg_link_done"
            DefaultText = "The conversation is now linked to {dept}."
        Case "err_general"
            DefaultText = "Something went wrong"
        Case "err_link"
            DefaultText = "Could not link the conversation"
        Case "err_remind"
            DefaultText = "Could not create the calendar reminder"
        Case "err_replies"
            DefaultText = "Could not check the replies"
        Case "err_row"
            DefaultText = "row"
        Case "dlg_nudge_confirm"
            DefaultText = "A reminder reply will be created on the same thread for the " & "departments below, which have not sent the deck yet. Continue?"
        Case "dlg_nudge_none"
            DefaultText = "None of your departments is waiting, so there is nobody to remind."
        Case "dlg_nudge_done"
            DefaultText = "{n} reminder draft(s) created in {box}. Review them before sending."
        Case "dlg_nudge_already"
            DefaultText = "already reminded today"
        Case "dlg_nudge_notsent"
            DefaultText = "Not sent yet, so there is nothing to remind about"
        Case "dlg_nudge_late"
            DefaultText = "NOTE: it is already past the deadline hour. The reminder asks for the " & "update before noon today, so it will arrive after the hour it names."
        Case "err_nudge"
            DefaultText = "Could not create the reminders"
        Case "dlg_unresolved"
            DefaultText = "Unknown recipient"
        Case "dlg_readonly"
            DefaultText = "This file is open read-only, so nothing would be recorded and the " & "round could be built twice. Close it, make sure no colleague has it " & "open, then open it again with write access."
        Case "dlg_test_noself"
            DefaultText = "Test mode: no address of your own to send the test to. Put your " & "e-mail in the optional address column of the team sheet."
        Case "dlg_trunc_list"
            DefaultText = "... the rest is not shown here. The log sheet lists every one."
        Case "dlg_rcpt_failed"
            DefaultText = "Could not read the recipients from the table"
        Case "mail_greet"
            DefaultText = AW(&H627, &H644, &H633, &H644, &H627, &H645) & " " & AW(&H639, &H644, &H64A, &H643, &H645) & " " & AW(&H648, &H631, &H62D, &H645, &H629) & " " & AW(&H627, &H644, &H644, &H647) & " " & AW(&H648, &H628, &H631, &H643, &H627, &H62A, &H647, &H60C)
        Case "mail_name"
            DefaultText = AW(&H645, &H639) & " " & AW(&H627, &H644, &H62A, &H62D, &H64A, &H629) & " " & AW(&H648, &H627, &H644, &H62A, &H642, &H62F, &H64A, &H631) & " {name}" & AW(&H60C)
        Case "mail_1"
            DefaultText = AW(&H646, &H623, &H645, &H644) & " " & AW(&H62F, &H639, &H645, &H643, &H645) & " " & AW(&H641, &H64A) & " " & AW(&H62A, &H62D, &H62F, &H64A, &H62B) & " " & AW(&H628, &H637, &H627, &H642, &H627, &H62A) & " " & AW(&H627, &H644, &H645, &H634, &H627, &H631, &H64A, &H639) & " " & AW(&H627, &H644, &H62E, &H627, &H635, &H629) & " " & AW(&H628, &H625, &H62F, &H627, &H631, &H62A, &H643, &H645) & " " & AW(&H644, &H627, &H62C, &H62A, &H645, &H627, &H639) & " " & AW(&H648, &H643, &H627, &H644, &H629) & " " & AW(&H631, &H623, &H633) & " " & AW(&H627, &H644, &H645, &H627, &H644) & " " & AW(&H627, &H644, &H628, &H634, &H631, &H64A) & "."
        Case "mail_b1"
            DefaultText = AW(&H646, &H623, &H645, &H644) & " " & AW(&H625, &H631, &H633, &H627, &H644) & " " & AW(&H627, &H644, &H628, &H637, &H627, &H642, &H627, &H62A) & " " & AW(&H642, &H628, &H644) & " " & AW(&H627, &H644, &H633, &H627, &H639, &H629) & " " & AW(&H661, &H662) & " " & AW(&H638, &H647, &H631, &H627, &H64B) & " " & AW(&H645, &H646) & " " & AW(&H64A, &H648, &H645) & " {due}."
        Case "mail_b2"
            DefaultText = AW(&H643, &H631, &H645, &H627, &H64B) & " " & AW(&H627, &H644, &H627, &H644, &H62A, &H632, &H627, &H645) & " " & AW(&H628, &H627, &H644, &H62A, &H639, &H62F, &H64A, &H644) & " " & AW(&H639, &H644, &H649) & " " & AW(&H627, &H644, &H646, &H645, &H648, &H630, &H62C) & " " & AW(&H627, &H644, &H645, &H631, &H641, &H642) & " " & AW(&H630, &H627, &H62A, &H647, &H60C) & " " & AW(&H648, &H645, &H631, &H627, &H639, &H627, &H629) & " " & AW(&H627, &H644, &H645, &H644, &H627, &H62D, &H638, &H627, &H62A) & " " & AW(&H627, &H644, &H645, &H633, &H62C, &H644, &H629) & "."
        Case "mail_b3"
            DefaultText = AW(&H646, &H623, &H645, &H644) & " " & AW(&H62A, &H648, &H636, &H64A, &H62D) & " " & AW(&H645, &H628, &H631, &H631, &H627, &H62A) & " " & AW(&H627, &H644, &H62A, &H623, &H62E, &H64A, &H631) & " " & AW(&H641, &H64A) & " " & AW(&H62D, &H627, &H644) & " " & AW(&H62A, &H639, &H62F, &H64A, &H644) & " " & AW(&H62A, &H627, &H631, &H64A, &H62E) & " " & AW(&H627, &H644, &H627, &H646, &H62A, &H647, &H627, &H621) & " " & AW(&H627, &H644, &H645, &H62A, &H648, &H642, &H639) & " " & AW(&H644, &H644, &H645, &H634, &H631, &H648, &H639, &H60C) & " " & AW(&H645, &H639) & " " & AW(&H623, &H647, &H645, &H64A, &H629) & " " & AW(&H62A, &H62D, &H62F, &H64A, &H62F) & " " & _
                AW(&H62A, &H627, &H631, &H64A, &H62E) & " " & AW(&H645, &H633, &H62A, &H647, &H62F, &H641) & " " & AW(&H644, &H644, &H627, &H646, &H62A, &H647, &H627, &H621, &H60C) & " " & AW(&H648, &H630, &H644, &H643) & " " & AW(&H644, &H636, &H645, &H627, &H646) & " " & AW(&H62F, &H642, &H629) & " " & AW(&H627, &H644, &H645, &H62A, &H627, &H628, &H639, &H629) & " " & AW(&H648, &H639, &H62F, &H645) & " " & AW(&H62A, &H623, &H62B, &H631) & " " & AW(&H645, &H624, &H634, &H631, &H627, &H62A) & " " & AW(&H62C, &H648, &H62F, &H629) & " " & AW(&H62A, &H633, &H644, &H64A, &H645) & " " & AW(&H627, &H644, &H645, &H634, &H627, &H631, &H64A, &H639) & "."
        Case "mail_note"
            DefaultText = AW(&H644, &H644, &H62A, &H646, &H648, &H64A, &H647) & ": " & AW(&H633, &H64A, &H62A, &H645) & " " & AW(&H62A, &H637, &H628, &H64A, &H642) & " " & AW(&H62A, &H642, &H64A, &H64A, &H645) & " " & AW(&H62C, &H648, &H62F, &H629) & " " & AW(&H62A, &H633, &H644, &H64A, &H645) & " " & AW(&H627, &H644, &H645, &H634, &H627, &H631, &H64A, &H639) & " " & AW(&H639, &H644, &H649) & " " & AW(&H62C, &H645, &H64A, &H639) & " " & AW(&H645, &H645, &H62B, &H644, &H64A) & " " & AW(&H648, &H643, &H627, &H644, &H629) & " " & AW(&H631, &H623, &H633) & " " & AW(&H627, &H644, &H645, &H627, &H644) & " " & AW(&H627, &H644, &H628, &H634, &H631, &H64A) & " " & AW(&H648, &H641, &H642) & " " & _
                AW(&H627, &H644, &H646, &H645, &H648, &H630, &H62C) & " " & AW(&H627, &H644, &H645, &H631, &H641, &H642) & "."
        Case "mail_end"
            DefaultText = AW(&H645, &H639) & " " & AW(&H62C, &H632, &H64A, &H644) & " " & AW(&H627, &H644, &H634, &H643, &H631) & " " & AW(&H648, &H627, &H644, &H62A, &H642, &H62F, &H64A, &H631, &H60C)
        Case "mail_sign"
            DefaultText = AW(&H625, &H62F, &H627, &H631, &H629) & " " & AW(&H623, &H62F, &H627, &H621) & " " & AW(&H648, &H643, &H627, &H644, &H629) & " " & AW(&H631, &H623, &H633) & " " & AW(&H627, &H644, &H645, &H627, &H644) & " " & AW(&H627, &H644, &H628, &H634, &H631, &H64A)
        Case "nudge_1"
            DefaultText = AW(&H644, &H644, &H62A, &H630, &H643, &H64A, &H631) & " " & AW(&H628, &H62A, &H632, &H648, &H64A, &H62F, &H646, &H627) & " " & AW(&H628, &H62A, &H62D, &H62F, &H64A, &H62B) & " " & AW(&H628, &H637, &H627, &H642, &H627, &H62A) & " " & AW(&H627, &H644, &H645, &H634, &H627, &H631, &H64A, &H639) & " " & AW(&H627, &H644, &H62E, &H627, &H635, &H629) & " " & AW(&H628, &H625, &H62F, &H627, &H631, &H62A, &H643, &H645) & " " & AW(&H644, &H627, &H62C, &H62A, &H645, &H627, &H639) & " " & AW(&H648, &H643, &H627, &H644, &H629) & " " & AW(&H631, &H623, &H633) & " " & AW(&H627, &H644, &H645, &H627, &H644) & " " & AW(&H627, &H644, &H628, &H634, &H631, &H64A, &H60C) & " " & _
                AW(&H648, &H630, &H644, &H643) & " " & AW(&H642, &H628, &H644) & " " & AW(&H627, &H644, &H633, &H627, &H639, &H629) & " " & AW(&H661, &H662) & " " & AW(&H638, &H647, &H631, &H627, &H64B) & " " & AW(&H645, &H646) & " " & AW(&H627, &H644, &H64A, &H648, &H645) & "."
        Case "nudge_end"
            DefaultText = AW(&H634, &H627, &H643, &H631, &H64A, &H646) & " " & AW(&H644, &H643, &H645) & " " & AW(&H62C, &H647, &H648, &H62F, &H643, &H645, &H60C)
        Case "w_all"
            DefaultText = AW(&H64A, &H636, &H627, &H641) & " " & AW(&H644, &H62C, &H645, &H64A, &H639) & " " & AW(&H627, &H644, &H625, &H62F, &H627, &H631, &H627, &H62A)
        Case "w_hold"
            DefaultText = AW(&H645, &H648, &H642, &H648, &H641)
        Case "w_test"
            DefaultText = AW(&H62A, &H62C, &H631, &H64A, &H628, &H64A)
        Case "w_draft"
            DefaultText = AW(&H645, &H633, &H648, &H62F, &H629)
        Case "w_delegate"
            DefaultText = AW(&H62A, &H641, &H648, &H64A, &H636)
        Case "w_linked"
            DefaultText = AW(&H645, &H631, &H62A, &H628, &H637, &H629)
        Case "w_notlinked"
            DefaultText = AW(&H63A, &H64A, &H631) & " " & AW(&H645, &H631, &H62A, &H628, &H637, &H629)
        Case "w_nudge"
            DefaultText = AW(&H62A, &H630, &H643, &H64A, &H631)
        Case "w_round"
            DefaultText = AW(&H627, &H644, &H62C, &H648, &H644, &H629)
        Case "w_cur"
            DefaultText = AW(&H627, &H644, &H62C, &H648, &H644, &H629) & " " & AW(&H627, &H644, &H62D, &H627, &H644, &H64A, &H629)
        Case "lbl_unknown"
            DefaultText = AW(&H63A, &H64A, &H631) & " " & AW(&H645, &H639, &H631, &H648, &H641, &H629)
        Case "lbl_none"
            DefaultText = AW(&H644, &H627) & " " & AW(&H62A, &H648, &H62C, &H62F) & " " & AW(&H625, &H62F, &H627, &H631, &H627, &H62A) & " " & AW(&H645, &H633, &H646, &H62F, &H629)
        Case "lbl_updated"
            DefaultText = AW(&H622, &H62E, &H631) & " " & AW(&H62A, &H62D, &H62F, &H64A, &H62B) & ": {date}"
        Case "hdr_mail"
            DefaultText = AW(&H627, &H644, &H628, &H631, &H64A, &H62F) & " " & AW(&H623, &H648) & " " & AW(&H627, &H644, &H627, &H633, &H645)
        Case "hdr_kind"
            DefaultText = AW(&H627, &H644, &H646, &H648, &H639)
        Case "hdr_note"
            DefaultText = AW(&H645, &H644, &H627, &H62D, &H638, &H629)
        Case "hdr_nudge"
            DefaultText = AW(&H62A, &H627, &H631, &H64A, &H62E) & " " & AW(&H627, &H644, &H62A, &H630, &H643, &H64A, &H631)
        Case "rem_title"
            DefaultText = AW(&H625, &H631, &H633, &H627, &H644) & " " & AW(&H62A, &H62D, &H62F, &H64A, &H62B) & " " & AW(&H627, &H644, &H645, &H634, &H627, &H631, &H64A, &H639) & " " & AW(&H644, &H644, &H625, &H62F, &H627, &H631, &H627, &H62A)
        Case "rem_body"
            DefaultText = AW(&H625, &H646, &H634, &H627, &H621) & " " & AW(&H645, &H633, &H648, &H62F, &H627, &H62A) & " " & AW(&H62A, &H62D, &H62F, &H64A, &H62B) & " " & AW(&H627, &H644, &H645, &H634, &H627, &H631, &H64A, &H639) & " " & AW(&H645, &H646) & " " & AW(&H645, &H644, &H641) & " " & AW(&H627, &H644, &H645, &H62A, &H627, &H628, &H639, &H629, &H60C) & " " & AW(&H62B, &H645) & " " & AW(&H645, &H631, &H627, &H62C, &H639, &H62A, &H647, &H627) & " " & AW(&H648, &H625, &H631, &H633, &H627, &H644, &H647, &H627) & "."
        Case "sep"
            DefaultText = " " & AW(&HB7) & " "
        Case "b_start"
            DefaultText = AW(&H625, &H646, &H634, &H627, &H621) & " " & AW(&H627, &H644, &H645, &H633, &H648, &H62F, &H627, &H62A)
        Case "b_link"
            DefaultText = AW(&H631, &H628, &H637) & " " & AW(&H645, &H62D, &H627, &H62F, &H62B, &H629) & " " & AW(&H625, &H62F, &H627, &H631, &H629)
        Case "b_replies"
            DefaultText = AW(&H62A, &H62D, &H62F, &H64A, &H62B) & " " & AW(&H62D, &H627, &H644, &H629) & " " & AW(&H627, &H644, &H631, &H62F, &H648, &H62F)
        Case "b_remind"
            DefaultText = AW(&H625, &H646, &H634, &H627, &H621) & " " & AW(&H62A, &H630, &H643, &H64A, &H631) & " " & AW(&H645, &H62A, &H643, &H631, &H631)
        Case "b_refresh"
            DefaultText = AW(&H62A, &H62D, &H62F, &H64A, &H62B) & " " & AW(&H627, &H644, &H635, &H641, &H62D, &H627, &H62A)
        Case "b_nudge"
            DefaultText = AW(&H62A, &H630, &H643, &H64A, &H631) & " " & AW(&H63A, &H64A, &H631) & " " & AW(&H627, &H644, &H645, &H62C, &H64A, &H628, &H64A, &H646)
        Case "st_replied"
            DefaultText = AW(&H631, &H62F, &H651, &H62A)
        Case "st_repnoatt"
            DefaultText = AW(&H631, &H62F, &H651, &H62A) & " " & AW(&H628, &H62F, &H648, &H646) & " " & AW(&H645, &H631, &H641, &H642)
        Case "st_waiting"
            DefaultText = AW(&H628, &H627, &H646, &H62A, &H638, &H627, &H631) & " " & AW(&H627, &H644, &H631, &H62F)
        Case "st_duenudge"
            DefaultText = AW(&H62A, &H633, &H62A, &H62D, &H642) & " " & AW(&H627, &H644, &H62A, &H630, &H643, &H64A, &H631)
        Case "st_noreply"
            DefaultText = AW(&H644, &H645) & " " & AW(&H62A, &H631, &H62F, &H651)
        Case "st_hold"
            DefaultText = AW(&H645, &H648, &H642, &H648, &H641, &H629)
        Case "st_notsent"
            DefaultText = AW(&H644, &H645) & " " & AW(&H62A, &H64F, &H631, &H633, &H644) & " " & AW(&H628, &H639, &H62F)
        Case "h_dept"
            DefaultText = AW(&H627, &H644, &H625, &H62F, &H627, &H631, &H629)
        Case "h_rep"
            DefaultText = AW(&H627, &H644, &H645, &H645, &H62B, &H644)
        Case "h_reply"
            DefaultText = AW(&H627, &H644, &H631, &H62F)
        Case "h_days"
            DefaultText = AW(&H623, &H64A, &H627, &H645) & " " & AW(&H627, &H644, &H639, &H645, &H644)
        Case "h_nudge"
            DefaultText = AW(&H627, &H644, &H62A, &H630, &H643, &H64A, &H631)
        Case "h_state"
            DefaultText = AW(&H627, &H644, &H62D, &H627, &H644, &H629)
        Case "overdue"
            DefaultText = AW(&H641, &H627, &H62A) & " " & AW(&H645, &H648, &H639, &H62F) & " " & AW(&H627, &H644, &H62C, &H648, &H644, &H629)
        Case "m_1"
            DefaultText = AW(&H64A, &H646, &H627, &H64A, &H631)
        Case "m_2"
            DefaultText = AW(&H641, &H628, &H631, &H627, &H64A, &H631)
        Case "m_3"
            DefaultText = AW(&H645, &H627, &H631, &H633)
        Case "m_4"
            DefaultText = AW(&H623, &H628, &H631, &H64A, &H644)
        Case "m_5"
            DefaultText = AW(&H645, &H627, &H64A, &H648)
        Case "m_6"
            DefaultText = AW(&H64A, &H648, &H646, &H64A, &H648)
        Case "m_7"
            DefaultText = AW(&H64A, &H648, &H644, &H64A, &H648)
        Case "m_8"
            DefaultText = AW(&H623, &H63A, &H633, &H637, &H633)
        Case "m_9"
            DefaultText = AW(&H633, &H628, &H62A, &H645, &H628, &H631)
        Case "m_10"
            DefaultText = AW(&H623, &H643, &H62A, &H648, &H628, &H631)
        Case "m_11"
            DefaultText = AW(&H646, &H648, &H641, &H645, &H628, &H631)
        Case "m_12"
            DefaultText = AW(&H62F, &H64A, &H633, &H645, &H628, &H631)
        Case "dw_1"
            DefaultText = AW(&H627, &H644, &H623, &H62D, &H62F)
        Case "dw_2"
            DefaultText = AW(&H627, &H644, &H627, &H62B, &H646, &H64A, &H646)
        Case "dw_3"
            DefaultText = AW(&H627, &H644, &H62B, &H644, &H627, &H62B, &H627, &H621)
        Case "dw_4"
            DefaultText = AW(&H627, &H644, &H623, &H631, &H628, &H639, &H627, &H621)
        Case "dw_5"
            DefaultText = AW(&H627, &H644, &H62E, &H645, &H64A, &H633)
        Case "dw_6"
            DefaultText = AW(&H627, &H644, &H62C, &H645, &H639, &H629)
        Case "dw_7"
            DefaultText = AW(&H627, &H644, &H633, &H628, &H62A)
        Case "d_mon"
            DefaultText = AW(&H627, &H644, &H627, &H62B, &H646, &H64A, &H646)
        Case "dlg_where"
            DefaultText = "The drafts are saved in: {box}"
        Case "dlg_delay_none"
            DefaultText = "NOTE: no shared delivery time is set, so these will go out as soon as " & "you press Send."
        Case "dlg_truncated"
            DefaultText = "Only the first 60 departments are shown."
        Case "dlg_skipped"
            DefaultText = "SKIPPED"
        Case "dlg_skipped_n"
            DefaultText = "Skipped:"
        Case "dlg_see_log"
            DefaultText = "The log sheet says why for each one."
        Case "dlg_test_goes"
            DefaultText = "TEST ROW - this draft goes to YOU only, with no Cc at all. Set the " & "status to active to see the real To and Cc."
        Case "dlg_nothread"
            DefaultText = "No conversation found for this department. Check the subject on the " & "departments sheet, or link the thread by hand."
        Case "dlg_nomail"
            DefaultText = "No e-mail address for the representative."
        Case "dlg_thread_gone"
            DefaultText = "The conversation could not be opened."
        Case "dlg_reply_failed"
            DefaultText = "Reply All failed"
        Case "dlg_history_rebuilt"
            DefaultText = "INCOMPLETE - the quoted history was lost, so this draft needs " & "checking."
        Case "dlg_subj_differs"
            DefaultText = "the subject differs from the one on the sheet -"
        Case Else: DefaultText = ""
    End Select
End Function

Private Function SetVal(r As Long) As Variant
    SetVal = SH("shSet").Cells(r, 3).Value
End Function

Private Function FontName() As String
    ' Quoted, because a name with a space ("Somar Bold") is invalid in CSS unquoted
    ' and the whole font-family rule would then be thrown away by the mail client.
    Dim v As String
    If gFont <> "" Then
        FontName = gFont
        Exit Function
    End If
    v = Trim$(CStr(Nz(SetVal(10), "")))
    If v = "" Then v = "Somar"
    gFont = "'" & Replace(v, "'", "") & "'"
    FontName = gFont
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
    If r < 1 Then Exit Function          ' Cells(0, 4) would raise
    Set ws = SH("shTeam")
    TeamEN = Trim$(CStr(ws.Cells(r, 4).Value))
    If TeamEN = "" Then TeamEN = "User " & (r - 5)
End Function

Private Function Blank(v As Variant) As Boolean
    ' CStr on an error value raises, and VBA's Or evaluates both sides, so the
    ' error case has to be taken out before the comparison
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then
        Blank = True
    Else
        Blank = (Trim$(CStr(v)) = "")
    End If
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
    Dim en As Long, ed As String
    Dim ws As Worksheet, s As Shape, i As Long, nm As String
    On Error GoTo Failed
    gStep = "setup"
    Fresh
    Application.ScreenUpdating = False
    Set ws = SH("shMain")
    For i = ws.Shapes.Count To 1 Step -1
        nm = ""
        On Error Resume Next
        nm = CStr(ws.Shapes(i).Name)
        Err.Clear
        On Error GoTo Failed
        If Left$(nm, 4) = "btn_" Or Left$(nm, 6) = "brand_" Then ws.Shapes(i).Delete
    Next i
    ' two rows, the everyday ones on top. Rows 11 and 13 are left tall for them.
    AddButton ws, TX("b_start"), "Btn_Start", 1, 1, True
    AddButton ws, TX("b_replies"), "Btn_Replies", 1, 2, False
    AddButton ws, TX("b_nudge"), "Btn_Nudge", 1, 3, False
    AddButton ws, TX("b_link"), "Btn_Link", 2, 1, False
    AddButton ws, TX("b_remind"), "Btn_Remind", 2, 2, False
    AddButton ws, TX("b_refresh"), "Btn_Refresh", 2, 3, False
    DrawGradient ws
    Refresh "setup"
    Application.ScreenUpdating = True
    MsgBox TX("dlg_setup"), vbInformation, TX("app_title")
    Exit Sub
Failed:
    en = Err.Number: ed = Err.Description     ' read first - Say and TX both clear it
    Application.ScreenUpdating = True
    MsgBox TX("err_general") & " [" & gStep & "] " & en & ": " & ed, _
           vbExclamation, TX("app_title")
End Sub

Private Sub AddButton(ws As Worksheet, label As String, macro As String, _
                      rowIdx As Long, colIdx As Long, primary As Boolean)
    Dim s As Shape, t As Double, l As Double, w As Double, h As Double, gap As Double
    Dim first As Double
    h = 34: gap = 10
    first = 190                                   ' the wide primary button
    If primary Then
        w = first
    Else
        w = 150
    End If
    If rowIdx = 1 Then
        t = ws.Range("B11").Top
    Else
        t = ws.Range("B13").Top
    End If
    l = ws.Range("B11").Left
    If colIdx > 1 Then
        ' the first slot on row one is the wide button; on the other rows it is not
        If rowIdx = 1 Then
            l = l + first + gap + (colIdx - 2) * (150 + gap)
        Else
            l = l + (colIdx - 1) * (150 + gap)
        End If
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
        If LCase$(Trim$(CStr(ws.Cells(r, 3).Value))) = u Then
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
    ' No default. It used to offer "1", so pressing Enter picked whoever happens to be
    ' first in the team sheet - and the username was then remembered against her name.
    ans = InputBox(TX("dlg_pick") & vbCrLf & vbCrLf & list, TX("app_title"), "")
    If Trim$(ans) = "" Then Exit Function
    If Not IsNumeric(ans) Then Exit Function
    i = CLng(ans)
    If i < 1 Or i > n Then Exit Function
    ' The username is written into the shared file, so the pick is confirmed once.
    If MsgBox(Replace(TX("dlg_who"), "{name}", TeamEN(TeamRowOfName(names(i)))) & vbCrLf & vbCrLf & _
              Replace(TX("dlg_who_sub"), "{depts}", DeptsOfEN(names(i))), _
              vbYesNo + vbQuestion, TX("app_title")) <> vbYes Then Exit Function
    AskWho = names(i)
    RememberUser names(i)
End Function

Private Function TeamRowOfName(person As String) As Long
    Dim ws As Worksheet, r As Long
    Set ws = SH("shTeam")
    For r = 6 To ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        If Norm(CStr(ws.Cells(r, 2).Value)) = Norm(person) Then
            TeamRowOfName = r
            Exit Function
        End If
    Next r
End Function

Private Sub RememberUser(person As String)
    ' The four of you share one file, and only one username was filled in by hand.
    ' The first time somebody picks her own name, her Windows username is written
    ' next to it, so from then on the file knows her without asking. An existing
    ' username is never overwritten, and a username already used on another row is
    ' not copied anywhere: a wrong guess here would offer somebody else's departments.
    Dim ws As Worksheet, r As Long, u As String, last As Long
    On Error Resume Next
    u = LCase$(Trim$(Environ$("USERNAME")))
    If u = "" Then Exit Sub
    Set ws = SH("shTeam")
    If ws Is Nothing Then Exit Sub
    last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = 6 To last
        If LCase$(Trim$(CStr(ws.Cells(r, 3).Value))) = u Then
            If Norm(CStr(ws.Cells(r, 2).Value)) = Norm(person) Then Exit Sub  ' right row
            ' It landed on somebody else's row - a mistyped number the first time.
            ' Clearing it lets the correct row take it now, instead of greeting her
            ' by a colleague's name, with a colleague's departments, for good.
            ws.Cells(r, 3).ClearContents
        End If
    Next r
    For r = 6 To last
        If Norm(CStr(ws.Cells(r, 2).Value)) = Norm(person) Then
            If Trim$(CStr(ws.Cells(r, 3).Value)) = "" Then
                ws.Cells(r, 3).Value = Trim$(Environ$("USERNAME"))
                SaveBook
            End If
            Exit Sub
        End If
    Next r
    Err.Clear
End Sub

' =====================================================================
'  dates
' =====================================================================
Private Function ShortDate(d As Date) As String
    ' day and month only - the year adds nothing to a deadline four days away
    ShortDate = Day(d) & " " & TX("m_" & Month(d))
End Function

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

Private Function NextWorkday(afterD As Date) As Date
    ' the first working day after a date: Friday and Saturday are skipped
    Dim d As Date, guard As Long
    d = afterD + 1
    Do While Weekday(d, vbSunday) = vbFriday Or Weekday(d, vbSunday) = vbSaturday
        d = d + 1
        guard = guard + 1
        If guard > 7 Then Exit Do
    Loop
    NextWorkday = d
End Function

Private Function DayWord(d As Date) As String
    ' the Arabic day name, so a reminder can name its own deadline day
    DayWord = TX("dw_" & Weekday(d, vbSunday))
End Function

Private Function StepWeeks() As Long
    ' how many weeks between rounds, kept inside sane limits whatever the cell holds
    On Error Resume Next
    StepWeeks = CLng(Nz(SetVal(12), 2))
    If Err.Number <> 0 Then
        Err.Clear
        StepWeeks = 2
    End If
    On Error GoTo 0
    If StepWeeks < 1 Then StepWeeks = 1
    If StepWeeks > 26 Then StepWeeks = 26
End Function

Private Function LastSent() As Date
    ' the most recent round anyone recorded
    Dim ws As Worksheet, r As Long, d As Date
    Set ws = SH("shDept")
    On Error Resume Next
    For r = D_FIRST To LastDeptRow()
        If Not Blank(ws.Cells(r, D_SENT).Value) Then
            If CDate(ws.Cells(r, D_SENT).Value) > d Then d = CDate(ws.Cells(r, D_SENT).Value)
        End If
        Err.Clear
    Next r
    On Error GoTo 0
    LastSent = d
End Function

Private Function DueThursday() As Date
    ' the next sending Thursday, counting in steps from the first round
    Dim start As Date, step_ As Long, d As Date
    step_ = StepWeeks()
    If Blank(SetVal(13)) Then
        d = Date
        Do While Weekday(d, vbSunday) <> vbThursday
            d = d + 1
        Loop
        DueThursday = d
        Exit Function
    End If
    start = Date
    On Error Resume Next
    start = CDate(SetVal(13))
    Err.Clear
    On Error GoTo 0
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

Private Sub Say(t As String)
    ' a short note on Excel's status bar, so a slow step never looks like a freeze
    On Error Resume Next
    If t = "" Then
        Application.StatusBar = False
    Else
        Application.StatusBar = t
    End If
    Err.Clear
End Sub

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
    ' Appends one or more recipients to a "a; b; c" list, skipping duplicates.
    '
    ' Semicolons always separate. A comma separates only inside a piece that holds
    ' real addresses: "Smith, John" is ONE person, and treating its comma as a
    ' separator produced two names that resolve to nobody - or, worse, to somebody
    ' else entirely.
    Dim outer As Variant, inner As Variant, i As Long, j As Long, a As String, b As String
    AddAddr = cur
    a = Trim$(addr)
    If a = "" Then Exit Function
    outer = Split(a, ";")
    For i = 0 To UBound(outer)
        a = Trim$(outer(i))
        If InStr(a, ",") > 0 And InStr(a, "@") > 0 Then a = Replace(a, ",", ";")
        inner = Split(a, ";")
        For j = 0 To UBound(inner)
            b = Trim$(inner(j))
            If b <> "" Then
                If InStr(1, ";" & Norm(AddAddr) & ";", ";" & Norm(b) & ";") = 0 Then
                    If AddAddr <> "" Then AddAddr = AddAddr & "; "
                    AddAddr = AddAddr & b
                End If
            End If
        Next j
    Next i
End Function

Private Function Delegated(r As Long) As Boolean
    Dim ws As Worksheet, untilD As Variant
    Set ws = SH("shDept")
    ' An address, not a name. With only the name filled in, this used to report a
    ' delegation, and To was then built from the empty address cell while the real
    ' representative was demoted to Cc - so the mail went to nobody in particular.
    If Blank(ws.Cells(r, D_ALTMAIL).Value) Then Exit Function
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
    ' only a real address goes in To - a bare name would resolve to an unreadable
    ' Exchange address, which is what used to show up as "/o=First Organization/ou=..."
    If Delegated(r) Then
        out = AddAddr("", CStr(Nz(ws.Cells(r, D_ALTMAIL).Value, "")))
    Else
        out = AddAddr("", CStr(Nz(ws.Cells(r, D_MAIL).Value, "")))
    End If
    out = AddAddr(out, ListFor(CStr(ws.Cells(r, D_NAME).Value), "To"))
    out = AddAddr(out, CStr(ws.Cells(r, D_XTO).Value))
    ToLine = out
End Function

Private Function CcLine(r As Long) As String
    Dim ws As Worksheet, out As String
    Set ws = SH("shDept")
    If Delegated(r) Then
        out = AddAddr("", CStr(Nz(ws.Cells(r, D_MAIL).Value, "")))
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
    ' The round message. The notice sits in a tinted box with a coloured edge, because
    ' it carries a consequence for the reader and has to survive a quick skim.
    Dim sF As String, sP As String, sB As String, nm As String, out As String, due As String
    sF = "font-family:" & FontName() & ",Arial,sans-serif;"
    sP = "margin:0 0 12px 0;" & sF & "font-size:11pt;color:#000000;"
    sB = "margin:0 0 8px 0;" & sF & "font-size:11pt;color:#000000;"
    nm = CallName(r)
    due = "<span style='color:#152446;font-weight:bold'>" & DayWord(dueD) & " " & ShortDate(dueD) & "</span>"
    out = "<div dir='rtl' style='text-align:right;" & sF & "font-size:11pt;color:#000000'>" & GradientBand()
    out = out & "<p style='" & sP & "'>" & TX("mail_greet") & "</p>"
    If nm <> "" Then
        out = out & "<p style='" & sP & "'>" & Replace(TX("mail_name"), "{name}", HtmlEsc(nm)) & "</p>"
    Else
        out = out & "<p style='" & sP & "'>" & Replace(TX("mail_name"), "{name}", "") & "</p>"
    End If
    out = out & "<p style='" & sP & "'>" & TX("mail_1") & "</p>"
    out = out & "<ul style='margin:0 0 16px 0;padding-right:24px;" & sF & "'>" & _
        "<li style='" & sB & "'>" & Replace(TX("mail_b1"), "{due}", due) & "</li>" & _
        "<li style='" & sB & "'>" & TX("mail_b2") & "</li>" & _
        "<li style='" & sB & "'>" & TX("mail_b3") & "</li></ul>"
    out = out & NoticeBox(TX("mail_note"))
    out = out & "<p style='" & sP & "margin-top:16px'>" & TX("mail_end") & "</p>"
    out = out & "<p style='" & sP & "margin-bottom:0;font-weight:bold'>" & TX("mail_sign") & "</p></div>"
    MailHtml = out
End Function

Private Function NoticeBox(txt As String) As String
    ' A table, not a styled div: Outlook ignores most box styling on a div, but it
    ' renders table cell backgrounds and borders reliably.
    Dim sF As String
    sF = "font-family:" & FontName() & ",Arial,sans-serif;"
    NoticeBox = "<table role='presentation' width='100%' cellpadding='0' cellspacing='0' " & _
        "border='0' style='border-collapse:collapse;width:100%'><tr>" & _
        "<td style='background:#F4FAFB;border-right:4px solid #4BB8C4;padding:12px 14px;" & _
        sF & "font-size:11pt;color:#152446;font-weight:bold'>" & txt & "</td></tr></table>"
End Function

Private Function NudgeHtml(r As Long, dueD As Date) As String
    ' The reminder: three short lines over the untouched thread. It carries no date -
    ' it is sent on the deadline day itself and asks for the update the same day.
    Dim sF As String, sP As String, nm As String, out As String
    sF = "font-family:" & FontName() & ",Arial,sans-serif;"
    sP = "margin:0 0 12px 0;" & sF & "font-size:11pt;color:#000000;"
    nm = CallName(r)
    out = "<div dir='rtl' style='text-align:right;" & sF & "font-size:11pt;color:#000000'>" & GradientBand()
    If nm <> "" Then
        out = out & "<p style='" & sP & "'>" & Replace(TX("mail_name"), "{name}", HtmlEsc(nm)) & "</p>"
    Else
        out = out & "<p style='" & sP & "'>" & Replace(TX("mail_name"), "{name}", "") & "</p>"
    End If
    out = out & "<p style='" & sP & "'>" & TX("nudge_1") & "</p>"
    out = out & "<p style='" & sP & "'>" & TX("nudge_end") & "</p>"
    out = out & "<p style='" & sP & "margin-bottom:0;font-weight:bold'>" & TX("mail_sign") & "</p></div>"
    NudgeHtml = out
End Function

' =====================================================================
'  main action: create the drafts
' =====================================================================
Public Sub Btn_Start()
    Dim en As Long, ed As String
    Dim who As String, ws As Worksheet, r As Long, rowsSel() As Long, n As Long
    Dim i As Long, msg As String, app As Object, reason As String, reasons() As String
    Dim made As Long, skipped As Long, dueD As Date, warn As String, st As String
    On Error GoTo Failed
    gStep = "init"
    Fresh
    If ReadOnlyStop() Then Exit Sub
    Set ws = SH("shDept")
    gStep = "who"
    who = AskWho()
    If who = "" Then Exit Sub
    dueD = NextMonday(Date)

    ' 1) the departments that belong to this person and are not on hold
    n = 0
    ReDim rowsSel(1 To 60)
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
    If n = 60 Then warn = warn & vbCrLf & TX("dlg_truncated")

    ' 2) review screen - a department that cannot be done properly is named and skipped
    gStep = "review"
    msg = ""
    If IsEmpty(DeliverAt()) Then
        msg = TX("dlg_delay_none") & vbCrLf & vbCrLf
    ElseIf CDate(DeliverAt()) > Now Then
        msg = TX("dlg_delay") & ": " & Format$(CDate(DeliverAt()), "hh:nn AM/PM") & vbCrLf & vbCrLf
    Else
        msg = TX("dlg_delay_past") & vbCrLf & vbCrLf        ' said first, not buried
    End If
    msg = msg & TX("dlg_confirm") & vbCrLf
    ' Every reason is worked out for EVERY department first. The list shown below is
    ' capped, and a department whose reason was never worked out would be attempted
    ' later instead of skipped.
    ReDim reasons(1 To n)
    For i = 1 To n
        reasons(i) = BlockedReason(rowsSel(i))
    Next i
    For i = 1 To n
        r = rowsSel(i)
        reason = reasons(i)
        msg = msg & vbCrLf & "- " & DeptEN(r)
        If Norm(CStr(ws.Cells(r, D_STATUS).Value)) = Norm(TX("w_test")) Then msg = msg & "  [" & TX("w_test") & "]"
        If reason <> "" Then
            msg = msg & vbCrLf & "   " & TX("dlg_skipped")
            msg = msg & vbCrLf & "   " & reason
        ElseIf Norm(CStr(ws.Cells(r, D_STATUS).Value)) = Norm(TX("w_test")) Then
            ' A test row is addressed to YOU with no Cc at all. Showing the real lists
            ' here made the draft look as though the Cc had been lost.
            msg = msg & vbCrLf & "   " & TX("dlg_test_goes")
        Else
            msg = msg & vbCrLf & "   " & TX("dlg_to") & ": " & Shown(ToLine(r))
            If Delegated(r) Then msg = msg & "   (" & TX("w_delegate") & ")"
            msg = msg & vbCrLf & "   " & TX("dlg_cc") & ": " & Shown(CcLine(r))
        End If
        msg = msg & vbCrLf
        ' A MsgBox shows about 1024 characters and silently drops the rest - including
        ' the departments at the bottom of the list and the line naming the mailbox.
        ' Past the limit the remaining departments are counted instead of described.
        If Len(msg) > 760 And i < n Then
            msg = msg & vbCrLf & "+ " & (n - i) & " more" & vbCrLf & TX("dlg_trunc_list") & vbCrLf
            Exit For
        End If
    Next i
    msg = msg & vbCrLf & Replace(TX("dlg_where"), "{box}", CStr(Nz(SetVal(6), "")))
    If warn <> "" Then msg = msg & vbCrLf & "---" & warn
    If MsgBox(msg, vbOKCancel + vbQuestion, TX("app_title")) <> vbOK Then Exit Sub

    ' 3) build the drafts
    gStep = "outlook"
    Set app = OutlookApp()
    If app Is Nothing Then
        MsgBox TX("dlg_no_outlook"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    Application.ScreenUpdating = False
    For i = 1 To n
        r = rowsSel(i)
        gStep = "draft " & i
        Say "HC Projects: " & i & "/" & n & "  " & DeptEN(r)
        reason = reasons(i)                    ' already worked out on the review screen
        If reason <> "" Then
            skipped = skipped + 1
            LogRow r, who, "", reason
        ElseIf MakeDraft(app, r, dueD, who, "round") Then
            made = made + 1
        Else
            skipped = skipped + 1
        End If
    Next i
    gStep = "wbsave"
    Say "HC Projects: finishing"
    SaveBook                      ' the send dates reach disk BEFORE anything is redrawn:
    gStep = "refresh"             ' a redraw that fails must not cost us the round
    Refresh "drafts"
    Application.ScreenUpdating = True
    SaveBook
    Say ""
    If made = 0 Then
        msg = TX("dlg_none")
    Else
        msg = Replace(Replace(TX("dlg_done"), "{n}", CStr(made)), _
                      "{box}", CStr(Nz(SetVal(6), "")))
    End If
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & TX("dlg_skipped_n") & " " & skipped & vbCrLf & TX("dlg_see_log")
    End If
    MsgBox msg, vbInformation, TX("app_title")
    Exit Sub
Failed:
    en = Err.Number: ed = Err.Description     ' read first - Say and TX both clear it
    Application.ScreenUpdating = True
    Say ""
    MsgBox TX("err_general") & " [" & gStep & "] " & en & ": " & ed, _
           vbExclamation, TX("app_title")
End Sub

Private Function Shown(list As String) As String
    ' An empty recipient line used to print as nothing at all, which is exactly how a
    ' silently emptied Cc list went unnoticed. It now says so in words.
    Dim n As Long, i As Long
    If Trim$(list) = "" Then
        Shown = "(EMPTY)"
        Exit Function
    End If
    For i = 1 To Len(list)
        If Mid$(list, i, 1) = ";" Then n = n + 1
    Next i
    Shown = Left$(list, 70)
    If Len(list) > 70 Then Shown = Shown & "..."
    Shown = Shown & "   [" & (n + 1) & "]"
End Function

Private Function BlockedReason(r As Long) As String
    ' what stops this department from getting a proper reply, if anything
    Dim ws As Worksheet
    Set ws = SH("shDept")
    If Blank(ws.Cells(r, D_CONV).Value) And Blank(ws.Cells(r, D_ENTRY).Value) _
       And Blank(ws.Cells(r, D_SUBJ).Value) Then
        BlockedReason = TX("dlg_nothread")
    ElseIf InStr(ToLine(r), "@") = 0 Then
        BlockedReason = TX("dlg_nomail")
    End If
End Function

Private Function ReadOnlyStop() As Boolean
    ' Four people share this file. If Excel handed this copy over read-only - usually
    ' because a colleague has it open - the drafts would still be created but the send
    ' date, the log and the reply status would all be lost without a word, and the same
    ' round could then be built again by somebody else. So the run stops before
    ' anything at all is created.
    On Error Resume Next
    If ThisWorkbook.ReadOnly Then
        Err.Clear
        On Error GoTo 0
        MsgBox TX("dlg_readonly"), vbExclamation, TX("app_title")
        ReadOnlyStop = True
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Sub SaveBook()
    ' a refused save must never look like a failed run
    On Error Resume Next
    ThisWorkbook.Save
    Err.Clear
End Sub

Private Function FindThreadItem(app As Object, r As Long) As Object
    ' The newest message of this department's thread, looked up in YOUR OWN Inbox and
    ' Sent Items. Your own mailbox is cached on this machine, so reading it is instant;
    ' the shared mailbox is not, and walking that is what used to take minutes.
    ' You are on every one of these threads, so the same conversation is here too.
    Dim ws As Worksheet, ns As Object, seed As Object, conv As String, subj As String
    Dim a As Object, b As Object, st As Object, floorD As Date
    Set ws = SH("shDept")
    Set ns = app.GetNamespace("MAPI")
    conv = Trim$(CStr(ws.Cells(r, D_CONV).Value))
    subj = CleanSubject(Trim$(CStr(ws.Cells(r, D_SUBJ).Value)))
    floorD = Date - 150

    On Error Resume Next
    If Not Blank(ws.Cells(r, D_ENTRY).Value) Then
        If Not Blank(ws.Cells(r, D_STORE).Value) Then
            Set seed = ns.GetItemFromID(CStr(ws.Cells(r, D_ENTRY).Value), CStr(ws.Cells(r, D_STORE).Value))
        End If
        If seed Is Nothing Then Set seed = ns.GetItemFromID(CStr(ws.Cells(r, D_ENTRY).Value))
        Err.Clear
    End If
    If conv = "" And Not seed Is Nothing Then          ' an older link: learn the id once
        conv = Trim$(CStr(seed.ConversationID))
        Err.Clear
        If conv <> "" Then ws.Cells(r, D_CONV).Value = conv
        Err.Clear
    End If
    If subj = "" And Not seed Is Nothing Then
        subj = CleanSubject(CStr(seed.Subject))
        Err.Clear
    End If
    Err.Clear
    On Error GoTo 0

    Set a = ScanFolder(ns, 6, conv, subj, floorD)       ' your Inbox
    Set b = ScanFolder(ns, 5, conv, subj, floorD)       ' your Sent Items
    Set FindThreadItem = Newer(a, b)

    ' Only if the thread is not in your mailbox at all - because you were left off it -
    ' is the shared mailbox opened, and only for that one department. It is the slow
    ' place to read, so it is never touched when your own copy was found.
    If FindThreadItem Is Nothing Then
        Set st = SharedStore(app)
        If Not st Is Nothing Then
            Set a = ScanInStore(st, 6, conv, subj, floorD)
            Set b = ScanInStore(st, 5, conv, subj, floorD)
            Set FindThreadItem = Newer(a, b)
        End If
    End If

    ' The conversation id is the ONLY key the reply check matches on, and until now it
    ' was learned only from a hand-linked message. A department found by its subject
    ' alone - the documented normal case - therefore never recorded one, and its
    ' replies could never be detected at all. Whatever the scan found IS this
    ' department's thread, so its id is written down here.
    If conv = "" And Not FindThreadItem Is Nothing Then
        On Error Resume Next
        conv = Trim$(CStr(FindThreadItem.ConversationID))
        If conv <> "" Then ws.Cells(r, D_CONV).Value = conv
        Err.Clear
        On Error GoTo 0
    End If

    If FindThreadItem Is Nothing Then Set FindThreadItem = seed
End Function

Private Function SharedStore(app As Object) As Object
    ' the agency mailbox, matched on its address or on its name as Outlook shows it
    Dim stx As Object, want As String
    want = Norm(FromAddress())
    If want = "" Then Exit Function
    If InStr(want, "@") > 0 Then want = Left$(want, InStr(want, "@") - 1)
    On Error Resume Next
    If Len(want) < 3 Then Exit Function
    For Each stx In app.GetNamespace("MAPI").Stores
        If InStr(Norm(CStr(stx.DisplayName)), want) > 0 Then
            Set SharedStore = stx
            Err.Clear
            On Error GoTo 0
            Exit Function
        End If
        Err.Clear
    Next stx
    Err.Clear
    On Error GoTo 0
End Function

Private Function ScanFolder(ns As Object, folderId As Long, conv As String, _
                            subj As String, floorD As Date) As Object
    Dim f As Object
    On Error Resume Next
    Set f = ns.GetDefaultFolder(folderId)      ' always YOUR mailbox, never a shared one
    Err.Clear
    On Error GoTo 0
    Set ScanFolder = ScanItems(f, (folderId = 5), conv, subj, floorD)
End Function

Private Function ScanInStore(st As Object, folderId As Long, conv As String, _
                             subj As String, floorD As Date) As Object
    Dim f As Object
    On Error Resume Next
    Set f = st.GetDefaultFolder(folderId)
    Err.Clear
    On Error GoTo 0
    Set ScanInStore = ScanItems(f, (folderId = 5), conv, subj, floorD)
End Function

Private Function ScanItems(f As Object, isSent As Boolean, conv As String, _
                           subj As String, floorD As Date) As Object
    ' Newest first, and stop the moment the mail is older than the floor. A thread that
    ' is written to every fortnight is always near the top, so this reads a few dozen
    ' items, never a whole folder.
    Dim items As Object, it As Object, n As Long, dt As Date, sorted_ As Boolean
    On Error Resume Next
    If f Is Nothing Then
        Err.Clear
        Exit Function
    End If
    Set items = f.Items
    If items Is Nothing Then
        Err.Clear
        Exit Function
    End If
    If isSent Then
        items.Sort "[SentOn]", True
    Else
        items.Sort "[ReceivedTime]", True
    End If
    Err.Clear
    Set it = items.GetFirst
    Do While Not it Is Nothing
        n = n + 1
        If n > 3000 Then Exit Do
        If it.Class = 43 Then
            dt = Now
            If isSent Then dt = it.SentOn Else dt = it.ReceivedTime
            If Err.Number <> 0 Then
                Err.Clear
                dt = Now
            End If
            ' The floor is only a shortcut, and it is only trustworthy once we have
            ' seen a recent item. Some stores refuse to sort; there the oldest mail
            ' comes first, and trusting the floor would end the search immediately.
            If dt >= floorD Then
                sorted_ = True
            ElseIf sorted_ Then
                Exit Do
            End If
            If SameThread(it, conv, subj) Then
                Set ScanItems = it
                Err.Clear
                On Error GoTo 0
                Exit Function
            End If
        End If
        Err.Clear
        Set it = items.GetNext
    Loop
    Err.Clear
    On Error GoTo 0
End Function

Private Function SameThread(it As Object, conv As String, subj As String) As Boolean
    ' the conversation id is exact; the subject is the fallback when nothing was linked
    On Error Resume Next
    If conv <> "" Then
        If Trim$(CStr(it.ConversationID)) = conv Then SameThread = True
    End If
    If Not SameThread Then
        ' compared with spaces removed and alef/ya/ta-marbuta levelled out, so a small
        ' difference in how the subject was typed does not lose the thread. A needle
        ' this short would match almost anything, so it is refused outright.
        If Len(Norm(subj)) >= 8 Then
            If InStr(1, Norm(CStr(it.Subject)), Norm(subj)) > 0 Then SameThread = True
        End If
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function Newer(a As Object, b As Object) As Object
    If a Is Nothing Then
        Set Newer = b
        Exit Function
    End If
    If b Is Nothing Then
        Set Newer = a
        Exit Function
    End If
    If ItemWhen(b) > ItemWhen(a) Then Set Newer = b Else Set Newer = a
End Function

Private Function ItemWhen(it As Object) As Date
    ' a sent message and a received one keep their time in different places
    Dim d1 As Date, d2 As Date
    On Error Resume Next
    d1 = CDate(it.SentOn)
    Err.Clear
    d2 = CDate(it.ReceivedTime)
    Err.Clear
    On Error GoTo 0
    If d2 > d1 Then ItemWhen = d2 Else ItemWhen = d1
End Function

Private Function FromAddress() As String
    FromAddress = Trim$(CStr(Nz(SetVal(6), "")))
End Function

Private Sub ApplyFrom(app As Object, m As Object, note As String)
    ' The draft lives in your own mailbox, so the sender has to be set on purpose.
    ' A matching account is best; otherwise Outlook takes the address directly, which
    ' is the normal way to send on behalf of a shared mailbox.
    Dim acc As Object, want As String, ok As Boolean
    want = FromAddress()
    If want = "" Then Exit Sub
    On Error Resume Next
    For Each acc In app.Session.Accounts
        If StrComp(CStr(acc.SmtpAddress), want, vbTextCompare) = 0 _
           Or Norm(CStr(acc.DisplayName)) = Norm(want) Then
            Err.Clear
            Set m.SendUsingAccount = acc
            If Err.Number = 0 Then ok = True
            Err.Clear
            Exit For
        End If
        Err.Clear
    Next acc
    If Not ok Then
        If InStr(want, "@") > 0 Then
            Err.Clear
            m.SentOnBehalfOfName = want
            If Err.Number = 0 Then ok = True
            Err.Clear
        End If
    End If
    If Not ok Then note = note & "from? "
    Err.Clear
    On Error GoTo 0
End Sub

Private Function DeliverAt() As Variant
    ' Today at the time set in the settings sheet.
    '
    ' The hour and minute are taken apart by hand rather than handed to CDate. The
    ' cell may hold the text "12:00" OR a real Excel time, and CDate on a real time
    ' goes through the Windows locale, which on an Arabic machine renders it in a
    ' shape CDate itself may then read back differently - or not at all.
    Dim h As Long, mi As Long
    If Not HourMin(SetVal(15), h, mi) Then Exit Function
    DeliverAt = DateSerial(Year(Date), Month(Date), Day(Date)) + TimeSerial(h, mi, 0)
End Function

Private Function HourMin(v As Variant, ByRef h As Long, ByRef mi As Long) As Boolean
    ' reads an hour and a minute out of a settings cell, whatever shape it is in:
    ' a real Excel time, the text "12:00", "12:00 PM", "12", or an Arabic AM/PM mark
    Dim s As String, i As Long, ch As String, a As String, b As String, pm As Boolean
    Dim am As Boolean, seen As Boolean
    On Error Resume Next
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function

    If VarType(v) = vbDate Then                  ' a real time: read it directly
        h = Hour(v): mi = Minute(v)
        Err.Clear
        On Error GoTo 0
        HourMin = True
        Exit Function
    End If
    If IsNumeric(v) Then                         ' a serial fraction of a day
        If CDbl(v) >= 0 And CDbl(v) < 1 Then
            h = Hour(CDate(v)): mi = Minute(CDate(v))
            Err.Clear
            On Error GoTo 0
            HourMin = True
            Exit Function
        End If
    End If

    s = Trim$(CStr(v))
    Err.Clear
    On Error GoTo 0
    If s = "" Then Exit Function
    ' PM in English or Arabic; the Arabic marks are a single letter
    If InStr(1, s, "PM", vbTextCompare) > 0 Then pm = True
    If InStr(1, s, "AM", vbTextCompare) > 0 Then am = True
    If InStr(s, ChrW(&H645)) > 0 Then pm = True
    If InStr(s, ChrW(&H635)) > 0 Then am = True
    For i = 1 To Len(s)                          ' the digits before and after the colon
        ch = Mid$(s, i, 1)
        If ch >= "0" And ch <= "9" Then
            If seen Then b = b & ch Else a = a & ch
        ElseIf ch = ":" Then
            If seen Then Exit For
            seen = True
        End If
    Next i
    If a = "" Then Exit Function
    If Len(a) > 2 Then Exit Function
    h = CLng(a)
    If b <> "" Then mi = CLng(Left$(b & "0", 2))
    If pm And h < 12 Then h = h + 12
    If am And h = 12 Then h = 0
    If h < 0 Or h > 23 Then Exit Function
    If mi < 0 Or mi > 59 Then Exit Function
    HourMin = True
End Function

Private Function MakeDraft(app As Object, r As Long, dueD As Date, who As String, _
                           kind As String) As Boolean
    ' Builds a true Reply All on the department's own thread.
    '
    ' Two rules keep the quoted history intact, and both were learned the hard way:
    '  1. the draft is never moved to another mailbox. Moving an unsent reply across
    '     stores makes Outlook rebuild it, and the rebuilt copy loses the quoted
    '     messages and shows raw Exchange addresses instead of names.
    '  2. the new text is spliced in just after <body>, so Outlook's own quoted block
    '     stays exactly where it was. If Outlook hands back an empty reply, the quoted
    '     block is rebuilt here from the original message rather than left missing.
    Dim ws As Worksheet, src As Object, m As Object
    Dim subj As String, bodyHtml As String, html As String
    Dim baseLen As Long, afterLen As Long, isTest As Boolean, note As String
    Dim bodyFailed As Boolean, isNudge As Boolean, toS As String, ccS As String

    isNudge = (kind = "nudge")

    Set ws = SH("shDept")
    isTest = (Norm(CStr(ws.Cells(r, D_STATUS).Value)) = Norm(TX("w_test")))
    On Error Resume Next

    ' --- the latest message in this department's thread
    Set src = FindThreadItem(app, r)
    Err.Clear
    If src Is Nothing Then
        On Error GoTo 0
        LogRow r, who, "", TX("dlg_thread_gone")
        Exit Function
    End If
    subj = CleanSubject(CStr(src.Subject))
    Err.Clear

    ' --- let Outlook build the reply: recipients, subject and history
    Set m = src.ReplyAll
    If m Is Nothing Or Err.Number <> 0 Then
        note = TX("dlg_reply_failed") & " (" & Err.Number & ")"
        Err.Clear
        On Error GoTo 0
        LogRow r, who, subj, note
        Exit Function
    End If
    Err.Clear
    ApplyFrom app, m, note

    ' --- make sure the quoted history is really there
    html = CStr(m.HTMLBody)
    Err.Clear
    If Len(InnerBody(html)) < 60 Then
        html = "<html><body>" & QuotedBlock(src) & "</body></html>"
        note = note & TX("dlg_history_rebuilt") & " "
    End If
    baseLen = Len(html)

    ' --- recipients and delivery time. The subject Outlook produced is left untouched:
    '     it already reads "RE: <the thread's own subject>", and rewriting it is one more
    '     way to end up with something that is not a real reply.
    If Trim$(CStr(ws.Cells(r, D_SUBJ).Value)) <> "" Then
        If Norm(CleanSubject(CStr(ws.Cells(r, D_SUBJ).Value))) <> Norm(subj) Then
            note = note & TX("dlg_subj_differs") & " "
        End If
    End If
    ' The recipient lists are read FIRST and the read is checked. Inside the blanket
    ' error trap, an unreadable cell used to make the whole SetRecipients statement
    ' vanish, and the draft then kept whatever ReplyAll had put there - every address
    ' on the original thread - while the log still said ok.
    Err.Clear
    toS = ToLine(r)
    ccS = CcLine(r)
    If Err.Number <> 0 Then
        note = note & TX("dlg_rcpt_failed") & " (" & Err.Number & ") "
        Err.Clear
        On Error GoTo 0
        LogRow r, who, subj, Trim$(note)
        Exit Function
    End If
    ' A test goes to YOU and to nobody else, which is what the departments sheet
    ' promises. It used to drop only the Cc and leave the real representative in To,
    ' so a "safe" test was one Send away from reaching a real department.
    If isTest Then
        toS = MyMail(app, who)
        ccS = ""
        If toS = "" Then
            note = note & TX("dlg_test_noself")
            Err.Clear
            On Error GoTo 0
            LogRow r, who, subj, Trim$(note)
            Exit Function
        End If
    End If
    note = note & SetRecipients(m, toS, ccS)
    ' The round is held until the shared delivery time so every department gets the
    ' same window. A reminder is not: it goes out when you send it, because it is sent
    ' on the deadline morning and holding it until noon would deliver it too late.
    If Not isNudge Then
        If Not IsEmpty(DeliverAt()) Then
            If CDate(DeliverAt()) > Now Then m.DeferredDeliveryTime = CDate(DeliverAt())
        End If
    End If
    Err.Clear

    ' --- the new text goes on top of the untouched history, and nothing else changes
    If isNudge Then
        bodyHtml = NudgeHtml(r, dueD)
    Else
        bodyHtml = MailHtml(r, dueD)
    End If
    Err.Clear
    m.HTMLBody = SpliceTop(html, bodyHtml)
    ' An empty bodyHtml would sail through the length test below, because the draft is
    ' then simply the old thread and is longer than nothing. It is caught here instead.
    If Err.Number <> 0 Or Len(bodyHtml) = 0 Then
        note = note & "body(" & Err.Number & "/" & Len(bodyHtml) & ") "
        Err.Clear
        bodyFailed = True
    End If

    Err.Clear
    m.Save
    If Err.Number <> 0 Then
        note = note & TX("err_row") & " " & Err.Number
        Err.Clear
        On Error GoTo 0
        LogRow r, who, subj, Trim$(note)
        Exit Function                     ' nothing was saved, so nothing is reported
    End If

    ' --- read the draft back: the new text and the history both have to be in there.
    '     A body that is barely longer than the new text alone means the history is gone.
    afterLen = Len(CStr(m.HTMLBody))
    Err.Clear
    note = note & WhereSaved(m)
    If bodyFailed Or afterLen < Len(bodyHtml) + baseLen \ 2 Then
        ' the saved draft is missing the new text, the history, or both. It is left in
        ' place to be looked at, but it is reported as not done rather than as a draft.
        note = note & "INCOMPLETE(" & afterLen & "<" & (Len(bodyHtml) + baseLen \ 2) & ")"
        Err.Clear
        On Error GoTo 0
        LogRow r, who, subj, Trim$(note)
        Exit Function
    End If
    note = note & "ok"
    ' The TIME, not just the day: a late answer to the PREVIOUS round that arrived
    ' this morning must not read as an answer to the round being built now.
    '
    ' A reminder must not touch the round's own send date - the reply status is read
    ' from it, and moving it would turn a department that replied back into waiting.
    ' A new round, on the other hand, clears the last round's reply and reminder
    ' dates, so the table always describes the round in hand.
    Err.Clear
    If isNudge Then
        ws.Cells(r, D_NUDGE).Value = Now
    Else
        ws.Cells(r, D_REPLY).ClearContents
        ws.Cells(r, D_ATT).ClearContents
        ws.Cells(r, D_NUDGE).ClearContents
        ws.Cells(r, D_SENT).Value = Now
    End If
    If Err.Number <> 0 Then
        ' the draft exists but the sheet refused the date, so this is not a clean run
        note = note & "STAMP(" & Err.Number & ") "
        Err.Clear
        On Error GoTo 0
        LogRow r, who, subj, Trim$(note)
        Exit Function
    End If
    Err.Clear
    On Error GoTo 0
    If isNudge Then
        LogRow r, who, subj, TX("w_draft") & " " & TX("w_nudge") & " " & Trim$(note)
    Else
        LogRow r, who, subj, TX("w_draft") & " " & Trim$(note)
    End If
    MakeDraft = True
End Function

Private Function MyMail(app As Object, who As String) As String
    ' Your own address, for a test draft: the team sheet first, then whatever Outlook
    ' is signed in as. Never the department's representative.
    Dim ws As Worksheet, r As Long, last As Long, acc As Object, v As String
    On Error Resume Next
    Set ws = SH("shTeam")
    If Not ws Is Nothing Then
        last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        For r = 6 To last
            If Norm(CStr(ws.Cells(r, 2).Value)) = Norm(who) Then
                v = Trim$(CStr(ws.Cells(r, 5).Value))
                Exit For
            End If
        Next r
    End If
    Err.Clear
    If InStr(v, "@") = 0 Then
        v = ""
        For Each acc In app.Session.Accounts
            v = Trim$(CStr(acc.SmtpAddress))
            Err.Clear
            If InStr(v, "@") > 0 Then Exit For
            v = ""
        Next acc
    End If
    Err.Clear
    On Error GoTo 0
    If InStr(v, "@") > 0 Then MyMail = v
End Function

Private Function WhereSaved(m As Object) As String
    ' which mailbox actually received the draft, so nobody hunts in the wrong Drafts
    Dim nm As String
    On Error Resume Next
    nm = CStr(m.Parent.Store.DisplayName)
    If nm = "" Then nm = CStr(m.Parent.Parent.Name)
    Err.Clear
    On Error GoTo 0
    If nm <> "" Then WhereSaved = "in[" & nm & "] "
End Function

Private Function InnerBody(html As String) As String
    ' what sits between <body ...> and </body>, or the whole string if there is no body tag
    Dim a As Long, b As Long
    a = InStr(1, LCase$(html), "<body")
    If a = 0 Then
        InnerBody = Trim$(html)
        Exit Function
    End If
    a = InStr(a, html, ">")
    If a = 0 Then Exit Function
    b = InStr(a, LCase$(html), "</body")
    If b = 0 Then b = Len(html) + 1
    InnerBody = Trim$(Mid$(html, a + 1, b - a - 1))
End Function

Private Function SpliceTop(html As String, inject As String) As String
    ' puts inject at the very start of the message body, leaving everything else alone
    Dim a As Long
    a = InStr(1, LCase$(html), "<body")
    If a = 0 Then
        SpliceTop = inject & html
        Exit Function
    End If
    a = InStr(a, html, ">")
    If a = 0 Then
        SpliceTop = inject & html
        Exit Function
    End If
    SpliceTop = Left$(html, a) & inject & Mid$(html, a + 1)
End Function

Private Function QuotedBlock(src As Object) As String
    ' a stand-in for Outlook's quoted block, used only when the reply comes back empty
    Dim sF As String, head As String, body As String, sentOn As String
    On Error Resume Next
    sF = "font-family:" & FontName() & ",Arial,sans-serif;font-size:11pt;color:#000000;"
    sentOn = ""
    Err.Clear
    sentOn = Format$(src.SentOn, "yyyy/mm/dd hh:nn")
    If Err.Number <> 0 Then
        Err.Clear
        sentOn = Format$(src.ReceivedTime, "yyyy/mm/dd hh:nn")
        Err.Clear
    End If
    head = "<div style='border-top:1px solid #B8C3D1;margin-top:18px;padding-top:10px;" & sF & "'>"
    head = head & "<p style='margin:0 0 10px 0;" & sF & "'>"
    head = head & "<b>From:</b> " & HtmlEsc(SenderOf(src)) & "<br>"
    head = head & "<b>Sent:</b> " & HtmlEsc(sentOn) & "<br>"
    head = head & "<b>To:</b> " & HtmlEsc(CStr(src.To)) & "<br>"
    If Trim$(CStr(src.CC)) <> "" Then head = head & "<b>Cc:</b> " & HtmlEsc(CStr(src.CC)) & "<br>"
    head = head & "<b>Subject:</b> " & HtmlEsc(CStr(src.Subject)) & "</p>"
    Err.Clear
    body = CStr(src.HTMLBody)
    If Err.Number <> 0 Then
        Err.Clear
        body = ""
    End If
    If Len(InnerBody(body)) < 10 Then
        body = "<div style='" & sF & "white-space:pre-wrap'>" & HtmlEsc(CStr(src.Body)) & "</div>"
        Err.Clear
    Else
        body = InnerBody(body)
    End If
    Err.Clear
    On Error GoTo 0
    QuotedBlock = head & body & "</div>"
End Function

Private Function SenderOf(src As Object) As String
    On Error Resume Next
    SenderOf = CStr(src.SentOnBehalfOfName)
    If Trim$(SenderOf) = "" Then SenderOf = CStr(src.SenderName)
    Err.Clear
    On Error GoTo 0
End Function

Private Function SetRecipients(m As Object, toList As String, ccList As String) As String
    ' Outlook is handed a plain "a; b; c" list and resolves it exactly as it would if a
    ' person had typed it. NOTHING is ever removed here. An earlier version deleted any
    ' entry the address book did not resolve on the spot, which quietly emptied a Cc list
    ' made of display names rather than addresses. Unresolved entries are only reported.
    Dim bad As String, i As Long
    On Error Resume Next
    For i = m.Recipients.Count To 1 Step -1       ' drop whoever ReplyAll brought along
        m.Recipients.Remove i
    Next i
    Err.Clear
    m.To = toList
    m.CC = ccList
    Err.Clear
    m.Recipients.ResolveAll
    Err.Clear
    bad = Unresolved(m)
    Err.Clear
    On Error GoTo 0
    If bad <> "" Then SetRecipients = TX("dlg_unresolved") & ": " & bad & " "
End Function

Private Function Unresolved(m As Object) As String
    ' names still unknown to the address book, reported so they can be corrected
    Dim i As Long, rcp As Object, out As String
    On Error Resume Next
    For i = 1 To m.Recipients.Count
        Set rcp = Nothing
        Set rcp = m.Recipients.Item(i)
        If Not rcp Is Nothing Then
            If Not rcp.Resolved Then out = out & CStr(rcp.Name) & " "
        End If
        Err.Clear
    Next i
    Err.Clear
    On Error GoTo 0
    Unresolved = Trim$(out)
End Function

Private Sub LogRow(r As Long, who As String, subj As String, state As String)
    Dim ws As Worksheet, dept As Worksheet
    On Error Resume Next
    Set ws = SH("shLog")
    Set dept = SH("shDept")
    ws.Rows(6).Insert
    ws.Range("B6:G6").Value = Array(Now, dept.Cells(r, D_NAME).Value, who, _
        Left$("To: " & ToLine(r) & "  |  Cc: " & CcLine(r), 250), Left$(subj, 120), state)
    ws.Range("B6").NumberFormat = "yyyy/mm/dd hh:mm"
    StyleRow ws, 6, 2, 7
    Err.Clear
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
    Dim ed As String
    Dim app As Object, item As Object, ws As Worksheet, r As Long, n As Long
    Dim list As String, rowsN() As Long, ans As String, i As Long
    On Error GoTo Failed
    Fresh
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
    SaveBook
    MsgBox Replace(TX("dlg_link_done"), "{dept}", DeptEN(r)), vbInformation, TX("app_title")
    Exit Sub
Failed:
    ed = Err.Description                      ' read first - TX clears it
    MsgBox TX("err_link") & ": " & ed, vbExclamation, TX("app_title")
End Sub

Private Function CleanSubject(s As String) As String
    Dim out As String, before As String
    out = Trim$(s)
    Do
        before = out
        out = StripPrefix(out, "re:")
        out = StripPrefix(out, "fw:")
        out = StripPrefix(out, "fwd:")
        out = StripPrefix(out, ChrW(&H631) & ChrW(&H62F) & ":")              ' Arabic "re:"
        out = StripPrefix(out, ChrW(&H62C) & ChrW(&H648) & ChrW(&H627) & ChrW(&H628) & ":")
    Loop While out <> before And out <> ""
    CleanSubject = out
End Function

Private Function StripPrefix(s As String, p As String) As String
    StripPrefix = s
    If Len(s) > Len(p) Then
        If LCase$(Left$(s, Len(p))) = LCase$(p) Then StripPrefix = Trim$(Mid$(s, Len(p) + 1))
    End If
End Function

' =====================================================================
'  replies
' =====================================================================
Public Sub Btn_Replies()
    ' Reads the mail and records who answered this round.
    '
    ' Two things here were learned the hard way:
    '  1. our own mail must not count. These rounds go out FROM the agency mailbox
    '     with the team in Cc, so a copy of what WE sent lands in our own Inbox, on
    '     the same conversation, dated after the round. Counted, it would mark every
    '     department as having answered. Anything sent by the agency mailbox or by one
    '     of the team is therefore skipped.
    '  2. the answer may not be in our Inbox at all. The representatives reply to the
    '     agency mailbox, and whoever is not in Cc never sees it. So if any department
    '     is still silent once our own Inbox has been read, the agency mailbox is read
    '     as well - once for the whole run, never once per department, which is what
    '     used to take minutes.
    Dim app As Object, ns As Object, f As Object, st As Object
    Dim ws As Worksheet, found As Long, lastR As Long, en As Long, ed As String
    On Error GoTo Failed
    gStep = "replies:init"
    Fresh
    Set ws = SH("shDept")
    Set app = OutlookApp()
    If app Is Nothing Then
        MsgBox TX("dlg_no_outlook"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    Set ns = app.GetNamespace("MAPI")
    Set f = ns.GetDefaultFolder(6)                     ' always YOUR OWN Inbox
    If f Is Nothing Then
        MsgBox TX("dlg_no_store"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    gOurs = "": gBox = ""                              ' rebuilt once per run
    Application.ScreenUpdating = False
    Say "HC Projects: checking replies"
    lastR = LastDeptRow()                              ' once, not once per message
    gStep = "replies:mine"
    ScanReplies f, ws, lastR, found, 1500

    If AnyWaiting(lastR) Then
        gStep = "replies:shared"
        Set f = Nothing
        On Error Resume Next
        Set st = SharedStore(app)
        If Not st Is Nothing Then Set f = st.GetDefaultFolder(6)
        Err.Clear
        On Error GoTo Failed
        If Not f Is Nothing Then
            Say "HC Projects: checking the agency mailbox"
            ScanReplies f, ws, lastR, found, 1500
        End If
    End If

    gStep = "replies:save"
    SaveBook                                           ' the dates are safe before
    gStep = "replies:refresh"                          ' the sheets are redrawn
    Refresh "replies"
    Application.ScreenUpdating = True
    SaveBook
    Say ""
    MsgBox Replace(TX("dlg_rep_done"), "{n}", CStr(found)), vbInformation, TX("app_title")
    Exit Sub
Failed:
    en = Err.Number: ed = Err.Description              ' read before anything clears it
    Application.ScreenUpdating = True
    Say ""
    MsgBox TX("err_replies") & " [" & gStep & "] " & en & ": " & ed, _
           vbExclamation, TX("app_title")
End Sub

Private Sub ScanReplies(f As Object, ws As Worksheet, lastR As Long, _
                        found As Long, cap As Long)
    ' one pass over one folder, newest first, stopping at the cap
    Dim items As Object, it As Object, r As Long, conv As String, scanned As Long
    Dim sentD As Variant, prev As Variant, after As Boolean, write_ As Boolean
    Dim got As Date, okTime As Boolean
    On Error Resume Next
    Set items = f.Items.Restrict("[ReceivedTime] >= '" & _
                Format$(Date - 120, "ddddd h:nn AMPM") & "'")
    If items Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    items.Sort "[ReceivedTime]", True
    Err.Clear
    Set it = items.GetFirst
    Do While Not it Is Nothing
        scanned = scanned + 1
        If scanned > cap Then Exit Do          ' newest first, so older mail adds nothing
        If it.Class = 43 Then
            If Not Ours(it) Then               ' our own copy of the round is not an answer
                conv = ""
                conv = CStr(it.ConversationID)
                Err.Clear
                okTime = False
                got = CDate(it.ReceivedTime)
                If Err.Number = 0 Then okTime = True
                Err.Clear
                If conv <> "" And okTime Then
                    For r = D_FIRST To lastR
                        If conv = CStr(ws.Cells(r, D_CONV).Value) Then
                            sentD = ws.Cells(r, D_SENT).Value
                            after = False
                            If Blank(sentD) Then
                                after = True
                            ElseIf IsDate(sentD) Then
                                If got >= CDate(sentD) Then after = True
                            End If
                            If after Then
                                prev = ws.Cells(r, D_REPLY).Value
                                write_ = False
                                If Blank(prev) Or Not IsDate(prev) Then
                                    write_ = True
                                ElseIf got > CDate(prev) Then
                                    write_ = True
                                End If
                                If write_ Then
                                    Err.Clear
                                    ws.Cells(r, D_REPLY).Value = got
                                    ' counted only when the sheet really took the value
                                    If Err.Number = 0 Then found = found + 1
                                    Err.Clear
                                End If
                                ' The reply that matters is the one carrying the deck.
                                ' It is recorded separately, because a department may
                                ' answer "noted" first and send the file days later.
                                If HasDeck(it) Then
                                    prev = ws.Cells(r, D_ATT).Value
                                    write_ = False
                                    If Blank(prev) Or Not IsDate(prev) Then
                                        write_ = True
                                    ElseIf got > CDate(prev) Then
                                        write_ = True
                                    End If
                                    If write_ Then ws.Cells(r, D_ATT).Value = got
                                    Err.Clear
                                End If
                            End If
                        End If
                    Next r
                End If
            End If
        End If
        Err.Clear
        Set it = items.GetNext
    Loop
    Err.Clear
End Sub

Private Function HasDeck(it As Object) As Boolean
    ' True when the message carries a real PowerPoint file.
    '
    ' Counting attachments alone would be useless: a logo in the signature is an
    ' attachment too, so every corporate mail would look as though it carried the
    ' deck. Only .pptx and .ppt count, which excludes signature images by itself.
    Dim i As Long, nm As String, dot As Long, ext As String
    On Error Resume Next
    For i = 1 To it.Attachments.Count
        nm = ""
        nm = LCase$(Trim$(CStr(it.Attachments.Item(i).FileName)))
        Err.Clear
        dot = InStrRev(nm, ".")
        If dot > 0 Then
            ext = Mid$(nm, dot + 1)
            If ext = "pptx" Or ext = "ppt" Or ext = "ppsx" Or ext = "pptm" Then
                Err.Clear
                On Error GoTo 0
                HasDeck = True
                Exit Function
            End If
        End If
    Next i
    Err.Clear
    On Error GoTo 0
End Function

Private Function WorkDays(fromD As Date, toD As Date) As Long
    ' Whole working days between the round going out and the answer arriving.
    ' An answer the same day is zero. Friday and Saturday are not counted.
    Dim d As Date, n As Long, guard As Long
    If toD <= fromD Then Exit Function
    d = Int(fromD)
    Do While d < Int(toD)
        d = d + 1
        guard = guard + 1
        If guard > 400 Then Exit Do
        If Weekday(d, vbSunday) <> vbFriday And Weekday(d, vbSunday) <> vbSaturday Then
            n = n + 1
        End If
    Loop
    WorkDays = n
End Function

Private Function AnyWaiting(lastR As Long) As Boolean
    ' is any department still silent? If none is, the agency mailbox is never opened
    Dim ws As Worksheet, r As Long
    Set ws = SH("shDept")
    Dim st As String
    For r = D_FIRST To lastR
        st = Replied(r)
        If st = TX("st_waiting") Or st = TX("st_duenudge") Or st = TX("st_repnoatt") Then
            AnyWaiting = True
            Exit Function
        End If
    Next r
End Function

Private Function Ours(it As Object) As Boolean
    ' True for mail the agency mailbox or one of the team sent, so that our own copy
    ' of the round is never read as somebody's answer to it.
    '
    ' The team is matched by exact name or address. The agency mailbox needs more than
    ' that: its address is HCPerformance-GMA@... while Outlook shows it as
    ' "HC Performance", and internal senders often arrive as a long Exchange path
    ' rather than an address at all. So for the mailbox itself the comparison is made
    ' on letters and digits only, and either string may be the start of the other.
    Dim ws As Worksheet, r As Long, last As Long, v As String
    On Error Resume Next
    If gOurs = "" Then
        gOurs = ";"
        Set ws = SH("shTeam")
        If Not ws Is Nothing Then
            last = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
            For r = 6 To last
                v = Norm(CStr(ws.Cells(r, 2).Value))        ' the Arabic name
                If Len(v) > 2 Then gOurs = gOurs & v & ";"
                v = Norm(CStr(ws.Cells(r, 4).Value))        ' the English name
                If Len(v) > 2 Then gOurs = gOurs & v & ";"
                v = Norm(CStr(ws.Cells(r, 5).Value))        ' the address, if given
                If Len(v) > 2 Then gOurs = gOurs & v & ";"
            Next r
        End If
        gBox = Squash(FromAddress())
        If InStr(gBox, "@") > 0 Then gBox = Left$(gBox, InStr(gBox, "@") - 1)
        Err.Clear
    End If
    If Listed(CStr(it.SenderName)) Then
        Ours = True
    ElseIf Listed(CStr(it.SenderEmailAddress)) Then
        Ours = True
    ElseIf IsBox(CStr(it.SenderName)) Then
        Ours = True
    ElseIf IsBox(CStr(it.SenderEmailAddress)) Then
        Ours = True
    ElseIf IsBox(CStr(it.SentOnBehalfOfName)) Then
        Ours = True
    End If
    Err.Clear
    On Error GoTo 0
End Function

Private Function Listed(who As String) As Boolean
    Dim v As String
    v = Norm(who)
    If v = "" Then Exit Function
    Listed = (InStr(gOurs, ";" & v & ";") > 0)
End Function

Private Function IsBox(who As String) As Boolean
    ' is this the agency mailbox, under any of the shapes Outlook shows it in?
    Dim v As String
    If gBox = "" Then Exit Function
    If Len(gBox) < 6 Then Exit Function
    v = Squash(who)
    If Len(v) < 6 Then Exit Function
    If InStr(v, "@") > 0 Then v = Left$(v, InStr(v, "@") - 1)
    If Len(v) < 6 Then Exit Function
    If Left$(gBox, Len(v)) = v Then
        IsBox = True
    ElseIf Left$(v, Len(gBox)) = gBox Then
        IsBox = True
    ElseIf InStr(v, gBox) > 0 Then
        IsBox = True                  ' an Exchange path with the name buried in it
    End If
End Function

Private Function Squash(s As String) As String
    ' letters, digits and "@" only, lower case: "HC Performance" and
    ' "HCPerformance-GMA" then compare as the same thing
    Dim i As Long, c As Long, ch As String
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or c = 64 Then
            Squash = Squash & ch
        End If
    Next i
    Squash = LCase$(Squash)
End Function

' =====================================================================
'  reminder for the departments that have not replied
' =====================================================================
Public Sub Btn_Nudge()
    Dim en As Long, ed As String
    ' A plain Reply All on the same thread, for your own departments that were sent
    ' the round and have not answered it. Everything else is left alone: the round's
    ' send date is not touched, and nothing is ever sent - only saved as a draft.
    Dim who As String, ws As Worksheet, r As Long, rowsSel() As Long, n As Long
    Dim i As Long, msg As String, app As Object, reason As String, reasons() As String
    Dim made As Long, skipped As Long, dueD As Date, st As String
    On Error GoTo Failed
    gStep = "nudge:init"
    Fresh
    If ReadOnlyStop() Then Exit Sub
    Set ws = SH("shDept")
    gStep = "nudge:who"
    who = AskWho()
    If who = "" Then Exit Sub
    dueD = Date          ' the reminder carries no date of its own

    ' 1) your departments that are still waiting for a reply
    n = 0
    ReDim rowsSel(1 To 60)
    For r = D_FIRST To LastDeptRow()
        If Norm(CStr(ws.Cells(r, D_OWNER).Value)) = Norm(who) And n < 60 Then
            st = Norm(CStr(ws.Cells(r, D_STATUS).Value))
            If st <> Norm(TX("w_hold")) And Trim$(CStr(ws.Cells(r, D_NAME).Value)) <> "" Then
                If NeedsNudge(r) Then
                    n = n + 1
                    rowsSel(n) = r
                End If
            End If
        End If
    Next r
    If n = 0 Then
        MsgBox TX("dlg_nudge_none"), vbInformation, TX("app_title")
        Exit Sub
    End If

    ' 2) review screen, same shape as the round
    gStep = "nudge:review"
    msg = ""
    ' The reminder asks for the update before noon TODAY, so sending it after noon
    ' asks for something already overdue. Worth saying before anything is built.
    If Not IsEmpty(DeliverAt()) Then
        If CDate(DeliverAt()) <= Now Then msg = TX("dlg_nudge_late") & vbCrLf & vbCrLf
    End If
    msg = msg & TX("dlg_nudge_confirm") & vbCrLf
    ' Every reason is worked out for EVERY department first. The list shown below is
    ' capped, and a department whose reason was never worked out would be attempted
    ' later instead of skipped.
    ReDim reasons(1 To n)
    For i = 1 To n
        reasons(i) = BlockedReason(rowsSel(i))
    Next i
    For i = 1 To n
        r = rowsSel(i)
        reason = reasons(i)
        msg = msg & vbCrLf & "- " & DeptEN(r)
        If Norm(CStr(ws.Cells(r, D_STATUS).Value)) = Norm(TX("w_test")) Then msg = msg & "  [" & TX("w_test") & "]"
        msg = msg & "  [" & DaysWaiting(r) & "]"
        If AlreadyNudged(r) Then msg = msg & "  [" & TX("dlg_nudge_already") & "]"
        If reason <> "" Then
            msg = msg & vbCrLf & "   " & TX("dlg_skipped")
            msg = msg & vbCrLf & "   " & reason
        ElseIf Norm(CStr(ws.Cells(r, D_STATUS).Value)) = Norm(TX("w_test")) Then
            ' A test row is addressed to YOU with no Cc at all. Showing the real lists
            ' here made the draft look as though the Cc had been lost.
            msg = msg & vbCrLf & "   " & TX("dlg_test_goes")
        Else
            msg = msg & vbCrLf & "   " & TX("dlg_to") & ": " & Shown(ToLine(r))
            If Delegated(r) Then msg = msg & "   (" & TX("w_delegate") & ")"
            msg = msg & vbCrLf & "   " & TX("dlg_cc") & ": " & Shown(CcLine(r))
        End If
        msg = msg & vbCrLf
        ' A MsgBox shows about 1024 characters and silently drops the rest - including
        ' the departments at the bottom of the list and the line naming the mailbox.
        ' Past the limit the remaining departments are counted instead of described.
        If Len(msg) > 760 And i < n Then
            msg = msg & vbCrLf & "+ " & (n - i) & " more" & vbCrLf & TX("dlg_trunc_list") & vbCrLf
            Exit For
        End If
    Next i
    msg = msg & vbCrLf & Replace(TX("dlg_where"), "{box}", CStr(Nz(SetVal(6), "")))
    If MsgBox(msg, vbOKCancel + vbQuestion, TX("app_title")) <> vbOK Then Exit Sub

    ' 3) build the reminder drafts
    gStep = "nudge:outlook"
    Set app = OutlookApp()
    If app Is Nothing Then
        MsgBox TX("dlg_no_outlook"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    Application.ScreenUpdating = False
    For i = 1 To n
        r = rowsSel(i)
        gStep = "nudge:draft " & i
        Say "HC Projects: " & i & "/" & n & "  " & DeptEN(r)
        reason = reasons(i)
        If reason <> "" Then
            skipped = skipped + 1
            LogRow r, who, "", TX("w_nudge") & " - " & reason
        ElseIf MakeDraft(app, r, dueD, who, "nudge") Then
            made = made + 1
        Else
            skipped = skipped + 1
        End If
    Next i
    gStep = "nudge:save"
    Say "HC Projects: finishing"
    SaveBook                      ' as above: the reminder dates are safe first
    gStep = "nudge:refresh"
    Refresh "nudge"
    Application.ScreenUpdating = True
    SaveBook
    Say ""
    If made = 0 Then
        msg = TX("dlg_none")
    Else
        msg = Replace(Replace(TX("dlg_nudge_done"), "{n}", CStr(made)), _
                      "{box}", CStr(Nz(SetVal(6), "")))
    End If
    If skipped > 0 Then
        msg = msg & vbCrLf & vbCrLf & TX("dlg_skipped_n") & " " & skipped & vbCrLf & TX("dlg_see_log")
    End If
    MsgBox msg, vbInformation, TX("app_title")
    Exit Sub
Failed:
    en = Err.Number: ed = Err.Description     ' read first - Say and TX both clear it
    Application.ScreenUpdating = True
    Say ""
    MsgBox TX("err_nudge") & " [" & gStep & "] " & en & ": " & ed, _
           vbExclamation, TX("app_title")
End Sub

Private Function DaysWaiting(r As Long) As String
    ' how long this department has had the round, so a reminder is a judged choice
    Dim v As Variant, n As Long
    v = SH("shDept").Cells(r, D_SENT).Value
    If Blank(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    n = DateDiff("d", Int(CDate(v)), Date)
    If n <= 0 Then
        DaysWaiting = "sent today"
    ElseIf n = 1 Then
        DaysWaiting = "1 day ago"
    Else
        DaysWaiting = n & " days ago"
    End If
End Function

Private Function AlreadyNudged(r As Long) As Boolean
    ' so a second run on the same day is a visible choice rather than a surprise
    Dim v As Variant
    v = SH("shDept").Cells(r, D_NUDGE).Value
    If Blank(v) Then Exit Function
    If Not IsDate(v) Then Exit Function
    AlreadyNudged = (Int(CDate(v)) = Date)
End Function

' =====================================================================
'  reminder
' =====================================================================
Public Sub Btn_Remind()
    Dim ed As String
    Dim app As Object, ap As Object, pat As Object, d As Date, step_ As Long
    Dim rh As Long, rm As Long
    On Error GoTo Failed
    Fresh
    Set app = OutlookApp()
    If app Is Nothing Then
        MsgBox TX("dlg_no_outlook"), vbExclamation, TX("app_title")
        Exit Sub
    End If
    step_ = StepWeeks()
    d = DueThursday()
    Do While Weekday(d, vbSunday) <> vbThursday      ' the mask says Thursday, so the
        d = d + 1                                    ' start date has to be one too
    Loop
    If Not HourMin(SetVal(14), rh, rm) Then
        rh = 9: rm = 0
    End If
    If DateSerial(Year(d), Month(d), Day(d)) + TimeSerial(rh, rm, 0) <= Now Then
        d = d + 7 * step_
    End If
    If ReminderExists(app) Then
        MsgBox TX("dlg_rem_exists"), vbInformation, TX("app_title")
        Exit Sub
    End If
    Set ap = app.CreateItem(1)                          ' olAppointmentItem
    ap.Subject = TX("rem_title")
    ap.Body = TX("rem_body")
    ap.Start = DateSerial(Year(d), Month(d), Day(d)) + TimeSerial(rh, rm, 0)
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
    ed = Err.Description                      ' read first - TX clears it
    MsgBox TX("err_remind") & ": " & ed, vbExclamation, TX("app_title")
End Sub

Private Function ReminderExists(app As Object) As Boolean
    Dim items As Object, it As Object, want As String, n As Long
    want = Norm(TX("rem_title"))
    On Error Resume Next
    Set items = app.GetNamespace("MAPI").GetDefaultFolder(9).Items     ' olFolderCalendar
    If items Is Nothing Then
        Err.Clear
        Exit Function
    End If
    For Each it In items
        n = n + 1
        If n > 2000 Then Exit For
        If Norm(CStr(it.Subject)) = want Then
            ReminderExists = True
            Exit For
        End If
        Err.Clear
    Next it
    Err.Clear
    On Error GoTo 0
End Function

' =====================================================================
'  refresh the two overview sheets
' =====================================================================
Public Sub Btn_Refresh()
    Dim en As Long, ed As String
    On Error GoTo Failed
    gStep = "refresh"
    Fresh
    Application.ScreenUpdating = False
    Refresh "manual"
    Application.ScreenUpdating = True
    SaveBook
    Say ""
    Exit Sub
Failed:
    en = Err.Number: ed = Err.Description     ' read first - Say and TX both clear it
    Application.ScreenUpdating = True
    Say ""
    MsgBox TX("err_general") & " [" & gStep & "] " & en & ": " & ed, _
           vbExclamation, TX("app_title")
End Sub

Private Sub Refresh(src As String)
    Dim ws As Worksheet, main As Worksheet, tr As Worksheet, r As Long, out As Long
    gStep = "refresh:start"
    Dim who As String, due As Date, linked As String, rounds As Long
    Dim prevDue As Date, late As Boolean, state_ As String
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
    prevDue = due - 7 * StepWeeks()
    late = False
    If LastSent() > DateSerial(2020, 1, 1) Then late = (LastSent() < prevDue)
    If late Then
        main.Range("C9").Value = ArabicDate(due) & "  -  " & TX("overdue")
    Else
        main.Range("C9").Value = ArabicDate(due)
    End If
    If late Then
        main.Range("C9").Font.Color = RGB(192, 57, 43)
    Else
        main.Range("C9").Font.Color = C_NAVY
    End If
    main.Range("B42").Value = Replace(TX("lbl_updated"), "{date}", Format$(Now, "yyyy/mm/dd hh:mm"))

    gStep = "refresh:clear"
    main.Range("B16:G41").Clear
    gStep = "refresh:track"
    WriteRound tr          ' the follow-up sheet keeps every past round - never cleared
    gStep = "refresh:blocks"
    EnsureBlocks
    gStep = "refresh:rows"
    out = 15
    rounds = 0
    For r = D_FIRST To LastDeptRow()
        If Trim$(CStr(ws.Cells(r, D_NAME).Value)) <> "" And out < 41 Then
            out = out + 1
            rounds = rounds + 1
            state_ = Replied(r)                 ' once per row, not four times
            If Blank(ws.Cells(r, D_CONV).Value) And Blank(ws.Cells(r, D_ENTRY).Value) Then
                linked = TX("w_notlinked")
            Else
                linked = TX("w_linked")
            End If
            main.Range(main.Cells(out, 2), main.Cells(out, 7)).Value = Array( _
                ws.Cells(r, D_NAME).Value, ws.Cells(r, D_OWNER).Value, ws.Cells(r, D_STATUS).Value, _
                linked, ws.Cells(r, D_SENT).Value, state_)
            StyleRow main, out, 2, 7
            main.Cells(out, 6).NumberFormat = "yyyy/mm/dd"
            If linked = TX("w_notlinked") Then main.Cells(out, 5).Font.Color = RGB(192, 57, 43)
            If state_ = TX("st_replied") Then
                main.Cells(out, 7).Interior.Color = RGB(228, 242, 211)
            ElseIf state_ = TX("st_waiting") Then
                main.Cells(out, 7).Interior.Color = RGB(250, 212, 208)
            End If

        End If
    Next r
End Sub

Private Sub WriteRound(tr As Worksheet)
    ' The follow-up sheet is a stack of rounds, newest on top, and nothing on it is
    ' ever erased. Each round is a block: a dark title row, a header row, then one row
    ' per department. Column A holds a hidden key so the block can be found again:
    '   HDR|<round>|<yyyymmdd>   the title row
    '   RND|<round>|<department> a department row
    ' Only the current round's rows are rewritten. Older blocks are left untouched,
    ' which is what makes this sheet both today's worklist and the archive.
    Dim ws As Worksheet, r As Long, top As Long, rowN As Long, n As Long
    Dim rnd_ As Long, sentD As Date, found As Boolean
    On Error GoTo Bail
    Set ws = SH("shDept")
    MigrateTrack tr
    sentD = RoundDate()
    If sentD < DateSerial(2000, 1, 1) Then Exit Sub      ' nothing has been sent yet
    rnd_ = RoundNo(tr, sentD, top)
    If top = 0 Then
        top = NewBlock(tr, rnd_, sentD)
    End If

    rowN = top + 2                                      ' first department row
    n = 0
    For r = D_FIRST To LastDeptRow()
        If Trim$(CStr(ws.Cells(r, D_NAME).Value)) <> "" Then
            n = n + 1
            If n > 40 Then Exit For
            If rowN > tr.Rows.Count - 4 Then Exit For
            ' write over this round's own row; push anything else down
            If Not Mine(tr, rowN, rnd_) Then tr.Rows(rowN).Insert
            BlockRow tr, rowN, rnd_, r
            rowN = rowN + 1
        End If
    Next r
    ' rows left over from when this block held more departments than it does now
    Do While Mine(tr, rowN, rnd_)
        tr.Rows(rowN).Delete
    Loop
    tr.Columns(1).Hidden = True
    Exit Sub
Bail:
    Err.Clear
End Sub

Private Function Mine(tr As Worksheet, rowN As Long, rnd_ As Long) As Boolean
    ' does this sheet row belong to the round being written?
    Mine = (InStr(1, CStr(tr.Cells(rowN, 1).Value), "RND|" & rnd_ & "|") = 1)
End Function

Private Sub MigrateTrack(tr As Worksheet)
    ' The follow-up sheet used to be one flat table with a header on row 5. The first
    ' time the new layout runs, that old table is cleared away - its contents describe
    ' the current round only, and the round is about to be written again as a block.
    Dim r As Long, last As Long
    On Error Resume Next
    last = tr.Cells(tr.Rows.Count, 2).End(xlUp).Row
    If last < 5 Then
        Err.Clear
        Exit Sub
    End If
    For r = 5 To last
        If Left$(CStr(tr.Cells(r, 1).Value), 4) = "HDR|" Then
            Err.Clear
            Exit Sub                       ' already migrated
        End If
    Next r
    tr.Range(tr.Cells(5, 1), tr.Cells(last + 2, 10)).Clear
    Err.Clear
End Sub

Private Function RoundDate() As Date
    ' the day this round went out: the newest send date on the departments sheet
    Dim ws As Worksheet, r As Long, v As Variant, best As Date
    Set ws = SH("shDept")
    For r = D_FIRST To LastDeptRow()
        v = ws.Cells(r, D_SENT).Value
        If Not Blank(v) Then
            If IsDate(v) Then
                If CDate(v) > best Then best = CDate(v)
            End If
        End If
    Next r
    RoundDate = best
End Function

Private Function RoundNo(tr As Worksheet, sentD As Date, ByRef top As Long) As Long
    ' Finds the block for this round, or works out the next round number.
    ' A send within six days of the newest block joins that block, so a colleague who
    ' builds her drafts a day later lands in the same round rather than starting one.
    Dim r As Long, k As String, hd As Date, hi As Long, last As Long
    top = 0
    last = tr.Cells(tr.Rows.Count, 1).End(xlUp).Row
    If last < 5 Then last = 5
    For r = 5 To last
        k = CStr(tr.Cells(r, 1).Value)
        If Left$(k, 4) = "HDR|" Then
            hi = HdrNo(k)
            hd = HdrDate(k)
            If hi > RoundNo Then RoundNo = hi
            If top = 0 And hd > DateSerial(2000, 1, 1) Then
                If Abs(DateDiff("d", Int(hd), Int(sentD))) <= 6 Then
                    top = r
                    RoundNo = hi
                    Exit Function
                End If
            End If
        End If
    Next r
    RoundNo = RoundNo + 1
End Function

Private Function HdrNo(k As String) As Long
    Dim p As Variant
    p = Split(k, "|")
    If UBound(p) >= 1 Then
        If IsNumeric(p(1)) Then HdrNo = CLng(p(1))
    End If
End Function

Private Function HdrDate(k As String) As Date
    Dim p As Variant, t As String
    p = Split(k, "|")
    If UBound(p) < 2 Then Exit Function
    t = CStr(p(2))
    If Len(t) <> 8 Then Exit Function
    On Error Resume Next
    HdrDate = DateSerial(CLng(Left$(t, 4)), CLng(Mid$(t, 5, 2)), CLng(Right$(t, 2)))
    Err.Clear
    On Error GoTo 0
End Function

Private Function NewBlock(tr As Worksheet, rnd_ As Long, sentD As Date) As Long
    ' a fresh block, inserted above everything, so the current round is always on top
    Dim top As Long, i As Long
    CloseBlock tr
    top = 5
    tr.Rows(top & ":" & (top + 1)).Insert
    tr.Cells(top, 1).Value = "HDR|" & rnd_ & "|" & Format$(sentD, "yyyymmdd")
    With tr.Range(tr.Cells(top, 2), tr.Cells(top, 7))
        .Merge
        .Value = TX("w_round") & " " & rnd_ & "  -  " & DayWord(sentD) & " " & ArabicDate(sentD)
        .Interior.Color = C_NAVY
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
    End With
    tr.Rows(top).RowHeight = 26
    tr.Cells(top + 1, 1).Value = "COL|" & rnd_
    tr.Range(tr.Cells(top + 1, 2), tr.Cells(top + 1, 7)).Value = Array( _
        TX("h_dept"), TX("h_rep"), TX("h_reply"), TX("h_days"), TX("h_nudge"), TX("h_state"))
    With tr.Range(tr.Cells(top + 1, 2), tr.Cells(top + 1, 7))
        .Interior.Color = RGB(57, 68, 92)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With
    For i = 2 To 7
        tr.Columns(i).ColumnWidth = IIf(i = 2, 30, IIf(i = 4, 24, 16))
    Next i
    NewBlock = top
End Function

Private Sub CloseBlock(tr As Worksheet)
    ' A round that has just ended is settled: a department still showing "due a
    ' reminder" or "waiting" never answered it, and the archive should say so plainly
    ' rather than keep a status that only made sense while the round was open.
    Dim r As Long, k As String, st As String, last As Long
    On Error Resume Next
    last = tr.Cells(tr.Rows.Count, 1).End(xlUp).Row
    For r = 5 To last
        k = CStr(tr.Cells(r, 1).Value)
        If Left$(k, 4) = "HDR|" And r > 5 Then Exit For    ' only the topmost block
        If Left$(k, 4) = "RND|" Then
            st = Trim$(CStr(tr.Cells(r, 7).Value))
            If st = TX("st_duenudge") Or st = TX("st_waiting") Then
                tr.Cells(r, 7).Value = TX("st_noreply")
                Paint tr.Cells(r, 7), TX("st_noreply")
            End If
        End If
    Next r
    Err.Clear
End Sub

Private Sub BlockRow(tr As Worksheet, rowN As Long, rnd_ As Long, r As Long)
    ' one department's line in the current block
    Dim ws As Worksheet, st As String, ans As Variant, days As Variant, ndg As Variant
    Set ws = SH("shDept")
    st = Replied(r)
    ans = AnswerAt(r)
    days = ""
    If Not IsEmpty(ans) Then days = WorkDays(CDate(ws.Cells(r, D_SENT).Value), CDate(ans))
    ndg = ws.Cells(r, D_NUDGE).Value
    If Blank(ndg) Then ndg = ""
    tr.Cells(rowN, 1).Value = "RND|" & rnd_ & "|" & ws.Cells(r, D_NAME).Value
    tr.Range(tr.Cells(rowN, 2), tr.Cells(rowN, 7)).Value = Array( _
        ws.Cells(r, D_NAME).Value, RepShown(r), _
        IIf(IsEmpty(ans), "", ans), days, ndg, st)
    StyleRow tr, rowN, 2, 7
    tr.Cells(rowN, 4).NumberFormat = "yyyy/mm/dd  hh:mm"
    tr.Cells(rowN, 6).NumberFormat = "yyyy/mm/dd"
    tr.Range(tr.Cells(rowN, 5), tr.Cells(rowN, 7)).HorizontalAlignment = xlCenter
    Paint tr.Cells(rowN, 7), st
End Sub

Private Sub Paint(c As Range, st As String)
    ' the status cell carries the colour, so the sheet reads at a glance
    If st = TX("st_replied") Then
        c.Interior.Color = RGB(228, 242, 211)
        c.Font.Color = RGB(63, 107, 42)
    ElseIf st = TX("st_repnoatt") Then
        c.Interior.Color = RGB(255, 246, 224)
        c.Font.Color = RGB(138, 98, 0)
    ElseIf st = TX("st_duenudge") Then
        c.Interior.Color = RGB(255, 228, 196)
        c.Font.Color = RGB(154, 79, 0)
        c.Font.Bold = True
    ElseIf st = TX("st_waiting") Or st = TX("st_noreply") Then
        c.Interior.Color = RGB(253, 233, 230)
        c.Font.Color = RGB(156, 51, 39)
    Else
        c.Font.Color = RGB(120, 128, 140)
    End If
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
    ' One of six states, and every one of them says what to do next:
    '   moved aside   - the department is on hold
    '   not sent yet  - the round has not gone out to it
    '   replied       - the deck arrived
    '   replied, no deck - they answered but sent no PowerPoint, so it is not done
    '   waiting       - no answer, but a reminder has already gone today
    '   due a reminder- no answer, no reminder today: this is the one to act on
    Dim ws As Worksheet, sentV As Variant, repV As Variant, attV As Variant, ndgV As Variant
    Set ws = SH("shDept")
    If Norm(CStr(ws.Cells(r, D_STATUS).Value)) = Norm(TX("w_hold")) Then
        Replied = TX("st_hold")
        Exit Function
    End If
    sentV = ws.Cells(r, D_SENT).Value
    If Blank(sentV) Or Not IsDate(sentV) Then
        Replied = TX("st_notsent")
        Exit Function
    End If
    attV = ws.Cells(r, D_ATT).Value
    If Not Blank(attV) Then
        If IsDate(attV) Then
            If CDate(attV) >= CDate(sentV) Then
                Replied = TX("st_replied")
                Exit Function
            End If
        End If
    End If
    repV = ws.Cells(r, D_REPLY).Value
    If Not Blank(repV) Then
        If IsDate(repV) Then
            If CDate(repV) >= CDate(sentV) Then
                Replied = TX("st_repnoatt")
                Exit Function
            End If
        End If
    End If
    ndgV = ws.Cells(r, D_NUDGE).Value
    If Not Blank(ndgV) Then
        If IsDate(ndgV) Then
            If Int(CDate(ndgV)) >= Int(CDate(sentV)) Then
                Replied = TX("st_waiting")
                Exit Function
            End If
        End If
    End If
    Replied = TX("st_duenudge")
End Function

Private Function NeedsNudge(r As Long) As Boolean
    ' A department needs a reminder when it has been sent the round, has NOT sent the
    ' deck, and has not already been reminded today. A department that answered
    ' without the PowerPoint still needs one - an acknowledgement is not the update.
    Dim st As String
    st = Replied(r)
    If st = TX("st_duenudge") Then NeedsNudge = True
    If st = TX("st_repnoatt") Then NeedsNudge = Not NudgedToday(r)
End Function

Private Function NudgedToday(r As Long) As Boolean
    Dim v As Variant
    v = SH("shDept").Cells(r, D_NUDGE).Value
    If Blank(v) Or Not IsDate(v) Then Exit Function
    NudgedToday = (Int(CDate(v)) = Date)
End Function

Private Function AnswerAt(r As Long) As Variant
    ' when the department answered: the deck reply if there is one, else the plain one
    Dim ws As Worksheet, v As Variant
    Set ws = SH("shDept")
    v = ws.Cells(r, D_ATT).Value
    If Blank(v) Or Not IsDate(v) Then v = ws.Cells(r, D_REPLY).Value
    If Blank(v) Or Not IsDate(v) Then Exit Function
    AnswerAt = CDate(v)
End Function


Private Sub EnsureBlocks()
    ' every department gets its own To/Cc block; new departments are appended
    On Error Resume Next
    Dim ls As Worksheet, ds As Worksheet, r As Long, last As Long, have As String, nm As String, k As String
    Set ls = SH("shList")
    Set ds = SH("shDept")
    If ls Is Nothing Or ds Is Nothing Then
        Err.Clear
        Exit Sub
    End If
    have = ";"
    last = ls.Cells(ls.Rows.Count, 1).End(xlUp).Row
    For r = 5 To last
        k = CStr(ls.Cells(r, 1).Value)
        If Left$(k, 5) = "LIST|" Then have = have & Norm(Mid$(k, 6)) & ";"
    Next r
    For r = D_FIRST To LastDeptRow()
        nm = Trim$(CStr(ds.Cells(r, D_NAME).Value))
        If nm <> "" Then
            If InStr(have, ";" & Norm(nm) & ";") = 0 Then
                AddBlock ls, nm
                have = have & Norm(nm) & ";"
            End If
        End If
    Next r
End Sub

Private Sub AddBlock(ls As Worksheet, nm As String)
    On Error Resume Next
    Dim r As Long, j As Long
    r = ls.Cells(ls.Rows.Count, 1).End(xlUp).Row
    If r < 5 Then r = 5 Else r = r + 14
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
