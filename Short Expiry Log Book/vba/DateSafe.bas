Attribute VB_Name = "DateSafe"
Option Explicit
' =====================================================================
'  DateSafe  -  locale-independent date handling for the Short Expiry
'               Log Book.
'
'  WHY THIS MODULE EXISTS
'  ----------------------
'  A date typed as "11/1/2026" is "11 Jan" on a d/m/y PC and "1 Nov" on
'  an m/d/y PC.  The old form accepted any text that IsDate() liked and
'  then assigned that TEXT to the cell.  Excel coerces text to a date
'  using en-US (m/d/y) first and the PC's regional format second, so
'  the same keystrokes produced different serial numbers on different
'  PCs, and text that neither rule could read stayed as text.
'
'  RULES ENFORCED HERE
'  -------------------
'  1. Only UNAMBIGUOUS input is accepted.  Anything that could be read
'     two ways ("11/1/2026") is rejected with a message telling the user
'     how to type it.
'  2. A date is always built with DateSerial(y, m, d) from explicitly
'     identified parts.  CDate / IsDate / Format-to-string are never used
'     to *interpret* user input.
'  3. A date is always written to a cell as a Double serial via
'     Range.Value2, never as text, so no coercion can happen.
'  4. Cells are given the display format "dd-mmm-yyyy" so every PC shows
'     the same thing regardless of its regional settings.
' =====================================================================

Public Const EXPIRY_NUMBER_FORMAT As String = "dd-mmm-yyyy"
Public Const STAMP_NUMBER_FORMAT As String = "dd-mmm-yyyy hh:mm"

' Sanity window for an expiry date (used by ExpiryPolicyWarning).
Public Const EXPIRY_MAX_YEARS_AHEAD As Long = 10

Public Const HOW_TO_TYPE As String = _
    "Type the expiry as one of:" & vbCrLf & _
    "   01112026        (ddmmyyyy, 8 digits)" & vbCrLf & _
    "   01-Nov-2026     (day-month name-year)" & vbCrLf & _
    "   2026-11-01      (yyyy-mm-dd)" & vbCrLf & _
    "   Nov-2026  or  11/2026   (month only = last day of month)"

' ---------------------------------------------------------------------
'  Month names
' ---------------------------------------------------------------------
Public Function MonthAbbr(ByVal m As Long) As String
    MonthAbbr = Choose(m, "Jan", "Feb", "Mar", "Apr", "May", "Jun", _
                          "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
End Function

Private Function MonthFull(ByVal m As Long) As String
    MonthFull = Choose(m, "January", "February", "March", "April", "May", "June", _
                          "July", "August", "September", "October", "November", "December")
End Function

' Returns 1-12 for an English month name / abbreviation (3+ letters,
' e.g. "nov", "NOV", "Novem", "November"), or 0 if not recognised.
' Also accepts the Office display-language month names as a courtesy.
Public Function MonthFromName(ByVal s As String) As Long
    Dim t As String, i As Long
    t = LCase$(Trim$(s))
    If Len(t) < 3 Then Exit Function
    For i = 1 To 12
        If Left$(LCase$(MonthFull(i)), Len(t)) = t Then
            MonthFromName = i
            Exit Function
        End If
    Next i
    ' Excel display language (e.g. Chinese UI) - exact match only
    For i = 1 To 12
        If LCase$(MonthName(i, True)) = t Or LCase$(MonthName(i)) = t Then
            MonthFromName = i
            Exit Function
        End If
    Next i
End Function

' ---------------------------------------------------------------------
'  Building / checking dates
' ---------------------------------------------------------------------
' True when y/m/d is a real calendar date (no DateSerial roll-over,
' e.g. 31 Feb becoming 3 Mar).
Public Function TryMakeDate(ByVal y As Long, ByVal m As Long, ByVal d As Long, ByRef result As Date) As Boolean
    If y < 1900 Or y > 9999 Then Exit Function
    If m < 1 Or m > 12 Then Exit Function
    If d < 1 Or d > 31 Then Exit Function
    result = DateSerial(y, m, d)
    TryMakeDate = (Year(result) = y And Month(result) = m And Day(result) = d)
End Function

Public Function LastDayOfMonth(ByVal y As Long, ByVal m As Long) As Date
    LastDayOfMonth = DateSerial(y, m + 1, 1) - 1
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    IsAllDigits = True
End Function

Private Function IsAllLetters(ByVal s As String) As Boolean
    Dim i As Long, c As String
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        c = LCase$(Mid$(s, i, 1))
        If c < "a" Or c > "z" Then Exit Function
    Next i
    IsAllLetters = True
End Function

' 2-digit years are taken as 20xx.
Private Function ExpandYear(ByVal y As Long, ByVal digits As Long) As Long
    If digits = 2 Then ExpandYear = 2000 + y Else ExpandYear = y
End Function

' ---------------------------------------------------------------------
'  TryParseExpiry  -  the ONLY entry point for interpreting typed dates
' ---------------------------------------------------------------------
' Accepts (all case-insensitive, separators / - . or space):
'     01112026            ddmmyyyy (8 digits)
'     01-Nov-2026         d Mon y      (Mon = 3+ letters, y = 2 or 4 digits)
'     Nov-01-2026         Mon d y
'     2026-11-01          yyyy-mm-dd
'     Nov-2026            Mon yyyy      -> last day of month
'     11/2026, 2026-11    mm/yyyy, yyyy-mm -> last day of month
'     20/6/2021           numeric d/m/y or m/d/y ONLY when one reading is
'                         impossible (a part > 12).  "11/1/2026" is rejected.
'
' assumeDMY:=True is for repairing legacy cells only - it resolves an
' ambiguous numeric date as day/month/year (the site convention).  The
' entry form must NEVER pass True.
Public Function TryParseExpiry(ByVal txt As String, ByRef result As Date, ByRef errMsg As String, _
                               Optional ByVal assumeDMY As Boolean = False) As Boolean
    Dim s As String, parts() As String, n As Long, i As Long
    Dim d As Long, m As Long, y As Long
    Dim p1 As String, p2 As String, p3 As String

    errMsg = ""
    s = Trim$(txt)
    ' normalise separators to a single "-"
    s = Replace(s, "/", "-")
    s = Replace(s, ".", "-")
    s = Replace(s, " ", "-")
    s = Replace(s, "\", "-")
    Do While InStr(s, "--") > 0
        s = Replace(s, "--", "-")
    Loop
    If Left$(s, 1) = "-" Then s = Mid$(s, 2)
    If Right$(s, 1) = "-" Then s = Left$(s, Len(s) - 1)

    If Len(s) = 0 Then
        errMsg = "Expiry date is empty."
        Exit Function
    End If

    ' ---- pure digits ----------------------------------------------------
    If IsAllDigits(s) Then
        Select Case Len(s)
            Case 8
                d = CLng(Left$(s, 2)): m = CLng(Mid$(s, 3, 2)): y = CLng(Right$(s, 4))
                If TryMakeDate(y, m, d, result) Then
                    TryParseExpiry = True
                Else
                    errMsg = "'" & txt & "' is not a valid ddmmyyyy date."
                End If
            Case 6
                errMsg = "'" & txt & "' is ambiguous (ddmmyy or yymmdd?). Use the 4-digit year: ddmmyyyy."
            Case Else
                errMsg = "'" & txt & "' is not a date. " & HOW_TO_TYPE
        End Select
        Exit Function
    End If

    parts = Split(s, "-")
    n = UBound(parts) + 1
    For i = 0 To n - 1
        parts(i) = Trim$(parts(i))
        If Len(parts(i)) = 0 Then
            errMsg = "'" & txt & "' is not a date. " & HOW_TO_TYPE
            Exit Function
        End If
    Next i

    Select Case n
        ' ---- two parts: month + year -------------------------------------
        Case 2
            p1 = parts(0): p2 = parts(1)
            If IsAllLetters(p1) And IsAllDigits(p2) And Len(p2) = 4 Then
                m = MonthFromName(p1): y = CLng(p2)
            ElseIf IsAllDigits(p1) And IsAllDigits(p2) And Len(p2) = 4 And Len(p1) <= 2 Then
                m = CLng(p1): y = CLng(p2)
            ElseIf IsAllDigits(p1) And Len(p1) = 4 And IsAllDigits(p2) And Len(p2) <= 2 Then
                y = CLng(p1): m = CLng(p2)
            ElseIf IsAllDigits(p1) And Len(p1) = 4 And IsAllLetters(p2) Then
                y = CLng(p1): m = MonthFromName(p2)
            Else
                errMsg = "'" & txt & "' is not a date. " & HOW_TO_TYPE
                Exit Function
            End If
            If m < 1 Or m > 12 Or y < 1900 Or y > 9999 Then
                errMsg = "'" & txt & "' has an invalid month or year."
                Exit Function
            End If
            result = LastDayOfMonth(y, m)
            TryParseExpiry = True

        ' ---- three parts ---------------------------------------------------
        Case 3
            p1 = parts(0): p2 = parts(1): p3 = parts(2)

            If IsAllDigits(p1) And Len(p1) = 4 And IsAllDigits(p2) And IsAllDigits(p3) Then
                ' yyyy-mm-dd
                y = CLng(p1): m = CLng(p2): d = CLng(p3)

            ElseIf IsAllDigits(p1) And IsAllLetters(p2) And IsAllDigits(p3) Then
                ' d-Mon-y
                d = CLng(p1): m = MonthFromName(p2): y = ExpandYear(CLng(p3), Len(p3))
                If m = 0 Then errMsg = "'" & p2 & "' is not a month name.": Exit Function

            ElseIf IsAllLetters(p1) And IsAllDigits(p2) And IsAllDigits(p3) Then
                ' Mon-d-y
                m = MonthFromName(p1): d = CLng(p2): y = ExpandYear(CLng(p3), Len(p3))
                If m = 0 Then errMsg = "'" & p1 & "' is not a month name.": Exit Function

            ElseIf IsAllDigits(p1) And IsAllDigits(p2) And IsAllDigits(p3) Then
                ' numeric a-b-y : accept only if exactly one reading is possible
                Dim a As Long, b As Long, okDMY As Boolean, okMDY As Boolean
                Dim dDMY As Date, dMDY As Date
                If Len(p3) <> 2 And Len(p3) <> 4 Then
                    errMsg = "'" & txt & "' - the year must be 2 or 4 digits. " & HOW_TO_TYPE
                    Exit Function
                End If
                a = CLng(p1): b = CLng(p2): y = ExpandYear(CLng(p3), Len(p3))
                okDMY = TryMakeDate(y, b, a, dDMY)
                okMDY = TryMakeDate(y, a, b, dMDY)
                If okDMY And okMDY And a <> b Then
                    If assumeDMY Then
                        result = dDMY
                        TryParseExpiry = True
                    Else
                        errMsg = "'" & txt & "' is AMBIGUOUS - it could be " & _
                                 ExpiryText(dDMY) & " or " & ExpiryText(dMDY) & "." & vbCrLf & HOW_TO_TYPE
                    End If
                ElseIf okDMY Then
                    result = dDMY
                    TryParseExpiry = True
                ElseIf okMDY Then
                    result = dMDY
                    TryParseExpiry = True
                Else
                    errMsg = "'" & txt & "' is not a valid date."
                End If
                Exit Function
            Else
                errMsg = "'" & txt & "' is not a date. " & HOW_TO_TYPE
                Exit Function
            End If

            If TryMakeDate(y, m, d, result) Then
                TryParseExpiry = True
            Else
                errMsg = "'" & txt & "' is not a valid calendar date."
            End If

        Case Else
            errMsg = "'" & txt & "' is not a date. " & HOW_TO_TYPE
    End Select
End Function

' ---------------------------------------------------------------------
'  Canonical text  (round-trips through TryParseExpiry on any locale)
' ---------------------------------------------------------------------
Public Function ExpiryText(ByVal d As Date) As String
    ExpiryText = Format$(Day(d), "00") & "-" & MonthAbbr(Month(d)) & "-" & Format$(Year(d), "0000")
End Function

' Friendly preview for the entry form, e.g. "Sun 01-Nov-2026  (424 days)"
Public Function ExpiryPreview(ByVal d As Date) As String
    Dim diff As Long
    diff = CLng(d) - CLng(Date)
    ExpiryPreview = Choose(Weekday(d, vbSunday), "Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat") & _
                    " " & ExpiryText(d)
    If diff < 0 Then
        ExpiryPreview = ExpiryPreview & "  (EXPIRED " & Abs(diff) & " days ago)"
    ElseIf diff = 0 Then
        ExpiryPreview = ExpiryPreview & "  (expires TODAY)"
    Else
        ExpiryPreview = ExpiryPreview & "  (" & diff & " days)"
    End If
End Function

' Returns "" when the date is plausible, otherwise a warning the caller
' should show and ask the user to confirm.
Public Function ExpiryPolicyWarning(ByVal d As Date) As String
    If d < Date Then
        ExpiryPolicyWarning = ExpiryText(d) & " is already in the past."
    ElseIf d > DateSerial(Year(Date) + EXPIRY_MAX_YEARS_AHEAD, Month(Date), Day(Date)) Then
        ExpiryPolicyWarning = ExpiryText(d) & " is more than " & EXPIRY_MAX_YEARS_AHEAD & " years away."
    End If
End Function

' ---------------------------------------------------------------------
'  Reading / writing cells
' ---------------------------------------------------------------------
' Writes a true date serial - never text - and fixes the display format.
Public Sub WriteDateCell(ByVal cell As Range, ByVal d As Date, Optional ByVal fmt As String = EXPIRY_NUMBER_FORMAT)
    cell.NumberFormat = fmt
    cell.Value2 = CDbl(d)
End Sub

Public Sub WriteStampCell(ByVal cell As Range, ByVal t As Date)
    WriteDateCell cell, t, STAMP_NUMBER_FORMAT
End Sub

' True when the cell already holds a real date (serial), False for text/empty.
Public Function CellHasRealDate(ByVal cell As Range) As Boolean
    CellHasRealDate = (VarType(cell.Value) = vbDate)
End Function
