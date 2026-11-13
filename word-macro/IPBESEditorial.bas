Attribute VB_Name = "IPBESEditorial"
' =============================================================================
' IPBES editorial checks for Microsoft Word
'
' Macros (run from View > Macros, or add them to the Quick Access Toolbar):
'   IPBES_CheckDocument      Marks possible style issues with Word comments
'                            starting "[IPBES]" and opens a summary report.
'                            Checks the masthead (if any) and heading numbering.
'                            Does NOT change the text.
'   IPBES_SafeFixes          Applies a small set of safe corrections with
'                            Track Changes switched on, so every change can be
'                            reviewed, accepted or rejected.
'   IPBES_CheckVoice         Marks prescriptive wording ("must", "should"...)
'                            for review. Separate because reports legitimately
'                            report speakers' views.
'   IPBES_FixFormatting      Applies the IPBES layout: IPBES Word styles (copied
'                            from the add-in), A4 page set-up, headings
'                            (Heading 1-4 -> CH1-CH4), body paragraphs
'                            (numbered or not), table captions and table
'                            layout. Style and text changes are tracked.
'   IPBES_ApplyChangeList    Applies a change list from Copilot or Claude (a
'                            table with columns "Original" and "Revised") to the
'                            active document as tracked changes, keeping all
'                            formatting. Paste the table into a new Word
'                            document, keep it open, then run this macro from
'                            the document to correct.
'   IPBES_RemoveIPBESComments  Deletes only the comments added by these macros.
'
' With tools/add_ribbon.py the add-in also gets an "IPBES" tab on the ribbon
' that runs these macros (entry points IPBESRibbon_* below).
'
' Rules: ipbes-editorial/references/house-style.md (UN Editorial Manual
' plus IPBES practice from the Plenary reports IPBES-1 to IPBES-12).
' Agreed text (decisions, approved summaries for policymakers, quotations) must
' not be changed: reject tracked changes there and treat comments as queries.
'
' This file is plain ASCII on purpose: non-ASCII characters are built with
' ChrW() so that the VBA editor imports them correctly.
' =============================================================================
Option Explicit

Private Const TAG As String = "[IPBES] "
Private Const MAX_COMMENTS_PER_RULE As Long = 25

Private Type IpbesRule
    Category As String
    FindText As String
    UseWildcards As Boolean
    MatchCase As Boolean
    WholeWord As Boolean
    Message As String
    Replacement As String
    AutoFix As Boolean
End Type

Private Rules() As IpbesRule
Private RuleCount As Long

' Report lines collected during a check
Private ReportLines As Collection

' Built-in style names, looked up once per run (TargetStyle runs for every paragraph).
Private bnHeading1 As String, bnHeading2 As String, bnHeading3 As String, bnHeading4 As String
Private bnTitle As String, bnCaption As String, bnListPara As String, bnListBullet As String, bnListNumber As String
Private bnNormal As String, bnBodyText As String, bnBodyText2 As String, bnBodyText3 As String
Private bnPlainText As String, bnNormalIndent As String

' ---------------------------------------------------------------------------
' Characters
' ---------------------------------------------------------------------------
Private Function EnDash() As String: EnDash = ChrW(8211): End Function
Private Function EmDash() As String: EmDash = ChrW(8212): End Function
Private Function RSQuo() As String: RSQuo = ChrW(8217): End Function
Private Function FullName() As String
    FullName = "Intergovernmental Science-Policy Platform on Biodiversity and Ecosystem Services"
End Function

' ---------------------------------------------------------------------------
' Rule table
' ---------------------------------------------------------------------------
Private Sub AddRule(ByVal cat As String, ByVal findText As String, ByVal wild As Boolean, _
                    ByVal mCase As Boolean, ByVal whole As Boolean, ByVal msg As String, _
                    Optional ByVal repl As String = "", Optional ByVal doFix As Boolean = False)
    RuleCount = RuleCount + 1
    ReDim Preserve Rules(1 To RuleCount)
    With Rules(RuleCount)
        .Category = cat
        .FindText = findText
        .UseWildcards = wild
        .MatchCase = mCase
        .WholeWord = whole
        .Message = msg
        .Replacement = repl
        .AutoFix = doFix
    End With
End Sub

Private Sub BuildRules()
    Dim stems As Variant, endings As Variant, s As Variant, e As Variant
    RuleCount = 0
    Erase Rules

    ' --- Platform name -----------------------------------------------------
    AddRule "Platform name", "Intergovernmental Platform on Biodiversity and Ecosystem Services", False, False, False, _
        "Never shorten the Platform's name. Write: " & FullName() & ".", FullName(), True
    AddRule "Platform name", "Science Policy Platform", False, False, False, _
        "Hyphen missing: Science-Policy Platform.", "Science-Policy Platform", True
    AddRule "Platform name", "Science-policy Platform", False, True, False, _
        "Capitalize: Science-Policy Platform.", "Science-Policy Platform", True
    AddRule "Platform name", "Intergovernmental science-policy platform", False, True, False, _
        "The Platform's name is capitalized: " & FullName() & "."
    AddRule "Platform name", "IPBES's", False, True, False, _
        "No possessive on acronyms: 'the IPBES secretariat', 'the mandate of IPBES'."
    AddRule "Platform name", "IPBES' ", False, True, False, _
        "No possessive on acronyms: 'the IPBES secretariat', 'the mandate of IPBES'."

    ' --- Internal codes and abbreviations not used in documents --------------
    AddCode "BBA", "business and biodiversity assessment"
    AddCode "TCA", "transformative change assessment"
    AddCode "NXS", "nexus assessment"
    AddCode "IAS", "invasive alien species assessment"
    AddCode "SUA", "sustainable use assessment"
    AddCode "GA1", "global assessment"
    AddCode "GA2", "second global assessment"
    AddCode "MTA", "monitoring assessment"
    AddCode "SPC", "spatial planning and connectivity assessment"
    AddCode "LDR", "land degradation assessment"
    AddCode "PPA", "pollinators assessment"
    AddCode "SCM", "scenarios and models assessment"
    AddCode "SPM", "summary for policymakers"
    AddCode "SPMs", "summaries for policymakers"
    AddCode "NCP", "nature" & RSQuo() & "s contributions to people"
    AddCode "ILK", "Indigenous and local knowledge"
    AddCode "IPLC", "Indigenous Peoples and local communities"
    AddCode "IPLCs", "Indigenous Peoples and local communities"
    AddCode "MEP", "Multidisciplinary Expert Panel (the Panel)"
    AddCode "TSU", "technical support unit"

    ' --- Abbreviations -------------------------------------------------------
    AddRule "Abbreviations", "e.g. ", False, True, False, _
        "IPBES practice: 'e.g.,' with a comma, inside parentheses only; in running text write 'for example'.", "e.g., ", True
    AddRule "Abbreviations", "i.e. ", False, True, False, _
        "IPBES practice: 'i.e.,' with a comma, inside parentheses only; in running text write 'that is'.", "i.e., ", True
    AddRule "Abbreviations", "and/or", False, False, False, _
        "Avoid 'and/or': use 'and' or 'or' (keep it in agreed text)."
    AddRule "Abbreviations", "US$", False, True, False, _
        "Write '$' for United States dollars (e.g., $200,000), not 'US$'."
    AddRule "Abbreviations", "USD", False, True, True, _
        "Write '$' for United States dollars, not 'USD'."
    AddRule "Abbreviations", "H.E.", False, True, False, "No honorifics such as 'H.E.'."
    AddRule "Abbreviations", "Excellency", False, False, True, "No honorifics such as 'His/Her Excellency'."
    AddRule "Abbreviations", "Dr. ", False, True, False, _
        "Use 'Ms.' or 'Mr.' (or the full name at first mention), not 'Dr.'."
    AddRule "Abbreviations", "Mrs.", False, True, False, "Use 'Ms.' unless the person prefers 'Mrs.'."
    AddRule "Abbreviations", "U.N.", False, True, False, "Write 'United Nations' (no abbreviation in running text)."

    ' --- Numbers, dates and time --------------------------------------------
    AddRule "Numbers", "[0-9]%", True, False, False, _
        "Write 'per cent' in running text; '%' only in table headnotes."
    AddRule "Numbers", "percent", False, False, True, "Write 'per cent' (two words).", "per cent", True
    AddRule "Numbers", "<[0-9]{1,2}:[0-9]{2}>", True, False, False, _
        "Time format: '10.15 a.m.', '3 p.m.', not '10:15'."
    AddRule "Numbers", "<[0-9]{1,2}.00 [ap].m.", True, False, False, _
        "Time format: '10 a.m.', not '10.00 a.m.'."
    AddRule "Numbers", "<[0-9]{1,2}-[0-9]{1,2} [JFMASOND][a-z]@ [0-9]{4}", True, True, False, _
        "Date range: 'from 3 to 8 February 2026' in text, '3" & EnDash() & "8 February 2026' in mastheads (en dash, not hyphen)."
    AddRule "Numbers", "<[JFMASOND][a-z]@ [0-9]{1,2}, [0-9]{4}", True, True, False, _
        "Date format: '12 February 2026' (day month year, no comma)."
    AddRule "Numbers", "this year", False, False, False, "Give the year rather than 'this year'."
    AddRule "Numbers", "next year", False, False, False, "Give the year rather than 'next year'."
    AddRule "Numbers", "last year", False, False, False, "Give the year rather than 'last year'."

    ' --- Punctuation ---------------------------------------------------------
    AddRule "Punctuation", " " & EmDash() & " ", False, False, False, _
        "Use a spaced en dash for an aside, never an em dash.", " " & EnDash() & " ", True
    AddRule "Punctuation", EmDash(), False, False, False, _
        "Use an en dash, never an em dash."
    AddRule "Punctuation", " - ", False, False, False, _
        "Use a spaced en dash for an aside, not a hyphen.", " " & EnDash() & " ", True
    AddRule "Punctuation", "  ", False, False, False, "Double space.", " ", True
    AddRule "Punctuation", " [;:\?\!]", True, False, False, "No space before ; : ? !"
    AddRule "Punctuation", Chr(34), True, False, False, _
        "Use curly quotation marks (IPBES_SafeFixes converts them)."
    AddRule "Punctuation", Chr(39), True, False, False, _
        "Use curly apostrophes and quotation marks (IPBES_SafeFixes converts them)."

    ' --- Spelling and compounds ----------------------------------------------
    stems = Array("recognis", "organis", "emphasis", "mobilis", "synthesis", "prioritis", "summaris", _
                  "utilis", "finalis", "characteris", "realis", "minimis", "maximis", "criticis", _
                  "harmonis", "standardis", "operationalis", "optimis", "categoris", "authoris", _
                  "institutionalis", "mainstreamis", "capitalis", "globalis", "localis", "visualis", "conceptualis")
    endings = Array("e", "ed", "es", "ing", "ation", "ations")
    For Each s In stems
        For Each e In endings
            AddRule "Spelling", s & e, False, False, True, _
                "UN/Oxford spelling uses -ize: " & Left(s, Len(s) - 1) & "z" & e & ".", _
                Left(s, Len(s) - 1) & "z" & e, True
        Next e
    Next s
    AddRule "Spelling", "whilst", False, False, True, "Write 'while'.", "while", True
    AddRule "Spelling", "amongst", False, False, True, "Write 'among'.", "among", True
    AddRule "Spelling", "program", False, False, True, _
        "Write 'programme' (except a computer program)."
    AddRule "Spelling", "programs", False, False, True, _
        "Write 'programmes' (except computer programs)."
    AddRule "Spelling", "wellbeing", False, False, True, "Write 'well-being'.", "well-being", True
    AddRule "Spelling", "policy makers", False, False, False, "Write 'policymakers'.", "policymakers", True
    AddRule "Spelling", "policy-makers", False, False, False, "Write 'policymakers'.", "policymakers", True
    AddRule "Spelling", "decision-makers", False, False, False, _
        "Noun: 'decision makers' (open); adjective: 'decision-making'.", "decision makers", True
    AddRule "Spelling", "inter-sessional", False, False, False, "Write 'intersessional'.", "intersessional", True
    AddRule "Spelling", "sub-regional", False, False, False, "Write 'subregional'.", "subregional", True
    AddRule "Spelling", "sub-region", False, False, True, "Write 'subregion'.", "subregion", True
    AddRule "Spelling", "co-operation", False, False, False, "Write 'cooperation'.", "cooperation", True
    AddRule "Spelling", "e-mail", False, False, False, "Write 'email'.", "email", True
    AddRule "Spelling", "on-line", False, False, False, "Write 'online'.", "online", True
    AddRule "Spelling", "nongovernmental", False, False, False, "Write 'non-governmental'.", "non-governmental", True
    AddRule "Spelling", "capacity building", False, False, False, _
        "Write 'capacity-building' (noun and adjective).", "capacity-building", True
    AddRule "Spelling", "judgment", False, False, True, _
        "Write 'judgement' (general sense); 'judgment' only for court decisions."

    ' --- Capitalization and terms --------------------------------------------
    AddRule "Terms", "Summary for Policymakers", False, True, False, _
        "Lower case in running text: 'summary for policymakers' (capitals only in a published title)."
    AddRule "Terms", "Summaries for Policymakers", False, True, False, _
        "Lower case: 'summaries for policymakers'."
    AddRule "Terms", "Work Programme", False, True, False, "Lower case: 'work programme'."
    AddRule "Terms", "Task Force", False, True, False, "Lower case: 'task force'."
    AddRule "Terms", "Technical Support Unit", False, True, False, "Lower case: 'technical support unit'."
    AddRule "Terms", "Contributions to People", False, True, False, _
        "Lower case: 'nature" & RSQuo() & "s contributions to people'."
    AddRule "Terms", "the plenary", False, True, False, _
        "IPBES practice: 'the Plenary' (capitalized) for the governing body."
    AddRule "Terms", "Secretariat", False, True, True, _
        "Lower case 'the secretariat' (capital only for the United Nations Secretariat or another body's name)."
    AddRule "Terms", "Member States", False, True, False, _
        "IPBES has 'members' (IPBES members, members of the Platform); 'Member States' means States Members of the United Nations."
    AddRule "Terms", "indigenous peoples", False, True, False, _
        "Capitalize: 'Indigenous Peoples' (IPBES practice since IPBES-10).", "Indigenous Peoples", True
    AddRule "Terms", "spatial planning assessment", False, False, False, _
        "Short name: 'spatial planning and connectivity assessment'.", "spatial planning and connectivity assessment", True
    AddRule "Terms", "Indigenous peoples", False, True, False, _
        "Capitalize: 'Indigenous Peoples'.", "Indigenous Peoples", True
    AddRule "Terms", "indigenous and local knowledge", False, True, False, _
        "Capitalize: 'Indigenous and local knowledge'.", "Indigenous and local knowledge", True
    AddRule "Terms", "indigenous knowledge", False, True, False, "Capitalize: 'Indigenous knowledge'.", "Indigenous knowledge", True
    AddRule "Terms", "Co-chair", False, True, True, _
        "'Co-Chair' for specific people (the Co-Chairs of the assessment); 'co-chairs' in generic lists."

    ' --- Countries (UNTERM short names) --------------------------------------
    AddRule "Countries", "Turkey", False, True, True, "UN name: T" & ChrW(252) & "rkiye.", "T" & ChrW(252) & "rkiye", True
    AddRule "Countries", "Czech Republic", False, True, False, "UN short name: Czechia.", "Czechia", True
    AddRule "Countries", "Vietnam", False, True, True, "UN name: Viet Nam.", "Viet Nam", True
    AddRule "Countries", "Swaziland", False, True, True, "UN name: Eswatini.", "Eswatini", True
    AddRule "Countries", "Cape Verde", False, True, False, "UN name: Cabo Verde.", "Cabo Verde", True
    AddRule "Countries", "Ivory Coast", False, True, False, "UN name: C" & ChrW(244) & "te d" & RSQuo() & "Ivoire.", _
        "C" & ChrW(244) & "te d" & RSQuo() & "Ivoire", True
    AddRule "Countries", "South Korea", False, True, False, "UN name: Republic of Korea."
    AddRule "Countries", "North Korea", False, True, False, "UN name: Democratic People" & RSQuo() & "s Republic of Korea."
    AddRule "Countries", "Russia", False, True, True, "UN name: Russian Federation."
    AddRule "Countries", "Syria", False, True, True, "UN name: Syrian Arab Republic."
    AddRule "Countries", "Laos", False, True, True, "UN name: Lao People" & RSQuo() & "s Democratic Republic."
    AddRule "Countries", "Holland", False, True, True, "UN name: Netherlands (Kingdom of the)."
    AddRule "Countries", "UK", False, True, True, _
        "Write 'United Kingdom of Great Britain and Northern Ireland' at first mention, then 'United Kingdom'."
    AddRule "Countries", "USA", False, True, True, _
        "Write 'United States of America' at first mention, then 'United States'."
    AddRule "Countries", "U.S.", False, True, False, "Write 'United States'."
    AddRule "Countries", "Macedonia", False, True, True, "UN name: North Macedonia (check context)."

    ' --- Inclusive language --------------------------------------------------
    AddRule "Inclusive language", "Chairman", False, False, True, "Use 'Chair'."
    AddRule "Inclusive language", "Chairperson", False, False, True, "IPBES uses 'Chair'."
    AddRule "Inclusive language", "Chairwoman", False, False, True, "Use 'Chair'."
    AddRule "Inclusive language", "he/she", False, False, False, _
        "Avoid 'he/she': use the plural, 'who', or omit the pronoun."
    AddRule "Inclusive language", "his/her", False, False, False, _
        "Avoid 'his/her': use the plural, 'the', or omit the possessive."
    AddRule "Inclusive language", "him/her", False, False, False, "Avoid 'him/her': rephrase."
    AddRule "Inclusive language", "(s)he", False, False, False, "Avoid '(s)he': rephrase."
    AddRule "Inclusive language", "s/he", False, False, False, "Avoid 's/he': rephrase."
    AddRule "Inclusive language", "mankind", False, False, True, "Use 'humankind' or 'humanity'.", "humankind", True
    AddRule "Inclusive language", "manpower", False, False, True, "Use 'staffing', 'staff' or 'workforce'."
    AddRule "Inclusive language", "man-made", False, False, False, "Use 'human-made', 'human-caused' or 'anthropogenic'."
    AddRule "Inclusive language", "spokesman", False, False, True, "Use 'spokesperson'.", "spokesperson", True
    AddRule "Inclusive language", "fishermen", False, False, True, "Use 'fishers'.", "fishers", True
    AddRule "Inclusive language", "layman", False, False, True, "Use 'non-specialist' or 'layperson'."
    AddRule "Inclusive language", "the disabled", False, False, False, "Use 'persons with disabilities'."
    AddRule "Inclusive language", "disabled person", False, False, False, "Use 'person with a disability' (person-first)."
    AddRule "Inclusive language", "disabled people", False, False, False, "Use 'persons with disabilities'."
    AddRule "Inclusive language", "handicapped", False, False, True, "Use 'persons with disabilities'."
    AddRule "Inclusive language", "special needs", False, False, False, "Avoid euphemisms: 'persons with disabilities'."
    AddRule "Inclusive language", "differently abled", False, False, False, "Avoid euphemisms: 'persons with disabilities'."
    AddRule "Inclusive language", "wheelchair-bound", False, False, False, "Use 'wheelchair user'."
    AddRule "Inclusive language", "confined to a wheelchair", False, False, False, "Use 'wheelchair user'."
    AddRule "Inclusive language", "suffers from", False, False, False, "Use 'has' or 'with' (e.g., 'person with ...')."
    AddRule "Inclusive language", "suffering from", False, False, False, "Use 'with' (e.g., 'person with ...')."
    AddRule "Inclusive language", "mentally ill", False, False, False, "Use 'person with a psychosocial disability'."
    AddRule "Inclusive language", "the blind", False, False, True, "Use 'blind persons' or 'persons with visual impairments'."
    AddRule "Inclusive language", "the deaf", False, False, True, "Use 'deaf persons' or 'persons who are deaf'."
End Sub

Private Sub AddCode(ByVal code As String, ByVal expansion As String)
    AddRule "Internal codes", code, False, True, True, _
        "Internal abbreviation, not used in IPBES documents. Write the short name: '" & expansion & _
        "' (full title at first mention)."
End Sub

' ---------------------------------------------------------------------------
' Public macros
' ---------------------------------------------------------------------------
Public Sub IPBES_CheckDocument()
    Dim i As Long, n As Long, total As Long, sample As String, doc As Document
    Set doc = ActiveDocument
    If doc Is Nothing Then Exit Sub
    Set ReportLines = New Collection
    Application.ScreenUpdating = False
    BuildRules

    For i = 1 To RuleCount
        If i Mod 10 = 1 Then
            Application.StatusBar = "IPBES check: rule " & i & " of " & RuleCount & " (" & Rules(i).Category & ")..."
            DoEvents
        End If
        sample = ""
        n = MarkRule(doc, Rules(i), sample)
        If n > 0 Then
            total = total + n
            ReportLines.Add Rules(i).Category & vbTab & n & vbTab & _
                IIf(Rules(i).AutoFix, "yes", "") & vbTab & Rules(i).Message & vbTab & Left(sample, 60)
        End If
    Next i

    Application.StatusBar = "IPBES check: masthead..."
    total = total + CheckMasthead(doc)
    Application.StatusBar = "IPBES check: headings..."
    total = total + CheckHeadings(doc)
    Application.StatusBar = "IPBES check: table and figure captions..."
    total = total + CheckCaptions(doc)
    Application.StatusBar = False

    Application.ScreenUpdating = True
    WriteReport doc, total, "Editorial check"
End Sub

Public Sub IPBES_CheckVoice()
    Dim words As Variant, w As Variant, r As IpbesRule, n As Long, total As Long, sample As String
    Set ReportLines = New Collection
    words = Array("must", "should", "need to", "needs to", "it is essential", "it is crucial", "it is vital", _
                  "it is imperative", "urgently", "groundbreaking", "unprecedented", "call on", "calls on")
    Application.ScreenUpdating = False
    For Each w In words
        Application.StatusBar = "IPBES voice check: '" & w & "'..."
        DoEvents
        r.Category = "Voice"
        r.FindText = CStr(w)
        r.UseWildcards = False
        r.MatchCase = False
        r.WholeWord = True
        r.Message = "Voice: IPBES is policy-relevant, not policy-prescriptive. In IPBES-authored narrative prefer " & _
                    "'could', 'options include', 'the assessment found that'. Keep if reporting a speaker or in agreed text."
        sample = ""
        n = MarkRule(ActiveDocument, r, sample)
        If n > 0 Then
            total = total + n
            ReportLines.Add "Voice" & vbTab & n & vbTab & "" & vbTab & "'" & w & "'" & vbTab & ""
        End If
    Next w
    Application.StatusBar = False
    Application.ScreenUpdating = True
    WriteReport ActiveDocument, total, "Voice check"
End Sub

Public Sub IPBES_SafeFixes()
    Dim doc As Document, oldTrack As Boolean, i As Long, n As Long, total As Long, k As Long
    Dim msg As String, c As Long
    Set doc = ActiveDocument
    If MsgBox("IPBES safe fixes" & vbCr & vbCr & _
              "Track Changes will be switched on, so every change can be reviewed." & vbCr & _
              "Fixes: Platform name, e.g./i.e. comma, -ize spelling, dashes, double spaces," & vbCr & _
              "curly quotes, country names and a few terms." & vbCr & vbCr & _
              "Reject any change inside agreed text (decisions, quotations)." & vbCr & vbCr & _
              "Continue?", vbYesNo + vbQuestion, "IPBES") <> vbYes Then Exit Sub

    oldTrack = doc.TrackRevisions
    If Not SetTracking(doc, True) Then Exit Sub
    Application.ScreenUpdating = False
    BuildRules
    Set ReportLines = New Collection

    For i = 1 To RuleCount
        If i Mod 10 = 1 Then
            Application.StatusBar = "IPBES safe fixes: rule " & i & " of " & RuleCount & " (" & Rules(i).Category & ")..."
            DoEvents
        End If
        If Rules(i).AutoFix Then
            If Rules(i).FindText = "  " Then
                n = 0
                For k = 1 To 10            ' runs of 3+ spaces need several passes
                    c = ReplaceRule(doc, Rules(i))
                    If c = 0 Then Exit For
                    n = n + c
                Next k
            Else
                n = ReplaceRule(doc, Rules(i))
            End If
            If n > 0 Then
                total = total + n
                ReportLines.Add Rules(i).Category & vbTab & n & vbTab & "fixed" & vbTab & _
                    Rules(i).FindText & " -> " & Rules(i).Replacement & vbTab & ""
            End If
        End If
    Next i

    Application.StatusBar = "IPBES safe fixes: quotation marks and apostrophes..."
    DoEvents
    n = FixStraightQuotes(doc)
    Application.StatusBar = False
    If n > 0 Then
        total = total + n
        ReportLines.Add "Punctuation" & vbTab & n & vbTab & "fixed" & vbTab & "straight quotes -> curly" & vbTab & ""
    End If

    SetTracking doc, oldTrack
    Application.ScreenUpdating = True
    msg = total & " change(s) made with Track Changes." & vbCr & _
          "Review them (Review > Next) and reject any inside agreed text."
    MsgBox msg, vbInformation, "IPBES"
End Sub

Public Sub IPBES_RemoveIPBESComments()
    Dim i As Long, n As Long, doc As Document
    Set doc = ActiveDocument
    For i = doc.Comments.Count To 1 Step -1
        If Left(doc.Comments(i).Range.Text, Len(TAG)) = TAG Then
            doc.Comments(i).Delete
            n = n + 1
        End If
    Next i
    MsgBox n & " IPBES comment(s) removed.", vbInformation, "IPBES"
End Sub

Public Sub IPBES_ApplyChangeList()
    Dim target As Document, listDoc As Document, d As Document
    Dim items As Collection, it As Variant
    Dim orig As String, rev As String, ruleTxt As String, itemNo As String
    Dim applied As Long, notFound As Long, ambiguous As Long, skipped As Long, res As Long
    Dim oldTrack As Boolean, candidates As Collection, names As String, choice As String, i As Long, iItem As Long

    Set target = ActiveDocument
    If HasChangeList(target) Then
        MsgBox "This document contains the change list." & vbCr & vbCr & _
               "Switch to the document you want to correct, then run IPBES_ApplyChangeList again.", vbExclamation, "IPBES"
        Exit Sub
    End If

    Set candidates = New Collection
    names = ""
    For Each d In Documents
        If Not d Is target Then
            names = names & "  - " & d.Name & vbCr
            If HasChangeList(d) Then candidates.Add d
        End If
    Next d

    If candidates.Count = 0 Then
        If Application.ProtectedViewWindows.Count > 0 Then
            MsgBox "A document is open in Protected View, so the macro cannot read it:" & vbCr & "  - " & _
                   Application.ProtectedViewWindows(1).Document.Name & vbCr & vbCr & _
                   "Click 'Enable Editing' in the yellow bar of that document, then run IPBES_ApplyChangeList again.", _
                   vbExclamation, "IPBES"
        Else
            MsgBox "No change list found." & vbCr & vbCr & _
                   "Other open documents checked:" & vbCr & IIf(Len(names) = 0, "  (none)" & vbCr, names) & vbCr & _
                   "The change list must be a table (or lines like | 1 | ... | ... |) whose first row has the " & _
                   "headings Original and Revised." & vbCr & vbCr & _
                   "1. Copy the change table from Copilot or Claude." & vbCr & _
                   "2. Paste it into a new Word document and keep that document open." & vbCr & _
                   "3. Return to the document to correct and run IPBES_ApplyChangeList.", vbExclamation, "IPBES"
        End If
        Exit Sub
    End If

    If candidates.Count = 1 Then
        Set listDoc = candidates(1)
    Else
        names = ""
        For i = 1 To candidates.Count
            names = names & i & ". " & candidates(i).Name & vbCr
        Next i
        choice = InputBox("Several open documents contain a change list:" & vbCr & names & vbCr & _
                          "Type the number of the one to use:", "IPBES", "1")
        If Not IsNumeric(choice) Then Exit Sub
        i = CLng(choice)
        If i < 1 Or i > candidates.Count Then Exit Sub
        Set listDoc = candidates(i)
    End If

    Set items = ReadChangeList(listDoc)
    If items.Count = 0 Then
        MsgBox "The change list in '" & listDoc.Name & "' has no rows under its headings.", vbExclamation, "IPBES"
        Exit Sub
    End If

    If MsgBox("Apply " & items.Count & " change(s) from '" & listDoc.Name & "'" & vbCr & _
              "to '" & target.Name & "' with Track Changes?" & vbCr & vbCr & _
              "Each change is made only if its original text is found exactly once." & vbCr & _
              "Review every tracked change afterwards.", vbYesNo + vbQuestion, "IPBES") <> vbYes Then Exit Sub

    Set ReportLines = New Collection
    oldTrack = target.TrackRevisions
    If Not SetTracking(target, True) Then Exit Sub
    Application.ScreenUpdating = False

    For Each it In items
        iItem = iItem + 1
        Application.StatusBar = "IPBES change list: item " & iItem & " of " & items.Count & "..."
        If iItem Mod 5 = 0 Then DoEvents
        itemNo = it(0): orig = it(1): rev = it(2): ruleTxt = it(3)
        If Len(orig) = 0 Or orig = rev Then
            skipped = skipped + 1
            ReportLines.Add "Skipped" & vbTab & itemNo & vbTab & "" & vbTab & ruleTxt & vbTab & _
                "Original empty or identical to revised"
        ElseIf InStr(orig, vbCr) > 0 Or InStr(orig, Chr(11)) > 0 Then
            skipped = skipped + 1
            ReportLines.Add "Skipped" & vbTab & itemNo & vbTab & "" & vbTab & ruleTxt & vbTab & _
                "Original spans more than one paragraph: apply by hand"
        Else
            res = ApplyOneChange(target, orig, rev, ruleTxt)
            Select Case res
                Case 1
                    applied = applied + 1
                    ReportLines.Add "Applied" & vbTab & itemNo & vbTab & "" & vbTab & ruleTxt & vbTab & Left(orig, 50) & " -> " & Left(rev, 50)
                Case 0
                    notFound = notFound + 1
                    ReportLines.Add "NOT FOUND" & vbTab & itemNo & vbTab & "" & vbTab & ruleTxt & vbTab & Left(orig, 100)
                Case Else
                    ambiguous = ambiguous + 1
                    ReportLines.Add "Found " & res & " times" & vbTab & itemNo & vbTab & "" & vbTab & ruleTxt & vbTab & _
                        Left(orig, 100) & " (commented, not applied)"
            End Select
        End If
    Next it
    Application.StatusBar = False

    SetTracking target, oldTrack
    Application.ScreenUpdating = True
    MsgBox applied & " change(s) applied with Track Changes. Review them with Review > Next." & _
           ItemsNotApplied(), vbInformation, "IPBES"
End Sub

' Lists the change-list items that were not applied, by status, for the closing message.
Private Function ItemsNotApplied() As String
    Dim ln As Variant, parts() As String, nf As String, amb As String, sk As String
    For Each ln In ReportLines
        parts = Split(CStr(ln), vbTab)
        If parts(0) = "NOT FOUND" Then
            nf = nf & IIf(Len(nf) > 0, ", ", "") & parts(1)
        ElseIf Left(parts(0), 6) = "Found " Then
            amb = amb & IIf(Len(amb) > 0, ", ", "") & parts(1)
        ElseIf parts(0) = "Skipped" Then
            sk = sk & IIf(Len(sk) > 0, ", ", "") & parts(1)
        End If
    Next ln
    If Len(nf) > 0 Then ItemsNotApplied = ItemsNotApplied & vbCr & vbCr & "Not found (check the original text): items " & Left(nf, 300)
    If Len(amb) > 0 Then ItemsNotApplied = ItemsNotApplied & vbCr & vbCr & "Found more than once, marked with a comment, not applied: items " & Left(amb, 300)
    If Len(sk) > 0 Then ItemsNotApplied = ItemsNotApplied & vbCr & vbCr & "Skipped (empty, identical or over several paragraphs): items " & Left(sk, 300)
End Function

' True if the document holds a change list (a table or "| ... |" text lines with
' the headings Original and Revised).
Private Function HasChangeList(ByVal d As Document) As Boolean
    HasChangeList = (FindChangeTable(d) > 0) Or (FindPipeHeader(d) > 0)
End Function

' Returns the index of the first table whose header row contains "Original" and "Revised", or 0.
Private Function FindChangeTable(ByVal d As Document) As Long
    Dim i As Long, h As String
    For i = 1 To d.Tables.Count
        h = ""
        On Error Resume Next
        h = LCase(d.Tables(i).Range.Paragraphs(1).Range.Text)
        h = h & " " & LCase(d.Tables(i).Rows(1).Range.Text)
        On Error GoTo 0
        If InStr(h, "original") > 0 And (InStr(h, "revised") > 0 Or InStr(h, "corrected") > 0) Then
            FindChangeTable = i
            Exit Function
        End If
    Next i
End Function

' Returns the index of the paragraph holding a "| Original | Revised |" header line, or 0.
Private Function FindPipeHeader(ByVal d As Document) As Long
    Dim i As Long, t As String
    For i = 1 To d.Paragraphs.Count
        If i > 2000 Then Exit Function
        t = LCase(Trim(d.Paragraphs(i).Range.Text))
        If Left(t, 1) = "|" And InStr(t, "original") > 0 And (InStr(t, "revised") > 0 Or InStr(t, "corrected") > 0) Then
            FindPipeHeader = i
            Exit Function
        End If
    Next i
End Function

' Reads the change list into a collection of Array(no, original, revised, rule).
Private Function ReadChangeList(ByVal d As Document) As Collection
    Dim items As New Collection, tbl As Table, r As Long, c As Long, hdr As String
    Dim colOrig As Long, colRev As Long, colRule As Long, colNo As Long
    Dim cells() As String, i As Long, startPara As Long, t As String, n As Long
    Dim vNo As String, vOrig As String, vRev As String, vRule As String

    If FindChangeTable(d) > 0 Then
        Set tbl = d.Tables(FindChangeTable(d))
        For c = 1 To tbl.Rows(1).Cells.Count
            SetColumn LCase(CellText(tbl.Cell(1, c))), c, colOrig, colRev, colRule, colNo
        Next c
        For r = 2 To tbl.Rows.Count
            n = n + 1
            vNo = CStr(n): vOrig = "": vRev = "": vRule = ""
            On Error Resume Next                  ' rows with merged cells may lack a column
            If colNo > 0 Then vNo = CellText(tbl.Cell(r, colNo))
            vOrig = CellText(tbl.Cell(r, colOrig))
            vRev = CellText(tbl.Cell(r, colRev))
            If colRule > 0 Then vRule = CellText(tbl.Cell(r, colRule))
            On Error GoTo 0
            items.Add Array(vNo, vOrig, vRev, vRule)
        Next r
    Else
        startPara = FindPipeHeader(d)
        If startPara = 0 Then Set ReadChangeList = items: Exit Function
        cells = SplitPipes(d.Paragraphs(startPara).Range.Text)
        For c = 0 To UBound(cells)
            SetColumn LCase(cells(c)), c + 1, colOrig, colRev, colRule, colNo
        Next c
        For i = startPara + 1 To d.Paragraphs.Count
            t = Trim(Replace(d.Paragraphs(i).Range.Text, vbCr, ""))
            If Left(t, 1) <> "|" Then Exit For
            If Not (Replace(Replace(Replace(Replace(t, "|", ""), "-", ""), ":", ""), " ", "") = "") Then
                cells = SplitPipes(t)
                n = n + 1
                vNo = CStr(n)
                If colNo > 0 Then vNo = PipeCell(cells, colNo)
                items.Add Array(vNo, PipeCell(cells, colOrig), PipeCell(cells, colRev), PipeCell(cells, colRule))
            End If
        Next i
    End If
    Set ReadChangeList = items
End Function

Private Sub SetColumn(ByVal hdr As String, ByVal c As Long, ByRef colOrig As Long, ByRef colRev As Long, _
                      ByRef colRule As Long, ByRef colNo As Long)
    hdr = Trim(hdr)
    If InStr(hdr, "original") > 0 Then colOrig = c
    If InStr(hdr, "revised") > 0 Or InStr(hdr, "corrected") > 0 Then colRev = c
    If InStr(hdr, "rule") > 0 Or InStr(hdr, "reason") > 0 Then colRule = c
    If hdr = "no." Or hdr = "no" Or hdr = "#" Then colNo = c
End Sub

' Splits "| a | b | c |" into an array of cleaned cells (0-based).
Private Function SplitPipes(ByVal t As String) As String()
    Dim parts() As String, i As Long
    t = Trim(Replace(Replace(t, vbCr, ""), Chr(7), ""))
    If Left(t, 1) = "|" Then t = Mid(t, 2)
    If Right(t, 1) = "|" Then t = Left(t, Len(t) - 1)
    parts = Split(t, "|")
    For i = 0 To UBound(parts)
        parts(i) = CleanCellText(parts(i))
    Next i
    SplitPipes = parts
End Function

Private Function PipeCell(cells() As String, ByVal col As Long) As String
    If col >= 1 And col - 1 <= UBound(cells) Then PipeCell = cells(col - 1)
End Function

' Text of a table cell without the end-of-cell marker.
Private Function CellText(ByVal c As Cell) As String
    Dim t As String
    t = c.Range.Text
    If Len(t) >= 2 Then t = Left(t, Len(t) - 2)          ' strip Chr(13) & Chr(7)
    CellText = CleanCellText(t)
End Function

' Removes surrounding spaces, quotes, backticks and Markdown bold from a cell.
Private Function CleanCellText(ByVal t As String) As String
    Dim f As String, l As String
    t = Trim(Replace(t, ChrW(160), " "))
    If Left(t, 2) = "**" And Right(t, 2) = "**" And Len(t) >= 4 Then t = Mid(t, 3, Len(t) - 4)
    If Left(t, 1) = "`" And Right(t, 1) = "`" And Len(t) >= 2 Then t = Mid(t, 2, Len(t) - 2)
    If Len(t) >= 2 Then
        f = Left(t, 1): l = Right(t, 1)
        If (f = Chr(34) Or f = ChrW(8220)) And (l = Chr(34) Or l = ChrW(8221)) Then t = Mid(t, 2, Len(t) - 2)
    End If
    CleanCellText = Trim(t)
End Function

' Applies one change as a tracked change. Returns 1 if applied, 0 if not found,
' or the number of matches (2 or more) if the text is not unique.
Private Function ApplyOneChange(ByVal doc As Document, ByVal orig As String, ByVal rev As String, ByVal ruleTxt As String) As Long
    Dim rng As Range, full As Range, found As Range, n As Long, key As String, guard As Long
    key = Replace(Left(orig, 250), "^", "^^")
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = key
        .MatchCase = True
        .MatchWildcards = False
        .MatchWholeWord = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With
    Do While rng.Find.Execute
        guard = guard + 1
        If guard > 200 Then Exit Do
        Set full = doc.Range(rng.Start, rng.Start + Len(orig))
        If NormQ(full.Text) = NormQ(orig) Then
            n = n + 1
            If n = 1 Then Set found = full.Duplicate
        End If
        rng.Collapse wdCollapseEnd
    Loop
    If n = 0 Then
        ApplyOneChange = 0
    ElseIf n > 1 Then
        doc.Comments.Add found, TAG & "Change list: this text occurs " & n & " times, so it was not changed. " & _
            "Apply by hand where needed: '" & Left(orig, 80) & "' -> '" & Left(rev, 80) & "'."
        ApplyOneChange = n
    Else
        ApplyDiff doc, found, rev, ruleTxt
        ApplyOneChange = 1
    End If
End Function

' Replaces only the part that differs (snapped to word boundaries), so formatting
' around the change is kept and the tracked change is easy to read.
Private Sub ApplyDiff(ByVal doc As Document, ByVal full As Range, ByVal rev As String, ByVal ruleTxt As String)
    Dim a As String, b As String, p As Long, s As Long, lo As Long, lr As Long, chg As Range, ins As String
    a = NormQ(full.Text)
    b = NormQ(rev)
    lo = Len(a): lr = Len(b)
    ' common prefix, snapped back to the start of a word
    Do While p < lo And p < lr
        If Mid(a, p + 1, 1) <> Mid(b, p + 1, 1) Then Exit Do
        p = p + 1
    Loop
    Do While p > 0
        If Mid(a, p, 1) = " " Then Exit Do
        p = p - 1
    Loop
    ' common suffix (not overlapping the prefix), snapped to a word boundary
    Do While s < lo - p And s < lr - p
        If Mid(a, lo - s, 1) <> Mid(b, lr - s, 1) Then Exit Do
        s = s + 1
    Loop
    Do While s > 0
        If lo - s = 0 Then Exit Do                ' suffix starts at the beginning of the text
        If Mid(a, lo - s + 1, 1) = " " Then Exit Do
        If Mid(a, lo - s, 1) = " " Then Exit Do
        s = s - 1
    Loop
    Set chg = doc.Range(full.Start + p, full.End - s)
    ins = Mid(rev, p + 1, lr - p - s)
    If chg.Start = chg.End Then
        chg.InsertAfter ins
    ElseIf Len(ins) = 0 Then
        chg.Delete
    Else
        chg.Text = ins
    End If
    If Len(ruleTxt) > 0 Then
        On Error Resume Next
        doc.Comments.Add chg, TAG & ruleTxt
        On Error GoTo 0
    End If
End Sub

' Normalizes quotes, apostrophes, non-breaking spaces and hyphens for comparison
' (one character for one character, so lengths are unchanged).
Private Function NormQ(ByVal t As String) As String
    t = Replace(t, ChrW(8216), "'")
    t = Replace(t, ChrW(8217), "'")
    t = Replace(t, ChrW(8220), Chr(34))
    t = Replace(t, ChrW(8221), Chr(34))
    t = Replace(t, ChrW(160), " ")
    t = Replace(t, Chr(30), "-")
    t = Replace(t, ChrW(8209), "-")
    NormQ = t
End Function

' ---------------------------------------------------------------------------
' Ribbon entry points (IPBES tab, see word-macro/ipbes-ribbon.xml).
' The control argument is declared As Object so that no extra library
' reference is needed.
' ---------------------------------------------------------------------------
Public Sub IPBESRibbon_CheckDocument(ByVal control As Object)
    IPBES_CheckDocument
End Sub

Public Sub IPBESRibbon_CheckVoice(ByVal control As Object)
    IPBES_CheckVoice
End Sub

Public Sub IPBESRibbon_SafeFixes(ByVal control As Object)
    IPBES_SafeFixes
End Sub

Public Sub IPBESRibbon_ApplyChangeList(ByVal control As Object)
    IPBES_ApplyChangeList
End Sub

Public Sub IPBESRibbon_FixFormatting(ByVal control As Object)
    IPBES_FixFormatting
End Sub

Public Sub IPBESRibbon_RemoveComments(ByVal control As Object)
    IPBES_RemoveIPBESComments
End Sub

' ---------------------------------------------------------------------------
' Formatting: IPBES styles, page set-up, headings, body paragraphs, tables
' ---------------------------------------------------------------------------
Public Sub IPBES_FixFormatting()
    Dim doc As Document, ans As Long, numbered As Boolean, alreadyIPBES As Boolean
    Dim oldTrack As Boolean, oldTrackFmt As Boolean, hadTrackFmt As Boolean
    Dim oldPagination As Boolean, oldView As Long, ur As Object
    Dim p As Paragraph, sName As String, target As String, txt As String, startBody As Long
    Dim t As Table, msg As String, tbls As Collection, ti As Long
    Dim nPage As Long, nHeadings As Long, nLabels As Long, nBody As Long, nNumbers As Long
    Dim nTitle As Long, nCaptions As Long, nTables As Long, nEmpty As Long, nLists As Long
    Dim stylesCopied As Boolean, typedMode As Boolean, typed As Long, nNumbered As Long, nRestarts As Long
    Dim otherStyles As String, nHeaders As Long, nParas As Long, iPara As Long, iTable As Long, nCleaned As Long

    Set doc = ActiveDocument
    If Not AddinHasStyles() Then
        MsgBox "This copy of the IPBES add-in does not contain the IPBES Word styles." & vbCr & vbCr & _
               "The secretariat editor can add them with tools/prepare_addin.py (see the README).", _
               vbExclamation, "IPBES"
        Exit Sub
    End If
    If doc.FullName = ThisDocument.FullName Then
        MsgBox "Run this macro on the document to format, not on the add-in itself.", vbExclamation, "IPBES"
        Exit Sub
    End If

    ans = MsgBox("IPBES formatting" & vbCr & vbCr & _
        "Applies the IPBES Word styles, page set-up (A4, IPBES margins), heading layout and table layout. " & _
        "Changes are tracked. Page set-up and the style definitions themselves are not, so work on a copy if in doubt." & vbCr & vbCr & _
        "Is this a report with numbered paragraphs (1., 2., 3. ...)?" & vbCr & vbCr & _
        "Yes = numbered paragraphs" & vbCr & "No = unnumbered paragraphs" & vbCr & "Cancel = stop", _
        vbYesNoCancel + vbQuestion, "IPBES")
    If ans = vbCancel Then Exit Sub
    numbered = (ans = vbYes)

    oldTrack = doc.TrackRevisions
    On Error Resume Next
    oldTrackFmt = doc.TrackFormatting
    hadTrackFmt = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If Not SetTracking(doc, True) Then Exit Sub
    On Error Resume Next
    doc.TrackFormatting = True
    ' Speed: no screen redraw, no background repagination, draft view; one Undo step for the run
    oldPagination = Options.Pagination
    Options.Pagination = False
    oldView = doc.ActiveWindow.View.Type
    doc.ActiveWindow.View.Type = wdNormalView
    Set ur = Application.UndoRecord
    ur.StartCustomRecord "IPBES formatting"
    Err.Clear
    On Error GoTo 0
    Application.ScreenUpdating = False
    Application.StatusBar = "IPBES formatting: preparing..."

    CacheBuiltinNames doc
    alreadyIPBES = UsesStyle(doc, "Normal_number") Or UsesStyle(doc, "CH1")

    ' 1. Style definitions from the add-in, then the key IPBES formats set explicitly
    '    (so the result does not depend on how Word merged the copied styles)
    '    A document that already uses the IPBES styles keeps its own definitions.
    If alreadyIPBES Then
        stylesCopied = True
    Else
        Application.StatusBar = "IPBES formatting: styles..."
        On Error Resume Next
        doc.CopyStylesFromTemplate ThisDocument.FullName
        stylesCopied = (Err.Number = 0)
        Err.Clear
        On Error GoTo 0
        EnsureIPBESStyles doc
    End If

    ' 2. Page set-up
    nPage = SetIPBESPage(doc)

    ' 3. Paragraphs (outside tables, after any masthead)
    startBody = BodyStart(doc)
    typedMode = numbered And (CountTypedNumbers(doc, startBody, numbered, alreadyIPBES) >= 3)
    Set tbls = TableRanges(doc)
    ti = 1
    nParas = doc.Paragraphs.Count
    For Each p In doc.Paragraphs
        iPara = iPara + 1
        If iPara Mod 200 = 0 Then
            Application.StatusBar = "IPBES formatting: paragraph " & iPara & " of " & nParas & "..."
            DoEvents
        End If
        If p.Range.Start >= startBody Then
            If Not InTableFast(p.Range.Start, tbls, ti) Then
                sName = StyleName(p)
                txt = Trim(Replace(Replace(Replace(p.Range.Text, vbCr, ""), Chr(11), ""), vbTab, " "))
                target = TargetStyle(doc, sName, txt, numbered, alreadyIPBES)
                If target = "HCH" Then
                    If HasHeadingLabel(p) Then
                        target = "CH1"
                    ElseIf nTitle = 0 And nHeadings = 0 Then
                        target = "BB_Title"
                    Else
                        target = "CH1"
                    End If
                End If
                Select Case target
                    Case "CH1", "CH2", "CH3", "CH4"
                        p.Style = doc.Styles(target)
                        ResetDirect p
                        If FixHeadingLabel(doc, p) Then nLabels = nLabels + 1
                        nHeadings = nHeadings + 1
                    Case "BB_Title"
                        p.Style = doc.Styles(target)
                        ResetDirect p
                        nTitle = nTitle + 1
                    Case "Title_table", "Title_figure"
                        p.Style = doc.Styles(target)
                        ResetDirect p
                        NormalizeCaption doc, p, IIf(target = "Title_table", "Table", "Figure")
                        If p.Range.Font.Bold <> True Then p.Range.Font.Bold = True
                        nCaptions = nCaptions + 1
                    Case "Normal_number", "Normal_No_number"
                        If Len(txt) = 0 Then
                            nEmpty = nEmpty + 1
                        Else
                            typed = 0
                            If numbered Then typed = RemoveTypedNumber(doc, p)
                            If typed > 0 Then nNumbers = nNumbers + 1
                            If typedMode And typed = 0 Then target = "Normal_No_number"
                            p.Style = doc.Styles(target)
                            ResetDirect p
                            If typedMode And typed = 1 And nNumbered > 0 Then
                                RestartNumbering p
                                nRestarts = nRestarts + 1
                            End If
                            If target = "Normal_number" Then nNumbered = nNumbered + 1
                            nBody = nBody + 1
                        End If
                    Case "LIST"
                        nLists = nLists + 1
                    Case ""
                        If IsIPBESStyle(sName) Then
                            ' already in an IPBES style: remove whole-paragraph manual font, size, colour
                            If CleanUniformFormatting(p) Then nCleaned = nCleaned + 1
                        ElseIf Len(txt) > 0 And sName <> bnNormal Then
                            NoteStyle otherStyles, sName
                        End If
                End Select
            End If
        End If
    Next p

    ' 4. Tables (not the masthead) and the caption line above them
    iTable = 0
    For Each t In doc.Tables
        iTable = iTable + 1
        Application.StatusBar = "IPBES formatting: table " & iTable & " of " & doc.Tables.Count & "..."
        DoEvents
        If t.Range.Start >= startBody And t.NestingLevel = 1 Then
            If FixTableCaption(doc, t) Then nCaptions = nCaptions + 1
            If alreadyIPBES Then
                FitTableToText t
            Else
                FormatTable doc, t
            End If
            nTables = nTables + 1
        End If
    Next t

    ' 5. Page headers and footers (symbol and page number, IPBES layout)
    Application.StatusBar = "IPBES formatting: headers and footers..."
    If Not alreadyIPBES Then nHeaders = FixHeadersFooters(doc)

    ' Restore settings
    On Error Resume Next
    If Not ur Is Nothing Then ur.EndCustomRecord
    If hadTrackFmt Then doc.TrackFormatting = oldTrackFmt
    doc.ActiveWindow.View.Type = oldView
    Options.Pagination = oldPagination
    Application.StatusBar = False
    On Error GoTo 0
    SetTracking doc, oldTrack
    Application.ScreenUpdating = True

    ' Short summary; the tracked changes are the detailed record
    msg = "IPBES formatting done. Review the tracked changes (Review > Next); Ctrl+Z undoes the whole run." & vbCr & vbCr
    If Not stylesCopied Then msg = msg & "! Could not copy the IPBES styles (is the document protected?)" & vbCr
    msg = msg & "Headings: " & nHeadings & IIf(nTitle > 0, " (and the title)", "") & vbCr
    msg = msg & "Body paragraphs: " & nBody & IIf(numbered, " (numbered)", " (unnumbered)") & vbCr
    If nNumbers > 0 Then msg = msg & "Typed paragraph numbers removed: " & nNumbers & vbCr
    If nCleaned > 0 Then msg = msg & "Manual font, size or colour removed from IPBES-styled paragraphs: " & nCleaned & vbCr
    If nRestarts > 0 Then msg = msg & "Numbering restarted at 1 (check these, e.g. decisions): " & nRestarts & vbCr
    If nCaptions + nTables > 0 Then msg = msg & "Tables: " & nTables & ", captions: " & nCaptions & vbCr
    If nPage > 0 Then msg = msg & "Sections set to A4 with IPBES margins: " & nPage & vbCr
    If nHeaders > 0 Then
        msg = msg & "Page headers and footers rebuilt (symbol, page numbers): " & nHeaders & vbCr
    ElseIf Not alreadyIPBES Then
        msg = msg & "Page headers left as they were (no IPBES/.. symbol found)" & vbCr
    End If
    If nEmpty + nLists > 0 Or Len(otherStyles) > 0 Then msg = msg & vbCr & "Left for you:" & vbCr
    If nEmpty > 0 Then msg = msg & "- " & nEmpty & " empty paragraph(s)" & vbCr
    If nLists > 0 Then msg = msg & "- " & nLists & " list paragraph(s): IPBES uses (a), (b), (i)" & vbCr
    If Len(otherStyles) > 0 Then msg = msg & "- unchanged styles: " & Left(Replace(otherStyles, "|", "; "), 200) & vbCr
    If alreadyIPBES Then msg = msg & vbCr & "The document already used the IPBES styles: its style definitions, headers, " & _
        "footers, table layout and Normal paragraphs were kept; only other paragraphs, captions and over-wide tables were changed." & vbCr
    msg = msg & vbCr & "Next: run Check document for heading numbering, lone subheadings and the masthead."
    MsgBox msg, vbInformation, "IPBES"
End Sub

' Sets the key IPBES paragraph styles explicitly (values from the IPBES report template,
' see ipbes-editorial/references/report-structure.md). Creates a style if it is missing.
Private Sub EnsureIPBESStyles(ByVal doc As Document)
    On Error Resume Next
    DefineStyle doc, "Normal", 10, False, 0, 0, 0, 0, False
    DefineStyle doc, "Normal-pool", 10, False, 0, 0, 0, 0, False
    DefineStyle doc, "BB_Title", 14, True, 2.2, 0, 16, 12, True
    DefineStyle doc, "CH1", 14, True, 2.2, -2.2, 12, 6, True
    DefineStyle doc, "CH2", 12, True, 2.2, -2.2, 12, 6, True
    DefineStyle doc, "CH3", 10, True, 2.2, -2.2, 12, 6, True
    DefineStyle doc, "CH4", 10, True, 2.2, -2.2, 6, 6, True
    DefineStyle doc, "Normal_number", 10, False, 2.2, 0, 0, 6, False
    DefineStyle doc, "Normal_No_number", 10, False, 2.2, 0, 0, 6, False
    DefineStyle doc, "Title_table", 10, True, 2.2, 0, 0, 3, True
    DefineStyle doc, "Title_figure", 10, True, 2.2, 0, 0, 3, True
    DefineStyle doc, "Header-pool", 9, True, 0, 0, 0, 0, False
    DefineStyle doc, "Footer-pool", 9, True, 0, 0, 0, 6, False
    ' heading layout "<tab>I.<tab>Title": number at 1.1 cm, title at the 2.2 cm hanging indent
    Dim h As Variant
    For Each h In Array("CH1", "CH2", "CH3", "CH4")
        With doc.Styles(h).ParagraphFormat.TabStops
            .ClearAll
            .Add CentimetersToPoints(1.1)
        End With
    Next h
    ' paragraph numbers: wrapped lines at 2.2 cm, text after the number at 3.3 cm
    With doc.Styles("Normal_number").ParagraphFormat
        .TabStops.ClearAll
        .TabStops.Add CentimetersToPoints(3.3)
    End With
    doc.Styles("Header-pool").ParagraphFormat.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
    doc.Styles("Header-pool").ParagraphFormat.Borders(wdBorderBottom).LineWidth = wdLineWidth050pt
    EnsureParagraphNumbering doc
    On Error GoTo 0
End Sub

Private Sub DefineStyle(ByVal doc As Document, ByVal sName As String, ByVal size As Single, ByVal isBold As Boolean, _
                        ByVal leftCm As Single, ByVal firstCm As Single, ByVal beforePt As Single, _
                        ByVal afterPt As Single, ByVal keepNext As Boolean)
    Dim st As Style
    On Error Resume Next
    Set st = doc.Styles(sName)
    If st Is Nothing Then Set st = doc.Styles.Add(sName, wdStyleTypeParagraph)
    If st Is Nothing Then Exit Sub
    With st.Font
        .Name = "Times New Roman"
        .Size = size
        .Bold = isBold
        .Italic = False
        .Color = wdColorAutomatic
    End With
    With st.ParagraphFormat
        .Alignment = wdAlignParagraphLeft
        .LeftIndent = CentimetersToPoints(leftCm)
        .FirstLineIndent = CentimetersToPoints(firstCm)
        .SpaceBefore = beforePt
        .SpaceAfter = afterPt
        .LineSpacingRule = wdLineSpaceSingle
        .KeepWithNext = keepNext
    End With
    On Error GoTo 0
End Sub

' Paragraph numbering of Normal_number: "1." at 2.2 cm with the text at 3.3 cm and wrapped
' lines at 2.2 cm; "(a)" at 3.3 cm; "(i)" at 4.4 cm (IPBES report layout).
Private Sub EnsureParagraphNumbering(ByVal doc As Document)
    Dim st As Style, lt As ListTemplate
    On Error Resume Next
    Set st = doc.Styles("Normal_number")
    If st Is Nothing Then Exit Sub
    Set lt = st.ListTemplate
    If lt Is Nothing Then
        Set lt = doc.ListTemplates.Add(OutlineNumbered:=True)
        st.LinkToListTemplate lt, 1
    End If
    With lt.ListLevels(1)
        .NumberStyle = wdListNumberStyleArabic
        .NumberFormat = "%1."
        .Alignment = wdListLevelAlignLeft
        .NumberPosition = CentimetersToPoints(2.2)
        .TextPosition = CentimetersToPoints(2.2)
        .TabPosition = CentimetersToPoints(3.3)
        .TrailingCharacter = wdTrailingTab
    End With
    With lt.ListLevels(2)
        .NumberStyle = wdListNumberStyleLowercaseLetter
        .NumberFormat = "(%2)"
        .Alignment = wdListLevelAlignLeft
        .NumberPosition = CentimetersToPoints(3.3)
        .TextPosition = CentimetersToPoints(2.2)
        .TabPosition = CentimetersToPoints(4.4)
        .TrailingCharacter = wdTrailingTab
    End With
    With lt.ListLevels(3)
        .NumberStyle = wdListNumberStyleLowercaseRoman
        .NumberFormat = "(%3)"
        .Alignment = wdListLevelAlignLeft
        .NumberPosition = CentimetersToPoints(4.4)
        .TextPosition = CentimetersToPoints(5.5)
        .TabPosition = CentimetersToPoints(5.5)
        .TrailingCharacter = wdTrailingTab
    End With
    On Error GoTo 0
End Sub

' True if the add-in itself holds the IPBES styles (added by tools/prepare_addin.py).
Private Function AddinHasStyles() As Boolean
    Dim st As Style
    On Error Resume Next
    Set st = ThisDocument.Styles("CH1")
    If Not st Is Nothing Then
        Set st = Nothing
        Set st = ThisDocument.Styles("Normal_number")
    End If
    AddinHasStyles = Not st Is Nothing
    On Error GoTo 0
End Function

' True if at least one paragraph uses the given style.
Private Function UsesStyle(ByVal doc As Document, ByVal sty As String) As Boolean
    Dim rng As Range, st As Style
    On Error Resume Next
    Set st = doc.Styles(sty)
    On Error GoTo 0
    If st Is Nothing Then Exit Function
    Set rng = doc.Content
    With rng.Find
        .ClearFormatting
        .Text = ""
        .Style = st
        .Format = True
        .Forward = True
        .Wrap = wdFindStop
        .MatchWildcards = False
    End With
    UsesStyle = rng.Find.Execute
End Function

Private Sub CacheBuiltinNames(ByVal doc As Document)
    bnHeading1 = BuiltinName(doc, wdStyleHeading1): bnHeading2 = BuiltinName(doc, wdStyleHeading2)
    bnHeading3 = BuiltinName(doc, wdStyleHeading3): bnHeading4 = BuiltinName(doc, wdStyleHeading4)
    bnTitle = BuiltinName(doc, wdStyleTitle): bnCaption = BuiltinName(doc, wdStyleCaption)
    bnListPara = BuiltinName(doc, wdStyleListParagraph): bnListBullet = BuiltinName(doc, wdStyleListBullet)
    bnListNumber = BuiltinName(doc, wdStyleListNumber): bnNormal = BuiltinName(doc, wdStyleNormal)
    bnBodyText = BuiltinName(doc, wdStyleBodyText): bnBodyText2 = BuiltinName(doc, wdStyleBodyText2)
    bnBodyText3 = BuiltinName(doc, wdStyleBodyText3): bnPlainText = BuiltinName(doc, wdStylePlainText)
    bnNormalIndent = BuiltinName(doc, wdStyleNormalIndent)
End Sub

' Ranges of the top-level tables, in document order. Word keeps Range objects up to date
' as text is inserted or deleted, so they stay valid while paragraphs are edited.
Private Function TableRanges(ByVal doc As Document) As Collection
    Dim c As New Collection, t As Table
    For Each t In doc.Tables
        c.Add t.Range
    Next t
    Set TableRanges = c
End Function

' Fast replacement for Range.Information(wdWithInTable) when positions are visited in order.
Private Function InTableFast(ByVal pos As Long, ByVal tbls As Collection, ByRef ti As Long) As Boolean
    Do While ti <= tbls.Count
        If tbls(ti).End > pos Then Exit Do
        ti = ti + 1
    Loop
    If ti <= tbls.Count Then InTableFast = (tbls(ti).Start <= pos)
End Function

' Local name of a built-in style (works in any Word language).
Private Function BuiltinName(ByVal doc As Document, ByVal id As Long) As String
    On Error Resume Next
    BuiltinName = doc.Styles(id).NameLocal
    On Error GoTo 0
End Function

' Decides the IPBES style for a paragraph, or "" to leave it alone.
Private Function TargetStyle(ByVal doc As Document, ByVal sName As String, ByVal txt As String, _
                             ByVal numbered As Boolean, ByVal alreadyIPBES As Boolean) As String
    Dim body As String
    body = IIf(numbered, "Normal_number", "Normal_No_number")
    If IsIPBESStyle(sName) Then Exit Function
    ' Caption lines are never body text
    If IsCaption(txt, "Table") Then TargetStyle = "Title_table": Exit Function
    If IsCaption(txt, "Figure") Then TargetStyle = "Title_figure": Exit Function
    ' Styles of the older UN document template (IPBES-1 and other DGACM documents)
    If sName Like "_ H _Ch*" Or sName Like "_ H_Ch*" Or sName Like "_ H __M*" Then TargetStyle = "HCH": Exit Function
    If sName Like "_ H_1*" Then TargetStyle = "CH2": Exit Function
    If sName Like "_ H_2*" Or sName Like "_ H_3*" Then TargetStyle = "CH3": Exit Function
    If sName Like "_ H_4*" Or sName Like "_ H_5*" Then TargetStyle = "CH4": Exit Function
    If sName Like "__Single Txt*" Or sName Like "_ Single Txt*" Or sName Like "SingleTxt*" Then TargetStyle = body: Exit Function
    If Len(bnNormal) = 0 Then CacheBuiltinNames doc
    Select Case sName
        Case bnHeading1: TargetStyle = "CH1"
        Case bnHeading2: TargetStyle = "CH2"
        Case bnHeading3: TargetStyle = "CH3"
        Case bnHeading4: TargetStyle = "CH4"
        Case bnTitle: TargetStyle = "BB_Title"
        Case bnCaption
            If LCase(Left(txt, 6)) = "figure" Then TargetStyle = "Title_figure" Else TargetStyle = "Title_table"
        Case bnListPara, bnListBullet, bnListNumber
            TargetStyle = "LIST"
        Case bnNormal
            If Not alreadyIPBES Then TargetStyle = body
        Case bnBodyText, bnBodyText2, bnBodyText3, bnPlainText, bnNormalIndent, "No Spacing", "Normal (Web)", "Default"
            TargetStyle = body
    End Select
End Function

' Start of the document body: after the masthead (tables, session block and any stray
' lines), i.e. at the first title or heading after the first table.
Private Function BodyStart(ByVal doc As Document) As Long
    Dim i As Long, txt As String, p As Paragraph, sName As String, firstTableEnd As Long
    If doc.Tables.Count = 0 Then Exit Function
    If Len(bnNormal) = 0 Then CacheBuiltinNames doc
    txt = UCase(Norm(doc.Tables(1).Range.Text))
    If doc.Tables(1).Range.Start > 200 Or InStr(txt, "UNITED") = 0 Then Exit Function
    firstTableEnd = doc.Tables(1).Range.End
    For Each p In doc.Paragraphs
        If p.Range.Start > 8000 Then Exit For
        If p.Range.Start >= firstTableEnd Then
            If Not p.Range.Information(wdWithInTable) Then
                sName = StyleName(p)
                If sName = "BB_Title" Or sName = "CH1" Or sName Like "_ H _Ch*" Or sName Like "_ H_Ch*" Or sName Like "_ H __M*" _
                   Or sName = bnTitle Or sName = bnHeading1 _
                   Or Left(Trim(p.Range.Text), 10) = "Report of " Then
                    BodyStart = p.Range.Start
                    Exit Function
                End If
            End If
        End If
    Next p
    For i = 1 To MinL(4, doc.Tables.Count)
        If doc.Tables(i).Range.Start < 3000 Then BodyStart = doc.Tables(i).Range.End
    Next i
End Function

' A4 and IPBES margins for every section. Returns the number of sections changed.
Private Function SetIPBESPage(ByVal doc As Document) As Long
    Dim sec As Section, n As Long, w As Single, h As Single
    On Error Resume Next
    For Each sec In doc.Sections
        With sec.PageSetup
            If .Orientation = wdOrientLandscape Then
                w = CentimetersToPoints(29.7): h = CentimetersToPoints(21)
            Else
                w = CentimetersToPoints(21): h = CentimetersToPoints(29.7)
            End If
            If Abs(.PageWidth - w) > 1 Or Abs(.PageHeight - h) > 1 Or Abs(.TopMargin - CentimetersToPoints(1.6)) > 1 Or _
               Abs(.BottomMargin - CentimetersToPoints(2.5)) > 1 Or Abs(.LeftMargin - CentimetersToPoints(2.5)) > 1 Or _
               Abs(.RightMargin - CentimetersToPoints(1.75)) > 1 Then
                ' only when something differs: every assignment makes Word lay out the document again
                n = n + 1
                .PageWidth = w
                .PageHeight = h
                .TopMargin = CentimetersToPoints(1.6)
                .BottomMargin = CentimetersToPoints(2.5)
                .LeftMargin = CentimetersToPoints(2.5)
                .RightMargin = CentimetersToPoints(1.75)
                .HeaderDistance = CentimetersToPoints(0.95)
                .FooterDistance = CentimetersToPoints(1.72)
            End If
        End With
    Next sec
    On Error GoTo 0
    SetIPBESPage = n
End Function

' Removes direct formatting so that the paragraph looks like its style.
' Bold, italic and underline are kept where they mark something in the text.
Private Sub ResetDirect(ByVal p As Paragraph)
    Dim r As Range
    Set r = p.Range
    On Error Resume Next
    r.ParagraphFormat.Reset
    If r.Font.Bold = 0 And r.Font.Italic = 0 And r.Font.Underline = 0 Then
        r.Font.Reset
    Else
        r.Font.Name = p.Style.Font.Name
        r.Font.Size = p.Style.Font.Size
        r.Font.Color = wdColorAutomatic
        r.Font.Shading.BackgroundPatternColor = wdColorAutomatic
        r.HighlightColorIndex = wdNoHighlight
    End If
    On Error GoTo 0
End Sub

' For a paragraph already in an IPBES style: removes manual font name, size, colour and
' highlighting applied to the WHOLE paragraph (e.g. a title made red Arial 14 pt). Bold,
' italic and underline are kept, and paragraphs with mixed formatting are not touched
' (Word reports mixed values as wdUndefined), so word-level formatting survives.
Private Function CleanUniformFormatting(ByVal p As Paragraph) As Boolean
    Dim r As Range, sf As Font, changed As Boolean
    On Error Resume Next
    Set r = p.Range
    Set sf = p.Style.Font
    If sf Is Nothing Then Exit Function
    If Len(r.Font.Name) > 0 And r.Font.Name <> sf.Name Then
        r.Font.Name = sf.Name
        changed = True
    End If
    If r.Font.Size <> wdUndefined And r.Font.Size <> sf.Size Then
        r.Font.Size = sf.Size
        changed = True
    End If
    If r.Font.Color <> wdUndefined And r.Font.Color <> sf.Color Then
        r.Font.Color = sf.Color
        changed = True
    End If
    If r.HighlightColorIndex <> wdUndefined And r.HighlightColorIndex <> wdNoHighlight Then
        r.HighlightColorIndex = wdNoHighlight
        changed = True
    End If
    On Error GoTo 0
    CleanUniformFormatting = changed
End Function

' "I. Introduction" -> tab, "I.", tab, "Introduction" (IPBES heading layout).
Private Function FixHeadingLabel(ByVal doc As Document, ByVal p As Paragraph) As Boolean
    Dim t As String, sp As Long, label As String, core As String, rng As Range
    t = Replace(p.Range.Text, vbCr, "")
    If InStr(t, vbTab) > 0 Then Exit Function
    sp = InStr(t, " ")
    If sp < 3 Or sp > 8 Then Exit Function
    label = Left(t, sp - 1)
    If Right(label, 1) <> "." Then Exit Function
    core = Left(label, Len(label) - 1)
    If Not (RomanValue(core) > 0 Or core Like "[A-Z]" Or core Like "#" Or core Like "##" Or core Like "([a-z])") Then Exit Function
    Set rng = doc.Range(p.Range.Start, p.Range.Start + sp)
    rng.Text = vbTab & label & vbTab
    FixHeadingLabel = True
End Function

' Removes a typed paragraph number ("12. " or "12.<tab>") when the style numbers
' paragraphs. Returns the number that was typed, or 0 if there was none.
Private Function RemoveTypedNumber(ByVal doc As Document, ByVal p As Paragraph) As Long
    Dim t As String, i As Long, v As Long
    t = p.Range.Text
    i = TypedNumberLength(t, v)
    If i = 0 Then Exit Function
    doc.Range(p.Range.Start, p.Range.Start + i).Delete
    RemoveTypedNumber = v
End Function

' Length of a leading "12. " / "12.<tab>" (with following spaces), and its value; 0 if none.
Private Function TypedNumberLength(ByVal t As String, ByRef v As Long) As Long
    Dim i As Long
    i = 1
    Do While i <= Len(t) And i <= 4
        If Not (Mid(t, i, 1) Like "#") Then Exit Do
        i = i + 1
    Loop
    If i = 1 Or i > Len(t) - 1 Then Exit Function
    If Mid(t, i, 1) <> "." Then Exit Function
    If Mid(t, i + 1, 1) <> " " And Mid(t, i + 1, 1) <> vbTab Then Exit Function
    v = CLng(Left(t, i - 1))
    i = i + 1
    Do While i <= Len(t) And (Mid(t, i, 1) = " " Or Mid(t, i, 1) = vbTab)
        i = i + 1
    Loop
    TypedNumberLength = i - 1
End Function

' Number of body paragraphs that start with a typed paragraph number.
Private Function CountTypedNumbers(ByVal doc As Document, ByVal startBody As Long, ByVal numbered As Boolean, _
                                   ByVal alreadyIPBES As Boolean) As Long
    Dim p As Paragraph, n As Long, v As Long, tg As String, txt As String, tbls As Collection, ti As Long
    Set tbls = TableRanges(doc)
    ti = 1
    For Each p In doc.Paragraphs
        If p.Range.Start >= startBody Then
            If Not InTableFast(p.Range.Start, tbls, ti) Then
                txt = Trim(Replace(Replace(p.Range.Text, vbCr, ""), Chr(11), ""))
                tg = TargetStyle(doc, StyleName(p), txt, numbered, alreadyIPBES)
                If tg = "Normal_number" Or tg = "Normal_No_number" Then
                    If TypedNumberLength(p.Range.Text, v) > 0 Then n = n + 1
                End If
            End If
        End If
    Next p
    CountTypedNumbers = n
End Function

' Restarts the paragraph numbering at 1 from this paragraph onwards.
Private Sub RestartNumbering(ByVal p As Paragraph)
    Dim lt As ListTemplate
    On Error Resume Next
    Set lt = p.Range.ListFormat.ListTemplate
    If lt Is Nothing Then Exit Sub
    p.Range.ListFormat.ApplyListTemplate ListTemplate:=lt, ContinuePreviousList:=False, _
        ApplyTo:=wdListApplyToThisPointForward
    On Error GoTo 0
End Sub

' True if the heading text starts with a number ("I.", "A.", "1.") or already has the tab layout.
Private Function HasHeadingLabel(ByVal p As Paragraph) As Boolean
    Dim t As String, sp As Long, core As String
    t = Replace(p.Range.Text, vbCr, "")
    If Left(t, 1) = vbTab Then t = Mid(t, 2)
    sp = InStr(Replace(t, vbTab, " "), " ")
    If sp < 3 Or sp > 8 Then Exit Function
    core = Left(t, sp - 1)
    If Right(core, 1) <> "." Then Exit Function
    core = Left(core, Len(core) - 1)
    HasHeadingLabel = (RomanValue(core) > 0 Or core Like "[A-Z]" Or core Like "#" Or core Like "##")
End Function

' True if the text is a caption such as "Table 1", "Table 1: Title", "Table IV. Title",
' "Table 1 Title" or "Table A.1", and not a sentence such as "Table 1 presents ...".
Private Function IsCaption(ByVal txt As String, ByVal capWord As String) As Boolean
    Dim rest As String, i As Long, c As String
    txt = Trim(Replace(Replace(txt, Chr(11), " "), vbTab, " "))
    If Len(txt) > 200 Then Exit Function
    If Left(txt, Len(capWord) + 1) <> capWord & " " Then Exit Function
    rest = Mid(txt, Len(capWord) + 2)
    ' the label: digits, Roman numerals, or A.1
    i = 1
    Do While i <= Len(rest)
        c = Mid(rest, i, 1)
        If Not (c Like "[0-9IVXA.]") Then Exit Do
        i = i + 1
    Loop
    If i = 1 Then Exit Function
    rest = Trim(Mid(rest, i))
    If Len(rest) = 0 Then IsCaption = True: Exit Function
    c = Left(rest, 1)
    If c = ":" Or c = "." Or c = "-" Or c = ChrW(8211) Or c = ChrW(8212) Then IsCaption = True: Exit Function
    IsCaption = (c Like "[A-Z(]")
End Function

' True for the IPBES styles themselves (never remapped).
Private Function IsIPBESStyle(ByVal sName As String) As Boolean
    Select Case sName
        Case "CH1", "CH2", "CH3", "CH4", "BB_Title", "AA_Title", "Normal_number", "Normal_No_number", "Normal-pool", _
             "Title_table", "Title_figure", "ZZ_Anx_header", "ZZ_Anx_title", "A_Spacer", "Header-pool", "Footer-pool", _
             "Footer-jobnumber", "Normal_Plain"
            IsIPBESStyle = True
    End Select
End Function

' Collects up to eight names of styles that were left unchanged.
Private Sub NoteStyle(ByRef list As String, ByVal sName As String)
    If InStr("|" & list & "|", "|" & sName & "|") > 0 Then Exit Sub
    If UBound(Split(list, "|")) >= 7 Then Exit Sub
    If Len(list) > 0 Then list = list & "|"
    list = list & sName
End Sub

' A paragraph "Table 1: Title" just above a table becomes "Table 1" + line break + "Title" in Title_table.
Private Function FixTableCaption(ByVal doc As Document, ByVal t As Table) As Boolean
    Dim pp As Paragraph, txt As String, k As Long
    If t.Range.Start = 0 Then Exit Function
    Set pp = doc.Range(0, t.Range.Start).Paragraphs.Last
    ' skip up to two empty paragraphs between the caption and the table
    For k = 1 To 2
        If pp Is Nothing Then Exit Function
        If pp.Range.Information(wdWithInTable) Then Exit Function
        If Len(CleanCaptionText(pp.Range.Text)) > 0 Then Exit For
        Set pp = pp.Previous
    Next k
    If pp Is Nothing Then Exit Function
    ' a headnote such as "(United States dollars)" between the caption and the table
    txt = CleanCaptionText(pp.Range.Text)
    If Left(txt, 1) = "(" And Right(txt, 1) = ")" And Len(txt) < 80 Then
        If StyleName(pp) <> "Title_table" Then
            pp.Style = doc.Styles("Title_table")
            ResetDirect pp
        End If
        CleanUniformFormatting pp
        If pp.Range.Font.Bold <> False Then pp.Range.Font.Bold = False      ' headnotes are not bold
        Set pp = pp.Previous
        If pp Is Nothing Then Exit Function
        If pp.Range.Information(wdWithInTable) Then Exit Function
    End If
    If Not IsCaption(CleanCaptionText(pp.Range.Text), "Table") Then Exit Function
    If StyleName(pp) <> "Title_table" Then
        pp.Style = doc.Styles("Title_table")
        ResetDirect pp
    End If
    NormalizeCaption doc, pp, "Table"
    CleanUniformFormatting pp
    If pp.Range.Font.Bold <> True Then pp.Range.Font.Bold = True            ' caption titles are bold
    FixTableCaption = True
End Function

' Caption text with line breaks turned into spaces and trimmed.
Private Function CleanCaptionText(ByVal t As String) As String
    t = Replace(Replace(Replace(t, vbCr, ""), Chr(11), " "), vbTab, " ")
    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop
    CleanCaptionText = Trim(t)
End Function

' Lays out a caption as "Table 3" + line break + title (IPBES layout), at the text indent.
' The text is rewritten only if it differs, so unchanged captions get no tracked change.
Private Sub NormalizeCaption(ByVal doc As Document, ByVal p As Paragraph, ByVal capWord As String)
    Dim txt As String, rest As String, num As String, i As Long, c As String, want As String, rng As Range
    txt = CleanCaptionText(p.Range.Text)
    If Left(txt, Len(capWord) + 1) <> capWord & " " Then Exit Sub
    rest = Mid(txt, Len(capWord) + 2)
    i = 1
    Do While i <= Len(rest)
        c = Mid(rest, i, 1)
        If Not (c Like "[0-9IVXA.]") Then Exit Do
        i = i + 1
    Loop
    num = Left(rest, i - 1)
    If Right(num, 1) = "." Then num = Left(num, Len(num) - 1)
    rest = Trim(Mid(rest, i))
    Do While Len(rest) > 0 And (Left(rest, 1) = ":" Or Left(rest, 1) = "." Or Left(rest, 1) = "-" Or _
                               Left(rest, 1) = ChrW(8211) Or Left(rest, 1) = ChrW(8212))
        rest = Trim(Mid(rest, 2))
    Loop
    want = capWord & " " & num
    If Len(rest) > 0 Then want = want & Chr(11) & rest
    Set rng = doc.Range(p.Range.Start, p.Range.End - 1)
    ' rewrite only if the caption really differs (spaces around the line break do not count),
    ' so correct captions in IPBES documents get no tracked change
    ' minimal tracked edit (only the part that differs), never a whole-caption replacement
    If CaptionKey(rng.Text) <> CaptionKey(want) Then ApplyDiff doc, rng, want, ""
    If Abs(p.LeftIndent - CentimetersToPoints(2.2)) > 1 Then p.LeftIndent = CentimetersToPoints(2.2)
    If p.FirstLineIndent <> 0 Then p.FirstLineIndent = 0
End Sub

' Caption text without spaces, tabs and non-breaking spaces (line breaks kept), for comparison.
Private Function CaptionKey(ByVal t As String) As String
    t = Replace(t, vbCr, "")
    t = Replace(t, " ", "")
    t = Replace(t, vbTab, "")
    t = Replace(t, ChrW(160), "")
    CaptionKey = t
End Function

' Rebuilds page headers and footers in the IPBES layout: the document symbol in
' Header-pool (left on even pages, right on odd pages) and the page number in
' Footer-pool. The first-page header and footer (masthead, job number) are kept.
' Returns the number of headers and footers rebuilt (0 if no symbol was found).
Private Function FixHeadersFooters(ByVal doc As Document) As Long
    Dim symbol As String, sec As Section, hf As HeaderFooter, i As Long, n As Long, kinds As Variant, k As Variant
    symbol = FindSymbol(doc)
    If Len(symbol) = 0 Then Exit Function
    On Error Resume Next
    doc.PageSetup.OddAndEvenPagesHeaderFooter = True
    kinds = Array(wdHeaderFooterPrimary, wdHeaderFooterEvenPages)
    For Each sec In doc.Sections
        For Each k In kinds
            Set hf = sec.Headers(k)
            If sec.Index = 1 Or Not hf.LinkToPrevious Then
                SetHeaderFooterText hf, symbol, "Header-pool", (k = wdHeaderFooterPrimary), False
                n = n + 1
            End If
            Set hf = sec.Footers(k)
            If sec.Index = 1 Or Not hf.LinkToPrevious Then
                SetHeaderFooterText hf, "", "Footer-pool", (k = wdHeaderFooterPrimary), True
                n = n + 1
            End If
        Next k
    Next sec
    On Error GoTo 0
    FixHeadersFooters = n
End Function

Private Sub SetHeaderFooterText(ByVal hf As HeaderFooter, ByVal txt As String, ByVal sty As String, _
                                ByVal alignRight As Boolean, ByVal pageNumber As Boolean)
    Dim r As Range, t As Long, i As Long, shp As Shape, story As Long, secIdx As Long
    Set r = hf.Range
    story = r.StoryType
    secIdx = r.Sections(1).Index
    ' lines or logos drawn as shapes in this header/footer (HeaderFooter.Shapes lists all
    ' header/footer shapes of the document, so check where each one is anchored)
    On Error Resume Next
    For i = hf.Shapes.Count To 1 Step -1
        Set shp = hf.Shapes(i)
        If shp.Anchor.StoryType = story Then
            If shp.Anchor.Sections(1).Index = secIdx Then shp.Delete
        End If
    Next i
    On Error GoTo 0
    For t = r.Tables.Count To 1 Step -1
        r.Tables(t).Delete
    Next t
    Set r = hf.Range
    r.Text = txt
    Set r = hf.Range
    r.Style = r.Document.Styles(sty)
    r.ParagraphFormat.Reset
    r.Font.Reset
    r.ParagraphFormat.Alignment = IIf(alignRight, wdAlignParagraphRight, wdAlignParagraphLeft)
    If pageNumber Then
        Set r = hf.Range
        r.Collapse wdCollapseStart
        r.Fields.Add r, wdFieldPage
    End If
End Sub

' The document symbol (IPBES/../..) from the page headers or the masthead.
Private Function FindSymbol(ByVal doc As Document) As String
    Dim sec As Section, i As Long, s As String
    On Error Resume Next
    For Each sec In doc.Sections
        For i = 1 To 3
            s = ExtractSymbol(Norm(sec.Headers(i).Range.Text))
            If Len(s) > 0 Then FindSymbol = s: Exit Function
        Next i
    Next sec
    If doc.Tables.Count > 0 Then
        For i = 1 To MinL(4, doc.Tables.Count)
            s = ExtractSymbol(Norm(doc.Tables(i).Range.Text))
            If Len(s) > 0 Then FindSymbol = s: Exit Function
        Next i
    End If
    On Error GoTo 0
End Function

' Shrinks a table proportionally if, after the 2.2 cm indent, it is wider than the text column.
Private Sub FitTableToText(ByVal t As Table)
    Dim avail As Single, c As Cell, widths() As Single, maxW As Single, r As Long, f As Single
    On Error Resume Next
    With t.Range.Sections(1).PageSetup
        avail = .PageWidth - .LeftMargin - .RightMargin - CentimetersToPoints(2.2)
    End With
    ReDim widths(1 To t.Rows.Count + 1)
    For Each c In t.Range.Cells
        r = c.RowIndex
        If r >= 1 And r <= UBound(widths) Then widths(r) = widths(r) + c.Width
    Next c
    For r = 1 To UBound(widths)
        If widths(r) > maxW Then maxW = widths(r)
    Next r
    If maxW <= avail + 1 Or maxW <= 0 Then Exit Sub
    f = avail / maxW
    t.AllowAutoFit = False
    For Each c In t.Range.Cells
        c.Width = c.Width * f
    Next c
    On Error GoTo 0
End Sub

' IPBES table layout: indented to the text, horizontal rules only (top, under the
' column heads, bottom), no shading, italic column heads, numbers right-aligned.
Private Sub FormatTable(ByVal doc As Document, ByVal t As Table)
    Dim c As Cell, txt As String, last As Long
    On Error Resume Next
    t.Rows.LeftIndent = CentimetersToPoints(2.2)
    FitTableToText t
    ' whole-table operations (one change each, instead of one per cell)
    t.Borders.Enable = False
    t.Shading.BackgroundPatternColor = wdColorAutomatic
    t.Range.Cells.Shading.BackgroundPatternColor = wdColorAutomatic
    t.Range.Style = doc.Styles("Normal-pool")
    ' numbers right-aligned: only cells that need it
    For Each c In t.Range.Cells
        If c.RowIndex > 1 Then
            txt = CellText(c)
            If IsNumericCell(txt) Then
                If c.Range.ParagraphFormat.Alignment <> wdAlignParagraphRight Then
                    c.Range.ParagraphFormat.Alignment = wdAlignParagraphRight
                End If
            End If
        End If
    Next c
    With t.Rows(1)
        .HeadingFormat = True
        .Range.Font.Italic = True
        .Range.Font.Bold = False
        .Borders(wdBorderTop).LineStyle = wdLineStyleSingle
        .Borders(wdBorderTop).LineWidth = wdLineWidth050pt
        .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
        .Borders(wdBorderBottom).LineWidth = wdLineWidth050pt
    End With
    last = t.Rows.Count
    t.Rows(last).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
    t.Rows(last).Borders(wdBorderBottom).LineWidth = wdLineWidth050pt
    On Error GoTo 0
End Sub

Private Function IsNumericCell(ByVal txt As String) As Boolean
    Dim t As String
    t = Trim(txt)
    If t = "-" Or t = ".." Or t = ChrW(8211) Or t = "(n/a)" Then IsNumericCell = True: Exit Function
    t = Replace(t, " ", ""): t = Replace(t, ChrW(160), ""): t = Replace(t, ",", "")
    t = Replace(t, "%", ""): t = Replace(t, "$", ""): t = Replace(t, "(", ""): t = Replace(t, ")", "")
    t = Replace(t, ChrW(8211), "-")
    If Len(t) = 0 Then Exit Function
    IsNumericCell = IsNumeric(t)
End Function

' ---------------------------------------------------------------------------
' Track Changes helpers
' ---------------------------------------------------------------------------
' Word refuses to change TrackRevisions while a comment card is selected in the
' comments pane (run-time error 4605). Adding comments can move the focus into
' the new comment, so the focus is moved back to the document text first.
Private Sub FocusDocumentText(ByVal doc As Document)
    On Error Resume Next
    doc.Activate
    doc.ActiveWindow.View.SeekView = wdSeekMainDocument
    If doc.ActiveWindow.Panes.Count > 1 Then doc.ActiveWindow.Panes(1).Activate
    If Selection.StoryType <> wdMainTextStory Then doc.Range(0, 0).Select
    On Error GoTo 0
End Sub

' Switches Track Changes on or off. Returns False (after telling the user) if
' Word does not allow it.
Private Function SetTracking(ByVal doc As Document, ByVal onOff As Boolean) As Boolean
    FocusDocumentText doc
    On Error Resume Next
    doc.TrackRevisions = onOff
    If Err.Number <> 0 Then
        Err.Clear
        doc.Range(0, 0).Select                  ' second attempt with the cursor at the start
        doc.TrackRevisions = onOff
    End If
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        Application.ScreenUpdating = True
        MsgBox "Word did not allow Track Changes to be switched " & IIf(onOff, "on", "off") & "." & vbCr & vbCr & _
               "Click once in the document text (not in a comment), then run the macro again." & _
               IIf(onOff, "", vbCr & "Track Changes may still be on: check Review > Track Changes."), _
               vbExclamation, "IPBES"
        SetTracking = False
        Exit Function
    End If
    On Error GoTo 0
    SetTracking = True
End Function

' ---------------------------------------------------------------------------
' Find helpers
' ---------------------------------------------------------------------------
Private Sub PrepareFind(ByVal rng As Range, r As IpbesRule)
    With rng.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Text = r.FindText
        .MatchWildcards = r.UseWildcards
        .MatchCase = r.MatchCase
        .MatchWholeWord = (r.WholeWord And Not r.UseWildcards)
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Forward = True
        .Wrap = wdFindStop
        .Format = False
    End With
End Sub

' Marks every match in the main text with a comment (up to a limit) and counts
' matches in the main text and footnotes.
Private Function MarkRule(ByVal doc As Document, r As IpbesRule, ByRef sample As String) As Long
    Dim story As Range, rng As Range, n As Long, guard As Long
    For Each story In doc.StoryRanges
        If story.StoryType = wdMainTextStory Or story.StoryType = wdFootnotesStory Then
            Set rng = story.Duplicate
            PrepareFind rng, r
            guard = 0
            Do While rng.Find.Execute
                guard = guard + 1
                If guard > 5000 Then Exit Do
                n = n + 1
                If Len(sample) = 0 Then sample = rng.Text
                If story.StoryType = wdMainTextStory And n <= MAX_COMMENTS_PER_RULE Then
                    doc.Comments.Add rng, TAG & r.Message & IIf(r.AutoFix, " [fixable with IPBES_SafeFixes]", "")
                End If
                rng.Collapse wdCollapseEnd
            Loop
        End If
    Next story
    MarkRule = n
End Function

Private Function CountRule(ByVal doc As Document, r As IpbesRule) As Long
    Dim story As Range, rng As Range, n As Long
    For Each story In doc.StoryRanges
        If story.StoryType = wdMainTextStory Or story.StoryType = wdFootnotesStory Then
            Set rng = story.Duplicate
            PrepareFind rng, r
            Do While rng.Find.Execute
                n = n + 1
                If n > 5000 Then Exit Do
                rng.Collapse wdCollapseEnd
            Loop
        End If
    Next story
    CountRule = n
End Function

' Counts the matches, then replaces them all (main text and footnotes).
' With MatchCase off, Word keeps the capitalization of the found text.
Private Function ReplaceRule(ByVal doc As Document, r As IpbesRule) As Long
    Dim story As Range, rng As Range, n As Long
    n = CountRule(doc, r)
    If n = 0 Then Exit Function
    For Each story In doc.StoryRanges
        If story.StoryType = wdMainTextStory Or story.StoryType = wdFootnotesStory Then
            Set rng = story.Duplicate
            PrepareFind rng, r
            rng.Find.Replacement.Text = r.Replacement
            rng.Find.Execute Replace:=wdReplaceAll
        End If
    Next story
    ReplaceRule = n
End Function

' Converts straight quotes and apostrophes to curly ones, one by one, choosing
' the opening or closing form from the preceding character.
Private Function FixStraightQuotes(ByVal doc As Document) As Long
    Dim story As Range, rng As Range, n As Long, prevChar As String, q As Variant, guard As Long
    For Each q In Array(Chr(34), Chr(39))
        For Each story In doc.StoryRanges
            If story.StoryType = wdMainTextStory Or story.StoryType = wdFootnotesStory Then
                Set rng = story.Duplicate
                With rng.Find
                    .ClearFormatting
                    .Text = CStr(q)
                    .MatchWildcards = True      ' with wildcards Word finds straight quotes only
                    .Forward = True
                    .Wrap = wdFindStop
                End With
                guard = 0
                Do While rng.Find.Execute
                    guard = guard + 1
                    If guard > 5000 Then Exit Do
                    If rng.Start = story.Start Then
                        prevChar = " "
                    Else
                        prevChar = doc.Range(rng.Start - 1, rng.Start).Text
                    End If
                    If IsOpeningContext(prevChar) Then
                        rng.Text = IIf(q = Chr(34), ChrW(8220), ChrW(8216))
                    Else
                        rng.Text = IIf(q = Chr(34), ChrW(8221), ChrW(8217))
                    End If
                    n = n + 1
                    rng.Collapse wdCollapseEnd
                Loop
            End If
        Next story
    Next q
    FixStraightQuotes = n
End Function

Private Function IsOpeningContext(ByVal c As String) As Boolean
    Select Case c
        Case " ", vbCr, vbLf, vbTab, Chr(11), "(", "[", "{", ChrW(8211), ChrW(8212), "-", "/", ChrW(160)
            IsOpeningContext = True
        Case Else
            IsOpeningContext = False
    End Select
End Function

' ---------------------------------------------------------------------------
' Masthead check (only if the document starts with a masthead table)
' ---------------------------------------------------------------------------
Private Function CheckMasthead(ByVal doc As Document) As Long
    Dim t1 As Table, i As Long, n As Long, txt As String, symbol As String
    Dim logoTable As Table, sessionTable As Table, symbolTable As Table, hdr As String
    Dim ratio As Double, sh As InlineShape

    If doc.Tables.Count = 0 Then Exit Function
    Set t1 = doc.Tables(1)
    txt = UCase(Norm(t1.Range.Text))
    If InStr(txt, "UNITED") = 0 Or InStr(txt, "NATIONS") = 0 Then Exit Function
    If t1.Range.Start > 200 Then Exit Function

    ' Locate the blocks among the first five tables
    For i = 1 To MinL(5, doc.Tables.Count)
        txt = Norm(doc.Tables(i).Range.Text)
        If InStr(txt, "Distr.") > 0 Then Set logoTable = doc.Tables(i)
        If InStr(txt, "session") > 0 And InStr(txt, "Plenary") > 0 Then Set sessionTable = doc.Tables(i)
        If symbolTable Is Nothing And InStr(txt, "IPBES/") > 0 Then Set symbolTable = doc.Tables(i)
    Next i

    ' 1. UN and partner logo strip
    If t1.Range.InlineShapes.Count = 0 Then
        n = n + Flag(doc, t1.Range, "Masthead: the UN and partner logo strip (UN emblem, UNEP, UNESCO, FAO, UNDP) is missing.")
    ElseIf t1.Range.InlineShapes.Count > 1 Then
        n = n + Flag(doc, t1.Range, "Masthead: " & t1.Range.InlineShapes.Count & " separate logo images. The current masthead uses " & _
            "one black-and-white strip: UN emblem, UNEP, UNESCO, FAO, UNDP (see assets/masthead).")
    Else
        Set sh = t1.Range.InlineShapes(1)
        If sh.Height > 0 Then
            ratio = sh.Width / sh.Height
            If Abs(ratio - 9.88) / 9.88 > 0.05 Then
                n = n + Flag(doc, sh.Range, "Masthead: the logo strip has a width/height ratio of " & Format(ratio, "0.0") & _
                    " (expected about 9.9). It may be distorted, cropped or the wrong image. Check that all five logos " & _
                    "(UN emblem, UNEP, UNESCO, FAO, UNDP) are present and in that order.")
            End If
        End If
    End If

    ' 2. IPBES logo
    If logoTable Is Nothing Then
        n = n + Flag(doc, t1.Range, "Masthead: no block with 'Distr.: General' found (IPBES logo, Platform name, distribution, date, 'Original: English').")
    Else
        If logoTable.Range.InlineShapes.Count = 0 Then
            n = n + Flag(doc, logoTable.Range, "Masthead: the IPBES logo is missing (black logo without strapline).")
        Else
            Set sh = logoTable.Range.InlineShapes(1)
            If sh.Height > 0 Then
                ratio = sh.Width / sh.Height
                If Abs(ratio - 2.15) / 2.15 > 0.08 Then
                    n = n + Flag(doc, sh.Range, "Masthead: the IPBES logo has a width/height ratio of " & Format(ratio, "0.00") & _
                        " (expected about 2.15 for the black logo without strapline). It may be distorted, include the " & _
                        "strapline or be the wrong version.")
                End If
            End If
            n = n + Flag(doc, sh.Range, "Masthead: confirm this is the BLACK IPBES logo without strapline (not colour, not reversed).", True)
        End If
        txt = Norm(logoTable.Range.Text)
        If InStr(txt, "Distr.: General") = 0 Then n = n + Flag(doc, logoTable.Range, "Masthead: 'Distr.: General' missing or misspelled.")
        If InStr(txt, "Original: English") = 0 Then n = n + Flag(doc, logoTable.Range, "Masthead: 'Original: English' missing or misspelled.")
        If Not HasDate(txt) Then n = n + Flag(doc, logoTable.Range, "Masthead: date of issue missing or not in the form '12 February 2026'.")
        If InStr(txt, FullName()) = 0 Then n = n + Flag(doc, logoTable.Range, "Masthead: the full Platform name is missing or altered: " & FullName() & ".")
    End If

    ' 3. Symbol and page headers
    If symbolTable Is Nothing Then
        n = n + Flag(doc, t1.Range, "Masthead: document symbol (IPBES/../..) not found.")
    Else
        symbol = ExtractSymbol(Norm(symbolTable.Range.Text))
        If Len(symbol) > 0 Then
            hdr = Norm(doc.Sections(1).Headers(wdHeaderFooterPrimary).Range.Text) & " " & _
                  Norm(doc.Sections(1).Headers(wdHeaderFooterEvenPages).Range.Text)
            If doc.Sections.Count > 1 Then hdr = hdr & " " & Norm(doc.Sections(2).Headers(wdHeaderFooterPrimary).Range.Text)
            If InStr(hdr, symbol) = 0 Then
                n = n + Flag(doc, symbolTable.Range, "Masthead: symbol " & symbol & " does not appear in the page headers. Check the headers.")
            End If
        End If
    End If

    ' 4. Session block
    If sessionTable Is Nothing Then
        n = n + Flag(doc, t1.Range, "Masthead: session block not found ('Plenary of the ... / Twelfth session / City, Country, dates').")
    Else
        txt = Norm(sessionTable.Range.Text)
        If InStr(txt, "Plenary of the " & FullName()) = 0 Then
            n = n + Flag(doc, sessionTable.Range, "Masthead: session block should read 'Plenary of the " & FullName() & "'.")
        End If
        If txt Like "*[0-9]-[0-9]*" Then
            n = n + Flag(doc, sessionTable.Range, "Masthead: use an en dash in the date range (3" & EnDash() & "8 February 2026).")
        End If
    End If

    ' 5. Spacing: empty paragraphs between the masthead blocks and the title
    n = n + CheckMastheadSpacing(doc)

    ' 6. Double spaces inside masthead tables
    For i = 1 To MinL(4, doc.Tables.Count)
        If InStr(Replace(doc.Tables(i).Range.Text, Chr(11), " "), "  ") > 0 Then
            n = n + Flag(doc, doc.Tables(i).Range, "Masthead: double space in this block.")
        End If
    Next i

    If n > 0 Then ReportLines.Add "Masthead" & vbTab & n & vbTab & "" & vbTab & "See comments in the masthead" & vbTab & ""
    CheckMasthead = n
End Function

Private Function CheckMastheadSpacing(ByVal doc As Document) As Long
    ' Flags empty paragraphs (other than A_Spacer) between the masthead blocks and the title.
    Dim p As Paragraph, n As Long, sName As String, count As Long
    For Each p In doc.Paragraphs
        count = count + 1
        If count > 60 Then Exit For
        If Not p.Range.Information(wdWithInTable) Then
            sName = StyleName(p)
            If sName = "BB_Title" Or sName = "CH1" Then Exit For
            If sName <> "A_Spacer" Then
                If Len(Trim(Replace(Replace(p.Range.Text, vbCr, ""), Chr(11), ""))) = 0 Then
                    n = n + Flag(doc, p.Range, "Masthead spacing: empty paragraph in style '" & sName & _
                        "'. Between masthead blocks use exactly one paragraph in style A_Spacer; remove extra empty paragraphs.")
                End If
            End If
        End If
    Next p
    CheckMastheadSpacing = n
End Function

Private Function HasDate(ByVal txt As String) As Boolean
    Dim m As Variant
    For Each m In Array("January", "February", "March", "April", "May", "June", "July", "August", _
                        "September", "October", "November", "December")
        If txt Like "*# " & m & " ####*" Then HasDate = True: Exit Function
    Next m
End Function

Private Function ExtractSymbol(ByVal txt As String) As String
    Dim p As Long, e As Long
    p = InStr(txt, "IPBES/")
    If p = 0 Then Exit Function
    e = p
    Do While e <= Len(txt)
        If Mid(txt, e, 1) = " " Then Exit Do
        e = e + 1
    Loop
    ExtractSymbol = Mid(txt, p, e - p)
End Function

' ---------------------------------------------------------------------------
' Heading checks: numbering gaps, lone subheadings, final full stops
' ---------------------------------------------------------------------------
Private Function CheckHeadings(ByVal doc As Document) As Long
    Dim p As Paragraph, sName As String, label As String, txt As String, n As Long
    Dim lastCH1 As Long, lastCH2 As Long, ch2Count As Long, lastCH1Range As Range, v As Long
    lastCH1 = 0: lastCH2 = 0: ch2Count = 0
    For Each p In doc.Paragraphs
        sName = StyleName(p)
        If sName = "CH1" Or sName = "CH2" Then
            SplitHeading p.Range.Text, label, txt
            If sName = "CH1" Then
                If ch2Count = 1 And Not lastCH1Range Is Nothing Then
                    n = n + Flag(doc, lastCH1Range, "Headings: this chapter has a single subheading (an 'A' needs a 'B').")
                End If
                ch2Count = 0: lastCH2 = 0
                If Len(label) = 0 Then
                    lastCH1 = 0                       ' decision or annex title: numbering restarts
                Else
                    v = RomanValue(Replace(label, ".", ""))
                    If v > 0 Then
                        If v <> lastCH1 + 1 Then n = n + Flag(doc, p.Range, "Headings: chapter numbering jumps from " & _
                            IIf(lastCH1 = 0, "the start", RomanText(lastCH1)) & " to " & label & ".")
                        lastCH1 = v
                    End If
                End If
                Set lastCH1Range = p.Range
            Else
                If Len(label) = 1 Or (Len(label) = 2 And Right(label, 1) = ".") Then
                    v = Asc(UCase(Left(label, 1))) - 64
                    If v >= 1 And v <= 26 Then
                        ch2Count = ch2Count + 1
                        If v <> lastCH2 + 1 Then n = n + Flag(doc, p.Range, "Headings: section lettering jumps to " & label & ".")
                        lastCH2 = v
                    End If
                End If
            End If
            If Len(txt) > 1 Then
                If Right(txt, 1) = "." And Not (Right(txt, 4) Like "#.#." Or Right(txt, 3) = "etc." Or Right(txt, 4) = "Add." ) Then
                    n = n + Flag(doc, p.Range, "Headings: no full stop at the end of a heading.")
                End If
            End If
        End If
    Next p
    If ch2Count = 1 And Not lastCH1Range Is Nothing Then
        n = n + Flag(doc, lastCH1Range, "Headings: this chapter has a single subheading (an 'A' needs a 'B').")
    End If
    If n > 0 Then ReportLines.Add "Headings" & vbTab & n & vbTab & "" & vbTab & "Numbering gaps, lone subheadings or final full stops" & vbTab & ""
    CheckHeadings = n
End Function

' ---------------------------------------------------------------------------
' Caption checks: style, layout, punctuation, formatting, numbering, position
' ---------------------------------------------------------------------------
Private Function CheckCaptions(ByVal doc As Document) As Long
    Dim p As Paragraph, nxt As Paragraph, txt As String, t2 As String, sName As String, n As Long
    Dim capWord As String, expected As String, num As String, sep As String, rest As String
    Dim v As Long, lastT As Long, lastF As Long, sinceAnnex As Boolean, nearTable As Boolean, k As Long
    Dim sf As Font
    For Each p In doc.Paragraphs
        sName = StyleName(p)
        txt = CleanCaptionText(p.Range.Text)
        If sName = "ZZ_Anx_header" Or txt = "Annex" Or Left(txt, 6) = "Annex " Then sinceAnnex = True
        capWord = ""
        If IsCaption(txt, "Table") Then
            capWord = "Table"
        ElseIf IsCaption(txt, "Figure") Then
            capWord = "Figure"
        End If
        If Len(capWord) > 0 Then
            If Not p.Range.Information(wdWithInTable) Then
                expected = IIf(capWord = "Table", "Title_table", "Title_figure")
                num = "": sep = "": rest = ""
                SplitCaption txt, capWord, num, sep, rest
                ' style
                If sName <> expected And Not (capWord = "Figure" And sName = "Title_table") Then
                    n = n + Flag(doc, p.Range, "Caption: use the style " & expected & ".")
                End If
                ' layout and punctuation
                If Len(sep) > 0 Then
                    n = n + Flag(doc, p.Range, "Caption: no punctuation after the number ('" & capWord & " " & num & "', then a line break, then the title).")
                ElseIf Len(rest) > 0 And InStr(p.Range.Text, Chr(11)) = 0 Then
                    n = n + Flag(doc, p.Range, "Caption layout: '" & capWord & " " & num & "', then a line break (Shift+Enter), then the title.")
                End If
                If Right(rest, 1) = "." Then n = n + Flag(doc, p.Range, "Caption: no full stop at the end of the title.")
                ' formatting
                Set sf = Nothing
                On Error Resume Next
                Set sf = p.Style.Font
                On Error GoTo 0
                If p.Range.Font.Bold <> True Then
                    n = n + Flag(doc, p.Range, "Caption: the number and title are bold (" & expected & ").")
                ElseIf Not sf Is Nothing Then
                    If (Len(p.Range.Font.Name) > 0 And p.Range.Font.Name <> sf.Name) Or _
                       (p.Range.Font.Size <> wdUndefined And p.Range.Font.Size <> sf.Size) Or _
                       (p.Range.Font.Color <> wdUndefined And p.Range.Font.Color <> sf.Color) Then
                        n = n + Flag(doc, p.Range, "Caption: font, size or colour differs from the style " & expected & " (Fix formatting removes it).")
                    End If
                End If
                ' numbering
                v = 0
                If num Like "#*" And IsNumeric(num) Then
                    v = CLng(num)
                ElseIf RomanValue(num) > 0 Then
                    v = RomanValue(num)
                    If capWord = "Table" Then n = n + Flag(doc, p.Range, "Tables are numbered with Arabic numerals (Table 1, Table 2 ...).")
                End If
                If v > 0 Then
                    If capWord = "Table" Then
                        If Not (v = lastT + 1 Or (v = 1 And sinceAnnex)) Then
                            n = n + Flag(doc, p.Range, "Table numbering: " & capWord & " " & num & " follows table " & lastT & _
                                IIf(v = 1, " (numbering restarts only in a new annex).", "."))
                        End If
                        lastT = v
                    Else
                        If Not (v = lastF + 1 Or (v = 1 And sinceAnnex)) Then
                            n = n + Flag(doc, p.Range, "Figure numbering: " & capWord & " " & num & " follows figure " & lastF & ".")
                        End If
                        lastF = v
                    End If
                    If v = 1 Then sinceAnnex = False
                End If
                ' position (tables): directly above the table, with at most a headnote and empty lines between
                If capWord = "Table" Then
                    nearTable = False
                    Set nxt = p.Next
                    k = 0
                    Do While Not nxt Is Nothing And k < 4
                        If nxt.Range.Information(wdWithInTable) Then nearTable = True: Exit Do
                        t2 = CleanCaptionText(nxt.Range.Text)
                        If Len(t2) > 0 Then
                            If Left(t2, 1) = "(" And Right(t2, 1) = ")" Then
                                If StyleName(nxt) <> "Title_table" Then
                                    n = n + Flag(doc, nxt.Range, "Headnote: use the style Title_table (not bold).")
                                ElseIf nxt.Range.Font.Bold = True Then
                                    n = n + Flag(doc, nxt.Range, "Headnote: not bold.")
                                End If
                            Else
                                Exit Do
                            End If
                        End If
                        Set nxt = nxt.Next
                        k = k + 1
                    Loop
                    If Not nearTable Then n = n + Flag(doc, p.Range, "Caption: should be directly above its table (only a headnote such as '(United States dollars)' in between).")
                End If
            End If
        End If
    Next p
    If n > 0 Then ReportLines.Add "Captions" & vbTab & n & vbTab & "" & vbTab & _
        "Table and figure captions: style, layout, punctuation, formatting, numbering, position" & vbTab & ""
    CheckCaptions = n
End Function

' Splits a caption into its number, the punctuation after it (if any) and the title.
Private Sub SplitCaption(ByVal txt As String, ByVal capWord As String, ByRef num As String, _
                         ByRef sep As String, ByRef rest As String)
    Dim i As Long, c As String, r As String
    r = Mid(txt, Len(capWord) + 2)
    i = 1
    Do While i <= Len(r)
        c = Mid(r, i, 1)
        If Not (c Like "[0-9IVXA.]") Then Exit Do
        i = i + 1
    Loop
    num = Left(r, i - 1)
    If Right(num, 1) = "." Then
        num = Left(num, Len(num) - 1)
        sep = "."
    End If
    rest = Trim(Mid(r, i))
    If Len(rest) > 0 Then
        c = Left(rest, 1)
        If c = ":" Or c = "-" Or c = "." Or c = ChrW(8211) Or c = ChrW(8212) Then
            sep = c
            rest = Trim(Mid(rest, 2))
        End If
    End If
End Sub

Private Sub SplitHeading(ByVal raw As String, ByRef label As String, ByRef txt As String)
    Dim parts() As String, i As Long, bits As New Collection
    raw = Replace(Replace(raw, vbCr, ""), Chr(11), " ")
    parts = Split(raw, vbTab)
    For i = LBound(parts) To UBound(parts)
        If Len(Trim(parts(i))) > 0 Then bits.Add Trim(parts(i))
    Next i
    label = "": txt = ""
    If bits.Count = 0 Then Exit Sub
    If bits.Count >= 2 And Len(bits(1)) <= 6 Then
        label = bits(1)
        For i = 2 To bits.Count
            txt = txt & IIf(Len(txt) > 0, " ", "") & bits(i)
        Next i
    Else
        For i = 1 To bits.Count
            txt = txt & IIf(Len(txt) > 0, " ", "") & bits(i)
        Next i
    End If
End Sub

Private Function RomanValue(ByVal s As String) As Long
    Dim i As Long, v As Long, cur As Long, nxt As Long
    s = UCase(Trim(s))
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        cur = RomanDigit(Mid(s, i, 1))
        If cur = 0 Then RomanValue = 0: Exit Function
        If i < Len(s) Then nxt = RomanDigit(Mid(s, i + 1, 1)) Else nxt = 0
        If cur < nxt Then v = v - cur Else v = v + cur
    Next i
    RomanValue = v
End Function

Private Function RomanDigit(ByVal c As String) As Long
    Select Case c
        Case "I": RomanDigit = 1
        Case "V": RomanDigit = 5
        Case "X": RomanDigit = 10
        Case "L": RomanDigit = 50
        Case "C": RomanDigit = 100
        Case Else: RomanDigit = 0
    End Select
End Function

Private Function RomanText(ByVal v As Long) As String
    Dim vals As Variant, syms As Variant, i As Long
    vals = Array(100, 90, 50, 40, 10, 9, 5, 4, 1)
    syms = Array("C", "XC", "L", "XL", "X", "IX", "V", "IV", "I")
    For i = 0 To 8
        Do While v >= vals(i)
            RomanText = RomanText & syms(i)
            v = v - vals(i)
        Loop
    Next i
End Function

' ---------------------------------------------------------------------------
' Utilities
' ---------------------------------------------------------------------------
Private Function Flag(ByVal doc As Document, ByVal rng As Range, ByVal msg As String, Optional ByVal quiet As Boolean = False) As Long
    On Error Resume Next
    doc.Comments.Add rng, TAG & msg
    On Error GoTo 0
    If Not quiet Then Flag = 1
End Function

Private Function MinL(ByVal a As Long, ByVal b As Long) As Long
    If a < b Then MinL = a Else MinL = b
End Function

Private Function StyleName(ByVal p As Paragraph) As String
    On Error Resume Next
    StyleName = p.Style.NameLocal
    On Error GoTo 0
End Function

' Normalizes text from Word: non-breaking hyphens/spaces, line breaks, cell marks.
Private Function Norm(ByVal s As String) As String
    s = Replace(s, Chr(30), "-")
    s = Replace(s, ChrW(8209), "-")
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, Chr(11), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, Chr(7), " ")
    s = Replace(s, vbTab, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Norm = Trim(s)
End Function

Private Sub WriteReport(ByVal src As Document, ByVal total As Long, ByVal title As String, _
                        Optional ByVal headers As String = "Category|Count|Safe fix|Rule|Example")
    Dim rep As Document, rng As Range, i As Long, tbl As Table, parts() As String, ln As Variant, h() As String
    Set rep = Documents.Add
    Set rng = rep.Content
    rng.Text = "IPBES " & title & ": " & src.Name & vbCr & _
        Format(Now, "d mmmm yyyy, h.nn") & vbCr & _
        total & " item(s) found." & IIf(total > 0, " Comments starting " & TAG & "were added to the document " & _
        "(at most " & MAX_COMMENTS_PER_RULE & " per rule). Remove them with IPBES_RemoveIPBESComments.", "") & vbCr & _
        "Do not change agreed text (decisions, approved summaries for policymakers, quotations): treat issues there as queries." & vbCr
    rep.Paragraphs(1).Range.Font.Bold = True
    rep.Paragraphs(1).Range.Font.Size = 14
    If ReportLines.Count = 0 Then Exit Sub
    Set rng = rep.Content
    rng.Collapse wdCollapseEnd
    Set tbl = rep.Tables.Add(rng, ReportLines.Count + 1, 5)
    tbl.Borders.Enable = True
    h = Split(headers, "|")
    For i = 0 To 4
        If i <= UBound(h) Then tbl.Cell(1, i + 1).Range.Text = h(i)
    Next i
    tbl.Rows(1).Range.Font.Bold = True
    i = 1
    For Each ln In ReportLines
        i = i + 1
        parts = Split(CStr(ln), vbTab)
        tbl.Cell(i, 1).Range.Text = parts(0)
        tbl.Cell(i, 2).Range.Text = parts(1)
        tbl.Cell(i, 3).Range.Text = parts(2)
        tbl.Cell(i, 4).Range.Text = parts(3)
        If UBound(parts) >= 4 Then tbl.Cell(i, 5).Range.Text = parts(4)
    Next ln
    rep.Saved = True
End Sub
