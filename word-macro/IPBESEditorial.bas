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
'   IPBES_RemoveIPBESComments  Deletes only the comments added by these macros.
'
' Rules: skills/ipbes-editorial/references/house-style.md (UN Editorial Manual
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
        sample = ""
        n = MarkRule(doc, Rules(i), sample)
        If n > 0 Then
            total = total + n
            ReportLines.Add Rules(i).Category & vbTab & n & vbTab & _
                IIf(Rules(i).AutoFix, "yes", "") & vbTab & Rules(i).Message & vbTab & Left(sample, 60)
        End If
    Next i

    total = total + CheckMasthead(doc)
    total = total + CheckHeadings(doc)

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
    doc.TrackRevisions = True
    Application.ScreenUpdating = False
    BuildRules
    Set ReportLines = New Collection

    For i = 1 To RuleCount
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

    n = FixStraightQuotes(doc)
    If n > 0 Then
        total = total + n
        ReportLines.Add "Punctuation" & vbTab & n & vbTab & "fixed" & vbTab & "straight quotes -> curly" & vbTab & ""
    End If

    doc.TrackRevisions = oldTrack
    Application.ScreenUpdating = True
    msg = total & " change(s) made with Track Changes." & vbCr & _
          "Review them (Review > Next) and reject any inside agreed text."
    MsgBox msg, vbInformation, "IPBES"
    WriteReport doc, total, "Safe fixes"
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

Private Sub WriteReport(ByVal src As Document, ByVal total As Long, ByVal title As String)
    Dim rep As Document, rng As Range, i As Long, tbl As Table, parts() As String, line As Variant
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
    tbl.Cell(1, 1).Range.Text = "Category"
    tbl.Cell(1, 2).Range.Text = "Count"
    tbl.Cell(1, 3).Range.Text = "Safe fix"
    tbl.Cell(1, 4).Range.Text = "Rule"
    tbl.Cell(1, 5).Range.Text = "Example"
    tbl.Rows(1).Range.Font.Bold = True
    i = 1
    For Each line In ReportLines
        i = i + 1
        parts = Split(CStr(line), vbTab)
        tbl.Cell(i, 1).Range.Text = parts(0)
        tbl.Cell(i, 2).Range.Text = parts(1)
        tbl.Cell(i, 3).Range.Text = parts(2)
        tbl.Cell(i, 4).Range.Text = parts(3)
        If UBound(parts) >= 4 Then tbl.Cell(i, 5).Range.Text = parts(4)
    Next line
    rep.Saved = True
End Sub
