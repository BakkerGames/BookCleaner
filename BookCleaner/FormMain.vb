' --------------------------------
' --- FormMain.vb - 06/29/2016 ---
' --------------------------------

' ----------------------------------------------------------------------------------------------------
' 06/29/2016 - SBakker
'            - Had to take "melange" out of the list for accents. "Dune" doesn't use the accents.
' 01/17/2015 - SBakker
'            - Removed HypenQuote logic, as it didn't do anything anymore.
'            - Added unaccented words "señor", "señora", "señorita".
' 08/20/2014 - SBakker
'            - Fixed HasErrorLines to work correctly and show message under files with "###" errors.
' 07/08/2014 - SBakker
'            - Leave <code> lines alone, no modification or error checking.
'            - Fixed unaccented word check to really have a lowercase CurrLine sent. Also lowercase
'              all words to be accented.
' 06/04/2014 - SBakker
'            - Added "coup de grâce", "smörgâsbord", and "Yucatán".
' 06/03/2014 - SBakker
'            - Ignore spacer lines with only a Null (&#0;).
' 05/24/2014 - SBakker
'            - Added ButtonStop to stop safely in the middle.
' 05/20/2014 - SBakker
'            - Fixed the logic for left and right quotes, so left quotes can't have a letter or a
'              quote-ending punctuation symbol (, . ? !) before it.
' 05/17/2014 - SBakker
'            - Added unaccented word "à la".
' 05/11/2014 - SBakker
'            - Added unaccented word "d'hôtel".
' 05/04/2014 - SBakker
'            - Added unaccented word "touché".
'            - Added bad formatting rule for " " and ' '.
' 04/29/2014 - SBakker
'            - Check for symbols after "..." and leave a space between them.
'            - Added checking for commonly accented words which are missing their accent.
'            - Added checking for bad formatting.
' 03/15/2014 - SBakker
'            - Added Bootstrapping to a a local area instead of using the ClickOnce install.
'            - Added the Settings Provider to save settings in the local area with the program.
'            - Added AboutMain.vb which shows the current path in the Status Bar.
' 11/10/2013 - SBakker
'            - Change "^ —" to "^—".
' 07/28/2013 - SBakker
'            - Removed ChangeHyphenSpaceMsg for now. Want to replace hyphen-space with an
'              HTML equivalent when I can find one.
'            - Added StatusBarMain to handle messages, so there's no popup message box.
' 04/01/2013 - SBakker
'            - Force all books to be UTF-8 if they are CodePage 1252 (ASCII).
' 03/10/2013 - SBakker
'            - Check for mismatched <i></i>, <b></b>, and <u></u> tags on the same line.
' 01/06/2013 - SBakker
'            - Make sure single and double quotes get fixed on EVERY line.
' 12/27/2012 - SBakker
'            - Removed old debug code.
' 08/15/2012 - SBakker
'            - Changed " <i>... " to " <i>...".
' 07/11/2012 - SBakker
'            - Change "images/" to "images\" so they are proper Windows paths.
'            - Don't add extra blank lines before lines starting with "~", "^", "|".
' 04/17/2012 - SBakker
'            - Check for ">" and "<" as well as letters when doing left/right quote check.
'              Then it will pick up the <i> tags as well.
' 04/15/2012 - SBakker
'            - Fixed so \" is ignored during left and right quote matchings. It will have to
'              be changed to a normal " during TXT2HTML.
' 04/14/2012 - SBakker
'            - Added mismatched quotes message, to find where left and right quotes don't go
'              in the proper sequence.
' 04/06/2012 - SBakker
'            - Added message "Has ### errors:" so that they aren't accidentally missed in
'              files which have already had the error lines added.
' 03/25/2012 - SBakker
'            - Removed space after vbTab + "<i>... " and fixed "'..." or "...'" issues.
' 03/19/2012 - SBakker
'            - Tried to fix issues with tags.
' 03/11/2012 - SBakker
'            - Check for letters OR digits before and after "...", and add a space after if
'              found.
' 02/20/2012 - SBakker
'            - Don't alter lines with "<pre>" tags.
' 02/17/2012 - SBakker
'            - Changed ""..."" to ""... "".
' 12/22/2011 - SBakker
'            - Fixed case with " <i> — " and " — </i> " making two spaces in the output, not
'              just one.
'            - Fixed so [a-z]..."[a-z] is replaced by [a-z]..." [a-z]. If this is incorrect,
'              putting the space outside the quote will be preserved properly.
' 10/25/2011 - SBakker
'            - Working on BookCleaner.
' ----------------------------------------------------------------------------------------------------

Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
Imports Arena_Utilities.FileUtils

Public Class FormMain

    Private Shared ReadOnly ObjName As String = System.Reflection.MethodBase.GetCurrentMethod().ReflectedType.FullName

    Private Const OddNumberOfQuotesMsg As String = "### Odd number of quotes above ###"
    Private Const MismatchedQuotesMsg As String = "### Mismatched left and right quotes ###"
    Private Const ChangeHyphenSpaceMsg As String = "### Change ""- "" to ""-·"" or "" — "" ###"
    Private Const MismatchedTagMsg_I As String = "### Mismatched number of <i></i> tags ###"
    Private Const MismatchedTagMsg_B As String = "### Mismatched number of <b></b> tags ###"
    Private Const MismatchedTagMsg_U As String = "### Mismatched number of <u></u> tags ###"
    Private Const NeedsAccentMsg As String = "### Word {1} above needs accent: {2} ###"
    Private Const BadFormattingMsg As String = "### Bad formatting: {1} ###"

    Private Const MaxBlankLines As Integer = 2

    Private UnaccentedWord As String = ""
    Private AccentedWord As String = ""
    Private BadFormattingPhrase As String = ""
    Private StopRequested As Boolean = False

    Private Sub FormMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Static FuncName As String = ObjName + "." + System.Reflection.MethodBase.GetCurrentMethod().Name

        Try
            If Arena_Bootstrap.BootstrapClass.CopyProgramsToLaunchPath Then
                Me.Close()
                Exit Sub
            End If
        Catch ex As Exception
            MessageBox.Show(FuncName + vbCrLf + ex.Message, My.Application.Info.AssemblyName, MessageBoxButtons.OK)
            Me.Close()
            Exit Sub
        End Try

        ' --- First call Upgrade to load setting from last version ---
        If My.Settings.CallUpgrade Then
            My.Settings.Upgrade()
            My.Settings.CallUpgrade = False
            My.Settings.Save()
        End If

        TextBoxFromPath.Text = My.Settings.DefaultPath

    End Sub

    Private Sub ButtonStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ButtonStart.Click
        Dim FileCount As Integer = 0
        Dim FileCountMsg As String = "Files Checked = "
        ' ---------------------------------------------
        If TextBoxFromPath.Text = "" Then Exit Sub
        If Not Directory.Exists(TextBoxFromPath.Text) Then Exit Sub
        ' --- Begin ---
        StopRequested = False
        ButtonStop.Enabled = True
        ButtonStop.Visible = True
        ButtonStart.Enabled = False
        ButtonStart.Visible = False
        ButtonStop.Focus()
        My.Settings.DefaultPath = TextBoxFromPath.Text
        My.Settings.Save()
        TextBoxFromPath.ReadOnly = True
        TextBoxResults.Text = ""
        ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString
        Application.DoEvents()
        Dim FromFiles() As String = Directory.GetFiles(TextBoxFromPath.Text, "*.txt")
        For Each FileName As String In FromFiles
            If StopRequested Then Exit For
            FixBook(FileName)
            FileCount += 1
            ToolStripStatusLabelMain.Text = FileCountMsg + FileCount.ToString
            Application.DoEvents()
        Next
        ' --- Done ---
        ToolStripStatusLabelMain.Text += " - Done"
        TextBoxFromPath.ReadOnly = False
        ButtonStart.Enabled = True
        ButtonStart.Visible = True
        ButtonStop.Enabled = False
        ButtonStop.Visible = False
        ButtonStart.Focus()
    End Sub

    Private Sub FixBook(ByVal FileName As String)
        Dim FromPath As String = FileName.Substring(0, FileName.LastIndexOf("\"c))
        Dim BaseFileName As String = FileName.Substring(FileName.LastIndexOf("\"c) + 1)
        Dim BaseFileNameNoExt As String = BaseFileName.Substring(0, BaseFileName.LastIndexOf("."c))
        Dim HeaderText As New StringBuilder
        Dim TargetText As New StringBuilder
        Dim FoundCover As Boolean = False
        Dim FoundTimestamp As Boolean = False
        Dim FileChanged As Boolean = False
        Dim LastLineError As Boolean = False
        Dim QuoteCount As Integer = 0
        Dim MismatchedQuotes As Boolean = False
        Dim MismatchedTag_I As Boolean = False
        Dim MismatchedTag_B As Boolean = False
        Dim MismatchedTag_U As Boolean = False
        Dim NeedsAccent As Boolean = False
        Dim BadFormatting As Boolean = False
        Dim PeriodCount As Integer = 0
        Dim OddNumQuotes As Boolean = False
        Dim FoundFirstChapter As Boolean = False
        Dim OrigLine As String
        Dim LastChar As Char = " "c
        Dim LastLine As String = ""
        Dim NumBlankLines As Integer = 0
        Dim AfterChapter As Boolean = False
        Dim BlockIndent As Boolean = False
        Dim HasErrorLines As Boolean = False
        Dim PoetryFile As Boolean = False
        ' ----------------------------------
        TextBoxFileName.Text = BaseFileName
        ' --- Fill in the lines ---
        Dim OrigEncoding As Encoding = GetFileEncoding(FileName)
        If OrigEncoding.EncodingName <> Encoding.UTF8.EncodingName Then
            FileChanged = True
        End If
        If OrigEncoding.CodePage = 1252 Then
            OrigEncoding = Encoding.UTF8
        End If
        HasErrorLines = False
        For Each CurrLine As String In File.ReadAllLines(FileName, OrigEncoding)
            ' --- Remove old error messages ---
            If CurrLine.StartsWith("###") Then
                LastLineError = True
                FileChanged = True
                Continue For ' --- Remove these lines, may be re-added below ---
            End If
            ' --- Ignore spacer lines with only a Null ---
            If CurrLine = "&#0;" Then
                TargetText.AppendLine(CurrLine)
                Continue For
            End If
            ' --- Handle <code> lines without modification ---
            If CurrLine.Contains("<code>") Then
                TargetText.AppendLine(CurrLine)
                Continue For
            End If
            ' --- Process current line ---
            OrigLine = CurrLine
            If String.IsNullOrWhiteSpace(CurrLine) And CurrLine <> "" Then
                CurrLine = ""
            End If
            ' --- Handle single quotes ---
            If CurrLine.Contains("`") Then CurrLine = CurrLine.Replace("`", "'")
            If CurrLine.Contains("‘") Then CurrLine = CurrLine.Replace("‘", "'")
            If CurrLine.Contains("’") Then CurrLine = CurrLine.Replace("’", "'")
            ' --- Handle double quotes ---
            If CurrLine.Contains(Chr(147)) Then CurrLine = CurrLine.Replace(Chr(147), """")
            If CurrLine.Contains(Chr(148)) Then CurrLine = CurrLine.Replace(Chr(148), """")
            If CurrLine.Contains("«") Then CurrLine = CurrLine.Replace("«", """")
            If CurrLine.Contains("»") Then CurrLine = CurrLine.Replace("»", """")
            If CurrLine.Contains("„") Then CurrLine = CurrLine.Replace("„", """")
            If CurrLine.Contains("''") Then CurrLine = CurrLine.Replace("''", """")
            ' --- Check for metadata ---
            If CurrLine.StartsWith("<") Then
                ' --- Ignore poetry ---
                If CurrLine.ToLower.Contains("content=""poetry""") Then
                    PoetryFile = True
                End If
                If CurrLine.ToLower.Contains("name=""timestamp""") Then
                    FoundTimestamp = True
                    Continue For ' --- Don't add this line yet ---
                End If
            ElseIf Not PoetryFile AndAlso
                Not CurrLine.Contains("<pre>") Then
                ' --- Normal lines ---
                If FoundTimestamp AndAlso CurrLine <> "" AndAlso Not CurrLine.StartsWith(vbTab) Then
                    FoundFirstChapter = True
                End If
                ' --- Handle errors from the last line ---
                If MismatchedQuotes Then
                    If Not FoundTimestamp Then
                        HeaderText.AppendLine(MismatchedQuotesMsg)
                    Else
                        TargetText.AppendLine(MismatchedQuotesMsg)
                    End If
                    If Not LastLineError Then
                        FileChanged = True
                    End If
                    MismatchedQuotes = False
                ElseIf OddNumQuotes Then
                    Dim PrintMsg As Boolean = False
                    If CurrLine <> "" AndAlso Not CurrLine.StartsWith(vbTab + """") AndAlso Not CurrLine.StartsWith(vbTab + "<") Then
                        PrintMsg = True
                    End If
                    If LastLine.EndsWith("""") Then
                        PrintMsg = True
                    End If
                    If PrintMsg Then
                        If Not FoundTimestamp Then
                            HeaderText.AppendLine(OddNumberOfQuotesMsg)
                        Else
                            TargetText.AppendLine(OddNumberOfQuotesMsg)
                        End If
                        HasErrorLines = True
                        If Not LastLineError Then
                            FileChanged = True
                        End If
                    End If
                    OddNumQuotes = False
                End If
                If MismatchedTag_I Then
                    If Not FoundTimestamp Then
                        HeaderText.AppendLine(MismatchedTagMsg_I)
                    Else
                        TargetText.AppendLine(MismatchedTagMsg_I)
                    End If
                    HasErrorLines = True
                    If Not LastLineError Then
                        FileChanged = True
                    End If
                    MismatchedTag_I = False
                End If
                If MismatchedTag_B Then
                    If Not FoundTimestamp Then
                        HeaderText.AppendLine(MismatchedTagMsg_B)
                    Else
                        TargetText.AppendLine(MismatchedTagMsg_B)
                    End If
                    HasErrorLines = True
                    If Not LastLineError Then
                        FileChanged = True
                    End If
                    MismatchedTag_B = False
                End If
                If MismatchedTag_U Then
                    If Not FoundTimestamp Then
                        HeaderText.AppendLine(MismatchedTagMsg_U)
                    Else
                        TargetText.AppendLine(MismatchedTagMsg_U)
                    End If
                    HasErrorLines = True
                    If Not LastLineError Then
                        FileChanged = True
                    End If
                    MismatchedTag_U = False
                End If
                If NeedsAccent Then
                    If Not FoundTimestamp Then
                        HeaderText.AppendLine(NeedsAccentMsg.Replace("{1}", UnaccentedWord).Replace("{2}", AccentedWord))
                    Else
                        TargetText.AppendLine(NeedsAccentMsg.Replace("{1}", UnaccentedWord).Replace("{2}", AccentedWord))
                    End If
                    HasErrorLines = True
                    If Not LastLineError Then
                        FileChanged = True
                    End If
                    NeedsAccent = False
                End If
                If BadFormatting Then
                    If Not FoundTimestamp Then
                        HeaderText.AppendLine(BadFormattingMsg.Replace("{1}", BadFormattingPhrase))
                    Else
                        TargetText.AppendLine(BadFormattingMsg.Replace("{1}", BadFormattingPhrase))
                    End If
                    HasErrorLines = True
                    If Not LastLineError Then
                        FileChanged = True
                    End If
                    BadFormatting = False
                End If
                ' --- Handle Spaces ---
                If CurrLine.EndsWith(" ") Then
                    CurrLine = CurrLine.TrimEnd
                End If
                If CurrLine.StartsWith(" ") AndAlso Not CurrLine.Contains(vbTab) Then
                    CurrLine = vbTab + CurrLine
                End If
                Do While CurrLine.Contains(Chr(160)) ' non-breaking space
                    CurrLine = CurrLine.Replace(Chr(160), " ")
                Loop
                Do While CurrLine.Contains("  ")
                    CurrLine = CurrLine.Replace("  ", " ")
                Loop
                Do While CurrLine.Contains(vbTab + " ")
                    CurrLine = CurrLine.Replace(vbTab + " ", vbTab)
                Loop
                Do While CurrLine.Contains(" " + vbTab)
                    CurrLine = CurrLine.Replace(" " + vbTab, vbTab)
                Loop
                ' --- Handle tabs ---
                LastChar = CChar(vbTab)
                For Each CurrChar As Char In CurrLine
                    If CurrChar = vbTab And LastChar <> vbTab Then
                        CurrLine = CurrLine.Replace(LastChar + CurrChar, LastChar + " ")
                    End If
                    LastChar = CurrChar
                Next
                Do While CurrLine.Contains("  ")
                    CurrLine = CurrLine.Replace("  ", " ")
                Loop
                ' --- Handle commas ---
                Do While CurrLine.Contains(" ,")
                    CurrLine = CurrLine.Replace(" ,", ",")
                Loop
                If CurrLine.EndsWith(",·") Then
                    CurrLine = CurrLine.Substring(0, CurrLine.Length - 1)
                End If
                ' --- Handle elipsis ---
                If CurrLine.Contains("…") Then
                    CurrLine = CurrLine.Replace("…", "...")
                End If
                If Not CurrLine.Contains(".....") Then
                    Do While CurrLine.Contains(". .")
                        CurrLine = CurrLine.Replace(". .", "...")
                    Loop
                    Do While CurrLine.Contains("....")
                        CurrLine = CurrLine.Replace("....", "...")
                    Loop
                    Do While CurrLine.Contains(" ...")
                        CurrLine = CurrLine.Replace(" ...", "...")
                    Loop
                    Do While CurrLine.Contains("... ")
                        CurrLine = CurrLine.Replace("... ", "...")
                    Loop
                    Do While CurrLine.Contains(".""...")
                        CurrLine = CurrLine.Replace(".""...", "."" ...")
                    Loop
                    Do While CurrLine.Contains(".'...")
                        CurrLine = CurrLine.Replace(".'...", ".' ...")
                    Loop
                    If CurrLine.Contains("...") Then
                        Dim TempLine As New StringBuilder
                        PeriodCount = 0
                        LastChar = " "c
                        For Each CurrChar As Char In CurrLine
                            If (Char.IsLetterOrDigit(LastChar) OrElse LastChar = "?" OrElse LastChar = "!" OrElse LastChar = ">"c OrElse LastChar = "'") AndAlso CurrChar = "."c Then
                                PeriodCount = 1
                            ElseIf LastChar = "."c AndAlso CurrChar = "."c Then
                                PeriodCount += 1
                            ElseIf LastChar = "."c AndAlso (Char.IsLetterOrDigit(CurrChar) OrElse CurrChar = "<"c OrElse Char.IsSymbol(CurrChar)) AndAlso PeriodCount = 3 Then
                                TempLine.Append(" "c)
                            Else
                                PeriodCount = 0
                            End If
                            TempLine.Append(CurrChar)
                            LastChar = CurrChar
                        Next
                        CurrLine = TempLine.ToString
                    End If
                    If CurrLine.Contains("... "" ") Then
                        CurrLine = CurrLine.Replace("... "" ", "... """)
                    End If
                    If CurrLine.EndsWith("... """) Then
                        CurrLine = CurrLine.Substring(0, CurrLine.Length - 5) + "..."""
                    End If
                    If CurrLine.EndsWith("... '") Then
                        CurrLine = CurrLine.Substring(0, CurrLine.Length - 5) + "...'"
                    End If
                    If CurrLine.Contains("...""") Then
                        Dim LastIndexOf As Integer = 0
                        Do
                            LastIndexOf = CurrLine.IndexOf("...""", LastIndexOf)
                            If LastIndexOf > 0 AndAlso LastIndexOf < CurrLine.Length - 4 Then
                                If Char.IsLetterOrDigit(CChar(CurrLine.Substring(LastIndexOf - 1, 1))) AndAlso
                                        Char.IsLetterOrDigit(CChar(CurrLine.Substring(LastIndexOf + 4, 1))) Then
                                    CurrLine = CurrLine.Insert(LastIndexOf + 3, " ")
                                End If
                            End If
                            LastIndexOf = CurrLine.IndexOf("...""", LastIndexOf + 1)
                        Loop While LastIndexOf >= 0
                    End If
                    If CurrLine.Contains("...'") Then
                        Dim LastIndexOf As Integer = 0
                        Do
                            LastIndexOf = CurrLine.IndexOf("...'", LastIndexOf)
                            If LastIndexOf > 0 AndAlso LastIndexOf < CurrLine.Length - 4 Then
                                If Char.IsLetterOrDigit(CChar(CurrLine.Substring(LastIndexOf - 1, 1))) AndAlso
                                        Char.IsLetterOrDigit(CChar(CurrLine.Substring(LastIndexOf + 4, 1))) Then
                                    CurrLine = CurrLine.Insert(LastIndexOf + 3, " ")
                                End If
                            End If
                            LastIndexOf = CurrLine.IndexOf("...'", LastIndexOf + 1)
                        Loop While LastIndexOf >= 0
                    End If
                    If CurrLine.Contains("<i>""... ") Then
                        CurrLine = CurrLine.Replace("<i>""... ", "<i>""...")
                    End If
                    If CurrLine.Contains("<i>'... ") Then
                        CurrLine = CurrLine.Replace("<i>'... ", "<i>'...")
                    End If
                    If CurrLine.Contains("""...""") Then
                        CurrLine = CurrLine.Replace("""...""", """... """)
                    End If
                    If CurrLine.Contains("'...'") Then
                        CurrLine = CurrLine.Replace("'...'", "'... '")
                    End If
                    If CurrLine.Contains(vbTab + "<i>... ") Then
                        CurrLine = CurrLine.Replace(vbTab + "<i>... ", vbTab + "<i>...")
                    End If
                    If CurrLine.Contains(" <i>... ") Then
                        CurrLine = CurrLine.Replace(" <i>... ", " <i>...")
                    End If
                    If CurrLine.Contains(vbTab + """... ") Then
                        CurrLine = CurrLine.Replace(vbTab + """... ", vbTab + """...")
                    End If
                    If CurrLine.Contains(vbTab + "'... ") Then
                        CurrLine = CurrLine.Replace(vbTab + "'... ", vbTab + "'...")
                    End If
                    If CurrLine.Contains(" '... ") Then
                        CurrLine = CurrLine.Replace(" '... ", " '...")
                    End If
                    Dim CurrIndex As Integer = 0
                    Do While CurrIndex < CurrLine.Length - 5
                        If Char.IsLetterOrDigit(CChar(CurrLine(CurrIndex))) AndAlso
                           (CurrLine(CurrIndex + 1) = """" OrElse CurrLine(CurrIndex + 1) = "'") AndAlso
                           CurrLine.Substring(CurrIndex + 2, 3) = "..." AndAlso
                           Char.IsLetterOrDigit(CChar(CurrLine(CurrIndex + 5))) Then
                            CurrLine = CurrLine.Insert(CurrIndex + 5, " ")
                        End If
                        CurrIndex += 1
                    Loop
                End If
                ' --- Handle dashes ---
                Do While CurrLine.Contains(" - ")
                    CurrLine = CurrLine.Replace(" - ", "—")
                Loop
                Do While CurrLine.Contains("–")
                    CurrLine = CurrLine.Replace("–", "—")
                Loop
                Do While CurrLine.Contains("--")
                    CurrLine = CurrLine.Replace("--", "—")
                Loop
                If CurrLine.Contains("—") Then
                    CurrLine = CurrLine.Replace("—", " — ")
                End If
                If CurrLine.Contains("- —") Then
                    CurrLine = CurrLine.Replace("- —", " — ")
                End If
                If CurrLine.Contains("— -") Then
                    CurrLine = CurrLine.Replace("— -", " — ")
                End If
                If CurrLine.Contains("""- ") Then
                    CurrLine = CurrLine.Replace("""- ", """ — ")
                End If
                If CurrLine.Contains(" -""") Then
                    CurrLine = CurrLine.Replace(" -""", " — """)
                End If
                ' '' --- Handle dash with quote ---
                ''If CurrLine.Contains(vbTab + """ — ") Then
                ''    CurrLine = CurrLine.Replace(vbTab + """ — ", vbTab + """—")
                ''End If
                ''If CurrLine.Contains(" — "" ") Then
                ''    CurrLine = CurrLine.Replace(" — "" ", "—"" ")
                ''End If
                ''If CurrLine.Contains(" "" — ") Then
                ''    CurrLine = CurrLine.Replace(" "" — ", " ""—")
                ''End If
                ''If CurrLine.EndsWith(" — """) Then
                ''    CurrLine = CurrLine.Substring(0, CurrLine.Length - 4) + "—"""
                ''End If
                ' --- Handle italics with dashes ---
                If CurrLine.Contains(Chr(160) + " —") Then
                    CurrLine = CurrLine.Replace(Chr(160) + " —", Chr(160) + "—")
                End If
                If CurrLine.Contains("^ —") Then
                    CurrLine = CurrLine.Replace("^ —", "^—")
                End If
                ' --- Handle italics ---
                FixTags(CurrLine, "i")
                ' --- Handle bold ---
                FixTags(CurrLine, "b")
                ' --- Handle underline ---
                FixTags(CurrLine, "u")
                ' --- Handle quotes ---
                If CurrLine.Contains("""'") Then
                    CurrLine = CurrLine.Replace("""'", """ '")
                End If
                If CurrLine.Contains("'""") Then
                    CurrLine = CurrLine.Replace("'""", "' """)
                End If
                QuoteCount = 0
                OddNumQuotes = False
                MismatchedQuotes = False
                LastChar = " "c
                If FoundTimestamp AndAlso CurrLine.StartsWith(vbTab) AndAlso Not CurrLine.StartsWith(vbTab + vbTab) Then
                    For Each CurrChar As Char In CurrLine
                        If CurrChar = vbTab Then
                            CurrChar = " "c
                        End If
                        If CurrChar = """"c AndAlso LastChar <> "\"c Then
                            QuoteCount += 1
                            If QuoteCount Mod 2 = 1 Then ' left quote
                                If Char.IsLetter(LastChar) OrElse
                                    LastChar = "."c OrElse
                                    LastChar = ","c OrElse
                                    LastChar = "?"c OrElse
                                    LastChar = "!"c Then
                                    MismatchedQuotes = True
                                End If
                            End If
                        ElseIf LastChar = """"c Then
                            If QuoteCount Mod 2 = 0 Then ' right quote
                                If Char.IsLetter(CurrChar) Then
                                    MismatchedQuotes = True
                                End If
                            End If
                        End If
                        If CurrChar = """"c AndAlso LastChar = "\"c Then
                            LastChar = " "c
                        Else
                            LastChar = CurrChar
                        End If
                    Next
                    If LastChar = """"c Then
                        If QuoteCount Mod 2 = 1 Then ' left quote at end of line
                            MismatchedQuotes = True
                        End If
                    End If
                    If QuoteCount Mod 2 = 1 Then
                        OddNumQuotes = True
                    End If
                End If
                If CurrLine.Contains("""<i>'") Then
                    CurrLine = CurrLine.Replace("""<i>'", """ <i>'")
                End If
                If CurrLine.Contains("'</i>""") Then
                    CurrLine = CurrLine.Replace("'</i>""", "'</i> """)
                End If
                ' --- Handle Spaces again ---
                If CurrLine.EndsWith(" ") Then
                    CurrLine = CurrLine.TrimEnd
                End If
                If CurrLine.StartsWith(" ") AndAlso Not CurrLine.Contains(vbTab) Then
                    CurrLine = vbTab + CurrLine
                End If
                Do While CurrLine.Contains("  ")
                    CurrLine = CurrLine.Replace("  ", " ")
                Loop
                Do While CurrLine.Contains(vbTab + " ")
                    CurrLine = CurrLine.Replace(vbTab + " ", vbTab)
                Loop
                Do While CurrLine.Contains(" " + vbTab)
                    CurrLine = CurrLine.Replace(" " + vbTab, vbTab)
                Loop
            End If
            ' --- Check for <image= tags ---
            If CurrLine.Contains("<image=") Then
                If CurrLine.Contains("images/") Then
                    CurrLine = CurrLine.Replace("images/", "images\")
                End If
            End If
            ' --- Check if line has changed ---
            LastLineError = False
            If String.IsNullOrWhiteSpace(CurrLine) AndAlso CurrLine <> "" Then
                CurrLine = ""
            End If
            If CurrLine <> OrigLine Then
                FileChanged = True
            End If
            If Not FoundTimestamp Then
                HeaderText.AppendLine(CurrLine)
            Else
                If CurrLine <> "" AndAlso
                    Not CurrLine.StartsWith("<") AndAlso
                    Not CurrLine.StartsWith("~") AndAlso
                    Not CurrLine.StartsWith("^") AndAlso
                    Not CurrLine.StartsWith("|") AndAlso
                    Not CurrLine.StartsWith("###") AndAlso
                    Not BlockIndent Then
                    If Not CurrLine.StartsWith(vbTab) Then
                        If Not AfterChapter Then
                            Do While NumBlankLines < MaxBlankLines
                                FileChanged = True
                                TargetText.AppendLine()
                                NumBlankLines += 1
                            Loop
                        End If
                        AfterChapter = True
                    ElseIf AfterChapter Then
                        If LastLine.StartsWith("|") Then
                            NumBlankLines = 1
                        End If
                        Do While NumBlankLines < 1
                            FileChanged = True
                            TargetText.AppendLine()
                            NumBlankLines += 1
                        Loop
                        AfterChapter = False
                    End If
                End If
                TargetText.AppendLine(CurrLine)
                If CurrLine = "" Then
                    NumBlankLines += 1
                Else
                    NumBlankLines = 0
                End If
                If CurrLine.StartsWith(vbTab + vbTab) Then
                    BlockIndent = True
                ElseIf CurrLine <> "" Then
                    BlockIndent = False
                End If
            End If
            If MismatchedTags(CurrLine, "i") Then
                MismatchedTag_I = True
            ElseIf MismatchedTags(CurrLine, "b") Then
                MismatchedTag_B = True
            ElseIf MismatchedTags(CurrLine, "u") Then
                MismatchedTag_U = True
            ElseIf NeedsAccentedWord(CurrLine) Then
                NeedsAccent = True
            ElseIf HasBadFormatting(CurrLine) Then
                BadFormatting = True
            End If
            LastLine = CurrLine
        Next
        If MismatchedQuotes Then
            If Not FoundTimestamp Then
                HeaderText.AppendLine(MismatchedQuotesMsg)
            Else
                TargetText.AppendLine(MismatchedQuotesMsg)
            End If
            HasErrorLines = True
            FileChanged = True
        ElseIf OddNumQuotes Then
            If Not FoundTimestamp Then
                HeaderText.AppendLine(OddNumberOfQuotesMsg)
            Else
                TargetText.AppendLine(OddNumberOfQuotesMsg)
            End If
            HasErrorLines = True
            FileChanged = True
        End If
        If MismatchedTag_I Then
            If Not FoundTimestamp Then
                HeaderText.AppendLine(MismatchedTagMsg_I)
            Else
                TargetText.AppendLine(MismatchedTagMsg_I)
            End If
            HasErrorLines = True
            FileChanged = True
        End If
        If MismatchedTag_B Then
            If Not FoundTimestamp Then
                HeaderText.AppendLine(MismatchedTagMsg_B)
            Else
                TargetText.AppendLine(MismatchedTagMsg_B)
            End If
            HasErrorLines = True
            FileChanged = True
        End If
        If MismatchedTag_U Then
            If Not FoundTimestamp Then
                HeaderText.AppendLine(MismatchedTagMsg_U)
            Else
                TargetText.AppendLine(MismatchedTagMsg_U)
            End If
            HasErrorLines = True
            FileChanged = True
        End If
        If FileChanged Then
            ' --- Update the timestamp if one exists ---
            If FoundTimestamp Then
                HeaderText.Append("<meta name=""timestamp"" content=""")
                HeaderText.Append(Today.ToString("yyyy-MM-dd"))
                HeaderText.AppendLine(""" />")
            End If
            File.WriteAllText(FileName, HeaderText.ToString + TargetText.ToString, Encoding.UTF8)
            TextBoxResults.AppendText(BaseFileName + vbCrLf)
            If HasErrorLines Then
                TextBoxResults.AppendText("   ### Has Errors ###" + vbCrLf)
            End If
        End If
    End Sub

    Private Function FixLine(ByVal CurrLine As String) As String
        ' --- Done ---
        Return CurrLine
    End Function

    Private Sub ExitToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ExitToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub FixTags(ByRef CurrLine As String, ByVal TagName As String)
        If Not CurrLine.Contains("<" + TagName + ">") AndAlso Not CurrLine.Contains("</" + TagName + ">") Then
            Exit Sub
        End If
        Do While CurrLine.Contains("<" + TagName + "> ")
            CurrLine = CurrLine.Replace("<" + TagName + "> ", "<" + TagName + ">")
        Loop
        Do While CurrLine.Contains(" </" + TagName + ">")
            CurrLine = CurrLine.Replace(" </" + TagName + ">", "</" + TagName + ">")
        Loop
        If CurrLine.Contains("""<" + TagName + ">— ") Then
            CurrLine = CurrLine.Replace("""<" + TagName + ">— ", """<" + TagName + "> — ")
        End If
        If CurrLine.Contains(" —</" + TagName + ">""") Then
            CurrLine = CurrLine.Replace(" —</" + TagName + ">""", " — </" + TagName + ">""")
        End If
        If CurrLine.Contains("'<" + TagName + ">— ") Then
            CurrLine = CurrLine.Replace("'<" + TagName + ">— ", "'<" + TagName + "> — ")
        End If
        If CurrLine.Contains(" —</" + TagName + ">'") Then
            CurrLine = CurrLine.Replace(" —</" + TagName + ">'", " — </" + TagName + ">'")
        End If
    End Sub

    Private Sub AboutToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AboutToolStripMenuItem.Click
        Dim TempAbout As New AboutMain
        TempAbout.ShowDialog()
        TempAbout = Nothing
    End Sub

    Private Function MismatchedTags(ByRef CurrLine As String, ByVal TagName As String) As Boolean
        Dim TagCount As Integer
        Dim TagPosStart As Integer
        Dim TagPosClose As Integer
        ' ------------------------
        If Not CurrLine.Contains("<" + TagName + ">") AndAlso Not CurrLine.Contains("</" + TagName + ">") Then
            Return False
        End If
        ' --- Adding Start and subtracting Close should equal zero ---
        TagCount = 0
        TagPosStart = CurrLine.IndexOf("<" + TagName + ">")
        Do While TagPosStart >= 0
            TagCount += 1
            TagPosStart = CurrLine.IndexOf("<" + TagName + ">", TagPosStart + 3)
        Loop
        TagPosClose = CurrLine.IndexOf("</" + TagName + ">")
        Do While TagPosClose >= 0
            TagCount -= 1
            TagPosClose = CurrLine.IndexOf("</" + TagName + ">", TagPosClose + 4)
        Loop
        If TagCount <> 0 Then
            Return True ' Not same number of Start and Close tags
        End If
        ' --- Check if tags out of order ---
        TagPosStart = CurrLine.IndexOf("<" + TagName + ">")
        TagPosClose = CurrLine.IndexOf("</" + TagName + ">")
        Do
            If TagPosClose < TagPosStart Then
                Return True ' Close before Start
            End If
            TagPosStart = CurrLine.IndexOf("<" + TagName + ">", TagPosStart + 3)
            If TagPosStart >= 0 AndAlso TagPosStart < TagPosClose Then
                Return True ' Two Starts before Close
            End If
            TagPosClose = CurrLine.IndexOf("</" + TagName + ">", TagPosClose + 4)
        Loop Until TagPosStart < 0
        ' --- Done ---
        Return False
    End Function

    Private Function NeedsAccentedWord(ByRef CurrLine As String) As Boolean
        Dim TempLowerLine As String = CurrLine.ToLower
        ' -------------------------------------------- 
        UnaccentedWord = "a la"
        AccentedWord = "à la"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "applique"
        AccentedWord = "appliqué"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "attache"
        AccentedWord = "attaché"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "blase"
        AccentedWord = "blasé"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "cafe"
        AccentedWord = "café"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "cliche"
        AccentedWord = "cliché"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "coup d'etat"
        AccentedWord = "coup d'état"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "coup de grace"
        AccentedWord = "coup de grâce"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "d'hotel"
        AccentedWord = "d'hôtel"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "deja vu"
        AccentedWord = "déjà vu"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "facade"
        AccentedWord = "façade"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "fianceé"
        AccentedWord = "fiancée"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "fiancee"
        AccentedWord = "fiancée"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "fiance"
        AccentedWord = "fiancé"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "maitre d"
        AccentedWord = "maître d"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        'UnaccentedWord = "melange"
        'AccentedWord = "mélange"
        'If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "naive"
        AccentedWord = "naïve"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "naïvete"
        AccentedWord = "naïveté"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "naiveté"
        AccentedWord = "naïveté"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "precis"
        AccentedWord = "précis"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "precís"
        AccentedWord = "précis"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "protege"
        AccentedWord = "protégé"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "protegé"
        AccentedWord = "protégé"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "protége"
        AccentedWord = "protégé"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "seance"
        AccentedWord = "séance"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "senor"
        AccentedWord = "señor"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "senora"
        AccentedWord = "señora"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "senorita"
        AccentedWord = "señorita"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "smorgasbord"
        AccentedWord = "smörgâsbord"
        If FindWordPlural(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "touche"
        AccentedWord = "touché"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        UnaccentedWord = "yucatan"
        AccentedWord = "yucatán"
        If FindWord(TempLowerLine, UnaccentedWord) Then Return True
        Return False
    End Function

    Private Function FindWord(ByRef CurrLine As String, ByVal UnaccentedWord As String) As Boolean
        ' --- CurrLine has been lowercased ---
        If CurrLine.Contains(UnaccentedWord) AndAlso Regex.IsMatch(CurrLine, ".*[^a-z]" + UnaccentedWord + "[^a-z].*") Then
            Return True
        End If
        Return False
    End Function

    Private Function FindWordPlural(ByRef CurrLine As String, ByVal UnaccentedWord As String) As Boolean
        ' --- CurrLine has been lowercased ---
        If CurrLine.Contains(UnaccentedWord) AndAlso Regex.IsMatch(CurrLine, ".*[^a-z]" + UnaccentedWord + "s*[^a-z].*") Then
            Return True
        End If
        Return False
    End Function

    Private Function HasBadFormatting(ByRef CurrLine As String) As Boolean
        Dim TempLowerLine As String = CurrLine.ToLower
        ' -------------------------------------------- 
        BadFormattingPhrase = "per cent"
        If FindWord(TempLowerLine, BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "~"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = " ,"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = ",,"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = ",..."
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "...,"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = ".<i>."
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = ".</i>."
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "</i> <i>"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "</i><i>"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "<i></i>"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "<i> </i>"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "<i><i>"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "</i></i>"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "<i>-"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "-</i>"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "-"""
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = """-"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = """ """
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        BadFormattingPhrase = "' '"
        If CurrLine.Contains(BadFormattingPhrase) Then Return True
        Return False
    End Function

    Private Sub ButtonStop_Click(sender As Object, e As EventArgs) Handles ButtonStop.Click
        StopRequested = True
    End Sub

End Class
