Imports System.IO
Imports System
Imports System.Drawing
Imports System.Text
Imports System.Text.RegularExpressions

Public Class MainForm



    Function lastIndex(ByVal fileContent As String, ByVal search As String) As Integer      'UDF to get last index of a string inside another string, although not required since used library function anyway, just a different function signature
        Dim lastIndexFound As Integer
        lastIndexFound = fileContent.LastIndexOf(search)
        Return lastIndexFound
    End Function
    Function getAccession(ByVal content As String) As String
        Dim startIndex, endIndex As Integer
        Dim result As String
        result = "NULL"
        startIndex = content.IndexOf("ACCESSION")
        If startIndex <> -1 Then
            endIndex = content.IndexOf(vbLf, startIndex + 1)
            result = content.Substring(startIndex, endIndex - startIndex).Replace("ACCESSION", "").Replace("""", "").Replace(" ", "")
        End If
        Return result
    End Function

    Function charCount(ByVal s1 As String, ByVal ch As Char)
        Return s1.Split(ch).Length - 1
    End Function

    Function hostFilter(ByVal search() As String, content As String) As Integer
        Dim i, required As Integer
        required = 0
        For i = 0 To search.Length - 1
            If (content.IndexOf(search(i)) <> -1) Then
                required = 1
            End If
        Next
        Return required
    End Function
    Function organismFilter(ByVal organism As String, content As String) As Integer
        Dim required As Integer

        'organism = "/organism=" & """" & organism & """"
        organism = "/organism=" & """" & organism           'changed from above because neuraminidase doesn't end on that line
        required = 0
        If (content.ToLower.Contains(organism.ToLower)) Then
            required = 1
        End If

        Return required
    End Function
    Private Function FindWords(ByVal searchString As String, ByVal fullString As String) As Integer

        Dim count As Integer
        Dim currIndex, i As Integer
        count = 0
        currIndex = 0
        i = 0

        Do
            currIndex = fullString.IndexOf(searchString, i)
            If (currIndex <> -1) Then
                count = count + 1
            Else
                Exit Do
            End If
            i = currIndex + searchString.Length             'was i = currIndex + searchString.Length - 2, changed to i = currIndex + searchString.Length
        Loop While (True)

        Return count

    End Function
    Function subTypeFilter(ByVal searchString As String, ByVal searchLine As String, ByVal fullString As String)
        Dim required, searchStartIndex, searchEndIndex As Integer


        required = 0
        searchStartIndex = fullString.IndexOf(searchLine)
        searchEndIndex = fullString.IndexOf(vbLf, searchStartIndex + 2)

        If (fullString.Substring(searchStartIndex, searchEndIndex - searchStartIndex + 1).ToLower.Contains(searchString.ToLower) = True) Then
            required = 1
        End If



        Return required
    End Function

    Function ifExists(ByVal search As String, ByVal fileContent As String) As String
        Dim found, required As Integer

        required = 0

        found = fileContent.IndexOf(search)
        If found <> -1 Then
            required = 1
        End If
        Return required
    End Function
    Function getDetails(ByVal search As String, ByVal fileContent As String) As String
        Dim startIndex, endIndex As Integer
        Dim result As String
        result = "Not Found"
        startIndex = fileContent.IndexOf(search)
        If startIndex <> -1 Then
            endIndex = fileContent.IndexOf(vbLf, startIndex + 1)
            result = fileContent.Substring(startIndex, endIndex - startIndex).Replace(search, "").Replace("""", "")
        End If
        Return result
    End Function
    Function convertSequence(ByVal Sequence As String) As String                        'generating new sequence which will contain only bases
        Dim Seq2, ch As String
        Dim ln, i As Integer
        ln = Sequence.Length
        Seq2 = ""
        For i = 0 To ln - 1
            ch = Sequence.Substring(i, 1)
            'If ch = "a" Or ch = "A" Or ch = "c" Or ch = "C" Or ch = "t" Or ch = "T" Or ch = "g" Or ch = "G" Or ch = "n" Or ch = "N" Then
            If ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z")) Then            'removing all non alphabet characters from a string
                Seq2 = Seq2 + ch
            End If
        Next
        Return Seq2
    End Function
    Function baseExtractor(ByVal Sequence As String, ByVal startIndex As Integer, ByVal endIndex As Integer)
        Dim res As String
        res = Sequence.Substring(startIndex - 1, endIndex - startIndex + 1)
        res = res.ToUpper
        Return res
    End Function



    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        'Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        'fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.Filter = "GenBank Files|*.gb"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            'strFileName = fd.FileName
            GV.tab1InpFile = fd.FileName
        End If
        TextBox1.ReadOnly = False
        TextBox1.Text = GV.tab1InpFile
        TextBox1.ReadOnly = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.tab1OutFolder = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox2.Text = GV.tab1OutFolder                                    'showing path in textbox for user's ease
            TextBox2.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        GV.tab1InpFile = TextBox1.Text
        GV.tab1OutFolder = TextBox2.Text

        If (GV.tab1InpFile = "" Or GV.tab1OutFolder = "" Or TextBox3.Text = "") Then
            MsgBox("Enter the Input File, the Output folder Path and the Organism fields and try again.")
            Return
        End If


        If ((File.Exists(GV.tab1InpFile)) = False Or (Directory.Exists(GV.tab1OutFolder) = False)) Then
            MsgBox("Invalid input file/output folder entered. Enter valid file/folder and try again." & vbNewLine & "Or select file/folder through the Browse File/Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If


        Dim fileContent, singleFileContent, accession, organism, CDS, Sequence, SequenceConverted As String
        Dim startIndex, endIndex, count, required, lastIndexOrigin, lastIndexDoubleSlash As Integer
        Dim key As String
        Dim keyIndex, geneCdsIndex, accessionIndex, versionIndex As Integer
        Dim fastaHeader, fastaContent, host, country, colDate, organismInFile As String
        fastaContent = ""

        organism = TextBox3.Text
        fileContent = ""

        If (System.IO.File.Exists(GV.tab1InpFile)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
            fileContent = File.ReadAllText(GV.tab1InpFile)
        End If

        fileContent = fileContent.Replace(vbCrLf, vbLf)



        'startIndex = fileContent.IndexOf("LOCUS", 3)
        'test = fileContent.Substring(startIndex - 4, 10)
        'MsgBox(test)

        startIndex = 0
        count = 1
        required = 0

        startIndex = fileContent.IndexOf("LOCUS", startIndex)
        Do
            If (startIndex >= fileContent.Length - 1 Or startIndex = -1) Then
                Exit Do
            End If
            startIndex = fileContent.IndexOf("LOCUS", startIndex)
            endIndex = fileContent.IndexOf(vbLf & "//", startIndex) + 3


            If Not Directory.Exists(GV.tab1OutFolder & "\Required") Then
                Directory.CreateDirectory(GV.tab1OutFolder & "\Required")
            End If
            If Not Directory.Exists(GV.tab1OutFolder & "\Skipped") Then
                Directory.CreateDirectory(GV.tab1OutFolder & "\Skipped")
            End If


            singleFileContent = fileContent.Substring(startIndex, endIndex - startIndex)


            accessionIndex = singleFileContent.IndexOf("ACCESSION")
            versionIndex = singleFileContent.IndexOf("VERSION")

            accession = singleFileContent.Substring(accessionIndex + "ACCESSION".Length, versionIndex - accessionIndex - "ACCESSION".Length - 1)
            accession = accession.Trim
            accession = accession.Replace(vbCr, "").Replace(vbLf, "")
            host = getDetails("/host=", singleFileContent)
            country = getDetails("/country=", singleFileContent)
            colDate = getDetails("/collection_date=", singleFileContent)
            organismInFile = getDetails("/organism=", singleFileContent)

            If (TextBox4.Text <> "") Then
                fastaHeader = ">ACCESSION=" & accession & " HOST=" & host & " COUNTRY=" & country & " Collection Date=" & colDate & " ORGANISM=" & organismInFile & " GENE=" & TextBox4.Text.ToString
            Else
                fastaHeader = ">ACCESSION=" & accession & " HOST=" & host & " COUNTRY=" & country & " Collection Date=" & colDate & " ORGANISM=" & organismInFile
            End If



            required = organismFilter(organism, singleFileContent)
            If (required = 0) Then
                GoTo JUMP
            End If



            lastIndexOrigin = lastIndex(singleFileContent, "ORIGIN")              'store last index/occurrence of ORIGIN String
            lastIndexDoubleSlash = lastIndex(singleFileContent, "//")             'store last index/occurrence of '//' String
            Sequence = singleFileContent.Substring(lastIndexOrigin + 6, lastIndexDoubleSlash - lastIndexOrigin - 7)       'taking the part of the txt file which contains only the sequence (but with line nos and spaces)
            Sequence = Sequence.Trim        'removing spaces from start and end, but doesnt work on spaces in between
            SequenceConverted = convertSequence(Sequence)       'UDF for removing line nos. and spaces, i.e. takes into account only letters be it upper case or lower


            If (TextBox4.Text <> "") Then           'doing for single gene, else part is for whole genome
                key = "/gene=" & """" & TextBox4.Text & """"

                keyIndex = singleFileContent.LastIndexOf(key)
                If keyIndex = -1 Then
                    GoTo JUMP2
                End If
                geneCdsIndex = singleFileContent.Substring(0, keyIndex - 2).LastIndexOf("CDS ")


                CDS = singleFileContent.Substring(geneCdsIndex + 16, singleFileContent.IndexOf(vbLf, geneCdsIndex + 2) - geneCdsIndex - 16)  'storing the range of CDS
                CDS = CDS.Trim      'removing spaces from start and end

                If (CDS.IndexOf("<") <> -1 Or CDS.IndexOf(">") <> -1) Then      'checking for incomplete sequences and skipping them
                    required = 0
                    GoTo JUMP
                End If

                Dim countCDSes, loopStartIndex, loopEndIndex, loopDotIndex, CDS_Dot_End_Index As Integer
                countCDSes = FindWords(",", CDS) + 1
                Dim gene_CDS_Index(countCDSes - 1, 1) As Integer

                If countCDSes = 1 Then              'i.e. not containing join
                    CDS_Dot_End_Index = lastIndex(CDS, ".")                 'store last occurrence of '.' in CDS range
                    gene_CDS_Index(0, 0) = CInt(CDS.Substring(0, CDS_Dot_End_Index - 1))       'store CDS start as integer
                    gene_CDS_Index(0, 1) = CInt(CDS.Substring(CDS_Dot_End_Index + 1, CDS.Length - CDS_Dot_End_Index - 1))        'store CDS end as integer
                Else                                'i.e. containing join
                    Dim CDS_Start, CDS_End, partCDS As String

                    loopStartIndex = CDS.IndexOf("(")

                    For i = 0 To gene_CDS_Index.GetLength(0) - 1
                        loopEndIndex = CDS.IndexOf(",", loopStartIndex + 1)
                        If loopEndIndex = -1 Then
                            loopEndIndex = CDS.IndexOf(")", loopStartIndex + 1)
                        End If
                        partCDS = CDS.Substring(loopStartIndex + 1, loopEndIndex - 1 - loopStartIndex)
                        loopDotIndex = partCDS.LastIndexOf(".")
                        CDS_Start = partCDS.Substring(0, loopDotIndex - 1)
                        CDS_End = partCDS.Substring(loopDotIndex + 1, partCDS.Length - loopDotIndex - 1)
                        gene_CDS_Index(i, 0) = CInt(CDS_Start)
                        gene_CDS_Index(i, 1) = CInt(CDS_End)
                        loopStartIndex = loopEndIndex
                        'If loopStartIndex >= gene_CDS_Index.GetLength(0) - 1 Then
                        '    Exit For
                        'End If
                    Next
                End If






                fastaContent = ""
                For i = 0 To gene_CDS_Index.GetLength(0) - 1
                    fastaContent = fastaContent & baseExtractor(SequenceConverted, gene_CDS_Index(i, 0), gene_CDS_Index(i, 1))
                Next


            Else        'i.e. if gene textbox is blank, then do for whole genome
                Dim firstCDSOccurrenceIndex, lastCDSOccurrenceIndex, dotIndex, firstBracketIndex, WG_Start_Index, WG_End_Index As Integer
                Dim firstCDSpart, lastCDSpart As String

                firstCDSOccurrenceIndex = singleFileContent.IndexOf("CDS ")
                lastCDSOccurrenceIndex = singleFileContent.LastIndexOf("CDS ")

                firstCDSpart = singleFileContent.Substring(firstCDSOccurrenceIndex + 16, singleFileContent.IndexOf(vbLf, firstCDSOccurrenceIndex + 2) - firstCDSOccurrenceIndex - 16)
                firstCDSpart.Trim()
                lastCDSpart = singleFileContent.Substring(lastCDSOccurrenceIndex + 16, singleFileContent.IndexOf(vbLf, lastCDSOccurrenceIndex + 2) - lastCDSOccurrenceIndex - 16)
                lastCDSpart.Trim()

                dotIndex = firstCDSpart.IndexOf(".")
                firstBracketIndex = firstCDSpart.IndexOf("(")
                If firstBracketIndex = -1 Then
                    WG_Start_Index = CInt(firstCDSpart.Substring(0, dotIndex))
                Else
                    WG_Start_Index = CInt(firstCDSpart.Substring(firstBracketIndex + 1, dotIndex - firstBracketIndex - 1))
                End If

                dotIndex = lastCDSpart.LastIndexOf(".")
                firstBracketIndex = lastCDSpart.LastIndexOf(")")
                If firstBracketIndex = -1 Then
                    WG_End_Index = CInt(lastCDSpart.Substring(dotIndex + 1, lastCDSpart.Length - dotIndex - 1))
                Else
                    WG_End_Index = CInt(lastCDSpart.Substring(dotIndex + 1, firstBracketIndex - dotIndex - 1))
                End If

                fastaContent = ""
                fastaContent = fastaContent & baseExtractor(SequenceConverted, WG_Start_Index, WG_End_Index)
            End If


JUMP:
            accession = getAccession(fileContent.Substring(startIndex, endIndex - startIndex))

            If (required = 1) Then
                If (accession <> "NULL") Then
                    File.WriteAllText(GV.tab1OutFolder & "\Required\" & accession & ".fasta", fastaHeader & vbLf & fastaContent)
                Else
                    File.WriteAllText(GV.tab1OutFolder & "\Required\" & count & ".fasta", fastaHeader & vbLf & fastaContent)
                End If
            Else
                If (accession <> "NULL") Then
                    File.WriteAllText(GV.tab1OutFolder & "\Skipped\" & accession & ".fasta", fastaHeader & vbLf & fastaContent)
                Else
                    File.WriteAllText(GV.tab1OutFolder & "\Skipped\" & count & ".fasta", fastaHeader & vbLf & fastaContent)
                End If
            End If


JUMP2:
            count = count + 1

            startIndex = endIndex + 2

        Loop While (True)

        MsgBox("DONE")
        'File.AppendAllText(GV.tab1OutFolder & "\Copied.txt", fileContent)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        'Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        'fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.Filter = "GenBank Files|*.gb"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            'strFileName = fd.FileName
            GV.tab2InpFile = fd.FileName
        End If
        TextBox5.ReadOnly = False
        TextBox5.Text = GV.tab2InpFile
        TextBox5.ReadOnly = True
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.tab2OutFolder = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox6.Text = GV.tab2OutFolder                                    'showing path in textbox for user's ease
            TextBox6.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        GV.tab2InpFile = TextBox5.Text
        GV.tab2OutFolder = TextBox6.Text

        If (GV.tab2InpFile = "" Or GV.tab2OutFolder = "" Or TextBox7.Text = "" Or TextBox8.Text = "") Then
            MsgBox("Enter the Input File, the Output folder Path, the Organism and the Gene fields and try again.")
            Return
        End If


        If ((File.Exists(GV.tab2InpFile)) = False Or (Directory.Exists(GV.tab2OutFolder) = False)) Then
            MsgBox("Invalid input file/output folder entered. Enter valid file/folder and try again." & vbNewLine & "Or select file/folder through the Browse File/Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If


        Dim fileContent, singleFileContent, accession, organism, translationString As String
        Dim startIndex, endIndex, count, required As Integer
        Dim key As String
        Dim keyIndex, geneCdsIndex, nextGeneCdsIndex, translationIndex, translationEndIndex, accessionIndex, versionIndex As Integer
        Dim fastaHeader, fastaContent, host, country, colDate, organismInFile As String
        fastaContent = ""

        organism = TextBox7.Text
        fileContent = ""

        If (System.IO.File.Exists(GV.tab2InpFile)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
            fileContent = File.ReadAllText(GV.tab2InpFile)
        End If

        fileContent = fileContent.Replace(vbCrLf, vbLf)



        'startIndex = fileContent.IndexOf("LOCUS", 3)
        'test = fileContent.Substring(startIndex - 4, 10)
        'MsgBox(test)

        startIndex = 0
        count = 1
        required = 0

        startIndex = fileContent.IndexOf("LOCUS", startIndex)
        Do
            If (startIndex >= fileContent.Length - 1 Or startIndex = -1) Then
                Exit Do
            End If
            startIndex = fileContent.IndexOf("LOCUS", startIndex)
            endIndex = fileContent.IndexOf(vbLf & "//", startIndex) + 3


            If Not Directory.Exists(GV.tab2OutFolder & "\Required") Then
                Directory.CreateDirectory(GV.tab2OutFolder & "\Required")
            End If
            If Not Directory.Exists(GV.tab2OutFolder & "\Skipped") Then
                Directory.CreateDirectory(GV.tab2OutFolder & "\Skipped")
            End If


            singleFileContent = fileContent.Substring(startIndex, endIndex - startIndex)


            accessionIndex = singleFileContent.IndexOf("ACCESSION")
            versionIndex = singleFileContent.IndexOf("VERSION")

            accession = singleFileContent.Substring(accessionIndex + "ACCESSION".Length, versionIndex - accessionIndex - "ACCESSION".Length - 1)
            accession = accession.Trim
            accession = accession.Replace(vbCr, "").Replace(vbLf, "")
            host = getDetails("/host=", singleFileContent)
            country = getDetails("/country=", singleFileContent)
            colDate = getDetails("/collection_date=", singleFileContent)
            organismInFile = getDetails("/organism=", singleFileContent)


            fastaHeader = ">ACCESSION=" & accession & " HOST=" & host & " COUNTRY=" & country & " Collection Date=" & colDate & " ORGANISM=" & organismInFile & " GENE=" & TextBox8.Text.ToString




            required = organismFilter(organism, singleFileContent)
            If (required = 0) Then
                GoTo JUMP
            End If







            key = "/gene=" & """" & TextBox8.Text & """"

            keyIndex = singleFileContent.LastIndexOf(key)
            If keyIndex = -1 Then
                GoTo JUMP2
            End If

            geneCdsIndex = singleFileContent.Substring(0, keyIndex - 2).LastIndexOf("CDS ")

            nextGeneCdsIndex = singleFileContent.IndexOf("CDS ", geneCdsIndex + 3)

            If nextGeneCdsIndex = -1 Then
                nextGeneCdsIndex = singleFileContent.IndexOf("ORIGIN ", geneCdsIndex + 3)
            End If

            translationString = singleFileContent.Substring(geneCdsIndex, nextGeneCdsIndex - geneCdsIndex).Replace(vbLf, "").Replace(" ", "")

            translationIndex = translationString.IndexOf("/translation=")
            If translationIndex = -1 Then
                required = 0
                GoTo JUMP
            Else

                translationEndIndex = translationString.IndexOf("""", translationIndex + 15)
                fastaContent = translationString.Substring(translationIndex, translationEndIndex - translationIndex)
                fastaContent = fastaContent.Replace("/translation=", "").Replace("""", "")
            End If











JUMP:
            accession = getAccession(fileContent.Substring(startIndex, endIndex - startIndex))

            If (required = 1) Then
                If (accession <> "NULL") Then
                    File.WriteAllText(GV.tab2OutFolder & "\Required\" & accession & ".fasta", fastaHeader & vbLf & fastaContent)
                Else
                    File.WriteAllText(GV.tab2OutFolder & "\Required\" & count & ".fasta", fastaHeader & vbLf & fastaContent)
                End If
            Else
                If (accession <> "NULL") Then
                    File.WriteAllText(GV.tab2OutFolder & "\Skipped\" & accession & ".fasta", fastaHeader & vbLf & fastaContent)
                Else
                    File.WriteAllText(GV.tab2OutFolder & "\Skipped\" & count & ".fasta", fastaHeader & vbLf & fastaContent)
                End If
            End If

JUMP2:

            count = count + 1

            startIndex = endIndex + 2

        Loop While (True)

        MsgBox("DONE")
        'File.AppendAllText(GV.tab1OutFolder & "\Copied.txt", fileContent)
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        'Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        'fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.Filter = "Fasta Files|*.fasta"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            'strFileName = fd.FileName
            GV.tab3InpFile = fd.FileName
        End If
        TextBox9.ReadOnly = False
        TextBox9.Text = GV.tab3InpFile
        TextBox9.ReadOnly = True
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click

        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.tab3OutFolder = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox10.Text = GV.tab3OutFolder                                    'showing path in textbox for user's ease
            TextBox10.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        GV.tab3InpFile = TextBox9.Text
        GV.tab3OutFolder = TextBox10.Text

        If (GV.tab3InpFile = "" Or GV.tab3OutFolder = "" Or TextBox10.Text = "") Then
            MsgBox("Enter the Input File, the Output folder Path and the Organism fields and try again.")
            Return
        End If


        If ((File.Exists(GV.tab3InpFile)) = False Or (Directory.Exists(GV.tab3OutFolder) = False)) Then
            MsgBox("Invalid input file/output folder entered. Enter valid file/folder and try again." & vbNewLine & "Or select file/folder through the Browse File/Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If


        Dim fileContent, singleFileContent, accession, organism, fastaHeader, lengthFolder As String
        Dim startIndex, endIndex, count, required, folder, noOfFiles, currFile As Integer
        Dim overwrite As Boolean

        organism = TextBox11.Text
        fileContent = ""

        If (System.IO.File.Exists(GV.tab3InpFile)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
            fileContent = File.ReadAllText(GV.tab3InpFile)
        End If

        fileContent = fileContent.Replace(vbCrLf, vbLf)

        If (CheckBox1.Checked = True) Then
            folder = 1
        Else
            folder = 0
        End If

        If (CheckBox2.Checked = True) Then
            overwrite = False
        Else
            overwrite = True
        End If

        'startIndex = fileContent.IndexOf("LOCUS", 3)
        'test = fileContent.Substring(startIndex - 4, 10)
        'MsgBox(test)

        startIndex = 0
        count = 1
        required = 0
        fileContent = fileContent.Replace(vbLf & vbLf, vbLf)
        startIndex = fileContent.IndexOf(">", startIndex)

        noOfFiles = charCount(fileContent, ">")
        currFile = 1

        Do
            If currFile Mod 100 = 0 Then
                Console.WriteLine("Working on file no. : [" & currFile & "] out of [" & noOfFiles & "] files.")
            End If
            currFile = currFile + 1
            If (startIndex >= fileContent.Length - 1 Or startIndex = -1) Then
                Exit Do
            End If
            startIndex = fileContent.IndexOf(">", startIndex)
            endIndex = fileContent.IndexOf(vbLf & ">", startIndex + 1)

            If (endIndex = -1) Then
                endIndex = fileContent.Length - 1
            End If

            If Not Directory.Exists(GV.tab3OutFolder & "\Required") Then
                Directory.CreateDirectory(GV.tab3OutFolder & "\Required")
            End If
            If Not Directory.Exists(GV.tab3OutFolder & "\Skipped") Then
                Directory.CreateDirectory(GV.tab3OutFolder & "\Skipped")
            End If


            singleFileContent = fileContent.Substring(startIndex, endIndex - startIndex)
            fastaHeader = singleFileContent.Substring(0, singleFileContent.IndexOf(vbLf))
            accession = fastaHeader.Substring(1, fastaHeader.IndexOf(" ") - 1).Replace("|", "-")

            If fastaHeader.ToLower.Contains(organism.ToLower) Then
                required = 1
            Else
                required = 0
                GoTo JUMP
            End If

            accession = fastaHeader.Substring(1, fastaHeader.IndexOf(" ") - 1)

            'added for virus variation db fasta files       eg. gb|KP739423|1-1410      for Influenza
            If accession.Contains("|") = True Then
                accession = accession.Substring(accession.IndexOf("|") + 1, (accession.IndexOf(":") - accession.IndexOf("|") - 1))
            ElseIf accession.Contains(":") = True Then      'for dengue 
                accession = accession.Substring(0, accession.IndexOf(":"))
            End If




            lengthFolder = singleFileContent.Substring(fastaHeader.Length, singleFileContent.Length - fastaHeader.Length).Replace(vbLf, "").Length.ToString



JUMP:

            lengthFolder = singleFileContent.Substring(fastaHeader.Length, singleFileContent.Length - fastaHeader.Length).Replace(vbLf, "").Length.ToString
            'pasted from above

            If (required = 1) Then
                If (folder = 1) Then
                    If Not Directory.Exists(GV.tab3OutFolder & "\Required\" & lengthFolder) Then         'create a new directory in output path with name as total no. of bases if it doesn't exist.
                        Directory.CreateDirectory(GV.tab3OutFolder & "\Required\" & lengthFolder)
                    End If

                    File.WriteAllText(GV.tab3OutFolder & "\Required\" & lengthFolder & "\" & accession & ".fasta", singleFileContent)
                Else
                    If (overwrite = True) Then
                        File.WriteAllText(GV.tab3OutFolder & "\Required\" & accession & ".fasta", singleFileContent)
                    ElseIf (overwrite = False) Then
                        If (File.Exists(GV.tab3OutFolder & "\Required\" & accession & ".fasta")) Then
                            File.WriteAllText(GV.tab3OutFolder & "\Required\" & accession & "-2.fasta", singleFileContent)
                        Else
                            File.WriteAllText(GV.tab3OutFolder & "\Required\" & accession & ".fasta", singleFileContent)
                        End If
                    End If
                End If
            Else
                File.WriteAllText(GV.tab3OutFolder & "\Skipped\" & accession & ".fasta", singleFileContent)
            End If

            'count = count + 1

            startIndex = endIndex + 1

        Loop While (True)

        MsgBox("DONE")
    End Sub
End Class
Public Class GV         'Global Variables, other classes to written after Main Class else code wont work

    Public Shared tab1InpFile, tab1OutFolder As String
    Public Shared tab2InpFile, tab2OutFolder As String
    Public Shared tab3InpFile, tab3OutFolder As String

End Class
