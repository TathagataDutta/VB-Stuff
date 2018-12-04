Imports System.IO
Imports System
Imports System.Drawing
Imports System.Text
Imports System.Text.RegularExpressions

Public Class MainForm
    Private Sub FolderBrowserDialog1_HelpRequest(sender As Object, e As EventArgs) Handles FolderBrowserDialog1.HelpRequest

    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk

    End Sub
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
            i = currIndex + searchString.Length - 2
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
            GV.inputGBFile = fd.FileName
        End If
        TextBox1.ReadOnly = False
        TextBox1.Text = GV.inputGBFile
        TextBox1.ReadOnly = True
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.outputFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox2.Text = GV.outputFolderPath                                    'showing path in textbox for user's ease
            TextBox2.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click

        GV.inputGBFile = TextBox1.Text
        GV.outputFolderPath = TextBox2.Text

        If (GV.inputGBFile = "" Or GV.outputFolderPath = "" Or TextBox3.Text = "") Then
            MsgBox("Enter the Input File, the Output folder Path and the Organism fields and try again.")
            Return
        End If


        If ((File.Exists(GV.inputGBFile)) = False Or (Directory.Exists(GV.outputFolderPath) = False)) Then
            MsgBox("Invalid input file/output folder entered. Enter valid file/folder and try again." & vbNewLine & "Or select file/folder through the Browse File/Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If

        Label2.Visible = True

        Dim filterHost() As String
        Dim fileContent, singleFileContent, accession, organism, CDS As String
        Dim startIndex, endIndex, count, required, annoFlag, occ, lastIndexCDS, CDS_End_Index, CDS_Start, CDS_End As Integer

        organism = TextBox3.Text
        fileContent = ""

        If (System.IO.File.Exists(GV.inputGBFile)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
            fileContent = File.ReadAllText(GV.inputGBFile)
        End If

        fileContent = fileContent.Replace(vbCrLf, vbLf)

        filterHost = RichTextBox1.Lines

        'startIndex = fileContent.IndexOf("LOCUS", 3)
        'test = fileContent.Substring(startIndex - 4, 10)
        'MsgBox(test)

        startIndex = 0
        count = 1
        required = 0

        Do
            startIndex = fileContent.IndexOf("LOCUS", startIndex)

            If (startIndex >= fileContent.Length - 1 Or startIndex = -1) Then
                Exit Do
            End If

            endIndex = fileContent.IndexOf(vbLf & "//", startIndex) + 3


            If Not Directory.Exists(GV.outputFolderPath & "\Required") Then
                Directory.CreateDirectory(GV.outputFolderPath & "\Required")
            End If
            If Not Directory.Exists(GV.outputFolderPath & "\Skipped") Then
                Directory.CreateDirectory(GV.outputFolderPath & "\Skipped")
            End If


            singleFileContent = fileContent.Substring(startIndex, endIndex - startIndex)

            If (filterHost.Length <> 0) Then
                required = hostFilter(filterHost, singleFileContent)
                If (required = 0) Then
                    GoTo JUMP
                End If
            End If


            required = organismFilter(organism, singleFileContent)
            If (required = 0) Then
                GoTo JUMP
            End If

            If (CheckBox1.Checked = True) Then
                required = ifExists("/country=", singleFileContent)
                If (required = 0) Then
                    GoTo JUMP
                End If
            End If

            If (CheckBox2.Checked = True) Then
                required = ifExists("/collection_date=", singleFileContent)
                If (required = 0) Then
                    GoTo JUMP
                End If
            End If

            If (CheckBox3.Checked = True) Then

                'rejecting files having gaps
                If (ifExists(" gap ", singleFileContent) = 1) Then
                    required = 0
                    GoTo JUMP
                End If

                'checking if annotations or CDS exists or not ... below
                annoFlag = ifExists(" mat_peptide ", singleFileContent) + ifExists(" misc_feature ", singleFileContent)
                If (annoFlag >= 1) Then
                    required = 1
                Else
                    required = 0
                    GoTo JUMP
                End If

                If (ifExists(" CDS ", singleFileContent) = 1) Then
                    required = 1
                Else
                    required = 0
                    GoTo JUMP
                End If


                'checking if multiple organism (eg. chimera)
                occ = FindWords("/organism=", singleFileContent)
                If (occ > 1) Then
                    required = 0
                    GoTo JUMP
                Else
                    required = 1
                End If


                occ = FindWords("misc_feature", singleFileContent) + FindWords("mat_peptide", singleFileContent)
                If (occ >= 11) Then
                    required = 1

                Else
                    required = 0
                    GoTo JUMP
                End If

            End If


            If (TextBox4.Text <> "") Then

                If (ifExists(" CDS ", singleFileContent) = 1) Then
                    required = 1
                Else
                    required = 0
                    GoTo JUMP
                End If

                lastIndexCDS = lastIndex(singleFileContent, "CDS ")    'store last index/occurrence of CDS String
                CDS = singleFileContent.Substring(lastIndexCDS + 16, 14)  'storing the range of CDS
                CDS = CDS.Trim      'removing spaces from start and end

                If (CDS.IndexOf("<") <> -1 Or CDS.IndexOf(">") <> -1) Then      'checking for incomplete sequences and skipping them
                    required = 0
                    GoTo JUMP
                End If

                CDS_End_Index = lastIndex(CDS, ".")                 'store last occurrence of '.' in CDS range
                CDS_Start = CInt(CDS.Substring(0, CDS_End_Index - 1))       'store CDS start as integer
                CDS_End = CInt(CDS.Substring(CDS_End_Index + 1, CDS.Length - CDS_End_Index - 1))        'store CDS end as integer

                If ((CDS_End - CDS_Start + 1) = CInt(TextBox4.Text)) Then
                    required = 1
                Else
                    required = 0
                End If

                If (required = 0) Then
                    GoTo JUMP
                End If

            End If

            If (TextBox5.Text <> "") Then
                required = subTypeFilter(TextBox5.Text, "ORGANISM", singleFileContent)
                If (required = 0) Then
                    GoTo JUMP
                End If
            End If

JUMP:

            accession = getAccession(fileContent.Substring(startIndex, endIndex - startIndex))

            If (required = 1) Then
                If (accession <> "NULL") Then
                    File.WriteAllText(GV.outputFolderPath & "\Required\" & accession & ".txt", singleFileContent)
                Else
                    File.WriteAllText(GV.outputFolderPath & "\Required\" & count & ".txt", singleFileContent)
                End If
            Else
                If (accession <> "NULL") Then
                    File.WriteAllText(GV.outputFolderPath & "\Skipped\" & accession & ".txt", singleFileContent)
                Else
                    File.WriteAllText(GV.outputFolderPath & "\Skipped\" & count & ".txt", singleFileContent)
                End If
            End If

            count = count + 1

            startIndex = endIndex + 2

        Loop While (True)
        Label2.Visible = False
        MsgBox("DONE")
        'File.AppendAllText(GV.outputFolderPath & "\Copied.txt", fileContent)

    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click

    End Sub
End Class

Public Class GV         'Global Variables, other classes to written after Main Class else code wont work

    Public Shared inputGBFile, outputFolderPath As String

End Class
