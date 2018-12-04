Imports Excel = Microsoft.Office.Interop.Excel
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

    Public Sub End_Excel_App(datestart As Date, dateEnd As Date)                           'UDF to kill all excel processes generated during execution of code to prevent memory leak
        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In xlp
            If Process.StartTime >= datestart And Process.StartTime <= dateEnd Then
                Process.Kill()
                'Exit For
            End If
        Next
    End Sub
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

    Function smallestStringArrayIndex(ByVal strArr() As String) As Integer
        Dim smallestLength, smallestIndex, i As Integer

        smallestLength = strArr(0).Length
        smallestIndex = 0
        For i = 1 To strArr.Length - 1
            If strArr(i).Length < smallestLength Then
                smallestLength = strArr(i).Length
                smallestIndex = i
            End If
        Next
        Return smallestIndex
    End Function
    Function subTypeFinder(ByVal searchLine As String, ByVal fullString As String)
        Dim subType, searchString As String
        Dim startIndex, endIndex, currIndex, subStartIndex, subEndIndex As Integer

        subType = "Error"

        startIndex = fullString.IndexOf(searchLine)
        endIndex = fullString.IndexOf(vbLf, startIndex + 2)

        searchString = fullString.Substring(startIndex, endIndex - startIndex)


        startIndex = 0
        currIndex = searchString.IndexOf("H", startIndex)
        '/organism="Influenza H A virus (A/NanChang/08/2010(H1N1))"

        While (currIndex <= searchString.Length - 1 Or currIndex <> -1)
            If ((searchString.Substring(currIndex + 2, 1) = "N" Or searchString.Substring(currIndex + 3, 1) = "N") And (searchString.Substring(currIndex + 1, 1) > "0" And searchString.Substring(currIndex + 1, 1) < "9")) Then
                subStartIndex = currIndex
                subEndIndex = searchString.IndexOf(")", currIndex + 1)
                Exit While
            Else
                startIndex = currIndex + 1
                currIndex = searchString.IndexOf("H", startIndex)
            End If
        End While

        subType = searchString.Substring(subStartIndex, subEndIndex - subStartIndex)

        Return subType
    End Function


    Public Sub graph_Gen(ByVal C As Integer, ByVal Seq As String, ByVal N As Integer, ByVal CDS_Start As Integer, ByVal CDS_End As Integer, ByVal outputPath As String, Optional ByVal type As Integer = -1)
        Dim X, Y, i As Integer
        Dim SumX, SumY, SumA, SumC, SumG, SumT As Long
        Dim ch, txtPath As String
        Dim skip As Boolean
        skip = False

        X = 0
        Y = 0
        GV.CountN = 0
        SumX = 0
        SumY = 0
        SumA = 0
        SumC = 0
        SumG = 0
        SumT = 0
        GV.MuX = 0
        GV.MuY = 0
        GV.gR = 0
        txtPath = ""

        '===============================================================================
        If (GV.csvFolderType = 1) Then      'if multiple folders required
            If Not Directory.Exists(outputPath & "\" & N) Then         'create a new directory in output path with name as total no. of bases if it doesn't exist.
                Directory.CreateDirectory(outputPath & "\" & N)
            End If
            txtPath = outputPath & "\" & N & "\" & GV.fileNameNoPath.Substring(0, GV.fileNameNoPath.Length() - 4)      'path of output .csv file with file name
        ElseIf (GV.csvFolderType = 2) Then
            If Not Directory.Exists(outputPath & "\CSV Files") Then         'folder to store all csv files
                Directory.CreateDirectory(outputPath & "\CSV Files")
            End If
            txtPath = outputPath & "\CSV Files\" & GV.fileNameNoPath.Substring(0, GV.fileNameNoPath.Length() - 4)      'path of output .csv file with file name
        ElseIf (GV.csvFolderType = 3) Then
            skip = True

        Else
            MsgBox("This error shouldn't occur. " & vbNewLine & " Contact admin If you see this." & "Exiting ...", vbCritical, "Danger")
            Application.Exit()
        End If

        If (type = 1) Then
            txtPath = txtPath & " [" & CDS_Start & " to " & CDS_End & "].csv"
        Else
            txtPath = txtPath & ".csv"
        End If


        'Dim FileDelete As String
        'FileDelete = "C:\testDelete.txt"

        If System.IO.File.Exists(txtPath) = True Then       'deleting .csv if already exists
            System.IO.File.Delete(txtPath)
        End If

        If skip = False Then
            File.Create(txtPath).Dispose()                      'create new .csv file; Dispose to overwrite but maybe doesn't work :/
        End If

        'Dim objWriter As New System.IO.StreamWriter(GV.outputFolderPath & "\" & N & "\" & GV.fileNameNoPath & ".csv", True)

        Dim sb As StringBuilder = New StringBuilder()       'sb to store all coordinates of 1 sequence in a string builder; it is faster to store it in a string and not append it to txt file everytime inside the loop

        '===============================================================================
        If C = 1 Then       'for graph selection with C=1 i.e. Nandy
            For i = CDS_Start - 1 To CDS_End - 1        'i to get index of string/character in the sequence
                ch = Seq.Substring(i, 1)                'store each character 1 by 1 for further process

                'generating graph coordinates below
                If (ch = "g" Or ch = "G") Then
                    X = X + 1
                    SumG = SumG + 1
                ElseIf (ch = "a" Or ch = "A") Then
                    X = X - 1
                    SumA = SumA + 1
                ElseIf (ch = "c" Or ch = "C") Then
                    Y = Y + 1
                    SumC = SumC + 1
                ElseIf (ch = "t" Or ch = "T") Then
                    Y = Y - 1
                    SumT = SumT + 1
                ElseIf ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z")) Then
                    GV.CountN = GV.CountN + 1       'storing bases which are not required, for eg. n,k,y etc.
                    Continue For                    'skipping rest of the loop and continue with next iteration
                End If
                If skip = False Then
                    sb.AppendLine(X & "," & Y)          'storing coordinates in string builder
                End If
                SumX = SumX + X
                SumY = SumY + Y
                'ProgressBar1.Value = (i - CDS_Start) / (CDS_End - CDS_Start) * 100          'individual file progress bar
                'Threading.Thread.Sleep(1)
            Next
            If skip = False Then
                File.AppendAllText(txtPath, sb.ToString())      'all coordinates stored in string builder appended to .csv file
            End If
            GV.MuX = SumX / N       'calc Mu X
            GV.MuY = SumY / N       'calc Mu Y
            GV.gR = Math.Pow(GV.MuX * GV.MuX + GV.MuY * GV.MuY, 0.5)        'calc graph radius

        ElseIf C = 2 Then       'for graph selection with C=2 i.e. Gates
            For i = CDS_Start - 1 To CDS_End - 1        'i to get index of string/character in the sequence
                ch = Seq.Substring(i, 1)                'store each character 1 by 1 for further process

                'generating graph coordinates below
                If (ch = "c" Or ch = "C") Then
                    X = X + 1
                    SumC = SumC + 1
                ElseIf (ch = "g" Or ch = "G") Then
                    X = X - 1
                    SumG = SumG + 1
                ElseIf (ch = "t" Or ch = "T") Then
                    Y = Y + 1
                    SumT = SumT + 1
                ElseIf (ch = "a" Or ch = "A") Then
                    Y = Y - 1
                    SumA = SumA + 1
                ElseIf ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z")) Then
                    GV.CountN = GV.CountN + 1       'storing bases which are not required, for eg. n,k,y etc.
                    Continue For                    'skipping rest of the loop and continue with next iteration
                End If
                If skip = False Then
                    sb.AppendLine(X & "," & Y)          'storing coordinates in string builder
                End If
                SumX = SumX + X
                SumY = SumY + Y
                'ProgressBar1.Value = (i - CDS_Start) / (CDS_End - CDS_Start) * 100          'individual file progress bar
                'Threading.Thread.Sleep(1)
            Next
            If skip = False Then
                File.AppendAllText(txtPath, sb.ToString())      'all coordinates stored in string builder appended to .csv file
            End If

            GV.MuX = SumX / N       'calc Mu X
            GV.MuY = SumY / N       'calc Mu Y
            GV.gR = Math.Pow(GV.MuX * GV.MuX + GV.MuY * GV.MuY, 0.5)        'calc graph radius

        ElseIf C = 3 Then       'for graph selection with C=3 i.e. Leong and Morgenthaler
            For i = CDS_Start - 1 To CDS_End - 1        'i to get index of string/character in the sequence
                ch = Seq.Substring(i, 1)                'store each character 1 by 1 for further process

                'generating graph coordinates below
                If (ch = "a" Or ch = "A") Then
                    X = X + 1
                    SumA = SumA + 1
                ElseIf (ch = "c" Or ch = "C") Then
                    X = X - 1
                    SumC = SumC + 1
                ElseIf (ch = "t" Or ch = "T") Then
                    Y = Y + 1
                    SumT = SumT + 1
                ElseIf (ch = "g" Or ch = "G") Then
                    Y = Y - 1
                    SumG = SumG + 1
                ElseIf ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z")) Then
                    GV.CountN = GV.CountN + 1       'storing bases which are not required, for eg. n,k,y etc.
                    Continue For                    'skipping rest of the loop and continue with next iteration
                End If
                If skip = False Then
                    sb.AppendLine(X & "," & Y)          'storing coordinates in string builder
                End If
                SumX = SumX + X
                SumY = SumY + Y
                'ProgressBar1.Value = (i - CDS_Start) / (CDS_End - CDS_Start) * 100          'individual file progress bar
                'Threading.Thread.Sleep(1)
            Next
            If skip = False Then
                File.AppendAllText(txtPath, sb.ToString())      'all coordinates stored in string builder appended to .csv file
            End If

            GV.MuX = SumX / N       'calc Mu X
            GV.MuY = SumY / N       'calc Mu Y
            GV.gR = Math.Pow(GV.MuX * GV.MuX + GV.MuY * GV.MuY, 0.5)        'calc graph radius

        ElseIf C = 4 Then       'for graph selection with C=4 i.e. Custom01
            For i = CDS_Start - 1 To CDS_End - 1        'i to get index of string/character in the sequence
                ch = Seq.Substring(i, 1)                'store each character 1 by 1 for further process

                'generating graph coordinates below
                If (ch = "t" Or ch = "T") Then
                    X = X + 1
                    SumT = SumT + 1
                ElseIf (ch = "a" Or ch = "A") Then
                    X = X - 1
                    SumA = SumA + 1
                ElseIf (ch = "g" Or ch = "G") Then
                    Y = Y + 1
                    SumG = SumG + 1
                ElseIf (ch = "c" Or ch = "C") Then
                    Y = Y - 1
                    SumC = SumC + 1
                ElseIf ((ch >= "a" And ch <= "z") Or (ch >= "A" And ch <= "Z")) Then
                    GV.CountN = GV.CountN + 1       'storing bases which are not required, for eg. n,k,y etc.
                    Continue For                    'skipping rest of the loop and continue with next iteration
                End If
                If skip = False Then
                    sb.AppendLine(X & "," & Y)          'storing coordinates in string builder
                End If
                SumX = SumX + X
                SumY = SumY + Y
                'ProgressBar1.Value = (i - CDS_Start) / (CDS_End - CDS_Start) * 100          'individual file progress bar
                'Threading.Thread.Sleep(1)
            Next

            If skip = False Then
                File.AppendAllText(txtPath, sb.ToString())      'all coordinates stored in string builder appended to .csv file
            End If




            GV.MuX = SumX / N       'calc Mu X
            GV.MuY = SumY / N       'calc Mu Y
            GV.gR = Math.Pow(GV.MuX * GV.MuX + GV.MuY * GV.MuY, 0.5)        'calc graph radius

        End If
        GV.SumA = SumA
        GV.SumC = SumC
        GV.SumG = SumG
        GV.SumT = SumT
    End Sub
    Public Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load         'defaulting graph selection as first index during program run.
        If GraphStyleCB.Items.Count > 0 Then
            GraphStyleCB.SelectedIndex = 0    ' The first item has index 0 '
        End If
        RadioButton3.Checked = True
    End Sub

    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click        'tab1 input folder path browse button
        'get input folder path containing .txt files
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.inputFolderPath = FolderBrowserDialog1.SelectedPath                  'storing in global variable
            TextBox1.Text = GV.inputFolderPath                                      'showing path in textbox for user's ease
            TextBox1.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Public Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click       'tab1 output folder path browse button
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.outputFolderPath = FolderBrowserDialog1.SelectedPath                 'storing in global variable
            TextBox2.Text = GV.outputFolderPath                                     'showing path in textbox for user's ease
            TextBox2.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Public Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click     'tab1 execute button
        'Executes main code; Execute Button
        GV.inputFolderPath = TextBox1.Text          'reading string from textbox in case user enters path manually
        GV.outputFolderPath = TextBox2.Text         'reading string from textbox in case user enters path manually

        If (GV.inputFolderPath = "" Or GV.outputFolderPath = "") Then
            MsgBox("Enter the Input Folder Path And/Or the Output folder Path and try again.")
            Return
        End If

        If ((Directory.Exists(GV.inputFolderPath)) = False Or (Directory.Exists(GV.outputFolderPath) = False)) Then
            MsgBox("Invalid input/output folder entered. Enter valid folder(s) and try again." & vbNewLine & "Or select folder(s) through the Browse Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If


        Dim fileEntries As String() = Directory.GetFiles(GV.inputFolderPath, "*.txt")   'Process the list of .txt files found in the directory. Storing the file names of txt files including the path in a string array.
        Dim fileName, fileContent, CDS, Sequence, SequenceConverted, definition As String           'fileName to store individual file name with path of txt files, fileContent to store contents of each txt file, CDS to store the CDS range, Sequence to store all bases with line nos and spaces, SequenceConverted to store only bases (with rejected bases)
        Dim ln, lastIndexCDS, lastIndexOrigin, lastIndexDoubleSlash, count, CDS_End_Index, lastIndexDefinition, lastIndexAccession As Integer    'ln to store length of string array containing all file names, lastIndexCDS to store the last occurrence of 'CDS' string in fileContent, lastIndexOrigin to store the last occurrence of 'ORIGIN' string in fileContent, same for '//'
        Dim CDS_Start, CDS_End, C As Integer        'CDS_Start to store CDS starting value, CDS_End to store CDS end value, C to store graph type selection
        Dim N, progCounter As Integer               'N to store no. of bases including rejected bases, progCounter to store current progress (out of total job)

        Dim incompleteCounter As Integer            'to store no of .txt files with incomplete sequences
        incompleteCounter = 0                       'initialize as 0
        'Dim MuX, MuY, gR As Double

        Array.Sort(fileEntries)

        ln = fileEntries.Length                     'store total no. of .txt files
        count = 1                                   'counts valid .txt files and stores Sl. No. after processing them to the excel file.

        '===============================================================================================

        Dim fileTest As String = GV.outputFolderPath & "\Complete Details.xlsx"         'output excel file name
        If File.Exists(fileTest) Then                                                   'delete excel file if already exists
            File.Delete(fileTest)
        End If

        Dim dateStart As Date = Date.Now            'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        Dim oExcel As Object                        'object of Excel
        oExcel = CreateObject("Excel.Application")
        Dim oBook As Excel.Workbook                 'Workbook = excel file
        Dim oSheet As Excel.Worksheet               'Worksheet = excel sheet (out of many sheets)




        oBook = oExcel.Workbooks.Add                'adding one workbook to store Complete details of processed .txt files
        oSheet = oExcel.Worksheets(1)               'adding a sheet to the workbook

        oSheet.Name = "Master Data"                 'renaming the sheet

        'Row 1 of sheet to contain headings
        oSheet.Range("A1").Value = "Sl. No."
        oSheet.Range("B1").Value = "File Name"
        oSheet.Range("C1").Value = "Definition"
        oSheet.Range("D1").Value = "Sequence"
        oSheet.Range("E1").Value = "CDS Start"
        oSheet.Range("F1").Value = "CDS End"
        oSheet.Range("G1").Value = "Total Base 'N'"
        oSheet.Range("H1").Value = "No. of Rejected base 'n',etc."
        oSheet.Range("I1").Value = "Mu X"
        oSheet.Range("J1").Value = "Mu Y"
        oSheet.Range("K1").Value = "gR"
        '===============================================================================================
        fileContent = ""

        'changing visibility of progress bars,etc. to visible during run time for users to see current progress
        ProgressBar2.Visible = True
        Label5.Visible = True
        'ProgressBar1.Visible = True
        'Label6.Visible = True
        progCounter = 1         'current progress out of total files
        For Each fileName In fileEntries        'for each loop taking all values 1 by 1 from the String array containing all file names with path
            If (System.IO.File.Exists(fileName)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
                fileContent = File.ReadAllText(fileName)
                lastIndexCDS = lastIndex(fileContent, "CDS ")    'store last index/occurrence of CDS String
                CDS = fileContent.Substring(lastIndexCDS + 16, 14)  'storing the range of CDS
                CDS = CDS.Trim      'removing spaces from start and end

                If (CDS.IndexOf("<") <> -1 Or CDS.IndexOf(">") <> -1) Then      'checking for incomplete sequences and skipping them
                    incompleteCounter = incompleteCounter + 1                   'counting the no. of skipped files i.e. files containing incomplete sequences
                    Continue For                                                'return to start of loop with increment i.e. next iteration
                End If
                lastIndexOrigin = lastIndex(fileContent, "ORIGIN")              'store last index/occurrence of ORIGIN String
                lastIndexDoubleSlash = lastIndex(fileContent, "//")             'store last index/occurrence of '//' String
                Sequence = fileContent.Substring(lastIndexOrigin + 6, lastIndexDoubleSlash - lastIndexOrigin - 7)       'taking the part of the txt file which contains only the sequence (but with line nos and spaces)
                Sequence = Sequence.Trim        'removing spaces from start and end, but doesnt work on spaces in between
                SequenceConverted = convertSequence(Sequence)       'UDF for removing line nos. and spaces, i.e. takes into account only letters be it upper case or lower

                CDS_End_Index = lastIndex(CDS, ".")                 'store last occurrence of '.' in CDS range
                CDS_Start = CInt(CDS.Substring(0, CDS_End_Index - 1))       'store CDS start as integer
                CDS_End = CInt(CDS.Substring(CDS_End_Index + 1, CDS.Length - CDS_End_Index - 1))        'store CDS end as integer

                N = CDS_End - CDS_Start + 1     'store total no. of bases


                lastIndexDefinition = lastIndex(fileContent, "DEFINITION")
                lastIndexAccession = lastIndex(fileContent, "ACCESSION")

                definition = fileContent.Substring(lastIndexDefinition + "DEFINITION".Length, lastIndexAccession - lastIndexDefinition - "DEFINITION".Length - 1)   'getting definition
                definition = definition.Trim    'remove spaces before start and after end of string
                definition = definition.Replace(vbCr, "").Replace(vbLf, "")     'remove new line/carriage return from string
                definition = Regex.Replace(definition, " {2,}", " ")    'remove excess spaces i.e. convert multiple spaces to just 1

                'C to store index of selected graph style
                If GraphStyleCB.Text = "Nandy" Then
                    C = 1
                ElseIf GraphStyleCB.Text = "Gates" Then
                    C = 2
                ElseIf GraphStyleCB.Text = "Leong and Morgenthaler" Then
                    C = 3
                ElseIf GraphStyleCB.Text = "Custom01" Then
                    C = 4
                Else
                    MsgBox("Enter correct Graph Type.", vbOK, "Incorrect Graph Type")
                End If

                GV.fileNameNoPath = fileName.Substring(lastIndex(fileName, "\") + 1, fileName.Length - lastIndex(fileName, "\") - 1)        'to store name of txt file only (without its path)
                graph_Gen(C, SequenceConverted, N, CDS_Start, CDS_End, GV.outputFolderPath)      'UDF to generate .csv files and calculate required stuff



                '===============================================================================================
                'writing to Master Sheet (excel)
                count = count + 1
                oSheet.Range("A" & count).Value = count - 1
                oSheet.Range("B" & count).Value = GV.fileNameNoPath
                oSheet.Range("C" & count).Value = definition
                oSheet.Range("D" & count).Value = SequenceConverted
                oSheet.Range("E" & count).Value = CDS_Start
                oSheet.Range("F" & count).Value = CDS_End
                oSheet.Range("G" & count).Value = N
                oSheet.Range("H" & count).Value = GV.CountN
                oSheet.Range("I" & count).Value = GV.MuX
                oSheet.Range("J" & count).Value = GV.MuY
                oSheet.Range("K" & count).Value = GV.gR
                '===============================================================================================
            End If
            ProgressBar2.Value = progCounter / fileEntries.Length * 100         'overall progress display
            progCounter = progCounter + 1
            'Threading.Thread.Sleep(5)
        Next

        'making progress bars, etc invisible after job is complete.
        ProgressBar2.Visible = False
        Label5.Visible = False
        'ProgressBar1.Visible = False
        'Label6.Visible = False
        '===============================================================================================
        'saving and closing excel, although it remains in the memory unable for users to see, can only be seen in processes, but taken care of through process kill function later on
        oBook.SaveAs(fileTest)
        oBook.Close()
        oBook = Nothing
        oExcel.Quit()
        oExcel = Nothing

        Dim dateEnd As Date = Date.Now      'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        End_Excel_App(dateStart, dateEnd)   'This closes excel process
        MsgBox("Job Complete", vbOKOnly, "DONE")
        If incompleteCounter > 0 Then
            MsgBox("No. of skipped Files cause of incomplete sequence : " & incompleteCounter, vbInformation, "Info")
        End If
        '===============================================================================================
    End Sub

    Public Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click    'Info Button
        MsgBox("1st TAB" & vbNewLine & "This program was developed to convert dna sequences downloaded from NCBI GenBank as .txt files into 2D Coordinates as .csv files which can further be used to plot 2D Graphs." & vbNewLine & vbNewLine & "It also generates a .xlsx file which contains the details of all .txt files containing dna bases and some calculated data based on the information obtained from the .txt files." & vbNewLine & vbNewLine & "2nd TAB" & vbNewLine & "This tab was added later on to access data from already summarized excel file and generate output(Coordinates as .csv files and another excel summary) based on custom CDS Range.", vbInformation, "Info")
    End Sub

    Public Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click   'Help Button
        MsgBox("For any issues, problems or feedback please contact me at:" & vbNewLine & "tathagata.dk@gmail.com", vbInformation, "Help")
    End Sub



    Public Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click   'tab2 browse .xlsx file button
        Dim fd As OpenFileDialog = New OpenFileDialog()
        'Dim strFileName As String

        fd.Title = "Open File Dialog"
        fd.InitialDirectory = "C:\"
        'fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.Filter = "Excel files|*.xlsx"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            'strFileName = fd.FileName
            GV.inputExcelFile = fd.FileName
        End If
        TextBox3.ReadOnly = False
        TextBox3.Text = GV.inputExcelFile
        TextBox3.ReadOnly = True
    End Sub

    Public Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click   'tab2 output folder path button
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.outputExcelPath = FolderBrowserDialog1.SelectedPath                  'storing in global variable
            TextBox4.Text = GV.outputExcelPath                                      'showing path in textbox for user's ease
            TextBox4.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Public Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click 'tab2 execute button
        'MsgBox(GV.csvFolderType)

        'Executes main code; Execute Button
        GV.inputExcelFile = TextBox3.Text           'reading string from textbox in case user enters path manually
        GV.outputExcelPath = TextBox4.Text          'reading string from textbox in case user enters path manually

        If (GV.inputExcelFile = "" Or GV.outputExcelPath = "") Then
            MsgBox("Enter the Input File And/Or the Output folder Path and try again.")
            Return
        End If

        If ((File.Exists(GV.inputExcelFile)) = False Or (Directory.Exists(GV.outputExcelPath) = False)) Then
            MsgBox("Invalid input file/output folder entered. Enter valid file/folder and try again." & vbNewLine & "Or select file/folder through the Browse File/Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If

        Dim ExApp As New Excel.Application
        Dim workbook As Excel.Workbook
        Dim worksheet As Excel.Worksheet

        Dim dateStart2 As Date = Date.Now            'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        workbook = ExApp.Workbooks.Open(GV.inputExcelFile)
        worksheet = workbook.Worksheets("Master Data")
        'MsgBox(worksheet.Cells(1, 1).Value & vbNewLine & worksheet.Cells(2, 1).Value & vbNewLine & worksheet.Cells(3, 1).Value)
        'MsgBox(worksheet.UsedRange.Rows.Count)
        'MsgBox(C)
        Dim rowCount, C, N, CDS_Start, CDS_End As Integer
        Dim st As String





        Dim fileTest As String = GV.outputExcelPath & "\Complete Details Edit 2.xlsx"         'output excel file name
        If File.Exists(fileTest) Then                                                   'delete excel file if already exists
            File.Delete(fileTest)
        End If


        Dim o2Excel As Object                        'object of Excel
        o2Excel = CreateObject("Excel.Application")
        Dim oWkBook As Excel.Workbook                 'Workbook = excel file
        Dim oWkSheet As Excel.Worksheet               'Worksheet = excel sheet (out of many sheets)
        oWkBook = o2Excel.Workbooks.Add                'adding one workbook to store Complete details of processed .txt files
        oWkSheet = o2Excel.Worksheets(1)               'adding a sheet to the workbook

        oWkSheet.Name = "Master Data"                 'renaming the sheet

        'Row 1 of sheet to contain headings
        oWkSheet.Range("A1").Value = "Sl. No."
        oWkSheet.Range("B1").Value = "File Name"
        oWkSheet.Range("C1").Value = "Definition"
        oWkSheet.Range("D1").Value = "Sequence"
        oWkSheet.Range("E1").Value = "CDS Start"
        oWkSheet.Range("F1").Value = "CDS End"
        oWkSheet.Range("G1").Value = "Total Base 'N'"
        oWkSheet.Range("H1").Value = "No. of Rejected base 'n',etc."
        oWkSheet.Range("I1").Value = "Mu X"
        oWkSheet.Range("J1").Value = "Mu Y"
        oWkSheet.Range("K1").Value = "gR"








        If GraphStyleCB.Text = "Nandy" Then
            C = 1
        ElseIf GraphStyleCB.Text = "Gates" Then
            C = 2
        ElseIf GraphStyleCB.Text = "Leong and Morgenthaler" Then
            C = 3
        ElseIf GraphStyleCB.Text = "Custom01" Then
            C = 4
        Else
            MsgBox("Enter correct Graph Type.", vbOK, "Incorrect Graph Type")
        End If


        rowCount = worksheet.UsedRange.Rows.Count

        st = ""

        ProgressBar3.Visible = True
        Label10.Visible = True

        For i = 2 To rowCount
            oWkSheet.Range("A" & i).Value = worksheet.Cells(i, 1)
            GV.fileNameNoPath = worksheet.Cells(i, 2).Value
            oWkSheet.Range("B" & i).Value = GV.fileNameNoPath
            oWkSheet.Range("C" & i).Value = worksheet.Cells(i, 3)
            st = worksheet.Cells(i, 4).Value
            oWkSheet.Range("D" & i).Value = st
            'oWkSheet.Range("D" & i).Value = worksheet.Cells(i, 4)
            CDS_Start = worksheet.Cells(i, 5).Value
            CDS_End = worksheet.Cells(i, 6).Value
            N = CDS_End - CDS_Start + 1
            oWkSheet.Range("E" & i).Value = worksheet.Cells(i, 5)
            oWkSheet.Range("F" & i).Value = worksheet.Cells(i, 6)
            graph_Gen(C, st, N, CDS_Start, CDS_End, GV.outputExcelPath)
            oWkSheet.Range("G" & i).Value = N
            oWkSheet.Range("H" & i).Value = GV.CountN
            oWkSheet.Range("I" & i).Value = GV.MuX
            oWkSheet.Range("J" & i).Value = GV.MuY
            oWkSheet.Range("K" & i).Value = GV.gR

            ProgressBar3.Value = i / rowCount * 100         'overall progress display

        Next

        ProgressBar3.Visible = False
        Label10.Visible = False

        oWkBook.SaveAs(fileTest)
        oWkBook.Close()
        oWkBook = Nothing
        o2Excel.Quit()
        o2Excel = Nothing

        workbook.Save()
        workbook.Close()
        ExApp.Quit()

        Dim dateEnd2 As Date = Date.Now      'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        End_Excel_App(dateStart2, dateEnd2)   'This closes excel process
        MsgBox("Job Complete", vbOKOnly, "DONE")


    End Sub

    Public Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged   'tab1 multiple output folder radio button
        GV.csvFolderType = 1
    End Sub

    Public Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged   'tab1 single output folder radio button
        GV.csvFolderType = 2
    End Sub
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        GV.csvFolderType = 3
    End Sub

    Public Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click     'kills all running excel processes
        Dim xlp() As Process = Process.GetProcessesByName("EXCEL")
        For Each Process As Process In xlp
            Process.Kill()
        Next
        MsgBox("All EXCEL processes killed.", vbInformation, "Process(es) Killed")
    End Sub

    Public Sub GraphStyleCB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GraphStyleCB.SelectedIndexChanged
        'ComboBox (Drop Down List) Selection based image output:
        If GraphStyleCB.SelectedIndex = 0 Then
            GraphStylePB.Image = My.Resources.Nandy
            'GraphStylePB2.Image = My.Resources.Nandy
        ElseIf GraphStyleCB.SelectedIndex = 1 Then
            GraphStylePB.Image = My.Resources.Gates
            'GraphStylePB2.Image = My.Resources.Gates
        ElseIf GraphStyleCB.SelectedIndex = 2 Then
            GraphStylePB.Image = My.Resources.Leong_and_Morgenthaler
            'GraphStylePB2.Image = My.Resources.Leong_and_Morgenthaler
        ElseIf GraphStyleCB.SelectedIndex = 3 Then
            GraphStylePB.Image = My.Resources.Custom01
            'GraphStylePB2.Image = My.Resources.Custom01
        Else
            MessageBox.Show("Wrong Choice", "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
    End Sub

    Private Sub Button14_Click(sender As Object, e As EventArgs) Handles Button14.Click     'tab3 input folder browse button
        'get input folder path containing .txt files
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.Tab3InpFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox6.Text = GV.Tab3InpFolderPath                                    'showing path in textbox for user's ease
            TextBox6.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click     'tab3 output folder browse button
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.Tab3OutFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox5.Text = GV.Tab3OutFolderPath                                    'showing path in textbox for user's ease
            TextBox5.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button12_Click(sender As Object, e As EventArgs) Handles Button12.Click     'tab3 execute button


        'Dim customKeyword As String

        'Executes main code; Execute Button
        GV.Tab3InpFolderPath = TextBox6.Text          'reading string from textbox in case user enters path manually
        GV.Tab3OutFolderPath = TextBox5.Text         'reading string from textbox in case user enters path manually

        If (GV.Tab3InpFolderPath = "" Or GV.Tab3OutFolderPath = "") Then
            MsgBox("Enter the Input Folder Path And/Or the Output folder Path and try again.")
            Return
        End If

        If ((Directory.Exists(GV.Tab3InpFolderPath)) = False Or (Directory.Exists(GV.Tab3OutFolderPath) = False)) Then
            MsgBox("Invalid input/output folder entered. Enter valid folder(s) and try again." & vbNewLine & "Or select folder(s) through the Browse Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If

        'If (TextBox7.Text = "") Then
        '    MsgBox("Enter keyword and try again", vbOKOnly, "Error")
        '    Return
        'End If

        Dim customSearch() As String
        Dim i As Integer
        customSearch = RichTextBox1.Lines

        If (customSearch.Length = 0) Then
            MsgBox("Enter keyword(s) and try again", vbOKOnly, "Error")
            Return
        End If

        For i = 0 To customSearch.GetUpperBound(0)
            customSearch(i) = "/product=" & """" & customSearch(i) & """"
            'MsgBox(customSearch(i))
        Next

        'customKeyword = TextBox7.Text
        'customKeyword = """" & customKeyword & """"
        'MsgBox(customKeyword & vbNewLine & customKeyword.Length())

        Dim fileEntries As String() = Directory.GetFiles(GV.Tab3InpFolderPath, "*.txt")   'Process the list of .txt files found in the directory. Storing the file names of txt files including the path in a string array.
        Dim fileName, fileContent, CDS, Sequence, SequenceConverted, definition, note As String           'fileName to store individual file name with path of txt files, fileContent to store contents of each txt file, CDS to store the CDS range, Sequence to store all bases with line nos and spaces, SequenceConverted to store only bases (with rejected bases)
        Dim ln, lastIndexCDS, lastIndexOrigin, lastIndexDoubleSlash, count, CDS_End_Index, lastIndexDefinition, lastIndexAccession, temp As Integer    'ln to store length of string array containing all file names, lastIndexCDS to store the last occurrence of 'CDS' string in fileContent, lastIndexOrigin to store the last occurrence of 'ORIGIN' string in fileContent, same for '//'
        Dim CDS_Start, CDS_End, C As Integer        'CDS_Start to store CDS starting value, CDS_End to store CDS end value, C to store graph type selection
        Dim N, progCounter As Integer               'N to store no. of bases including rejected bases, progCounter to store current progress (out of total job)

        Dim incompleteCounter As Integer            'to store no of .txt files with incomplete sequences
        incompleteCounter = 0                       'initialize as 0

        Dim doesntContainKeyword As Integer
        doesntContainKeyword = 0

        Array.Sort(fileEntries)

        ln = fileEntries.Length                     'store total no. of .txt files
        count = 1                                   'counts valid .txt files and stores Sl. No. after processing them to the excel file.

        '===============================================================================================

        Dim fileTest As String = GV.Tab3OutFolderPath & "\Complete Details.xlsx"         'output excel file name
        If File.Exists(fileTest) Then                                                   'delete excel file if already exists
            File.Delete(fileTest)
        End If

        Dim dateStart As Date = Date.Now            'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        Dim o3Excel As Object                        'object of Excel
        o3Excel = CreateObject("Excel.Application")
        Dim o3Book As Excel.Workbook                 'Workbook = excel file
        Dim o3Sheet As Excel.Worksheet               'Worksheet = excel sheet (out of many sheets)




        o3Book = o3Excel.Workbooks.Add                'adding one workbook to store Complete details of processed .txt files
        o3Sheet = o3Excel.Worksheets(1)               'adding a sheet to the workbook

        o3Sheet.Name = "Master Data"                 'renaming the sheet

        'Row 1 of sheet to contain headings
        o3Sheet.Range("A1").Value = "Sl. No."
        o3Sheet.Range("B1").Value = "File Name"
        o3Sheet.Range("C1").Value = "Definition"
        o3Sheet.Range("D1").Value = "Sequence"
        o3Sheet.Range("E1").Value = "CDS Start"
        o3Sheet.Range("F1").Value = "CDS End"
        o3Sheet.Range("G1").Value = "Total Base 'N'"
        o3Sheet.Range("H1").Value = "No. of Rejected base 'n',etc."
        o3Sheet.Range("I1").Value = "Mu X"
        o3Sheet.Range("J1").Value = "Mu Y"
        o3Sheet.Range("K1").Value = "gR"
        '===============================================================================================
        fileContent = ""

        'changing visibility of progress bars,etc. to visible during run time for users to see current progress
        ProgressBar4.Visible = True
        Label11.Visible = True
        'ProgressBar1.Visible = True
        'Label6.Visible = True
        progCounter = 1         'current progress out of total files

        Dim notFoundKeywordPath = GV.Tab3OutFolderPath & "\Skipped Files.rtf"

        If System.IO.File.Exists(notFoundKeywordPath) = True Then       'deleting .txt file if already exists
            System.IO.File.Delete(notFoundKeywordPath)
        End If
        File.Create(notFoundKeywordPath).Dispose()

        Dim notFoundSB As StringBuilder = New StringBuilder()   'not found file names in string builder
        Dim foundInFile As Boolean = False
        For Each fileName In fileEntries        'for each loop taking all values 1 by 1 from the String array containing all file names with path
            If (System.IO.File.Exists(fileName)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
                fileContent = File.ReadAllText(fileName)

                For i = 0 To customSearch.GetUpperBound(0)
                    'Console.WriteLine("Looking for [" + customSearch(i) + "] in file :" + fileName)
                    lastIndexCDS = lastIndex(fileContent, customSearch(i))


                    If (lastIndexCDS < 0) Then          'if not found
                        note = "/note=" & customSearch(i).Substring(9, customSearch(i).Length - 9)
                        lastIndexCDS = lastIndex(fileContent, note)
                        If (lastIndexCDS >= 0) Then
                            foundInFile = True
                            Exit For
                        End If
                        foundInFile = False
                        Continue For
                    Else                                'if found
                        foundInFile = True
                        'Console.WriteLine("========*****=======Found [" + customSearch(i) + "] in file :" + fileName)
                        Exit For
                    End If
                    'indirectly breaking from outer loop
                Next

                If foundInFile = False Then
                    doesntContainKeyword = doesntContainKeyword + 1     'if not found after going through full string
                    notFoundSB.AppendLine(Path.GetFileName(fileName))      'add file name only excluding path
                    GoTo JUMP
                End If


                'lastIndexCDS = lastIndex(fileContent, customSearch(i))    'store last index/occurrence of custom Keyword String

                'If (lastIndexCDS < 0) Then
                '    doesntContainKeyword = doesntContainKeyword + 1
                '    Continue For
                'End If


                'CDS = fileContent.Substring(lastIndexCDS + 16, 14)  'storing the range of CDS
                temp = lastIndexCDS
                lastIndexCDS = lastIndex(fileContent.Substring(0, temp), "mat_peptide")


                If (lastIndexCDS = -1) Then
                    lastIndexCDS = lastIndex(fileContent.Substring(0, temp), "misc_feature")
                End If

                CDS = fileContent.Substring(lastIndexCDS + 13, 30)  'storing the range of CDS


                CDS = CDS.Trim      'removing spaces from start and end

                'MsgBox(CDS)

                If (CDS.IndexOf("<") <> -1 Or CDS.IndexOf(">") <> -1) Then      'checking for incomplete sequences and skipping them
                    incompleteCounter = incompleteCounter + 1                   'counting the no. of skipped files i.e. files containing incomplete sequences
                    Continue For                                                'return to start of loop with increment i.e. next iteration
                End If
                lastIndexOrigin = lastIndex(fileContent, "ORIGIN")              'store last index/occurrence of ORIGIN String
                lastIndexDoubleSlash = lastIndex(fileContent, "//")             'store last index/occurrence of '//' String
                Sequence = fileContent.Substring(lastIndexOrigin + 6, lastIndexDoubleSlash - lastIndexOrigin - 7)       'taking the part of the txt file which contains only the sequence (but with line nos and spaces)
                Sequence = Sequence.Trim        'removing spaces from start and end, but doesnt work on spaces in between
                SequenceConverted = convertSequence(Sequence)       'UDF for removing line nos. and spaces, i.e. takes into account only letters be it upper case or lower

                CDS_End_Index = lastIndex(CDS, ".")                 'store last occurrence of '.' in CDS range
                CDS_Start = CInt(CDS.Substring(0, CDS_End_Index - 1))       'store CDS start as integer
                CDS_End = CInt(CDS.Substring(CDS_End_Index + 1, CDS.Length - CDS_End_Index - 1))        'store CDS end as integer

                N = CDS_End - CDS_Start + 1     'store total no. of bases


                lastIndexDefinition = lastIndex(fileContent, "DEFINITION")
                lastIndexAccession = lastIndex(fileContent, "ACCESSION")

                definition = fileContent.Substring(lastIndexDefinition + "DEFINITION".Length, lastIndexAccession - lastIndexDefinition - "DEFINITION".Length - 1)   'getting definition
                definition = definition.Trim    'remove spaces before start and after end of string
                definition = definition.Replace(vbCr, "").Replace(vbLf, "")     'remove new line/carriage return from string
                definition = Regex.Replace(definition, " {2,}", " ")    'remove excess spaces i.e. convert multiple spaces to just 1

                'C to store index of selected graph style
                If GraphStyleCB.Text = "Nandy" Then
                    C = 1
                ElseIf GraphStyleCB.Text = "Gates" Then
                    C = 2
                ElseIf GraphStyleCB.Text = "Leong and Morgenthaler" Then
                    C = 3
                ElseIf GraphStyleCB.Text = "Custom01" Then
                    C = 4
                Else
                    MsgBox("Enter correct Graph Type.", vbOK, "Incorrect Graph Type")
                End If

                GV.fileNameNoPath = fileName.Substring(lastIndex(fileName, "\") + 1, fileName.Length - lastIndex(fileName, "\") - 1)        'to store name of txt file only (without its path)
                graph_Gen(C, SequenceConverted, N, CDS_Start, CDS_End, GV.Tab3OutFolderPath)      'UDF to generate .csv files and calculate required stuff



                '===============================================================================================
                'writing to Master Sheet (excel)
                count = count + 1
                o3Sheet.Range("A" & count).Value = count - 1
                o3Sheet.Range("B" & count).Value = GV.fileNameNoPath
                o3Sheet.Range("C" & count).Value = definition
                o3Sheet.Range("D" & count).Value = SequenceConverted
                o3Sheet.Range("E" & count).Value = CDS_Start
                o3Sheet.Range("F" & count).Value = CDS_End
                o3Sheet.Range("G" & count).Value = N
                o3Sheet.Range("H" & count).Value = GV.CountN
                o3Sheet.Range("I" & count).Value = GV.MuX
                o3Sheet.Range("J" & count).Value = GV.MuY
                o3Sheet.Range("K" & count).Value = GV.gR
                '===============================================================================================
            End If

JUMP:


            ProgressBar4.Value = progCounter / fileEntries.Length * 100         'overall progress display
            progCounter = progCounter + 1
            'Threading.Thread.Sleep(5)
        Next

        File.AppendAllText(notFoundKeywordPath, notFoundSB.ToString())

        'making progress bars, etc invisible after job is complete.
        ProgressBar4.Visible = False
        Label11.Visible = False
        'ProgressBar1.Visible = False
        'Label6.Visible = False
        '===============================================================================================
        'saving and closing excel, although it remains in the memory unable for users to see, can only be seen in processes, but taken care of through process kill function later on
        o3Book.SaveAs(fileTest)
        o3Book.Close()
        o3Book = Nothing
        o3Excel.Quit()
        o3Excel = Nothing

        Dim dateEnd As Date = Date.Now      'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        End_Excel_App(dateStart, dateEnd)   'This closes excel process
        MsgBox("Job Complete", vbOKOnly, "DONE")
        If incompleteCounter > 0 Then
            MsgBox("No. of skipped Files cause of incomplete sequence : " & incompleteCounter, vbInformation, "Info")
        End If

        If doesntContainKeyword > 0 Then
            MsgBox("No. of skipped Files cause of not containing keyword : " & doesntContainKeyword, vbInformation, "Info")
        End If

        '===============================================================================================
    End Sub


    Private Sub Button20_Click(sender As Object, e As EventArgs) Handles Button20.Click     'tab4 input folder button
        'get input folder path containing .txt files
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.Tab4InpFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox8.Text = GV.Tab4InpFolderPath                                    'showing path in textbox for user's ease
            TextBox8.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button19_Click(sender As Object, e As EventArgs) Handles Button19.Click     'tab4 output folder button
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.Tab4OutFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox7.Text = GV.Tab4OutFolderPath                                    'showing path in textbox for user's ease
            TextBox7.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub
    Private Sub Button18_Click(sender As Object, e As EventArgs) Handles Button18.Click     'tab4 execute button
        'Dim customKeyword As String

        'Executes main code; Execute Button
        GV.Tab4InpFolderPath = TextBox8.Text          'reading string from textbox in case user enters path manually
        GV.Tab4OutFolderPath = TextBox7.Text         'reading string from textbox in case user enters path manually

        If (GV.Tab4InpFolderPath = "" Or GV.Tab4OutFolderPath = "") Then
            MsgBox("Enter the Input Folder Path And/Or the Output folder Path and try again.")
            Return
        End If

        If ((Directory.Exists(GV.Tab4InpFolderPath)) = False Or (Directory.Exists(GV.Tab4OutFolderPath) = False)) Then
            MsgBox("Invalid input/output folder entered. Enter valid folder(s) and try again." & vbNewLine & "Or select folder(s) through the Browse Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If





        Dim customSearch() As String
        Dim i, d As Integer
        customSearch = RichTextBox2.Lines

        If (customSearch.Length = 0) Then
            MsgBox("Enter keyword(s) and try again", vbOKOnly, "Error")
            Return
        End If

        If Integer.TryParse(TextBox9.Text, d) = False Then
            MsgBox("Length of first segment should not be BLANK and should not contain character(s)" & vbNewLine & "Enter valid length and try again.", vbCritical, "Invalid Folder")
            Return
        End If

        If (TextBox9.Text = "" Or CInt(TextBox9.Text) Mod 3 <> 0) Then
            MsgBox("Length of first segment should be a multiple of 3." & vbNewLine & "Enter valid length and try again.", vbCritical, "Invalid Folder")
            Return
        End If

        For i = 0 To customSearch.GetUpperBound(0)
            customSearch(i) = "/product=" & """" & customSearch(i) & """"
        Next



        Dim fileEntries As String() = Directory.GetFiles(GV.Tab4InpFolderPath, "*.txt")   'Process the list of .txt files found in the directory. Storing the file names of txt files including the path in a string array.
        Dim fileName, fileContent, CDS, Sequence, SequenceConverted, definition, note As String           'fileName to store individual file name with path of txt files, fileContent to store contents of each txt file, CDS to store the CDS range, Sequence to store all bases with line nos and spaces, SequenceConverted to store only bases (with rejected bases)
        Dim ln, lastIndexCDS, lastIndexOrigin, lastIndexDoubleSlash, count, CDS_End_Index, lastIndexDefinition, lastIndexAccession, temp As Integer    'ln to store length of string array containing all file names, lastIndexCDS to store the last occurrence of 'CDS' string in fileContent, lastIndexOrigin to store the last occurrence of 'ORIGIN' string in fileContent, same for '//'
        Dim CDS_Start, CDS_End, C As Integer        'CDS_Start to store CDS starting value, CDS_End to store CDS end value, C to store graph type selection
        Dim N, progCounter As Integer               'N to store no. of bases including rejected bases, progCounter to store current progress (out of total job)
        Dim host, country, colDate As String


        Dim incompleteCounter As Integer            'to store no of .txt files with incomplete sequences
        incompleteCounter = 0                       'initialize as 0

        Dim doesntContainKeyword As Integer
        doesntContainKeyword = 0

        Array.Sort(fileEntries)

        ln = fileEntries.Length                     'store total no. of .txt files
        count = 1                                   'counts valid .txt files and stores Sl. No. after processing them to the excel file.

        '===============================================================================================

        Dim fileTest As String = GV.Tab4OutFolderPath & "\Complete Details.xlsx"         'output excel file name
        If File.Exists(fileTest) Then                                                   'delete excel file if already exists
            File.Delete(fileTest)
        End If

        Dim dateStart As Date = Date.Now            'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        Dim o4Excel As Object                        'object of Excel
        o4Excel = CreateObject("Excel.Application")
        Dim o4Book As Excel.Workbook                 'Workbook = excel file
        Dim o4Sheet As Excel.Worksheet               'Worksheet = excel sheet (out of many sheets)




        o4Book = o4Excel.Workbooks.Add                'adding one workbook to store Complete details of processed .txt files
        o4Sheet = o4Excel.Worksheets(1)               'adding a sheet to the workbook

        o4Sheet.Name = "Master Data"                 'renaming the sheet

        'Row 1 of sheet to contain headings
        o4Sheet.Range("A1").Value = "Sl. No."
        o4Sheet.Range("B1").Value = "File Name"
        o4Sheet.Range("C1").Value = "Definition"
        o4Sheet.Range("D1").Value = "Host"
        o4Sheet.Range("E1").Value = "Country"
        o4Sheet.Range("F1").Value = "Collection Date"
        o4Sheet.Range("G1").Value = "Sequence"
        o4Sheet.Range("H1").Value = "CDS Start"
        o4Sheet.Range("I1").Value = "CDS End"
        o4Sheet.Range("J1").Value = "Total Base 'N'"
        o4Sheet.Range("K1").Value = "No. of Rejected base 'n',etc."
        o4Sheet.Range("L1").Value = "CDS Start (Seg 1)"
        o4Sheet.Range("M1").Value = "CDS End (Seg 1)"
        o4Sheet.Range("N1").Value = "Length (Seg 1)"
        o4Sheet.Range("O1").Value = "Mu X (Seg 1)"
        o4Sheet.Range("P1").Value = "Mu Y (Seg 1)"
        o4Sheet.Range("Q1").Value = "gR (Seg 1)"
        o4Sheet.Range("R1").Value = "CDS Start (Seg 2)"
        o4Sheet.Range("S1").Value = "CDS End (Seg 2)"
        o4Sheet.Range("T1").Value = "Length (Seg 2)"
        o4Sheet.Range("U1").Value = "Mu X (Seg 2)"
        o4Sheet.Range("V1").Value = "Mu Y (Seg 2)"
        o4Sheet.Range("W1").Value = "gR (Seg 2)"
        '===============================================================================================
        fileContent = ""

        'changing visibility of progress bars,etc. to visible during run time for users to see current progress
        ProgressBar1.Visible = True
        Label14.Visible = True
        'ProgressBar1.Visible = True
        'Label6.Visible = True
        progCounter = 1         'current progress out of total files

        Dim notFoundKeywordPath = GV.Tab4OutFolderPath & "\Skipped Files.rtf"

        If System.IO.File.Exists(notFoundKeywordPath) = True Then       'deleting .txt file if already exists
            System.IO.File.Delete(notFoundKeywordPath)
        End If
        File.Create(notFoundKeywordPath).Dispose()

        Dim notFoundSB As StringBuilder = New StringBuilder()   'not found file names in string builder
        Dim foundInFile As Boolean = False
        For Each fileName In fileEntries        'for each loop taking all values 1 by 1 from the String array containing all file names with path
            If (System.IO.File.Exists(fileName)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
                fileContent = File.ReadAllText(fileName)

                For i = 0 To customSearch.GetUpperBound(0)
                    'Console.WriteLine("Looking for [" + customSearch(i) + "] in file :" + fileName)
                    lastIndexCDS = lastIndex(fileContent, customSearch(i))


                    If (lastIndexCDS < 0) Then          'if not found
                        note = "/note=" & customSearch(i).Substring(9, customSearch(i).Length - 9)
                        lastIndexCDS = lastIndex(fileContent, note)
                        If (lastIndexCDS >= 0) Then
                            foundInFile = True
                            Exit For
                        End If
                        foundInFile = False
                        Continue For
                    Else                                'if found
                        foundInFile = True
                        'Console.WriteLine("========*****=======Found [" + customSearch(i) + "] in file :" + fileName)
                        Exit For
                    End If
                    'indirectly breaking from outer loop
                Next

                If foundInFile = False Then
                    doesntContainKeyword = doesntContainKeyword + 1     'if not found after going through full string
                    notFoundSB.AppendLine(Path.GetFileName(fileName))      'add file name only excluding path
                    GoTo JUMP
                End If


                'lastIndexCDS = lastIndex(fileContent, customSearch(i))    'store last index/occurrence of custom Keyword String

                'If (lastIndexCDS < 0) Then
                '    doesntContainKeyword = doesntContainKeyword + 1
                '    Continue For
                'End If


                'CDS = fileContent.Substring(lastIndexCDS + 16, 14)  'storing the range of CDS
                temp = lastIndexCDS
                lastIndexCDS = lastIndex(fileContent.Substring(0, temp), "mat_peptide")


                If (lastIndexCDS = -1) Then
                    lastIndexCDS = lastIndex(fileContent.Substring(0, temp), "misc_feature")
                End If

                CDS = fileContent.Substring(lastIndexCDS + 13, 30)  'storing the range of CDS


                CDS = CDS.Trim      'removing spaces from start and end

                'MsgBox(CDS)

                If (CDS.IndexOf("<") <> -1 Or CDS.IndexOf(">") <> -1) Then      'checking for incomplete sequences and skipping them
                    incompleteCounter = incompleteCounter + 1                   'counting the no. of skipped files i.e. files containing incomplete sequences
                    Continue For                                                'return to start of loop with increment i.e. next iteration
                End If
                lastIndexOrigin = lastIndex(fileContent, "ORIGIN")              'store last index/occurrence of ORIGIN String
                lastIndexDoubleSlash = lastIndex(fileContent, "//")             'store last index/occurrence of '//' String
                Sequence = fileContent.Substring(lastIndexOrigin + 6, lastIndexDoubleSlash - lastIndexOrigin - 7)       'taking the part of the txt file which contains only the sequence (but with line nos and spaces)
                Sequence = Sequence.Trim        'removing spaces from start and end, but doesnt work on spaces in between
                SequenceConverted = convertSequence(Sequence)       'UDF for removing line nos. and spaces, i.e. takes into account only letters be it upper case or lower

                CDS_End_Index = lastIndex(CDS, ".")                 'store last occurrence of '.' in CDS range
                CDS_Start = CInt(CDS.Substring(0, CDS_End_Index - 1))       'store CDS start as integer
                CDS_End = CInt(CDS.Substring(CDS_End_Index + 1, CDS.Length - CDS_End_Index - 1))        'store CDS end as integer

                N = CDS_End - CDS_Start + 1     'store total no. of bases


                lastIndexDefinition = lastIndex(fileContent, "DEFINITION")
                lastIndexAccession = lastIndex(fileContent, "ACCESSION")




                definition = fileContent.Substring(lastIndexDefinition + "DEFINITION".Length, lastIndexAccession - lastIndexDefinition - "DEFINITION".Length - 1)   'getting definition
                definition = definition.Trim    'remove spaces before start and after end of string
                definition = definition.Replace(vbCr, "").Replace(vbLf, "")     'remove new line/carriage return from string
                definition = Regex.Replace(definition, " {2,}", " ")    'remove excess spaces i.e. convert multiple spaces to just 1

                host = getDetails("/host=", fileContent)
                country = getDetails("/country=", fileContent)
                colDate = getDetails("/collection_date=", fileContent)






                'C to store index of selected graph style
                If GraphStyleCB.Text = "Nandy" Then
                    C = 1
                ElseIf GraphStyleCB.Text = "Gates" Then
                    C = 2
                ElseIf GraphStyleCB.Text = "Leong and Morgenthaler" Then
                    C = 3
                ElseIf GraphStyleCB.Text = "Custom01" Then
                    C = 4
                Else
                    MsgBox("Enter correct Graph Type.", vbOK, "Incorrect Graph Type")
                End If

                GV.fileNameNoPath = fileName.Substring(lastIndex(fileName, "\") + 1, fileName.Length - lastIndex(fileName, "\") - 1)        'to store name of txt file only (without its path)




                '===============================================================================================
                'writing to Master Sheet (excel)
                count = count + 1
                o4Sheet.Range("A" & count).Value = count - 1
                o4Sheet.Range("B" & count).Value = GV.fileNameNoPath
                o4Sheet.Range("C" & count).Value = definition
                o4Sheet.Range("D" & count).Value = host
                o4Sheet.Range("E" & count).Value = country
                o4Sheet.Range("F" & count).Value = colDate

                o4Sheet.Range("G" & count).Value = SequenceConverted
                o4Sheet.Range("H" & count).Value = CDS_Start
                o4Sheet.Range("I" & count).Value = CDS_End
                o4Sheet.Range("J" & count).Value = N
                o4Sheet.Range("K" & count).Value = GV.CountN

                graph_Gen(C, SequenceConverted, N, CDS_Start, CDS_Start + d - 1, GV.Tab4OutFolderPath, 1)      'UDF to generate .csv files and calculate required stuff

                o4Sheet.Range("L" & count).Value = CDS_Start
                o4Sheet.Range("M" & count).Value = CDS_Start + d - 1
                o4Sheet.Range("N" & count).Value = (CDS_Start + d - 1) - (CDS_Start) + 1

                o4Sheet.Range("O" & count).Value = GV.MuX
                o4Sheet.Range("P" & count).Value = GV.MuY
                o4Sheet.Range("Q" & count).Value = GV.gR

                graph_Gen(C, SequenceConverted, N, CDS_Start + d, CDS_End, GV.Tab4OutFolderPath, 1)      'UDF to generate .csv files and calculate required stuff

                o4Sheet.Range("R" & count).Value = CDS_Start + d
                o4Sheet.Range("S" & count).Value = CDS_End
                o4Sheet.Range("T" & count).Value = (CDS_End) - (CDS_Start + d) + 1

                o4Sheet.Range("U" & count).Value = GV.MuX
                o4Sheet.Range("V" & count).Value = GV.MuY
                o4Sheet.Range("W" & count).Value = GV.gR




                '===============================================================================================
            End If

JUMP:


            ProgressBar1.Value = progCounter / fileEntries.Length * 100         'overall progress display
            progCounter = progCounter + 1
            'Threading.Thread.Sleep(5)
        Next

        File.AppendAllText(notFoundKeywordPath, notFoundSB.ToString())

        'making progress bars, etc invisible after job is complete.
        ProgressBar1.Visible = False
        Label14.Visible = False
        'ProgressBar1.Visible = False
        'Label6.Visible = False
        '===============================================================================================
        'saving and closing excel, although it remains in the memory unable for users to see, can only be seen in processes, but taken care of through process kill function later on
        o4Book.SaveAs(fileTest)
        o4Book.Close()
        o4Book = Nothing
        o4Excel.Quit()
        o4Excel = Nothing

        Dim dateEnd As Date = Date.Now      'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        End_Excel_App(dateStart, dateEnd)   'This closes excel process
        MsgBox("Job Complete", vbOKOnly, "DONE")
        If incompleteCounter > 0 Then
            MsgBox("No. of skipped Files cause of incomplete sequence : " & incompleteCounter, vbInformation, "Info")
        End If

        If doesntContainKeyword > 0 Then
            MsgBox("No. of skipped Files cause of not containing keyword : " & doesntContainKeyword, vbInformation, "Info")
        End If

        '===============================================================================================
    End Sub
    Private Sub Button25_Click(sender As Object, e As EventArgs) Handles Button25.Click     'tab5 input folder browse button
        'get input folder path containing .txt files
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.tab5InpFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox12.Text = GV.tab5InpFolderPath                                    'showing path in textbox for user's ease
            TextBox12.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button24_Click(sender As Object, e As EventArgs) Handles Button24.Click     'tab5 output folder browse button
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.Tab5OutFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox11.Text = GV.Tab5OutFolderPath                                    'showing path in textbox for user's ease
            TextBox11.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub
    Private Sub Button23_Click(sender As Object, e As EventArgs) Handles Button23.Click     'tab5 execute button


        'Executes main code; Execute Button
        GV.tab5InpFolderPath = TextBox12.Text          'reading string from textbox in case user enters path manually
        GV.Tab5OutFolderPath = TextBox11.Text         'reading string from textbox in case user enters path manually

        If (GV.tab5InpFolderPath = "" Or GV.Tab5OutFolderPath = "") Then
            MsgBox("Enter the Input Folder Path And/Or the Output folder Path and try again.")
            Return
        End If

        If ((Directory.Exists(GV.tab5InpFolderPath)) = False Or (Directory.Exists(GV.Tab5OutFolderPath) = False)) Then
            MsgBox("Invalid input/output folder entered. Enter valid folder(s) and try again." & vbNewLine & "Or select folder(s) through the Browse Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If





        Dim customSearch() As String
        Dim i As Integer
        customSearch = RichTextBox3.Lines

        If (customSearch.Length = 0) Then
            MsgBox("Enter keyword(s) and try again", vbOKOnly, "Error")
            Return
        End If

        'If Integer.TryParse(TextBox9.Text, d) = False Then
        '    MsgBox("Length of first segment should not be BLANK and should not contain character(s)" & vbNewLine & "Enter valid length and try again.", vbCritical, "Invalid Folder")
        '    Return
        'End If

        'If (TextBox9.Text = "" Or CInt(TextBox9.Text) Mod 3 <> 0) Then
        '    MsgBox("Length of first segment should be a multiple of 3." & vbNewLine & "Enter valid length and try again.", vbCritical, "Invalid Folder")
        '    Return
        'End If

        For i = 0 To customSearch.GetUpperBound(0)
            customSearch(i) = "/product=" & """" & customSearch(i) & """"
        Next



        Dim fileEntries As String() = Directory.GetFiles(GV.tab5InpFolderPath, "*.txt")   'Process the list of .txt files found in the directory. Storing the file names of txt files including the path in a string array.
        Dim fileName, fileContent, CDS, Sequence, SequenceConverted, definition, note As String           'fileName to store individual file name with path of txt files, fileContent to store contents of each txt file, CDS to store the CDS range, Sequence to store all bases with line nos and spaces, SequenceConverted to store only bases (with rejected bases)
        Dim ln, lastIndexCDS, lastIndexOrigin, lastIndexDoubleSlash, count, CDS_End_Index, lastIndexDefinition, lastIndexAccession, temp As Integer    'ln to store length of string array containing all file names, lastIndexCDS to store the last occurrence of 'CDS' string in fileContent, lastIndexOrigin to store the last occurrence of 'ORIGIN' string in fileContent, same for '//'
        Dim WG_CDS_Start, WG_CDS_End, C As Integer        'CDS_Start to store CDS starting value, CDS_End to store CDS end value, C to store graph type selection
        Dim gene_CDS_Start, gene_CDS_End As Integer
        Dim progCounter As Integer               'progCounter to store current progress (out of total job)
        Dim host, country, colDate, organism As String
        Dim geneName As String

        Dim incompleteCounter As Integer            'to store no of .txt files with incomplete sequences
        incompleteCounter = 0                       'initialize as 0

        Dim doesntContainKeyword As Integer
        doesntContainKeyword = 0

        Array.Sort(fileEntries)

        ln = fileEntries.Length                     'store total no. of .txt files
        count = 1                                   'counts valid .txt files and stores Sl. No. after processing them to the excel file.

        '===============================================================================================

        geneName = customSearch(smallestStringArrayIndex(customSearch)).Substring("/product=".Length + 1, customSearch(smallestStringArrayIndex(customSearch)).Length - "/product=".Length - 2)

        Dim fileTest As String = GV.Tab5OutFolderPath & "\Complete Details of " & geneName & ".xlsx"         'output excel file name
        If File.Exists(fileTest) Then                                                   'delete excel file if already exists
            File.Delete(fileTest)
        End If

        Dim dateStart As Date = Date.Now            'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        Dim o5Excel As Object                        'object of Excel
        o5Excel = CreateObject("Excel.Application")
        Dim o5Book As Excel.Workbook                 'Workbook = excel file
        Dim o5Sheet As Excel.Worksheet               'Worksheet = excel sheet (out of many sheets)




        o5Book = o5Excel.Workbooks.Add                'adding one workbook to store Complete details of processed .txt files
        o5Sheet = o5Excel.Worksheets(1)               'adding a sheet to the workbook

        o5Sheet.Name = "Master Data"                 'renaming the sheet

        'Row 1 of sheet to contain headings
        o5Sheet.Range("A1").Value = "Sl. No."
        o5Sheet.Range("B1").Value = "File Name"
        o5Sheet.Range("C1").Value = "Definition"
        o5Sheet.Range("D1").Value = "Organism"
        o5Sheet.Range("E1").Value = "Host"
        o5Sheet.Range("F1").Value = "Country"
        o5Sheet.Range("G1").Value = "Collection Date"

        o5Sheet.Range("H1").Value = "gR (WG)"

        o5Sheet.Range("I1").Value = "gR (" & geneName & ")"

        'o5Sheet.Range("H1").Value = "gR (" & customSearch(smallestStringArrayIndex(customSearch)).Substring("/product=".Length + 1, customSearch(smallestStringArrayIndex(customSearch)).Length - "/product=".Length - 2) & ")"
        '===============================================================================================
        fileContent = ""

        'changing visibility of progress bars,etc. to visible during run time for users to see current progress
        ProgressBar5.Visible = True
        Label20.Visible = True
        'ProgressBar1.Visible = True
        'Label6.Visible = True
        progCounter = 1         'current progress out of total files

        Dim notFoundKeywordPath = GV.Tab5OutFolderPath & "\Skipped Files.rtf"

        If System.IO.File.Exists(notFoundKeywordPath) = True Then       'deleting .txt file if already exists
            System.IO.File.Delete(notFoundKeywordPath)
        End If
        File.Create(notFoundKeywordPath).Dispose()

        Dim notFoundSB As StringBuilder = New StringBuilder()   'not found file names in string builder
        Dim foundInFile As Boolean = False
        For Each fileName In fileEntries        'for each loop taking all values 1 by 1 from the String array containing all file names with path
            If (System.IO.File.Exists(fileName)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
                fileContent = File.ReadAllText(fileName)

                For i = 0 To customSearch.GetUpperBound(0)
                    'Console.WriteLine("Looking for [" + customSearch(i) + "] in file :" + fileName)
                    lastIndexCDS = lastIndex(fileContent, customSearch(i))


                    If (lastIndexCDS < 0) Then          'if not found
                        note = "/note=" & customSearch(i).Substring(9, customSearch(i).Length - 9)
                        lastIndexCDS = lastIndex(fileContent, note)
                        If (lastIndexCDS >= 0) Then
                            foundInFile = True
                            Exit For
                        End If
                        foundInFile = False
                        Continue For
                    Else                                'if found
                        foundInFile = True
                        'Console.WriteLine("========*****=======Found [" + customSearch(i) + "] in file :" + fileName)
                        Exit For
                    End If
                    'indirectly breaking from outer loop
                Next

                If foundInFile = False Then
                    doesntContainKeyword = doesntContainKeyword + 1     'if not found after going through full string
                    notFoundSB.AppendLine(Path.GetFileName(fileName))      'add file name only excluding path
                    GoTo JUMP
                End If


                'lastIndexCDS = lastIndex(fileContent, customSearch(i))    'store last index/occurrence of custom Keyword String

                'If (lastIndexCDS < 0) Then
                '    doesntContainKeyword = doesntContainKeyword + 1
                '    Continue For
                'End If


                'CDS = fileContent.Substring(lastIndexCDS + 16, 14)  'storing the range of CDS of required gene (eg. E, NS5 etc.)
                temp = lastIndexCDS
                lastIndexCDS = lastIndex(fileContent.Substring(0, temp), "mat_peptide")


                If (lastIndexCDS = -1) Then
                    lastIndexCDS = lastIndex(fileContent.Substring(0, temp), "misc_feature")
                End If

                CDS = fileContent.Substring(lastIndexCDS + 13, 30)  'storing the range of CDS


                CDS = CDS.Trim      'removing spaces from start and end

                'MsgBox(CDS)

                If (CDS.IndexOf("<") <> -1 Or CDS.IndexOf(">") <> -1) Then      'checking for incomplete sequences and skipping them
                    incompleteCounter = incompleteCounter + 1                   'counting the no. of skipped files i.e. files containing incomplete sequences
                    Continue For                                                'return to start of loop with increment i.e. next iteration
                End If
                lastIndexOrigin = lastIndex(fileContent, "ORIGIN")              'store last index/occurrence of ORIGIN String
                lastIndexDoubleSlash = lastIndex(fileContent, "//")             'store last index/occurrence of '//' String
                Sequence = fileContent.Substring(lastIndexOrigin + 6, lastIndexDoubleSlash - lastIndexOrigin - 7)       'taking the part of the txt file which contains only the sequence (but with line nos and spaces)
                Sequence = Sequence.Trim        'removing spaces from start and end, but doesnt work on spaces in between
                SequenceConverted = convertSequence(Sequence)       'UDF for removing line nos. and spaces, i.e. takes into account only letters be it upper case or lower

                CDS_End_Index = lastIndex(CDS, ".")                 'store last occurrence of '.' in CDS range
                gene_CDS_Start = CInt(CDS.Substring(0, CDS_End_Index - 1))       'store CDS start as integer
                gene_CDS_End = CInt(CDS.Substring(CDS_End_Index + 1, CDS.Length - CDS_End_Index - 1))        'store CDS end as integer

                'N = CDS_End - CDS_Start + 1     'store total no. of bases


                'finding whole genome CDS below
                lastIndexCDS = lastIndex(fileContent, "CDS ")    'store last index/occurrence of CDS String
                CDS = fileContent.Substring(lastIndexCDS + 16, 14)  'storing the range of CDS
                CDS = CDS.Trim      'removing spaces from start and end

                If (CDS.IndexOf("<") <> -1 Or CDS.IndexOf(">") <> -1) Then      'checking for incomplete sequences and skipping them
                    incompleteCounter = incompleteCounter + 1                   'counting the no. of skipped files i.e. files containing incomplete sequences
                    Continue For                                                'return to start of loop with increment i.e. next iteration
                End If

                CDS_End_Index = lastIndex(CDS, ".")                 'store last occurrence of '.' in CDS range
                WG_CDS_Start = CInt(CDS.Substring(0, CDS_End_Index - 1))       'store CDS start as integer
                WG_CDS_End = CInt(CDS.Substring(CDS_End_Index + 1, CDS.Length - CDS_End_Index - 1))        'store CDS end as integer
                'found whole genome CDS i.e. work done 


                lastIndexDefinition = lastIndex(fileContent, "DEFINITION")
                lastIndexAccession = lastIndex(fileContent, "ACCESSION")




                definition = fileContent.Substring(lastIndexDefinition + "DEFINITION".Length, lastIndexAccession - lastIndexDefinition - "DEFINITION".Length - 1)   'getting definition
                definition = definition.Trim    'remove spaces before start and after end of string
                definition = definition.Replace(vbCr, "").Replace(vbLf, "")     'remove new line/carriage return from string
                definition = Regex.Replace(definition, " {2,}", " ")    'remove excess spaces i.e. convert multiple spaces to just 1

                host = getDetails("/host=", fileContent)
                country = getDetails("/country=", fileContent)
                colDate = getDetails("/collection_date=", fileContent)
                organism = getDetails("/organism=", fileContent)






                'C to store index of selected graph style
                If GraphStyleCB.Text = "Nandy" Then
                    C = 1
                ElseIf GraphStyleCB.Text = "Gates" Then
                    C = 2
                ElseIf GraphStyleCB.Text = "Leong and Morgenthaler" Then
                    C = 3
                ElseIf GraphStyleCB.Text = "Custom01" Then
                    C = 4
                Else
                    MsgBox("Enter correct Graph Type.", vbOK, "Incorrect Graph Type")
                End If

                GV.fileNameNoPath = fileName.Substring(lastIndex(fileName, "\") + 1, fileName.Length - lastIndex(fileName, "\") - 1)        'to store name of txt file only (without its path)




                '===============================================================================================
                'writing to Master Sheet (excel)
                count = count + 1
                o5Sheet.Range("A" & count).Value = count - 1
                o5Sheet.Range("B" & count).Value = GV.fileNameNoPath
                o5Sheet.Range("C" & count).Value = definition
                o5Sheet.Range("D" & count).Value = organism
                o5Sheet.Range("E" & count).Value = host
                o5Sheet.Range("F" & count).Value = country
                o5Sheet.Range("G" & count).Value = colDate



                graph_Gen(C, SequenceConverted, WG_CDS_End - WG_CDS_Start + 1, WG_CDS_Start, WG_CDS_End, GV.Tab5OutFolderPath, 1)      'UDF to generate .csv files and calculate required stuff


                o5Sheet.Range("H" & count).Value = GV.gR

                graph_Gen(C, SequenceConverted, gene_CDS_End - gene_CDS_Start + 1, gene_CDS_Start, gene_CDS_End, GV.Tab5OutFolderPath, 1)      'UDF to generate .csv files and calculate required stuff


                o5Sheet.Range("I" & count).Value = GV.gR




                '===============================================================================================
            End If

JUMP:


            ProgressBar5.Value = progCounter / fileEntries.Length * 100         'overall progress display
            progCounter = progCounter + 1
            'Threading.Thread.Sleep(5)
        Next

        File.AppendAllText(notFoundKeywordPath, notFoundSB.ToString())

        'making progress bars, etc invisible after job is complete.
        ProgressBar5.Visible = False
        Label20.Visible = False
        'ProgressBar1.Visible = False
        'Label6.Visible = False
        '===============================================================================================
        'saving and closing excel, although it remains in the memory unable for users to see, can only be seen in processes, but taken care of through process kill function later on
        o5Book.SaveAs(fileTest)
        o5Book.Close()
        o5Book = Nothing
        o5Excel.Quit()
        o5Excel = Nothing

        Dim dateEnd As Date = Date.Now      'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        End_Excel_App(dateStart, dateEnd)   'This closes excel process
        MsgBox("Job Complete", vbOKOnly, "DONE")
        If incompleteCounter > 0 Then
            MsgBox("No. of skipped Files cause of incomplete sequence : " & incompleteCounter, vbInformation, "Info")
        End If

        If doesntContainKeyword > 0 Then
            MsgBox("No. of skipped Files cause of not containing keyword : " & doesntContainKeyword, vbInformation, "Info")
        End If
    End Sub

    Private Sub Button29_Click(sender As Object, e As EventArgs) Handles Button29.Click     'tab6 input folder
        'get input folder path containing .txt files
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.tab6InpFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox13.Text = GV.tab6InpFolderPath                                    'showing path in textbox for user's ease
            TextBox13.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button28_Click(sender As Object, e As EventArgs) Handles Button28.Click     'tab6 output folder
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.Tab6OutFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox10.Text = GV.Tab6OutFolderPath                                    'showing path in textbox for user's ease
            TextBox10.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub



    Private Sub Button27_Click(sender As Object, e As EventArgs) Handles Button27.Click     'tab6 execute button
        'Executes main code; Execute Button
        GV.tab6InpFolderPath = TextBox13.Text          'reading string from textbox in case user enters path manually
        GV.Tab6OutFolderPath = TextBox10.Text         'reading string from textbox in case user enters path manually

        If (GV.tab6InpFolderPath = "" Or GV.Tab6OutFolderPath = "") Then
            MsgBox("Enter the Input Folder Path And/Or the Output folder Path and try again.")
            Return
        End If

        If ((Directory.Exists(GV.tab6InpFolderPath)) = False Or (Directory.Exists(GV.Tab6OutFolderPath) = False)) Then
            MsgBox("Invalid input/output folder entered. Enter valid folder(s) and try again." & vbNewLine & "Or select folder(s) through the Browse Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If

        If (TextBox14.Text = "" Or TextBox15.Text = "") Then
            MsgBox("Enter the length of segment 1 and segment 2 and try again.")
            Return
        End If




        'Dim customSearch() As String
        'Dim i As Integer
        'customSearch = RichTextBox4.Lines

        'If (customSearch.Length = 0) Then
        '    MsgBox("Enter keyword(s) and try again", vbOKOnly, "Error")
        '    Return
        'End If


        'For i = 0 To customSearch.GetUpperBound(0)
        '    customSearch(i) = "/product=" & """" & customSearch(i) & """"
        'Next



        Dim fileEntries As String() = Directory.GetFiles(GV.tab6InpFolderPath, "*.txt")   'Process the list of .txt files found in the directory. Storing the file names of txt files including the path in a string array.
        Dim fileName, fileContent, CDS, Sequence, SequenceConverted, definition As String           'fileName to store individual file name with path of txt files, fileContent to store contents of each txt file, CDS to store the CDS range, Sequence to store all bases with line nos and spaces, SequenceConverted to store only bases (with rejected bases)
        Dim ln, lastIndexCDS, lastIndexOrigin, lastIndexDoubleSlash, count, CDS_End_Index, lastIndexDefinition, lastIndexAccession As Integer    'ln to store length of string array containing all file names, lastIndexCDS to store the last occurrence of 'CDS' string in fileContent, lastIndexOrigin to store the last occurrence of 'ORIGIN' string in fileContent, same for '//'
        Dim WG_CDS_Start, WG_CDS_End, C As Integer        'CDS_Start to store CDS starting value, CDS_End to store CDS end value, C to store graph type selection
        'Dim gene_CDS_Start, gene_CDS_End As Integer
        Dim progCounter As Integer               'progCounter to store current progress (out of total job)
        Dim host, country, colDate, organism, subType As String
        'Dim geneName As String
        Dim lenSeg1, lenSeg2 As Integer

        Dim incompleteCounter As Integer            'to store no of .txt files with incomplete sequences
        incompleteCounter = 0                       'initialize as 0

        Dim doesntContainKeyword As Integer
        doesntContainKeyword = 0

        Array.Sort(fileEntries)

        ln = fileEntries.Length                     'store total no. of .txt files
        count = 1                                   'counts valid .txt files and stores Sl. No. after processing them to the excel file.

        '===============================================================================================

        'geneName = customSearch(smallestStringArrayIndex(customSearch)).Substring("/product=".Length + 1, customSearch(smallestStringArrayIndex(customSearch)).Length - "/product=".Length - 2)

        Dim fileTest As String = GV.Tab6OutFolderPath & "\Complete Details of All.xlsx"         'output excel file name
        If File.Exists(fileTest) Then                                                   'delete excel file if already exists
            File.Delete(fileTest)
        End If

        Dim dateStart As Date = Date.Now            'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        Dim o6Excel As Object                        'object of Excel
        o6Excel = CreateObject("Excel.Application")
        Dim o6Book As Excel.Workbook                 'Workbook = excel file
        Dim o6Sheet As Excel.Worksheet               'Worksheet = excel sheet (out of many sheets)




        o6Book = o6Excel.Workbooks.Add                'adding one workbook to store Complete details of processed .txt files
        o6Sheet = o6Excel.Worksheets(1)               'adding a sheet to the workbook

        o6Sheet.Name = "Master Data"                 'renaming the sheet

        'Row 1 of sheet to contain headings
        o6Sheet.Range("A1").Value = "Sl. No."
        o6Sheet.Range("B1").Value = "File Name"
        o6Sheet.Range("C1").Value = "Definition"
        o6Sheet.Range("D1").Value = "Organism"
        o6Sheet.Range("E1").Value = "Sub Type"

        o6Sheet.Range("F1").Value = "Host"
        o6Sheet.Range("G1").Value = "Country"
        o6Sheet.Range("H1").Value = "Collection Date"

        o6Sheet.Range("I1").Value = "gR (WG)"

        o6Sheet.Range("J1").Value = "gR (Seg 1)"
        o6Sheet.Range("K1").Value = "gR (Seg 2)"
        o6Sheet.Range("L1").Value = "gR (Seg 3)"

        'o5Sheet.Range("H1").Value = "gR (" & customSearch(smallestStringArrayIndex(customSearch)).Substring("/product=".Length + 1, customSearch(smallestStringArrayIndex(customSearch)).Length - "/product=".Length - 2) & ")"
        '===============================================================================================
        fileContent = ""

        'changing visibility of progress bars,etc. to visible during run time for users to see current progress
        ProgressBar6.Visible = True
        Label23.Visible = True

        progCounter = 1         'current progress out of total files

        lenSeg1 = CInt(TextBox14.Text)
        lenSeg2 = CInt(TextBox15.Text)

        'Dim foundInFile As Boolean = False
        For Each fileName In fileEntries        'for each loop taking all values 1 by 1 from the String array containing all file names with path
            If (System.IO.File.Exists(fileName)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
                fileContent = File.ReadAllText(fileName)

                lastIndexOrigin = lastIndex(fileContent, "ORIGIN")              'store last index/occurrence of ORIGIN String
                lastIndexDoubleSlash = lastIndex(fileContent, "//")             'store last index/occurrence of '//' String
                Sequence = fileContent.Substring(lastIndexOrigin + 6, lastIndexDoubleSlash - lastIndexOrigin - 7)       'taking the part of the txt file which contains only the sequence (but with line nos and spaces)
                Sequence = Sequence.Trim        'removing spaces from start and end, but doesnt work on spaces in between
                SequenceConverted = convertSequence(Sequence)       'UDF for removing line nos. and spaces, i.e. takes into account only letters be it upper case or lower


                'finding whole genome CDS below
                lastIndexCDS = lastIndex(fileContent, "CDS ")    'store last index/occurrence of CDS String
                CDS = fileContent.Substring(lastIndexCDS + 16, 14)  'storing the range of CDS
                CDS = CDS.Trim      'removing spaces from start and end

                If (CDS.IndexOf("<") <> -1 Or CDS.IndexOf(">") <> -1) Then      'checking for incomplete sequences and skipping them
                    incompleteCounter = incompleteCounter + 1                   'counting the no. of skipped files i.e. files containing incomplete sequences
                    Continue For                                                'return to start of loop with increment i.e. next iteration
                End If

                CDS_End_Index = lastIndex(CDS, ".")                 'store last occurrence of '.' in CDS range
                WG_CDS_Start = CInt(CDS.Substring(0, CDS_End_Index - 1))       'store CDS start as integer
                WG_CDS_End = CInt(CDS.Substring(CDS_End_Index + 1, CDS.Length - CDS_End_Index - 1))        'store CDS end as integer
                'found whole genome CDS i.e. work done 


                lastIndexDefinition = lastIndex(fileContent, "DEFINITION")
                lastIndexAccession = lastIndex(fileContent, "ACCESSION")




                definition = fileContent.Substring(lastIndexDefinition + "DEFINITION".Length, lastIndexAccession - lastIndexDefinition - "DEFINITION".Length - 1)   'getting definition
                definition = definition.Trim    'remove spaces before start and after end of string
                definition = definition.Replace(vbCr, "").Replace(vbLf, "")     'remove new line/carriage return from string
                definition = Regex.Replace(definition, " {2,}", " ")    'remove excess spaces i.e. convert multiple spaces to just 1

                host = getDetails("/host=", fileContent)
                country = getDetails("/country=", fileContent)
                colDate = getDetails("/collection_date=", fileContent)
                organism = getDetails("/organism=", fileContent)
                subType = subTypeFinder("ORGANISM", fileContent)

                'C to store index of selected graph style
                If GraphStyleCB.Text = "Nandy" Then
                    C = 1
                ElseIf GraphStyleCB.Text = "Gates" Then
                    C = 2
                ElseIf GraphStyleCB.Text = "Leong and Morgenthaler" Then
                    C = 3
                ElseIf GraphStyleCB.Text = "Custom01" Then
                    C = 4
                Else
                    MsgBox("Enter correct Graph Type.", vbOK, "Incorrect Graph Type")
                End If

                GV.fileNameNoPath = fileName.Substring(lastIndex(fileName, "\") + 1, fileName.Length - lastIndex(fileName, "\") - 1)        'to store name of txt file only (without its path)




                '===============================================================================================
                'writing to Master Sheet (excel)
                count = count + 1
                o6Sheet.Range("A" & count).Value = count - 1
                o6Sheet.Range("B" & count).Value = GV.fileNameNoPath
                o6Sheet.Range("C" & count).Value = definition
                o6Sheet.Range("D" & count).Value = organism
                o6Sheet.Range("E" & count).Value = subType
                o6Sheet.Range("F" & count).Value = host
                o6Sheet.Range("G" & count).Value = country
                o6Sheet.Range("H" & count).Value = colDate



                graph_Gen(C, SequenceConverted, WG_CDS_End - WG_CDS_Start + 1, WG_CDS_Start, WG_CDS_End, GV.Tab6OutFolderPath, 1)      'UDF to generate .csv files and calculate required stuff
                o6Sheet.Range("I" & count).Value = GV.gR

                graph_Gen(C, SequenceConverted, WG_CDS_End - WG_CDS_Start + 1, WG_CDS_Start, WG_CDS_Start + lenSeg1 - 1, GV.Tab6OutFolderPath, 1)      'UDF to generate .csv files and calculate required stuff
                o6Sheet.Range("J" & count).Value = GV.gR

                graph_Gen(C, SequenceConverted, WG_CDS_End - WG_CDS_Start + 1, WG_CDS_Start + lenSeg1, WG_CDS_Start + lenSeg1 + lenSeg2 - 1, GV.Tab6OutFolderPath, 1)      'UDF to generate .csv files and calculate required stuff
                o6Sheet.Range("K" & count).Value = GV.gR

                graph_Gen(C, SequenceConverted, WG_CDS_End - WG_CDS_Start + 1, WG_CDS_Start + lenSeg1 + lenSeg2, WG_CDS_End, GV.Tab6OutFolderPath, 1)      'UDF to generate .csv files and calculate required stuff
                o6Sheet.Range("L" & count).Value = GV.gR


                '===============================================================================================
            End If

JUMP:


            ProgressBar6.Value = progCounter / fileEntries.Length * 100         'overall progress display
            progCounter = progCounter + 1
            'Threading.Thread.Sleep(5)
        Next


        'making progress bars, etc invisible after job is complete.
        ProgressBar6.Visible = False
        Label23.Visible = False
        'ProgressBar1.Visible = False
        'Label6.Visible = False
        '===============================================================================================
        'saving and closing excel, although it remains in the memory unable for users to see, can only be seen in processes, but taken care of through process kill function later on
        o6Book.SaveAs(fileTest)
        o6Book.Close()
        o6Book = Nothing
        o6Excel.Quit()
        o6Excel = Nothing

        Dim dateEnd As Date = Date.Now      'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        End_Excel_App(dateStart, dateEnd)   'This closes excel process
        MsgBox("Job Complete", vbOKOnly, "DONE")
        If incompleteCounter > 0 Then
            MsgBox("No. of skipped Files cause of incomplete sequence : " & incompleteCounter, vbInformation, "Info")
        End If

        If doesntContainKeyword > 0 Then
            MsgBox("No. of skipped Files cause of not containing keyword : " & doesntContainKeyword, vbInformation, "Info")
        End If
    End Sub

    Private Sub Button32_Click(sender As Object, e As EventArgs) Handles Button32.Click             'tab7 inpput folder
        'get input folder path containing .txt files
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.tab7InpFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox17.Text = GV.tab7InpFolderPath                                    'showing path in textbox for user's ease
            TextBox17.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button31_Click(sender As Object, e As EventArgs) Handles Button31.Click             'tab7 output folder
        'get output folder path to generate Combined data in .xlsx and individual .csv files for graph generation
        If FolderBrowserDialog1.ShowDialog = Windows.Forms.DialogResult.OK Then     'get the selected path value
            GV.Tab7OutFolderPath = FolderBrowserDialog1.SelectedPath                'storing in global variable
            TextBox16.Text = GV.Tab7OutFolderPath                                    'showing path in textbox for user's ease
            TextBox16.ReadOnly = True                                                'making the textbox uneditable i.e. path cannot be changed if selected throught Browse (Button1) button
        End If
    End Sub

    Private Sub Button30_Click(sender As Object, e As EventArgs) Handles Button30.Click             'tab7 execute button
        'Executes main code; Execute Button
        GV.tab7InpFolderPath = TextBox17.Text          'reading string from textbox in case user enters path manually
        GV.Tab7OutFolderPath = TextBox16.Text         'reading string from textbox in case user enters path manually

        If (GV.tab7InpFolderPath = "" Or GV.Tab7OutFolderPath = "") Then
            MsgBox("Enter the Input Folder Path And/Or the Output folder Path and try again.")
            Return
        End If

        If ((Directory.Exists(GV.tab7InpFolderPath)) = False Or (Directory.Exists(GV.Tab7OutFolderPath) = False)) Then
            MsgBox("Invalid input/output folder entered. Enter valid folder(s) and try again." & vbNewLine & "Or select folder(s) through the Browse Folder button(s) to prevent such error.", vbCritical, "Invalid Folder")
            Return
        End If




        Dim fileEntries As String() = Directory.GetFiles(GV.tab7InpFolderPath, "*.fasta")   'Process the list of .txt files found in the directory. Storing the file names of txt files including the path in a string array.
        Dim fileName, fileContent, Sequence, SequenceConverted, fastaHeader, year As String           'fileName to store individual file name with path of txt files, fileContent to store contents of each txt file, CDS to store the CDS range, Sequence to store all bases with line nos and spaces, SequenceConverted to store only bases (with rejected bases)
        Dim headerStartIndex, headerEndIndex, SeqStartIndex, SeqEndIndex, ln, count, C As Integer
        Dim progCounter As Integer               'progCounter to store current progress (out of total job)

        Array.Sort(fileEntries)

        ln = fileEntries.Length                     'store total no. of .txt files
        count = 1                                   'counts valid .txt files and stores Sl. No. after processing them to the excel file.
        year = "Not Found"

        '===============================================================================================

        'geneName = customSearch(smallestStringArrayIndex(customSearch)).Substring("/product=".Length + 1, customSearch(smallestStringArrayIndex(customSearch)).Length - "/product=".Length - 2)

        Dim fileTest As String = GV.Tab7OutFolderPath & "\Complete Details.xlsx"         'output excel file name
        If File.Exists(fileTest) Then                                                   'delete excel file if already exists
            File.Delete(fileTest)
        End If

        Dim dateStart As Date = Date.Now            'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        Dim o7Excel As Object                        'object of Excel
        o7Excel = CreateObject("Excel.Application")
        Dim o7Book As Excel.Workbook                 'Workbook = excel file
        Dim o7Sheet As Excel.Worksheet               'Worksheet = excel sheet (out of many sheets)




        o7Book = o7Excel.Workbooks.Add                'adding one workbook to store Complete details of processed .txt files
        o7Sheet = o7Excel.Worksheets(1)               'adding a sheet to the workbook

        o7Sheet.Name = "Master Data"                 'renaming the sheet

        'Row 1 of sheet to contain headings
        o7Sheet.Range("A1").Value = "Sl. No."
        o7Sheet.Range("B1").Value = "File Name"
        o7Sheet.Range("C1").Value = "Fasta Header"
        o7Sheet.Range("D1").Value = "gR"
        o7Sheet.Range("E1").Value = "A Count"
        o7Sheet.Range("F1").Value = "C Count"
        o7Sheet.Range("G1").Value = "G Count"
        o7Sheet.Range("H1").Value = "T Count"
        o7Sheet.Range("I1").Value = "No. of bases"
        o7Sheet.Range("J1").Value = "Year"
        'o5Sheet.Range("H1").Value = "gR (" & customSearch(smallestStringArrayIndex(customSearch)).Substring("/product=".Length + 1, customSearch(smallestStringArrayIndex(customSearch)).Length - "/product=".Length - 2) & ")"
        '===============================================================================================
        fileContent = ""

        'changing visibility of progress bars,etc. to visible during run time for users to see current progress
        ProgressBar7.Visible = True
        Label18.Visible = True

        progCounter = 1         'current progress out of total files


        'Dim foundInFile As Boolean = False
        For Each fileName In fileEntries        'for each loop taking all values 1 by 1 from the String array containing all file names with path
            If (System.IO.File.Exists(fileName)) Then       'if file exists (ofcourse it does, unless deleted at runtime midway xD) then store its entire contents to fileContent as String
                fileContent = File.ReadAllText(fileName)

                'lastIndexOrigin = lastIndex(fileContent, "ORIGIN")              'store last index/occurrence of ORIGIN String
                'lastIndexDoubleSlash = lastIndex(fileContent, "//")             'store last index/occurrence of '//' String
                headerStartIndex = fileContent.IndexOf(">")
                headerEndIndex = fileContent.IndexOf(vbLf)
                SeqStartIndex = headerEndIndex + 1
                SeqEndIndex = fileContent.Length - 1


                fastaHeader = fileContent.Substring(headerStartIndex, headerEndIndex - headerStartIndex)
                Sequence = fileContent.Substring(SeqStartIndex, SeqEndIndex - SeqStartIndex + 1)
                Sequence = Sequence.Trim        'removing spaces from start and end, but doesnt work on spaces in between
                SequenceConverted = convertSequence(Sequence)       'UDF for removing line nos. and spaces, i.e. takes into account only letters be it upper case or lower




                'C to store index of selected graph style
                If GraphStyleCB.Text = "Nandy" Then
                    C = 1
                ElseIf GraphStyleCB.Text = "Gates" Then
                    C = 2
                ElseIf GraphStyleCB.Text = "Leong and Morgenthaler" Then
                    C = 3
                ElseIf GraphStyleCB.Text = "Custom01" Then
                    C = 4
                Else
                    MsgBox("Enter correct Graph Type.", vbOK, "Incorrect Graph Type")
                End If

                GV.fileNameNoPath = fileName.Substring(lastIndex(fileName, "\") + 1, fileName.Length - lastIndex(fileName, "\") - 1)        'to store name of txt file only (without its path)

                If fastaHeader.ToLower.Contains("year=") Then
                    year = fastaHeader.ToLower.Substring(fastaHeader.ToLower.IndexOf("year") + 5, 4)        'considering 4 digit year
                Else
                    year = "Not Found"
                End If



                '===============================================================================================
                'writing to Master Sheet (excel)
                count = count + 1
                o7Sheet.Range("A" & count).Value = count - 1
                o7Sheet.Range("B" & count).Value = GV.fileNameNoPath
                o7Sheet.Range("C" & count).Value = fastaHeader


                graph_Gen(C, SequenceConverted, SequenceConverted.Length, 1, SequenceConverted.Length, GV.Tab7OutFolderPath)
                o7Sheet.Range("D" & count).Value = GV.gR
                o7Sheet.Range("E" & count).Value = GV.SumA
                o7Sheet.Range("F" & count).Value = GV.SumC
                o7Sheet.Range("G" & count).Value = GV.SumG
                o7Sheet.Range("H" & count).Value = GV.SumT
                o7Sheet.Range("I" & count).Value = SequenceConverted.Length.ToString
                o7Sheet.Range("J" & count).Value = year


                '===============================================================================================
            End If


            ProgressBar7.Value = progCounter / fileEntries.Length * 100         'overall progress display
            progCounter = progCounter + 1
            'Threading.Thread.Sleep(5)
        Next


        'making progress bars, etc invisible after job is complete.
        ProgressBar7.Visible = False
        Label18.Visible = False
        'ProgressBar1.Visible = False
        'Label6.Visible = False
        '===============================================================================================
        'saving and closing excel, although it remains in the memory unable for users to see, can only be seen in processes, but taken care of through process kill function later on
        o7Book.SaveAs(fileTest)
        o7Book.Close()
        o7Book = Nothing
        o7Excel.Quit()
        o7Excel = Nothing

        Dim dateEnd As Date = Date.Now      'storing current date (with Time) dynamically for killing EXCEL.EXE process later on to prevent memory leak
        End_Excel_App(dateStart, dateEnd)   'This closes excel process
        MsgBox("Job Complete", vbOKOnly, "DONE")
    End Sub

    '========================================================PREV AND NEXT TAB BUTTONS=========================================================
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click   'tab1 next tab button, i.e. takes to tab2
        TabControl1.SelectTab(1)
    End Sub

    Public Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click   'tab2 prev tab button, i.e. takes to tab1
        TabControl1.SelectTab(0)
    End Sub

    Private Sub Button16_Click(sender As Object, e As EventArgs) Handles Button16.Click     'tab2 next tab button, i.e. takes to tab3
        TabControl1.SelectTab(2)
    End Sub

    Private Sub Button15_Click(sender As Object, e As EventArgs) Handles Button15.Click     'tab3 prev tab button, i.e. takes to tab2
        TabControl1.SelectTab(1)
    End Sub

    Private Sub Button21_Click(sender As Object, e As EventArgs) Handles Button21.Click     'tab3 next tab button, i.e. takes to tab4
        TabControl1.SelectTab(3)
    End Sub

    Private Sub Button17_Click(sender As Object, e As EventArgs) Handles Button17.Click     'tab4 prev tab button, i.e. takes to tab3
        TabControl1.SelectTab(2)
    End Sub
End Class

Public Class GV         'Global Variables, other classes to written after Main Class else code wont work
    Public Shared CountN, csvFolderType As Integer         'countN to store total no. of rejected bases. , csvFolderType to save output .csv files type: whether in multiple folders depending on length of gene or in a single folder
    Public Shared MuX, MuY, gR, SumA, SumC, SumG, SumT As Double    'self explanatory
    Public Shared inputFolderPath, outputFolderPath, fileNameNoPath As String       'inputFolderPath to store input folder path, outputFolderPath to store output folder path, fileNameNoPath to store name of input txt file only (without its path)
    Public Shared inputExcelFile, outputExcelPath As String  'inputExcelFile to store edited CDS excel file as input, outputExcelPath to give output excel after operating on edited CDS excel file
    Public Shared Tab3InpFolderPath, Tab3OutFolderPath As String
    Public Shared Tab4InpFolderPath, Tab4OutFolderPath As String
    Public Shared tab5InpFolderPath, Tab5OutFolderPath As String
    Public Shared tab6InpFolderPath, Tab6OutFolderPath As String
    Public Shared tab7InpFolderPath, Tab7OutFolderPath As String

End Class