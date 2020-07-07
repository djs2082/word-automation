' Dim strLine1
' Dim strLine2
' Dim strLine3
' Dim question
' Dim file_name
' question=""

' Set excelObj=CreateObject("Excel.Application")
' excelObj.DisplayAlerts = False
' excelFile="config.xlsx"
' Set wrkBookObj=excelObj.WorkBooks.open("C:\Users\Admin\Desktop\djs_auto\bin\" & excelFile)
' sheetName="Sheet1"
' Set sheetObj=wrkBookObj.Worksheets(sheetName)

Set wshShell = CreateObject( "WScript.Shell" )
path=wshShell.ExpandEnvironmentStrings( "%MYPATH%" )
MsgBox(path)
wshShell = Nothing

' Create object of MS Word
' Set objWord = CreateObject("Word.Application")
' objWord.Caption = "IML Practical"
' objWord.Visible = True
' Set objDoc = objWord.Documents.Add()
' Set objSelection = objWord.Selection
' WScript.Sleep 5000
' set shl = createobject("wscript.shell")
' shl.SendKeys "{ENTER}"
' prac_name=InputBox("Enter the practical Name:")
 ' Write The Practical Name in file
' objSelection.ParagraphFormat.Alignment = 1
' objSelection.Font.Name = sheetObj.Range("B12").value
' objSelection.Font.Size = sheetObj.Range("C12").value
' objSelection.TypeText "" & prac_name &""
' objSelection.TypeParagraph()


' Read the questions and sorce code file names from file questions.txt
' Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile(sheetObj.Range("B2").value,1)
' do while not objFileToRead.AtEndOfStream
' strLine1 = objFileToRead.ReadLine()
' i=0
' a=split(strLine1," ")
' b=UBound(a)
' file_name=a(b)
' For i=0 to b-1
' question=question & " " & a(i)
' Next

' Paste the question in file
' objSelection.ParagraphFormat.Alignment = 0
' objSelection.Font.Size = sheetObj.Range("C13").value
' objSelection.Font.Name = sheetObj.Range("B13").value
' objSelection.TypeText "" & question & ""
' objSelection.TypeParagraph()

' Paste the code of question
' Set objFileToRead2 = CreateObject("Scripting.FileSystemObject").OpenTextFile(sheetObj.Range("B3").value & "\" & file_name,1)
' do while not objFileToRead2.AtEndOfStream
' strLine2 = objFileToRead2.ReadLine()
' If Len(strLine2) > 0 Then
' objSelection.Font.Size = sheetObj.Range("C14").value
' objSelection.Font.Name = sheetObj.Range("B14").value
' objSelection.TypeText "" & strLine2 & ""
' objSelection.TypeParagraph()
' End If

' loop
' objFileToRead2.Close
' Set objFileToRead2 = Nothing

' Run the program
' SET oShell = WScript.CreateObject("Wscript.Shell")
' Dim source_code_path
' source_code_path = "C:\Users\Admin\Desktop\python\" & file_name 
' Dim currentCommand 
' currentCommand = "cmd /k " & Chr(34) & source_code_path & Chr(34)
' WScript.echo currentCommand
' oShell.run currentCommand,1,True



' s_name=Replace(file_name,".","_")
Set WshShell = CreateObject("WScript.Shell") 
' file_name=chr(34) & sheetObj.Range("B3").value &"\"& file_name & chr(34)
' g="abcd " & chr(34) & a & chr(34)

' start "cmd /K ss2.bat %1 %2 %3"
' WshShell.Run "cmd /k " & sheetObj.Range("B4").value & "\ss3.bat " & s_name & " " & file_name & " " & sheetObj.Range("B7").value, 1, true

' start "" /MAX "cmd /K ss2.bat %1 %2 %3"
' MsgBox(chr(34) & chr(34)&" /MAX " & chr(34) & "cmd /k " & "ss2.bat")
' cs="start " & chr(34) & chr(34)&" /MAX " & chr(34) & "cmd /k "
' MsgBox(cs)
 ' WshShell.Run cs & " ss2.bat", 1, true
' Set WshShell = Nothing

' Set objShape = objDoc.Shapes

' src=sheetObj.Range("B7").value&"\" & s_name & ".jpg"
' MsgBox(src)
' Set objShape = objSelection.InlineShapes.AddPicture(src)
' objSelection.TypeParagraph
' objShape.AddPicture(src)
' WScript.Sleep 2000
' oShell.run "cmd.exe /C copy ""S:\Claims\Sound.wav"" ""C:\WINDOWS\Media\Sound.wav"" "

' Set oShell = WScript.CreateObject ("WScript.Shell")
' oShell.run "cmd.exe /C nircmd.exe screenshot dj.jpg" "C:\WINDOWS\Media\Sound.wav"
' Set oShell = 'Nothing'

' question=""

' loop

' myErr = objDoc.SpellingErrors.Count 
' If myErr = 0 Then 
 ' Msgbox "No spelling errors found." 
' Else 
 ' Msgbox myErr & " spelling errors found." 
' End If
' objFileToRead.Close
' Set objFileToRead = Nothing
' objDoc.SaveAs(sheetObj.Range("B5").value&"\testdoc.doc")
' objWord.Quit
' excelObj.DisplayAlerts = False
' wrkBookObj.Save
' wrkBookObj.Close True
' excelObj.Quit
' WScript.Quit