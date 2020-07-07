
Dim strLine1
Dim strLine2
Dim strLine3
Dim question
Dim file_name
question=""
'Create object of MS Word
Set objWord = CreateObject("Word.Application")
objWord.Caption = "IML Practical"
objWord.Visible = True
Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection
prac_name=InputBox("Enter the practical Name:")
WScript.Sleep 1000
set shl = createobject("wscript.shell")
shl.SendKeys "{ENTER}"
 'Get the Practical Name from User


'Write The Practical Name in file
objSelection.ParagraphFormat.Alignment = 1
objSelection.Font.Name = "Arial Black"
objSelection.Font.Size = "24"
objSelection.TypeText "" & prac_name &""
objSelection.TypeParagraph()


'Read the questions and sorce code file names from file questions.txt
Set objFileToRead = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\Admin\Desktop\questions.txt",1)
do while not objFileToRead.AtEndOfStream
strLine1 = objFileToRead.ReadLine()


i=0
a=split(strLine1," ")
b=UBound(a)

file_name=a(b)

For i=0 to b-1
question=question & " " & a(i)
Next

'Paste the question in file
objSelection.ParagraphFormat.Alignment = 0
objSelection.Font.Size = "20"
objSelection.Font.Name = "BahnSchrift SemiBold"
objSelection.TypeText "" & question & ""
objSelection.TypeParagraph()

'Paste the code of question
Set objFileToRead2 = CreateObject("Scripting.FileSystemObject").OpenTextFile("C:\Users\Admin\Desktop\python\" & file_name,1)


do while not objFileToRead2.AtEndOfStream
strLine2 = objFileToRead2.ReadLine()
If Len(strLine2) > 0 Then
objSelection.Font.Size = "14"
objSelection.Font.Name = "Calibri"
objSelection.TypeText "" & strLine2 & ""
objSelection.TypeParagraph()
End If

loop
objFileToRead2.Close
Set objFileToRead2 = Nothing

'Run the program
' SET oShell = WScript.CreateObject("Wscript.Shell")
' Dim source_code_path
' source_code_path = "C:\Users\Admin\Desktop\python\" & file_name 
' Dim currentCommand 
' currentCommand = "cmd /k " & Chr(34) & source_code_path & Chr(34)
' WScript.echo currentCommand
' oShell.run currentCommand,1,True



s_name=Replace(file_name,".","_")

Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run "cmd /k " & chr(34) & "ss3.bat " & s_name & Chr(34), 1, true
Set WshShell = Nothing

Set objShape = objDoc.Shapes

src="C:\Users\Admin\Desktop\" & s_name & ".jpg"
MsgBox(src)
Set objShape = objSelection.InlineShapes.AddPicture(src)
objSelection.TypeParagraph
' objShape.AddPicture("C:\Users\Admin\Desktop\dj.png")
' WScript.Sleep 2000
' oShell.run "cmd.exe /C copy ""S:\Claims\Sound.wav"" ""C:\WINDOWS\Media\Sound.wav"" "

' Set oShell = WScript.CreateObject ("WScript.Shell")
' oShell.run "cmd.exe /C nircmd.exe screenshot dj.jpg" "C:\WINDOWS\Media\Sound.wav"
' Set oShell = 'Nothing'

question=""

loop

myErr = objDoc.SpellingErrors.Count 
If myErr = 0 Then 
 Msgbox "No spelling errors found." 
Else 
 Msgbox myErr & " spelling errors found." 
End If
objFileToRead.Close
Set objFileToRead = Nothing
objDoc.SaveAs("C:\Users\Admin\Desktop\testdoc.doc")
objWord.Quit
