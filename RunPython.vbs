Function Main(inputs, outputFolder, params)
Dim WshShell, fso, loc, cmd, sFile, pyExe, inputFile, outFolder 
Const Quote = """"
 inputFile = inputs(0)
 WScript.Echo(Recode("Входящий файл XLS: " & inputFile, "cp866", "windows-1251"))
 outFolder = outputFolder
 WScript.Echo(Recode("Папка вывода: " & outFolder, "cp866", "windows-1251"))
 PyFile = params(1)
 

Set fso = CreateObject("Scripting.FileSystemObject")
loc = fso.GetAbsolutePathName(".")
sFile = "C:\Esko\bg_data_fastserverscrrunnt_v100\Scripts\WindowsScript\" & PyFile  & " " & Quote & inputFile & Quote & " " & Quote & outFolder & Quote 
WScript.Echo(Recode("sFile: " & sFile, "cp866", "windows-1251"))
pyExe = "C:\Python37\python.exe "
WScript.Echo pyExe + sFile



Set WshShell = CreateObject("WScript.Shell")
WshShell.Run (pyExe & sFile),,true

if WScript.Arguments.Count = 0 then
    loc = fso.GetAbsolutePathName(".")
else
    loc = WScript.Arguments(0)
end if




'Блок вывода в консоль файла log.txt
FileLog = outputFolder  & "\log.txt"


if  FSO.FileExists(FileLog)  Then
	Set File = FSO.OpenTextFile(FileLog, 1)
	Str1 = File.ReadAll
	File.Close
	WScript.Echo(Recode(Str1, "cp866", "windows-1251"))
else
  'Выводим информацию 
	WScript.Echo(Recode("Файл не логирования не найден по адресу " & FileLog, "cp866", "windows-1251"))
end if


End function
'Функция смены кодировки для вывода кирилицы в консоль'
			Function Recode(StrText, SrcCode, DestCode)
			    With CreateObject("ADODB.Stream")
			        .Type = 2
			        .Mode = 3
			        .Charset = DestCode
			        .Open
			        .WriteText (strText)
			        .Position = 0
			        .Charset = SrcCode
			        Recode = .ReadText
			        .Close
			    end with
			End Function


 Dim inputs()
 Dim outputFolder
 Dim params()

 Main inputs, outputFolder, params
 Main = "OK"