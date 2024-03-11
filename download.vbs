' VBScript to download a file from the internet

Dim httpRequest, stream
Set httpRequest = CreateObject("MSXML2.XMLHTTP")

' Specify the URL of the file to download
Dim fileUrl
fileUrl = "https://github.com/GloAlgara/ucases.github.io/blob/main/calc.sh"

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
Else
    MsgBox "Failed to download the file. Status: " & httpRequest.Status
End If

Set httpRequest = Nothing
