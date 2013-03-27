'closeMillennium.vbs checks every idleTime seconds to see if any instances of
'Millennium have been idle (based on a change in UserModeTime). If an instance
'has been idle between checks, it is closed. 
'
'Created by Brian Samek (brian.samek@gmail.com) on 11/28/2012.

'Interval in seconds to check if Millennium has been idle
idleTime = 1800

'Connect to WMI on local computer.
strComputer = "."
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\cimv2")

'Create the Dictionary object to keep track of PIDs and UserModeTime.
Set PIDandTime = CreateObject("Scripting.Dictionary")

'The main loop.
Do While True

    'Get all processes named java.exe. We will make sure they are
    'Millennium in a moment.
    Set Processes = objWMIService.ExecQuery _
        ("Select * from Win32_Process Where Name = 'java.exe'")

    For Each Process in Processes

        'Check if it's actually Millennium.
        If InStr(Process.ExecutablePath, "Millennium") > 0 Then

            'Get process's User Mode Time.
            processTime = CSng(Process.UserModeTime) / 10000000

            If PIDandTime.exists(Process.ProcessID) Then

                'If the time hasn't increased, close Millennium.
                If processTime < PIDandTime.Item(Process.ProcessID) + .1 Then
                    Process.Terminate

                'If the time has increased, put the new time in the
                'Dictionary.
                Else
                    PIDandTime.Item(Process.ProcessID) = processTime
                End If

            'If this process isn't in the Dictionary, add it.
            Else
                PIDandTime.add Process.ProcessID, processTime
            End If
        End If
    Next

    Wscript.Sleep(idleTime*1000)
Loop