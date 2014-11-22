' --------------------------------------------------------------
' ACP-Weather Script
' User Customized
' --------------------------------------------------------------
'
' *** Function IsProcessRunning - determines if specificed Windows process is running
Function IsProcessRunning( strServer, strProcess )
   Dim Process, strObject
   IsProcessRunning = False
   strObject = "winmgmts://" & strServer
   For Each Process in GetObject( strObject ).InstancesOf( "win32_process" )
      If UCase( Process.name ) = UCase( strProcess ) Then
         IsProcessRunning = True
         Exit Function
      End If
   Next
End Function
'
'
Sub Main()
   Dim FMX
   Console.PrintLine "Weather Safety... script initiated"
   On Error Resume Next                        ' Best efforts...
   '
   ' *** Close Dome
   If Dome.ScopeClearsClosedDome Then
       Console.PrintLine "...closing roof/shutter."
       Dome.CloseShutter                       ' Harmless if no dome/roof
   End If
   '
   ' *** If autofocus enabled and program running, interupt it...
   If Prefs.AutoFocus.Enabled Then
      If Prefs.AutoFocus.UsePWI Then
           Console.PrintLine "...halting PWI AutoFocus (if needed)"
           Set PWAF = CreateObject("PlaneWave.AutoFocus")
           PWAF.StopAutofocus
       Else
           Console.PrintLine "...halting FocusMax (if needed)"
           If isprocessrunning(".","FocusMax.exe") Then
              Console.Printline "...FocusMax running, checking for active operation"
              Set FMX = CreateObject("FocusMax.FocusControl")
              If FMX.FocusAsyncStatus = -1 Then   ' If FMx is actually doing something, stop it
                 Console.PrintLine "...halting active FocusMax operation"
                 FMX.Halt
              End If
           End If
           Set FMX = Nothing
           ' Console.PrintLine "...success."
       End If
       Util.WaitForMilliseconds 1000
   End If
   '
   ' *** Park Telescope 
   If Telescope.Connected Then
      Console.PrintLine "...scope is connect, parking now"
      Telescope.Park     ' Default setting in ACP, this parks and closes the dome/roof in that order
      Console.PrintLine "...scope is parked"
   End If
   '
   Console.PrintLine "Weather Safety...script completed successfully"
End Sub
