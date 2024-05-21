'-----------------------------------------------------------------------------------
'-----------------------------GPS DriftCast Main Program----------------------------
'-----------------------------------------------------------------------------------
'Dave Snyder, 5/4/2024

Option Explicit

Sub main()

Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim z As Integer

Dim hoffset() As Integer
Dim LaunchLat As Double
Dim LaunchLon As Double
Dim LaunchDate As String
Dim LaunchStart As String
Dim numLaunches As Integer

Dim responses As Variant
Dim stringResponses() As String
Dim model() As String


Dim windSpdTerm As String
Dim windDirTerm As String
Dim Altitude() As Long
Dim WindDir() As Long
Dim WindSpeed() As Long
Dim insertPosition() As Long
Dim newValue As Long
Dim PosPlaceHolder As Long
Dim events() As Long
Dim descents(1) As Double
Dim DriftDist() As Variant
Dim DriftLat() As Variant
Dim DriftLon() As Variant
Dim DriftAlt() As Variant
Dim MainOnlyFlag As Boolean
Dim MainOnlyString As String

Dim launchTimes() As String
Dim WeatherCockTableFlag As Boolean
Dim WeatherCockString As String
Dim ref As Integer
Dim ApoLat As Double
Dim ApoLon As Double
Dim ws As Worksheet
Dim rowNum As Integer
Dim distBtwn As Double
Dim HighWindWarning As Boolean
Dim name As String
Dim landLat() As Variant
Dim landLon() As Variant
Dim landTags() As Variant
Dim initAlt() As Variant
Dim initWindSpd() As Variant
Dim initWindDir() As Variant
Dim LSelev As Double
Dim WaiverRadius As Double
Dim WaiverLat As Double
Dim WaiverLon As Double
Dim Colors As Variant
Dim FN As String
Dim FP As String
Dim FN_LandingScatter As String
Dim FP_LandingScatter As String
Dim FN_FlightScatter As String
Dim FP_FlightScatter As String
Dim endProgramFlag As Boolean
Dim LaunchSite As String
Dim LS_Array() As Variant
Dim SiteFlag As Boolean
endProgramFlag = False


    
' Get file location and file name from named cells for output KML files
FP = ThisWorkbook.Names("FileLocation").RefersToRange.value
FN = ThisWorkbook.Names("FileName").RefersToRange.value

'Clean up a "\" at end of file path if needed
    If Right(FP, 1) = "\" Then
        FP = Left(FP, Len(FP) - 1)
    End If
           
'Check to see if FP is valid
    If Dir(FP, vbDirectory) = "" Then
        MsgBox "Invalid folder location: " & FP, vbExclamation, "Error"
        End
    End If
   
' Append tags for output files
FN_FlightScatter = FN & "_" & "FlightScatter"
FP_FlightScatter = FP & "\" & FN_FlightScatter & ".kml"


FN_LandingScatter = FN & "_" & "LandingScatter"
FP_LandingScatter = FP & "\" & FN_LandingScatter & ".kml"

'Debug.Print FP

' Get User Input from Excel Sheet
LaunchSite = Range("LaunchSite").value
LaunchDate = CStr(Range("LaunchDate").value)
LaunchStart = CStr(Range("LaunchStart").value)
numLaunches = Range("NumLaunches").value

MainOnlyString = Range("MainOnlyFlag").value
MainOnlyFlag = RecoveryToBoolean(MainOnlyString)

WeatherCockString = Range("WeatherCockString").value
WeatherCockTableFlag = ConvertToBoolean(WeatherCockString)

'Get Launch Site information from Data Table dependent on user site drop down selection
LS_Array = GetLSData(LaunchSite)

'Check to see if there are errors with user input of Laucn Site Table

If LS_Array(1) = 100 Then
    
    MsgBox ("Site not found")
    End

ElseIf LS_Array(1) = 200 Then
    
    End
    
Else

    LaunchLat = LS_Array(1)
    LaunchLon = LS_Array(2)
    LSelev = LS_Array(3)
    WaiverRadius = LS_Array(4)
    WaiverLat = LS_Array(5)
    WaiverLon = LS_Array(6)
End If

If WaiverRadius = 0 Then
WaiverLat = LaunchLat
WaiverLon = LaunchLon
End If


'Set Name for sheet to be used for pointing and filling at a glance cells
name = "GPS DriftCast"
Set ws = ThisWorkbook.Worksheets(name)

'Set colors for KML Files
Colors = Array( _
    "ff0000ff", "ffffffff", "ffffff00", "ff00ff00", _
    "ff00ffff", "ffff00ff", "ff808080", "ffffa500", _
    "ff000000", "ff008000", "ff0000ff", "ff800080" _
)


windSpdTerm = "WindSpeed"
windDirTerm = "WindDirection"

' Error handling for inputs for number of launches exceeding 12
If numLaunches > 12 Then
MsgBox ("Launch Window cannot be more than 12 hours")
End
End If


If MainOnlyFlag = True Then
    ReDim events(0) As Long
    events(0) = CLng(Range("ApoEvent").value)
     
    Else
    ReDim events(1) As Long
    events(0) = CLng(Range("ApoEvent").value)
    events(1) = CLng(Range("MainEvent").value)

End If

descents(0) = CDbl(Range("DrogueDescentRate").value)
descents(1) = CDbl(Range("MainDescentRate").value)


'Initialize hoffset array for each of the forecast calls
ReDim hoffset(numLaunches - 1)
For i = 0 To UBound(hoffset)

 If i = 0 Then
    'Get first hoffoset
  hoffset(i) = LaunchOffset(LaunchDate, LaunchStart)
    
      If hoffset(i) > 380 Then
        MsgBox ("Cannot forecast more than 380 hours into the future")
        End
      ElseIf hoffset(i) < -24 Then
        MsgBox ("Cannot forecast more than 24 hours into the past.")
        End
      End If
 
    
Else
  'On progressive loops, calculate subsequent hoffsets from first
  hoffset(i) = hoffset(i - 1) + 1
  
      If hoffset(i) > 380 Then
        MsgBox ("Cannot forecast more than 380 hours into the future.")
        End
      ElseIf hoffset(i) < -24 Then
        MsgBox ("Cannot forecast more than 24 hours into the past.")
        End
     End If
End If
Next i

'For i = 0 To UBound(hoffset)
'Debug.Print hoffset(i)
'Next i

    ' Call the GetWinds function to get the forecasts from Winds Aloft website
    responses = GetWinds(LaunchLat, LaunchLon, hoffset)
    
    ' Cast responses to an array of strings
    ' To condition for use in the next step
    
    ReDim stringResponses(LBound(responses) To UBound(responses))
    ReDim model(LBound(responses) To UBound(responses))
    For i = LBound(responses) To UBound(responses)
        stringResponses(i) = CStr(responses(i))
        model(i) = GetModel(stringResponses(i))
        If model(i) = "Open-Meteo" And events(0) > 40000 Then
            MsgBox ("Your apogee is too high for the forecast Open-Meteo. Apogee of 40,000ft AGL or less supported by Open-Meteo at this time")
            End
        End If
            
    Next i
  'Debug.Print model(0)
        ' Print the string responses
   ' For i = LBound(stringResponses) To UBound(stringResponses)
   '     Debug.Print "hoffset(" & i & "): " & stringResponses(i)
   ' Next i

        ' Output the response array to debug sheet
        ThisWorkbook.Sheets("Debug").Columns("A").ClearContents
        
        For i = LBound(responses) To UBound(responses)
            ThisWorkbook.Sheets("Debug").Cells(i + 1, 1).value = responses(i)
         Next i
 
     
    ReDim LandingArray(0 To UBound(stringResponses))
     
    ' Calculate launch times for launch window
    ReDim launchTimes(0 To numLaunches - 1) As String
    launchTimes = CalculateLaunchTimes(LaunchStart, numLaunches)

    ' Display the launch times
    'For i = LBound(launchTimes) To UBound(launchTimes)
    '    Debug.Print "Launch " & i + 1 & ": " & launchTimes(i)
    'Next i
 
' Clear Contents of Output for at a glance
' F6 to I17

Set ws = ThisWorkbook.Worksheets(name)
ws.Range("E7:M18").ClearContents

rowNum = 7  ' Starting row for at a glance output

'Redimension variables to proper size based on user input
ReDim landLat(UBound(stringResponses) + 1)
ReDim landLon(UBound(stringResponses) + 1)
ReDim landTags(UBound(stringResponses) + 1)

'Initialize Flight Scatter Plot KML File

InitializeKMLFile FP_FlightScatter


'--------------------------------------------------------------------------
'------------------------------MAIN LOOP-----------------------------------
'--------------------------------------------------------------------------
'--------------------------------------------------------------------------


For k = 0 To UBound(stringResponses) 'Run the main loop for the number of forecasts requested

 
  initAlt = GetAltitudeArray(stringResponses(k), model(k))
  initWindDir = ExtractValues(stringResponses(k), model(k), windDirTerm)
  initWindSpd = ExtractValues(stringResponses(k), model(k), windSpdTerm)
  


  'Sometimes, altitude array from forecasts has duplicate data which causes problems because alt, wind dir, wind speed arrays
  ' are not the same size, so we clean them up quick
  
  initAlt = RemoveDuplicatesFromArray(initAlt)
  
  'For i = 0 To UBound(initAlt)
  'Debug.Print initAlt(i), initWindDir(i), initWindSpd(i)
  'Next i
  
'--------------------------------------------------------------------------
'Weathercock Set up

Dim q As Integer
Dim AvgWCDir As Double
Dim AvgWCSpd As Double
Dim AvgSurfaceWindDir As Double
AvgWCDir = 0
AvgWCSpd = 0

Dim b As Integer
Dim WCDist As Double

Dim WCwindSpeed() As Variant
Dim WCwindDist() As Variant
Dim TableApogee() As Variant
Dim Apogee As Double

If model(k) = "RAP" And events(0) <= 1000 Then
    q = 1
ElseIf model(k) = "RAP" And events(0) > 1000 Then
    q = 3
ElseIf model(k) = "Open-Meteo" Then
    q = 0
End If


If WeatherCockTableFlag = False Then
    ApoLat = LaunchLat
    ApoLon = LaunchLon
    WCDist = 0
       
For b = UBound(initAlt) To UBound(initAlt) - q Step -1
AvgWCDir = initWindDir(b) + AvgWCDir
AvgWCSpd = initWindSpd(b) + AvgWCSpd ' WindSpeed in knots
'Debug.Print Altitude(b); WindDir(b); WindSpeed(b))

Next b
AvgWCDir = (AvgWCDir / (q + 1))

AvgWCSpd = Round((AvgWCSpd / (q + 1)) * 1.15078, 0) ' Get Average Wind Speed during portion of ascent, convert to MPH to compare with User Table
AvgSurfaceWindDir = AvgWCDir

If AvgWCSpd >= 20 Then
HighWindWarning = True
End If

Else

For b = UBound(initAlt) To UBound(initAlt) - q Step -1
AvgWCDir = initWindDir(b) + AvgWCDir
AvgWCSpd = initWindSpd(b) + AvgWCSpd ' WindSpeed in knots
'Debug.Print Altitude(b); WindDir(b); WindSpeed(b)
  
Next b

AvgWCDir = (AvgWCDir / (q + 1)) ' Get Average Wind Direction during portion of ascent
AvgWCSpd = Round((AvgWCSpd / (q + 1)) * 1.15078, 0) ' Get Average Wind Speed during portion of ascent, convert to MPH to compare with User Table

AvgSurfaceWindDir = AvgWCDir

'Debug.Print AvgWCDir
'Debug.Print "Avg WS", AvgWCSpd

'Invert weathercock direction so it can be used to move the rocket upwind
If AvgWCDir <= 180 Then
AvgWCDir = AvgWCDir + 180
Else
AvgWCDir = AvgWCDir - 180
End If

If AvgWCSpd >= 20 Then
HighWindWarning = True
End If
    ' Call subroutine to get the weathercock distance from WC Wind speed and distance table in sheet
    
    Call GetWeatherCockData(WCwindSpeed, WCwindDist, TableApogee, name, endProgramFlag)
        ' Check if the flag is set to end the program
    If endProgramFlag Then
        MsgBox "Program ended.", vbInformation
        Exit Sub
    End If
    
        For i = LBound(WCwindSpeed) To UBound(WCwindSpeed)
                
            If WCwindSpeed(i) >= AvgWCSpd Then
                
                WCDist = LinearInterpolate(AvgWCSpd, WCwindSpeed(i - 1), WCwindDist(i - 1), WCwindSpeed(i), WCwindDist(i))
                events(0) = LinearInterpolate(AvgWCSpd, WCwindSpeed(i - 1), TableApogee(i - 1), WCwindSpeed(i), TableApogee(i))
                
                Exit For
            ElseIf i = UBound(WCwindSpeed) Then
                WCDist = WCwindDist(UBound(WCwindDist))
                events(0) = TableApogee(UBound(TableApogee))
                
                                
            End If
        Next i

'Debug.Print WCDist


ApoLat = Lat_Brg_Dist(LaunchLat, WCDist, AvgWCDir)
ApoLon = Lon_Brg_Dist(LaunchLon, LaunchLat, ApoLat, WCDist, AvgWCDir)

End If




'End Weathercock set up
'--------------------------------------------------------------------------
  

        
   ' Initialize the arrays with the initial values
    ReDim Altitude(LBound(initAlt) To UBound(initAlt))
    ReDim WindDir(LBound(initWindDir) To UBound(initWindDir))
    ReDim WindSpeed(LBound(initWindSpd) To UBound(initWindSpd))
    ReDim insertPosition(0 To UBound(events))
    
    i = 0 'reset i back to 0
    
    For i = LBound(initAlt) To UBound(initAlt)
        Altitude(i) = initAlt(i)
    Next i
    
    For i = LBound(initWindDir) To UBound(initWindDir)
        WindDir(i) = initWindDir(i)
    Next i

    For i = LBound(initWindSpd) To UBound(initWindSpd)
        WindSpeed(i) = initWindSpd(i)
    Next i
'Debug.Print UBound(initAlt), UBound(initWindDir), UBound(initWindSpd)



'Insert apogee and main event into altitude arrays
    ' Loop through the new values array
    For i = LBound(events) To UBound(events)
        newValue = events(i)
        ' Find the position to insert the new value
        For j = LBound(Altitude) To UBound(Altitude)
            If Altitude(j) < newValue Then
                insertPosition(i) = j
                Exit For
            End If
        Next j

        ' Resize the arrays to accommodate altitude and main events
        ReDim Preserve Altitude(LBound(Altitude) To UBound(Altitude) + 1)
        ReDim Preserve WindDir(LBound(WindDir) To UBound(WindDir) + 1)
        ReDim Preserve WindSpeed(LBound(WindSpeed) To UBound(WindSpeed) + 1)

        ' Shift elements to make space for the altitude and main event
        For j = UBound(Altitude) To insertPosition(i) + 1 Step -1
            Altitude(j) = Altitude(j - 1)
            WindDir(j) = WindDir(j - 1)
            WindSpeed(j) = WindSpeed(j - 1)
        Next j

        ' Insert Altitude and Main event into arrays, Linear interpolate Wind speed and direction at these values for calcs
        Altitude(insertPosition(i)) = newValue
        WindDir(insertPosition(i)) = LinearInterpolate(Altitude(insertPosition(i)), Altitude(insertPosition(i) - 1), WindDir(insertPosition(i) - 1), Altitude(insertPosition(i) + 1), WindDir(insertPosition(i) + 1))
        WindSpeed(insertPosition(i)) = LinearInterpolate(Altitude(insertPosition(i)), Altitude(insertPosition(i) - 1), WindSpeed(insertPosition(i) - 1), Altitude(insertPosition(i) + 1), WindSpeed(insertPosition(i) + 1))
        
        
    Next i

    ReDim DriftAlt(0 To (UBound(Altitude) - insertPosition(0)) + 1)
    ReDim DriftDist(UBound(DriftAlt))
    ReDim DriftLat(UBound(DriftAlt))
    ReDim DriftLon(UBound(DriftAlt))
      
    z = 0
    z = insertPosition(0)
    For i = 1 To UBound(DriftAlt) ' start at 1, leave 0 blank for launch site information, 1 position is apogee information
        
        DriftAlt(i) = Altitude(z) ' Flip table
        
        z = z + 1
    Next i
     


'Set up drift data tables - Constants
'------------------------
'0 position, ground level launch site
DriftDist(0) = CDbl(0#)
DriftLat(0) = LaunchLat
DriftLon(0) = LaunchLon
DriftAlt(0) = CDbl(0#)
'Debug.Print "Distance       Altitude      Latitude     Longitude"
'Debug.Print DriftDist(0), DriftAlt(0), DriftLat(0), DriftLon(0)

'1 position is constant, apogee data

   
DriftDist(1) = CDbl(0#)
DriftLat(1) = ApoLat  'LaunchLat  'Later this could be ApoLat with weathercock correction
DriftLon(1) = ApoLon  'LaunchLon  'Later this could be ApoLon with weathercock correction

'Debug.Print DriftDist(1), DriftAlt(1), DriftLat(1), DriftLon(1)
ref = insertPosition(0) ' this variable sets the location for the needed information within WindDir() and WindSpeed() arrays, 0 position should be apogee position
ref = ref + 1

If MainOnlyFlag = True Then
    For i = 2 To UBound(DriftAlt)
        
        DriftDist(i) = DriftDistance(DriftAlt(i - 1), DriftAlt(i), (WindSpeed(ref) + WindSpeed(ref - 1)) / 2, descents(1))
        DriftLat(i) = Lat_Brg_Dist(DriftLat(i - 1), DriftDist(i), (WindDir(ref) + WindDir(ref - 1)) / 2)
        DriftLon(i) = Lon_Brg_Dist(DriftLon(i - 1), DriftLat(i - 1), DriftLat(i), DriftDist(i), (WindDir(ref) + WindDir(ref - 1)) / 2)
        'Debug.Print DriftDist(i), DriftAlt(i), DriftLat(i), DriftLon(i)
        ref = ref + 1
        
    Next i

    
Else
    For i = 2 To UBound(DriftAlt)
            
        If DriftAlt(i) <= events(0) Then
            DriftDist(i) = DriftDistance(DriftAlt(i - 1), DriftAlt(i), (WindSpeed(ref) + WindSpeed(ref - 1)) / 2, descents(0))
            DriftLat(i) = Lat_Brg_Dist(DriftLat(i - 1), DriftDist(i), (WindDir(ref) + WindDir(ref - 1)) / 2)
            DriftLon(i) = Lon_Brg_Dist(DriftLon(i - 1), DriftLat(i - 1), DriftLat(i), DriftDist(i), (WindDir(ref) + WindDir(ref - 1)) / 2)
            'Debug.Print DriftDist(i), DriftAlt(i), DriftLat(i), DriftLon(i)
            ref = ref + 1
            
        Else
            DriftDist(i) = DriftDistance(DriftAlt(i - 1), DriftAlt(i), (WindSpeed(ref) + WindSpeed(ref - 1)) / 2, descents(1))
            DriftLat(i) = Lat_Brg_Dist(DriftLat(i - 1), DriftDist(i), (WindDir(ref) + WindDir(ref - 1)) / 2)
            DriftLon(i) = Lon_Brg_Dist(DriftLon(i - 1), DriftLat(i - 1), DriftLat(i), DriftDist(i), (WindDir(ref) + WindDir(ref - 1)) / 2)
            'Debug.Print DriftDist(i), DriftAlt(i), DriftLat(i), DriftLon(i)
            ref = ref + 1
            
        End If
        
    Next i
    
End If

    
    landLat(k + 1) = DriftLat(UBound(DriftAlt))
    landLon(k + 1) = DriftLon(UBound(DriftAlt))
   
    'Write each Fligh Path to single KML File
    
    AddKMLPath DriftLat, DriftLon, DriftAlt, launchTimes(k), Colors(k), LSelev, FP_FlightScatter, WaiverRadius, WaiverLat, WaiverLon
    
    
    distBtwn = GetDistBetween(LaunchLat, DriftLat(UBound(DriftAlt)), LaunchLon, DriftLon(UBound(DriftAlt)))
    Set ws = ThisWorkbook.Worksheets(name)
    'Write Data to Output Cells F7 to J18 available
    ws.Range("E" & rowNum).value = launchTimes(k)
    ws.Range("F" & rowNum).value = model(k)
    ws.Range("G" & rowNum).value = Round(AvgWCSpd, 0)
    ws.Range("H" & rowNum).value = Round(AvgSurfaceWindDir, 0)
    ws.Range("I" & rowNum).value = Round(events(0), 0)
    ws.Range("J" & rowNum).value = Round(WCDist, 0)
    ws.Range("K" & rowNum).value = DriftLat(UBound(DriftAlt))
    ws.Range("L" & rowNum).value = DriftLon(UBound(DriftAlt))
    ws.Range("M" & rowNum).value = distBtwn
    
    
    rowNum = rowNum + 1
    


Next k

'Finalze Flight Path KML File
FinalizeKMLFile FP_FlightScatter

'For i = 0 To UBound(stringResponses)
'Debug.Print landLat(i), landLon(i)
'Next i
For i = 0 To UBound(landTags)
    If i = 0 Then
        landTags(0) = "Launch Site"
    Else
    landTags(i) = launchTimes(i - 1)
    End If
   
Next i


landLat(0) = LaunchLat
landLon(0) = LaunchLon

'For i = 0 To UBound(DriftAlt)
'Debug.Print Round((DriftAlt(i) + 833) / 3.281, 2), DriftLat(i), DriftLon(i)
'Next i

GenerateLandingScatterKML FP_LandingScatter, landLat(), landLon(), landTags(), WaiverRadius, WaiverLat, WaiverLon

If HighWindWarning = True Then
MsgBox ("Some surface level winds are 20MPH or above. Consider flying at a different time or day" & vbCrLf & "Apogee and weathercock distance values derived from max wind speed for those sims" & vbCrLf & "" & vbCrLf & "Your sims were saved to the file path specified")

Else

MsgBox ("All surface level winds were under 20MPH" & vbCrLf & "Your sims were saved to the file path specified")
End If


AskToOpenKMLFile FP


End Sub





