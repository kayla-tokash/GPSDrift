'-----------------------------------------------------------------------------------
'----------------------Main functions and subs for GPSDriftCast---------------------
'-----------------------------------------------------------------------------------
Option Explicit

Function GetWinds(latitude As Double, longitude As Double, hoffset() As Integer) As Variant
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.XMLHTTP")
    
    Dim url As String
    Dim i As Integer
    Dim response() As Variant
    
    ReDim response(LBound(hoffset) To UBound(hoffset))
    
    For i = LBound(hoffset) To UBound(hoffset)
        url = "https://windsaloft.us/winds.php?lat=" & latitude & "&lon=" & longitude & "&hourOffset=" & hoffset(i) & "&referrer=GPSDRIFTCAST"
        
        'Debug.Print url
        
        xmlhttp.Open "GET", url, False
        xmlhttp.send
        
        If xmlhttp.Status = 200 Then
            response(i) = xmlhttp.responseText
        Else
            response(i) = "Error: " & xmlhttp.statusText
        End If
        'Debug.Print response(i)
    Next i
    
    GetWinds = response
End Function


Function LaunchOffset(Ldate As String, Lstart As String) As Integer
' Function takes in the launch date and launch time
' Returns the offset in hours between the current date and time(now in DateDiff function) and the given launch date and start time

Dim LaunchDate As Date

Lstart = WorksheetFunction.Text(Lstart, " h AM/PM") ' extra space before formatting is to make formatting work with DateDiff function
Ldate = WorksheetFunction.Text(Ldate, "mm/dd/yyyy")
'Debug.Print Lstart

LaunchDate = Ldate & Lstart

LaunchOffset = DateDiff("h", now, LaunchDate)
'Debug.Print LaunchOffset

End Function


Function LinearInterpolate(x, x0, y0, x1, y1) As Double
    ' Linear interpolation formula: y = y0 + (x - x0) * ((y1 - y0) / (x1 - x0))
    If x1 - x0 = 0 Then
        LinearInterpolate = CVErr(xlErrDiv0)
    Else
        LinearInterpolate = y0 + (x - x0) * ((y1 - y0) / (x1 - x0))
    End If
End Function



Function Lat_Brg_Dist(StartLat, Distance, WindDir) As Double
Dim DistLat As Double
Dim r As Integer
Dim Bearing As Double
r = 6371 'Radius of the Earth, km
DistLat = Distance / 3280.84
Bearing = WindDir - 180
Lat_Brg_Dist = Round(WorksheetFunction.Degrees(WorksheetFunction.Asin(Sin(WorksheetFunction.Radians(StartLat)) * Cos(DistLat / r) + Cos(WorksheetFunction.Radians(StartLat)) * Sin(DistLat / r) * Cos(WorksheetFunction.Radians(Bearing)))), 6)
End Function

Function Lon_Brg_Dist(StartLon, StartLat, EndLat, Distance, WindDir) As Double
Dim DistLon As Double
Dim Bearing As Double
DistLon = Distance
Dim r As Integer
r = 6371 'Radius of the Earth, km
DistLon = Distance / 3280.84
Bearing = WindDir - 180

Lon_Brg_Dist = Round(WorksheetFunction.Degrees(WorksheetFunction.Radians(StartLon) + WorksheetFunction.Atan2(Cos(DistLon / r) - (Sin(WorksheetFunction.Radians(StartLat)) * Sin(WorksheetFunction.Radians(EndLat))), Sin(WorksheetFunction.Radians(Bearing)) * Sin(DistLon / r) * Cos(WorksheetFunction.Radians(StartLat)))), 6)

End Function

Function DriftDistance(StartAlt, EndAlt, WindSpeed, DescentRate) As Double
Dim TotalAlt As Double
Dim DescentTime As Double

TotalAlt = StartAlt - EndAlt ' TotalAlt is in ft
DescentTime = TotalAlt / CDbl(DescentRate) ' DescentTime is in seconds, Descent Rate is in ft/s

'Drift distance is in feet, if windspeed is mph, multiply windspeed by 1.4667
' if windspeed is in knots, multiply by 1.68781
' wind speed needs to be converted to ft/s from incoming value

DriftDistance = Round(DescentTime * (WindSpeed * 1.68781), 2) 'ft

End Function

Function GetDistBetween(StartLat, EndLat, StartLon, EndLon)

Dim r As Integer
r = 6371 ' Radius of Earth, km

GetDistBetween = Round(3280.84 * (WorksheetFunction.Acos(Sin(WorksheetFunction.Radians(StartLat)) * Sin(WorksheetFunction.Radians(EndLat)) + Cos(WorksheetFunction.Radians(StartLat)) * Cos(WorksheetFunction.Radians(EndLat)) * Cos(WorksheetFunction.Radians(EndLon) - WorksheetFunction.Radians(StartLon))) * r), 0)

End Function

Function ConvertToBoolean(ByVal value As String) As Boolean
    If UCase(value) = "YES" Then
        ConvertToBoolean = True
    ElseIf UCase(value) = "NO" Then
        ConvertToBoolean = False
    Else
        ' Handle invalid input (optional)
        MsgBox "Invalid value in named range. Expected 'Yes' or 'No'."
    End If
End Function

Function RecoveryToBoolean(ByVal value As String) As Boolean
    If UCase(value) = "MAIN ONLY" Then
        RecoveryToBoolean = True
    ElseIf UCase(value) = "DUAL DEPLOY" Then
        RecoveryToBoolean = False
    Else
        ' Handle invalid input (optional)
        MsgBox "Invalid value in named range. Expected 'Main Only' or 'Dual Deploy'."
    End If
End Function

Function CalculateLaunchTimes(startTime As String, numLaunches As Integer) As Variant
    Dim launchTimes() As String
    Dim dt As Date
    Dim i As Integer

    ' Convert start time to Date type
    dt = TimeValue(startTime)

    ' Resize the array to hold numLaunches - 1 elements
    ReDim launchTimes(0 To numLaunches - 1)
    launchTimes(0) = WorksheetFunction.Text(dt, "hAM/PM")
    ' Loop to calculate launch times
    For i = 1 To numLaunches - 1
        ' Add an hour to the time
        dt = dt + TimeSerial(1, 0, 0)
        ' Format the new time to "hAM/PM"
        launchTimes(i) = WorksheetFunction.Text(dt, "hAM/PM")
    Next i

    ' Return the array of launch times
    CalculateLaunchTimes = launchTimes
End Function

Sub GetWeatherCockData(ByRef arrWindSpeed() As Variant, ByRef arrWindDist() As Variant, ByRef arrApogee() As Variant, wsName As String, ByRef endProgramFlag As Boolean)
    Dim ws As Worksheet
    Dim rngWindSpeed As Range
    Dim rngWindDist As Range
    Dim rngApogee As Range
    Dim cell As Range
    Dim hasBlanks As Boolean

    ' Get the active worksheet
    Set ws = ThisWorkbook.Worksheets(wsName)

    ' Get the named range "WC_WindSpeed"
    On Error Resume Next
    Set rngWindSpeed = ws.Range("WC_WindSpeed")
    On Error GoTo 0

    ' Check if named range exists
    If Not rngWindSpeed Is Nothing Then
        ' Check for blank cells
        hasBlanks = False
        For Each cell In rngWindSpeed
            If IsEmpty(cell) Then
                hasBlanks = True
                Exit For
            End If
        Next cell
        ' Display message if blank cells found
        If hasBlanks Then
            MsgBox "WC_WindSpeed contains blank cells. Please ensure all cells are filled.", vbExclamation
            endProgramFlag = True ' Set flag to end the main program
            Exit Sub ' End this subroutine
        End If
        ' Store the values in the named range into a 1-dimensional array
        arrWindSpeed = Application.Transpose(rngWindSpeed.value)
    End If

    ' Get the named range "WC_WindDist"
    On Error Resume Next
    Set rngWindDist = ws.Range("WC_WindDist")
    On Error GoTo 0

    ' Check if named range exists
    If Not rngWindDist Is Nothing Then
        ' Check for blank cells
        hasBlanks = False
        For Each cell In rngWindDist
            If IsEmpty(cell) Then
                hasBlanks = True
                Exit For
            End If
        Next cell
        ' Display message if blank cells found
        If hasBlanks Then
            MsgBox "WC_WindDist contains blank cells. Please ensure all cells are filled.", vbExclamation
            endProgramFlag = True ' Set flag to end the main program
            Exit Sub ' End this subroutine
        End If
        ' Store the values in the named range into a 1-dimensional array
        arrWindDist = Application.Transpose(rngWindDist.value)
    End If

    ' Get the named range "WC_Apogee"
    On Error Resume Next
    Set rngApogee = ws.Range("WC_Apogee")
    On Error GoTo 0

    ' Check if named range exists
    If Not rngApogee Is Nothing Then
        ' Check for blank cells
        hasBlanks = False
        For Each cell In rngApogee
            If IsEmpty(cell) Then
                hasBlanks = True
                Exit For
            End If
        Next cell
        ' Display message if blank cells found
        If hasBlanks Then
            MsgBox "WC_Apogee contains blank cells. Please ensure all cells are filled.", vbExclamation
            endProgramFlag = True ' Set flag to end the main program
            Exit Sub ' End this subroutine
        End If
        ' Store the values in the named range into a 1-dimensional array
        arrApogee = Application.Transpose(rngApogee.value)
    End If
End Sub


Function RemoveDuplicatesFromArray(arr As Variant) As Variant
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    
    Dim result() As Variant
    Dim element As Variant
    Dim i As Long
    
    ' Remove duplicates
    For Each element In arr
        If Not dict.Exists(element) Then
            dict.Add element, Nothing
        End If
    Next element
    
    ' Resize result array
    ReDim result(0 To dict.Count - 1)
    
    ' Fill result array
    For Each element In dict.Keys
        result(i) = element
        i = i + 1
    Next element
    
    RemoveDuplicatesFromArray = result
End Function

Function GetLSData(targetName As String) As Variant
    Dim ws As Worksheet
    Dim LS_Arr(1 To 6) As Variant ' Array to store latitude, longitude, elevation, and waiver
    Dim rowIndex As Long
    Dim siteCell As Range
    Dim siteRange As Range
    Dim elevation As Double
    Dim LaunchLatitude As Double
    Dim LaunchLongitude As Double
    Dim WaiverRadius As Double
    Dim WaiverLat As Double
    Dim WaiverLon As Double
    Dim errorMessage As String
    Dim anyBlank As Boolean
    Dim i As Integer
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("GPS DriftCast")
    
    ' Find the target site in the SITE named range
    Set siteRange = ws.Range("SITE")
    Set siteCell = siteRange.Find(What:=targetName, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' If target site is found, get the elevation, latitude, and longitude
    If Not siteCell Is Nothing Then
        rowIndex = siteCell.Row - siteRange.Row + 1
        
        ' Check if any of the values are blank
        anyBlank = False
        If IsEmpty(ws.Range("LATITUDE").Cells(rowIndex, 1).value) Then
            errorMessage = errorMessage & "Launch Latitude blank for: " & targetName & vbCrLf
            anyBlank = True
        Else
            LaunchLatitude = CDbl(ws.Range("LATITUDE").Cells(rowIndex, 1).value)
            If LaunchLatitude < -90 Or LaunchLatitude > 90 Then
                errorMessage = errorMessage & "The provided latitude for site: " & targetName & " is outside the valid range of -90 to 90." & vbCrLf
                anyBlank = True
            End If
        End If
        
        If IsEmpty(ws.Range("LONGITUDE").Cells(rowIndex, 1).value) Then
            errorMessage = errorMessage & "Launch Longitude blank for: " & targetName & vbCrLf
            anyBlank = True
        Else
            LaunchLongitude = CDbl(ws.Range("LONGITUDE").Cells(rowIndex, 1).value)
            If LaunchLongitude < -180 Or LaunchLongitude > 180 Then
                errorMessage = errorMessage & "The provided longitude for site: " & targetName & " is outside the valid range of -180 to 180." & vbCrLf
                anyBlank = True
            End If
        End If
        
        ' If any value is blank or out of range, show error message, set LS_Arr(1) = 200, and write zeros for the rest of the arrays
        If anyBlank Then
            MsgBox errorMessage, vbExclamation, "Invalid Value(s) Found"
            LS_Arr(1) = 200
            For i = 2 To 6
                LS_Arr(i) = 0
            Next i
            GetLSData = LS_Arr
            Exit Function
        End If
        
        ' Get the waiver radius, latitude, and longitude values
        WaiverRadius = CDbl(ws.Range("WAIVER_RAD").Cells(rowIndex, 1).value) ' Updated variable name
        elevation = CDbl(ws.Range("ELEVATION").Cells(rowIndex, 1).value)
        WaiverLat = CDbl(ws.Range("WAIVER_LAT").Cells(rowIndex, 1).value)
        WaiverLon = CDbl(ws.Range("WAIVER_LON").Cells(rowIndex, 1).value)
        
        ' Store the values in the array
        LS_Arr(1) = LaunchLatitude
        LS_Arr(2) = LaunchLongitude
        LS_Arr(3) = elevation
        LS_Arr(4) = WaiverRadius
        LS_Arr(5) = WaiverLat
        LS_Arr(6) = WaiverLon
    Else
        ' If site is not found, set LS_Arr(1) = 100 and write zeros for the rest of the arrays
        LS_Arr(1) = 100
        For i = 2 To 6
            LS_Arr(i) = 0
        Next i
    End If
    
    ' Return the array
    GetLSData = LS_Arr
End Function



