'-----------------------------------------------------------------------------------
'----------------------------File Handling subroutines------------------------------
'-----------------------------------------------------------------------------------

'-----------------------------------------------------------------------------------
'----------------Subs/Functions to open KML File upon Completion--------------------
'-----------------------------------------------------------------------------------
Option Explicit

Dim CircleDrawn As Boolean

#If VBA7 Then
    Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As LongPtr, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As LongPtr
#Else
    Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Sub OpenKML_FilePrompt(initialDir As String)
    Dim chosenFile As Variant
     
    ' Change the current drive and directory to the initialDir
    ChDrive Left(initialDir, 1)
    ChDir initialDir
    
    ' Prompt the user to choose a file
    chosenFile = Application.GetOpenFilename("KML Files (*.kml), *.kml", , "Select KML file to open", , False)
    
    ' Check if the user selected a file
    If Not IsEmpty(chosenFile) Then
        ' Open the selected file
        ShellExecute 0, "open", chosenFile, vbNullString, vbNullString, vbNormalFocus
    End If
End Sub


Sub AskToOpenKMLFile(FILE_PATH As String)
    Dim response As Integer

    ' Ask the user if they want to open a file
    response = MsgBox("Do you want to open one of the KML Scatter Plot files?", vbYesNo, "Open File")

    ' Check the user's response
    If response = vbYes Then
    
            OpenKML_FilePrompt FILE_PATH
            
    Else
        MsgBox "File opening canceled"
    End If
End Sub

'-----------------------------------------------------------------------------------
'--------------Subroutines for writing landing scatter plot KML File----------------
'-----------------------------------------------------------------------------------
Sub GenerateLandingScatterKML(FilePath As String, latitudes As Variant, longitudes As Variant, tags As Variant, WaiverRadius_NM As Double, WaiverLat As Double, WaiverLon As Double)
    Dim fso As Object
    Dim ts As Object
    Dim i As Long
    Dim WaiverName As String
    
    
    If latitudes(0) = WaiverLat Then
        WaiverName = ""
    Else
        WaiverName = "Waiver Center"
    End If
    
    
    

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.CreateTextFile(FilePath, True)

    ' Convert WaiverRadius from feet to kilometers
    Dim WaiverRadiusKm As Double
    WaiverRadiusKm = WaiverRadius_NM * 0.539957 ' Convert nautical miles to KM for plot

    ' Write KML header
    ts.WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
    ts.WriteLine "<kml xmlns=""http://www.opengis.net/kml/2.2"">"
    ts.WriteLine "<Document>"

    ' Define style for the first tag (red color)
    ts.WriteLine "<Style id=""redStyle"">"
    ts.WriteLine "<IconStyle>"
    ts.WriteLine "<color>ff0000ff</color>"
    ts.WriteLine "</IconStyle>"
    ts.WriteLine "</Style>"

    ' Loop through each point and write Placemark
    For i = LBound(latitudes) To UBound(latitudes)
        ts.WriteLine "<Placemark>"
        If i = 0 Then
            ts.WriteLine "<styleUrl>#redStyle</styleUrl>"
        End If
        ts.WriteLine "<name>" & tags(i) & "</name>"
        ts.WriteLine "<Point>"
        ts.WriteLine "<coordinates>" & longitudes(i) & "," & latitudes(i) & ",0</coordinates>"
        ts.WriteLine "</Point>"
        ts.WriteLine "</Placemark>"
    Next i

    ' Write the placemark for the Waiver Center
    ts.WriteLine "<Placemark>"
    ts.WriteLine "<name>" & WaiverName & "</name>"
    ts.WriteLine "<Style>"
    ts.WriteLine "<IconStyle>"
    ts.WriteLine "<color>ff0000ff</color>"
    ts.WriteLine "</IconStyle>"
    ts.WriteLine "</Style>"
    ts.WriteLine "<Point>"
    ts.WriteLine "<altitudeMode>clampToGround</altitudeMode>"
    ts.WriteLine "<coordinates>" & WaiverLon & "," & WaiverLat & ",0</coordinates>"
    ts.WriteLine "</Point>"
    ts.WriteLine "</Placemark>"

    ' Plot circle clamped to the ground
    ts.WriteLine "<Placemark>"
    ts.WriteLine "<name>Waiver Radius</name>"
    ts.WriteLine "<Style>"
    ts.WriteLine "<LineStyle>"
    ts.WriteLine "<color>ff0000ff</color>"
    ts.WriteLine "<width>2</width>"
    ts.WriteLine "</LineStyle>"
    ts.WriteLine "<PolyStyle>"
    ts.WriteLine "<color>1aff0000</color>" ' Red color with some transparency
    ts.WriteLine "</PolyStyle>"
    ts.WriteLine "</Style>"
    ts.WriteLine "<Polygon>"
    ts.WriteLine "<outerBoundaryIs>"
    ts.WriteLine "<LinearRing>"
    ts.WriteLine "<coordinates>"
    
    Dim centerLat As Double
    Dim centerLon As Double
    centerLat = WaiverLat
    centerLon = WaiverLon
    
    Dim j As Integer
    For j = 0 To 360 Step 10
        Dim circleLat As Double
        Dim circleLon As Double
        circleLat = centerLat + (WaiverRadiusKm / 111.32) * Sin(j * 3.14159265358979 / 180)
        circleLon = centerLon + (WaiverRadiusKm / (111.32 * Cos(centerLat * 3.14159265358979 / 180))) * Cos(j * 3.14159265358979 / 180)
        ts.WriteLine circleLon & "," & circleLat & ",0"
    Next j
    
    ts.WriteLine "</coordinates>"
    ts.WriteLine "</LinearRing>"
    ts.WriteLine "</outerBoundaryIs>"
    ts.WriteLine "</Polygon>"
    ts.WriteLine "</Placemark>"

    ' Write KML footer
    ts.WriteLine "</Document>"
    ts.WriteLine "</kml>"

    ts.Close
    Set ts = Nothing
    Set fso = Nothing
    CircleDrawn = False

End Sub




'-----------------------------------------------------------------------------------
'------------------Subroutines for writing Flight Scatter KML Files-----------------
'-----------------------------------------------------------------------------------


Sub AddKMLPath(ByVal latitude As Variant, ByVal longitude As Variant, ByVal Altitude As Variant, ByVal Tag As String, ByVal MainColor As String, ByVal LaunchSiteElevation As Double, ByVal FILE_PATH As String, ByVal WaiverRadius_NM As Double, ByVal WaiverLat As Double, ByVal WaiverLon As Double)
    Dim fso As Object
    Dim file As Object
    Dim i As Integer
    Dim lineColor As String
    Dim WaiverRadiusKm As Double
    Dim WaiverName As String
    
    ' Convert variants to arrays
    Dim latArray() As Variant
    Dim lonArray() As Variant
    Dim altArray() As Variant
    latArray = latitude
    lonArray = longitude
    altArray = Altitude
    

    If latitude(0) = WaiverLat Then
        WaiverName = ""
    Else
        WaiverName = "Waiver Center"
    End If
    
    WaiverRadiusKm = WaiverRadius_NM * 0.539957 ' Convert nautical mile to km for plotting
    ' Define line color based on MainColor
    lineColor = MainColor
    
    ' Open the file for appending
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(FILE_PATH, 8)
    
    ' Draw circle if it hasn't been drawn yet
    If Not CircleDrawn Then
        ' Write the circle placemark centered at the Waiver Lat and Waiver Lon
        file.WriteLine "<Placemark>"
        file.WriteLine "<name>Waiver Center</name>"
        file.WriteLine "<Style>"
        file.WriteLine "<IconStyle>"
        file.WriteLine "<color>ff0000ff</color>"
        file.WriteLine "</IconStyle>"
        file.WriteLine "</Style>"
        file.WriteLine "<Point>"
        file.WriteLine "<altitudeMode>clampToGround</altitudeMode>"
        file.WriteLine "<coordinates>" & WaiverLon & "," & WaiverLat & ",0</coordinates>"
        file.WriteLine "</Point>"
        file.WriteLine "</Placemark>"
        
        ' Write the circle polygon
        file.WriteLine "<Placemark>"
        file.WriteLine "<name>" & WaiverName & "</name>"
        file.WriteLine "<Style>"
        file.WriteLine "<LineStyle>"
        file.WriteLine "<color>ff0000ff</color>"
        file.WriteLine "<width>2</width>"
        file.WriteLine "</LineStyle>"
        file.WriteLine "<PolyStyle>"
        file.WriteLine "<color>1aff0000</color>" ' Red color with some transparency
        file.WriteLine "</PolyStyle>"
        file.WriteLine "</Style>"
        file.WriteLine "<Polygon>"
        file.WriteLine "<outerBoundaryIs>"
        file.WriteLine "<LinearRing>"
        file.WriteLine "<coordinates>"
        
        ' Calculate circle points
        Dim centerLat As Double
        Dim centerLon As Double
        centerLat = WaiverLat
        centerLon = WaiverLon
    Dim j As Integer
    For j = 0 To 360 Step 10
        Dim circleLat As Double
        Dim circleLon As Double
        circleLat = centerLat + (WaiverRadiusKm / 111.32) * Sin(j * 3.14159265358979 / 180)
        circleLon = centerLon + (WaiverRadiusKm / (111.32 * Cos(centerLat * 3.14159265358979 / 180))) * Cos(j * 3.14159265358979 / 180)
        
        file.WriteLine circleLon & "," & circleLat & ",0"
        Next j
        
        file.WriteLine "</coordinates>"
        file.WriteLine "</LinearRing>"
        file.WriteLine "</outerBoundaryIs>"
        file.WriteLine "</Polygon>"
        file.WriteLine "</Placemark>"
        
        ' Set CircleDrawn to True
        CircleDrawn = True
    End If
    
    ' Write the flight path coordinates with altitude, tag, and color
    file.WriteLine "<Placemark>"
    file.WriteLine "<name>Fllight Path, " & Tag & "</name>"
    
    ' flight path track style
    file.WriteLine "<Style>"
    file.WriteLine "<LineStyle>"
    file.WriteLine "<color>" & lineColor & "</color>"
    file.WriteLine "<width>4</width>"
    file.WriteLine "</LineStyle>"
    file.WriteLine "</Style>"
    
    ' flight path coordinates
    file.WriteLine "<LineString>"
    file.WriteLine "<altitudeMode>absolute</altitudeMode>"
    file.WriteLine "<tessellate>1</tessellate>"
    file.WriteLine "<coordinates>"
    For i = LBound(latArray) To UBound(latArray)
        file.WriteLine lonArray(i) & "," & latArray(i) & "," & Round((altArray(i) + LaunchSiteElevation) / 3.281, 2)
    Next i
    file.WriteLine "</coordinates>"
    file.WriteLine "</LineString>"
    
    file.WriteLine "</Placemark>"
    
    ' Write the ground path coordinates with altitude, tag, and color
    file.WriteLine "<Placemark>"
    file.WriteLine "<name>Ground Track, " & Tag & "</name>"
    
    ' Ground track style (same color as main track, line width 1)
    file.WriteLine "<Style>"
    file.WriteLine "<LineStyle>"
    file.WriteLine "<color>" & lineColor & "</color>"
    file.WriteLine "<width>1</width>"
    file.WriteLine "</LineStyle>"
    file.WriteLine "</Style>"
    
    ' Ground path coordinates
    file.WriteLine "<LineString>"
    file.WriteLine "<altitudeMode>clampToGround</altitudeMode>"
    file.WriteLine "<tessellate>1</tessellate>"
    file.WriteLine "<coordinates>"
    For i = LBound(latArray) To UBound(latArray)
        file.WriteLine lonArray(i) & "," & latArray(i) & ",0"
    Next i
    file.WriteLine "</coordinates>"
    file.WriteLine "</LineString>"
    
    file.WriteLine "</Placemark>" ' Closing the ground track Placemark
    
    ' Write the placemark for the last coordinate of the ground track
    file.WriteLine "<Placemark>"
    file.WriteLine "<name>" & Tag & "</name>"
    file.WriteLine "<Style>"
    file.WriteLine "<IconStyle>"
    file.WriteLine "<color>" & "00FFFFFF" & "</color>"
    file.WriteLine "</IconStyle>"
    file.WriteLine "</Style>"
    file.WriteLine "<Point>"
    file.WriteLine "<altitudeMode>clampToGround</altitudeMode>"
    file.WriteLine "<coordinates>" & lonArray(UBound(lonArray)) & "," & latArray(UBound(latArray)) & "," & "0" & "</coordinates>"
    file.WriteLine "</Point>"
    file.WriteLine "</Placemark>"
    
    ' Close the file
    file.Close
    Set file = Nothing
    Set fso = Nothing
    
End Sub



Sub InitializeKMLFile(ByVal FILE_PATH As String)
    Dim fso As Object
    Dim file As Object
    
    ' Open the file for appending
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(FILE_PATH, 2, True)
    
    ' Write KML header
    file.WriteLine "<?xml version=""1.0"" encoding=""UTF-8""?>"
    file.WriteLine "<kml xmlns=""http://www.opengis.net/kml/2.2"">"
    file.WriteLine "<Document>"
    
    ' Close the file
    file.Close
    Set file = Nothing
    Set fso = Nothing
End Sub



Sub FinalizeKMLFile(ByVal FILE_PATH As String)
    Dim fso As Object
    Dim file As Object
    
    ' Open the file for appending
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set file = fso.OpenTextFile(FILE_PATH, 8)
    
    ' Write KML footer
    file.WriteLine "</Document>"
    file.WriteLine "</kml>"
    
    ' Close the file
    file.Close
    Set file = Nothing
    Set fso = Nothing
    CircleDrawn = False
End Sub




