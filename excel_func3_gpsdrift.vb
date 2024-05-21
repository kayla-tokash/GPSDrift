Option Explicit

Sub Get_LS_Elevation()
    Dim ws As Worksheet
    Dim targetName As String
    Dim elevation As Integer
    Dim latitude As Double
    Dim longitude As Double
    Dim waiver As Double
    
    ' Set the worksheet
    Set ws = ThisWorkbook.Sheets("GPS DriftCast")
    
    ' Get the target site name
    targetName = ws.Range("LS_Elev_Target").value
    
    ' Find the target site in the SITE named range
    Dim siteRange As Range
    Set siteRange = ws.Range("SITE")
    Dim siteCell As Range
    Set siteCell = siteRange.Find(What:=targetName, LookIn:=xlValues, LookAt:=xlWhole)
    
    ' If target site is found, get the elevation, latitude, and longitude
    If Not siteCell Is Nothing Then
        Dim rowIndex As Long
        rowIndex = siteCell.Row - siteRange.Row + 1
        
        ' Get the latitude and longitude values
        If IsEmpty(ws.Range("LATITUDE").Cells(rowIndex, 1).value) Then
            MsgBox "Latitude is blank for: " & targetName, vbExclamation
            Exit Sub
        Else
            latitude = CDbl(ws.Range("LATITUDE").Cells(rowIndex, 1).value)
            If latitude < -90 Or latitude > 90 Then
                MsgBox "The provided latitude for site: " & targetName & " is outside the valid range of -90 to 90.", vbExclamation
                Exit Sub
            End If
        End If
        
        If IsEmpty(ws.Range("LONGITUDE").Cells(rowIndex, 1).value) Then
            MsgBox "Longitude is blank for: " & targetName, vbExclamation
            Exit Sub
        Else
            longitude = CDbl(ws.Range("LONGITUDE").Cells(rowIndex, 1).value)
            If longitude < -180 Or longitude > 180 Then
                MsgBox "The provided longitude for site: " & targetName & " is outside the valid range of -180 to 180.", vbExclamation
                Exit Sub
            End If
        End If
        
        ' Get the elevation value
        elevation = GetElevation(CStr(latitude), CStr(longitude))
        
        ' Set the elevation value
        ws.Range("ELEVATION").Cells(rowIndex, 1).value = elevation
        
        ' Use the latitude and longitude values as needed
        MsgBox "Elevation found and updated for: " & targetName
    Else
        MsgBox "Site not found.", vbExclamation
        End
    End If
End Sub


Function GetElevation(LAT As String, LON As String) As Integer
Dim response As String
    Dim xhr As Object
    Dim responseText As String
    Dim startIndex As Integer
    Dim endIndex As Integer
    Dim elevation As Double
    Dim url As String


url = "https://api.open-elevation.com/api/v1/lookup?locations=" & LAT & "," & LON
'url = "https://api.open-elevation.com/api/v1/lookup?locations=45.546058,-92.92464"
'Debug.Print url


    ' Create a new XMLHTTP request object
    Set xhr = CreateObject("MSXML2.XMLHTTP")

    ' Open a connection to the URL
    xhr.Open "GET", url, False

    ' Send the request
    xhr.send

    ' Get the response text
    responseText = xhr.responseText
    'Debug.Print responseText
    
    
    ' Clean up
    Set xhr = Nothing

    
    ' Find the index of the elevation value in the response text
    startIndex = InStr(responseText, """elevation"":") + Len("""elevation"":")
    endIndex = InStr(startIndex, responseText, "}")
    
    ' Extract the elevation value and convert it to a Double
    elevation = CDbl(Mid(responseText, startIndex, endIndex - startIndex))
    
    ' Return the elevation
    GetElevation = CInt(elevation * 3.281) 'Convert m to ft
  
   
End Function

