'-----------------------------------------------------------------------------------
'-------------Functions and Subs used to read data from Winds Aloft-----------------
'-----------------------------------------------------------------------------------

Option Explicit

Function GetAltitudeArray(response As String, model As String) As Variant
    Dim altitudeArr As Variant
    Dim startPos As Long
    Dim endPos As Long
    Dim altString As String
    Dim matches As Object
    Dim i As Integer
    Dim temp As Variant

    ' Determine the correct array name based on the model
    If model = "RAP" Then
        altString = """altFtRaw"""
    ElseIf model = "Open-Meteo" Then
        altString = """altFt"""
    Else
        GetAltitudeArray = CVErr(xlErrValue) ' Return error if model is neither "RAP" nor "Open-Meteo"
        Exit Function
    End If

    ' Find the starting position of the array
    startPos = InStrRev(response, altString)

    ' Find the ending position of the array
    endPos = InStr(startPos, response, "]")

    ' Extract the array as a string
    altString = Mid(response, startPos, endPos - startPos + 1)

    ' Use regular expressions to extract numbers from the string
    Set matches = GetRegexMatches(altString, "-?\d+")
    
    ' Convert the matches to a double array
    ReDim altitudeArr(0 To matches.Count - 1)
    For i = 0 To matches.Count - 1
        altitudeArr(i) = CDbl(matches.Item(i))
    Next i

    ' Reverse the array
    For i = LBound(altitudeArr) To (UBound(altitudeArr) - 1) / 2
        temp = altitudeArr(i)
        altitudeArr(i) = altitudeArr(UBound(altitudeArr) - i)
        altitudeArr(UBound(altitudeArr) - i) = temp
    Next i

    ' Return the reversed double array
    GetAltitudeArray = altitudeArr
End Function

Function GetRegexMatches(inputString As String, pattern As String) As Object
    Dim regex As Object
    Dim matches As Object
    
    Set regex = CreateObject("VBScript.RegExp")
    regex.Global = True
    regex.MultiLine = True
    regex.IgnoreCase = False
    regex.pattern = pattern
    
    Set matches = regex.Execute(inputString)
    Set GetRegexMatches = matches
End Function


Function ExtractValues(apiResponse As String, model As String, searchStr As String) As Variant
    Dim term As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim valuesStr As String
    Dim valuesArray() As String
    Dim i As Integer
    Dim temp As String

    ' Determine the term based on the model and search string
    If model = "RAP" And searchStr = "WindSpeed" Then
        term = "speedRaw"
    ElseIf model = "RAP" And searchStr = "WindDirection" Then
        term = "directionRaw"
    ElseIf model = "Open-Meteo" And searchStr = "WindSpeed" Then
        term = "speed"
    ElseIf model = "Open-Meteo" And searchStr = "WindDirection" Then
        term = "direction"
    Else
        Exit Function ' Exit if the model and search string combination is invalid
    End If

    ' Find the start position of the term key
    startPos = InStr(apiResponse, """" & term & """:{") + Len("""" & term & """:{")
    
    ' Find the end position of the term key
    endPos = InStr(startPos, apiResponse, "}")
    
    ' Extract the substring containing the term values
    valuesStr = Mid(apiResponse, startPos, endPos - startPos)
    
    ' Split the term values into an array
    valuesArray = Split(valuesStr, ",")
    
    ' Resize the array to remove the key-value pairs and keep only the values
    ReDim Values(0 To UBound(valuesArray))
    For i = 0 To UBound(valuesArray)
        ' Trim the value to remove leading/trailing spaces and double quotes (for RAP data)
        Values(i) = Val(Replace(Trim(Split(valuesArray(i), ":")(1)), """", ""))
    Next i
    
    ' Reverse the array
    For i = LBound(Values) To (UBound(Values) - 1) / 2
        temp = Values(i)
        Values(i) = Values(UBound(Values) - i)
        Values(UBound(Values) - i) = temp
    Next i

    ExtractValues = Values
End Function


Function GetModel(jsonString As String) As String
    Dim startPos As Integer
    Dim endPos As Integer
    Dim modelStartPos As Integer
    Dim modelEndPos As Integer
    
    ' Find the position of the "model" key in the JSON string
    modelStartPos = InStr(jsonString, """model"":""") + Len("""model"":""")
    
    ' Find the position of the end of the "model" value
    startPos = modelStartPos
    Do While Mid(jsonString, startPos, 1) <> ","
        startPos = startPos + 1
        ' Check if startPos exceeds the length of the string
        If startPos > Len(jsonString) Then
            startPos = Len(jsonString)
            Exit Do
        End If
    Loop
    endPos = startPos - 2 ' Adjust to exclude the comma and the trailing quote
    
    ' Extract the "model" value
    GetModel = Mid(jsonString, modelStartPos, endPos - modelStartPos + 1)
End Function


