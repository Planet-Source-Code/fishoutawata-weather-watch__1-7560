Attribute VB_Name = "modMain"
Public objHTTP  As New MSXML.XMLHTTPRequest
Public strCity  As String
Public strState As String

Public Type WeatherGrid
    CurrentStat As String
    Temp        As String
    Wind        As String
    DewPoint    As String
    Humidity    As String
    Visibility  As String
    Barometer   As String
    SunRise     As String
    SunSet      As String
End Type

Public Function CheckCity(City As String)

Dim strCity      As String
Dim strTempCity  As String
Dim strChrHolder As String
Dim Counter      As Integer

strCity = Trim(City)

'If city name is more then two words we need to add an underscore to join them
If InStr(strCity, " ") > 0 Then
    For Counter = 1 To Len(strCity)
        strChrHolder = Mid(strCity, Counter, 1)
        If strChrHolder <> Chr(32) Then
            strTempCity = strTempCity & strChrHolder
        Else
            strTempCity = strTempCity & "_"
        End If
    Next
    strCity = strTempCity
End If

CheckCity = strCity

End Function

Public Function ParseData(Request As String) As WeatherGrid

Dim strRequest     As String
Dim StartFrom      As Long
Dim EndAt          As Long
Dim RetVal         As Long
Dim strCurrentStat As String
Dim strTemp        As String
Dim strWind        As String
Dim strDewPoint    As String
Dim strHumitity    As String
Dim strVisibility  As String
Dim strBarometer   As String
Dim strSunrise     As String
Dim strSunset      As String

strRequest = Request

'Search Webpage for the Data we need

'Get Current Status
RetVal = InStr(strRequest, "as reported at")
StartFrom = RetVal
RetVal = InStr(StartFrom, strRequest, "<B>")
StartFrom = RetVal + 3
EndAt = InStr(StartFrom, strRequest, "</B>")
strCurrentStat = Mid(strRequest, StartFrom, EndAt - StartFrom)

'Get Current Temp
RetVal = InStr(strRequest, "Temp:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "&deg;F")
    StartFrom = RetVal
    strTemp = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current Wind Speed
RetVal = InStr(StartFrom, strRequest, "Wind:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "mph")
    StartFrom = RetVal
    strWind = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current DewPoint
RetVal = InStr(StartFrom, strRequest, "Dewpoint:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "&deg;F")
    StartFrom = RetVal
    strDewPoint = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current Humidity
RetVal = InStr(StartFrom, strRequest, "Rel. Humidity:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "%")
    StartFrom = RetVal
    strHumidity = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current Visibility
RetVal = InStr(StartFrom, strRequest, "Visibility:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "miles")
    StartFrom = RetVal
    strVisibility = CheckIt(Mid(strRequest, StartFrom - 3, 3))
End If

'Get Current Barometer
RetVal = InStr(StartFrom, strRequest, "Barometer:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "inches")
    StartFrom = RetVal
    strBarometer = CheckOther(Mid(strRequest, StartFrom - 6, 6))
End If

'Get Current Sunrise
RetVal = InStr(StartFrom, strRequest, "Sunrise:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "am")
    StartFrom = RetVal
    strSunrise = CheckOther(Mid(strRequest, StartFrom - 6, 6))
End If

'Get Current Sunset
RetVal = InStr(StartFrom, strRequest, "Sunset:")
If RetVal > 0 Then
    StartFrom = RetVal
    RetVal = InStr(StartFrom, strRequest, "pm")
    StartFrom = RetVal
    strSunset = CheckOther(Mid(strRequest, StartFrom - 6, 6))
End If

ParseData.CurrentStat = strCurrentStat
ParseData.Temp = strTemp
ParseData.Wind = strWind
ParseData.DewPoint = strDewPoint
ParseData.Humidity = strHumidity
ParseData.Visibility = strVisibility
ParseData.Barometer = strBarometer
ParseData.SunRise = strSunrise
ParseData.SunSet = strSunset

End Function
Private Function CheckIt(Tmp As String) As String

Dim strTmp       As String
Dim strTempTmp   As String
Dim strChrHolder As String
Dim Counter      As Integer

strTmp = Tmp

For Counter = 1 To Len(strTmp)
    strChrHolder = Mid(strTmp, Counter, 1)
    If IsNumeric(strChrHolder) Then
        strTempTmp = strTempTmp & strChrHolder
    End If
Next

CheckIt = strTempTmp

End Function

Private Function CheckOther(Tmp As String) As String

Dim strTmp       As String
Dim strTempTmp   As String
Dim strChrHolder As String
Dim Counter      As Integer

strTmp = Tmp

For Counter = 1 To Len(strTmp)
    strChrHolder = Mid(strTmp, Counter, 1)
    If strChrHolder <> ">" Then
        strTempTmp = strTempTmp & strChrHolder
    End If
Next

CheckOther = strTempTmp

End Function
