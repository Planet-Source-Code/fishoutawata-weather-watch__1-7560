VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Weather Watch"
   ClientHeight    =   3330
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   7725
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7725
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picStats 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2895
      ScaleWidth      =   7725
      TabIndex        =   6
      Top             =   420
      Width           =   7725
      Begin VB.Label lblReportedOn 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   90
         TabIndex        =   24
         Top             =   120
         Width           =   7545
      End
      Begin VB.Image imgWeatherImage 
         Height          =   1230
         Left            =   810
         Top             =   510
         Width           =   1230
      End
      Begin VB.Label lblCurrentStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   390
         TabIndex        =   23
         Top             =   1980
         Width           =   2115
      End
      Begin VB.Label lbltemp 
         BackStyle       =   0  'Transparent
         Caption         =   "Temp:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         TabIndex        =   22
         Top             =   540
         Width           =   1185
      End
      Begin VB.Label lblWind 
         BackStyle       =   0  'Transparent
         Caption         =   "Wind:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         TabIndex        =   21
         Top             =   810
         Width           =   1185
      End
      Begin VB.Label lblDewPoint 
         BackStyle       =   0  'Transparent
         Caption         =   "DewPoint:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         TabIndex        =   20
         Top             =   1095
         Width           =   1185
      End
      Begin VB.Label lblHumidity 
         BackStyle       =   0  'Transparent
         Caption         =   "Humidity:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         TabIndex        =   19
         Top             =   1365
         Width           =   1185
      End
      Begin VB.Label lblVisiblity 
         BackStyle       =   0  'Transparent
         Caption         =   "Visiblity:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         TabIndex        =   18
         Top             =   1635
         Width           =   1185
      End
      Begin VB.Label lblBaroMeter 
         BackStyle       =   0  'Transparent
         Caption         =   "Barometer:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         TabIndex        =   17
         Top             =   1905
         Width           =   1185
      End
      Begin VB.Label lblSunrise 
         BackStyle       =   0  'Transparent
         Caption         =   "Sunrise:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         TabIndex        =   16
         Top             =   2190
         Width           =   1185
      End
      Begin VB.Label lblSunset 
         BackStyle       =   0  'Transparent
         Caption         =   "Sunset:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3570
         TabIndex        =   15
         Top             =   2460
         Width           =   1185
      End
      Begin VB.Label lblTempVal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   14
         Top             =   540
         Width           =   1155
      End
      Begin VB.Label lblWindVal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   13
         Top             =   810
         Width           =   2865
      End
      Begin VB.Label lblDewPointVal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   12
         Top             =   1110
         Width           =   1155
      End
      Begin VB.Label lblHumidityVal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   11
         Top             =   1380
         Width           =   1155
      End
      Begin VB.Label lblVisibilityVal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   10
         Top             =   1650
         Width           =   1155
      End
      Begin VB.Label lblBarometerVal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   9
         Top             =   1920
         Width           =   1155
      End
      Begin VB.Label lblSunriseVal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   8
         Top             =   2190
         Width           =   1155
      End
      Begin VB.Label lblSunsetVal 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   4800
         TabIndex        =   7
         Top             =   2460
         Width           =   1155
      End
   End
   Begin VB.PictureBox picWeatherToday 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   0
      ScaleHeight     =   405
      ScaleWidth      =   7875
      TabIndex        =   0
      Top             =   0
      Width           =   7875
      Begin VB.ComboBox cmbState 
         Height          =   315
         Left            =   5610
         TabIndex        =   2
         Top             =   30
         Width           =   735
      End
      Begin VB.ComboBox cmbCity 
         Height          =   315
         Left            =   3000
         TabIndex        =   1
         Top             =   30
         Width           =   1785
      End
      Begin VB.Image imgGo 
         Height          =   345
         Left            =   6720
         MouseIcon       =   "frmMain.frx":0442
         MousePointer    =   99  'Custom
         Picture         =   "frmMain.frx":074C
         Top             =   0
         Width           =   390
      End
      Begin VB.Label lblState 
         BackStyle       =   0  'Transparent
         Caption         =   "State"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   5070
         TabIndex        =   5
         Top             =   90
         Width           =   705
      End
      Begin VB.Label lblCity 
         BackStyle       =   0  'Transparent
         Caption         =   "City"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2610
         TabIndex        =   4
         Top             =   90
         Width           =   975
      End
      Begin VB.Label lblWeatherToday 
         BackStyle       =   0  'Transparent
         Caption         =   "Weather Today"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   60
         TabIndex        =   3
         Top             =   60
         Width           =   2145
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private objWeather As WeatherGrid

Private Sub Form_Load()

lbltemp.Visible = False
lblWind.Visible = False
lblBaroMeter.Visible = False
lblVisiblity.Visible = False
lblDewPoint.Visible = False
lblHumidity.Visible = False
lblSunrise.Visible = False
lblSunset.Visible = False

End Sub

Private Sub Form_Resize()

picWeatherToday.Width = Me.Width
picStats.Height = Me.Height
picStats.Width = Me.Width

End Sub

Private Sub imgGo_Click()

Dim strRequest As String
Dim strWebPage As String

'Check for City and State
If cmbCity = "" Or cmbState = "" Then
    MsgBox "You must Select and City and State", vbCritical, "Weather Today"
    Exit Sub
End If

'Fill UDT with Values from Form
strCity = LCase(CheckCity(cmbCity.Text))
strState = LCase(Trim(cmbState.Text))

'Retrieve Web Page
strWebPage = "http://www.weather.com/weather/cities/us_" & strState & "_" & strCity & ".html"
objHTTP.open "GET", strWebPage, False
objHTTP.send

'Check status of page
If objHTTP.Status <> "200" Then
    MsgBox "Cannot find City/State combination, please " & vbNewLine & _
           "make sure you spelled the City correctly", vbCritical, "Weather Today"
    Exit Sub
End If

'Pass Response back from Server
strRequest = objHTTP.responseText

'Show what city we are loading
frmMain.lblReportedOn = cmbCity.Text & ", " & cmbState.Text & " Reported on " & Now
lbltemp.Visible = True
lblWind.Visible = True
lblBaroMeter.Visible = True
lblVisiblity.Visible = True
lblDewPoint.Visible = True
lblHumidity.Visible = True
lblSunrise.Visible = True
lblSunset.Visible = True

'Get Current Status
objWeather = ParseData(strRequest)

'Load Image
Select Case objWeather.CurrentStat
    Case "Fair"
        imgWeatherImage.Picture = LoadPicture(App.Path & "\Fair.gif")
    Case "Sunny"
        imgWeatherImage.Picture = LoadPicture(App.Path & "\Fair.gif")
    Case "Cloudy"
        imgWeatherImage.Picture = LoadPicture(App.Path & "\Cloudy.gif")
    Case Else
        imgWeatherImage.Picture = LoadPicture(App.Path & "\PartlyCloudy.gif")
End Select

'Load Form Labels
lblCurrentStatus.Caption = objWeather.CurrentStat
lblTempVal.Caption = objWeather.Temp & " deg"
lblWindVal.Caption = objWeather.Wind & " mph"
lblDewPointVal.Caption = objWeather.DewPoint & " deg"
lblHumidityVal.Caption = objWeather.Humidity & " %"
lblVisibilityVal.Caption = objWeather.Visibility & " miles"
lblBarometerVal.Caption = objWeather.Barometer & " inches"
lblSunriseVal.Caption = objWeather.SunRise & " am"
lblSunsetVal.Caption = objWeather.SunSet & " pm"

End Sub

Private Sub mnuExit_Click()

'End Program
Unload Me

End Sub
