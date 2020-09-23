Attribute VB_Name = "modStart"
Option Explicit



Sub Main()
'
'First check for registration
'Returns the codes
' 0 - No error,trial
' 1 - No error,registered
' 2 - Expired..real version
' 3 - Expired..trial version
' 4 - Harddiskkey not matched
'
Dim lCode As Integer
  lCode = CheckRegistration
  Load frmSplash
  frmSplash.LoadEditor lCode
End Sub

Public Function CheckRegistration() As Integer
'
'Check the registration
'
Dim lKey As String
Dim lDatereg As Date
Dim lDuration As Integer
Dim lDays As Integer
  If GetSetting(App.Title, "regkey", "regkey") = "" Then
    'For trial version
    If Not IsExpired(True) Then
      lDatereg = Decode(GetMySetting("StrRgd"))
      lDays = 15
      If DateDiff("d", lDatereg, Date) >= 0 Then
        'Remaining days
        CheckRegistration = 0
      End If
    Else 'Expired
      CheckRegistration = 3
    End If
  Else
    'For registered version
    'harddiskkey verification
    lKey = GetSetting(App.Title, "regkey", "regkey")
    lKey = Mid(lKey, 1, InStrRev(lKey, "-") - 1)
    lKey = Replace(lKey, "-", "")
    If GetDecodedKey(GetHardDiskKey(), 20) = lKey Then
        'isExpired
        If Not IsExpired Then
          lDatereg = Decode(GetMySetting("StrRgd"))
          lDuration = val(GetSetting(App.Title, "regvalid", "regvalid"))
          lDays = DateDiff("d", lDatereg, DateAdd("m", lDuration, lDatereg))
          If DateDiff("d", lDatereg, Date) >= 0 Then
            'remainging days
            CheckRegistration = 1
          End If
        Else
          'expired
          CheckRegistration = 2
        End If
    Else
      CheckRegistration = 4
    End If
  End If
End Function

Private Function IsExpired(Optional pTrial As Boolean) As Boolean
'
'StrRgd - Registered date
'LstOpn - Last time it is opened
'
Dim strDate                 As String
Dim oldDate                 As String
Dim strSplash               As String
    On Error GoTo Cerr
    
    strDate = GetMySetting("StrRgd")
    oldDate = GetMySetting("LstOpn")
    If strDate = "" Then 'Fresh User
        If pTrial Then 'For trial version
          SaveMySetting "StrRgd", Encode(Date)
          strDate = Date
        Else
          SaveMySetting "StrRgd", Encode(CDate(GetSetting(App.Title, "regdate", "regdate")))
          strDate = CDate(GetSetting(App.Title, "regdate", "regdate"))
        End If
        SaveMySetting "LstOpn", Encode(Now)
        
        oldDate = Now
    Else
        strDate = Decode(strDate)
        If oldDate = "" Then    'For Checking Installed trial version date
            IsExpired = True
        Else
            oldDate = Decode(oldDate)
        End If
        If Not IsExpired Then         'For Malpractice attempting to Modify system date
            If CDate(oldDate) < CDate(strDate) Then IsExpired = True
            If CDate(oldDate) > Now Then IsExpired = True
        End If
    End If
    If Not IsExpired Then             'For Checking Expiry
        SaveMySetting "LstOpn", Encode(Now)
        If val(DateDiff("d", Now, CDate(strDate) + IIf(pTrial, 15, 365))) <= 0 Then IsExpired = True
    End If
    Exit Function
Cerr:
    IsExpired = False
End Function

