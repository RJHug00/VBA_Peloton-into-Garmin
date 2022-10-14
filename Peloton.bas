Attribute VB_Name = "Peloton"
Option Explicit

  ' Specify location for the file.  Desktop would be another low-mystery location.
Private Const FIT_FILE_NAME As String = "%USERPROFILE%\Downloads\PelotonRide_~YYYYmmDD_HHMM~.FIT"
Private expandedFileName As String

  ' Constant air temperature associated with every second of Peloton ride data
Private Const FIT_DATA_TEMPERATURE As Integer = 21  ' Celsius for 70 degrees

  ' Garmin prefers 0:254 while Peloton records it 0:100
Private Const PELOTON_TO_GARMIN_RESISTANCE_FACTOR As Double = 2.54   '1.0

  ' Normally zero, this causes us to operate on other than the most recent Peloton ride
Private Const DESIRED_RIDE_INDEX As Integer = 0

  ' Normally zero, this makes a ride appear on Garmin to have occurred at a different date/time
Private Const DEVELOPMENT_TIMESTAMP_BIAS As Long = (86200 * 1) ' + One day

Private Const SECONDS_PER_HOUR As Long = 3600
Private Const SECONDS_PER_DAY As Long = SECONDS_PER_HOUR * 24
Private Const WINDOWS_EPOC_BIAS As Long = 25569                 ' days to add to Peloton timestamp

  ' This may need to be fiddled with for your timezone
Private Const UTC_OFFSET_SECONDS As Long = -4& * SECONDS_PER_HOUR               ' US/Eastern is 4 hours
Private Const UTC_OFFSET_FLOAT As Double = UTC_OFFSET_SECONDS / SECONDS_PER_DAY ' fraction of a day

Private Const METERS_PER_MILE As Double = 1609.344
Private Const METERS_PER_SECOND_FACTOR As Double = METERS_PER_MILE / SECONDS_PER_HOUR  ' MPH to meters-per-sec

Private Const PELOTON As String = "https://api.onepeloton.com/"
Private Const PELOTON_API As String = PELOTON & "api/"

  ' Garmin doesn't display anything about these values, so it might not be worth inclusion in FIT file
Private Const INCLUDE_DEVELOPER_DATA_LOGIC As Boolean = False

  ' Users modifying the logic in any substantial way should get their own unique guid [w/ version]
  ' and allow this one to correspond to the processing logic the original author is maintaining.

Private Const APPLICATION_GUID As String = "b1baaeee-d578-4708-97fe-5d7863b1e998"
Private Const APPLICATION_VSN As Integer = "100"     ' version 1.00 times 100

' Our implementation of assigning Local message types.  No re-use at this time.
Private Enum LCL_TYPE
  gFILE_ID = 0
  gCREATOR = 1
  gDEVICE_INFO = 2
  gDEVELOPER_ID = 3
  gFIELD_DEF = 4
  gEVENT = 5
  gRECORD = 6
  gLAP = 7
  gSESSION = 8
  gACTIVITY = 9
  gMAX_ASSIGNABLE = 15
End Enum

Public Sub ConvertPelotonRideToGarminActivity()
Attribute ConvertPelotonRideToGarminActivity.VB_Description = "Invoke the macro to retrieve one Peloton bike ride and convert it to an uploadable Garmin FIT file."
Attribute ConvertPelotonRideToGarminActivity.VB_ProcData.VB_Invoke_Func = "a\n14"

  Dim wasCalculating As XlCalculation: wasCalculating = Application.Calculation
  Application.Calculation = xlCalculationManual
  Application.ScreenUpdating = False
  Application.Cursor = xlWait

  Dim mStart As Date: mStart = Now()
  Debug.Print mStart & " Starting"

  ' Old-school 'c' and assembly language programmers prefer variables declared
  ' up top to declutter the actual processing logic. You do you; I'll do me.

  Dim oHttpReq As Object, j As Object
  Dim tp() As FIT_Definition_Element, v() As Byte, JSON As String
  Dim numDataBytes As Integer, numEventDataBytes As Integer, i As Integer
  Dim thisSpeed As Double, thisSpeedIntegral As Long
  Dim pelotonStartTime As Long, startTimestamp As Long, currentTimestamp As Long
  Dim totalDistance As Double, totalDistanceIntegral As Long: totalDistanceIntegral = 0
  Dim pelotonDay As Long, pelotonSeconds As Long, windowsDay As Long, windowsStartTime As Date
  Dim hasHeartRateData As Boolean: hasHeartRateData = False
  Dim numLapsEmitted As Integer: numLapsEmitted = 0
  Dim samples As Long, sampleI As Long

'========================================================================================
'  Get data for the most recent Peloton ride through API calls
'========================================================================================

  Set oHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
  oHttpReq.SetTimeouts 30000, 30000, 30000, 30000
  Set j = CreateObject("Scripting.Dictionary")

  JSON = "{" & vbLf  ' [Peloton doesn't recognize single tics for JSON property names nor values]
  JSON = JSON & """username_or_email"": """ & Environ("PELOTON_USER") & """," & vbLf
  JSON = JSON & """password"": """ & Environ("PELOTON_KEY") & """" & vbLf & "}"

  oHttpReq.Open "POST", PELOTON & "auth/login", False
  oHttpReq.setRequestHeader "Content-Type", "application/json"
  oHttpReq.Send JSON

  FillMemory JSON, 2 * Len(JSON), 0 ' clear out that unicode memory

  Set j = JSONparser(IIf(oHttpReq.Status = 200, oHttpReq.responseText, ""))

  ' ----------
  ' There exists potential for an enhancement where I display a list of the last N activities
  ' and let the user click the one they want converted.  Such an interface could theoretically
  ' accept login credentials and allow for entry of the air temperature to use. I'm not fond
  ' of that approach as it demands the user re-type their username and password every time.
  ' ----------

  Dim objString As String: objString = "obj.data(" & CStr(DESIRED_RIDE_INDEX) & ")"

  oHttpReq.Open "GET", PELOTON_API & "user/" & j("obj.user_id") & _
                       "/workouts?limit=" & CStr(DESIRED_RIDE_INDEX + 1) & "&page=0", False
  oHttpReq.setRequestHeader "cookie", "peloton_session_id=" & j("obj.session_id")
  oHttpReq.setRequestHeader "peloton-platform", "web"       ' rumored to be needed, but successful without it
  oHttpReq.Send

  Set j = JSONparser(IIf(oHttpReq.Status = 200, oHttpReq.responseText, ""))

  ' It will only process a completed cycling activity.  You're forced to manually fiddle with
  ' DESIRED_RIDE_INDEX if you recorded a run or workout *after* the ride.
  If j(objString & ".fitness_discipline") <> "cycling" Then
    MsgBox "Most-recent Peloton activity is not 'cycling'" & vbLf & _
           "adjust DESIRED_RIDE_INDEX to workaround this.", vbOKOnly, "Error"
    Exit Sub
  End If
  ' I'm not sure what it would mean for an activity to be other than "COMPLETE"
  Debug.Assert j(objString & ".status") = "COMPLETE"

  ' Perforamce_graph REST endpoint doesn't give us the startTime we want; get it here.
  pelotonStartTime = CLng(j(objString & ".start_time"))
  startTimestamp = pelotonStartTime - GARMIN_EPOC_BIAS + DEVELOPMENT_TIMESTAMP_BIAS ' development bias is normally zero
  currentTimestamp = startTimestamp

  pelotonDay = pelotonStartTime / SECONDS_PER_DAY
  If pelotonDay * SECONDS_PER_DAY > pelotonStartTime Then pelotonDay = pelotonDay - 1  ' force rounding down
  pelotonSeconds = pelotonStartTime - (pelotonDay * SECONDS_PER_DAY)                   ' seconds into that day
  windowsDay = pelotonDay + WINDOWS_EPOC_BIAS
  windowsStartTime = CDbl(windowsDay) + pelotonSeconds / SECONDS_PER_DAY    ' in UTC
  windowsStartTime = windowsStartTime + UTC_OFFSET_FLOAT                    ' negative for western hemisphere

  'queryString can't specify "&fields=power,cadence,..." like we'd want. Morons.
  oHttpReq.Open "GET", PELOTON_API & "workout/" & j(objString & ".id") & _
                       "/performance_graph?every_n=1", False
  oHttpReq.Send

  Set j = JSONparser(IIf(oHttpReq.Status = 200, oHttpReq.responseText, ""))

  ' There can be time-in-heart-zone metrics in here, but Peloton zones
  ' are hardwired percentages of MaxHR and are different than Garmin zones.
  ' Z1 <= 65% MHR   Z2 <= 75% MHR   Z3 <= 85% MHR   Z4 <= 95% MHR   Z5 > 95% MHR

  ' I compute the duration I send to Garmin, but this is Peloton's take.
  'Dim total_duration As Long: total_duration = CLng(j("obj.duration"))                       ' seconds
  ' Currently unused, this could fuel the likes of a progress bar.
  'Dim predicted_runtime As Long: predicted_runtime = total_duration / 30  ' divisor empirically obtained from a 25-mile ride

  ' I tell Garmin the distance I accumulate second-by-second; Peloton supplies its distance.
  'Dim total_distance As Double: total_distance = CDbl(j("obj.summaries(1).value"))           ' Miles

  ' I don't think Garmin lets us tell him the kilo-joules, but its here.
  'Dim total_output As Integer: total_output = CInt(j("obj.summaries(0).value"))              ' Kjoules

  Dim total_calories As Integer: total_calories = CInt(j("obj.summaries(2).value"))           ' kcal

  Debug.Assert j("obj.metrics(0).display_name") = "Output"
  Dim powers() As Integer: powers = GetFilteredIntegers(j, "obj.metrics(0).values*")
  samples = UBound(powers)
  Debug.Assert samples <= 99999   ' maximum "points" (?) Garmin upload can contain (27 hours)
  Dim avg_power As Integer: avg_power = CInt(j("obj.metrics(0).average_value"))               ' Watts
  Dim max_power As Integer: max_power = CInt(j("obj.metrics(0).max_value"))

  Debug.Assert j("obj.metrics(1).display_name") = "Cadence"
  Dim cadences() As Integer: cadences = GetFilteredIntegers(j, "obj.metrics(1).values*")
  Debug.Assert UBound(cadences) = samples
  Dim avg_cadence As Integer: avg_cadence = CInt(j("obj.metrics(1).average_value"))           ' RPM
  Dim max_cadence As Integer: max_cadence = CInt(j("obj.metrics(1).max_value"))

  Debug.Assert j("obj.metrics(2).display_name") = "Resistance"
  Dim resistances() As Integer: resistances = GetFilteredIntegers(j, "obj.metrics(2).values*")
  Debug.Assert UBound(resistances) = samples
  Dim avg_resist As Integer: avg_resist = CInt(j("obj.metrics(2).average_value"))             ' percent
  Dim max_resist As Integer: max_resist = CInt(j("obj.metrics(2).max_value"))

  Debug.Assert j("obj.metrics(3).display_name") = "Speed"
  Dim speeds() As Double: speeds = GetFilteredDoubles(j, "obj.metrics(3).values*")
  Debug.Assert UBound(speeds) = samples
  Dim avg_speed As Double: avg_speed = CDbl(j("obj.metrics(3).average_value")) * METERS_PER_SECOND_FACTOR * 1000 ' convert MPH
  Dim max_speed As Double: max_speed = CDbl(j("obj.metrics(3).max_value")) * METERS_PER_SECOND_FACTOR * 1000     ' to meters/second
  Dim avg_speed_integral As Long: avg_speed_integral = CLng(avg_speed)
  Dim max_speed_integral As Long: max_speed_integral = CLng(max_speed)

  Dim HRs() As Integer: HRs = GetFilteredIntegers(j, "obj.metrics(4).values*")
  Dim avg_HR As Integer, max_HR As Integer
  If UBound(HRs) > 0 Then
    Debug.Assert j("obj.metrics(4).display_name") = "Heart Rate"
    Debug.Assert UBound(HRs) = samples
    hasHeartRateData = True
    avg_HR = CInt(j("obj.metrics(4).average_value"))                                        ' bpm
    max_HR = CInt(j("obj.metrics(4).max_value"))
  Else
    avg_HR = -1: max_HR = -1
  End If

'========================================================================================
'  Format a FIT file for Garmin upload
'========================================================================================

  ' -------------------------------------------------------------------------------
  ' Create file, write fake FIT header, then the 'FILE_ID' and DEVICE_INFO messages

  expandedFileName = ExpandFileName(FIT_FILE_NAME, windowsStartTime)

  Open_FIT_File FIT_File_Type.ACTIVITY, startTimestamp, _
                CInt(Environ("GARMIN_DEVICE_MODEL_NUM")), _
                getAULongFromLargeDecimal(Environ("GARMIN_DEVICE_ID_NUM")), _
                expandedFileName

 'Emit_Creator     ' Not essential, but perhaps a future enhancement
  Emit_Device_Info

  If INCLUDE_DEVELOPER_DATA_LOGIC Then

    ' -------------------------------------------------------------------
    ' DEVELOPER_DATA_ID message if we're adding at least one custom field

    ReDim tp(2): i = 0
    tp(i) = FIT_Field_STRING(1, 16): i = i + 1 ' Application ID
    tp(i) = FIT_Field_UINT32(4): i = i + 1     ' version
    tp(i) = FIT_Field_UINT08(3): i = i + 1     ' developer data index

    ReDim v(Emit_FIT_Definition_Rec(LCL_TYPE.gDEVELOPER_ID, GARMIN_TYPE.gDEVELOPER_DATA_ID, tp))

    v(3) = CInt("&H0" & Mid(APPLICATION_GUID, 1, 2))
    v(2) = CInt("&H0" & Mid(APPLICATION_GUID, 3, 2))
    v(1) = CInt("&H0" & Mid(APPLICATION_GUID, 5, 2))
    v(0) = CInt("&H0" & Mid(APPLICATION_GUID, 7, 2))
    v(5) = CInt("&H0" & Mid(APPLICATION_GUID, 10, 2))
    v(4) = CInt("&H0" & Mid(APPLICATION_GUID, 12, 2))
    v(7) = CInt("&H0" & Mid(APPLICATION_GUID, 15, 2))
    v(6) = CInt("&H0" & Mid(APPLICATION_GUID, 17, 2))
    v(9) = CInt("&H0" & Mid(APPLICATION_GUID, 20, 2))
    v(8) = CInt("&H0" & Mid(APPLICATION_GUID, 22, 2))
    v(10) = CInt("&H0" & Mid(APPLICATION_GUID, 25, 2))
    v(11) = CInt("&H0" & Mid(APPLICATION_GUID, 27, 2))
    v(12) = CInt("&H0" & Mid(APPLICATION_GUID, 29, 2))
    v(13) = CInt("&H0" & Mid(APPLICATION_GUID, 31, 2))
    v(14) = CInt("&H0" & Mid(APPLICATION_GUID, 33, 2))
    v(15) = CInt("&H0" & Mid(APPLICATION_GUID, 35, 2))
    Set_4 v, 16, APPLICATION_VSN
    v(20) = 0                                       ' first developer data index

    Emit_FIT_Data_Rec LCL_TYPE.gDEVELOPER_ID, v

    ' --------------------------------------------------
    ' FIELD DESCRIPTION message

    ReDim tp(4): i = 0
    tp(i) = FIT_Field_STRING(3, 10): i = i + 1  ' field name           *
    tp(i) = FIT_Field_STRING(8, 7): i = i + 1   ' units                *
    tp(i) = FIT_Field_UINT08(0): i = i + 1      ' Developer data index *
    tp(i) = FIT_Field_UINT08(1): i = i + 1      ' field definition #   *
    tp(i) = FIT_Field_UINT08(2): i = i + 1      ' fit base type id

    ReDim v(Emit_FIT_Definition_Rec(LCL_TYPE.gFIELD_DEF, GARMIN_TYPE.gFIELD_DESCRIPTION, tp))

    ' Its a fairly significant waste of time to perfect this now, while
    ' Garmin Connect has no capability to display the field I'm adding.

    SetString v, 0, "Undefined"   ' field name
    SetString v, 10, "Percent"    ' units
    v(17) = 0                     ' developer data index
    v(18) = 0                     ' first custom field
    v(19) = 2                     ' base type: UINT8

    Emit_FIT_Data_Rec LCL_TYPE.gFIELD_DEF, v

  End If    ' INCLUDE_DEVELOPER_DATA_LOGIC

  ' ----------------------------------------------------------------------------------------
  ' EVENT Definition and START EVENT message (Best Practice, but works w/ Garmin if omitted)

  ReDim tp(4): i = 0
  tp(i) = FIT_Field_UINT32(&HFD): i = i + 1   ' Timestamp
  tp(i) = FIT_Field_UINT32(3): i = i + 1      ' "data" ???
  tp(i) = FIT_Field_ENUMER(0): i = i + 1      ' event
  tp(i) = FIT_Field_ENUMER(1): i = i + 1      ' event type
  tp(i) = FIT_Field_UINT08(4): i = i + 1      ' event group

  numEventDataBytes = Emit_FIT_Definition_Rec(LCL_TYPE.gEVENT, GARMIN_TYPE.gEVENT, tp)

  ReDim v(numEventDataBytes)
  Set_4 v, 0, startTimestamp
  Set_4 v, 4, 0               ' data
  v(8) = 0                    ' evt
  v(9) = 0                    ' type = "Start"
  v(10) = 0                   ' group

  Emit_FIT_Data_Rec LCL_TYPE.gEVENT, v

  '==========================================================
  ' RECORD message for the actual second-by-second data points

  ReDim tp(8): i = 0
  tp(i) = FIT_Field_UINT32(&HFD): i = i + 1   ' Timestamp in Garmin's EPOC
  tp(i) = FIT_Field_UINT32(5): i = i + 1      ' Distance so far in meters * 100
  tp(i) = FIT_Field_UINT32(73): i = i + 1     ' Enhanced Speed in meters*1000 per second [replaces field #6]
  tp(i) = FIT_Field_UINT16(6): i = i + 1      ' Speed in meters*1000 per second          [legacy field]
  tp(i) = FIT_Field_UINT16(7): i = i + 1      ' Power in watts
  tp(i) = FIT_Field_UINT08(4): i = i + 1      ' Cadence in rpm
  tp(i) = FIT_Field_UINT08(10): i = i + 1     ' Resistance
  tp(i) = FIT_Field_UINT08(3): i = i + 1      ' Heart Rate in bpm
  tp(i) = FIT_Field_SINT08(13): i = i + 1     ' Temperature

  numDataBytes = Emit_FIT_Definition_Rec(LCL_TYPE.gRECORD, GARMIN_TYPE.gRECORD, tp, developerData:=INCLUDE_DEVELOPER_DATA_LOGIC)

  If INCLUDE_DEVELOPER_DATA_LOGIC Then

    ' augment the definition to include developer field(s)
    ReDim tp(0)
    tp(0).gblIndex = 0: tp(0).length = 1: tp(0).typeCode = 0  ' a one-byte field for developer index 0

    numDataBytes = numDataBytes + Emit_FIT_Developer_Definitions(tp)

  End If

  ReDim v(numDataBytes)

  ' Iterate over the Peloton data samples and emit Garmin FIT messages

  'Note: if we get merging watch data working, we might want to use
  '      the temperature recorded by the watch.  TBR
  v(19) = FIT_DATA_TEMPERATURE    ' every second will have this constant air temperature

  For sampleI = 1 To samples

    thisSpeed = speeds(sampleI) * METERS_PER_SECOND_FACTOR  ' MPH -> meters/second
    thisSpeedIntegral = thisSpeed * 1000#                   ' let it round for this purpose

    Set_4 v, 0, currentTimestamp
    Set_4 v, 4, totalDistanceIntegral    ' meters times 100 (e.g. 286 = 2.86 meters)
    Set_4 v, 8, thisSpeedIntegral        ' eventually replaces the UInt16 value
    Set_2 v, 12, CInt(thisSpeedIntegral) ' meters per second * 1000
    Set_2 v, 14, powers(sampleI)         ' watts
    v(16) = cadences(sampleI)            ' rpm
    v(17) = resistances(sampleI) * PELOTON_TO_GARMIN_RESISTANCE_FACTOR ' 1.0 or 2.54
    If hasHeartRateData Then
      v(18) = HRs(sampleI)               ' pulse
    Else
      v(18) = &HFF    ' TODO: Test to see what Garmin does with this
    End If

    Emit_FIT_Data_Rec LCL_TYPE.gRECORD, v

    ' Compute the distance for the NEXT second using this second's speed
    ' Let it round up/down (the rounding error gets lost by the next sample's arithmetic).
    totalDistance = totalDistance + (thisSpeed * 100#)  ' speed in (meters * 100) / sec
    totalDistanceIntegral = totalDistance               ' rounding to nearest millimeter - LOL

    currentTimestamp = currentTimestamp + 1

  Next sampleI

  currentTimestamp = currentTimestamp - 1   ' back out the final increment inside the loop

  ' ------------------------------------------
  ' EVENT Message for STOP-ALL (defined above)

  ReDim v(numEventDataBytes)
  Set_4 v, 0, currentTimestamp
  Set_4 v, 4, 0
  v(8) = 0
  v(9) = 4     ' STOP-ALL code
  v(10) = 0

  Emit_FIT_Data_Rec LCL_TYPE.gEVENT, v

  ' -------------------------------------------------------------
  ' LAP Message - Best Practice, but Garmin accepts files w/o any

  Dim elapsedTime As Long: elapsedTime = (currentTimestamp - startTimestamp) * 1000

  ReDim tp(21): i = 0
  tp(i) = FIT_Field_UINT32(&HFD): i = i + 1   ' Timestamp
  tp(i) = FIT_Field_UINT32(2): i = i + 1      ' start Time
  tp(i) = FIT_Field_UINT32(7): i = i + 1      ' total elapsed time
  tp(i) = FIT_Field_UINT32(8): i = i + 1      ' total timer time
  tp(i) = FIT_Field_UINT32(52): i = i + 1     ' total moving time
  tp(i) = FIT_Field_UINT32(9): i = i + 1      ' total distance
  tp(i) = FIT_Field_UINT32(110): i = i + 1    ' enhanced avg speed (eventually replaces #13)
  tp(i) = FIT_Field_UINT32(111): i = i + 1    ' enhanced max speed (eventually replaces #14)
  tp(i) = FIT_Field_UINT16(13): i = i + 1     ' avg speed  (legacy field; see #110)
  tp(i) = FIT_Field_UINT16(14): i = i + 1     ' max speed  (legacy field; see #111)
  tp(i) = FIT_Field_UINT16(19): i = i + 1     ' avg power
  tp(i) = FIT_Field_UINT16(20): i = i + 1     ' max power
  tp(i) = FIT_Field_UINT16(11): i = i + 1     ' total calories
  tp(i) = FIT_Field_ENUMER(0): i = i + 1      ' event
  tp(i) = FIT_Field_ENUMER(1): i = i + 1      ' event type
  tp(i) = FIT_Field_UINT08(15): i = i + 1     ' avg HR
  tp(i) = FIT_Field_UINT08(16): i = i + 1     ' max HR
  tp(i) = FIT_Field_UINT08(17): i = i + 1     ' avg cadence
  tp(i) = FIT_Field_UINT08(18): i = i + 1     ' max cadence
  tp(i) = FIT_Field_UINT08(24): i = i + 1     ' lap_trigger
  tp(i) = FIT_Field_ENUMER(25): i = i + 1     ' sport
  tp(i) = FIT_Field_UINT08(26): i = i + 1     ' event group

  ReDim v(Emit_FIT_Definition_Rec(LCL_TYPE.gLAP, GARMIN_TYPE.gLAP, tp))

  Set_4 v, 0, currentTimestamp                ' Mandatory
  Set_4 v, 4, startTimestamp                  ' Mandatory
  Set_4 v, 8, elapsedTime                     ' Mandatory total elapsed
  Set_4 v, 12, elapsedTime                    ' Mandatory total timer time
  Set_4 v, 16, elapsedTime                    ' total moving time
  Set_4 v, 20, totalDistanceIntegral          ' meters * 100
  Set_4 v, 24, avg_speed_integral             ' enhanced avg speed
  Set_4 v, 28, max_speed_integral             ' enhanced max speed
  Set_2 v, 32, CInt(avg_speed_integral)       ' avg speed m/s (legacy)
  Set_2 v, 34, CInt(max_speed_integral)       ' max speed m/s (legacy)
  Set_2 v, 36, avg_power                      ' avg power
  Set_2 v, 38, max_power                      ' max power
  Set_2 v, 40, total_calories                 ' [k]calories
  v(42) = 9                                   ' event = lap
  v(43) = 1                                   ' event type = stop
  v(44) = IIf(hasHeartRateData, avg_HR, &HFF) ' avg HR
  v(45) = IIf(hasHeartRateData, max_HR, &HFF) ' max HR
  v(46) = avg_cadence                         ' avg cadence
  v(47) = max_cadence                         ' max cadence
  v(48) = 0                                   ' trigger = End
  v(49) = 2                                   ' sport = cycling
  v(50) = 0                                   ' event group

  Emit_FIT_Data_Rec LCL_TYPE.gLAP, v

  numLapsEmitted = numLapsEmitted + 1

  ' ---------------------------------------
  '  SESSION message - Mandatory for Garmin

  ReDim tp(26): i = 0
  tp(i) = FIT_Field_UINT32(&HFD): i = i + 1   ' Mandatory Timestamp
  tp(i) = FIT_Field_UINT32(2): i = i + 1      ' Mandatory Start Time
  tp(i) = FIT_Field_UINT32(7): i = i + 1      ' Mandatory elapsed Time
  tp(i) = FIT_Field_UINT32(8): i = i + 1      ' Mandatory timer time
  tp(i) = FIT_Field_UINT32(9): i = i + 1      ' distance
  tp(i) = FIT_Field_UINT32(124): i = i + 1    ' enhanced avg speed (eventually replaces #14)
  tp(i) = FIT_Field_UINT32(125): i = i + 1    ' enhanced max speed (eventually replaces #15)
  tp(i) = FIT_Field_UINT16(14): i = i + 1     ' avg speed (legacy field, see #?)
  tp(i) = FIT_Field_UINT16(15): i = i + 1     ' max speed (legacy field, see #?)
  tp(i) = FIT_Field_UINT16(20): i = i + 1     ' avg power
  tp(i) = FIT_Field_UINT16(21): i = i + 1     ' max power
  tp(i) = FIT_Field_UINT16(22): i = i + 1     ' ascent
  tp(i) = FIT_Field_UINT16(23): i = i + 1     ' descent
  tp(i) = FIT_Field_UINT16(25): i = i + 1     ' lap1 index=0
  tp(i) = FIT_Field_UINT16(26): i = i + 1     ' num_laps=1
  tp(i) = FIT_Field_UINT16(11): i = i + 1     ' calories
  tp(i) = FIT_Field_UINT16(&HFE): i = i + 1   ' message_indx (dunno)
  tp(i) = FIT_Field_ENUMER(0): i = i + 1      ' event=8=session
  tp(i) = FIT_Field_ENUMER(1): i = i + 1      ' event_type=1
  tp(i) = FIT_Field_ENUMER(5): i = i + 1      ' sport=2=cycling
  tp(i) = FIT_Field_ENUMER(6): i = i + 1      ' sub_sport=6=indoor
  tp(i) = FIT_Field_UINT08(16): i = i + 1     ' avg HR
  tp(i) = FIT_Field_UINT08(17): i = i + 1     ' max HR
  tp(i) = FIT_Field_UINT08(18): i = i + 1     ' avg cadence
  tp(i) = FIT_Field_UINT08(19): i = i + 1     ' max cadence
  tp(i) = FIT_Field_UINT08(27): i = i + 1     ' event grp=0
  tp(i) = FIT_Field_UINT08(28): i = i + 1     ' trigger=0=activity_end
 'tp(i) = FIT_Field_STRING(&H6E): i = i + 1   ' Garmin currently ignores activity name

  ReDim v(Emit_FIT_Definition_Rec(LCL_TYPE.gSESSION, GARMIN_TYPE.gSESSION, tp))

  Set_4 v, 0, currentTimestamp                ' Mandatory
  Set_4 v, 4, startTimestamp                  ' Mandatory
  Set_4 v, 8, elapsedTime                     ' Mandatory total elapsed
  Set_4 v, 12, elapsedTime                    ' Mandatory total timer time
  Set_4 v, 16, totalDistanceIntegral          ' meters * 100
  Set_4 v, 20, avg_speed_integral             ' enhanced avg speed m/s
  Set_4 v, 24, max_speed_integral             ' enhanced max speed m/s
  Set_2 v, 28, CInt(avg_speed_integral)       ' avg speed m/s (legacy field)
  Set_2 v, 30, CInt(max_speed_integral)       ' max speed m/s (legacy field)
  Set_2 v, 32, avg_power                      ' avg power
  Set_2 v, 34, max_power                      ' max power
  Set_2 v, 36, 0                              ' ascent
  Set_2 v, 38, 0                              ' descent
  Set_2 v, 40, 0                              ' first lap index
  Set_2 v, 42, numLapsEmitted                 ' num laps
  Set_2 v, 44, total_calories                 ' [k]calories
  Set_2 v, 46, 0                              ' message index
  v(48) = 8                                   ' event = session
  v(49) = 1                                   ' event_type = stop
  v(50) = 2                                   ' sport = cycling
  v(51) = 5                                   ' sub_sport = SPIN (indoor bike = 6)
  v(52) = IIf(hasHeartRateData, avg_HR, &HFF) ' avg HR
  v(53) = IIf(hasHeartRateData, max_HR, &HFF) ' max HR
  v(54) = avg_cadence                         ' avg cadence
  v(55) = max_cadence                         ' max cadence
  v(56) = 0                                   ' event group
  v(57) = 0                                   ' trigger = activity End
  'SetString v, 58, "Peloton Ride"            ' Garmin currently ignores what we send

  Emit_FIT_Data_Rec LCL_TYPE.gSESSION, v

  ' ----------------------------------------
  '  ACTIVITY message - Mandatory for Garmin

  ReDim tp(7): i = 0
  tp(i) = FIT_Field_UINT32(&HFD): i = i + 1   ' Mandatory timestamp
  tp(i) = FIT_Field_UINT32(0): i = i + 1      ' Mandatory elapsed time
  tp(i) = FIT_Field_UINT32(5): i = i + 1      ' local timestamp
  tp(i) = FIT_Field_UINT16(1): i = i + 1      ' num sessions
  tp(i) = FIT_Field_ENUMER(2): i = i + 1      ' type
  tp(i) = FIT_Field_ENUMER(3): i = i + 1      ' event
  tp(i) = FIT_Field_ENUMER(4): i = i + 1      ' event_type
  tp(i) = FIT_Field_UINT08(6): i = i + 1      ' event_group

  ReDim v(Emit_FIT_Definition_Rec(LCL_TYPE.gACTIVITY, GARMIN_TYPE.gACTIVITY, tp))

  Set_4 v, 0, currentTimestamp ' Mandatory
  Set_4 v, 4, elapsedTime      ' Mandatory Elapsed Time in seconds * 1000
  Set_4 v, 8, 0                ' local timestamp
  Set_2 v, 12, 1               ' num_sessions
  v(14) = 0                    ' type (manual)
  v(15) = 26                   ' event (activity)
  v(16) = 1                    ' event_type (stop)
  v(17) = 0                    ' event_group

  Emit_FIT_Data_Rec LCL_TYPE.gACTIVITY, v

  ' ----------------------------------------------------------------------------------------
  ' Rewrite the FIT Header with checksum, write the final file checksum, and close the file.
  ' No further file operations will work until another Open_FIT_File() is performed.

  Close_FIT_File

  Dim mEnd As Date: mEnd = Now()
  Debug.Print mEnd & " FIT generation complete (" & (mEnd - mStart) * 86400 & ") seconds"

  ' --------------------------
  ' Upload that file to Garmin

  UploadFITFile expandedFileName

  Dim mEnd2 As Date: mEnd2 = Now()
  Debug.Print mEnd2 & " Upload Complete (" & (mEnd2 - mEnd) * 86400 & ") seconds"

  Application.Cursor = xlDefault
  Application.ScreenUpdating = True
  Application.Calculation = wasCalculating

  MsgBox "Conversion & Upload complete " & vbLf & " file is " & expandedFileName
End Sub

Public Function ExpandFileName(inString As String, timestamp As Date) As String
  Dim v() As String, s As String: s = inString
  Dim token As String, i As Integer
  ' the filename can be patterned to include date and/or time
  i = InStr(s, "~")
  If i > 0 Then
    token = Mid(s, i, 32): token = Left(token, InStr(2, token, "~"))
    s = Replace(s, token, Format(timestamp, Mid(token, 2, Len(token) - 2)))
  End If
  ' tokens delimited by '%' are environment variables to be expanded
  v = Split(s, "%")
  While UBound(v) > 0
    s = Replace(s, "%" & v(1) & "%", Environ(v(1))) ' replace all occurrences
    v = Split(s, "%")                               ' look for more variables
  Wend
  ExpandFileName = s   ' return the substitution string
End Function

Private Function getAULongFromLargeDecimal(s As String) As Long
  Dim s1 As LongLong: s1 = CLngLng(s)
  Dim s2 As Long: CopyMemory s2, s1, 4
  getAULongFromLargeDecimal = s2
End Function

' convert a 8-char hex string to an unsigned long - this function isn't used right now
'Private Function getAULongFromHexString(hexString As String) As Long
'  Dim xr As String: xr = Mid(hexString, 2)
'  Select Case Left(hexString, 1)
'    Case "8": getAULongFromHexString = (CLng("&h" & xr)) Or &H80000000
'    Case "9": getAULongFromHexString = (CLng("&h1" & xr)) Or &H80000000
'    Case "A": getAULongFromHexString = (CLng("&h2" & xr)) Or &H80000000
'    Case "B": getAULongFromHexString = (CLng("&h3" & xr)) Or &H80000000
'    Case "C": getAULongFromHexString = (CLng("&h4" & xr)) Or &H80000000
'    Case "D": getAULongFromHexString = (CLng("&h5" & xr)) Or &H80000000
'    Case "E": getAULongFromHexString = (CLng("&h6" & xr)) Or &H80000000
'    Case "F": getAULongFromHexString = (CLng("&h7" & xr)) Or &H80000000
'    Case Else: getAULongFromHexString = CLng("&h" & hexString)
'  End Select
'End Function

Public Sub ExploreDownloadsFolder()
  Dim s As String
  s = Left(FIT_FILE_NAME, InStrRev(FIT_FILE_NAME, "\") - 1)
  s = Replace(s, "%USERPROFILE%", Environ("USERPROFILE"))
  Shell "explorer.exe """ & s & """", vbNormalFocus
End Sub
