Attribute VB_Name = "FIT_Protocol"
Option Explicit

Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As Any, src As Any, ByVal length As Long)
Public Declare PtrSafe Sub FillMemory Lib "kernel32" Alias "RtlFillMemory" (dest As Any, ByVal size As Long, ByVal fill As Byte)

Public Enum ARCH
  LITTLE = 0        ' Garmin watches are little-endian and we want GC to believe the watch generated the FIT data
  BIG = 1           ' Windows on Intel is big-endian
End Enum
Public Const ARCHITECTURE As Integer = ARCH.LITTLE    ' Don't change without testing; GUID handling almost certainly WRONG

Public Type FIT_Definition_Element
  gblIndex As Byte
  typeCode As Byte
  length As Byte
End Type

Public Enum FIT_File_Type  ' Folder       SDK_name
  FILE_ID = 0
  DEVICE = 1               ' GARMIN       CAPABILITIES
  SETTINGS = 2             ' Settings     DEVICE_SETTINGS
  SPORT = 3                ' Sports       USER_PROFILE
  ACTIVITY = 4             ' Activites    HRM_PROFILE
  WORKOUT = 5              ' Workouts     SDM_PROFILE
  COURSE = 6               ' Courses      BIKE_PROFILE
  SCHEDULES = 7            ' Schedules    ZONES_TARGET
  LOCATIONS = 8            ' Locations    HR_ZONE
  WEIGHT_ = 9              ' Weight       POWER_ZONE
  TOTALS = 10              ' Totals       MET_ZONE
  GOALS = 11               ' Goals
  SPORT_2 = 12             ' [from Profile.xlsx]
  BLOOD_PRESSURE = 14      ' Blood Pressure
  MONITORING_A = 15        ' Monitoring   GOAL
  SESSION = 18             ' [from Profile.xlsx]
  LAP = 19                 ' [from Profile.xlsx]
  ACTIVITY_SUMMARY = 20    ' Activities   RECORD
  MONITORING_DAILY = 28    '              SCHEDULE
  MONITORING_B = 32        ' Monitoring   COURSE_POINT
  SEGMENT = 34             ' Segments     ACTIVITY
  SEGMENT_LIST = 35        ' Segments     SOFTWARE
  EXD_CONFIGURATION = 40   ' Settings
  MFG_MIN = &HF7           ' here to &HFE is for manufacturer stuff (?)
End Enum

Public Enum GARMIN_TYPE    ' Global Message Types
  gFILE_ID = 0
  gDEVICE_SETTINGS = 2
  gUSER_PROFILE = 3
  gZONES_TARGET = 7
  gSPORT = 12
  gSESSION = 18
  gLAP = 19
  gRECORD = 20
  gEVENT = 21
  gDEVICE_INFO = 23
  gACTIVITY = 34
  gCREATOR = 49
  gGPS_META = 160
  gFIELD_DESCRIPTION = 206
  gDEVELOPER_DATA_ID = 207
End Enum

Public Enum FIT_Base_Type
  ENUMERATION = 0
  SINT8 = 1
  UINT8 = 2
  UINT8Z = 10
  SINT16 = &H83
  UINT16 = &H84
  UINT16Z = &H8B
  SINT32 = &H85
  UINT32 = &H86
  UINT32Z = &H8C
  STRNG = 7
  FLOAT32 = &H88
  FLOAT64 = &H89
End Enum

' Garmin EPOC 31-Dec-1989; Peloton EPOC 1-Jan-1970
Public Const GARMIN_EPOC_BIAS As Long = 631065600  ' seconds to SUBTRACT from Peloton times

Private Const FIT_MANUFACTURER As Integer = 1    ' 255 = "MANUFACTURER_DEVELOPMENT" (&H1 is Garmin)
Private Const FIT_PROTOCOL_VSN As Byte = 32      ' SDK fit.h: FIT_PROTOCOL_VERSION_20 (2 << 4)
Private Const FIT_PROFILE_VSN As Integer = 2183  ' SDK fit.h: FIT_PROFILE_VERSION

Public Type endian_flipper
  b0 As Byte: b1 As Byte: b2 As Byte: b3 As Byte
End Type

Private fileno As Integer, filePos As Long
Private vData() As Byte, numBytes As Long
Private crc_tbl(15) As Long, crc As Long           ' low 16 bits treated as as UINT16
Private saved_device_product_code As Integer
Private saved_device_serial_num As Long
Private saved_file_creation_timestamp As Long

Public Sub Open_FIT_File(fileTypeCode As FIT_File_Type, creationTimestamp As Long, _
                         garmin_device_product_code As Integer, garmin_device_serial_num As Long, _
                         fName As String)

  Call Initialize_CRC_table   ' VBA can't init array content at declaration time
  crc = 0

  ' set these aside so we can re-use them in Emit_Device_Info()
  saved_file_creation_timestamp = creationTimestamp
  saved_device_product_code = garmin_device_product_code
  saved_device_serial_num = garmin_device_serial_num

  ' Left alone, VBA writing less data to a pre-existing file leaves the old file's size intact.
  ' There's a sophisticated approach using Win32 API to truncate, but that's just silly for this.
  On Error Resume Next
  Kill fName
  On Error GoTo 0

  fileno = FreeFile
  Open fName For Binary Access Write As #fileno
  filePos = 1

  ' Emit 14 placeholder bytes for the FIT header we'll output later.
  ReDim vData(13)
  For numBytes = 0 To 13: vData(numBytes) = 0: Next numBytes
  Put #fileno, 1, vData: filePos = filePos + 14

  Dim tp() As FIT_Definition_Element, i As Integer

  ReDim tp(5): i = 0                         ' FILE_ID definition and data messages
  tp(i) = FIT_Field_UINT32Z(3): i = i + 1    ' serial# (zero is illegal)
  tp(i) = FIT_Field_UINT32(4): i = i + 1     ' creation time
  tp(i) = FIT_Field_UINT16(1): i = i + 1     ' manufacturer
  tp(i) = FIT_Field_UINT16(2): i = i + 1     ' product code
  tp(i) = FIT_Field_UINT16(5): i = i + 1     ' number (not used for activity files)
  tp(i) = FIT_Field_ENUMER(0) ': i = i + 1   ' file type code

  ReDim vData(Emit_FIT_Definition_Rec(GARMIN_TYPE.gFILE_ID, GARMIN_TYPE.gFILE_ID, tp))

  ' Its a fairly well-known thing that a Garmin device must [appear to] have recorded the activity
  ' for it to 'count' when it comes to challenges. I'm not wild about this sort of misrepresentation
  ' built-into a system, but having the distance, time, and exertion of Peloton rides 'count'
  ' in the Garmin world almost demands this kind of system-spoofing.

  Set_4 vData, 0, garmin_device_serial_num
  Set_4 vData, 4, creationTimestamp
  Set_2 vData, 8, FIT_MANUFACTURER
  Set_2 vData, 10, garmin_device_product_code
  Set_2 vData, 12, &HFFFF
  vData(14) = fileTypeCode     ' Caller specified from the FIT_File_Type enum

  Emit_FIT_Data_Rec GARMIN_TYPE.gFILE_ID, vData

End Sub

Public Sub Emit_Device_Info(Optional unused As Long)

  Dim tp(6) As FIT_Definition_Element, i As Integer: i = 0

  tp(i) = FIT_Field_UINT32(&HFD): i = i + 1    ' timestamp
  tp(i) = FIT_Field_UINT32Z(3): i = i + 1      ' serial#
  tp(i) = FIT_Field_UINT16(2): i = i + 1       ' manufacturer
  tp(i) = FIT_Field_UINT16(4): i = i + 1       ' product
 'tp(i) = FIT_Field_UINT16(5): i = i + 1       ' software - all indications are that this is optional
  tp(i) = FIT_Field_UINT08(0): i = i + 1       ' device index
  tp(i) = FIT_Field_UINT08(1): i = i + 1       ' device type
  tp(i) = FIT_Field_ENUMER(25) ': i = i + 1    ' source type

  ReDim vData(Emit_FIT_Definition_Rec(2, GARMIN_TYPE.gDEVICE_INFO, tp))

  Set_4 vData, 0, saved_file_creation_timestamp
  Set_4 vData, 4, saved_device_serial_num
  Set_2 vData, 8, FIT_MANUFACTURER      ' GARMIN
  Set_2 vData, 10, saved_device_product_code
  vData(12) = 0                         ' device index
  vData(13) = &HFF                      ' no device type here
  vData(14) = 5                         ' source type

  Emit_FIT_Data_Rec 2, vData

  Set_4 vData, 4, 0                     ' no serial#
  vData(12) = 1                         ' device index
  vData(13) = 4                         ' device type

  Emit_FIT_Data_Rec 2, vData

  'vData(12) = 2                        ' next assignable device index

'  Peloton Interactive doesn't have an assigned ID as a manufacturer
'  HRM-PRO has device number 3300, but that's somewhat unique to me
'  Specialized has a manufacturer code, but that's just my Peloton shoes

'= TYPE=6 NAME=device_info NUMBER=23
'--- timestamp=1032360847=2022-09-17T14:54:07Z
'--- serial_number=3369083060=3369083060       ' HRM-PRO
'--- cum_operating_time=108554=108554 s
'--- unknown15=3=3
'--- unknown24=24649908=24649908
'--- manufacturer=1=garmin
'--- garmin_product=3300=3300                  ' HRM-PRO product code
'--- software_version=880=8.80
'--- battery_voltage=725=2.83 V
'--- device_index=2=device2                     ' index 2 = "device 2" ????
'--- antplus_device_type=120=heart_rate         ' antplus device type is heart_rate
'--- hardware_version=66=66
'--- battery_status=3=ok
'--- ant_network=1=antplus
'--- source_type=1=antplus
'--- xxx29=76=76,116,41,25,188,208
'--- xxx30=1=1

End Sub

Public Function Emit_FIT_Definition_Rec(recType As Byte, globalMsgIndex As Integer, _
                                        types() As FIT_Definition_Element, _
                                        Optional developerData As Boolean = False) As Integer

  Dim lastElement As Integer: lastElement = UBound(types)
  Dim numElements As Integer: numElements = lastElement + 1
  Dim i As Integer, j As Integer, numDataBytes As Integer: numDataBytes = 0
  numBytes = (numElements * 3) + 6: ReDim vData(numBytes - 1)

  vData(0) = IIf(developerData, &H60, &H40) Or recType  ' Definition
  vData(1) = 0                                          ' Reserved
  vData(2) = ARCHITECTURE                               ' Endian indicator
  Set_2 vData, 3, globalMsgIndex
  vData(5) = numElements                                ' # fields
  j = 6
  For i = 0 To lastElement
    vData(j) = types(i).gblIndex: vData(j + 1) = types(i).length: vData(j + 2) = types(i).typeCode
    j = j + 3: numDataBytes = numDataBytes + types(i).length
  Next i

  Put #fileno, filePos, vData: UpdateCkSum vData: filePos = filePos + numBytes

  Emit_FIT_Definition_Rec = numDataBytes - 1   ' tell caller how big his data array must be

End Function

Public Function Emit_FIT_Developer_Definitions(types() As FIT_Definition_Element) As Integer

  Dim lastElement As Integer: lastElement = UBound(types)
  Dim numElements As Integer: numElements = lastElement + 1
  Dim i As Integer, j As Integer, numDataBytes As Integer: numDataBytes = 0
  numBytes = (numElements * 3) + 1: ReDim vData(numBytes - 1)

  vData(0) = numElements
  j = 1
  For i = 0 To lastElement
    vData(j) = types(i).gblIndex
    vData(j + 1) = types(i).length
    vData(j + 2) = types(i).typeCode
    j = j + 3
    numDataBytes = numDataBytes + types(i).length
  Next i

  Put #fileno, filePos, vData: UpdateCkSum vData: filePos = filePos + numBytes

  Emit_FIT_Developer_Definitions = numDataBytes

End Function

Public Sub Emit_FIT_Data_Rec(recType As Byte, v() As Byte)

  Put #fileno, filePos, recType
  Compute_Checksum recType
  Put #fileno, filePos + 1, v
  'UpdateCkSum v                ' worked fine; the below is slightly faster
  Dim i As Integer, n As Integer: n = UBound(v)
  For i = 0 To n: Compute_Checksum v(i): Next i
  filePos = filePos + n + 2

End Sub

Public Sub Close_FIT_File(Optional unused As Long)

  Dim dataLen As Long: dataLen = filePos - 14 - 1  ' not counting the header nor the trailing file CRC
  Dim crcShort As Integer

  ' Output the file checksum at the very end-of-file, following the data messages
  ReDim vData(1)
  CopyMemory crcShort, crc, 2
  Set_2 vData, 0, crcShort
  Put #fileno, filePos, vData

  crc = 0
  ' Write the FIT header record at the beginning of the file
  numBytes = 12: ReDim vData(numBytes - 1)
  vData(0) = 14                           ' I use the "newer" 14-byte header
  vData(1) = FIT_PROTOCOL_VSN             ' vsn 2.0 is (2<<4) = 32
  Set_2 vData, 2, FIT_PROFILE_VSN
  Set_4 vData, 4, dataLen
  vData(8) = &H2E: vData(9) = &H46: vData(10) = &H49: vData(11) = &H54 ' ".FIT"
  Put #fileno, 1, vData                   ' first 12 bytes
  UpdateCkSum vData

  ReDim vData(1)
  CopyMemory crcShort, crc, 2
  Set_2 vData, 0, crcShort
  Put #fileno, 13, vData                  ' add the FIT header checksum

  Close #fileno
  fileno = -1     ' make it invalid until next Open_FIT_File() call

End Sub

Public Function FIT_Field_UINT32(gblIndex As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = 4: x.typeCode = FIT_Base_Type.UINT32
  FIT_Field_UINT32 = x
End Function
Public Function FIT_Field_UINT32Z(gblIndex As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = 4: x.typeCode = FIT_Base_Type.UINT32Z
  FIT_Field_UINT32Z = x
End Function
Public Function FIT_Field_UINT16(gblIndex As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = 2: x.typeCode = FIT_Base_Type.UINT16
  FIT_Field_UINT16 = x
End Function
Public Function FIT_Field_UINT16Z(gblIndex As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = 2: x.typeCode = FIT_Base_Type.UINT16Z
  FIT_Field_UINT16Z = x
End Function
Public Function FIT_Field_UINT08(gblIndex As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = 1: x.typeCode = FIT_Base_Type.UINT8
  FIT_Field_UINT08 = x
End Function
Public Function FIT_Field_SINT08(gblIndex As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = 1: x.typeCode = FIT_Base_Type.SINT8
  FIT_Field_SINT08 = x
End Function
Public Function FIT_Field_UINT08Z(gblIndex As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = 1: x.typeCode = FIT_Base_Type.UINT8Z
  FIT_Field_UINT08Z = x
End Function
Public Function FIT_Field_ENUMER(gblIndex As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = 1: x.typeCode = FIT_Base_Type.ENUMERATION
  FIT_Field_ENUMER = x
End Function
Public Function FIT_Field_STRING(gblIndex As Integer, sLen As Integer) As FIT_Definition_Element
  Dim x As FIT_Definition_Element
  x.gblIndex = gblIndex: x.length = sLen: x.typeCode = FIT_Base_Type.STRNG
  FIT_Field_STRING = x
End Function

Private Sub Initialize_CRC_table()  ' crc_table is an array of longs (32-bit)
  crc_tbl(0) = 0: crc_tbl(1) = &HCC01&: crc_tbl(2) = &HD801&: crc_tbl(3) = &H1400&
  crc_tbl(4) = &HF001&: crc_tbl(5) = &H3C00&: crc_tbl(6) = &H2800&: crc_tbl(7) = &HE401&
  crc_tbl(8) = &HA001&: crc_tbl(9) = &H6C00&: crc_tbl(10) = &H7800&: crc_tbl(11) = &HB401&
  crc_tbl(12) = &H5000&: crc_tbl(13) = &H9C01&: crc_tbl(14) = &H8801&: crc_tbl(15) = &H4400&
End Sub

Private Sub Compute_Checksum(b As Byte)

  Dim tmp As Long  ' crc_tbl and crc itself are 32-bit longs, but high 2 bytes should always be zero

  tmp = crc_tbl(crc And &HF)
  crc = (crc And &HFFF0) / 16              ' high 3 nibbles of USHORT >> 4
  crc = crc Xor tmp Xor crc_tbl(b And &HF) ' low nibble of data byte as an index

  tmp = crc_tbl(crc And &HF)
  crc = (crc And &HFFF0) / 16                      ' high 3 nibbles of USHORT >> 4
  crc = crc Xor tmp Xor crc_tbl((b And &HF0) / 16) ' high nibble of data byte as an index

End Sub

Private Sub UpdateCkSum(v() As Byte)
  Dim i As Integer, n As Integer: n = UBound(v)
  For i = 0 To n: Compute_Checksum v(i): Next i
End Sub

Public Sub Set_2(ByRef v() As Byte, i As Integer, val As Integer)
  Dim x As endian_flipper
  CopyMemory x, val, 2
  If ARCHITECTURE = ARCH.BIG Then
    v(i) = x.b1
    v(i + 1) = x.b0
  Else                ' LITTLE, like Garmin watches
    v(i) = x.b0
    v(i + 1) = x.b1
  End If
End Sub

Public Sub Set_4(ByRef v() As Byte, i As Integer, val As Long)
  Dim x As endian_flipper
  CopyMemory x, val, 4
  If ARCHITECTURE = ARCH.BIG Then
    v(i) = x.b3
    v(i + 1) = x.b2
    v(i + 2) = x.b1
    v(i + 3) = x.b0
  Else                ' LITTLE, like Garmin watches
    v(i) = x.b0
    v(i + 1) = x.b1
    v(i + 2) = x.b2
    v(i + 3) = x.b3
  End If
End Sub

Public Sub SetString(ByRef v() As Byte, i As Integer, val As String)
  Dim j As Integer
  For j = 1 To Len(val)
    v(i) = Asc(Mid(val, j, 1)): i = i + 1
  Next j
End Sub

Public Sub DecodeWatchInfoFromFITfile()

  Dim bytArr(128) As Byte, i As Integer, j As Integer, fileno As Integer: fileno = FreeFile
  Dim IDdefIndex As Integer, IDdataIndex As Integer, accumulated_offset As Integer
  Dim watch_id_num As Long, watch_prod_code As Long
  Dim f As endian_flipper, bigEndian As Boolean
  Dim selectedFile As String, watch_id As String

  With Application.FileDialog(msoFileDialogFilePicker)
    .AllowMultiSelect = False
    .InitialFileName = Environ("USERPROFILE") & "\Downloads"
    .Filters.Clear
    .Filters.Add "FIT Files", "*.FIT"
    .Show
    If .SelectedItems.Count < 1 Then Exit Sub
    selectedFile = .SelectedItems(1)
  End With

  Open selectedFile For Binary Access Read As #fileno
  Get #fileno, , bytArr   ' (0) tells us the header record length
  Close #fileno

  IDdefIndex = bytArr(0)                                      ' where FILE_ID definition begins
  bigEndian = bytArr(IDdefIndex + 2) <> ARCH.LITTLE           ' We anticipate Garmin watches being LITTLE
  IDdataIndex = IDdefIndex + (bytArr(IDdefIndex + 5) * 3) + 6 ' where FILE_ID data begins

  accumulated_offset = 1
  For i = IDdefIndex + 6 To IDdataIndex - 1 Step 3
    Select Case bytArr(i)
      Case 0:                                   ' FIT file type
        If bytArr(IDdataIndex + accumulated_offset) <> FIT_File_Type.ACTIVITY Then
          MsgBox "FIT file is not of type ACTIVITY", vbOKOnly, "Error": Exit Sub
        End If
      Case 2:                                   ' product code
        j = IDdataIndex + accumulated_offset
        If bigEndian Then
          f.b1 = bytArr(j): f.b0 = bytArr(j + 1)
        Else
          f.b0 = bytArr(j): f.b1 = bytArr(j + 1)
        End If
        CopyMemory watch_prod_code, f, 2
      Case 3:                                   ' ID number
        j = IDdataIndex + accumulated_offset
        If bigEndian Then
          f.b3 = bytArr(j): f.b2 = bytArr(j + 1): f.b1 = bytArr(j + 2): f.b0 = bytArr(j + 3)
        Else
          f.b0 = bytArr(j): f.b1 = bytArr(j + 1): f.b2 = bytArr(j + 2): f.b3 = bytArr(j + 3)
        End If
        CopyMemory watch_id_num, f, 4
    End Select
    accumulated_offset = accumulated_offset + bytArr(i + 1)
  Next i

  If watch_id_num < 0 Then      ' high-order bit (#31) on?
    watch_id = StringDecimalAddition("2147483648", CStr(watch_id_num And &H7FFFFFFF))
  Else
    watch_id = CStr(watch_id_num)
  End If

  If InStr(Application.Name, "Excel") > 0 Then
    With ActiveSheet
      .Range("D22").Value = "Watch Product Code:"
      .Range("E22").Value = CStr(watch_prod_code)
      .Range("D23").Value = "Watch Identity:"
      .Range("E23").Value = watch_id
    End With
  Else
    MsgBox "Watch Prouct Code: " & CStr(watch_prod_code) & vbLf & _
           "Watch Identity: " & watch_id, vbOKOnly, "Results"
  End If
End Sub

' The year: 2022.
' The challenge: Adding two base-10 integers larger than 2 billion in a MS Office application
' The solution: Elementary school blackboard addition [ay yi fuckin' ay]
Private Function StringDecimalAddition(m As String, n As String) As String
  Dim vM As Integer, vN As Integer: vM = Len(m): vN = Len(n)
  Dim i As Integer, iSum As Integer, carry As Integer: carry = 0
  Const ZILCHES As String = "0000000000000000000"
  ' Left pad the shorter string with zeros
  Dim digits As Integer: digits = vM: If vN > vM Then digits = vN
  If vM < digits Then m = Left(ZILCHES, digits - vM) & m
  If vN < digits Then n = Left(ZILCHES, digits - vN) & n
  For i = digits To 1 Step -1
    iSum = CInt(Mid(m, i, 1)) + CInt(Mid(n, i, 1)) + carry
    If iSum > 9 Then carry = 1 Else carry = 0
    StringDecimalAddition = Right(CStr(iSum), 1) & StringDecimalAddition
  Next i
  If carry = 1 Then StringDecimalAddition = "1" & StringDecimalAddition
End Function

