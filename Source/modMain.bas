Attribute VB_Name = "modMain"
Option Explicit
Public Const AppName = "TemperatureMonitor"
Public Const VersionDate As String = "05-Feb-16"
Public Const CurrentDBversion As Long = 10
Public Const BeginPacket As String = "^"

'error descriptions
Public Const ErrNum = "A number is required."
Public Const ErrDate = "A date is required."
Public Const ErrBoo = "True/False is required."
Public Const InputErr = vbObjectError + 1001    'global error #
Public Const ErrPassword = vbObjectError + 1002
Public Const ErrDBtype = vbObjectError + 1003     'wrong database type
Public Const ErrLowVersion = vbObjectError + 1004 'database version too low
Public Const ErrHighVersion = vbObjectError + 1005 'database version too high
Public Const ErrFileNotFound = vbObjectError + 1006  'database file doesn't exist

Public Enum GMStorageTypes
    GMSTall = 0
    GMSTBins = 1
    GMSTWarehouses = 2
End Enum

Public Enum SensorStatus
    ssEnabled
    ssDisabled
    ssError
    ssAlarmDisabled
    ssAlarmCondition
End Enum

Public Enum SensorTypes
    stTemperature
End Enum

'// packet description:
'// start,packet type,break,data,break,sensor Rom Code,break
'// packet types:                                                examples:
'// 0  heartbeat to signal still connected                       ^0|||       heartbeat
'// 1  command sensors to report either 0 or a specific Mac      ^1||0|      all sensors report
'// 2  set sensor number                                         ^2|6|Mac|   set sensor with address Mac to sensor # 6
'// 3  set bin number                                            ^3|12|Mac|  set sensor with address Mac to bin # 12
'// 4  get sensor number                                         ^4||Mac|    get sensor number for sensor with adddress Mac
'// 5  get bin number                                            ^5||Mac|    get bin number for sensor with address Mac
'// 6  get signal strength                                       ^6|||       get wifi signal strength
'// 7  get board mac address                                     ^7|||       get 'The Thing' Mac address


Public Enum PacketData
    pdHeartBeat
    pdGetTemperatures
    pdSetSenNum
    pdSetBinNum
    pdGetSenNum
    pdGetBinNum
    pdGetSignalStrength
    pdGetClientMac
End Enum

Public Type Packet
    DataType As PacketData
    DataString As String
    DataID As Long  'control index ID
End Type

Private Type tagInitCommonControlsEx
   lngSize As Long
   lngICC As Long
End Type

Private Declare Function InitCommonControlsEx Lib "comctl32.dll" _
   (iccex As tagInitCommonControlsEx) As Boolean
Private Const ICC_USEREX_CLASSES = &H200

Public AD As clsAppData
Public LastFolder As String
Public MainDB As DAO.Database
Public Prog As clsMain
Public DBconnected As Boolean

Private Function InitCommonControlsVB() As Boolean
   On Error Resume Next
   Dim iccex As tagInitCommonControlsEx
   ' Ensure CC available:
   With iccex
       .lngSize = LenB(iccex)
       .lngICC = ICC_USEREX_CLASSES
   End With
   InitCommonControlsEx iccex
   InitCommonControlsVB = (Err.Number = 0)
   On Error GoTo 0
End Function

Private Sub Main()
    If App.PrevInstance Then End
    LoadHelp
    InitCommonControlsVB
    Set AD = New clsAppData
    LastFolder = AD.Folders(App_Folders_Database)
    BuildNameLists
    frmStart.Show
ErrExit:
    Exit Sub
ErrHandler:
    AD.DisplayError Err.Number, "modMain", "Main", Err.Description
    Resume ErrExit
End Sub

