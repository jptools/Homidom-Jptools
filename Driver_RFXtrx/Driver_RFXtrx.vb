﻿'Option Strict On
Imports HoMIDom
Imports HoMIDom.HoMIDom.Server
Imports HoMIDom.HoMIDom.Device
Imports STRGS = Microsoft.VisualBasic.Strings
Imports VB = Microsoft.VisualBasic
Imports System.IO.Ports
Imports System.Math
Imports System.Net.Sockets
Imports System.Threading
Imports System.Globalization
Imports System.Text
Imports System.IO
Imports System.Media

' Auteur : David
' Date : 21/09/2014
' Version from RFXCOM : Version 12.0.0.13     16-09-2014
'------------------------------------------------------------------------------------
'                          Protocol License Agreement                      
'                                                                    
' The RFXtrx protocols are owned by RFXCOM, and are protected under applicable
' copyright laws.
'
' ==================================================================================
' = It is only allowed to use this software or any part of it for RFXCOM products. =
' ==================================================================================
'
' The above Protocol License Agreement and the permission notice shall be included
' in all software using the RFXtrx protocols.
'
' Any use in violation of the foregoing restrictions may subject the user to criminal
' sanctions under applicable laws, as well as to civil liability for the breach of the
' terms and conditions of this license.
'-------------------------------------------------------------------------------------


''' <summary>Class Driver_RFXTrx, permet de communiquer avec le RFXtrx Ethernet/COM</summary>
''' <remarks>Pour la version USB, necessite l'installation du driver USB RFXCOM</remarks>
<Serializable()> Public Class Driver_RFXtrx
    Implements HoMIDom.HoMIDom.IDriver

#Region "Variables génériques"
    '!!!Attention les variables ci-dessous doivent avoir une valeur par défaut obligatoirement
    'aller sur l'adresse http://www.somacon.com/p113.php pour avoir un ID
    Dim _ID As String = "3D9D5D42-475B-11E1-B117-64314824019B"
    Dim _Nom As String = "RFXtrx"
    Dim _Enable As Boolean = False
    Dim _Description As String = "RFXtrx USB/Ethernet Interface"
    Dim _StartAuto As Boolean = False
    Dim _Protocol As String = "RF"
    Dim _IsConnect As Boolean = False
    Dim _IP_TCP As String = ""
    Dim _Port_TCP As String = ""
    Dim _IP_UDP As String = "@"
    Dim _Port_UDP As String = "@"
    Dim _Com As String = "COM2"
    Dim _Refresh As Integer = 0
    Dim _Modele As String = "@"
    Dim _Version As String = My.Application.Info.Version.ToString
    Dim _OsPlatform As String = "3264"
    Dim _Picture As String = ""
    Dim _Server As HoMIDom.HoMIDom.Server
    Dim _Device As HoMIDom.HoMIDom.Device
    Dim _DeviceSupport As New ArrayList
    Dim _Parametres As New ArrayList
    Dim _LabelsDriver As New ArrayList
    Dim _LabelsDevice As New ArrayList
    Dim MyTimer As New Timers.Timer
    Dim _IdSrv As String
    Dim _DeviceCommandPlus As New List(Of HoMIDom.HoMIDom.Device.DeviceCommande)
    Dim _AutoDiscover As Boolean = False

    'param avancé
    Dim _DEBUG As Boolean = False
    'Dim _PARAMMODE As String = "20011111111111111011111111"
    Dim _PARAMMODE_1_frequence As Integer = 2 '1 : type frequence (310, 315, 433, 868.30, 868.30 FSK, 868.35, 868.35 FSK, 868.95)
    Dim _PARAMMODE_2_undec As Integer = 0 '2 : UNDEC
    Dim _PARAMMODE_3_novatis As Integer = 0 '3 : novatis --> NOT USED ANYMORE 200
    Dim _PARAMMODE_4_proguard As Integer = 1 '4 : proguard
    Dim _PARAMMODE_5_fs20 As Integer = 1 '5 : FS20
    Dim _PARAMMODE_6_lacrosse As Integer = 1 '6 : Lacrosse
    Dim _PARAMMODE_7_hideki As Integer = 1 '7 : Hideki
    Dim _PARAMMODE_8_ad As Integer = 1 '8 : AD
    Dim _PARAMMODE_9_mertik As Integer = 1 '9 : Mertik 111111
    Dim _PARAMMODE_10_visonic As Integer = 1 '10 : Visonic
    Dim _PARAMMODE_11_ati As Integer = 1 '11 : ATI
    Dim _PARAMMODE_12_oregon As Integer = 1 '12 : Oregon
    Dim _PARAMMODE_13_meiantech As Integer = 1 '13 : Meiantech
    Dim _PARAMMODE_14_heeu As Integer = 1 '14 : HEEU
    Dim _PARAMMODE_15_ac As Integer = 1 '15 : AC
    Dim _PARAMMODE_16_arc As Integer = 1 '16 : ARC
    Dim _PARAMMODE_17_x10 As Integer = 1 '17 : X10 11111111
    Dim _PARAMMODE_18_blindst0 As Integer = 0 '18 : BlindsT0
    Dim _PARAMMODE_19_Imagintronix As Integer = 1 '19 : Imagintronix
    Dim _PARAMMODE_20_sx As Integer = 1 '20 : SX
    Dim _PARAMMODE_21_rsl As Integer = 1 '21 : RSL
    Dim _PARAMMODE_22_lighting4 As Integer = 1 '22 : LIGHTING4
    Dim _PARAMMODE_23_fineoffset As Integer = 1 '23 : FINEOFFSET
    Dim _PARAMMODE_24_rubicson As Integer = 1 '24 : RUBICSON
    Dim _PARAMMODE_25_ae As Integer = 1 '25 : AE
    Dim _PARAMMODE_26_blindst1 As Integer = 1 '26 : BlindsT1

#End Region

#Region "Variables Internes"


    Enum ICMD As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        cmnd = 4
        msg1 = 5
        msg2 = 6
        msg3 = 7
        msg4 = 8
        msg5 = 9
        msg6 = 10
        msg7 = 11
        msg8 = 12
        msg9 = 13
        size = 13

        'Interface Control
        pTypeInterfaceControl = &H0
        sTypeInterfaceCommand = &H0

        'Interface commands
        cmdRESET = &H0 ' reset the receiver/transceiver
        cmdSTATUS = &H2 ' request firmware versions and configuration of the interface
        cmdSETMODE = &H3 ' set the configuration of the interface
        cmdSAVE = &H6 ' save receiving modes of the receiver/transceiver in non-volatile memory
        cmdStartRec = &H7   'start RFXtrx receiver device

        cmd310 = &H50 ' select 310MHz in the 310/315 transceiver
        cmd315 = &H51 ' select 315MHz in the 310/315 transceiver
        cmd433r = &H52 ' select 433.92MHz in the 433.92 receiver
        cmd433 = &H53 ' select 433.92MHz in the 433.92 transceiver
        cmd800 = &H55 ' select 868.00MHz ASK in the 868 transceiver
        cmd800F = &H56 ' select 868.00MHz FSK in the 868 transceiver
        cmd830 = &H57 ' select 868.30MHz ASK in the 868 transceiver
        cmd830F = &H58 ' select 868.30MHz FSK in the 868 transceiver
        cmd835 = &H59 ' select 868.35MHz ASK in the 868 transceiver
        cmd835F = &H5A ' select 868.35MHz FSK in the 868 transceiver
        cmd895 = &H5B ' select 868.95MHz in the 868 transceiver
    End Enum

    Enum IRESPONSE As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        cmnd = 4
        msg1 = 5
        msg2 = 6
        msg3 = 7
        msg4 = 8
        msg5 = 9
        msg6 = 10
        msg7 = 11
        msg8 = 12
        msg9 = 13
        msg10 = 14
        msg11 = 15
        msg12 = 16
        msg13 = 17
        msg14 = 18
        msg15 = 19
        msg16 = 20
        size = 20

        pTypeInterfaceMessage = &H1
        sTypeInterfaceResponse = &H0
        sTypeUnknownRFYremote = &H1
        sTypeExtError = &H2
        sTypeRFYremoteList = &H3
        sTypeRecStarted = &H7
        sTypeInterfaceWrongCommand = &HFF

        recType310 = &H50
        recType315 = &H51
        recType43392 = &H52
        trxType43392 = &H53
        trxType43342 = &H54
        recType86800 = &H55
        recType86800FSK = &H56
        recType86830 = &H57
        recType86830FSK = &H58
        recType86835 = &H59
        recType86835FSK = &H5A
        recType86895 = &H5B

        msg3_undec = &H80
        msg3_IMAGINTRONIX = &H40
        msg3_SX = &H20
        msg3_RSL = &H10
        msg3_LIGHTING4 = &H8
        msg3_FINEOFFSET = &H4
        msg3_RUBICSON = &H2
        msg3_AE = &H1

        msg4_BlindsT1 = &H80
        msg4_BlindsT0 = &H40
        msg4_PROGUARD = &H20
        msg4_FS20 = &H10
        msg4_LCROS = &H8
        msg4_HID = &H4
        msg4_AD = &H2
        msg4_MERTIK = &H1

        msg5_VISONIC = &H80
        msg5_ATI = &H40
        msg5_OREGON = &H20
        msg5_MEI = &H10
        msg5_HEU = &H8
        msg5_AC = &H4
        msg5_ARC = &H2
        msg5_X10 = &H1

        msg6_RFU7 = &H80
        msg6_RFU6 = &H40
        msg6_RFU5 = &H20
        msg6_RFU4 = &H10
        msg6_RFU3 = &H8
        msg6_RFU2 = &H4
        msg6_RFU1 = &H2
        msg6_KEELOQ = &H1
    End Enum

    Enum RXRESPONSE As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        msg = 4
        size = 4

        pTypeRecXmitMessage = &H2
        sTypeReceiverLockError = &H0
        sTypeTransmitterResponse = &H1
    End Enum

    Enum UNDECODED As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        msg1 = 4
        'msg2 to msg32 depending on RF packet length
        size = 36   'maximum size

        pTypeUndecoded = &H3
        sTypeUac = &H0
        sTypeUarc = &H1
        sTypeUati = &H2
        sTypeUhideki = &H3
        sTypeUlacrosse = &H4
        sTypeUad = &H5
        sTypeUmertik = &H6
        sTypeUoregon1 = &H7
        sTypeUoregon2 = &H8
        sTypeUoregon3 = &H9
        sTypeUproguard = &HA
        sTypeUvisonic = &HB
        sTypeUnec = &HC
        sTypeUfs20 = &HD
        sTypeUrsl = &HE
        sTypeUblinds = &HF
        sTypeUrubicson = &H10
        sTypeUae = &H11
        sTypeUfineoffset = &H12
        sTypeUrgb = &H13
        sTypeUrfy = &H14
        sTypeUselectplus = &H15
    End Enum

    Enum LIGHTING1 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        housecode = 4
        unitcode = 5
        cmnd = 6
        filler = 7 'bits 3-0
        rssi = 7   'bits 7-4
        size = 7

        pTypeLighting1 = &H10
        sTypeX10 = &H0
        sTypeARC = &H1
        sTypeAB400D = &H2
        sTypeWaveman = &H3
        sTypeEMW200 = &H4
        sTypeIMPULS = &H5
        sTypeRisingSun = &H6
        sTypePhilips = &H7
        sTypeEnergenie = &H8
        sTypeEnergenie5 = &H9
        sTypeGDR2 = &HA

        sOff = 0
        sOn = 1
        sDim = 2
        sBright = 3
        sAllOff = 5
        sAllOn = 6
        sChime = 7
    End Enum

    Enum LIGHTING2 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        id3 = 6
        id4 = 7
        unitcode = 8
        cmnd = 9
        level = 10
        filler = 11 'bits 3-0
        rssi = 11   'bits 7-4
        size = 11

        pTypeLighting2 = &H11
        sTypeAC = &H0
        sTypeHEU = &H1
        sTypeANSLUT = &H2
        sTypeKambrook = &H3

        sOff = 0
        sOn = 1
        sSetLevel = 2
        sGroupOff = 3
        sGroupOn = 4
        sSetGroupLevel = 5
    End Enum

    Enum LIGHTING3 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        system = 4
        channel8_1 = 5
        channel10_9 = 6
        cmnd = 7
        filler = 8  'bits 3-0
        rssi = 8    'bits 7-4
        size = 8

        pTypeLighting3 = &H12
        sTypeKoppla = &H0

        sBright = &H0
        sDim = &H8
        sOn = &H10
        sLevel1 = &H11
        sLevel2 = &H12
        sLevel3 = &H13
        sLevel4 = &H14
        sLevel5 = &H15
        sLevel6 = &H16
        sLevel7 = &H17
        sLevel8 = &H18
        sLevel9 = &H19
        sOff = &H1A
        sProgram = &H1C
    End Enum

    Enum LIGHTING4 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        cmd1 = 4
        cmd2 = 5
        cmd3 = 6
        pulsehigh = 7
        pulselow = 8
        filler = 9  'bits 3-0
        rssi = 9    'bits 7-4
        size = 9

        pTypeLighting4 = &H13
        sTypePT2262 = &H0
    End Enum

    Enum LIGHTING5 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        id3 = 6
        unitcode = 7
        cmnd = 8
        level = 9
        filler = 10 'bits 3-0
        rssi = 10   'bits 7-4
        size = 10

        pTypeLighting5 = &H14
        sTypeLightwaveRF = &H0
        sTypeEMW100 = &H1
        sTypeBBSB = &H2
        sTypeMDREMOTE = &H3
        sTypeRSL = &H4
        sTypeLivolo = &H5
        sTypeTRC02 = &H6
        sTypeAoke = &H7
        sTypeTRC02_2 = &H8
        sTypeEurodomest = &H9
        sTypeLivoloAppliance = &HA

        sOff = 0
        sOn = 1
        sGroupOff = 2
        sLearn = 2
        sGroupOn = 3
        sMood1 = 3
        sMood2 = 4
        sMood3 = 5
        sMood4 = 6
        sMood5 = 7
        sUnlock = 10
        sLock = 11
        sAllLock = 12
        sClose = 13
        sStop = 14
        sOpen = 15
        sSetLevel = 16
        sColourPalette = 17
        sColourTone = 18
        sColourCycle = 19

        sPower = 0
        sLight = 1
        sBright = 2
        sDim = 3
        s100 = 4
        s50 = 5
        s25 = 6
        sModePlus = 7
        sSpeedMin = 8
        sSpeedPlus = 9
        sModeMin = 10

        sLivoloAllOff = 0
        sLivoloGang1Toggle = 1
        sLivoloGang2Toggle = 2   'dim+ for dimmer
        sLivoloGang3Toggle = 3   'dim- for dimmer
        sLivoloGang4Toggle = 4
        sLivoloGang5Toggle = 5
        sLivoloGang6Toggle = 6
        sLivoloGang7Toggle = 7
        sLivoloGang8Toggle = 8
        sLivoloGang9Toggle = 9
        sLivoloGang10Toggle = 10

        sRGBoff = 0
        sRGBon = 1
        sRGBbright = 2
        sRGBdim = 3
        sRGBcolorplus = 4
        sRGBcolormin = 5
    End Enum

    Enum LIGHTING6 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        groupcode = 6
        unitcode = 7
        cmnd = 8
        cmndseqnbr = 9
        seqnbr2 = 10
        filler = 11 'bits 3-0
        rssi = 11   'bits 7-4
        size = 11

        pTypeLighting6 = &H15
        sTypeBlyss = &H0

        sOn = 0
        sOff = 1
        sGroupOn = 2
        sGroupOff = 3
    End Enum

    Enum CHIME As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        sound = 6   'sound or id3
        filler = 7 'bits 3-0
        rssi = 7   'bits 7-4
        size = 7

        'types for Chime
        pTypeChime = &H16
        sTypeByronSX = &H0
        sTypeByronMP001 = &H1
        sTypeSelectPlus = &H2 'SelectPlus200689101
        sTypeRFU = &H3 'not used
        sTypeEnvivo = &H4 'Envivo ENV-1348

        sSound0 = 1
        sSound1 = 3
        sSound2 = 5
        sSound3 = 9
        sSound4 = 13
        sSound5 = 14
        sSound6 = 6
        sSound7 = 2
    End Enum

    Enum FAN As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        id3 = 6
        cmnd = 7
        filler = 8 'bits 3-0
        rssi = 8   'bits 7-4
        size = 8

        'types for Fan
        pTypeFan = &H17
        sTypeSiemensSF01 = &H0

        sTimer = 1
        sMin = 2
        sLearn = 3
        sPlus = 4
        sConfirm = 5
        sLight = 6
    End Enum

    Enum CURTAIN1 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        housecode = 4
        unitcode = 5
        cmnd = 6
        filler = 7 'bits 3-0
        rssi = 7   'bits 7-4
        size = 7

        'types for Curtain
        pTypeCurtain = &H18
        sTypeHarrison = &H0

        sOpen = 0
        sClose = 1
        sStop = 2
        sProgram = 3
    End Enum

    Enum BLINDS1 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        id3 = 6
        unitcode = 7    'bits 3-0
        id4 = 7         'bits 7-4  at BlindsT6 & T7
        cmnd = 8
        filler = 9 'bits 3-0
        rssi = 9   'bits 7-4
        size = 9

        'types for Blindss
        pTypeBlinds = &H19
        BlindsT0 = &H0    'RollerTrol or Hasta new
        BlindsT1 = &H1    'Hasta old
        BlindsT2 = &H2    'A-OK RF01
        BlindsT3 = &H3    'A-OK AC114
        BlindsT4 = &H4    'RAEX
        BlindsT5 = &H5    'Media Mount
        BlindsT6 = &H6    'DC106
        BlindsT7 = &H7    'Forest
        BlindsT8 = &H8    'Chamberlain CS4330CN
        BlindsT9 = &H9    'Sunpery
        BlindsT10 = &HA   'Dolat

        sOpen = 0
        sClose = 1
        sStop = 2
        sConfirm = 3
        sLimit = 4
        sLowerLimit = 5
        sDeleteLimits = 6
        sChangeDirection = 7
        sLeft = 8
        sRight = 9
        s9ChangeDirection = 6
        s9ImA = 7
        s9ImCenter = 8
        s9ImB = 9
        s9EraseCurrentCh = 10
        s9EraseAllCh = 11
        s10LearnMaster = 4
        s10EraseCurrentCh = 5
        s10ChangeDirection = 6
    End Enum

    Enum RFY As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        id3 = 6
        unitcode = 7
        cmnd = 8
        rfu1 = 9
        rfu2 = 10
        rfu3 = 11
        filler = 12 'bits 3-0
        rssi = 12   'bits 7-4
        size = 12

        'types for Blinds
        pTypeRFY = &H1A
        RFY = &H0
        RFYext = &H1    'not yet used
        GEOM = &H2

        sStop = 0
        sUp = 1
        sUpStop = 2
        sDown = 3
        sDownStop = 4
        sUpDown = 5
        sListRemotes = 6
        sProgram = 7
        s2SecProgram = 8
        s7SecProgram = 9
        s2SecStop = 10
        s5SecStop = 11
        s5SecUpDown = 12
        sEraseThis = 13
        sEraseAll = 14
        s05SecUP = 15
        s05SecDown = 16
        s2SecUP = 17
        s2SecDown = 18
        sEnableSunWind = 19
        sDisableSun = 20
    End Enum

    Enum SECURITY1 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        id3 = 6
        status = 7
        rssi = 8            'bits 7-4
        battery_level = 8   'bits 3-0
        filler = 8
        size = 8

        'Security1
        pTypeSecurity1 = &H20
        sTypeSecX10 = &H0
        sTypeSecX10M = &H1
        sTypeSecX10R = &H2
        sTypeKD101 = &H3
        sTypePowercodeSensor = &H4
        sTypePowercodeMotion = &H5
        sTypeCodesecure = &H6
        sTypePowercodeAux = &H7
        sTypeMeiantech = &H8
        sTypeSA30 = &H9

        'status security
        sStatusNormal = &H0
        sStatusNormalDelayed = &H1
        sStatusAlarm = &H2
        sStatusAlarmDelayed = &H3
        sStatusMotion = &H4
        sStatusNoMotion = &H5
        sStatusPanic = &H6
        sStatusPanicOff = &H7
        sStatusIRbeam = &H8
        sStatusArmAway = &H9
        sStatusArmAwayDelayed = &HA
        sStatusArmHome = &HB
        sStatusArmHomeDelayed = &HC
        sStatusDisarm = &HD
        sStatusLightOff = &H10
        sStatusLightOn = &H11
        sStatusLIGHTING2Off = &H12
        sStatusLIGHTING2On = &H13
        sStatusDark = &H14
        sStatusLight = &H15
        sStatusBatLow = &H16
        sStatusPairKD101 = &H17
        sStatusNormalTamper = &H80
        sStatusNormalDelayedTamper = &H81
        sStatusAlarmTamper = &H82
        sStatusAlarmDelayedTamper = &H83
        sStatusMotionTamper = &H84
        sStatusNoMotionTamper = &H85
    End Enum

    Enum SECURITY2 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        id3 = 6
        id4 = 7
        id5 = 8
        id6 = 9
        id7 = 10
        id8 = 11
        id9 = 12
        id10 = 13
        id11 = 14
        id12 = 15
        id13 = 16
        id14 = 17
        id15 = 18
        id16 = 19
        id17 = 20
        id18 = 21
        id19 = 22
        id20 = 23
        id21 = 24
        id22 = 25
        id23 = 26
        id24 = 27
        rssi = 28           'bits 7-4
        battery_level = 28  'bits 3-0
        size = 28

        'Security2 KeeLoq
        pTypeSecurity2 = &H21
        sTypeSec2Classic = &H0  'KeeLoq packet
    End Enum

    Enum CAMERA1 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        housecode = 4
        cmnd = 5
        filler = 6 'bits 3-0
        rssi = 6   'bits 7-4
        size = 6

        'Camera1
        pTypeCamera = &H28
        sTypeNinja = &H0

        sLeft = 0
        sRight = 1
        sUp = 2
        sDown = 3
        sPosition1 = 4
        sProgramPosition1 = 5
        sPosition2 = 6
        sProgramPosition2 = 7
        sPosition3 = 8
        sProgramPosition3 = 9
        sPosition4 = 10
        sProgramPosition4 = 11
        sCenter = 12
        sProgramCenterPosition = 13
        sSweep = 14
        sProgramSweep = 15
    End Enum

    Enum REMOTE As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id = 4
        cmnd = 5
        toggle = 6       'bit 0
        cmndtype = 6       'bits 3-1
        rssi = 6         'bits 7-4
        size = 6

        'Remotes
        pTypeRemote = &H30
        sTypeATI = &H0
        sTypeATIplus = &H1
        sTypeMedion = &H2
        sTypePCremote = &H3
        sTypeATIrw2 = &H4
    End Enum

    Enum THERMOSTAT1 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        temperature = 6
        set_point = 7
        status = 8  'bits 1-0
        filler = 8  'bits 6-2
        mode = 8    'bit 7
        battery_level = 9   'bits 3-0
        rssi = 9            'bits 7-4
        size = 9

        'Thermostat1
        pTypeThermostat1 = &H40
        sTypeDigimax = &H0    'Digimax with long packet 
        sTypeDigimaxShort = &H1   'Digimax with short packet (no set point)

        sDemand = 1
        sNoDemand = 2
        sInitializing = 3
    End Enum

    Enum THERMOSTAT2 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        unitcode = 4
        cmnd = 5
        filler = 6  'bits 3-0
        rssi = 6    'bits 7-4
        size = 6

        'Thermostat2
        pTypeThermostat2 = &H41
        sTypeHE105 = &H0  'HE105
        sTypeRTS10 = &H1  'RTS10

        sOff = 0
        sOn = 1
        sProgram = 2
    End Enum

    Enum THERMOSTAT3 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        unitcode1 = 4
        unitcode2 = 5
        unitcode3 = 6
        cmnd = 7
        filler = 8   'bits 3-0
        rssi = 8     'bits 7-4
        size = 8

        'Thermostat3
        pTypeThermostat3 = &H42
        sTypeMertikG6RH4T1 = &H0  'Mertik G6R-H4T1
        sTypeMertikG6RH4TB = &H1  'Mertik G6R-H4TB
        sTypeMertikG6RH4TD = &H2  'Mertik G6R-H4TD

        sOff = 0
        sOn = 1
        sUp = 2
        sDown = 3
        sRunUp = 4
        Off2nd = 4
        sRunDown = 5
        On2nd = 5
        sStop = 6
    End Enum

    Enum RADIATOR1 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        id3 = 6
        id4 = 7
        unitcode = 8
        cmnd = 9
        temperature = 10
        tempPoint5 = 11
        filler = 12 'bits 3-0
        rssi = 12   'bits 7-4
        size = 12

        pTypeRadiator1 = &H48
        sTypeSmartwares = &H0

        sNight = 0
        sDay = 1
        sSetTemp = 2
    End Enum

    Enum BBQ As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        sensor1h = 6
        sensor1l = 7
        sensor2h = 8
        sensor2l = 9
        battery_level = 10   'bits 3-0
        rssi = 10            'bits 7-4
        size = 10

        'Temperature
        pTypeBBQ = &H4E
        sTypeBBQ1 = &H1  'Maverick ET-732
    End Enum

    Enum TEMP_RAIN As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        temperatureh = 6    'bits 6-0
        tempsign = 6        'bit 7
        temperaturel = 7
        raintotal1 = 8
        raintotal2 = 9
        battery_level = 10  'bits 3-0
        rssi = 10           'bits 7-4
        size = 10

        'temperature+humidity
        pTypeTEMP_RAIN = &H4F
        sTypeTR1 = &H1    'WS1200
    End Enum

    Enum TEMP As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        temperatureh = 6    'bits 6-0
        tempsign = 6        'bit 7
        temperaturel = 7
        battery_level = 8   'bits 3-0
        rssi = 8            'bits 7-4
        size = 8

        'Temperature
        pTypeTEMP = &H50
        sTypeTEMP1 = &H1  'THR128/138, THC138
        sTypeTEMP2 = &H2  'THC238/268,THN132,THWR288,THRN122,THN122,AW129/131
        sTypeTEMP3 = &H3  'THWR800
        sTypeTEMP4 = &H4  'RTHN318
        sTypeTEMP5 = &H5  'LaCrosse TX3
        sTypeTEMP6 = &H6  'TS15C
        sTypeTEMP7 = &H7  'Viking 02811/Proove TSS330,311346
        sTypeTEMP8 = &H8  'WS2300
        sTypeTEMP9 = &H9  'RUBiCSON
        sTypeTEMP10 = &HA   'TFA 30.3133
        sTypeTEMP11 = &HB   'WT0122
    End Enum

    Enum HUM As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        humidity = 6
        humidity_status = 7
        battery_level = 8  'bits 3-0
        rssi = 8           'bits 7-4
        size = 8

        'humidity
        pTypeHUM = &H51
        sTypeHUM1 = &H1  'LaCrosse TX3
        sTypeHUM2 = &H2  'LaCrosse WS2300

        snormal = 0
        scomfort = 1
        sdry = 2
        swet = 3
    End Enum

    Enum TEMP_HUM As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        temperatureh = 6    'bits 6-0
        tempsign = 6        'bit 7
        temperaturel = 7
        humidity = 8
        humidity_status = 9
        battery_level = 10  'bits 3-0
        rssi = 10           'bits 7-4
        size = 10

        'temperature+humidity
        pTypeTEMP_HUM = &H52
        sTypeTH1 = &H1    'THGN122/123,/THGN132,THGR122/228/238/268
        sTypeTH2 = &H2    'THGR810/THGN800
        sTypeTH3 = &H3    'RTGR328
        sTypeTH4 = &H4    'THGR328
        sTypeTH5 = &H5    'WTGR800
        sTypeTH6 = &H6    'THGR918,THGRN228,THGN500
        sTypeTH7 = &H7    'TFA TS34C, Cresta
        sTypeTH8 = &H8    'Esic
        sTypeTH9 = &H9    'viking 02038/Proove TSS320,311501
        sTypeTH10 = &HA    'Rubicson
        sTypeTH11 = &HB    'EW109
        sTypeTH12 = &HC    'Imagintronix soil sensor
        sTypeTH13 = &HD    'Alecto WS1700
        sTypeTH14 = &HE    'Alecto
    End Enum

    Enum BARO As Byte
        'barometric
        pTypeBARO = &H53  'not used
    End Enum

    Enum TEMP_HUM_BARO As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        temperatureh = 6    'bits 6-0
        tempsign = 6        'bit 7
        temperaturel = 7
        humidity = 8
        humidity_status = 9
        baroh = 10
        barol = 11
        forecast = 12
        battery_level = 13  'bits 3-0
        rssi = 13           'bits 7-4
        size = 13

        'temperature+humidity+baro
        pTypeTEMP_HUM_BARO = &H54
        sTypeTHB1 = &H1   'BTHR918
        sTypeTHB2 = &H2   'BTHR918N,BTHR968
    End Enum

    Enum RAIN As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        rainrateh = 6
        rainratel = 7
        raintotal1 = 8
        raintotal2 = 9
        raintotal3 = 10
        battery_level = 11  'bits 3-0
        rssi = 11           'bits 7-4
        size = 11

        'rain
        pTypeRAIN = &H55
        sTypeRAIN1 = &H1   'RGR126/682/918
        sTypeRAIN2 = &H2   'PCR800
        sTypeRAIN3 = &H3   'TFA
        sTypeRAIN4 = &H4   'UPM
        sTypeRAIN5 = &H5   'WS2300
        sTypeRAIN6 = &H6   'TX5
        sTypeRAIN7 = &H7   'Alecto
    End Enum

    Enum WIND As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        directionh = 6
        directionl = 7
        av_speedh = 8
        av_speedl = 9
        gusth = 10
        gustl = 11
        temperatureh = 12    'bits 6-0
        tempsign = 12        'bit 7
        temperaturel = 13
        chillh = 14    'bits 6-0
        chillsign = 14        'bit 7
        chilll = 15
        battery_level = 16  'bits 3-0
        rssi = 16           'bits 7-4
        size = 16

        'wind
        pTypeWIND = &H56
        sTypeWIND1 = &H1   'WTGR800
        sTypeWIND2 = &H2   'WGR800
        sTypeWIND3 = &H3   'STR918,WGR918
        sTypeWIND4 = &H4   'TFA
        sTypeWIND5 = &H5   'UPM
        sTypeWIND6 = &H6   'WS2300
        sTypeWIND7 = &H7   'Alecto WS4500
    End Enum

    Enum UV As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        uv = 6
        temperatureh = 7    'bits 6-0
        tempsign = 7        'bit 7
        temperaturel = 8
        battery_level = 9   'bits 3-0
        rssi = 9            'bits 7-4
        size = 9

        'uv
        pTypeUV = &H57
        sTypeUV1 = &H1   'UVN128,UV138
        sTypeUV2 = &H2   'UVN800
        sTypeUV3 = &H3   'TFA
    End Enum

    Enum DT As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        yy = 6
        mm = 7
        dd = 8
        dow = 9
        hr = 10
        min = 11
        sec = 12
        battery_level = 13  'bits 3-0
        rssi = 13           'bits 7-4
        size = 13

        'date & time
        pTypeDT = &H58
        sTypeDT1 = &H1   'RTGR328N
    End Enum

    Enum CURRENT As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        count = 6
        ch1h = 7
        ch1l = 8
        ch2h = 9
        ch2l = 10
        ch3h = 11
        ch3l = 12
        battery_level = 13  'bits 3-0
        rssi = 13           'bits 7-4
        size = 13

        'current
        pTypeCURRENT = &H59
        sTypeELEC1 = &H1   'CM113,Electrisave
    End Enum

    Enum ENERGY As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        count = 6
        instant1 = 7
        instant2 = 8
        instant3 = 9
        instant4 = 10
        total1 = 11
        total2 = 12
        total3 = 13
        total4 = 14
        total5 = 15
        total6 = 16
        battery_level = 17  'bits 3-0
        rssi = 17           'bits 7-4
        size = 17

        'energy
        pTypeENERGY = &H5A
        sTypeELEC2 = &H1   'CM119/160
        sTypeELEC3 = &H2   'CM180
    End Enum

    Enum CURRENT_ENERGY As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        count = 6
        ch1h = 7
        ch1l = 8
        ch2h = 9
        ch2l = 10
        ch3h = 11
        ch3l = 12
        total1 = 13
        total2 = 14
        total3 = 15
        total4 = 16
        total5 = 17
        total6 = 18
        battery_level = 19  'bits 3-0
        rssi = 19           'bits 7-4
        size = 19

        'current-energy
        pTypeCURRENTENERGY = &H5B
        sTypeELEC4 = &H1   'CM180i
    End Enum

    Enum POWER As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        voltage = 6
        currentH = 7
        currentL = 8
        powerH = 9
        powerL = 10
        energyH = 11
        energyL = 12
        pf = 13
        freq = 14
        filler = 15  'bits 3-0
        rssi = 15    'bits 7-4
        size = 15

        'current-energy
        pTypePOWER = &H5C
        sTypeELEC5 = &H1   'Revolt
    End Enum

    Enum WEIGHT As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        weighthigh = 6
        weightlow = 7
        filler = 8   'bits 3-0
        rssi = 8     'bits 7-4
        size = 8

        'weight scales
        pTypeWEIGHT = &H5D
        sTypeWEIGHT1 = &H1   'BWR102
        sTypeWEIGHT2 = &H2   'GR101
    End Enum

    Enum GAS As Byte
        'gas
        pTypeGAS = &H5E  'not used
    End Enum

    Enum WATER As Byte
        'water
        pTypeWATER = &H5F  'not used
    End Enum

    Enum RFXSENSOR As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id = 4
        msg1 = 5
        msg2 = 6
        filler = 7  'bits 3-0
        rssi = 7    'bits 7-4
        size = 7

        'RFXSensor
        pTypeRFXSensor = &H70
        sTypeTemp = &H0
        sTypeAD = &H1
        sTypeVolt = &H2
        sTypeMessage = &H3
    End Enum

    Enum RFXMETER As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        id1 = 4
        id2 = 5
        count1 = 6
        count2 = 7
        count3 = 8
        count4 = 9
        filler = 10 'bits 3-0
        rssi = 10   'bits 7-4
        size = 10

        'RFXMeter
        pTypeRFXMeter = &H71
        sTypeCount = &H0
        sTypeInterval = &H1
        sTypeCalib = &H2
        sTypeAddr = &H3
        sTypeCounterReset = &H4
        sTypeCounterSet = &HB
        sTypeSetInterval = &HC
        sTypeSetCalib = &HD
        sTypeSetAddr = &HE
        sTypeIdent = &HF
    End Enum

    Enum FS20 As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        hc1 = 4
        hc2 = 5
        addr = 6
        cmd1 = 7
        cmd2 = 8
        filler = 9 'bits 3-0
        rssi = 9   'bits 7-4
        size = 9

        'FS20
        pTypeFS20 = &H72
        sTypeFS20 = &H0
        sTypeFHT8V = &H1
        sTypeFHT80 = &H2
    End Enum

    Enum RAW As Byte
        packetlength = 0
        packettype = 1
        subtype = 2
        seqnbr = 3
        repeat = 4
        uint1_msb = 5
        uint1_lsb = 6

        size = 254

        'RAW transmit
        pTypeRAW = &H7F
        sTypeRAW = &H0
    End Enum

    'liste des variables de base
    Dim WithEvents RS232Port As New SerialPort
    Private gRecComPortEnabled As Boolean = False
    'Private Resettimer As Integer = 0
    Private trxType As Integer = 0
    Private Shared rfxtrxlock As New Object

    Private recbuf(60), recbytes As Byte
    Private bytecnt As Integer = 0
    Private message As String
    Private bytSeqNbr As Byte = 0
    Private bytRemoteToggle As Byte = 0
    Private bytFWversion As Byte
    Private bytCmndSeqNbr As Byte = 0
    Private bytCmndSeqNbr2 As Byte = 0

    'La Crosse TX5 rain sensor data
    Private shortTotalRain As Short = 0
    Private byteFlipCount As Byte = 0

    Private client As TcpClient
    Private stream As NetworkStream
    Private tcp As Boolean
    Private maxticks As Byte = 0
    Private LogFile As StreamWriter
    Private LogActive As Boolean = False
    Private TCPData(1024) As Byte

    Private port_name As String = ""
    Private dateheurelancement As DateTime
    Dim adressetoint() As String = {"00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "0A", "0B", "0C", "0D", "0E", "0F", "10", "11", "12", "13", "14", "15", "16", "17", "18", "19", "1A", "1B", "1C", "1D", "1E", "1F", "20", "21", "22", "23", "24", "25", "26", "27", "28", "29", "2A", "2B", "2C", "2D", "2E", "2F", "30", "31", "32", "33", "34", "35", "36", "37", "38", "39", "3A", "3B", "3C", "3D", "3E", "3F", "40", "41", "42", "43", "44", "45", "46", "47", "48", "49", "4A", "4B", "4C", "4D", "4E", "4F", "50", "51", "52", "53", "54", "55", "56", "57", "58", "59", "5A", "5B", "5C", "5D", "5E", "5F", "60", "61", "62", "63", "64", "65", "66", "67", "68", "69", "6A", "6B", "6C", "6D", "6E", "6F", "70", "71", "72", "73", "74", "75", "76", "77", "78", "79", "7A", "7B", "7C", "7D", "7E", "7F", "80", "81", "82", "83", "84", "85", "86", "87", "88", "89", "8A", "8B", "8C", "8D", "8E", "8F", "90", "91", "92", "93", "94", "95", "96", "97", "98", "99", "9A", "9B", "9C", "9D", "9E", "9F", "A0", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "AA", "AB", "AC", "AD", "AE", "AF", "B0", "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "BA", "BB", "BC", "BD", "BE", "BF", "C0", "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "CA", "CB", "CC", "CD", "CE", "CF", "D0", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "DA", "DB", "DC", "DD", "DE", "DF", "E0", "E1", "E2", "E3", "E4", "E5", "E6", "E7", "E8", "E9", "EA", "EB", "EC", "ED", "EE", "EF", "F0", "F1", "F2", "F3", "F4", "F5", "F6", "F7", "F8", "F9", "FA", "FB", "FC", "FD", "FE", "FF"}
    Dim adressetoint2() As String = {"0", "1", "2", "3"}
    Dim unittoint() As String = {"1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", "13", "14", "15", "16"}
    Dim messagerecu As String
    'Private WithEvents tmrRead As New System.Timers.Timer

    'old
    'Private WithEvents tmrRead As New System.Timers.Timer
    'Private messagetemp, messagelast, adresselast, valeurlast, recbuf_last As String
    'Private nblast As Integer = 0
    'Private BufferIn(8192) As Byte
    'Const GETSW As Byte = &H30
    'Const MODEBLK As Byte = &H31
    'Const PING As Byte = &H32
    'Const MODERBRB48 As Byte = &H33
    'Const MODECONT As Byte = &H35
    'Const MODEBRB48 As Byte = &H37
    'Private protocolsynchro As Integer = MODEBRB48
    'Private ack As Boolean = False
    'Private ack_ok As Boolean = True

#End Region

#Region "Propriétés génériques"
    Public WriteOnly Property IdSrv As String Implements HoMIDom.HoMIDom.IDriver.IdSrv
        Set(ByVal value As String)
            _IdSrv = value
        End Set
    End Property

    Public Property Server() As HoMIDom.HoMIDom.Server Implements HoMIDom.HoMIDom.IDriver.Server
        Get
            Return _Server
        End Get
        Set(ByVal value As HoMIDom.HoMIDom.Server)
            _Server = value
        End Set
    End Property
    Public ReadOnly Property DeviceSupport() As ArrayList Implements HoMIDom.HoMIDom.IDriver.DeviceSupport
        Get
            Return _DeviceSupport
        End Get
    End Property
    Public Property Parametres() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.Parametres
        Get
            Return _Parametres
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _Parametres = value
        End Set
    End Property

    Public Property LabelsDriver() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.LabelsDriver
        Get
            Return _LabelsDriver
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _LabelsDriver = value
        End Set
    End Property
    Public Property LabelsDevice() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.LabelsDevice
        Get
            Return _LabelsDevice
        End Get
        Set(ByVal value As System.Collections.ArrayList)
            _LabelsDevice = value
        End Set
    End Property

    Public Property COM() As String Implements HoMIDom.HoMIDom.IDriver.COM
        Get
            Return _Com
        End Get
        Set(ByVal value As String)
            _Com = value
        End Set
    End Property
    Public ReadOnly Property Description() As String Implements HoMIDom.HoMIDom.IDriver.Description
        Get
            Return _Description
        End Get
    End Property
    Public Event DriverEvent(ByVal DriveName As String, ByVal TypeEvent As String, ByVal Parametre As Object) Implements HoMIDom.HoMIDom.IDriver.DriverEvent
    Public Property Enable() As Boolean Implements HoMIDom.HoMIDom.IDriver.Enable
        Get
            Return _Enable
        End Get
        Set(ByVal value As Boolean)
            _Enable = value
        End Set
    End Property
    Public ReadOnly Property ID() As String Implements HoMIDom.HoMIDom.IDriver.ID
        Get
            Return _ID
        End Get
    End Property
    Public Property IP_TCP() As String Implements HoMIDom.HoMIDom.IDriver.IP_TCP
        Get
            Return _IP_TCP
        End Get
        Set(ByVal value As String)
            _IP_TCP = value
        End Set
    End Property
    Public Property IP_UDP() As String Implements HoMIDom.HoMIDom.IDriver.IP_UDP
        Get
            Return _IP_UDP
        End Get
        Set(ByVal value As String)
            _IP_UDP = value
        End Set
    End Property
    Public ReadOnly Property IsConnect() As Boolean Implements HoMIDom.HoMIDom.IDriver.IsConnect
        Get
            Return _IsConnect
        End Get
    End Property
    Public Property Modele() As String Implements HoMIDom.HoMIDom.IDriver.Modele
        Get
            Return _Modele
        End Get
        Set(ByVal value As String)
            _Modele = value
        End Set
    End Property
    Public ReadOnly Property Nom() As String Implements HoMIDom.HoMIDom.IDriver.Nom
        Get
            Return _Nom
        End Get
    End Property
    Public Property Picture() As String Implements HoMIDom.HoMIDom.IDriver.Picture
        Get
            Return _Picture
        End Get
        Set(ByVal value As String)
            _Picture = value
        End Set
    End Property
    Public Property Port_TCP() As String Implements HoMIDom.HoMIDom.IDriver.Port_TCP
        Get
            Return _Port_TCP
        End Get
        Set(ByVal value As String)
            _Port_TCP = value
        End Set
    End Property
    Public Property Port_UDP() As String Implements HoMIDom.HoMIDom.IDriver.Port_UDP
        Get
            Return _Port_UDP
        End Get
        Set(ByVal value As String)
            _Port_UDP = value
        End Set
    End Property
    Public ReadOnly Property Protocol() As String Implements HoMIDom.HoMIDom.IDriver.Protocol
        Get
            Return _Protocol
        End Get
    End Property
    Public Property Refresh() As Integer Implements HoMIDom.HoMIDom.IDriver.Refresh
        Get
            Return _Refresh
        End Get
        Set(ByVal value As Integer)
            _Refresh = value
        End Set
    End Property
    Public ReadOnly Property Version() As String Implements HoMIDom.HoMIDom.IDriver.Version
        Get
            Return _Version
        End Get
    End Property
    Public ReadOnly Property OsPlatform() As String Implements HoMIDom.HoMIDom.IDriver.OsPlatform
        Get
            Return _OsPlatform
        End Get
    End Property
    Public Property StartAuto() As Boolean Implements HoMIDom.HoMIDom.IDriver.StartAuto
        Get
            Return _StartAuto
        End Get
        Set(ByVal value As Boolean)
            _StartAuto = value
        End Set
    End Property
    Public Property AutoDiscover() As Boolean Implements HoMIDom.HoMIDom.IDriver.AutoDiscover
        Get
            Return _AutoDiscover
        End Get
        Set(ByVal value As Boolean)
            _AutoDiscover = value
        End Set
    End Property
#End Region

#Region "Fonctions génériques"
    ''' <summary>
    ''' Retourne la liste des Commandes avancées
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCommandPlus() As List(Of DeviceCommande)
        Return _DeviceCommandPlus
    End Function

    ''' <summary>Execute une commande avancée</summary>
    ''' <param name="Command"></param>
    ''' <param name="Param"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function ExecuteCommand(ByVal MyDevice As Object, ByVal Command As String, Optional ByVal Param() As Object = Nothing) As Boolean
        Dim retour As Boolean = False
        Try
            If MyDevice IsNot Nothing Then
                'Pas de commande demandée donc erreur
                If Command = "" Then
                    Return False
                Else
                    Write(MyDevice, Command, Param(0), Param(1))
                    Return True
                End If
            Else
                Return False
            End If
        Catch ex As Exception
            WriteLog("ERR: ExecuteCommand exception : " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>Permet de vérifier si un champ est valide</summary>
    ''' <param name="Champ"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function VerifChamp(ByVal Champ As String, ByVal Value As Object) As String Implements HoMIDom.HoMIDom.IDriver.VerifChamp
        Try
            Dim retour As String = "0"
            Select Case UCase(Champ)


            End Select
            Return retour
        Catch ex As Exception
            Return "Une erreur est apparue lors de la vérification du champ " & Champ & ": " & ex.ToString
        End Try
    End Function

    ''' <summary>Démarrer le du driver</summary>
    ''' <remarks></remarks>
    Public Sub Start() Implements HoMIDom.HoMIDom.IDriver.Start
        '_IsConnect = True
        Dim retour As String

        'récupération des paramétres avancés
        Try
            _DEBUG = _Parametres.Item(0).Valeur
            '_PARAMMODE = _Parametres.Item(1).Valeur
            '_AUTODISCOVER = _Parametres.Item(2).Valeur
            'correction si valeur correspond à ancienne valeur parammode de type "20011111111111111011111111"
            If CStr(_Parametres.Item(1).Valeur).Length = 26 Then
                WriteLog("ERR: Anciens Paramétres avancés trouvés. Conversion de l'ancienne valeur au nouveau format. Veuillez vérifier que les nouveaux paramétres sont corrects.")
                _PARAMMODE_1_frequence = _Parametres.Item(1).Valeur.Substring(0, 1)
                _PARAMMODE_2_undec = _Parametres.Item(1).Valeur.Substring(1, 1)
                _PARAMMODE_3_novatis = _Parametres.Item(1).Valeur.Substring(2, 1)
                _PARAMMODE_4_proguard = _Parametres.Item(1).Valeur.Substring(3, 1)
                _PARAMMODE_5_fs20 = _Parametres.Item(1).Valeur.Substring(4, 1)
                _PARAMMODE_6_lacrosse = _Parametres.Item(1).Valeur.Substring(5, 1)
                _PARAMMODE_7_hideki = _Parametres.Item(1).Valeur.Substring(6, 1)
                _PARAMMODE_8_ad = _Parametres.Item(1).Valeur.Substring(7, 1)
                _PARAMMODE_9_mertik = _Parametres.Item(1).Valeur.Substring(8, 1)
                _PARAMMODE_10_visonic = _Parametres.Item(1).Valeur.Substring(9, 1)
                _PARAMMODE_11_ati = _Parametres.Item(1).Valeur.Substring(10, 1)
                _PARAMMODE_12_oregon = _Parametres.Item(1).Valeur.Substring(11, 1)
                _PARAMMODE_13_meiantech = _Parametres.Item(1).Valeur.Substring(12, 1)
                _PARAMMODE_14_heeu = _Parametres.Item(1).Valeur.Substring(13, 1)
                _PARAMMODE_15_ac = _Parametres.Item(1).Valeur.Substring(14, 1)
                _PARAMMODE_16_arc = _Parametres.Item(1).Valeur.Substring(15, 1)
                _PARAMMODE_17_x10 = _Parametres.Item(1).Valeur.Substring(16, 1)
                _PARAMMODE_18_blindst0 = _Parametres.Item(1).Valeur.Substring(17, 1)
                _PARAMMODE_19_Imagintronix = _Parametres.Item(1).Valeur.Substring(18, 1)
                _PARAMMODE_20_sx = _Parametres.Item(1).Valeur.Substring(19, 1)
                _PARAMMODE_21_rsl = _Parametres.Item(1).Valeur.Substring(20, 1)
                _PARAMMODE_22_lighting4 = _Parametres.Item(1).Valeur.Substring(21, 1)
                _PARAMMODE_23_fineoffset = _Parametres.Item(1).Valeur.Substring(22, 1)
                _PARAMMODE_24_rubicson = _Parametres.Item(1).Valeur.Substring(23, 1)
                _PARAMMODE_25_ae = _Parametres.Item(1).Valeur.Substring(24, 1)
                _PARAMMODE_26_blindst1 = _Parametres.Item(1).Valeur.Substring(25, 1)

                _Parametres.Item(1).Valeur = _PARAMMODE_1_frequence
                _Parametres.Item(2).Valeur = _PARAMMODE_2_undec
                _Parametres.Item(3).Valeur = _PARAMMODE_3_novatis
                _Parametres.Item(4).Valeur = _PARAMMODE_4_proguard
                _Parametres.Item(5).Valeur = _PARAMMODE_5_fs20
                _Parametres.Item(6).Valeur = _PARAMMODE_6_lacrosse
                _Parametres.Item(7).Valeur = _PARAMMODE_7_hideki
                _Parametres.Item(8).Valeur = _PARAMMODE_8_ad
                _Parametres.Item(9).Valeur = _PARAMMODE_9_mertik
                _Parametres.Item(10).Valeur = _PARAMMODE_10_visonic
                _Parametres.Item(11).Valeur = _PARAMMODE_11_ati
                _Parametres.Item(12).Valeur = _PARAMMODE_12_oregon
                _Parametres.Item(13).Valeur = _PARAMMODE_13_meiantech
                _Parametres.Item(14).Valeur = _PARAMMODE_14_heeu
                _Parametres.Item(15).Valeur = _PARAMMODE_15_ac
                _Parametres.Item(16).Valeur = _PARAMMODE_16_arc
                _Parametres.Item(17).Valeur = _PARAMMODE_17_x10
                _Parametres.Item(18).Valeur = _PARAMMODE_18_blindst0
                _Parametres.Item(19).Valeur = _PARAMMODE_19_Imagintronix
                _Parametres.Item(20).Valeur = _PARAMMODE_20_sx
                _Parametres.Item(21).Valeur = _PARAMMODE_21_rsl
                _Parametres.Item(22).Valeur = _PARAMMODE_22_lighting4
                _Parametres.Item(23).Valeur = _PARAMMODE_23_fineoffset
                _Parametres.Item(24).Valeur = _PARAMMODE_24_rubicson
                _Parametres.Item(25).Valeur = _PARAMMODE_25_ae
                _Parametres.Item(26).Valeur = _PARAMMODE_26_blindst1

            ElseIf CStr(_Parametres.Item(1).Valeur).Length > 1 Then
                WriteLog("ERR: Erreur dans les paramétres avancés. utilisation des valeur par défaut")
                _Parametres.Item(1).Valeur = _PARAMMODE_1_frequence
                _Parametres.Item(2).Valeur = _PARAMMODE_2_undec
                _Parametres.Item(3).Valeur = _PARAMMODE_3_novatis
                _Parametres.Item(4).Valeur = _PARAMMODE_4_proguard
                _Parametres.Item(5).Valeur = _PARAMMODE_5_fs20
                _Parametres.Item(6).Valeur = _PARAMMODE_6_lacrosse
                _Parametres.Item(7).Valeur = _PARAMMODE_7_hideki
                _Parametres.Item(8).Valeur = _PARAMMODE_8_ad
                _Parametres.Item(9).Valeur = _PARAMMODE_9_mertik
                _Parametres.Item(10).Valeur = _PARAMMODE_10_visonic
                _Parametres.Item(11).Valeur = _PARAMMODE_11_ati
                _Parametres.Item(12).Valeur = _PARAMMODE_12_oregon
                _Parametres.Item(13).Valeur = _PARAMMODE_13_meiantech
                _Parametres.Item(14).Valeur = _PARAMMODE_14_heeu
                _Parametres.Item(15).Valeur = _PARAMMODE_15_ac
                _Parametres.Item(16).Valeur = _PARAMMODE_16_arc
                _Parametres.Item(17).Valeur = _PARAMMODE_17_x10
                _Parametres.Item(18).Valeur = _PARAMMODE_18_blindst0
                _Parametres.Item(19).Valeur = _PARAMMODE_19_Imagintronix
                _Parametres.Item(20).Valeur = _PARAMMODE_20_sx
                _Parametres.Item(21).Valeur = _PARAMMODE_21_rsl
                _Parametres.Item(22).Valeur = _PARAMMODE_22_lighting4
                _Parametres.Item(23).Valeur = _PARAMMODE_23_fineoffset
                _Parametres.Item(24).Valeur = _PARAMMODE_24_rubicson
                _Parametres.Item(25).Valeur = _PARAMMODE_25_ae
                _Parametres.Item(26).Valeur = _PARAMMODE_26_blindst1

            Else
                'situation normale, on recupere chaque parametre
                _PARAMMODE_1_frequence = _Parametres.Item(1).Valeur
                _PARAMMODE_2_undec = _Parametres.Item(2).Valeur
                _PARAMMODE_3_novatis = _Parametres.Item(3).Valeur
                _PARAMMODE_4_proguard = _Parametres.Item(4).Valeur
                _PARAMMODE_5_fs20 = _Parametres.Item(5).Valeur
                _PARAMMODE_6_lacrosse = _Parametres.Item(6).Valeur
                _PARAMMODE_7_hideki = _Parametres.Item(7).Valeur
                _PARAMMODE_8_ad = _Parametres.Item(8).Valeur
                _PARAMMODE_9_mertik = _Parametres.Item(9).Valeur
                _PARAMMODE_10_visonic = _Parametres.Item(10).Valeur
                _PARAMMODE_11_ati = _Parametres.Item(11).Valeur
                _PARAMMODE_12_oregon = _Parametres.Item(12).Valeur
                _PARAMMODE_13_meiantech = _Parametres.Item(13).Valeur
                _PARAMMODE_14_heeu = _Parametres.Item(14).Valeur
                _PARAMMODE_15_ac = _Parametres.Item(15).Valeur
                _PARAMMODE_16_arc = _Parametres.Item(16).Valeur
                _PARAMMODE_17_x10 = _Parametres.Item(17).Valeur
                _PARAMMODE_18_blindst0 = _Parametres.Item(18).Valeur
                _PARAMMODE_19_Imagintronix = _Parametres.Item(19).Valeur
                _PARAMMODE_20_sx = _Parametres.Item(20).Valeur
                _PARAMMODE_21_rsl = _Parametres.Item(21).Valeur
                _PARAMMODE_22_lighting4 = _Parametres.Item(22).Valeur
                _PARAMMODE_23_fineoffset = _Parametres.Item(23).Valeur
                _PARAMMODE_24_rubicson = _Parametres.Item(24).Valeur
                _PARAMMODE_25_ae = _Parametres.Item(25).Valeur
                _PARAMMODE_26_blindst1 = _Parametres.Item(26).Valeur
            End If

        Catch ex As Exception
            WriteLog("ERR: Erreur dans les paramétres avancés. utilisation des valeur par défaut : " & ex.Message)
        End Try

        'ouverture du port suivant le Port Com ou IP
        Try
            If _Com <> "" Then
                retour = ouvrir(_Com)
            ElseIf _IP_TCP <> "" Then
                retour = ouvrir(_IP_TCP)
            Else
                retour = "ERR: Port Com ou IP_TCP non défini. Impossible d'ouvrir le port !"
            End If
            'traitement du message de retour
            If STRGS.Left(retour, 4) = "ERR:" Then
                retour = STRGS.Right(retour, retour.Length - 5)
                WriteLog("ERR: Driver non démarré : " & retour)
            Else
                'le driver est démarré, on log puis on lance les handlers
                WriteLog("Driver démarré : " & retour)
                retour = lancer()
                If STRGS.Left(retour, 4) = "ERR:" Then
                    WriteLog("ERR: Start driver non lancé, arrêt du driver")
                    [Stop]()
                Else
                    WriteLog(retour)
                    'les handlers sont lancés, on configure le RFXtrx
                    retour = configurer()
                    If STRGS.Left(retour, 4) = "ERR:" Then
                        retour = STRGS.Right(retour, retour.Length - 5)
                        WriteLog("ERR: Start " & retour)
                        [Stop]()
                    Else
                        WriteLog(retour)
                    End If
                End If
            End If
        Catch ex As Exception
            WriteLog("ERR: Start Exception " & ex.Message)
        End Try
    End Sub

    ''' <summary>Arrêter le du driver</summary>
    ''' <remarks></remarks>
    Public Sub [Stop]() Implements HoMIDom.HoMIDom.IDriver.Stop
        Dim retour As String
        Try
            retour = fermer()
            If STRGS.Left(retour, 4) = "ERR:" Then
                retour = STRGS.Right(retour, retour.Length - 5)
                WriteLog("Stop " & retour)
            Else
                WriteLog("Stop " & retour)
            End If
        Catch ex As Exception
            WriteLog("ERR: Stop Exception " & ex.Message)
        End Try
    End Sub

    ''' <summary>Re-Démarrer le du driver</summary>
    ''' <remarks></remarks>
    Public Sub Restart() Implements HoMIDom.HoMIDom.IDriver.Restart
        [Stop]()
        Start()
    End Sub

    ''' <summary>Intérroger un device</summary>
    ''' <param name="Objet">Objet représetant le device à interroger</param>
    ''' <remarks>pas utilisé</remarks>
    Public Sub Read(ByVal Objet As Object) Implements HoMIDom.HoMIDom.IDriver.Read
        'pas utilisé
        If _Enable = False Then Exit Sub
    End Sub

    ''' <summary>Commander un device</summary>
    ''' <param name="Objet">Objet représetant le device à interroger</param>
    ''' <param name="Command">La commande à passer</param>
    ''' <param name="Parametre1"></param>
    ''' <param name="Parametre2"></param>
    ''' <remarks></remarks>
    Public Sub Write(ByVal Objet As Object, ByVal Command As String, Optional ByVal Parametre1 As Object = Nothing, Optional ByVal Parametre2 As Object = Nothing) Implements HoMIDom.HoMIDom.IDriver.Write
        Try
            If _Enable = False Then Exit Sub
            If _IsConnect = False Then
                WriteLog("Le driver n'est pas démarré, impossible d'écrire sur le port")
                Exit Sub
            End If
            If _DEBUG Then WriteLog("DBG: WRITE Device " & Objet.Name & " <-- " & Command)
            'suivant le protocole, on lance la bonne fonction
            'AC / ACEU / ANSLUT / ARC / BLYSS / ELROAB400D / EMW100 / EMW200 / IMPULS / LIGHTWAVERF / PHILIPS / RISINGSUN / WAVEMAN / X10
            Select Case UCase(Objet.Modele)
                Case "AC" 'AC : Chacon...
                    If IsNothing(parametre1) Then
                        send_Lighting2(Objet.Adresse1, Command, LIGHTING2.sTypeAC)
                    Else
                        If IsNumeric(parametre1) Then
                            send_Lighting2(Objet.Adresse1, Command, LIGHTING2.sTypeAC, CInt(parametre1))
                        Else
                            WriteLog("ERR: WRITE Le parametre " & CStr(Parametre1) & " n'est pas un entier")
                        End If
                    End If
                Case "ACEU" 'AC norme Europe
                    If IsNothing(Parametre1) Then
                        send_Lighting2(Objet.Adresse1, Command, LIGHTING2.sTypeHEU)
                    Else
                        If IsNumeric(Parametre1) Then
                            send_Lighting2(Objet.Adresse1, Command, LIGHTING2.sTypeHEU, CInt(Parametre1))
                        Else
                            WriteLog("ERR: WRITE Le parametre " & CStr(Parametre1) & " n'est pas un entier")
                        End If
                    End If
                Case "ANSLUT"
                    If IsNothing(Parametre1) Then
                        send_Lighting2(Objet.Adresse1, Command, LIGHTING2.sTypeANSLUT)
                    Else
                        If IsNumeric(Parametre1) Then
                            send_Lighting2(Objet.Adresse1, Command, LIGHTING2.sTypeANSLUT, CInt(Parametre1))
                        Else
                            WriteLog("ERR: WRITE Le parametre " & CStr(Parametre1) & " n'est pas un entier")
                        End If
                    End If
                Case "ARC" : send_lighting1(Objet.Adresse1, Command, LIGHTING1.sTypeARC)
                Case "BLYSS" : send_lighting6(Objet.Adresse1, Command, LIGHTING6.sTypeBlyss)
                Case "ELROAB400D" : send_lighting1(Objet.Adresse1, Command, LIGHTING1.sTypeAB400D)
                Case "EMW100"
                    If Not IsNothing(Parametre1) Then
                        If IsNumeric(Parametre1) Then
                            send_lighting5(Objet.Adresse1, Command, LIGHTING5.sTypeEMW100, CInt(Parametre1))
                        Else
                            WriteLog("ERR: WRITE Le parametre " & CStr(Parametre1) & " n'est pas un entier")
                        End If
                    Else
                        WriteLog("ERR: WRITE Il manque un parametre")
                    End If
                Case "EMW200" : send_lighting1(Objet.Adresse1, Command, LIGHTING1.sTypeEMW200)
                Case "IMPULS" : send_lighting1(Objet.Adresse1, Command, LIGHTING1.sTypeIMPULS)
                Case "LIGHTWAVERF"
                    If Not IsNothing(Parametre1) Then
                        If IsNumeric(Parametre1) Then
                            send_lighting5(Objet.Adresse1, Command, LIGHTING5.sTypeLightwaveRF, CInt(Parametre1))
                        Else
                            WriteLog("ERR: WRITE Le parametre " & CStr(Parametre1) & " n'est pas un entier")
                        End If
                    Else
                        WriteLog("ERR: WRITE Il manque un parametre")
                    End If
                Case "PHILIPS" : send_lighting1(Objet.Adresse1, Command, LIGHTING1.sTypePhilips)
                Case "RISINGSUN" : send_lighting1(Objet.Adresse1, Command, LIGHTING1.sTypeRisingSun)
                Case "SECURITY-KD101" : send_security(Objet.Adresse1, Command, SECURITY1.sTypeKD101)
                Case "SECURITY-SA30" : send_security(Objet.Adresse1, Command, SECURITY1.sTypeSA30)
                Case "SECURITY-X10" : send_security(Objet.Adresse1, Command, SECURITY1.sTypeSecX10)
                Case "SECURITY-MEIANTECH" : send_security(Objet.Adresse1, Command, SECURITY1.sTypeMeiantech)
                Case "SECURITY-VISONICDOOR" : send_security(Objet.Adresse1, Command, SECURITY1.sTypePowercodeSensor)
                Case "SECURITY-VISONICMOTION" : send_security(Objet.Adresse1, Command, SECURITY1.sTypePowercodeMotion)
                Case "SECURITY-VISONICCONTACT" : send_security(Objet.Adresse1, Command, SECURITY1.sTypePowercodeAux)
                Case "WAVEMAN" : send_lighting1(Objet.Adresse1, Command, LIGHTING1.sTypeWaveman)
                Case "X10" : send_lighting1(Objet.Adresse1, Command, LIGHTING1.sTypeX10)
                Case "aucun" : WriteLog("ERR: WRITE Pas de protocole d'emission pour " & Objet.Name)
                Case "" : WriteLog("ERR: WRITE Pas de protocole d'emission pour " & Objet.Name)
                Case Else : WriteLog("ERR: WRITE Protocole non géré : " & Objet.Modele.ToString.ToUpper)
            End Select
        Catch ex As Exception
            WriteLog("ERR: WRITE" & ex.ToString)
        End Try
    End Sub

    ''' <summary>Fonction lancée lors de la suppression d'un device</summary>
    ''' <param name="DeviceId">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub DeleteDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.DeleteDevice

    End Sub

    ''' <summary>Fonction lancée lors de l'ajout d'un device</summary>
    ''' <param name="DeviceId">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub NewDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.NewDevice

    End Sub

    ''' <summary>ajout des commandes avancées pour les devices</summary>
    ''' <remarks></remarks>
    Private Sub add_devicecommande(ByVal nom As String, ByVal description As String, ByVal nbparam As Integer)
        Try
            Dim x As New DeviceCommande
            x.NameCommand = nom
            x.DescriptionCommand = description
            x.CountParam = nbparam
            _DeviceCommandPlus.Add(x)
        Catch ex As Exception
            WriteLog("ERR: add_devicecommande Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>ajout Libellé pour le Driver</summary>
    ''' <param name="nom">Nom du champ : HELP</param>
    ''' <param name="labelchamp">Nom à afficher : Aide</param>
    ''' <param name="tooltip">Tooltip à afficher au dessus du champs dans l'admin</param>
    ''' <remarks></remarks>
    Private Sub Add_LibelleDriver(ByVal Nom As String, ByVal Labelchamp As String, ByVal Tooltip As String, Optional ByVal Parametre As String = "")
        Try
            Dim y0 As New HoMIDom.HoMIDom.Driver.cLabels
            y0.LabelChamp = Labelchamp
            y0.NomChamp = UCase(Nom)
            y0.Tooltip = Tooltip
            y0.Parametre = Parametre
            _LabelsDriver.Add(y0)
        Catch ex As Exception
            WriteLog("ERR: Add_LibelleDriver Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Ajout Libellé pour les Devices</summary>
    ''' <param name="nom">Nom du champ : HELP</param>
    ''' <param name="labelchamp">Nom à afficher : Aide, si = "@" alors le champ ne sera pas affiché</param>
    ''' <param name="tooltip">Tooltip à afficher au dessus du champs dans l'admin</param>
    ''' <remarks></remarks>
    Private Sub Add_LibelleDevice(ByVal Nom As String, ByVal Labelchamp As String, ByVal Tooltip As String, Optional ByVal Parametre As String = "")
        Try
            Dim ld0 As New HoMIDom.HoMIDom.Driver.cLabels
            ld0.LabelChamp = Labelchamp
            ld0.NomChamp = UCase(Nom)
            ld0.Tooltip = Tooltip
            ld0.Parametre = Parametre
            _LabelsDevice.Add(ld0)
        Catch ex As Exception
            WriteLog("ERR: Add_LibelleDevice Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>ajout de parametre avancés</summary>
    ''' <param name="nom">Nom du parametre (sans espace)</param>
    ''' <param name="description">Description du parametre</param>
    ''' <param name="valeur">Sa valeur</param>
    ''' <remarks></remarks>
    Private Sub add_paramavance(ByVal nom As String, ByVal description As String, ByVal valeur As Object)
        Try
            Dim x As New HoMIDom.HoMIDom.Driver.Parametre
            x.Nom = nom
            x.Description = description
            x.Valeur = valeur
            _Parametres.Add(x)
        Catch ex As Exception
            WriteLog("ERR: add_devicecommande Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Creation d'un objet de type</summary>
    ''' <remarks></remarks>
    Public Sub New()
        Try
            _Version = Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString

            'Parametres avancés
            add_paramavance("Debug", "Activer le Debug complet (True/False)", False)
            'add_paramavance("ParamMode", "Paramétres (ex: 20011111111111111011111111)", "20011111111111111011111111")
            'add_paramavance("AutoDiscover", "Permet de créer automatiquement des composants si ceux-ci n'existent pas encore (True/False)", False)
            add_paramavance("Frequence", "Frequence utilisée : 0=310, 1=315, 2=433, 3=868.30, 4=868.30 FSK, 5=868.35, 6=868.35 FSK, 7=868.95", 2)
            add_paramavance("Protocole Undec", "Protocole UNDEC 0=disable 1=enable", 1)
            add_paramavance("Protocole Novatis", "0=disable 1=enable", 0)
            add_paramavance("Protocole Proguard", "0=disable 1=enable", 1)
            add_paramavance("Protocole FS20", "0=disable 1=enable", 1)
            add_paramavance("Protocole Lacrosse", "0=disable 1=enable", 1)
            add_paramavance("Protocole Hideki", "0=disable 1=enable", 1)
            add_paramavance("Protocole AD", "0=disable 1=enable", 1)
            add_paramavance("Protocole Mertik", "0=disable 1=enable", 1)
            add_paramavance("Protocole Visonic", "0=disable 1=enable", 1)
            add_paramavance("Protocole ATI", "0=disable 1=enable", 1)
            add_paramavance("Protocole oregon", "0=disable 1=enable", 1)
            add_paramavance("Protocole meiantech", "0=disable 1=enable", 1)
            add_paramavance("Protocole heeu", "0=disable 1=enable", 1)
            add_paramavance("Protocole AC", "0=disable 1=enable", 1)
            add_paramavance("Protocole ARC", "0=disable 1=enable", 1)
            add_paramavance("Protocole X10", "0=disable 1=enable", 1)
            add_paramavance("Protocole Blinds t0", "0=disable 1=enable", 0)
            add_paramavance("Protocole Imagintronix", "0=disable 1=enable", 0)
            add_paramavance("Protocole SX", "0=disable 1=enable", 1)
            add_paramavance("Protocole RSL", "0=disable 1=enable", 1)
            add_paramavance("Protocole Lighting4", "0=disable 1=enable", 0)
            add_paramavance("Protocole Fineoffset", "0=disable 1=enable", 1)
            add_paramavance("Protocole Rubicson", "0=disable 1=enable", 1)
            add_paramavance("Protocole AE", "0=disable 1=enable", 1)
            add_paramavance("Protocole Blinds t1", "0=disable 1=enable", 1)

            'liste des devices compatibles
            _DeviceSupport.Add(ListeDevices.APPAREIL.ToString)
            _DeviceSupport.Add(ListeDevices.BAROMETRE.ToString)
            _DeviceSupport.Add(ListeDevices.BATTERIE.ToString)
            _DeviceSupport.Add(ListeDevices.COMPTEUR.ToString)
            _DeviceSupport.Add(ListeDevices.CONTACT.ToString)
            _DeviceSupport.Add(ListeDevices.DETECTEUR.ToString)
            _DeviceSupport.Add(ListeDevices.DIRECTIONVENT.ToString)
            _DeviceSupport.Add(ListeDevices.ENERGIEINSTANTANEE.ToString)
            _DeviceSupport.Add(ListeDevices.ENERGIETOTALE.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUEBOOLEEN.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUESTRING.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUEVALUE.ToString)
            _DeviceSupport.Add(ListeDevices.HUMIDITE.ToString)
            _DeviceSupport.Add(ListeDevices.LAMPE.ToString)
            _DeviceSupport.Add(ListeDevices.PLUIECOURANT.ToString)
            _DeviceSupport.Add(ListeDevices.PLUIETOTAL.ToString)
            _DeviceSupport.Add(ListeDevices.SWITCH.ToString)
            _DeviceSupport.Add(ListeDevices.TELECOMMANDE.ToString)
            _DeviceSupport.Add(ListeDevices.TEMPERATURE.ToString)
            _DeviceSupport.Add(ListeDevices.TEMPERATURECONSIGNE.ToString)
            _DeviceSupport.Add(ListeDevices.UV.ToString)
            _DeviceSupport.Add(ListeDevices.VITESSEVENT.ToString)
            _DeviceSupport.Add(ListeDevices.VOLET.ToString)

            'ajout des commandes avancées pour les devices
            'add_devicecommande("COMMANDE", "DESCRIPTION", nbparametre)
            add_devicecommande("GROUP_ON", "Protocole AC/ACEU/ARC/BLYSS : ON sur le groupe du composant", 1)
            add_devicecommande("GROUP_OFF", "Protocole AC/ACEU/ARC/BLYSS : OFF sur le groupe du composant", 1)
            add_devicecommande("GROUP_DIM", "Protocole AC/ACEU : DIM sur le groupe du composant", 1)
            add_devicecommande("BRIGHT", "Protocole X10 : Bright", 0)
            add_devicecommande("ALL_LIGHT_ON", "Protocole X10/EMW200/PHILIPS : ALL_LIGHT_ON", 0)
            add_devicecommande("ALL_LIGHT_OFF", "Protocole X10/EMW200/PHILIPS : ALL_LIGHT_OFF", 0)
            add_devicecommande("CHIME", "Protocole ARC : Chime", 0)
            add_devicecommande("PANIC", "Protocole SECURITY : Send PANIC alarm", 0)
            add_devicecommande("END_PANIC", "Protocole SECURITY : Send END_PANIC command", 0)
            add_devicecommande("MOTION", "Protocole SECURITY : Send MOTION alarm", 0)
            add_devicecommande("NO_MOTION", "Protocole SECURITY : Send NO_MOTION command", 0)
            add_devicecommande("PAIR", "Protocole SECURITY : Send PAIR command", 0)
            add_devicecommande("ALARM", "Protocole SECURITY : Send ALARM command", 0)
            add_devicecommande("NORMAL", "Protocole SECURITY : Send NORMAL command", 0)
            add_devicecommande("DISARM", "Protocole SECURITY : Send ALARM command", 0)
            add_devicecommande("ARM_AWAY", "Protocole SECURITY : Send ARM_AWAY command", 0)
            add_devicecommande("ARM_HOME", "Protocole SECURITY : Send ARM_HOME command", 0)
            add_devicecommande("DARK_DETECTED", "Protocole SECURITY : Send DARK_DETECTED command", 0)
            add_devicecommande("LIGHT_DETECTED", "Protocole SECURITY : Send LIGHT_DETECTED command", 0)

            'Libellé Driver
            Add_LibelleDriver("HELP", "Aide...", "Pas d'aide actuellement...")

            'Libellé Device
            Add_LibelleDevice("ADRESSE1", "Adresse", "Adresse du composant. Le format dépend du protocole")
            Add_LibelleDevice("ADRESSE2", "@", "")
            Add_LibelleDevice("SOLO", "@", "")
            Add_LibelleDevice("MODELE", "Protocole", "Nom du protocole à utiliser : aucun/AC/ACEU/ANSLUT/ARC/BLYSS/ELROAB400D/EMW100/EMW200/LIGHTWAVERF/IMPULS/PHILIPS/RISINGSUN/SECURITY-KD101/SECURITY-MEIANTECH/SECURITY-SA30/SECURITY-VISONICDOOR/SECURITY-VISONICMOTION/SECURITY-VISONICCONTACT/SECURITY-X10/WAVEMAN/X10", "aucun|AC|ACEU|ANSLUT|ARC|BLYSS|ELROAB400D|EMW100|EMW200|IMPULS|LIGHTWAVERF|PHILIPS|RISINGSUN|SECURITY-KD101|SECURITY-MEIANTECH|SECURITY-SA30|SECURITY-VISONICDOOR|SECURITY-VISONICMOTION|SECURITY-VISONICCONTACT|SECURITY-X10|WAVEMAN|X10")
            Add_LibelleDevice("REFRESH", "@", "")
            'Add_LibelleDevice("LASTCHANGEDUREE", "LastChange Durée", "")

        Catch ex As Exception
            WriteLog("ERR: New Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Si refresh >0 gestion du timer</summary>
    ''' <remarks>PAS UTILISE CAR IL FAUT LANCER UN TIMER QUI LANCE/ARRETE CETTE FONCTION dans Start/Stop</remarks>
    Private Sub TimerTick(ByVal source As Object, ByVal e As System.Timers.ElapsedEventArgs)

    End Sub

#End Region

#Region "Fonctions Internes"

    ''' <summary>Ouvrir le port COM/ETHERNET</summary>
    ''' <param name="numero">Nom/Numero du port COM/Adresse IP: COM2</param>
    ''' <remarks></remarks>
    Private Function ouvrir(ByVal numero As String) As String
        'Forcer le . 
        Try
            'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
            'My.Application.ChangeCulture("en-US")
            If Not _IsConnect Then
                port_name = numero 'pour se rapeller du nom du port
                If VB.Left(numero, 3) <> "COM" Then
                    'RFXtrx est un modele ethernet
                    If My.Computer.Network.Ping(numero) Then
                        tcp = True
                        client = New TcpClient(numero, _Port_TCP)
                        _IsConnect = True
                        dateheurelancement = DateTime.Now
                        Return ("Port IP " & port_name & ":" & _Port_TCP & " ouvert")
                    Else
                        Return ("ERR: Le RFXtrx de répond pas au ping")
                    End If

                Else
                    'RFXtrx est un modele usb
                    tcp = False
                    RS232Port.PortName = port_name 'nom du port : COM1
                    RS232Port.BaudRate = 38400 'vitesse du port 300, 600, 1200, 2400, 4800, 9600, 14400, 19200, 38400, 57600, 115200
                    RS232Port.Parity = Parity.None 'pas de parité
                    RS232Port.StopBits = StopBits.One '1 bit d'arrêt par octet
                    RS232Port.DataBits = 8 'nombre de bit par octet
                    'RS232Port.Encoding = System.Text.Encoding.GetEncoding(1252)  'Extended ASCII (8-bits)
                    RS232Port.Handshake = Handshake.None
                    RS232Port.ReadBufferSize = CInt(4096)
                    'RS232Port.ReceivedBytesThreshold = 1
                    RS232Port.ReadTimeout = 100
                    RS232Port.WriteTimeout = 500
                    RS232Port.Open()
                    _IsConnect = True
                    If RS232Port.IsOpen Then
                        GC.SuppressFinalize(RS232Port.BaseStream)
                        RS232Port.DtrEnable = True
                        RS232Port.RtsEnable = True
                        RS232Port.DiscardInBuffer()
                    End If
                    gRecComPortEnabled = True
                    Return ("Port " & port_name & " ouvert")
                End If
            Else
                Return ("Port " & port_name & " dejà ouvert")
            End If
        Catch ex As Exception
            Return ("ERR: " & ex.Message)
        End Try
    End Function

    ''' <summary>Lances les handlers sur le port</summary>
    ''' <remarks></remarks>
    Private Function lancer() As String
        Try
            'lancer les handlers
            If tcp Then
                Try
                    stream = client.GetStream()
                    stream.BeginRead(TCPData, 0, 1024, AddressOf TCPDataReceived, Nothing)
                    Return "Handler IP OK"
                Catch ex As Exception
                    WriteLog("ERR: LANCER GETSTREAM Exception : " & ex.Message)
                    Return "ERR: Handler IP"
                End Try
            Else
                Try
                    AddHandler RS232Port.DataReceived, New SerialDataReceivedEventHandler(AddressOf DataReceived)
                    AddHandler RS232Port.ErrorReceived, New SerialErrorReceivedEventHandler(AddressOf ReadErrorEvent)
                    Return "Handler COM OK"
                Catch ex As Exception
                    WriteLog("ERR: LANCER Serial Exception : " & ex.Message)
                    Return "ERR: Handler COM"
                End Try
            End If
            recbuf(0) = 0
            maxticks = 0

            ''tmrRead.Enabled = True
            'If tmrRead.Enabled Then MyTimer.Stop()
            'tmrRead.Interval = 100
            'tmrRead.Start()
        Catch ex As Exception
            WriteLog("ERR: LANCER Serial Exception : " & ex.Message)
            Return "ERR: Exception"
        End Try
    End Function

    ''' <summary>Configurer le RFXtrx</summary>
    ''' <remarks></remarks>
    Private Function configurer() As String
        'configurer le RFXtrx
        Try
            'get firmware version
            SendCommand(ICMD.cmdRESET, "Reset receiver/transceiver:")
            System.Threading.Thread.Sleep(2000)
            'configure Transceiver mode
            'SetMode(_PARAMMODE)
            SetMode2()
            System.Threading.Thread.Sleep(2000)

            dateheurelancement = DateTime.Now

            Return "Configuration OK"
        Catch ex As Exception
            WriteLog("ERR: LANCER Configuration Exception : " & ex.Message)
            Return "ERR: Configuration NOK"
        End Try
    End Function

    ''' <summary>Ferme la connexion au port</summary>
    ''' <remarks></remarks>
    Private Function fermer() As String
        Try
            If _IsConnect Then
                'fermeture des ports
                _IsConnect = False
                'If tmrRead.Enabled Then tmrRead.Stop()
                If tcp Then
                    client.Close()
                    stream.Close()
                    Return ("Port IP fermé")
                Else
                    'suppression de l'attente de données à lire
                    gRecComPortEnabled = False
                    RemoveHandler RS232Port.DataReceived, AddressOf DataReceived
                    RemoveHandler RS232Port.ErrorReceived, AddressOf ReadErrorEvent
                    If (Not (RS232Port Is Nothing)) Then ' The COM port exists.
                        If RS232Port.IsOpen Then
                            Dim limite As Integer = 0
                            'vidage des tampons
                            RS232Port.DiscardInBuffer()
                            RS232Port.DiscardOutBuffer()
                            'au cas on verifie si encore quelque chose à lire
                            Do While (RS232Port.BytesToWrite > 0 And limite < 100) ' Wait for the transmit buffer to empty.
                                limite = limite + 1
                            Loop
                            limite = 0
                            Do While (RS232Port.BytesToRead > 0 And limite < 100) ' Wait for the receipt buffer to empty.
                                limite = limite + 1
                            Loop

                            GC.ReRegisterForFinalize(RS232Port.BaseStream)
                            RS232Port.Close()
                            RS232Port.Dispose()
                            Return ("Port " & port_name & " fermé")
                        End If
                        Return ("Port " & port_name & "  est déjà fermé")
                    End If
                    Return ("Port " & port_name & " n'existe pas")
                End If
                tcp = False
            End If
        Catch ex As UnauthorizedAccessException
            Return ("ERR: Port " & port_name & " IGNORE") ' The port may have been removed. Ignore.
        Catch ex As Exception
            Return ("ERR: Port " & port_name & " : " & ex.Message)
        End Try
        Return "ERR: Not defined"
    End Function

    ''' <summary>ecrire sur le port</summary>
    ''' <param name="commande">premier paquet à envoyer</param>
    ''' <remarks></remarks>
    Private Function ecrire(ByVal commande() As Byte) As String
        Try
            SyncLock rfxtrxlock 'lock pour etre sur de ne pas faire deux operations en meme temps 
                If tcp Then
                    stream.Write(commande, 0, commande.Length)
                    bytSeqNbr = CByte(bytSeqNbr + 1)
                Else
                    RS232Port.Write(commande, 0, commande.Length)
                    bytSeqNbr = CByte(bytSeqNbr + 1)
                End If
            End SyncLock
            Return ""
        Catch ex As Exception
            Return ("ERR: " & ex.Message)
        End Try
    End Function

    Private Sub SendCommand(ByVal command As Byte, ByRef message As String)
        Try

            Dim kar(ICMD.size) As Byte
            kar(ICMD.packetlength) = ICMD.size
            kar(ICMD.packettype) = ICMD.pTypeInterfaceControl
            kar(ICMD.subtype) = ICMD.sTypeInterfaceCommand
            kar(ICMD.seqnbr) = bytSeqNbr
            kar(ICMD.cmnd) = command
            kar(ICMD.msg1) = 0
            kar(ICMD.msg2) = 0
            kar(ICMD.msg3) = 0
            kar(ICMD.msg4) = 0
            kar(ICMD.msg5) = 0
            kar(ICMD.msg6) = 0
            kar(ICMD.msg7) = 0
            kar(ICMD.msg8) = 0
            kar(ICMD.msg9) = 0

            For Each bt As Byte In kar
                message = message + VB.Right("0" & Hex(bt), 2)
            Next
            WriteLog("SendCommand : " & message)

            Try
                ecrire(kar)
            Catch exc As Exception
                WriteLog("ERR: SendCommand: Unable to write to port")
            End Try
            kar = Nothing
        Catch ex As Exception
            WriteLog("ERR: SendCommand Exception : " & ex.Message)
        End Try
    End Sub
    'Private Sub ecrirecommande(ByVal kar As Byte())
    '    Dim message As String
    '    Try
    '        Dim temp, tcpdata(0) As Byte
    '        Dim ok(0) As Byte
    '        Dim intIndex, intEnd As Integer
    '        Dim Finish As Double
    '        ack_ok = True

    '        Finish = VB.DateAndTime.Timer + 3.0   ' wait for ACK, max 3-seconds

    '        Do While (ack = False)
    '            If VB.DateAndTime.Timer > Finish Then
    '                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "RFXtrx", "No ACK received witin 3 seconds !")
    '                ack_ok = False
    '                Exit Do
    '            End If

    '            If tcp = True Then
    '                ' As long as there is information, read one byte at a time and output it.
    '                While stream.DataAvailable
    '                    stream.Read(tcpdata, 0, 1)
    '                    temp = tcpdata(0)
    '                    If temp = protocolsynchro Then
    '                        _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "RFXtrx", "ACK  => " & VB.Right("0" & Hex(temp), 2))
    '                    ElseIf temp = &H5A Then
    '                        _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "RFXtrx", "NAK  => " & VB.Right("0" & Hex(temp), 2))
    '                    End If
    '                    mess = True
    '                End While
    '            Else
    '                Try
    '                    ' As long as there is information, read one byte at a time and 
    '                    '   output it.
    '                    While (port.BytesToRead() > 0)
    '                        ' Write the output to the screen.
    '                        temp = port.ReadByte()
    '                        If temp = protocolsynchro Then
    '                            _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "RFXtrx", "ACK  => " & VB.Right("0" & Hex(temp), 2))
    '                        ElseIf temp = &H5A Then
    '                            _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "RFXtrx", "NAK  => " & VB.Right("0" & Hex(temp), 2))
    '                        End If
    '                        mess = True
    '                    End While
    '                Catch exc As Exception
    '                    ' An exception is raised when there is no information to read : Don't do anything here, just let the exception go.
    '                End Try
    '            End If

    '            If mess Then
    '                ack = True
    '                mess = False
    '            End If
    '        Loop

    '        ack = False

    '        ' Write a user specified Command to the Port.
    '        Try
    '            ecrire(kar)
    '        Catch exc As Exception
    '            ' Warn the user.
    '            _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "RFXtrx", "Unable to write to port")
    '            ack_ok = False
    '        Finally

    '        End Try

    '        message = VB.Right("0" & Hex(kar(0)), 2)
    '        intEnd = ((kar(0) And &HF8) / 8)
    '        If (kar(0) And &H7) <> 0 Then
    '            intEnd += 1
    '        End If
    '        For intIndex = 1 To intEnd
    '            message = message + VB.Right("0" & Hex(kar(intIndex)), 2)
    '        Next
    '        _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "RFXtrx Ecrirecommande", message)
    '        ack = False
    '    Catch ex As Exception
    '        _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "RFXtrx ecrirecommande", ex.Message)
    '    End Try
    'End Sub

    ''' <summary>Executer lors de la reception d'une donnée sur le port</summary>
    ''' <remarks></remarks>
    Private Sub DataReceived(ByVal sender As Object, ByVal e As SerialDataReceivedEventArgs)
        Try
            While gRecComPortEnabled And RS232Port.BytesToRead() > 0
                ProcessReceivedChar(CByte(RS232Port.ReadByte()))
            End While
        Catch Ex As Exception
            WriteLog("ERR: Datareceived Exception : " & Ex.Message)
        End Try
    End Sub

    ''' <summary>Executer lors de la reception d'une erreur sur le port</summary>
    ''' <remarks></remarks>
    Private Sub ReadErrorEvent(ByVal sender As Object, ByVal ev As SerialErrorReceivedEventArgs)
        Try
            While gRecComPortEnabled And RS232Port.BytesToRead() > 0
                ProcessReceivedChar(CByte(RS232Port.ReadByte()))
            End While
        Catch Ex As Exception
            WriteLog("ERR: ReadErrorEvent Exception : " & Ex.Message)
        End Try
    End Sub

    ''' <summary>Executer lors de la reception d'une donnée sur le port IP</summary>
    ''' <remarks></remarks>
    Private Sub TCPDataReceived(ByVal ar As IAsyncResult)
        Dim intCount As Integer
        Try
            If _IsConnect Then
                intCount = stream.EndRead(ar)
                ProcessNewTCPData(TCPData, 0, intCount)
                stream.BeginRead(TCPData, 0, 1024, AddressOf TCPDataReceived, Nothing)
            End If
        Catch Ex As Exception
            WriteLog("ERR: TCPDatareceived Exception : " & Ex.Message)
        End Try
    End Sub

    ''' <summary>Traite les données IP recu</summary>
    ''' <remarks></remarks>
    Private Sub ProcessNewTCPData(ByVal Bytes() As Byte, ByVal offset As Integer, ByVal count As Integer)
        Dim intIndex As Integer
        Try
            For intIndex = offset To offset + count - 1
                ProcessReceivedChar(Bytes(intIndex))
            Next
        Catch ex As Exception
            WriteLog("ERR: ProcessNewTCPData Exception : " & ex.Message)
        End Try
    End Sub

    ' ''' <summary>xxx</summary>
    ' ''' <remarks></remarks>
    'Private Sub tmrRead_Elapsed(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tmrRead.Elapsed
    '    Try
    '        If Resettimer <= 0 Then
    '            If recbytes <> 0 Then 'one or more bytes received
    '                maxticks += 1
    '                If maxticks > 3 Then 'flush buffer due to 400ms timeout
    '                    maxticks = 0
    '                    recbytes = 0
    '                    _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "RFXtrx", "Buffer flushed due to timeout")
    '                End If
    '            End If
    '        Else
    '            Resettimer = Resettimer - 1    ' decrement resettimer
    '            If Resettimer = 0 Then
    '                If gRecComPortEnabled Then
    '                    RS232Port.DiscardInBuffer()
    '                Else
    '                    'stream.Flush() 'flush not yet supported
    '                End If
    '                SendCommand(ICMD.STATUS, "Get Status:")
    '                maxticks = 0
    '            End If
    '        End If
    '    Catch ex As Exception
    '        _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "RFXtrx tmrRead_Elapsed", ex.Message)
    '    End Try
    'End Sub

    ''' <summary>Traite chaque byte reçu</summary>
    ''' <param name="sComChar">Byte recu</param>
    ''' <remarks></remarks>
    Private Sub ProcessReceivedChar(ByVal sComChar As Byte)
        Try
            'If Resettimer <> 0 Then
            '    Exit Sub 'ignore received characters after a reset cmd until resettimer = 0
            'End If

            maxticks = 0    'reset receive timeout

            If recbytes = 0 Then    '1st char of a packet received
                If sComChar <> 0 Then
                    If LogActive Then
                        LogFile.WriteLine()
                    End If
                Else
                    Return  'ignore 1st char if 00
                End If
            End If

            recbuf(recbytes) = sComChar 'store received char
            recbytes += 1               'increment char counter

            If recbytes > recbuf(0) Then 'all bytes of the packet received?
                'Write the output to the screen for DEBUG
                messagerecu = messagerecu & VB.Right("0" & Hex(sComChar), 2)
                If _DEBUG Then WriteLog("DBG: Message Reçu : " & messagerecu)
                messagerecu = ""

                decode_messages()  'decode message
                recbytes = 0    'set to zero to receive next message
            Else
                messagerecu = messagerecu & VB.Right("0" & Hex(sComChar), 2) 'get message recu for debug
            End If

        Catch ex As Exception
            WriteLog("ERR: ProcessReceivedChar Exception : " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Decode messages"
    Private Sub decode_messages()
        Try
            Select Case recbuf(1)
                Case ICMD.pTypeInterfaceControl : decode_InterfaceControl()
                Case IRESPONSE.pTypeInterfaceMessage : decode_InterfaceMessage()
                Case RXRESPONSE.pTypeRecXmitMessage : decode_RecXmitMessage()
                Case UNDECODED.pTypeUndecoded : decode_UNDECODED()
                Case LIGHTING1.pTypeLighting1 : decode_Lighting1()
                Case LIGHTING2.pTypeLighting2 : decode_Lighting2()
                Case LIGHTING3.pTypeLighting3 : decode_Lighting3()
                Case LIGHTING4.pTypeLighting4 : decode_Lighting4()
                Case LIGHTING5.pTypeLighting5 : decode_Lighting5()
                Case LIGHTING6.pTypeLighting6 : decode_Lighting6()
                Case CHIME.pTypeChime : decode_Chime()
                Case CURTAIN1.pTypeCurtain : decode_Curtain1()
                Case SECURITY1.pTypeSecurity1 : decode_Security1()
                Case CAMERA1.pTypeCamera : decode_Camera1()
                Case BLINDS1.pTypeBlinds : decode_BLINDS1()
                Case RFY.pTypeRFY : decode_RFY()
                Case REMOTE.pTypeRemote : decode_Remote()
                Case THERMOSTAT1.pTypeThermostat1 : decode_Thermostat1()
                Case THERMOSTAT2.pTypeThermostat2 : decode_Thermostat2()
                Case THERMOSTAT3.pTypeThermostat3 : decode_Thermostat3()
                Case BBQ.pTypeBBQ : decode_BBQ()
                Case TEMP_RAIN.pTypeTEMP_RAIN : decode_TempRain()
                Case TEMP.pTypeTEMP : decode_Temp()
                Case HUM.pTypeHUM : decode_Hum()
                Case TEMP_HUM.pTypeTEMP_HUM : decode_TempHum()
                Case BARO.pTypeBARO : decode_Baro()
                Case TEMP_HUM_BARO.pTypeTEMP_HUM_BARO : decode_TempHumBaro()
                Case RAIN.pTypeRAIN : decode_Rain()
                Case WIND.pTypeWIND : decode_Wind()
                Case UV.pTypeUV : decode_UV()
                Case DT.pTypeDT : decode_DateTime()
                Case CURRENT.pTypeCURRENT : decode_Current()
                Case ENERGY.pTypeENERGY : decode_Energy()
                Case CURRENT_ENERGY.pTypeCURRENTENERGY : decode_Current_Energy()
                Case POWER.pTypePOWER : decode_Power()
                Case GAS.pTypeGAS : decode_Gas()
                Case WATER.pTypeWATER : decode_Water()
                Case WEIGHT.pTypeWEIGHT : decode_Weight()
                Case RFXSENSOR.pTypeRFXSensor : decode_RFXSensor()
                Case RFXMETER.pTypeRFXMeter : decode_RFXMeter()
                Case FS20.pTypeFS20 : decode_FS20()
                Case RAW.pTypeRAW : decode_RAW()
                Case Else : _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "RFXtrx decode_messages", "ERROR: Unknown Packet type:" & Hex(recbuf(1)))
            End Select
        Catch ex As Exception
            WriteLog("ERR: decode_messages : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_InterfaceControl()
        Try
            Dim messagelog As String = ""
            Select Case recbuf(ICMD.subtype)
                Case ICMD.sTypeInterfaceCommand
                    'WriteMessage("subtype           = Mode Command")
                    'WriteMessage("Sequence nbr      = " & recbuf(ICMD.seqnbr).ToString)
                    messagelog = "Interface: Command="
                    Select Case recbuf(ICMD.cmnd)
                        Case &H0 : messagelog &= "Reset the receiver/transceiver. No answer is transmitted!"
                        Case &H1, &H4, &H5, &H7, &H52, &H53, &H54 : messagelog &= "not used"
                        Case &H2 : messagelog &= "Get Status, return firmware versions and configuration of the interface"
                        Case &H3
                            messagelog &= "Set Mode msg1-msg5, return also the firmware version and configuration of the interface"
                            messagelog &= " Frequency="
                            Select Case recbuf(ICMD.msg1)
                                Case &H50 : messagelog &= "310MHz"
                                Case &H51 : messagelog &= "315MHz"
                                Case &H52 : messagelog &= "433.92MHz receive only"
                                Case &H53 : messagelog &= "433.92MHz transceiver"
                                Case &H55 : messagelog &= "868.00MHz"
                                Case &H56 : messagelog &= "868.00MHz FSK"
                                Case &H57 : messagelog &= "868.30MHz"
                                Case &H58 : messagelog &= "868.30MHz FSK"
                                Case &H59 : messagelog &= "868.35MHz"
                                Case &H5A : messagelog &= "868.35MHz FSK"
                                Case &H5B : messagelog &= "868.95MHz"
                            End Select
                            If (recbuf(ICMD.msg3) And IRESPONSE.msg3_undec) = 0 Then messagelog &= "Undec=off" Else messagelog &= "Undec=on"
                            If (recbuf(ICMD.msg3) And IRESPONSE.msg3_IMAGINTRONIX) = 0 Then messagelog &= ", Imagintronix=off" Else messagelog &= ", Imagintronix=on"
                            If (recbuf(ICMD.msg3) And IRESPONSE.msg3_SX) = 0 Then messagelog &= ", ByronSX=off" Else messagelog &= ", ByronSX=on"
                            If (recbuf(ICMD.msg3) And IRESPONSE.msg3_RSL) = 0 Then messagelog &= ", RSL=off" Else messagelog &= ", RSL=on"
                            If (recbuf(ICMD.msg3) And IRESPONSE.msg3_LIGHTING4) = 0 Then messagelog &= ", LIGHTING4=off" Else messagelog &= ", LIGHTING4=on"
                            If (recbuf(ICMD.msg3) And &H4) = 0 Then messagelog &= ", FineOffset=off" Else messagelog &= ", FineOffset=on"
                            If (recbuf(ICMD.msg3) And &H2) = 0 Then messagelog &= ", Rubicson=off" Else messagelog &= ", Rubicson=on"
                            If (recbuf(ICMD.msg3) And &H1) = 0 Then messagelog &= ", AE Blyss=off" Else messagelog &= ", AE Blyss=on"
                            If (recbuf(ICMD.msg4) And &H80) = 0 Then messagelog &= ", BlindsT1/T2/T3/T4=off" Else messagelog &= ", BlindsT1/T2/T3/T4=on"
                            If (recbuf(ICMD.msg4) And &H40) = 0 Then messagelog &= ", BlindsT0=off" Else messagelog &= ", BlindsT0=on"
                            If (recbuf(ICMD.msg4) And &H20) = 0 Then messagelog &= ", Proguard=off" Else messagelog &= ", Proguard=on"
                            If (recbuf(ICMD.msg4) And &H10) = 0 Then messagelog &= ", FS20=off" Else messagelog &= ", FS20=on"
                            If (recbuf(ICMD.msg4) And &H8) = 0 Then messagelog &= ", La Crosse=off" Else messagelog &= ", La Crosse=on"
                            If (recbuf(ICMD.msg4) And &H4) = 0 Then messagelog &= ", Hideki/UPM=off" Else messagelog &= ", Hideki/UPM=on"
                            If (recbuf(ICMD.msg4) And &H2) = 0 Then messagelog &= ", AD LightwaveRF=off" Else messagelog &= ", AD LightwaveRF=on"
                            If (recbuf(ICMD.msg4) And &H1) = 0 Then messagelog &= ", Mertik Maxitrol=off" Else messagelog &= ", Mertik Maxitrol=on"
                            If (recbuf(ICMD.msg5) And &H80) = 0 Then messagelog &= ", Visonic=off" Else messagelog &= ", Visonic=on"
                            If (recbuf(ICMD.msg5) And &H40) = 0 Then messagelog &= ", ATI=off" Else messagelog &= ", ATI=on"
                            If (recbuf(ICMD.msg5) And &H20) = 0 Then messagelog &= ", Oregon Scientific=off" Else messagelog &= ", Oregon Scientific=on"
                            If (recbuf(ICMD.msg5) And &H10) = 0 Then messagelog &= ", Meiantech/Atlantic=off" Else messagelog &= ", Meiantech/Atlantic=on"
                            If (recbuf(ICMD.msg5) And &H8) = 0 Then messagelog &= ", HomeEasy EU=off" Else messagelog &= ", HomeEasy EU=on"
                            If (recbuf(ICMD.msg5) And &H4) = 0 Then messagelog &= ", AC HomeEasy,KAKU..=off" Else messagelog &= ", AC HomeEasy,KAKU.=on"
                            If (recbuf(ICMD.msg5) And &H2) = 0 Then messagelog &= ", ARC HomeEasy,.....=off" Else messagelog &= ", ARC HomeEasy,.....=on"
                            If (recbuf(ICMD.msg5) And &H1) = 0 Then messagelog &= ", X10=off" Else messagelog &= ", X10=on"
                        Case &H6 : messagelog &= "Save receiving modes of the receiver/transceiver in non-volatile memory"
                        Case &H8 : messagelog &= "T1 – for internal use by RFXCOM"
                        Case &H9 : messagelog &= "T2 – for internal use by RFXCOM"
                        Case &H50 : messagelog &= "select 310MHz in the 310/315 transceiver"
                        Case &H51 : messagelog &= "select 315MHz in the 310/315 transceiver"
                        Case &H55 : messagelog &= "select 868.00MHz in the 868 transceiver"
                        Case &H56 : messagelog &= "select 868.00MHz FSK in the 868 transceiver"
                        Case &H57 : messagelog &= "select 868.30MHz in the 868 transceiver"
                        Case &H58 : messagelog &= "select 868.30MHz FSK in the 868 transceiver"
                        Case &H59 : messagelog &= "select 868.35MHz in the 868 transceiver"
                        Case &H5A : messagelog &= "select 868.35MHz FSK in the 868 transceiver"
                        Case &H5B : messagelog &= "select 868.95MHz in the 868 transceiver"
                        Case Else : messagelog &= "ERROR: Unknown cmnd:" & Hex(recbuf(ICMD.cmnd))
                    End Select
                Case Else
                    messagelog &= "ERROR: Unknown subtype:" & Hex(recbuf(ICMD.subtype))
            End Select
            WriteLog(messagelog)
            messagelog = Nothing
        Catch ex As Exception
            WriteLog("ERR: decode_InterfaceControl : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_InterfaceMessage()
        Try
            Dim messagelog As String = ""
            Select Case recbuf(IRESPONSE.subtype)
                Case IRESPONSE.sTypeInterfaceResponse
                    '        WriteMessage("subtype           = Interface Response")
                    '        WriteMessage("Sequence nbr      = " & recbuf(IRESPONSE.seqnbr).ToString)
                    Select Case recbuf(IRESPONSE.cmnd)
                        Case ICMD.cmdSTATUS, ICMD.cmdSETMODE, ICMD.cmdSAVE, ICMD.cmd310, ICMD.cmd315, ICMD.cmd800, ICMD.cmd800F, ICMD.cmd830, ICMD.cmd830F, ICMD.cmd835, ICMD.cmd835F, ICMD.cmd895
                            messagelog = "Interface: Command="
                            Select Case recbuf(IRESPONSE.cmnd)
                                Case ICMD.cmdSTATUS : messagelog &= "Get Status"
                                Case ICMD.cmdSETMODE : messagelog &= "Set Mode"
                                Case ICMD.cmdSAVE : messagelog &= "Save Settings"
                                Case ICMD.cmd310 : messagelog &= "Select 310MHz"
                                Case ICMD.cmd315 : messagelog &= "Select 315MHz"
                                Case ICMD.cmd800 : messagelog &= "Select 868.00MHz"
                                Case ICMD.cmd800F : messagelog &= "Select 868.00MHz FSK"
                                Case ICMD.cmd830 : messagelog &= "Select 868.30MHz"
                                Case ICMD.cmd830F : messagelog &= "Select 868.30MHz FSK"
                                Case ICMD.cmd835 : messagelog &= "Select 868.35MHz"
                                Case ICMD.cmd835F : messagelog &= "Select 868.35MHz FSK"
                                Case ICMD.cmd895 : messagelog &= "Select 868.95MHz"
                                Case Else : messagelog &= "Error unknown response"
                            End Select
                            messagelog &= ", Type="
                            Select Case recbuf(IRESPONSE.msg1)
                                Case IRESPONSE.recType310 : messagelog &= "Transceiver 310MHz"
                                Case IRESPONSE.recType315 : messagelog &= "Receiver 315MHz"
                                Case IRESPONSE.recType43392 : messagelog &= "Receiver 433.92MHz"
                                Case IRESPONSE.trxType43392 : messagelog &= "Transceiver 433.92MHz"
                                Case IRESPONSE.trxType43342 : messagelog &= "Transceiver 433.42MHz"
                                Case IRESPONSE.recType86800 : messagelog &= "Receiver 868.00MHz"
                                Case IRESPONSE.recType86800FSK : messagelog &= "Receiver 868.00MHz FSK"
                                Case IRESPONSE.recType86830 : messagelog &= "Receiver 868.30MHz"
                                Case IRESPONSE.recType86830FSK : messagelog &= "Receiver 868.30MHz FSK"
                                Case IRESPONSE.recType86835 : messagelog &= "Receiver 868.35MHz"
                                Case IRESPONSE.recType86835FSK : messagelog &= "Receiver 868.35MHz FSK"
                                Case IRESPONSE.recType86895 : messagelog &= "Receiver 868.95MHz"
                                Case Else : messagelog &= "Receiver unknown"
                            End Select
                            trxType = recbuf(IRESPONSE.msg1)
                            messagelog &= ", Firmware=" & recbuf(IRESPONSE.msg2)
                            messagelog &= ", Hardware=" & recbuf(IRESPONSE.msg6) & "." & recbuf(IRESPONSE.msg7)
                            WriteLog(messagelog)
                            bytFWversion = recbuf(IRESPONSE.msg2)

                            messagelog = "Protocole: "
                            If (recbuf(IRESPONSE.msg3) And IRESPONSE.msg3_undec) <> 0 Then messagelog &= "Undec=on" Else messagelog &= "Undec=off"
                            If (recbuf(IRESPONSE.msg5) And IRESPONSE.msg5_X10) <> 0 Then messagelog &= ", X10=on" Else messagelog &= ", X10=off"
                            If (recbuf(IRESPONSE.msg5) And IRESPONSE.msg5_ARC) <> 0 Then messagelog &= ", ARC=on" Else messagelog &= ", ARC=off"
                            If (recbuf(IRESPONSE.msg5) And IRESPONSE.msg5_AC) <> 0 Then messagelog &= ", AC=on" Else messagelog &= ", AC=off"
                            If (recbuf(IRESPONSE.msg5) And IRESPONSE.msg5_HEU) <> 0 Then messagelog &= ", HomeEasyEU=on" Else messagelog &= ", HomeEasyEU=off"
                            If (recbuf(IRESPONSE.msg5) And IRESPONSE.msg5_MEI) <> 0 Then messagelog &= ", Meiantech=on" Else messagelog &= ", Meiantech=off"
                            If (recbuf(IRESPONSE.msg5) And IRESPONSE.msg5_OREGON) <> 0 Then messagelog &= ", OregonScientific=on" Else messagelog &= ", OregonScientific=off"
                            If (recbuf(IRESPONSE.msg5) And IRESPONSE.msg5_ATI) <> 0 Then messagelog &= ", ATI=on" Else messagelog &= ", ATI=off"
                            If (recbuf(IRESPONSE.msg5) And IRESPONSE.msg5_VISONIC) <> 0 Then messagelog &= ", Visonic=on" Else messagelog &= ", Visonic=off"
                            If (recbuf(IRESPONSE.msg4) And IRESPONSE.msg4_MERTIK) <> 0 Then messagelog &= ", Mertik=on" Else messagelog &= ", Mertik=off"
                            If (recbuf(IRESPONSE.msg4) And IRESPONSE.msg4_AD) <> 0 Then messagelog &= ", AD=on" Else messagelog &= ", AD=off"
                            If (recbuf(IRESPONSE.msg4) And IRESPONSE.msg4_HID) <> 0 Then messagelog &= ", Hideki=on" Else messagelog &= ", Hideki=off"
                            If (recbuf(IRESPONSE.msg4) And IRESPONSE.msg4_LCROS) <> 0 Then messagelog &= ", LaCrosse=on" Else messagelog &= ", LaCrosse=off"
                            If (recbuf(IRESPONSE.msg4) And IRESPONSE.msg4_FS20) <> 0 Then messagelog &= ", FS20=on" Else messagelog &= ", FS20=off"
                            If (recbuf(IRESPONSE.msg4) And IRESPONSE.msg4_PROGUARD) <> 0 Then messagelog &= ", ProGuard=on" Else messagelog &= ", ProGuard=off"
                            If (recbuf(IRESPONSE.msg4) And IRESPONSE.msg4_BlindsT0) <> 0 Then messagelog &= ", BlindsT0=on" Else messagelog &= ", BlindsT0=off"
                            If (recbuf(IRESPONSE.msg4) And IRESPONSE.msg4_BlindsT1) <> 0 Then messagelog &= ", BlindsT1=on" Else messagelog &= ", BlindsT1=off"
                            If (recbuf(IRESPONSE.msg3) And IRESPONSE.msg3_AE) <> 0 Then messagelog &= ", AE=on" Else messagelog &= ", AE=off"
                            If (recbuf(IRESPONSE.msg3) And IRESPONSE.msg3_RUBICSON) <> 0 Then messagelog &= ", RUBiCSON=on" Else messagelog &= ", RUBiCSON=off"
                            If (recbuf(IRESPONSE.msg3) And IRESPONSE.msg3_FINEOFFSET) <> 0 Then messagelog &= ", FineOffset=on" Else messagelog &= ", FineOffset=off"
                            If (recbuf(IRESPONSE.msg3) And IRESPONSE.msg3_LIGHTING4) <> 0 Then messagelog &= ", Lighting4=on" Else messagelog &= ", Lighting4=off"
                            If (recbuf(IRESPONSE.msg3) And IRESPONSE.msg3_RSL) <> 0 Then messagelog &= ", RSL=on" Else messagelog &= ", RSL=off"
                            If (recbuf(IRESPONSE.msg3) And IRESPONSE.msg3_SX) <> 0 Then messagelog &= ", Byron SX=on" Else messagelog &= ", Byron SX=off"
                            If (recbuf(IRESPONSE.msg3) And IRESPONSE.msg3_IMAGINTRONIX) <> 0 Then messagelog &= ", Imagintronix=on" Else messagelog &= ", Imagintronix=off"

                            WriteLog(messagelog)
                            'Case ICMD.ENABLEALL : WriteLog("Réponse à : Enable All RF")
                            'Case ICMD.UNDECODED : WriteLog("Réponse à : UNDECODED on")
                            'Case ICMD.SAVE : WriteLog("Réponse à : Save")
                            'Case ICMD.DISX10 : WriteLog("Réponse à : Disable X10 RF")
                            'Case ICMD.DISARC : WriteLog("Réponse à : Disable ARC RF")
                            'Case ICMD.DISAC : WriteLog("Réponse à : Disable AC RF")
                            'Case ICMD.DISHEU : WriteLog("Réponse à : Disable HomeEasy EU RF")
                            'Case ICMD.DISKOP : WriteLog("Réponse à : Disable Meiantech RF")
                            'Case ICMD.DISOREGON : WriteLog("Réponse à : Disable Oregon Scientific RF")
                            'Case ICMD.DISATI : WriteLog("Réponse à : Disable ATI remote RF")
                            'Case ICMD.DISVISONIC : WriteLog("Réponse à : Disable Visonic RF")
                            'Case ICMD.DISMERTIK : WriteLog("Réponse à : Disable Mertik RF")
                            'Case ICMD.DISAD : WriteLog("Réponse à : Disable AD RF")
                            'Case ICMD.DISHID : WriteLog("Réponse à : Disable Hideki RF")
                            'Case ICMD.DISLCROS : WriteLog("Réponse à : Disable La Crosse RF")
                            'Case ICMD.DISNOVAT : WriteLog("Réponse à : Disable Novatis RF")

                        Case IRESPONSE.sTypeUnknownRFYremote : WriteLog("ERR: Unknown RFY remote! Use the Program command to create a remote in the RFXtrx433E, Sequence nbr= " & recbuf(IRESPONSE.seqnbr).ToString)
                        Case IRESPONSE.sTypeExtError : WriteLog("ERR: No RFXtrx433E hardware detected")
                        Case IRESPONSE.sTypeRFYremoteList
                            If recbuf(IRESPONSE.msg2) = 0 And recbuf(IRESPONSE.msg3) = 0 And recbuf(IRESPONSE.msg4) = 0 And recbuf(IRESPONSE.msg5) = 0 Then
                                WriteLog("ERR: RFY remote:" & recbuf(IRESPONSE.msg1).ToString & " is empty")
                            Else
                                WriteLog("ERR: RFY remote:" & recbuf(IRESPONSE.msg1).ToString & " ID:" & VB.Right("0" & Hex(recbuf(IRESPONSE.msg2)), 2) & " " & VB.Right("0" & Hex(recbuf(IRESPONSE.msg3)), 2) & " " & VB.Right("0" & Hex(recbuf(IRESPONSE.msg4)), 2) & " unitnbr:" & Hex(recbuf(IRESPONSE.msg5)))
                            End If
                        Case IRESPONSE.sTypeInterfaceWrongCommand : WriteLog("ERR: Wrong command received from application")
                        Case Else : WriteLog("ERR: decode_InterfaceMessage : Données incorrectes reçues : type=" & Hex(recbuf(IRESPONSE.packettype)) & ", Sub type=" & Hex(recbuf(IRESPONSE.subtype)) & " cmnd=" & Hex(recbuf(IRESPONSE.cmnd)))
                    End Select
            End Select
            messagelog = Nothing
        Catch ex As Exception
            WriteLog("ERR: decode_InterfaceMessage : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_RecXmitMessage()
        Try
            Select Case recbuf(RXRESPONSE.subtype)
                Case RXRESPONSE.sTypeReceiverLockError
                    If _DEBUG Then WriteLog("ERR: Receiver lock error")
                Case RXRESPONSE.sTypeTransmitterResponse
                    Select Case recbuf(RXRESPONSE.msg)
                        Case &H0 : If _DEBUG Then WriteLog("Transmitter Response : ACK, data correct transmitted")
                        Case &H1 : If _DEBUG Then WriteLog("Transmitter Response : ACK, but transmit started after 6 seconds delay anyway with RF receive data detected")
                        Case &H2 : If _DEBUG Then WriteLog("ERR: Transmitter Response : NAK, transmitter did not lock on the requested transmit frequency")
                        Case &H3 : If _DEBUG Then WriteLog("ERR: Transmitter Response : NAK, AC address zero in id1-id4 not allowed")
                        Case Else : If _DEBUG Then WriteLog("ERR: decode_RecXmitMessage : Type de message reçu incorrect : type=" & Hex(recbuf(RXRESPONSE.msg)))
                    End Select
                Case Else : If _DEBUG Then WriteLog("ERR: decode_RecXmitMessage : Données incorrectes reçues : type=" & Hex(recbuf(RXRESPONSE.packettype)) & ", Sub type=" & Hex(recbuf(RXRESPONSE.subtype)))
            End Select
        Catch ex As Exception
            WriteLog("ERR: decode_RecXmitMessage : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_UNDECODED()
        Try
            Dim messagelog As String = ""
            messagelog = "UNDECODED "
            Select Case recbuf(UNDECODED.subtype)
                Case UNDECODED.sTypeUac : messagelog &= "AC:"
                Case UNDECODED.sTypeUarc : messagelog &= "ARC:"
                Case UNDECODED.sTypeUati : messagelog &= "ATI:"
                Case UNDECODED.sTypeUhideki : messagelog &= "HIDEKI:"
                Case UNDECODED.sTypeUlacrosse : messagelog &= "LACROSSE:"
                Case UNDECODED.sTypeUad : messagelog &= "AD:"
                Case UNDECODED.sTypeUmertik : messagelog &= "MERTIK:"
                Case UNDECODED.sTypeUoregon1 : messagelog &= "OREGON1:"
                Case UNDECODED.sTypeUoregon2 : messagelog &= "OREGON2:"
                Case UNDECODED.sTypeUoregon3 : messagelog &= "OREGON3:"
                Case UNDECODED.sTypeUproguard : messagelog &= "PROGUARD:"
                Case UNDECODED.sTypeUvisonic : messagelog &= "VISONIC:"
                Case UNDECODED.sTypeUnec : messagelog &= "NEC:"
                Case UNDECODED.sTypeUfs20 : messagelog &= "FS20:"
                Case UNDECODED.sTypeUrsl : messagelog &= "RSL:"
                Case UNDECODED.sTypeUblinds : messagelog &= "Blinds:"
                Case UNDECODED.sTypeUrubicson : messagelog &= "RUBICSON:"
                Case UNDECODED.sTypeUae : messagelog &= "AE:"
                Case UNDECODED.sTypeUfineoffset : messagelog &= "FineOffset:"
                Case UNDECODED.sTypeUrgb : messagelog &= "RGB:"
                Case UNDECODED.sTypeUrfy : messagelog &= "RFY:"
                Case Else : messagelog = "ERR: UNDECODED Unknown Sub type for Packet type=" & Hex(recbuf(UNDECODED.packettype)) & ": " & Hex(recbuf(UNDECODED.subtype))
            End Select
            For i = 0 To recbuf(UNDECODED.packetlength) - UNDECODED.msg1
                messagelog &= VB.Right("0" & Hex(recbuf(UNDECODED.msg1 + i)), 2)
            Next

            If recbuf(UNDECODED.subtype) = UNDECODED.sTypeUoregon3 And (recbuf(UNDECODED.msg1) = &H54 Or recbuf(UNDECODED.msg1) = &H8C) Then

                Dim amp1, amp2, amp3 As Single
                amp1 = CSng(((recbuf(UNDECODED.msg1 + 5) And &H7) * 16 + ((recbuf(UNDECODED.msg1 + 4) And &HF0) >> 4))) / 14
                amp2 = CSng(((recbuf(UNDECODED.msg1 + 7) And &H7) * 16 + ((recbuf(UNDECODED.msg1 + 6) And &HF0) >> 4))) / 14
                amp3 = CSng(((recbuf(UNDECODED.msg1 + 9) And &H7) * 16 + ((recbuf(UNDECODED.msg1 + 8) And &HF0) >> 4))) / 14
                messagelog &= (" amp1=" & amp1.ToString & " amp2=" & amp2.ToString & " amp3=" & amp3.ToString)

                If recbuf(UNDECODED.msg1) = &H8C Then
                    Dim kwh As Integer
                    kwh = recbuf(UNDECODED.msg1 + 10) >> 4
                    kwh = kwh + recbuf(UNDECODED.msg1 + 11) * 16
                    kwh = kwh + recbuf(UNDECODED.msg1 + 12) * 16 * 256
                    messagelog &= (" pwr=" & (kwh / 223.666).ToString)
                End If

                WriteLog(messagelog)
            End If
        Catch ex As Exception
            WriteLog("ERR: decode_UNDECODED Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Lighting1()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(LIGHTING1.subtype)
                Case LIGHTING1.sTypeX10
                    '        WriteMessage("subtype       = X10")
                    '        WriteMessage("Sequence nbr  = " & recbuf(LIGHTING1.seqnbr).ToString)
                    adresse = Chr(recbuf(LIGHTING1.housecode)) & recbuf(LIGHTING1.unitcode).ToString
                    Select Case recbuf(LIGHTING1.cmnd)
                        Case LIGHTING1.sOff : valeur = "OFF"
                        Case LIGHTING1.sOn : valeur = "ON"
                        Case LIGHTING1.sDim : valeur = "DIM"
                        Case LIGHTING1.sBright : valeur = "BRIGHT"
                        Case LIGHTING1.sAllOn : valeur = "ALL_ON"
                        Case LIGHTING1.sAllOff : valeur = "ALL_OFF"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING1.sTypeARC
                    '        WriteMessage("subtype       = ARC")
                    '        WriteMessage("Sequence nbr  = " & recbuf(LIGHTING1.seqnbr).ToString)
                    adresse = Chr(recbuf(LIGHTING1.housecode)) & recbuf(LIGHTING1.unitcode).ToString
                    Select Case recbuf(LIGHTING1.cmnd)
                        Case LIGHTING1.sOff : valeur = "OFF"
                        Case LIGHTING1.sOn : valeur = "ON"
                        Case LIGHTING1.sAllOn : valeur = "ALL_ON"
                        Case LIGHTING1.sAllOff : valeur = "ALL_OFF"
                        Case LIGHTING1.sChime : valeur = "CHIME"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING1.sTypeAB400D, LIGHTING1.sTypeWaveman, LIGHTING1.sTypeEMW200, LIGHTING1.sTypeIMPULS, LIGHTING1.sTypeRisingSun, LIGHTING1.sTypeGDR2, LIGHTING1.sTypeEnergenie, LIGHTING1.sTypeEnergenie5
                    'Select Case recbuf(LIGHTING1.subtype)
                    '    Case LIGHTING1.sTypeAB400D
                    '        WriteMessage("ELRO AB400")
                    '    Case LIGHTING1.sTypeWaveman
                    '        WriteMessage("Waveman")
                    '    Case LIGHTING1.sTypeEMW200
                    '        WriteMessage("EMW200")
                    '    Case LIGHTING1.sTypeIMPULS
                    '        WriteMessage("IMPULS")
                    '    Case LIGHTING1.sTypeRisingSun
                    '        WriteMessage("RisingSun")
                    '    Case LIGHTING1.sTypeEnergenie
                    '        WriteMessage("Energenie ENER010")
                    '    Case LIGHTING1.sTypeEnergenie5
                    '        WriteMessage("Energenie 5-gang")
                    '    Case LIGHTING1.sTypeGDR2
                    '        WriteMessage("COCO GDR2")
                    'End Select
                    adresse = Chr(recbuf(LIGHTING1.housecode)) & recbuf(LIGHTING1.unitcode).ToString
                    Select Case recbuf(LIGHTING1.cmnd)
                        Case LIGHTING1.sOff : valeur = "OFF"
                        Case LIGHTING1.sOn : valeur = "ON"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING1.sTypePhilips
                    '        WriteMessage("subtype       = Philips SBC")
                    '        WriteMessage("Sequence nbr  = " & recbuf(LIGHTING1.seqnbr).ToString)
                    adresse = Chr(recbuf(LIGHTING1.housecode)) & recbuf(LIGHTING1.unitcode).ToString
                    Select Case recbuf(LIGHTING1.cmnd)
                        Case LIGHTING1.sOff : valeur = "OFF"
                        Case LIGHTING1.sOn : valeur = "ON"
                        Case LIGHTING1.sAllOn : valeur = "ALL_ON"
                        Case LIGHTING1.sAllOff : valeur = "ALL_OFF"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case Else
                    WriteLog("ERR: decode_Lighting1 : Unknown Sub type for Packet type=" & Hex(recbuf(LIGHTING1.packettype)) & ": " & Hex(recbuf(LIGHTING1.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(LIGHTING1.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Lighting1 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Lighting2()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(LIGHTING2.subtype)
                Case LIGHTING2.sTypeAC, LIGHTING2.sTypeHEU, LIGHTING2.sTypeANSLUT
                    '        Select Case recbuf(LIGHTING2.subtype)
                    '            Case LIGHTING2.sTypeAC
                    '                WriteMessage("subtype       = AC")
                    '            Case LIGHTING2.sTypeHEU
                    '                WriteMessage("subtype       = HomeEasy EU")
                    '            Case LIGHTING2.sTypeANSLUT
                    '                WriteMessage("subtype       = ANSLUT")
                    '        End Select
                    '        WriteMessage("Sequence nbr  = " & recbuf(LIGHTING2.seqnbr).ToString)
                    adresse = Hex(recbuf(LIGHTING2.id1)) & VB.Right("0" & Hex(recbuf(LIGHTING2.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING2.id3)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING2.id4)), 2) & "-" & recbuf(LIGHTING2.unitcode).ToString
                    Select Case recbuf(LIGHTING2.cmnd)
                        Case LIGHTING2.sOff : valeur = "OFF"
                        Case LIGHTING2.sOn : valeur = "ON"
                        Case LIGHTING2.sSetLevel : valeur = "SET_LEVEL:" & recbuf(LIGHTING2.level).ToString
                        Case LIGHTING2.sGroupOff : valeur = "GROUP_OFF"
                        Case LIGHTING2.sGroupOn : valeur = "GROUP_ON"
                        Case LIGHTING2.sSetGroupLevel : valeur = "SET_GROUP_LEVEL:" & recbuf(LIGHTING2.level).ToString
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING2.sTypeKambrook
                    adresse = Chr(recbuf(LIGHTING2.id1) + &H41) & VB.Right("0" & Hex(recbuf(LIGHTING2.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING2.id3)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING2.id4)), 2) & "-" & recbuf(LIGHTING2.unitcode).ToString
                    Select Case recbuf(LIGHTING2.cmnd)
                        Case LIGHTING2.sOff : valeur = "OFF"
                        Case LIGHTING2.sOn : valeur = "ON"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                Case Else : WriteLog("ERR: decode_Lighting2 : Unknown Sub type for Packet type=" & Hex(recbuf(LIGHTING2.packettype)) & ": " & Hex(recbuf(LIGHTING2.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(LIGHTING2.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Lighting2 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Lighting3()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(LIGHTING3.subtype)
                Case LIGHTING3.sTypeKoppla
                    '        WriteMessage("subtype       = Ikea Koppla")
                    '        WriteMessage("Sequence nbr  = " & recbuf(LIGHTING3.seqnbr).ToString)
                    '        WriteMessage("Command       = "
                    adresse = "IKEAKOPPLA"
                    If recbuf(LIGHTING3.channel8_1) = 0 And recbuf(LIGHTING3.channel10_9) = 0 Then
                        WriteLog("ERR: decode_Lighting3: No channel selected")
                    Else
                        If (recbuf(LIGHTING3.channel8_1) And &H1) <> 0 Then adresse &= "1"
                        If (recbuf(LIGHTING3.channel8_1) And &H2) <> 0 Then adresse &= "2"
                        If (recbuf(LIGHTING3.channel8_1) And &H4) <> 0 Then adresse &= "3"
                        If (recbuf(LIGHTING3.channel8_1) And &H8) <> 0 Then adresse &= "4"
                        If (recbuf(LIGHTING3.channel8_1) And &H10) <> 0 Then adresse &= "5"
                        If (recbuf(LIGHTING3.channel8_1) And &H20) <> 0 Then adresse &= "6"
                        If (recbuf(LIGHTING3.channel8_1) And &H40) <> 0 Then adresse &= "7"
                        If (recbuf(LIGHTING3.channel8_1) And &H80) <> 0 Then adresse &= "8"
                        If (recbuf(LIGHTING3.channel10_9) And &H1) <> 0 Then adresse &= "9"
                        If (recbuf(LIGHTING3.channel10_9) And &H2) <> 0 Then adresse &= "10"
                    End If
                    Select Case recbuf(LIGHTING3.cmnd)
                        Case LIGHTING3.sBright : valeur = "BRIGHT"
                        Case LIGHTING3.sOff : valeur = "OFF"
                        Case LIGHTING3.sOn : valeur = "ON"
                        Case LIGHTING3.sDim : valeur = "DIM"
                        Case LIGHTING3.sLevel1 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:1")
                        Case LIGHTING3.sLevel2 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:2")
                        Case LIGHTING3.sLevel3 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:3")
                        Case LIGHTING3.sLevel4 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:4")
                        Case LIGHTING3.sLevel5 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:5")
                        Case LIGHTING3.sLevel6 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:6")
                        Case LIGHTING3.sLevel7 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:7")
                        Case LIGHTING3.sLevel8 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:8")
                        Case LIGHTING3.sLevel9 : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " SET_LEVEL:9")
                        Case LIGHTING3.sProgram : WriteLog("decode_Lighting3: Commande non gérée IkeaKoppla : " & adresse & " PROGRAM")
                        Case Else : WriteLog("ERR: decode_Lighting3 : Commande Inconnu IkeaKoppla : " & adresse)
                    End Select
                    If (valeur <> "") Then WriteRetour(adresse, "", valeur)
                Case Else : WriteLog("ERR: decode_Lighting3 : Unknown Sub type for Packet type=" & Hex(recbuf(LIGHTING3.packettype)) & ": " & Hex(recbuf(LIGHTING3.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(LIGHTING3.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Lighting3 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Lighting4()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(LIGHTING4.subtype)
                Case LIGHTING4.sTypePT2262
                    'WriteMessage("subtype       = PT2262")
                    'WriteMessage("Sequence nbr  = " & recbuf(LIGHTING4.seqnbr).ToString)
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING4.cmd1)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING4.cmd2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING4.cmd3)), 2)

                    If (recbuf(LIGHTING4.cmd1) And &H80) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd1) And &H40) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd1) And &H20) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd1) And &H10) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd1) And &H8) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd1) And &H4) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd1) And &H2) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd1) And &H1) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd2) And &H80) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd2) And &H40) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd2) And &H20) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd2) And &H10) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd2) And &H8) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd2) And &H4) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd2) And &H2) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd2) And &H1) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd3) And &H80) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd3) And &H40) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd3) And &H20) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd3) And &H10) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd3) And &H8) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd3) And &H4) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd3) And &H2) = 0 Then valeur &= "0" Else valeur &= "1"
                    If (recbuf(LIGHTING4.cmd3) And &H1) = 0 Then valeur &= "0" Else valeur &= "1"

                    'valeur &= " Pulse=" & ((recbuf(LIGHTING4.pulsehigh) * 256) + recbuf(LIGHTING4.pulselow)).ToString & " usec"

                    WriteLog("decode_Lighting4: Commande non gérée PT2262 : " & adresse & "=" & valeur)
                    'WriteRetour(adresse, "", valeur)
                Case Else : WriteLog("ERR: decode_Lighting4 : Unknown Sub type for Packet type=" & Hex(recbuf(LIGHTING4.packettype)) & ": " & Hex(recbuf(LIGHTING4.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(LIGHTING4.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Lighting4 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Lighting5()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(LIGHTING5.subtype)
                Case LIGHTING5.sTypeLightwaveRF
                    'WriteMessage("subtype       = LightwaveRF")
                    'WriteMessage("Sequence nbr  = " & recbuf(LIGHTING5.seqnbr).ToString)
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING5.id1)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id3)), 2) & "-" & recbuf(LIGHTING5.unitcode).ToString
                    Select Case recbuf(LIGHTING5.cmnd)
                        Case LIGHTING5.sOff : valeur = "OFF"
                        Case LIGHTING5.sOn : valeur = "ON"
                        Case LIGHTING5.sGroupOff : valeur = "GROUP_OFF"
                        Case LIGHTING5.sMood1 : valeur = "GROUP_Mood_1"
                        Case LIGHTING5.sMood2 : valeur = "GROUP_Mood_2"
                        Case LIGHTING5.sMood3 : valeur = "GROUP_Mood_3"
                        Case LIGHTING5.sMood4 : valeur = "GROUP_Mood_4"
                        Case LIGHTING5.sMood5 : valeur = "GROUP_Mood_5"
                        Case LIGHTING5.sUnlock : valeur = "UNLOCK"
                        Case LIGHTING5.sLock : valeur = "LOCK"
                        Case LIGHTING5.sAllLock : valeur = "ALL_LOCK"
                        Case LIGHTING5.sClose : valeur = "CLOSE_INLINE_RELAY"
                        Case LIGHTING5.sStop : valeur = "STOP_INLINE_RELAY"
                        Case LIGHTING5.sOpen : valeur = "OPEN_INLINE_RELAY"
                        Case LIGHTING5.sSetLevel : valeur = CInt((recbuf(LIGHTING5.level) * 3.2)).ToString 'Dim level "Set dim level to: " & CInt((recbuf(LIGHTING5.level) * 3.2)).ToString & "%"
                        Case LIGHTING5.sColourPalette : If recbuf(LIGHTING5.level) = 0 Then valeur = "Colour Palette (Even command)" Else valeur = "Colour Palette (Odd command)"
                        Case LIGHTING5.sColourTone : If recbuf(LIGHTING5.level) = 0 Then valeur = "Colour Tone (Even command)" Else valeur = "Colour Tone (Odd command)"
                        Case LIGHTING5.sColourCycle : If recbuf(LIGHTING5.level) = 0 Then valeur = "Colour Cycle (Even command)" Else valeur = "Colour Cycle (Odd command)"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING5.sTypeEMW100
                    'WriteMessage("subtype       = EMW100")
                    'WriteMessage("Sequence nbr  = " & recbuf(LIGHTING5.seqnbr).ToString)
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING5.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id3)), 2) & "-" & recbuf(LIGHTING5.unitcode).ToString
                    Select Case recbuf(LIGHTING5.cmnd)
                        Case LIGHTING5.sOff : valeur = "OFF"
                        Case LIGHTING5.sOn : valeur = "ON"
                        Case LIGHTING5.sLearn : valeur = "LEARN"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING5.sTypeBBSB
                    'WriteMessage("subtype       = BBSB new")
                    ' WriteMessage("Sequence nbr  = " & recbuf(LIGHTING5.seqnbr).ToString)
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING5.id1)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id3)), 2) & "-" & recbuf(LIGHTING5.unitcode).ToString
                    Select Case recbuf(LIGHTING5.cmnd)
                        Case LIGHTING5.sOff : valeur = "OFF"
                        Case LIGHTING5.sOn : valeur = "ON"
                        Case LIGHTING5.sGroupOff : valeur = "GROUP_OFF"
                        Case LIGHTING5.sGroupOn : valeur = "GROUP_ON"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING5.sTypeMDREMOTE
                    'WriteMessage("subtype       = MD Remote")
                    ' WriteMessage("Sequence nbr  = " & recbuf(LIGHTING5.seqnbr).ToString)
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING5.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id3)), 2)
                    Select Case recbuf(LIGHTING5.cmnd)
                        Case LIGHTING5.sPower : valeur = "Power"
                        Case LIGHTING5.sLight : valeur = "Light"
                        Case LIGHTING5.sBright : valeur = "Bright+"
                        Case LIGHTING5.sDim : valeur = "Bright-"
                        Case LIGHTING5.s100 : valeur = "100%"
                        Case LIGHTING5.s50 : valeur = "50%"
                        Case LIGHTING5.s25 : valeur = "25%"
                        Case LIGHTING5.sModePlus : valeur = "Mode+"
                        Case LIGHTING5.sSpeedMin : valeur = "Speed-"
                        Case LIGHTING5.sSpeedPlus : valeur = "Speed+"
                        Case LIGHTING5.sModeMin : valeur = "Mode-"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING5.sTypeRSL
                    'WriteMessage("subtype       = Conrad RSL")
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING5.id1)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id3)), 2) & "-" & recbuf(LIGHTING5.unitcode).ToString
                    Select Case recbuf(LIGHTING5.cmnd)
                        Case LIGHTING5.sOff : valeur = "OFF"
                        Case LIGHTING5.sOn : valeur = "ON"
                        Case LIGHTING5.sGroupOff : valeur = "GROUP_OFF"
                        Case LIGHTING5.sGroupOn : valeur = "GROUP_ON"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING5.sTypeTRC02, LIGHTING5.sTypeTRC02_2
                    'WriteMessage("subtype       = RGB TRC02")
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING5.id1)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id3)), 2)
                    Select Case recbuf(LIGHTING5.cmnd)
                        Case LIGHTING5.sRGBoff : valeur = "OFF"
                        Case LIGHTING5.sRGBon : valeur = "ON"
                        Case LIGHTING5.sRGBbright : valeur = "Bright+"
                        Case LIGHTING5.sRGBdim : valeur = "Bright-"
                        Case LIGHTING5.sRGBcolorplus : valeur = "Color+"
                        Case LIGHTING5.sRGBcolormin : valeur = "Color-"
                        Case Else : valeur = recbuf(LIGHTING5.cmnd).ToString
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING5.sTypeAoke
                    'WriteMessage("subtype       = Aoke relay")
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING5.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id3)), 2) & "-" & recbuf(LIGHTING5.unitcode).ToString
                    Select Case recbuf(LIGHTING5.cmnd)
                        Case LIGHTING5.sOff : valeur = "OFF"
                        Case LIGHTING5.sOn : valeur = "ON"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case LIGHTING5.sTypeEurodomest
                    'WriteMessage("subtype       =Eurodomest")
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING5.id1)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id2)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING5.id3)), 2) & "-" & recbuf(LIGHTING5.unitcode).ToString
                    Select Case recbuf(LIGHTING5.cmnd)
                        Case LIGHTING5.sOff : valeur = "OFF"
                        Case LIGHTING5.sOn : valeur = "ON"
                        Case LIGHTING5.sGroupOff : valeur = "GROUP_OFF"
                        Case LIGHTING5.sGroupOn : valeur = "GROUP_ON"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case Else : WriteLog("ERR: decode_Lighting5 : Unknown Sub type for Packet type=" & Hex(recbuf(LIGHTING5.packettype)) & ": " & Hex(recbuf(LIGHTING5.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(LIGHTING5.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Lighting5 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Lighting6()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(LIGHTING6.subtype)
                Case LIGHTING6.sTypeBlyss
                    'WriteMessage("subtype       = BLYSS")
                    'WriteMessage("Sequence nbr  = " & recbuf(LIGHTING6.seqnbr).ToString)
                    adresse = VB.Right("0" & Hex(recbuf(LIGHTING6.id1)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING6.id2)), 2) & "-" & Chr(recbuf(LIGHTING6.groupcode)) & recbuf(LIGHTING6.unitcode).ToString
                    'WriteMessage("ID            = " & VB.Right("0" & Hex(recbuf(LIGHTING6.id1)), 2) & VB.Right("0" & Hex(recbuf(LIGHTING6.id2)), 2))
                    'WriteMessage("groupcode     = " & Chr(recbuf(LIGHTING6.groupcode)))
                    'WriteMessage("unitcode      = " & recbuf(LIGHTING6.unitcode).ToString)
                    Select Case recbuf(LIGHTING6.cmnd)
                        Case LIGHTING6.sOff : valeur = "OFF"
                        Case LIGHTING6.sOn : valeur = "ON"
                        Case LIGHTING6.sGroupOff : valeur = "GROUP_OFF"
                        Case LIGHTING6.sGroupOn : valeur = "GROUP_ON"
                        Case Else : valeur = "UNKNOWN"
                    End Select

                    'sync the sequence numbers on received Blyss commands
                    If recbuf(LIGHTING6.cmndseqnbr) = 4 Then bytCmndSeqNbr = 0 Else bytCmndSeqNbr = recbuf(LIGHTING6.cmndseqnbr) + 1
                    If recbuf(LIGHTING6.seqnbr2) < 145 Then bytCmndSeqNbr2 = recbuf(LIGHTING6.seqnbr2) + 1 Else bytCmndSeqNbr2 = 1

                    WriteRetour(adresse, "", valeur)
                Case Else : WriteLog("ERR: decode_Lighting6 : Unknown Sub type for Packet type=" & Hex(recbuf(LIGHTING6.packettype)) & ": " & Hex(recbuf(LIGHTING6.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(LIGHTING6.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Lighting6 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Chime()
        'Byron MP001 receive is not implemented in the RFXtrx433E firmware
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(CHIME.subtype)
                Case CHIME.sTypeByronSX
                    'WriteMessage("subtype       = Byron SX")
                    'WriteMessage("ID            = " & VB.Right("0" & Hex(recbuf(CHIME.id1)), 2) & VB.Right("0" & Hex(recbuf(CHIME.id2)), 2))
                    adresse = VB.Right("0" & Hex(recbuf(CHIME.id1)), 2) & VB.Right("0" & Hex(recbuf(CHIME.id2)), 2)
                    Select Case recbuf(CHIME.sound)
                        Case CHIME.sSound0, CHIME.sSound4 : valeur = "Sound: Tubular 3 notes"
                        Case CHIME.sSound1, CHIME.sSound5 : valeur = "Sound: Big Ben"
                        Case CHIME.sSound2, CHIME.sSound6 : valeur = "Sound: Tubular 2 notes"
                        Case CHIME.sSound3, CHIME.sSound7 : valeur = "Sound: Solo"
                        Case Else : valeur = "Sound: Undefined:" & Hex(recbuf(CHIME.sound))
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case CHIME.sTypeByronMP001
                    'WriteMessage("subtype       = Byron MP001")
                    'WriteMessage("ID            = " & VB.Right("0" & Hex(recbuf(CHIME.id1)), 2) & VB.Right("0" & Hex(recbuf(CHIME.id2)), 2) & VB.Right("0" & Hex(recbuf(CHIME.sound)), 2))
                    adresse = VB.Right("0" & Hex(recbuf(CHIME.id1)), 2) & VB.Right("0" & Hex(recbuf(CHIME.id2)), 2) & VB.Right("0" & Hex(recbuf(CHIME.sound)), 2)
                    If recbuf(CHIME.id1) And &H40 = &H40 Then valeur = "Switch 1=Off" Else valeur = "Switch 1=On"
                    If recbuf(CHIME.id1) And &H10 = &H10 Then valeur = "Switch 2=Off" Else valeur = "Switch 2=On"
                    If recbuf(CHIME.id1) And &H4 = &H4 Then valeur = "Switch 3=Off" Else valeur = "Switch 3=On"
                    If recbuf(CHIME.id1) And &H1 = &H1 Then valeur = "Switch 4=Off" Else valeur = "Switch 4=On"
                    If recbuf(CHIME.id2) And &H40 = &H40 Then valeur = "Switch 5=Off" Else valeur = "Switch 5=On"
                    If recbuf(CHIME.id2) And &H10 = &H10 Then valeur = "Switch 6=Off" Else valeur = "Switch 6=On"
                    WriteRetour(adresse, "", valeur)
                Case CHIME.sTypeSelectPlus
                    'WriteMessage("subtype       = SelectPlus")
                    'WriteMessage("ID            = " & VB.Right("0" & Hex(recbuf(CHIME.id1)), 2) & VB.Right("0" & Hex(recbuf(CHIME.id2)), 2) & VB.Right("0" & Hex(recbuf(CHIME.sound)), 2))
                    adresse = VB.Right("0" & Hex(recbuf(CHIME.id1)), 2) & VB.Right("0" & Hex(recbuf(CHIME.id2)), 2) & VB.Right("0" & Hex(recbuf(CHIME.sound)), 2)
                    valeur = "CHIME"
                Case CHIME.sTypeEnvivo
                    'WriteMessage("subtype       = Envivo ENV-1348")
                    'WriteMessage("ID            = " & VB.Right("0" & Hex(recbuf(CHIME.id1)), 2) & VB.Right("0" & Hex(recbuf(CHIME.id2)), 2))
                    adresse = VB.Right("0" & Hex(recbuf(CHIME.id1)), 2) & VB.Right("0" & Hex(recbuf(CHIME.id2)), 2)
                    valeur = "CHIME"
                Case Else : WriteLog("ERR: decode_Chime : Unknown Sub type for Packet type=" & Hex(recbuf(CHIME.packettype)) & ": " & Hex(recbuf(CHIME.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(CHIME.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Chime Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Curtain1()
        Try
            'decoding of this type is only implemented for use by simulate and verbose
            'Curtain1 receive is not implemented in the RFXtrx433 firmware
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(CURTAIN1.subtype)
                Case CURTAIN1.sTypeHarrison
                    'WriteMessage("subtype       = Harrison")
                    'WriteMessage("Sequence nbr  = " & recbuf(CURTAIN1.seqnbr).ToString)
                    adresse = Chr(recbuf(CURTAIN1.housecode)) & recbuf(CURTAIN1.unitcode).ToString
                    Select Case recbuf(CURTAIN1.cmnd)
                        Case CURTAIN1.sOpen : valeur = "OPEN"
                        Case CURTAIN1.sClose : valeur = "CLOSE"
                        Case CURTAIN1.sStop : valeur = "STOP"
                        Case CURTAIN1.sProgram : valeur = "PROGRAM"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case Else : WriteLog("ERR: decode_Curtain1 : Unknown Sub type for Packet type=" & Hex(recbuf(CURTAIN1.packettype)) & ": " & Hex(recbuf(CURTAIN1.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(CURTAIN1.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Curtain1 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Security1()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            'Select Case recbuf(SECURITY1.subtype)
            '    Case SECURITY1.sTypeSecX10
            '        WriteMessage("subtype       = X10 security")
            '    Case SECURITY1.sTypeSecX10M
            '        WriteMessage("subtype       = X10 security motion")
            '    Case SECURITY1.sTypeSecX10R
            '        WriteMessage("subtype       = X10 security remote")
            '    Case SECURITY1.sTypeKD101
            '        WriteMessage("subtype       = KD101 smoke detector")
            '    Case SECURITY1.sTypePowercodeSensor
            '        WriteMessage("subtype       = Visonic PowerCode sensor - primary contact")
            '    Case SECURITY1.sTypePowercodeMotion
            '        WriteMessage("subtype       = Visonic PowerCode motion")
            '    Case SECURITY1.sTypeCodesecure
            '        WriteMessage("subtype       = Visonic CodeSecure")
            '    Case SECURITY1.sTypePowercodeAux
            '        WriteMessage("subtype       = Visonic PowerCode sensor - auxiliary contact")
            '    Case SECURITY1.sTypeMeiantech
            '        WriteMessage("subtype       = Meiantech/Atlantic/Aidebao")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(SECURITY1.packettype)) & ": " & Hex(recbuf(SECURITY1.subtype)))
            'End Select

            adresse = VB.Right("0" & Hex(recbuf(SECURITY1.id1)), 2) & VB.Right("0" & Hex(recbuf(SECURITY1.id2)), 2) & VB.Right("0" & Hex(recbuf(SECURITY1.id3)), 2)
            Select Case recbuf(SECURITY1.status)
                Case SECURITY1.sStatusNormal : valeur = "NORMAL"
                Case SECURITY1.sStatusNormalDelayed : valeur = "NORMAL DELAYED"
                Case SECURITY1.sStatusAlarm : valeur = "ALARM"
                Case SECURITY1.sStatusAlarmDelayed : valeur = "ALARM DELAYED"
                Case SECURITY1.sStatusMotion : valeur = "MOTION"
                Case SECURITY1.sStatusNoMotion : valeur = "NO MOTION"
                Case SECURITY1.sStatusPanic : valeur = "PANIC"
                Case SECURITY1.sStatusPanicOff : valeur = "PANIC END"
                Case SECURITY1.sStatusIRbeam : valeur = "IR BEAM BLOCKED"
                Case SECURITY1.sStatusArmAway
                    valeur = "ARM AWAY"
                    If recbuf(SECURITY1.subtype) = SECURITY1.sTypeMeiantech Then valeur = "Group2 OR Arm Away"
                Case SECURITY1.sStatusArmAwayDelayed : valeur = "ARM AWAY DELAYED"
                Case SECURITY1.sStatusArmHome
                    valeur = "ARM HOME"
                    If recbuf(SECURITY1.subtype) = SECURITY1.sTypeMeiantech Then valeur = "Group3 OR ARM HOME"
                Case SECURITY1.sStatusArmHomeDelayed : valeur = "ARM HOME DELAYED"
                Case SECURITY1.sStatusDisarm
                    valeur = "DISARM"
                    If recbuf(SECURITY1.subtype) = SECURITY1.sTypeMeiantech Then valeur = "Group1 OR DISARM"
                Case SECURITY1.sStatusLightOff : valeur = "LIGHT OFF"
                Case SECURITY1.sStatusLightOn : valeur = "LIGHT ON"
                Case SECURITY1.sStatusLIGHTING2Off : valeur = "LIGHT 2 OFF"
                Case SECURITY1.sStatusLIGHTING2On : valeur = "LIGHT 2 ON"
                Case SECURITY1.sStatusDark : valeur = "DARK DETECTED"
                Case SECURITY1.sStatusLight : valeur = "LIGHT DETECTED"
                Case SECURITY1.sStatusBatLow : valeur = "BATTERY LOW MS10 OR XX18 SENSOR"
                Case SECURITY1.sStatusPairKD101 : valeur = "PAIR KD101"
                Case SECURITY1.sStatusNormalTamper : valeur = "NORMAL + TAMPER"
                Case SECURITY1.sStatusNormalDelayedTamper : valeur = "NORMAL DELAYED + TAMPER"
                Case SECURITY1.sStatusAlarmTamper : valeur = "ALARM + TAMPER"
                Case SECURITY1.sStatusAlarmDelayedTamper : valeur = "ALARM DELAYED + TAMPER"
                Case SECURITY1.sStatusMotionTamper : valeur = "MOTION + TAMPER"
                Case SECURITY1.sStatusNoMotionTamper : valeur = "NO MOTION + TAMPER"
            End Select
            WriteRetour(adresse, "", valeur)
            If recbuf(SECURITY1.subtype) <> SECURITY1.sTypeKD101 Then    'KD101 does not support battery low indication
                If (recbuf(SECURITY1.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            End If
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(SECURITY1.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Security1 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Camera1()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(CAMERA1.subtype)
                Case CAMERA1.sTypeNinja 'X10 Ninja/Robocam
                    Select Case recbuf(CAMERA1.cmnd)
                        Case CAMERA1.sLeft : valeur = "LEFT"
                        Case CAMERA1.sRight : valeur = "RIGHT"
                        Case CAMERA1.sUp : valeur = "UP"
                        Case CAMERA1.sDown : valeur = "DOWN"
                        Case CAMERA1.sPosition1 : valeur = "POSITION 1"
                        Case CAMERA1.sProgramPosition1 : valeur = "POSITION 1 PROGRAM"
                        Case CAMERA1.sPosition2 : valeur = "POSITION 2"
                        Case CAMERA1.sProgramPosition2 : valeur = "POSITION 2 PROGRAM"
                        Case CAMERA1.sPosition3 : valeur = "POSITION 3"
                        Case CAMERA1.sProgramPosition3 : valeur = "POSITION 3 PROGRAM"
                        Case CAMERA1.sPosition4 : valeur = "POSITION 4"
                        Case CAMERA1.sProgramPosition4 : valeur = "POSITION 4 PROGRAM"
                        Case CAMERA1.sCenter : valeur = "CENTER"
                        Case CAMERA1.sProgramCenterPosition : valeur = "CENTER PROGRAM"
                        Case CAMERA1.sSweep : valeur = "SWEEP"
                        Case CAMERA1.sProgramSweep : valeur = "SWEEP PROGRAM"
                        Case Else : valeur = "UNKNOWN"
                    End Select
                    adresse = Chr(recbuf(CAMERA1.housecode))
                    WriteRetour(adresse, "", valeur)
                    If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(CAMERA1.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
                Case Else : WriteLog("ERR: decode_Camera1 : Unknown Sub type for Packet type=" & Hex(recbuf(CAMERA1.packettype)) & ": " & Hex(recbuf(CAMERA1.subtype)))
            End Select
        Catch ex As Exception
            WriteLog("ERR: decode_Camera1 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_BLINDS1()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            'Select Case recbuf(BLINDS1.subtype)
            '    Case BLINDS1.BlindsT0
            '        WriteMessage("subtype       = Safy / RollerTrol / Hasta new")
            '    Case BLINDS1.BlindsT1
            '        WriteMessage("subtype       = Hasta old")
            '    Case BLINDS1.BlindsT2
            '        WriteMessage("subtype       = A-OK RF01")
            '    Case BLINDS1.BlindsT3
            '        WriteMessage("subtype       = A-OK AC114")
            '    Case BLINDS1.BlindsT4
            '        WriteMessage("subtype       = RAEX")
            '    Case BLINDS1.BlindsT5
            '        WriteMessage("subtype       = Media Mount")
            '    Case BLINDS1.BlindsT6
            '        WriteMessage("subtype       = DC106")
            '    Case BLINDS1.BlindsT7
            '        WriteMessage("subtype       = Forest")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(BLINDS1.packettype)) & ": " & Hex(recbuf(BLINDS1.subtype)))
            'End Select
            'WriteMessage("Sequence nbr  = " & recbuf(BLINDS1.seqnbr).ToString)
            Select Case recbuf(BLINDS1.subtype)
                Case BLINDS1.BlindsT0, BLINDS1.BlindsT1
                    adresse = VB.Right("0" & Hex(recbuf(BLINDS1.id2)), 2) & VB.Right("0" & Hex(recbuf(BLINDS1.id3)), 2)
                    If recbuf(BLINDS1.unitcode) = 0 Then adresse &= "-ALL" Else adresse &= "-" & recbuf(BLINDS1.unitcode).ToString
                Case BLINDS1.BlindsT6, BLINDS1.BlindsT7
                    adresse = VB.Right("0" & Hex(recbuf(BLINDS1.id1)), 2) & VB.Right("0" & Hex(recbuf(BLINDS1.id2)), 2) & VB.Right("0" & Hex(recbuf(BLINDS1.id3)), 2) & Hex(recbuf(BLINDS1.id4) >> 4)
                    If recbuf(BLINDS1.unitcode) = 0 Then adresse &= "-ALL" Else adresse &= "-" & (recbuf(BLINDS1.unitcode) And &HF).ToString
                Case Else
                    adresse = VB.Right("0" & Hex(recbuf(BLINDS1.id1)), 2) & VB.Right("0" & Hex(recbuf(BLINDS1.id2)), 2) & VB.Right("0" & Hex(recbuf(BLINDS1.id3)), 2)
            End Select

            Select Case recbuf(BLINDS1.cmnd)
                Case BLINDS1.sOpen : valeur = "OPEN"
                Case BLINDS1.sStop : valeur = "STOP"
                Case BLINDS1.sClose : valeur = "CLOSE"
                Case BLINDS1.sConfirm : valeur = "CONFIRM"
                Case BLINDS1.sLimit
                    If recbuf(BLINDS1.subtype) = BLINDS1.BlindsT4 Then
                        valeur = "SET UPPER LIMIT"
                    Else
                        valeur = "SET LIMIT"
                    End If
                Case BLINDS1.sLowerLimit : valeur = "SET LOWER LIMIT"
                Case BLINDS1.sDeleteLimits : valeur = "DELETE LIMITS"
                Case BLINDS1.sChangeDirection : valeur = "CHANGE DIRECTION"
                Case BLINDS1.sLeft : valeur = "LEFT"
                Case BLINDS1.sRight : valeur = "RIGHT"
                Case Else : valeur = "UNKNOWN"
            End Select
            WriteRetour(adresse, "", valeur)
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(BLINDS1.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_BLINDS1 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_RFY()
        'decoding of this type is only implemented for use by simulate and verbose
        'RFY receive is not implemented in the RFXtrx433E firmware
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            'Select Case recbuf(RFY.subtype)
            '    Case RFY.RFY
            '        WriteMessage("subtype       = RFY")
            '    Case RFY.RFYext
            '        WriteMessage("subtype       = not used!")
            '    Case RFY.GEOM
            '        WriteMessage("subtype       = GEOM")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(RFY.packettype)) & ": " & Hex(recbuf(RFY.subtype)))
            'End Select

            Select Case recbuf(RFY.subtype)
                Case RFY.RFY
                    adresse = VB.Right("0" & Hex(recbuf(RFY.id1)), 2) & VB.Right("0" & Hex(recbuf(RFY.id2)), 2) & VB.Right("0" & Hex(recbuf(RFY.id3)), 2)
                    If recbuf(RFY.unitcode) = 0 Then adresse &= "-ALL" Else adresse &= "-" & (recbuf(RFY.unitcode) And &HF).ToString
            End Select

            Select Case recbuf(RFY.cmnd)
                Case RFY.sStop : valeur = "stop"
                Case RFY.sUp : valeur = "up"
                Case RFY.sUpStop : valeur = "up + stop"
                Case RFY.sDown : valeur = "down"
                Case RFY.sDownStop : valeur = "down + stop"
                Case RFY.sUpDown : valeur = "up + down"
                Case RFY.sListRemotes : valeur = "List remotes"
                Case RFY.sProgram : valeur = "program"
                Case RFY.s2SecProgram : valeur = "> 2 seconds: program"
                Case RFY.s7SecProgram : valeur = "> 7 seconds: program"
                Case RFY.s2SecStop : valeur = "> 2 second: stop"
                Case RFY.s5SecStop : valeur = "> 5 seconds: stop"
                Case RFY.s5SecUpDown : valeur = "> 5 seconds: up + down"
                Case RFY.sEraseThis : valeur = "Erase this remote"
                Case RFY.sEraseAll : valeur = "Erase all remotes"
                Case RFY.s05SecUP : valeur = "< 0.5 seconds: up"
                Case RFY.s05SecDown : valeur = "< 0.5 seconds: down"
                Case RFY.s2SecUP : valeur = "> 2 seconds: up"
                Case RFY.s2SecDown : valeur = "> 2 seconds: down"
                Case Else : valeur = "UNKNOWN"
            End Select

            'WriteMessage("rfu1          = " & VB.Right("0" & Hex(recbuf(RFY.rfu1)), 2))
            'WriteMessage("rfu2          = " & VB.Right("0" & Hex(recbuf(RFY.rfu2)), 2))
            'WriteMessage("rfu3          = " & VB.Right("0" & Hex(recbuf(RFY.rfu3)), 2))

            WriteRetour(adresse, "", valeur)
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(RFY.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_RFY Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Remote()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""

            Select Case recbuf(REMOTE.subtype)
                Case REMOTE.sTypeATI 'ATI Remote Wonder
                    adresse = recbuf(REMOTE.id).ToString
                    Select Case recbuf(REMOTE.cmnd)
                        Case &H0 : valeur = "A"
                        Case &H1 : valeur = "B"
                        Case &H2 : valeur = "power"
                        Case &H3 : valeur = "TV"
                        Case &H4 : valeur = "DVD"
                        Case &H5 : valeur = "?"
                        Case &H6 : valeur = "Guide"
                        Case &H7 : valeur = "Drag"
                        Case &H8 : valeur = "VOL+"
                        Case &H9 : valeur = "VOL-"
                        Case &HA : valeur = "MUTE"
                        Case &HB : valeur = "CHAN+"
                        Case &HC : valeur = "CHAN-"
                        Case &HD : valeur = "1"
                        Case &HE : valeur = "2"
                        Case &HF : valeur = "3"
                        Case &H10 : valeur = "4"
                        Case &H11 : valeur = "5"
                        Case &H12 : valeur = "6"
                        Case &H13 : valeur = "7"
                        Case &H14 : valeur = "8"
                        Case &H15 : valeur = "9"
                        Case &H16 : valeur = "txt"
                        Case &H17 : valeur = "0"
                        Case &H18 : valeur = "snapshot ESC"
                        Case &H19 : valeur = "C"
                        Case &H1A : valeur = "^"
                        Case &H1B : valeur = "D"
                        Case &H1C : valeur = "TV/RADIO"
                        Case &H1D : valeur = "<"
                        Case &H1E : valeur = "OK"
                        Case &H1F : valeur = ">"
                        Case &H20 : valeur = "<-"
                        Case &H21 : valeur = "E"
                        Case &H22 : valeur = "v"
                        Case &H23 : valeur = "F"
                        Case &H24 : valeur = "Rewind"
                        Case &H25 : valeur = "Play"
                        Case &H26 : valeur = "Fast forward"
                        Case &H27 : valeur = "Record"
                        Case &H28 : valeur = "Stop"
                        Case &H29 : valeur = "Pause"
                        Case &H2C : valeur = "TV"
                        Case &H2D : valeur = "VCR"
                        Case &H2E : valeur = "RADIO"
                        Case &H2F : valeur = "TV Preview"
                        Case &H30 : valeur = "Channel list"
                        Case &H31 : valeur = "Video Desktop"
                        Case &H32 : valeur = "red"
                        Case &H33 : valeur = "green"
                        Case &H34 : valeur = "yellow"
                        Case &H35 : valeur = "blue"
                        Case &H36 : valeur = "rename TAB"
                        Case &H37 : valeur = "Acquire image"
                        Case &H38 : valeur = "edit image"
                        Case &H39 : valeur = "Full screen"
                        Case &H3A : valeur = "DVD Audio"
                        Case &H70 : valeur = "Cursor-left"
                        Case &H71 : valeur = "Cursor-right"
                        Case &H72 : valeur = "Cursor-up"
                        Case &H73 : valeur = "Cursor-down"
                        Case &H74 : valeur = "Cursor-up-left"
                        Case &H75 : valeur = "Cursor-up-right"
                        Case &H76 : valeur = "Cursor-down-right"
                        Case &H77 : valeur = "Cursor-down-left"
                        Case &H78 : valeur = "V"
                        Case &H79 : valeur = "V-End"
                        Case &H7C : valeur = "X"
                        Case &H7D : valeur = "X-End"
                        Case Else : valeur = "unknown"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case REMOTE.sTypeATIplus 'ATI Remote Wonder Plus
                    adresse = recbuf(REMOTE.id).ToString
                    Select Case recbuf(REMOTE.cmnd)
                        Case &H0 : valeur = "A"
                        Case &H1 : valeur = "B"
                        Case &H2 : valeur = "power"
                        Case &H3 : valeur = "TV"
                        Case &H4 : valeur = "DVD"
                        Case &H5 : valeur = "?"
                        Case &H6 : valeur = "Guide"
                        Case &H7 : valeur = "Drag"
                        Case &H8 : valeur = "VOL+"
                        Case &H9 : valeur = "VOL-"
                        Case &HA : valeur = "MUTE"
                        Case &HB : valeur = "CHAN+"
                        Case &HC : valeur = "CHAN-"
                        Case &HD : valeur = "1"
                        Case &HE : valeur = "2"
                        Case &HF : valeur = "3"
                        Case &H10 : valeur = "4"
                        Case &H11 : valeur = "5"
                        Case &H12 : valeur = "6"
                        Case &H13 : valeur = "7"
                        Case &H14 : valeur = "8"
                        Case &H15 : valeur = "9"
                        Case &H16 : valeur = "txt"
                        Case &H17 : valeur = "0"
                        Case &H18 : valeur = "Open Setup Menu"
                        Case &H19 : valeur = "C"
                        Case &H1A : valeur = "^"
                        Case &H1B : valeur = "D"
                        Case &H1C : valeur = "FM"
                        Case &H1D : valeur = "<"
                        Case &H1E : valeur = "OK"
                        Case &H1F : valeur = ">"
                        Case &H20 : valeur = "Max/Restore window"
                        Case &H21 : valeur = "E"
                        Case &H22 : valeur = "v"
                        Case &H23 : valeur = "F"
                        Case &H24 : valeur = "Rewind"
                        Case &H25 : valeur = "Play"
                        Case &H26 : valeur = "Fast forward"
                        Case &H27 : valeur = "Record"
                        Case &H28 : valeur = "Stop"
                        Case &H29 : valeur = "Pause"
                        Case &H2A : valeur = "TV2"
                        Case &H2B : valeur = "Clock"
                        Case &H2C : valeur = "i"
                        Case &H2D : valeur = "ATI"
                        Case &H2E : valeur = "RADIO"
                        Case &H2F : valeur = "TV Preview"
                        Case &H30 : valeur = "Channel list"
                        Case &H31 : valeur = "Video Desktop"
                        Case &H32 : valeur = "red"
                        Case &H33 : valeur = "green"
                        Case &H34 : valeur = "yellow"
                        Case &H35 : valeur = "blue"
                        Case &H36 : valeur = "rename TAB"
                        Case &H37 : valeur = "Acquire image"
                        Case &H38 : valeur = "edit image"
                        Case &H39 : valeur = "Full screen"
                        Case &H3A : valeur = "DVD Audio"
                        Case &H70 : valeur = "Cursor-left"
                        Case &H71 : valeur = "Cursor-right"
                        Case &H72 : valeur = "Cursor-up"
                        Case &H73 : valeur = "Cursor-down"
                        Case &H74 : valeur = "Cursor-up-left"
                        Case &H75 : valeur = "Cursor-up-right"
                        Case &H76 : valeur = "Cursor-down-right"
                        Case &H77 : valeur = "Cursor-down-left"
                        Case &H78 : valeur = "Left Mouse Button"
                        Case &H79 : valeur = "V-End"
                        Case &H7C : valeur = "Right Mouse Button"
                        Case &H7D : valeur = "X-End"
                        Case Else : valeur = "unknown"
                    End Select
                    '        If (recbuf(REMOTE.toggle) And &H1) = &H1 Then
                    '            WriteMessage("  (button press = odd)")
                    '        Else
                    '            WriteMessage("  (button press = even)")
                    '        End If
                    WriteRetour(adresse, "", valeur)
                Case REMOTE.sTypeATIrw2 'ATI Remote Wonder II
                    adresse = recbuf(REMOTE.id).ToString
                    'Select Case recbuf(REMOTE.cmndtype) And &HE
                    '    Case &H0 : WriteMessage("PC")
                    '    Case &H2 : WriteMessage("AUX1")
                    '    Case &H4 : WriteMessage("AUX2")
                    '    Case &H6 : WriteMessage("AUX3")
                    '    Case &H8 : WriteMessage("AUX4")
                    '    Case Else : WriteMessage("unknown")
                    'End Select
                    Select Case recbuf(REMOTE.cmnd)
                        Case &H0 : valeur = "A"
                        Case &H1 : valeur = "B"
                        Case &H2 : valeur = "power"
                        Case &H3 : valeur = "TV"
                        Case &H4 : valeur = "DVD"
                        Case &H5 : valeur = "?"
                        Case &H6 : valeur = "Guide"
                        Case &H7 : valeur = "Drag"
                        Case &H8 : valeur = "VOL+"
                        Case &H9 : valeur = "VOL-"
                        Case &HA : valeur = "MUTE"
                        Case &HB : valeur = "CHAN+"
                        Case &HC : valeur = "CHAN-"
                        Case &HD : valeur = "1"
                        Case &HE : valeur = "2"
                        Case &HF : valeur = "3"
                        Case &H10 : valeur = "4"
                        Case &H11 : valeur = "5"
                        Case &H12 : valeur = "6"
                        Case &H13 : valeur = "7"
                        Case &H14 : valeur = "8"
                        Case &H15 : valeur = "9"
                        Case &H16 : valeur = "txt"
                        Case &H17 : valeur = "0"
                        Case &H18 : valeur = "Open Setup Menu"
                        Case &H19 : valeur = "C"
                        Case &H1A : valeur = "^"
                        Case &H1B : valeur = "D"
                        Case &H1C : valeur = "FM"
                        Case &H1D : valeur = "<"
                        Case &H1E : valeur = "OK"
                        Case &H1F : valeur = ">"
                        Case &H20 : valeur = "Max/Restore window"
                        Case &H21 : valeur = "E"
                        Case &H22 : valeur = "v"
                        Case &H23 : valeur = "F"
                        Case &H24 : valeur = "Rewind"
                        Case &H25 : valeur = "Play"
                        Case &H26 : valeur = "Fast forward"
                        Case &H27 : valeur = "Record"
                        Case &H28 : valeur = "Stop"
                        Case &H29 : valeur = "Pause"
                        Case &H2C : valeur = "i"
                        Case &H2D : valeur = "ATI"
                        Case &H3B : valeur = "PC"
                        Case &H3C : valeur = "AUX1"
                        Case &H3D : valeur = "AUX2"
                        Case &H3E : valeur = "AUX3"
                        Case &H3F : valeur = "AUX4"
                        Case &H70 : valeur = "Cursor-left"
                        Case &H71 : valeur = "Cursor-right"
                        Case &H72 : valeur = "Cursor-up"
                        Case &H73 : valeur = "Cursor-down"
                        Case &H74 : valeur = "Cursor-up-left"
                        Case &H75 : valeur = "Cursor-up-right"
                        Case &H76 : valeur = "Cursor-down-right"
                        Case &H77 : valeur = "Cursor-down-left"
                        Case &H78 : valeur = "Left Mouse Button"
                        Case &H7C : valeur = "Right Mouse Button"
                        Case Else : valeur = "unknown"
                    End Select
                    '        If (recbuf(REMOTE.toggle) And &H1) = &H1 Then
                    '            WriteMessage("  (button press = odd)")
                    '        Else
                    '            WriteMessage("  (button press = even)")
                    '        End If
                    WriteRetour(adresse, "", valeur)
                Case REMOTE.sTypeMedion 'Medion Remote
                    adresse = recbuf(REMOTE.id).ToString
                    Select Case recbuf(REMOTE.cmnd)
                        Case &H0 : valeur = "Mute"
                        Case &H1 : valeur = "B"
                        Case &H2 : valeur = "power"
                        Case &H3 : valeur = "TV"
                        Case &H4 : valeur = "DVD"
                        Case &H5 : valeur = "Photo"
                        Case &H6 : valeur = "Music"
                        Case &H7 : valeur = "Drag"
                        Case &H8 : valeur = "VOL-"
                        Case &H9 : valeur = "VOL+"
                        Case &HA : valeur = "MUTE"
                        Case &HB : valeur = "CHAN+"
                        Case &HC : valeur = "CHAN-"
                        Case &HD : valeur = "1"
                        Case &HE : valeur = "2"
                        Case &HF : valeur = "3"
                        Case &H10 : valeur = "4"
                        Case &H11 : valeur = "5"
                        Case &H12 : valeur = "6"
                        Case &H13 : valeur = "7"
                        Case &H14 : valeur = "8"
                        Case &H15 : valeur = "9"
                        Case &H16 : valeur = "txt"
                        Case &H17 : valeur = "0"
                        Case &H18 : valeur = "snapshot ESC"
                        Case &H19 : valeur = "DVD MENU"
                        Case &H1A : valeur = "^"
                        Case &H1B : valeur = "Setup"
                        Case &H1C : valeur = "TV/RADIO"
                        Case &H1D : valeur = "<"
                        Case &H1E : valeur = "OK"
                        Case &H1F : valeur = ">"
                        Case &H20 : valeur = "<-"
                        Case &H21 : valeur = "E"
                        Case &H22 : valeur = "v"
                        Case &H23 : valeur = "F"
                        Case &H24 : valeur = "Rewind"
                        Case &H25 : valeur = "Play"
                        Case &H26 : valeur = "Fast forward"
                        Case &H27 : valeur = "Record"
                        Case &H28 : valeur = "Stop"
                        Case &H29 : valeur = "Pause"
                        Case &H2C : valeur = "V"
                        Case &H2D : valeur = "VCR"
                        Case &H2E : valeur = "RADIO"
                        Case &H2F : valeur = "TV Preview"
                        Case &H30 : valeur = "Channel list"
                        Case &H31 : valeur = "Video Desktop"
                        Case &H32 : valeur = "red"
                        Case &H33 : valeur = "green"
                        Case &H34 : valeur = "yellow"
                        Case &H35 : valeur = "blue"
                        Case &H36 : valeur = "rename TAB"
                        Case &H37 : valeur = "Acquire image"
                        Case &H38 : valeur = "edit image"
                        Case &H39 : valeur = "Full screen"
                        Case &H3A : valeur = "DVD Audio"
                        Case &H70 : valeur = "Cursor-left"
                        Case &H71 : valeur = "Cursor-right"
                        Case &H72 : valeur = "Cursor-up"
                        Case &H73 : valeur = "Cursor-down"
                        Case &H74 : valeur = "Cursor-up-left"
                        Case &H75 : valeur = "Cursor-up-right"
                        Case &H76 : valeur = "Cursor-down-right"
                        Case &H77 : valeur = "Cursor-down-left"
                        Case &H78 : valeur = "V"
                        Case &H79 : valeur = "V-End"
                        Case &H7C : valeur = "X"
                        Case &H7D : valeur = "X-End"
                        Case Else : valeur = "unknown"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case REMOTE.sTypePCremote 'PC Remote
                    adresse = recbuf(REMOTE.id).ToString
                    Select Case recbuf(REMOTE.cmnd)
                        Case &H2 : valeur = "0"
                        Case &H82 : valeur = "1"
                        Case &HD1 : valeur = "MP3"
                        Case &H42 : valeur = "2"
                        Case &HD2 : valeur = "DVD"
                        Case &HC2 : valeur = "3"
                        Case &HD3 : valeur = "CD"
                        Case &H22 : valeur = "4"
                        Case &HD4 : valeur = "PC or SHIFT-4"
                        Case &HA2 : valeur = "5"
                        Case &HD5 : valeur = "SHIFT-5"
                        Case &H62 : valeur = "6"
                        Case &HE2 : valeur = "7"
                        Case &H12 : valeur = "8"
                        Case &H92 : valeur = "9"
                        Case &HC0 : valeur = "CH-"
                        Case &H40 : valeur = "CH+"
                        Case &HE0 : valeur = "VOL-"
                        Case &H60 : valeur = "VOL+"
                        Case &HA0 : valeur = "MUTE"
                        Case &H3A : valeur = "INFO"
                        Case &H38 : valeur = "REW"
                        Case &HB8 : valeur = "FF"
                        Case &HB0 : valeur = "PLAY"
                        Case &H64 : valeur = "PAUSE"
                        Case &H63 : valeur = "STOP"
                        Case &HB6 : valeur = "MENU"
                        Case &HFF : valeur = "REC"
                        Case &HC9 : valeur = "EXIT"
                        Case &HD8 : valeur = "TEXT"
                        Case &HD9 : valeur = "SHIFT-TEXT"
                        Case &HF2 : valeur = "TELETEXT"
                        Case &HD7 : valeur = "SHIFT-TELETEXT"
                        Case &HBA : valeur = "A+B"
                        Case &H52 : valeur = "ENT"
                        Case &HD6 : valeur = "SHIFT-ENT"
                        Case &H70 : valeur = "Cursor-left"
                        Case &H71 : valeur = "Cursor-right"
                        Case &H72 : valeur = "Cursor-up"
                        Case &H73 : valeur = "Cursor-down"
                        Case &H74 : valeur = "Cursor-up-left"
                        Case &H75 : valeur = "Cursor-up-right"
                        Case &H76 : valeur = "Cursor-down-right"
                        Case &H77 : valeur = "Cursor-down-left"
                        Case &H78 : valeur = "Left mouse"
                        Case &H79 : valeur = "Left mouse-End"
                        Case &H7B : valeur = "Drag"
                        Case &H7C : valeur = "Right mouse"
                        Case &H7D : valeur = "Right mouse-End"
                        Case Else : valeur = "unknown"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case Else : WriteLog("ERR: decode_Remote : Unknown Sub type for Packet type=" & Hex(recbuf(REMOTE.packettype)) & ": " & Hex(recbuf(REMOTE.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(REMOTE.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Remote Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Thermostat1()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            'Select Case recbuf(THERMOSTAT1.subtype)
            '    Case THERMOSTAT1.Digimax
            '        WriteMessage("subtype       = Digimax")
            '    Case THERMOSTAT1.DigimaxShort
            '        WriteMessage("subtype       = Digimax with short format")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(THERMOSTAT1.packettype)) & ":" & Hex(recbuf(THERMOSTAT1.subtype)))
            'End Select
            adresse = ((recbuf(THERMOSTAT1.id1) * 256 + recbuf(THERMOSTAT1.id2))).ToString

            valeur = recbuf(THERMOSTAT1.temperature).ToString '°C
            WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, valeur)

            If recbuf(THERMOSTAT1.subtype) = THERMOSTAT1.sTypeDigimax Then
                valeur = recbuf(THERMOSTAT1.set_point).ToString '°C
                WriteRetour(adresse, ListeDevices.TEMPERATURECONSIGNE.ToString, valeur)

                If (recbuf(THERMOSTAT1.mode) And &H80) = 0 Then
                    WriteLog("decode_Thermostat1 - Mode=heating")
                Else
                    WriteLog("decode_Thermostat1 - Mode=Cooling")
                End If
                Select Case (recbuf(THERMOSTAT1.status) And &H3)
                    Case 0 : WriteLog("decode_Thermostat1 Status = no status available")
                    Case 1 : WriteLog("decode_Thermostat1 Status = demand")
                    Case 2 : WriteLog("decode_Thermostat1 Status = no demand")
                    Case 3 : WriteLog("decode_Thermostat1 Status  = initializing")
                End Select
            End If
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(THERMOSTAT1.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Thermostat1 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Thermostat2()
        Try
            'decoding of this type is only implemented for use by simulate and verbose
            'HE105 receive is not implemented in the RFXtrx433 firmware
            'and RTS10 commands are received as Thermostat1 commands
            Dim adresse As String = ""
            Dim valeur As String = ""
            'Select Case recbuf(THERMOSTAT2.subtype)
            '    Case THERMOSTAT2.sTypeHE105
            '        WriteMessage("subtype       = HE105")
            '    Case THERMOSTAT2.sTypeRTS10
            '        WriteMessage("subtype       = RTS10/RFS10/TLX1206")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(THERMOSTAT2.packettype)) & ": " & Hex(recbuf(THERMOSTAT2.subtype)))
            'End Select
            adresse = recbuf(THERMOSTAT2.unitcode).ToString

            If recbuf(THERMOSTAT2.subtype) = THERMOSTAT2.sTypeHE105 Then
                valeur = "switches 1- 5 = "
                If (recbuf(THERMOSTAT2.unitcode) And &H10) = 0 Then valeur &= "OFF " Else valeur &= "ON  "
                If (recbuf(THERMOSTAT2.unitcode) And &H8) = 0 Then valeur &= "OFF " Else valeur &= "ON  "
                If (recbuf(THERMOSTAT2.unitcode) And &H4) = 0 Then valeur &= "OFF " Else  : valeur &= "ON  "
                If (recbuf(THERMOSTAT2.unitcode) And &H2) = 0 Then valeur &= "OFF " Else valeur &= "ON  "
                If (recbuf(THERMOSTAT2.unitcode) And &H1) = 0 Then valeur &= "OFF " Else valeur &= "ON  "
            End If
            valeur &= " Command= "
            Select Case recbuf(THERMOSTAT2.cmnd)
                Case THERMOSTAT2.sOff : valeur &= " OFF"
                Case THERMOSTAT2.sOn : valeur &= " ON"
                Case THERMOSTAT2.sProgram : valeur &= "Program RTS10"
                Case Else : valeur &= "UNKNOWN"
            End Select
            WriteRetour(adresse, "", valeur)
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(THERMOSTAT2.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Thermostat2 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Thermostat3()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            'Select Case recbuf(THERMOSTAT3.subtype)
            '    Case THERMOSTAT3.sTypeMertikG6RH4T1
            '        WriteMessage("subtype       = Mertik G6R-H4T1")
            '    Case THERMOSTAT3.sTypeMertikG6RH4TB
            '        WriteMessage("subtype       = Mertik G6R-H4TB")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(THERMOSTAT3.packettype)) & ":" & Hex(recbuf(THERMOSTAT3.subtype)))
            'End Select
            adresse = VB.Right("0" & Hex(recbuf(THERMOSTAT3.unitcode1)), 2) & VB.Right("0" & Hex(recbuf(THERMOSTAT3.unitcode2)), 2) & VB.Right("0" & Hex(recbuf(THERMOSTAT3.unitcode3)), 2)

            Select Case recbuf(THERMOSTAT3.cmnd)
                Case 0 : valeur = "OFF"
                Case 1 : valeur = "ON"
                Case 2 : valeur = "UP"
                Case 3 : valeur = "DOWN"
                Case 4 : If recbuf(THERMOSTAT3.subtype) = THERMOSTAT3.sTypeMertikG6RH4T1 Then valeur = "RUN_UP" Else valeur = "2ND_OFF"
                Case 5 : If recbuf(THERMOSTAT3.subtype) = THERMOSTAT3.sTypeMertikG6RH4T1 Then valeur = "RUN_DOWN" Else valeur = "2ND_ON"
                Case 6 : If recbuf(THERMOSTAT3.subtype) = THERMOSTAT3.sTypeMertikG6RH4T1 Then valeur = "STOP" Else valeur = "UNKNOWN"
                Case Else : valeur = "UNKNOWN"
            End Select
            WriteRetour(adresse, "", valeur)

            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(THERMOSTAT3.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Thermostat3 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_BBQ()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(BBQ.subtype)
                Case BBQ.sTypeBBQ1
                    'WriteMessage("subtype       = BBQ1 - Maverick ET-732")
                    'WriteMessage("Sequence nbr  = " & recbuf(BBQ.seqnbr).ToString)
                    adresse = (recbuf(BBQ.id1) * 256 + recbuf(BBQ.id2)).ToString
                    WriteRetour(adresse & "-1", ListeDevices.TEMPERATURE.ToString, (recbuf(BBQ.sensor1h) * 256 + recbuf(BBQ.sensor1l)).ToString)
                    WriteRetour(adresse & "-2", ListeDevices.TEMPERATURE.ToString, (recbuf(BBQ.sensor2h) * 256 + recbuf(BBQ.sensor2l)).ToString)
                Case Else : WriteLog("ERR: decode_BBQ : Unknown Sub type for Packet type=" & Hex(recbuf(BBQ.packettype)) & ": " & Hex(recbuf(BBQ.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(BBQ.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
            If (recbuf(BBQ.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
        Catch ex As Exception
            WriteLog("ERR: decode_BBQ Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_TempRain()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(TEMP_RAIN.subtype)
                Case TEMP_RAIN.sTypeTR1
                    'WriteMessage("subtype       = TR1 - WS1200")
                    'WriteMessage("Sequence nbr  = " & recbuf(TEMP_RAIN.seqnbr).ToString)
                    adresse = (recbuf(TEMP_RAIN.id1) * 256 + recbuf(TEMP_RAIN.id2)).ToString
                    If (recbuf(TEMP_RAIN.tempsign) And &H80) = 0 Then
                        WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, Math.Round((recbuf(TEMP_RAIN.temperatureh) * 256 + recbuf(TEMP_RAIN.temperaturel)) / 10, 2).ToString)
                    Else
                        WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, Math.Round(((recbuf(TEMP_RAIN.temperatureh) And &H7F) * 256 + recbuf(TEMP_RAIN.temperaturel)) / 10, 2).ToString)
                    End If
                    WriteRetour(adresse, ListeDevices.PLUIETOTAL.ToString, Math.Round((recbuf(TEMP_RAIN.raintotal1) * 256 + recbuf(TEMP_RAIN.raintotal2)) / 10, 2).ToString)
                Case Else : WriteLog("ERR: decode_TEMP_RAIN : Unknown Sub type for Packet type=" & Hex(recbuf(TEMP_RAIN.packettype)) & ": " & Hex(recbuf(TEMP_RAIN.subtype)))
            End Select

            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(TEMP_RAIN.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
            If (recbuf(TEMP_RAIN.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
        Catch ex As Exception
            WriteLog("ERR: decode_TempRain Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Temp()
        Try
            'Select Case recbuf(TEMP.subtype)
            '    Case TEMP.sTypeTEMP1 : WriteMessage("subtype       = TEMP1 - THR128/138, THC138 channel " & recbuf(TEMP.id2).ToString)
            '    Case TEMP.sTypeTEMP2 : WriteMessage("subtype       = TEMP2 - THC238/268,THN132,THWR288,THRN122,THN122,AW129/131 channel " & recbuf(TEMP.id2).ToString)
            '    Case TEMP.sTypeTEMP3 : WriteMessage("subtype       = TEMP3 - THWR800")
            '    Case TEMP.sTypeTEMP4 : WriteMessage("subtype       = TEMP4 - RTHN318  channel " & recbuf(TEMP.id2).ToString)
            '    Case TEMP.sTypeTEMP5 : WriteMessage("subtype       = TEMP5 - LaCrosse TX2, TX3, TX4, TX17")
            '    Case TEMP.sTypeTEMP6 : WriteMessage("subtype       = TEMP6 - TS15C")
            '    Case TEMP.sTypeTEMP7 : WriteMessage("subtype       = TEMP7 - Viking 02811")
            '    Case TEMP.sTypeTEMP8 :  WriteMessage("subtype       = TEMP8 - LaCrosse WS2300")
            '    Case TEMP.sTypeTEMP9
            '        If recbuf(TEMP.id2) = &HFF Then
            '            WriteMessage("subtype       = TEMP9 - RUBiCSON 48659 stektermometer")
            '        Else
            '            WriteMessage("subtype       = TEMP9 - RUBiCSON 48695")
            '        End If
            '    Case TEMP.sTypeTEMP10 : WriteMessage("subtype       = TEMP10 - TFA 30.3133")
            '    Case Else : WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(TEMP.packettype)) & ":" & Hex(recbuf(TEMP.subtype)))
            'End Select
            'WriteMessage("Sequence nbr  = " & recbuf(TEMP.seqnbr).ToString)
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(TEMP.id1) * 256 + recbuf(TEMP.id2)).ToString
            If (recbuf(TEMP.tempsign) And &H80) = 0 Then
                valeur = Math.Round((recbuf(TEMP.temperatureh) * 256 + recbuf(TEMP.temperaturel)) / 10, 2).ToString
            Else
                valeur = (-(Math.Round(((recbuf(TEMP.temperatureh) And &H7F) * 256 + recbuf(TEMP.temperaturel)) / 10, 2))).ToString
            End If
            WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, valeur)
            If (recbuf(TEMP.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(TEMP.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Temp Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Hum()
        Try
            'Select Case recbuf(HUM.subtype)
            '    Case HUM.sTypeHUM1
            '        WriteMessage("subtype       = HUM1 - LaCrosse TX3")
            '    Case HUM.sTypeHUM2
            '        WriteMessage("subtype       = HUM2 - LaCrosse WS2300")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(HUM.packettype)) & ":" & Hex(recbuf(HUM.subtype)))
            'End Select
            'WriteMessage("Sequence nbr  = " & recbuf(HUM.seqnbr).ToString)
            'Select Case recbuf(HUM.humidity_status)
            '    Case HUM.snormal
            '        WriteMessage("Status        = Normal")
            '    Case HUM.scomfort
            '        WriteMessage("Status        = Comfortable")
            '    Case HUM.sdry
            '        WriteMessage("Status        = Dry")
            '    Case HUM.swet
            '        WriteMessage("Status        = Wet")
            'End Select
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(HUM.id1) * 256 + recbuf(HUM.id2)).ToString
            valeur = recbuf(HUM.humidity).ToString
            WriteRetour(adresse, ListeDevices.HUMIDITE.ToString, valeur)
            If (recbuf(HUM.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(HUM.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Hum Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_TempHum()
        Try
            'Select Case recbuf(TEMP_HUM.subtype)
            '    Case TEMP_HUM.sTypeTH1
            '        WriteMessage("subtype       = TH1 - THGN122/123/132,THGR122/228/238/268")
            '        WriteMessage("                channel " & recbuf(TEMP_HUM.id2).ToString)
            '    Case TEMP_HUM.sTypeTH2
            '        WriteMessage("subtype       = TH2 - THGN800,TGHN801,THGR810")
            '        WriteMessage("                channel " & recbuf(TEMP_HUM.id2).ToString)
            '    Case TEMP_HUM.sTypeTH3
            '        WriteMessage("subtype       = TH3 - RTGN318,RTGR328,RTGR368")
            '        WriteMessage("                channel " & recbuf(TEMP_HUM.id2).ToString)
            '    Case TEMP_HUM.sTypeTH4
            '        WriteMessage("subtype       = TH4 - THGR328")
            '        WriteMessage("                channel " & recbuf(TEMP_HUM.id2).ToString)
            '    Case TEMP_HUM.sTypeTH5
            '        WriteMessage("subtype       = TH5 - WTGR800")
            '    Case TEMP_HUM.sTypeTH6
            '        WriteMessage("subtype       = TH6 - THGR918/928,THGRN228,THGN500")
            '        WriteMessage("                channel " & recbuf(TEMP_HUM.id2).ToString)
            '    Case TEMP_HUM.sTypeTH7
            '        WriteMessage("subtype       = TH7 - Cresta, TFA TS34C")
            '        If recbuf(TEMP_HUM.id1) < &H40 Then
            '            WriteMessage("                channel 1")
            '        ElseIf recbuf(TEMP_HUM.id1) < &H60 Then
            '            WriteMessage("                channel 2")
            '        ElseIf recbuf(TEMP_HUM.id1) < &H80 Then
            '            WriteMessage("                channel 3")
            '        ElseIf recbuf(TEMP_HUM.id1) > &H9F And recbuf(TEMP_HUM.id1) < &HC0 Then
            '            WriteMessage("                channel 4")
            '        ElseIf recbuf(TEMP_HUM.id1) < &HE0 Then
            '            WriteMessage("                channel 5")
            '        Else
            '            WriteMessage("                channel ??")
            '        End If
            '    Case TEMP_HUM.sTypeTH8
            '        WriteMessage("subtype       = TH8 - WT260,WT260H,WT440H,WT450,WT450H")
            '        WriteMessage("                channel " & recbuf(TEMP_HUM.id2).ToString)
            '    Case TEMP_HUM.sTypeTH9
            '        WriteMessage("subtype       = TH9 - Viking 02038, 02035 (02035 has no humidity)")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(TEMP_HUM.packettype)) & ":" & Hex(recbuf(TEMP_HUM.subtype)))
            'End Select
            'WriteMessage("Sequence nbr  = " & recbuf(TEMP_HUM.seqnbr).ToString)
            'Select Case recbuf(TEMP_HUM.humidity_status)
            '    Case HUM.snormal
            '        WriteMessage("Status        = Normal")
            '    Case HUM.scomfort
            '        WriteMessage("Status        = Comfortable")
            '    Case HUM.sdry
            '        WriteMessage("Status        = Dry")
            '    Case HUM.swet
            '        WriteMessage("Status        = Wet")
            'End Select
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(TEMP_HUM.id1) * 256 + recbuf(TEMP_HUM.id2)).ToString
            If (recbuf(TEMP_HUM.tempsign) And &H80) = 0 Then
                valeur = Math.Round((recbuf(TEMP_HUM.temperatureh) * 256 + recbuf(TEMP_HUM.temperaturel)) / 10, 2).ToString
            Else
                valeur = (-(Math.Round(((recbuf(TEMP_HUM.temperatureh) And &H7F) * 256 + recbuf(TEMP_HUM.temperaturel)) / 10, 2))).ToString
            End If
            WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, valeur)
            valeur = recbuf(TEMP_HUM.humidity).ToString
            WriteRetour(adresse, ListeDevices.HUMIDITE.ToString, valeur)
            If recbuf(TEMP_HUM.subtype) = TEMP_HUM.sTypeTH6 Then
                'battery low < 10% (0=10% 1=20% 2=30%....9=100%)
                If recbuf(TEMP_HUM.battery_level) = 0 Then
                    WriteBattery(adresse, "0")
                Else
                    WriteBattery(adresse, ((recbuf(TEMP_HUM.battery_level) + 1) * 10) & "%")
                End If
            Else
                If (recbuf(TEMP_HUM.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            End If
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(TEMP_HUM.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_TempHum Exception : " & ex.Message)
        End Try
    End Sub
    'Not implemented
    Private Sub decode_Baro()
        'WriteMessage("Not implemented")
    End Sub

    Private Sub decode_TempHumBaro()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(TEMP_HUM_BARO.id1) * 256 + recbuf(TEMP_HUM_BARO.id2)).ToString
            'Select Case recbuf(TEMP_HUM_BARO.subtype)
            '    Case TEMP_HUM_BARO.sTypeTHB1
            '        WriteMessage("subtype       = THB1 - BTHR918")
            '        WriteMessage("                channel " & recbuf(TEMP_HUM_BARO.id2).ToString)
            '    Case TEMP_HUM_BARO.sTypeTHB2
            '        WriteMessage("subtype       = THB2 - BTHR918N, BTHR968")
            '        WriteMessage("                channel " & recbuf(TEMP_HUM_BARO.id2).ToString)
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(TEMP_HUM_BARO.packettype)) & ":" & Hex(recbuf(TEMP_HUM_BARO.subtype)))
            'End Select

            If (TEMP_HUM_BARO.tempsign And &H80) = 0 Then
                valeur = Math.Round((recbuf(TEMP_HUM_BARO.temperatureh) * 256 + recbuf(TEMP_HUM_BARO.temperaturel)) / 10, 2).ToString '°C
            Else
                valeur = (-(Math.Round(((recbuf(TEMP_HUM_BARO.temperatureh) And &H7F) * 256 + recbuf(TEMP_HUM_BARO.temperaturel)) / 10, 2))).ToString '°C
            End If
            WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, valeur)

            valeur = recbuf(TEMP_HUM_BARO.humidity).ToString
            WriteRetour(adresse, ListeDevices.HUMIDITE.ToString, valeur)

            'Select Case recbuf(TEMP_HUM_BARO.humidity_status)
            '    Case HUM.snormal
            '        WriteMessage("Status        = Normal")
            '    Case HUM.scomfort
            '        WriteMessage("Status        = Comfortable")
            '    Case HUM.sdry
            '        WriteMessage("Status        = Dry")
            '    Case HUM.swet
            '        WriteMessage("Status        = Wet")
            'End Select

            valeur = (recbuf(TEMP_HUM_BARO.baroh) * 256 + recbuf(TEMP_HUM_BARO.barol)).ToString
            WriteRetour(adresse, ListeDevices.BAROMETRE.ToString, valeur)

            'Select Case recbuf(TEMP_HUM_BARO.forecast)
            '    Case &H0
            '        WriteMessage("Forecast      = No information available")
            '    Case &H1
            '        WriteMessage("Forecast      = Sunny")
            '    Case &H2
            '        WriteMessage("Forecast      = Partly Cloudy")
            '    Case &H3
            '        WriteMessage("Forecast      = Cloudy")
            '    Case &H4
            '        WriteMessage("Forecast      = Rain")
            'End Select

            If (recbuf(TEMP_HUM_BARO.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(TEMP_HUM_BARO.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_TempHumBaro Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Rain()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(RAIN.id1) * 256 + recbuf(RAIN.id2)).ToString
            'Select Case recbuf(RAIN.subtype)
            '    Case RAIN.sTypeRAIN1
            '        WriteMessage("subtype       = RAIN1 - RGR126/682/918/928")
            '    Case RAIN.sTypeRAIN2
            '        WriteMessage("subtype       = RAIN2 - PCR800")
            '    Case RAIN.sTypeRAIN3
            '        WriteMessage("subtype       = RAIN3 - TFA")
            '    Case RAIN.sTypeRAIN4
            '        WriteMessage("subtype       = RAIN4 - UPM RG700")
            '    Case RAIN.sTypeRAIN5
            '        WriteMessage("subtype       = RAIN5 - LaCrosse WS2300")
            '    Case RAIN.sTypeRAIN6
            '        WriteMessage("subtype       = RAIN6 - LaCrosse TX5")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(RAIN.packettype)) & ":" & Hex(recbuf(RAIN.subtype)))
            'End Select
            If recbuf(RAIN.subtype) = RAIN.sTypeRAIN1 Then
                valeur = ((recbuf(RAIN.rainrateh) * 256) + recbuf(RAIN.rainratel)).ToString 'mm/h
            ElseIf recbuf(RAIN.subtype) = RAIN.sTypeRAIN2 Then
                valeur = (((recbuf(RAIN.rainrateh) * 256) + recbuf(RAIN.rainratel)) / 100).ToString ' mm/h
            End If
            WriteRetour(adresse, ListeDevices.PLUIECOURANT.ToString, valeur)


            If recbuf(RAIN.subtype) = RAIN.sTypeRAIN6 Then
                If shortTotalRain = 0 Then   'TX5 not yet received before
                    byteFlipCount = recbuf(RAIN.raintotal3)
                    shortTotalRain = 0.001   'indicate TX5 received
                Else
                    If byteFlipCount > recbuf(RAIN.raintotal3) Then
                        shortTotalRain = (recbuf(RAIN.raintotal3) + 16 - byteFlipCount) * 0.266
                    Else
                        shortTotalRain = (recbuf(RAIN.raintotal3) - byteFlipCount) * 0.266
                    End If
                    byteFlipCount = recbuf(RAIN.raintotal3)
                End If
                valeur = Math.Round(shortTotalRain, 2).ToString 'mm  (since program start)
            Else
                valeur = ((recbuf(RAIN.raintotal1) * 65535) + recbuf(RAIN.raintotal2) * 256 + recbuf(RAIN.raintotal3)).ToString 'mm
            End If
            WriteRetour(adresse, ListeDevices.PLUIETOTAL.ToString, valeur)

            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(RAIN.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
            If (recbuf(RAIN.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
        Catch ex As Exception
            WriteLog("ERR: decode_Rain Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Wind()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(WIND.id1) * 256 + recbuf(WIND.id2)).ToString
            'Select Case recbuf(WIND.subtype)
            '    Case WIND.sTypeWIND1
            '        WriteMessage("subtype       = WIND1 - WTGR800")
            '    Case WIND.sTypeWIND2
            '        WriteMessage("subtype       = WIND2 - WGR800")
            '    Case WIND.sTypeWIND3
            '        WriteMessage("subtype       = WIND3 - STR918/928, WGR918")
            '    Case WIND.sTypeWIND4
            '        WriteMessage("subtype       = WIND4 - TFA")
            '    Case WIND.sTypeWIND5
            '        WriteMessage("subtype       = WIND5 - UPM WDS500")
            '    Case WIND.sTypeWIND6
            '        WriteMessage("subtype       = WIND6 - LaCrosse WS2300")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(WIND.packettype)) & ":" & Hex(recbuf(WIND.subtype)))
            'End Select

            valeur = ((recbuf(WIND.directionh) * 256) + recbuf(WIND.directionl)).ToString 'direction en degré
            WriteRetour(adresse, ListeDevices.DIRECTIONVENT.ToString, valeur)
            If recbuf(WIND.subtype) <> WIND.sTypeWIND5 Then
                'valeur = (((recbuf(WIND.av_speedh) * 256) + recbuf(WIND.av_speedl)) / 10).ToString 'Vitesse en mtr/sec
                valeur = (((recbuf(WIND.av_speedh) * 256) + recbuf(WIND.av_speedl)) * 0.36).ToString 'Vitesse en km/hr
                WriteRetour(adresse, ListeDevices.VITESSEVENT.ToString, valeur)
            End If

            'WriteMessage("Wind gust     = " & (((recbuf(WIND.gusth) * 256) + recbuf(WIND.gustl)) / 10).ToString & " mtr/sec")

            If recbuf(WIND.subtype) = WIND.sTypeWIND4 Then
                If (WIND.tempsign And &H80) = 0 Then
                    valeur = Math.Round((recbuf(WIND.temperatureh) * 256 + recbuf(WIND.temperaturel)) / 10, 2).ToString '°C
                Else
                    valeur = (-(Math.Round(((recbuf(WIND.temperatureh) And &H7F) * 256 + recbuf(WIND.temperaturel)) / 10, 2))).ToString '°C
                End If
                WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, valeur)

                If (WIND.chillsign And &H80) = 0 Then
                    '        WriteMessage("Chill         = " & Math.Round((recbuf(WIND.chillh) * 256 + recbuf(WIND.chilll)) / 10, 2).ToString & " °C")
                Else
                    '        WriteMessage("Chill         = -" & Math.Round(((recbuf(WIND.chillh) And &H7F) * 256 + recbuf(WIND.chilll)) / 10, 2).ToString & " °C")
                End If
            End If

            If recbuf(WIND.subtype) = WIND.sTypeWIND3 Then
                'battery low < 10% (0=10% 1=20% 2=30%....9=100%)
                If recbuf(WIND.battery_level) = 0 Then
                    WriteBattery(adresse, "0")
                Else
                    WriteBattery(adresse, ((recbuf(WIND.battery_level) + 1) * 10) & "%")
                End If
            Else
                If (recbuf(WIND.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            End If
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(WIND.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Wind Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_UV()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(UV.id1) * 256 + recbuf(UV.id2)).ToString
            'Select Case recbuf(UV.subtype)
            '    Case UV.sTypeUV1
            '        WriteMessage("Subtype       = UV1 - UVN128,UVR128,UV138")
            '    Case UV.sTypeUV2
            '        WriteMessage("Subtype       = UV2 - UVN800")
            '    Case UV.sTypeUV3
            '        WriteMessage("Subtype       = UV3 - TFA")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(UV.packettype)) & ":" & Hex(recbuf(UV.subtype)))
            'End Select

            valeur = (recbuf(UV.uv) / 10).ToString
            WriteRetour(adresse, ListeDevices.UV.ToString, valeur)

            If recbuf(UV.subtype) = UV.sTypeUV3 Then
                If (UV.tempsign And &H80) = 0 Then
                    valeur = Math.Round((recbuf(UV.temperatureh) * 256 + recbuf(UV.temperaturel)) / 10, 2).ToString '°C
                Else
                    valeur = (-(Math.Round(((recbuf(UV.temperatureh) And &H7F) * 256 + recbuf(UV.temperaturel)) / 10, 2))).ToString '°C
                End If
                WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, valeur)
            End If

            'If recbuf(UV.uv) < 3 Then
            '    WriteMessage("Description = Low")
            'ElseIf recbuf(UV.uv) < 6 Then
            '    WriteMessage("Description = Medium")
            'ElseIf recbuf(UV.uv) < 8 Then
            '    WriteMessage("Description = High")
            'ElseIf recbuf(UV.uv) < 11 Then
            '    WriteMessage("Description = Very high")
            'Else
            '    WriteMessage("Description = Dangerous")
            'End If

            If (recbuf(UV.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(UV.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_UV Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_DateTime()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(DT.id1) * 256 + recbuf(DT.id2)).ToString
            'Select Case recbuf(DT.subtype)
            '    Case DT.sTypeDT1
            '        WriteMessage("Subtype       = DT1 - RTGR328N")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(DT.packettype)) & ":" & Hex(recbuf(DT.subtype)))
            'End Select
            valeur = recbuf(DT.yy).ToString + "/" & recbuf(DT.mm).ToString & "/" & recbuf(DT.dd).ToString + " " + recbuf(DT.hr).ToString + ":" & recbuf(DT.min).ToString & ":" & recbuf(DT.sec).ToString
            WriteRetour(adresse, "", valeur)

            If (recbuf(DT.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(UV.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_DateTime Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Current()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(CURRENT.id1) * 256 + recbuf(CURRENT.id2)).ToString
            'Select Case recbuf(CURRENT.subtype)
            '    Case CURRENT.sTypeELEC1
            '        WriteMessage("subtype       = ELEC1 - OWL CM113, Electrisave, cent-a-meter")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(CURRENT.packettype)) & ":" & Hex(recbuf(CURRENT.subtype)))
            'End Select
            'WriteMessage("Count         = " & recbuf(CURRENT.count).ToString)

            valeur = ((recbuf(CURRENT.ch1h) * 256 + recbuf(CURRENT.ch1l)) / 10).ToString 'ampere channel 1
            WriteRetour(adresse & "-1", "", valeur)
            valeur = ((recbuf(CURRENT.ch2h) * 256 + recbuf(CURRENT.ch2h)) / 10).ToString 'ampere channel 2
            WriteRetour(adresse & "-2", "", valeur)
            valeur = ((recbuf(CURRENT.ch3l) * 256 + recbuf(CURRENT.ch3l)) / 10).ToString 'ampere channel 3
            WriteRetour(adresse & "-3", "", valeur)

            If (recbuf(CURRENT.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(CURRENT.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Current Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Energy()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(ENERGY.id1) * 256 + recbuf(ENERGY.id2)).ToString
            'Select Case recbuf(ENERGY.subtype)
            '    Case ENERGY.ELEC2
            '        WriteMessage("subtype       = ELEC2 - OWL CM119, CM160")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(ENERGY.packettype)) & ":" & Hex(recbuf(ENERGY.subtype)))
            'End Select

            'WriteMessage("Count         = " & recbuf(ENERGY.count).ToString)

            valeur = (CLng(recbuf(ENERGY.instant1)) * &H1000000 + recbuf(ENERGY.instant2) * &H10000 + recbuf(ENERGY.instant3) * &H100 + recbuf(ENERGY.instant4)).ToString 'Watt
            WriteRetour(adresse, ListeDevices.ENERGIEINSTANTANEE.ToString, valeur)
            valeur = ((CDbl(recbuf(ENERGY.total1)) * &H10000000000 + CDbl(recbuf(ENERGY.total2)) * &H100000000 + CDbl(recbuf(ENERGY.total3)) * &H1000000 + recbuf(ENERGY.total4) * &H10000 + recbuf(ENERGY.total5) * &H100 + recbuf(ENERGY.total6)) / 223.666).ToString 'Watt / h
            'WriteMessage("total usage   = " & Math.Round(usage, 1).ToString & " Wh")
            WriteRetour(adresse, ListeDevices.ENERGIETOTALE.ToString, valeur)

            If (recbuf(ENERGY.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(ENERGY.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Energy Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Current_Energy()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(CURRENT_ENERGY.id1) * 256 + recbuf(CURRENT_ENERGY.id2)).ToString
            'Select Case recbuf(CURRENT_ENERGY.subtype)
            '    Case CURRENT_ENERGY.sTypeELEC4
            '        WriteMessage("subtype       = ELEC4 - OWL CM180i")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(CURRENT_ENERGY.packettype)) & ":" & Hex(recbuf(CURRENT_ENERGY.subtype)))
            'End Select

            'WriteMessage("Count         = " & recbuf(CURRENT_ENERGY.count).ToString)
            valeur = ((recbuf(CURRENT_ENERGY.ch1h) * 256 + recbuf(CURRENT_ENERGY.ch1l)) / 10).ToString 'ampere channel 1
            WriteRetour(adresse & "-1", ListeDevices.ENERGIEINSTANTANEE.ToString, valeur)
            valeur = ((recbuf(CURRENT_ENERGY.ch2h) * 256 + recbuf(CURRENT_ENERGY.ch2l)) / 10).ToString 'ampere channel 2
            WriteRetour(adresse & "-2", ListeDevices.ENERGIEINSTANTANEE.ToString, valeur)
            valeur = ((recbuf(CURRENT_ENERGY.ch3h) * 256 + recbuf(CURRENT_ENERGY.ch3l)) / 10).ToString 'ampere channel 3
            WriteRetour(adresse & "-3", ListeDevices.ENERGIEINSTANTANEE.ToString, valeur)

            valeur = Math.Round(((CDbl(recbuf(CURRENT_ENERGY.total1)) * &H10000000000 + CDbl(recbuf(CURRENT_ENERGY.total2)) * &H100000000 + CDbl(recbuf(CURRENT_ENERGY.total3)) * &H1000000 + recbuf(CURRENT_ENERGY.total4) * &H10000 + recbuf(CURRENT_ENERGY.total5) * &H100 + recbuf(CURRENT_ENERGY.total6)) / 223.666), 1).ToString 'Watt / h
            WriteRetour(adresse, ListeDevices.ENERGIETOTALE.ToString, valeur)

            If (recbuf(CURRENT_ENERGY.battery_level) And &HF) = 0 Then WriteBattery(adresse, "0") Else WriteBattery(adresse, "100")
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(CURRENT_ENERGY.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Current_Energy Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Power()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(POWER.subtype)
                Case POWER.sTypeELEC5
                    'WriteMessage("subtype       = ELEC5 - Revolt")
                    adresse = (recbuf(POWER.id1) * 256 + recbuf(POWER.id2)).ToString
                    'WriteMessage("Voltage       = " & recbuf(POWER.voltage).ToString & " Volt")
                    'WriteMessage("Current       = " & ((recbuf(POWER.currentH) * 256 + recbuf(POWER.currentL)) / 100).ToString & " Ampere")
                    WriteRetour(adresse, ListeDevices.ENERGIEINSTANTANEE.ToString, ((recbuf(POWER.powerH) * 256 + recbuf(POWER.powerL)) / 10).ToString) 'W
                    WriteRetour(adresse, ListeDevices.ENERGIETOTALE.ToString, ((recbuf(POWER.energyH) * 256 + recbuf(POWER.energyL)) / 100).ToString) 'kWh
                    'WriteMessage("power factor  = " & (recbuf(POWER.pf) / 100).ToString)
                    'WriteMessage("Frequency     = " & recbuf(POWER.freq).ToString)
                Case Else : WriteLog("ERR: decode_Power : Unknown Sub type for Packet type=" & Hex(recbuf(POWER.packettype)) & ": " & Hex(recbuf(POWER.subtype)))
            End Select

            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(POWER.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Power Exception : " & ex.Message)
        End Try
    End Sub

    'Not implemented
    Private Sub decode_Gas()
        Try
            'WriteMessage("Not implemented")
        Catch ex As Exception
            WriteLog("ERR: decode_Gas Exception : " & ex.Message)
        End Try
    End Sub
    'Not implemented
    Private Sub decode_Water()
        Try
            'WriteMessage("Not implemented")
        Catch ex As Exception
            WriteLog("ERR: decode_Water Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_Weight()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            adresse = (recbuf(WEIGHT.id1) * 256 + recbuf(WEIGHT.id2)).ToString
            'Select Case recbuf(WEIGHT.subtype)
            '    Case WEIGHT.sTypeWEIGHT1
            '        WriteMessage("subtype       = BWR102")
            '    Case WEIGHT.sTypeWEIGHT2
            '        WriteMessage("subtype       = GR101")
            '    Case Else
            '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(WEIGHT.packettype)) & ":" & Hex(recbuf(WEIGHT.subtype)))
            'End Select

            valeur = ((recbuf(WEIGHT.weighthigh) * 25.6) + recbuf(WEIGHT.weightlow) / 10).ToString 'kg
            WriteRetour(adresse, "", valeur)

            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(WEIGHT.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_Weight Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_RFXSensor()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(RFXSENSOR.subtype)
                Case RFXSENSOR.sTypeTemp 'Temperature
                    adresse = recbuf(RFXSENSOR.id).ToString
                    If (recbuf(RFXSENSOR.msg1) And &H80) = 0 Then 'positive temperature?
                        valeur = Math.Round(((recbuf(RFXSENSOR.msg1) * 256 + recbuf(RFXSENSOR.msg2)) / 100), 2).ToString '°C
                    Else
                        valeur = Math.Round((0 - ((recbuf(RFXSENSOR.msg1) And &H7F) * 256 + recbuf(RFXSENSOR.msg2)) / 100), 2).ToString '°C
                    End If
                    WriteRetour(adresse, ListeDevices.TEMPERATURE.ToString, valeur)
                Case RFXSENSOR.sTypeAD 'A/D
                    adresse = recbuf(RFXSENSOR.id).ToString
                    valeur = (recbuf(RFXSENSOR.msg1) * 256 + recbuf(RFXSENSOR.msg2)).ToString ' mV
                    WriteRetour(adresse, "", valeur)
                Case RFXSENSOR.sTypeVolt 'Voltage
                    adresse = recbuf(RFXSENSOR.id).ToString
                    valeur = (recbuf(RFXSENSOR.msg1) * 256 + recbuf(RFXSENSOR.msg2)).ToString ' mV
                    WriteRetour(adresse, "", valeur)
                Case RFXSENSOR.sTypeMessage 'Message
                    adresse = recbuf(RFXSENSOR.id).ToString
                    Select Case recbuf(RFXSENSOR.msg2)
                        Case &H1 : valeur = "sensor addresses incremented"
                        Case &H2 : valeur = "battery low detected"
                        Case &H81 : valeur = "no 1-wire device connected"
                        Case &H82 : valeur = "1-Wire ROM CRC error"
                        Case &H83 : valeur = "1-Wire device connected is not a DS18B20 or DS2438"
                        Case &H84 : valeur = "no end of read signal received from 1-Wire device"
                        Case &H85 : valeur = "1-Wire scratchpad CRC error"
                        Case Else : valeur = "ERROR: unknown message"
                    End Select
                    WriteLog("decode_RFXSensor Message : " & valeur & " de " & adresse)
                    valeur = (recbuf(RFXSENSOR.msg1) * 256 + recbuf(RFXSENSOR.msg2)).ToString
                    WriteRetour(adresse, "", valeur)
                Case Else : WriteLog("ERR: decode_RFXSensor : Unknown Sub type for Packet type=" & Hex(recbuf(RFXSENSOR.packettype)) & ": " & Hex(recbuf(RFXSENSOR.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(RFXSENSOR.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_RFXSensor Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_RFXMeter()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Dim counter As Long

            Select Case recbuf(RFXMETER.subtype)
                Case RFXMETER.sTypeCount
                    'WriteMessage("subtype       = RFXMeter counter")
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    counter = (CLng(recbuf(RFXMETER.count1)) << 24) + (CLng(recbuf(RFXMETER.count2)) << 16) + (CLng(recbuf(RFXMETER.count3)) << 8) + recbuf(RFXMETER.count4)
                    valeur = counter.ToString 'WriteMessage("if RFXPwr     = " & (counter / 1000).ToString & " kWh")
                    WriteRetour(adresse, "", valeur)
                Case RFXMETER.sTypeInterval
                    'WriteMessage("subtype       = RFXMeter new interval time set")
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    Select Case recbuf(RFXMETER.count3)
                        Case &H1 : valeur = "Interval: 30 sec"
                        Case &H2 : valeur = "Interval: 1 min"
                        Case &H4 : valeur = "Interval: 6 min"
                        Case &H8 : valeur = "Interval: 12 min"
                        Case &H10 : valeur = "Interval: 15 min"
                        Case &H20 : valeur = "Interval: 30 min"
                        Case &H40 : valeur = "Interval: 45 min"
                        Case &H80 : valeur = "Interval: 60 min"
                        Case Else : valeur = "Interval: illegal value"
                    End Select
                    WriteRetour(adresse, "", "CFG: " & valeur)
                Case RFXMETER.sTypeCalib
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    counter = CLng(((CLng(recbuf(RFXMETER.count2) And &H3F) << 16) + (CLng(recbuf(RFXMETER.count3)) << 8) + recbuf(RFXMETER.count4)) / 1000)
                    Select Case (recbuf(RFXMETER.count2) And &HC0)
                        Case &H0 : WriteRetour(adresse, "", "CFG: Calibrate mode for channel 1 : " & counter.ToString & " msec - RFXPwr        = " & Convert.ToString(Round(1 / ((16 * counter) / (3600000 / 62.5)), 3)) & " kW")
                        Case &H40 : WriteRetour(adresse, "", "CFG: Calibrate mode for channel 2 : " & counter.ToString & " msec - RFXPwr        = " & Convert.ToString(Round(1 / ((16 * counter) / (3600000 / 62.5)), 3)) & " kW")
                        Case &H80 : WriteRetour(adresse, "", "CFG: Calibrate mode for channel 3 : " & counter.ToString & " msec - RFXPwr        = " & Convert.ToString(Round(1 / ((16 * counter) / (3600000 / 62.5)), 3)) & " kW")
                    End Select
                Case RFXMETER.sTypeAddr
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    WriteRetour(adresse, "", "CFG: New address set, push button for next address")
                Case RFXMETER.sTypeCounterReset
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    Select Case (recbuf(RFXMETER.count2) And &HC0)
                        Case &H0 : WriteRetour(adresse, "", "CFG: Push the button for next mode within 5 seconds or else RESET COUNTER channel 1 will be executed")
                        Case &H40 : WriteRetour(adresse, "", "CFG: Push the button for next mode within 5 seconds or else RESET COUNTER channel 2 will be executed")
                        Case &H80 : WriteRetour(adresse, "", "CFG: Push the button for next mode within 5 seconds or else RESET COUNTER channel 3 will be executed")
                    End Select
                Case RFXMETER.sTypeCounterSet
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    valeur = ((CLng(recbuf(RFXMETER.count1)) << 24) + (CLng(recbuf(RFXMETER.count2)) << 16) + (CLng(recbuf(RFXMETER.count3)) << 8) + recbuf(RFXMETER.count4)).ToString
                    Select Case (recbuf(RFXMETER.count2) And &HC0)
                        Case &H0 : WriteRetour(adresse, "", "CFG: Counter channel 1 is reset to zero - Valeur:" & valeur)
                        Case &H40 : WriteRetour(adresse, "", "CFG: Counter channel 2 is reset to zero - Valeur:" & valeur)
                        Case &H80 : WriteRetour(adresse, "", "CFG: Counter channel 3 is reset to zero - Valeur:" & valeur)
                    End Select
                Case RFXMETER.sTypeSetInterval
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    WriteRetour(adresse, "", "CFG: Push the button for next mode within 5 seconds or else SET INTERVAL MODE will be entered")
                Case RFXMETER.sTypeSetCalib
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    Select Case (recbuf(RFXMETER.count2) And &HC0)
                        Case &H0 : WriteRetour(adresse, "", "CFG: Push the button for next mode within 5 seconds or else CALIBRATION mode for channel 1 will be executed")
                        Case &H40 : WriteRetour(adresse, "", "CFG: Push the button for next mode within 5 seconds or else CALIBRATION mode for channel 2 will be executed")
                        Case &H80 : WriteRetour(adresse, "", "CFG: Push the button for next mode within 5 seconds or else CALIBRATION mode for channel 3 will be executed")
                    End Select
                Case RFXMETER.sTypeSetAddr
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    WriteRetour(adresse, "", "CFG: Push the button for next mode within 5 seconds or else SET ADDRESS MODE will be entered")
                Case RFXMETER.sTypeIdent
                    'WriteMessage("subtype       = RFXMeter identification")
                    'WriteMessage("Sequence nbr  = " & recbuf(RFXMETER.seqnbr).ToString)
                    adresse = (recbuf(RFXMETER.id1) * 256 + recbuf(RFXMETER.id2)).ToString
                    WriteRetour(adresse, "", "CFG: FW version" & Hex(recbuf(RFXMETER.count3)))
                    Select Case recbuf(RFXMETER.count4)
                        Case &H1 : valeur = "Interval: 30 sec"
                        Case &H2 : valeur = "Interval: 1 min"
                        Case &H4 : valeur = "Interval: 6 min"
                        Case &H8 : valeur = "Interval: 12 min"
                        Case &H10 : valeur = "Interval: 15 min"
                        Case &H20 : valeur = "Interval: 30 min"
                        Case &H40 : valeur = "Interval: 45 min"
                        Case &H80 : valeur = "Interval: 60 min"
                        Case Else : valeur = "Interval: illegal value"
                    End Select
                    WriteRetour(adresse, "", "CFG: " & valeur)
                Case Else
                    '        WriteMessage("ERROR: Unknown Sub type for Packet type=" & Hex(recbuf(RFXMETER.packettype)) & ":" & Hex(recbuf(RFXMETER.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(RFXMETER.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_RFXMeter Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_FS20()
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(FS20.subtype)
                Case FS20.sTypeFS20
                    'WriteMessage("subtype       = FS20")
                    'WriteMessage("Sequence nbr  = " & recbuf(FS20.seqnbr).ToString)
                    'WriteMessage("House code    = " & VB.Right("0" & Hex(recbuf(FS20.hc1)), 2) & VB.Right("0" & Hex(recbuf(FS20.hc2)), 2))
                    'WriteMessage("Address       = " & VB.Right("0" & Hex(recbuf(FS20.addr)), 2))
                    adresse = VB.Right("0" & Hex(recbuf(FS20.hc1)), 2) & VB.Right("0" & Hex(recbuf(FS20.hc2)), 2) & VB.Right("0" & Hex(recbuf(FS20.addr)), 2)
                    Select Case (recbuf(FS20.cmd1) And &H1F)
                        Case &H0 : valeur = "0"
                        Case &H1 : valeur = "6.25"
                        Case &H2 : valeur = "12.5"
                        Case &H3 : valeur = "18.75"
                        Case &H4 : valeur = "2"
                        Case &H5 : valeur = "31.25"
                        Case &H6 : valeur = "37.5"
                        Case &H7 : valeur = "43.75"
                        Case &H8 : valeur = "50"
                        Case &H9 : valeur = "56.25"
                        Case &HA : valeur = "62.5"
                        Case &HB : valeur = "68.75"
                        Case &HC : valeur = "75"
                        Case &HD : valeur = "81.25"
                        Case &HE : valeur = "87.5"
                        Case &HF : valeur = "93.75"
                        Case &H10 : valeur = "100"
                        Case &H11 : valeur = "ON" 'On ( at last dim level set)"
                        Case &H12 : valeur = "TOGGLE" '"Toggle between Off and On (last dim level set)"
                        Case &H13 : valeur = "BRIGHT" '"Bright one step"
                        Case &H14 : valeur = "DIM" 'Dim one step"
                        Case &H15 : valeur = "START_DIM_CYCLE" 'Start dim cycle"
                        Case &H16 : valeur = "Program(Timer)"
                        Case &H17 : valeur = "Request status from a bidirectional device"
                        Case &H18 : valeur = "OFF for timer period"
                        Case &H19 : valeur = "ON (100%) for timer period"
                        Case &H1A : valeur = "ON ( at last dim level set) for timer period"
                        Case &H1B : valeur = "RESET"
                        Case Else : valeur = "ERROR: Unknown"
                    End Select
                    'If (recbuf(FS20.cmd1) And &H80) = 0 Then
                    '    WriteMessage("                command to receiver")
                    'Else
                    '    WriteMessage("                response from receiver")
                    'End If
                    'If (recbuf(FS20.cmd1) And &H40) = 0 Then
                    '    WriteMessage("                unidirectional command")
                    'Else
                    '    WriteMessage("                bidirectional command")
                    'End If
                    'If (recbuf(FS20.cmd1) And &H20) = 0 Then
                    '    WriteMessage("                additional cmd2 byte not present")
                    'Else
                    '    WriteMessage("                additional cmd2 byte present")
                    'End If
                    'If (recbuf(FS20.cmd1) And &H20) <> 0 Then
                    '    WriteMessage("Cmd2          = " & VB.Right("0" & Hex(recbuf(FS20.cmd2)), 2))
                    'End If
                    WriteRetour(adresse, "", valeur)
                Case FS20.sTypeFHT8V
                    'WriteMessage("subtype       = FHT 8V valve")
                    'WriteMessage("Sequence nbr  = " & recbuf(FS20.seqnbr).ToString)
                    'WriteMessage("House code    = " & VB.Right("0" & Hex(recbuf(FS20.hc1)), 2) & VB.Right("0" & Hex(recbuf(FS20.hc2)), 2))
                    'WriteMessage("Address       = " & VB.Right("0" & Hex(recbuf(FS20.addr)), 2))
                    adresse = VB.Right("0" & Hex(recbuf(FS20.hc1)), 2) & VB.Right("0" & Hex(recbuf(FS20.hc2)), 2) & VB.Right("0" & Hex(recbuf(FS20.addr)), 2)
                    'If (recbuf(FS20.cmd1) And &H80) = 0 Then
                    '    WriteMessage("new command")
                    'Else
                    '    WriteMessage("repeated command")
                    'End If
                    'If (recbuf(FS20.cmd1) And &H40) = 0 Then
                    '    WriteMessage("                unidirectional command")
                    'Else
                    '    WriteMessage("                bidirectional command")
                    'End If
                    'If (recbuf(FS20.cmd1) And &H20) = 0 Then
                    '    WriteMessage("                additional cmd2 byte not present")
                    'Else
                    '    WriteMessage("                additional cmd2 byte present")
                    'End If
                    'If (recbuf(FS20.cmd1) And &H10) = 0 Then
                    '    WriteMessage("                battery empty beep not enabled")
                    'Else
                    '    WriteMessage("                enable battery empty beep")
                    'End If
                    Select Case (recbuf(FS20.cmd1) And &HF)
                        Case &H0 : valeur = "Synchronize now : valve position: " & VB.Right("0" & Hex(recbuf(FS20.cmd2)), 2) & " is " & (CInt(recbuf(FS20.cmd2) / 2.55)).ToString & "%"
                        Case &H1 : valeur = "open valve"
                        Case &H2 : valeur = "close valve"
                        Case &H6 : valeur = "open valve at percentage level : valve position: " & VB.Right("0" & Hex(recbuf(FS20.cmd2)), 2) & " is " & (CInt(recbuf(FS20.cmd2) / 2.55)).ToString & "%"
                        Case &H8 : valeur = "relative offset (cmd2 bit 7=direction, bit 5-0 offset value)"
                        Case &HA : valeur = "decalcification cycle : valve position: " & VB.Right("0" & Hex(recbuf(FS20.cmd2)), 2) & " is " & (CInt(recbuf(FS20.cmd2) / 2.55)).ToString & "%"
                        Case &HC : valeur = "synchronization active : count down is " & (recbuf(FS20.cmd2) >> 1).ToString & " seconds"
                        Case &HE : valeur = "test, drive valve and produce an audible signal"
                        Case &HF : valeur = "pair valve (cmd2 bit 7-1 is count down in seconds, bit 0=1) : count down is " & CStr(recbuf(FS20.cmd2) >> 1) & " seconds"
                        Case Else : valeur = "ERROR: Unknown"
                    End Select
                    WriteRetour(adresse, "", valeur)
                Case FS20.sTypeFHT80
                    'WriteMessage("subtype       = FHT80 door/window sensor")
                    'WriteMessage("Sequence nbr  = " & recbuf(FS20.seqnbr).ToString)
                    'WriteMessage("House code    = " & VB.Right("0" & Hex(recbuf(FS20.hc1)), 2) & VB.Right("0" & Hex(recbuf(FS20.hc2)), 2))
                    'WriteMessage("Address       = " & VB.Right("0" & Hex(recbuf(FS20.addr)), 2))
                    adresse = VB.Right("0" & Hex(recbuf(FS20.hc1)), 2) & VB.Right("0" & Hex(recbuf(FS20.hc2)), 2) & VB.Right("0" & Hex(recbuf(FS20.addr)), 2)
                    Select Case (recbuf(FS20.cmd1) And &HF)
                        Case &H1 : valeur = "OPEN"
                        Case &H2 : valeur = "CLOSE"
                        Case &HC : valeur = "CFG: synchronization active"
                        Case Else : valeur = "ERROR: Unknown command"
                    End Select
                    'If (recbuf(FS20.cmd1) And &H80) = 0 Then
                    '    WriteMessage("                new command")
                    'Else
                    '    WriteMessage("                repeated command")
                    'End If
                    WriteRetour(adresse, "", valeur)
                Case Else : WriteLog("ERR: Unknown Sub type for Packet type=" & Hex(recbuf(FS20.packettype)) & ":" & Hex(recbuf(FS20.subtype)))
            End Select
            If _DEBUG Then WriteLog("DBG: Signal Level : " & (recbuf(FS20.rssi) >> 4).ToString & " (Adresse:" & adresse & ")")
        Catch ex As Exception
            WriteLog("ERR: decode_FS20 Exception : " & ex.Message)
        End Try
    End Sub

    Private Sub decode_RAW()
        'decoding of this type is only implemented for use by simulate and verbose
        Try
            Dim adresse As String = ""
            Dim valeur As String = ""
            Select Case recbuf(RAW.subtype)
                Case RAW.sTypeRAW
                    'WriteMessage("Packet Length = " & recbuf(RAW.packetlength).ToString)
                    'WriteMessage("subtype       = RAW transmit")
                    'WriteMessage("Sequence nbr  = " & recbuf(RAW.seqnbr).ToString)
                    'WriteMessage("Repeat        = " & recbuf(RAW.repeat).ToString)
                    WriteLog("decode_RAW: Packet Length = " & recbuf(RAW.packetlength).ToString & ", subtype = RAW transmit, Repeat = " & recbuf(RAW.repeat).ToString)
                Case Else : WriteLog("ERR: decode_RAW : Unknown Sub type for Packet type=" & Hex(recbuf(RAW.packettype)) & ": " & Hex(recbuf(RAW.subtype)))
            End Select
        Catch ex As Exception
            WriteLog("ERR: decode_RAW Exception : " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Send messages"

    Private Sub SetMode2()
        Try
            Dim temp As String = ""

            Dim kar(ICMD.size) As Byte
            kar(ICMD.packetlength) = ICMD.size
            kar(ICMD.packettype) = ICMD.pTypeInterfaceControl
            kar(ICMD.subtype) = ICMD.sTypeInterfaceCommand
            kar(ICMD.seqnbr) = bytSeqNbr
            kar(ICMD.cmnd) = ICMD.cmdSETMODE

            'type frequence
            Select Case _PARAMMODE_1_frequence
                Case 0 : kar(ICMD.msg1) = IRESPONSE.recType310
                Case 1 : kar(ICMD.msg1) = IRESPONSE.recType315
                Case 2 : kar(ICMD.msg1) = IRESPONSE.recType43392
                Case 3 : kar(ICMD.msg1) = IRESPONSE.recType86830
                Case 4 : kar(ICMD.msg1) = IRESPONSE.recType86830FSK
                Case 5 : kar(ICMD.msg1) = IRESPONSE.recType86835
                Case 6 : kar(ICMD.msg1) = IRESPONSE.recType86835FSK
                Case 7 : kar(ICMD.msg1) = IRESPONSE.recType86895
            End Select

            'UNDEC
            If _PARAMMODE_2_undec = 1 Then kar(ICMD.msg3) = &H80 Else kar(ICMD.msg3) = 0

            kar(ICMD.msg4) = 0

            If _PARAMMODE_18_blindst0 = 1 Then
                'All other protocol receiving is disabled if BlindsT0 enabled, BlindsT0 receive is only used to read the address code of the remote
                kar(ICMD.msg4) = IRESPONSE.msg4_BlindsT0
                kar(ICMD.msg5) = 0
            Else
                If _PARAMMODE_19_Imagintronix = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_IMAGINTRONIX
                If _PARAMMODE_20_sx = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_SX
                If _PARAMMODE_21_rsl = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_RSL
                If _PARAMMODE_22_lighting4 = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_LIGHTING4
                If _PARAMMODE_23_fineoffset = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_FINEOFFSET
                If _PARAMMODE_24_rubicson = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_RUBICSON
                If _PARAMMODE_25_ae = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_AE

                If _PARAMMODE_26_blindst1 = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_BlindsT1
                If _PARAMMODE_4_proguard = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_PROGUARD
                If _PARAMMODE_5_fs20 = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_FS20
                If _PARAMMODE_6_lacrosse = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_LCROS
                If _PARAMMODE_7_hideki = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_HID
                If _PARAMMODE_8_ad = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_AD
                If _PARAMMODE_9_mertik = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_MERTIK

                kar(ICMD.msg5) = 0
                If _PARAMMODE_10_visonic = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_VISONIC
                If _PARAMMODE_11_ati = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_ATI
                If _PARAMMODE_12_oregon = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_OREGON
                If _PARAMMODE_13_meiantech = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_MEI
                If _PARAMMODE_14_heeu = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_HEU
                If _PARAMMODE_15_ac = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_AC
                If _PARAMMODE_16_arc = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_ARC
                If _PARAMMODE_17_x10 = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_X10

            End If
            If _DEBUG Then
                For Each bt As Byte In kar
                    temp = temp & VB.Right("0" & Hex(bt), 2) & " "
                Next
                WriteLog("DBG: Setmode : Commande envoyée : " & temp)
            End If
            ecrire(kar)
        Catch ex As Exception
            WriteLog("ERR: SetMode Exception : " & ex.Message)
        End Try
    End Sub
    'Private Sub SetMode(ByVal paramMode As String)
    '    Try
    '        Dim temp As String = ""
    '        'paramMode 20011111111111111011111111  
    '        '1 : type frequence (310, 315, 433, 868.30, 868.30 FSK, 868.35, 868.35 FSK, 868.95)
    '        '2 : UNDEC
    '        '3 : novatis --> NOT USED ANYMORE 200
    '        '4 : proguard
    '        '5 : FS20
    '        '6 : Lacrosse
    '        '7 : Hideki
    '        '8 : AD
    '        '9 : Mertik 111111
    '        '10 : Visonic
    '        '11 : ATI
    '        '12 : Oregon
    '        '13 : Meiantech
    '        '14 : HEEU
    '        '15 : AC
    '        '16 : ARC
    '        '17 : X10 11111111

    '        '18 : BlindsT0
    '        '19 : RFU6
    '        '20 : RFU5
    '        '21 : RFU4
    '        '22 : RFU3 --> LIGHTING4
    '        '23 : FINEOFFSET
    '        '24 : RUBICSON
    '        '25 : AE
    '        '26 : BlindsT1

    '        If paramMode.Length <> 26 Then
    '            WriteLog("ERR: Setmode : ParamMode incorrect : " & paramMode & " (valeur par défaut utilisée : 20011111111111111011111111)")
    '            paramMode = "20011111111111111011111111"
    '        Else
    '            WriteLog("DBG: Setmode : ParamMode utilisé : " & paramMode)
    '        End If


    '        Dim kar(ICMD.size) As Byte
    '        kar(ICMD.packetlength) = ICMD.size
    '        kar(ICMD.packettype) = ICMD.pTypeInterfaceControl
    '        kar(ICMD.subtype) = ICMD.sTypeInterfaceCommand
    '        kar(ICMD.seqnbr) = bytSeqNbr
    '        kar(ICMD.cmnd) = ICMD.cmdSETMODE

    '        'type frequence
    '        Select Case CInt(paramMode.Substring(0, 1))
    '            Case 0 : kar(ICMD.msg1) = IRESPONSE.recType310
    '            Case 1 : kar(ICMD.msg1) = IRESPONSE.recType315
    '            Case 2 : kar(ICMD.msg1) = IRESPONSE.recType43392
    '            Case 3 : kar(ICMD.msg1) = IRESPONSE.recType86830
    '            Case 4 : kar(ICMD.msg1) = IRESPONSE.recType86830FSK
    '            Case 5 : kar(ICMD.msg1) = IRESPONSE.recType86835
    '            Case 6 : kar(ICMD.msg1) = IRESPONSE.recType86835FSK
    '            Case 7 : kar(ICMD.msg1) = IRESPONSE.recType86895
    '        End Select

    '        'UNDEC
    '        If CInt(paramMode.Substring(1, 1)) = 1 Then kar(ICMD.msg3) = &H80 Else kar(ICMD.msg3) = 0

    '        If CInt(paramMode.Substring(17, 1)) = 1 Then
    '            'All other protocol receiving is disabled if BlindsT0 enabled, BlindsT0 receive is only used to read the address code of the remote
    '            kar(ICMD.msg4) = 0
    '            kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_BlindsT0
    '            kar(ICMD.msg5) = 0
    '        Else
    '            If CInt(paramMode.Substring(18, 1)) = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_RFU6
    '            If CInt(paramMode.Substring(19, 1)) = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_RFU5
    '            If CInt(paramMode.Substring(20, 1)) = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_RFU4
    '            If CInt(paramMode.Substring(21, 1)) = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_LIGHTING4
    '            If CInt(paramMode.Substring(22, 1)) = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_FINEOFFSET
    '            If CInt(paramMode.Substring(23, 1)) = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_RUBICSON
    '            If CInt(paramMode.Substring(24, 1)) = 1 Then kar(ICMD.msg3) = kar(ICMD.msg3) Or IRESPONSE.msg3_AE

    '            kar(ICMD.msg4) = 0
    '            If CInt(paramMode.Substring(25, 1)) = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_BlindsT1
    '            If CInt(paramMode.Substring(3, 1)) = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_PROGUARD
    '            If CInt(paramMode.Substring(4, 1)) = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_FS20
    '            If CInt(paramMode.Substring(5, 1)) = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_LCROS
    '            If CInt(paramMode.Substring(6, 1)) = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_HID
    '            If CInt(paramMode.Substring(7, 1)) = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_AD
    '            If CInt(paramMode.Substring(8, 1)) = 1 Then kar(ICMD.msg4) = kar(ICMD.msg4) Or IRESPONSE.msg4_MERTIK

    '            kar(ICMD.msg5) = 0
    '            If CInt(paramMode.Substring(9, 1)) = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_VISONIC
    '            If CInt(paramMode.Substring(10, 1)) = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_ATI
    '            If CInt(paramMode.Substring(11, 1)) = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_OREGON
    '            If CInt(paramMode.Substring(12, 1)) = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_MEI
    '            If CInt(paramMode.Substring(13, 1)) = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_HEU
    '            If CInt(paramMode.Substring(14, 1)) = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_AC
    '            If CInt(paramMode.Substring(15, 1)) = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_ARC
    '            If CInt(paramMode.Substring(16, 1)) = 1 Then kar(ICMD.msg5) = kar(ICMD.msg5) Or IRESPONSE.msg5_X10

    '        End If
    '        If _DEBUG Then
    '            For Each bt As Byte In kar
    '                temp = temp & VB.Right("0" & Hex(bt), 2) & " "
    '            Next
    '            WriteLog("DBG: Setmode : Commande envoyée : " & temp)
    '        End If
    '        ecrire(kar)
    '    Catch ex As Exception
    '        WriteLog("ERR: SetMode Exception : " & ex.Message)
    '    End Try
    'End Sub

    ''' <summary>Converti un house de type A, B... en byte</summary>
    ''' <param name="housecode">HouseCode du type A (de A1)</param>
    ''' <returns>Byte représentant le housecode</returns>
    Private Function convert_housecode(ByVal housecode As String) As Byte
        Try
            Dim temp As Byte
            Select Case housecode
                Case "A" : temp = 0 + &H41
                Case "B" : temp = 1 + &H41
                Case "C" : temp = 2 + &H41
                Case "D" : temp = 3 + &H41
                Case "E" : temp = 4 + &H41
                Case "F" : temp = 5 + &H41
                Case "G" : temp = 6 + &H41
                Case "H" : temp = 7 + &H41
                Case "I" : temp = 8 + &H41
                Case "J" : temp = 9 + &H41
                Case "K" : temp = 10 + &H41
                Case "L" : temp = 11 + &H41
                Case "M" : temp = 12 + &H41
                Case "N" : temp = 13 + &H41
                Case "O" : temp = 14 + &H41
                Case "P" : temp = 15 + &H41
                Case Else : WriteLog("ERR: convert_housecode HouseCode Incorrect : " & housecode)
            End Select
            Return temp
        Catch ex As Exception
            WriteLog("ERR: convert_housecode Exception : " & ex.Message)
            Return 0 + &H41
        End Try
    End Function

    ''' <summary>Converti un id... en byte (Blyss)</summary>
    ''' <param name="id">id (de type A1)</param>
    ''' <returns>Byte représentant le housecode</returns>
    Private Function convert_id(ByVal id As String) As Byte
        Try
            Dim t(2), temp As Byte


            Select Case id.Substring(0, 1)
                Case "0" : t(0) = &H0
                Case "1" : t(0) = &H10
                Case "2" : t(0) = &H20
                Case "3" : t(0) = &H30
                Case "4" : t(0) = &H40
                Case "5" : t(0) = &H50
                Case "6" : t(0) = &H60
                Case "7" : t(0) = &H70
                Case "8" : t(0) = &H80
                Case "9" : t(0) = &H90
                Case "A" : t(0) = &HA0
                Case "B" : t(0) = &HB0
                Case "C" : t(0) = &HC0
                Case "D" : t(0) = &HD0
                Case "E" : t(0) = &HE0
                Case "F" : t(0) = &HF0
                Case Else : WriteLog("ERR: convert_id Id Incorrect : " & id)
            End Select

            Select Case id.Substring(1, 1)
                Case "0" : t(1) = &H0
                Case "1" : t(1) = &H1
                Case "2" : t(1) = &H2
                Case "3" : t(1) = &H3
                Case "4" : t(1) = &H4
                Case "5" : t(1) = &H5
                Case "6" : t(1) = &H6
                Case "7" : t(1) = &H7
                Case "8" : t(1) = &H8
                Case "9" : t(1) = &H9
                Case "A" : t(1) = &HA
                Case "B" : t(1) = &HB
                Case "C" : t(1) = &HC
                Case "D" : t(1) = &HD
                Case "E" : t(1) = &HE
                Case "F" : t(1) = &HF
                Case Else : WriteLog("ERR: convert_id Id Incorrect : " & id)
            End Select
            temp = (t(0) Or t(1))
            WriteLog("DBG: INPUT: " & id & ", OUTPUT : " & t(0) & " || " & t(1) & " = " & temp)

            Return (temp)
        Catch ex As Exception
            WriteLog("ERR: convert_id Exception : " & ex.Message)
            Return 0 + &H41
        End Try
    End Function

    ''' <summary>Gestion du protocole X10/ARC/ELRO AB400D/Waveman/EMW200/Impuls/RisingSun/Philips SBC</summary>
    ''' <param name="adresse">Adresse du type A1</param>
    ''' <param name="commande">commande ON, OFF, BRIGHT, DIM, ALL_LIGHT_ON, ALL_LIGHT_OFF</param>
    ''' <param name="type">0=X10, 1=ARC, 2=ELRO AB400D, 3=Waveman, 4=EMW200, 5=Impuls, 6=RisingSun, 7=Philips SBC</param>
    ''' <remarks></remarks>
    Private Sub send_lighting1(ByVal adresse As String, ByVal commande As String, ByVal type As Integer)
        Try
            Dim kar(LIGHTING1.size) As Byte
            Dim temp As String = ""

            'verification format adresse
            If Not (adresse.Length = 2 Or adresse.Length = 3) Then
                WriteLog("ERR: Send lighting1 : Adresse invalide : " & adresse)
                Exit Sub
            End If
            Select Case type
                Case LIGHTING1.sTypeX10, LIGHTING1.sTypeARC, LIGHTING1.sTypeWaveman
                    If CInt(adresse.Substring(1, adresse.Length - 1)) > 16 Or (adresse.Substring(0, 1) <> "A" And adresse.Substring(0, 1) <> "B" And adresse.Substring(0, 1) <> "C" And adresse.Substring(0, 1) <> "D" And adresse.Substring(0, 1) <> "E" And adresse.Substring(0, 1) <> "F" And adresse.Substring(0, 1) <> "G" And adresse.Substring(0, 1) <> "H" And adresse.Substring(0, 1) <> "I" And adresse.Substring(0, 1) <> "J" And adresse.Substring(0, 1) <> "K" And adresse.Substring(0, 1) <> "L" And adresse.Substring(0, 1) <> "M" And adresse.Substring(0, 1) <> "N" And adresse.Substring(0, 1) <> "O" And adresse.Substring(0, 1) <> "P") Then
                        WriteLog("ERR: Send lighting1 : Adresse (X10, ARC, Waveman) invalide : " & adresse)
                        Exit Sub
                    End If
                Case LIGHTING1.sTypeAB400D, LIGHTING1.sTypeIMPULS
                    If CInt(adresse.Substring(1, adresse.Length - 1)) > 64 Or (adresse.Substring(0, 1) <> "A" And adresse.Substring(0, 1) <> "B" And adresse.Substring(0, 1) <> "C" And adresse.Substring(0, 1) <> "D" And adresse.Substring(0, 1) <> "E" And adresse.Substring(0, 1) <> "F" And adresse.Substring(0, 1) <> "G" And adresse.Substring(0, 1) <> "H" And adresse.Substring(0, 1) <> "I" And adresse.Substring(0, 1) <> "J" And adresse.Substring(0, 1) <> "K" And adresse.Substring(0, 1) <> "L" And adresse.Substring(0, 1) <> "M" And adresse.Substring(0, 1) <> "N" And adresse.Substring(0, 1) <> "O" And adresse.Substring(0, 1) <> "P") Then
                        WriteLog("ERR: Send lighting1 : Adresse (ELRO AB400D, Impuls) invalide : " & adresse)
                        Exit Sub
                    End If
                Case LIGHTING1.sTypeEMW200
                    If CInt(adresse.Substring(1, adresse.Length - 1)) > 4 Or (adresse.Substring(0, 1) <> "A" And adresse.Substring(0, 1) <> "B" And adresse.Substring(0, 1) <> "C") Then
                        WriteLog("ERR: Send lighting1 : Adresse (EMW200) invalide : " & adresse)
                        Exit Sub
                    End If
                Case LIGHTING1.sTypePhilips
                    If CInt(adresse.Substring(1, adresse.Length - 1)) > 8 Or (adresse.Substring(0, 1) <> "A" And adresse.Substring(0, 1) <> "B" And adresse.Substring(0, 1) <> "C" And adresse.Substring(0, 1) <> "D" And adresse.Substring(0, 1) <> "E" And adresse.Substring(0, 1) <> "F" And adresse.Substring(0, 1) <> "G" And adresse.Substring(0, 1) <> "H" And adresse.Substring(0, 1) <> "I" And adresse.Substring(0, 1) <> "J" And adresse.Substring(0, 1) <> "K" And adresse.Substring(0, 1) <> "L" And adresse.Substring(0, 1) <> "M" And adresse.Substring(0, 1) <> "N" And adresse.Substring(0, 1) <> "O" And adresse.Substring(0, 1) <> "P") Then
                        WriteLog("ERR: Send lighting1 : Adresse (ELRO AB400D, Impuls) invalide : " & adresse)
                        Exit Sub
                    End If
            End Select

            kar(LIGHTING1.packetlength) = LIGHTING1.size
            kar(LIGHTING1.packettype) = LIGHTING1.pTypeLighting1
            kar(LIGHTING1.subtype) = CByte(type)
            kar(LIGHTING1.seqnbr) = bytSeqNbr
            kar(LIGHTING1.housecode) = convert_housecode(adresse.Substring(0, 1))
            kar(LIGHTING1.unitcode) = CByte(adresse.Substring(1, adresse.Length - 1))

            Select Case commande
                Case "OFF"
                    kar(LIGHTING1.cmnd) = LIGHTING1.sOff
                    WriteRetourSend(adresse, "", "OFF")
                Case "ON"
                    kar(LIGHTING1.cmnd) = LIGHTING1.sOn
                    WriteRetourSend(adresse, "", "ON")
                Case "ALL_LIGHT_OFF"
                    If type = LIGHTING1.sTypeX10 Or type = LIGHTING1.sTypeEMW200 Or type = LIGHTING1.sTypePhilips Or type = LIGHTING1.sTypeARC Then
                        kar(LIGHTING1.cmnd) = LIGHTING1.sAllOff
                        WriteRetourSend(adresse, "", "OFF")
                        'traiter toutes les lights


                    Else
                        WriteLog("ERR: Send lighting1 : Commande ALL_LIGHT_OFF indisponible pour ce protocole")
                        Exit Sub
                    End If
                Case "ALL_LIGHT_ON"
                    If type = LIGHTING1.sTypeX10 Or type = LIGHTING1.sTypeEMW200 Or type = LIGHTING1.sTypePhilips Or type = LIGHTING1.sTypeARC Then
                        kar(LIGHTING1.cmnd) = LIGHTING1.sAllOn
                        WriteRetourSend(adresse, "", "ON")
                        'traiter toutes les lights


                    Else
                        WriteLog("ERR: Send lighting1 : Commande ALL_LIGHT_ON indisponible pour ce protocole")
                        Exit Sub
                    End If
                Case "BRIGHT"
                    If type = LIGHTING1.sTypeX10 Then
                        kar(LIGHTING1.cmnd) = LIGHTING1.sBright
                        WriteRetourSend(adresse, "", "ON")
                    Else
                        WriteLog("ERR: Send lighting1 : Commande BRIGHT indisponible pour ce protocole")
                        Exit Sub
                    End If
                Case "DIM"
                    If type = LIGHTING1.sTypeX10 Then
                        kar(LIGHTING1.cmnd) = LIGHTING1.sDim
                        WriteRetourSend(adresse, "", "ON")
                    Else
                        WriteLog("ERR: Send lighting1 : Commande DIM indisponible pour ce protocole")
                        Exit Sub
                    End If
                Case "CHIME"
                    If type = LIGHTING1.sTypeARC Then
                        kar(LIGHTING1.cmnd) = LIGHTING1.sChime
                        kar(LIGHTING1.unitcode) = 8
                        WriteRetourSend(adresse, "", "CHIME")
                    Else
                        WriteLog("ERR: Send lighting1 : Commande CHIME indisponible pour ce protocole")
                        Exit Sub
                    End If
                Case Else
                    WriteLog("ERR: Send lighting1 : Commande invalide : " & commande)
                    Exit Sub
            End Select
            kar(LIGHTING1.filler) = 0
            ecrire(kar)
            If _DEBUG Then
                For Each bt As Byte In kar
                    temp = temp & VB.Right("0" & Hex(bt), 2) & " "
                Next
                WriteLog("DBG: Send lighting1 : commande envoyée : " & temp)
            End If
        Catch ex As Exception
            WriteLog("ERR: Send lighting1 Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Gestion du protocole AC / HEEU / ANSLUT</summary>
    ''' <param name="adresse">Adresse du type 02F4416-1</param>
    ''' <param name="commande">commande ON, OFF, DIM, GROUP_OFF, GROUP_ON, GROUP_DIM</param>
    ''' <param name="type">0=AC / 1=HEEU / 2=ANSLUT</param>
    ''' <param name="dimlevel">Level pour Dim de 1 à 16</param>
    ''' <remarks></remarks>
    Private Sub send_Lighting2(ByVal adresse As String, ByVal commande As String, ByVal type As Integer, Optional ByVal dimlevel As Integer = 1)
        Try
            Dim kar(LIGHTING2.size) As Byte
            Dim temp As String = ""
            If Not (adresse.Length = 9 Or adresse.Length = 10) Then
                WriteLog("ERR: Send Lighting2 : Adresse invalide : " & adresse)
                Exit Sub
            End If
            kar(LIGHTING2.packetlength) = LIGHTING2.size
            kar(LIGHTING2.packettype) = LIGHTING2.pTypeLighting2
            kar(LIGHTING2.seqnbr) = bytSeqNbr
            kar(LIGHTING2.subtype) = CByte(type)
            '  WriteLog("DBG: Send AC : Commande Envoyée : " & commande)
            Select Case commande
                Case "OFF"
                    kar(LIGHTING2.cmnd) = 0
                    WriteRetourSend(adresse, "", "OFF")
                Case "ON"
                    kar(LIGHTING2.cmnd) = 1
                    WriteRetourSend(adresse, "", "ON")
                Case "GROUP_OFF"
                    kar(LIGHTING2.cmnd) = 3
                    WriteRetourSend(adresse, "", "OFF")
                    'traiter toutes le groupe



                Case "GROUP_ON"
                    kar(LIGHTING2.cmnd) = 4
                    WriteRetourSend(adresse, "", "ON")
                    'traiter toutes le groupe



                Case "GROUP_DIM"
                    kar(LIGHTING2.cmnd) = 5
                    WriteRetourSend(adresse, "", CStr(dimlevel))
                    'traiter toutes le groupe



                Case "DIM"
                    kar(LIGHTING2.cmnd) = 2
                    WriteRetourSend(adresse, "", CStr(dimlevel))

                Case "OUVERTURE"
                    If dimlevel = 0 Then
                        kar(LIGHTING2.cmnd) = 0
                        WriteRetourSend(adresse, "", "OFF")
                    Else
                        kar(LIGHTING2.cmnd) = 1
                        WriteRetourSend(adresse, "", "ON")
                    End If
                Case Else
                    WriteLog("ERR: Send AC : Commande invalide : " & commande)
                    Exit Sub
            End Select
            Try
                Dim adressetab As String() = adresse.Split(CChar("-"))
                kar(LIGHTING2.unitcode) = CByte(adressetab(1))
                kar(LIGHTING2.id1) = CByte(adressetab(0).Substring(0, 1))
                kar(LIGHTING2.id2) = CByte(Array.IndexOf(adressetoint, adressetab(0).Substring(1, 2)))
                kar(LIGHTING2.id3) = CByte(Array.IndexOf(adressetoint, adressetab(0).Substring(3, 2)))
                kar(LIGHTING2.id4) = CByte(Array.IndexOf(adressetoint, adressetab(0).Substring(5, 2)))
            Catch ex As Exception
                WriteLog("ERR: Send Lighting2 Exception : Adresse incorrecte")
            End Try

            If dimlevel = 0 Then
                dimlevel = 0
            ElseIf dimlevel < 7 Then
                dimlevel = 1
            ElseIf dimlevel < 14 Then
                dimlevel = 2
            ElseIf dimlevel < 21 Then
                dimlevel = 3
            ElseIf dimlevel < 28 Then
                dimlevel = 4
            ElseIf dimlevel < 34 Then
                dimlevel = 5
            ElseIf dimlevel < 40 Then
                dimlevel = 6
            ElseIf dimlevel < 46 Then
                dimlevel = 7
            ElseIf dimlevel < 53 Then
                dimlevel = 8
            ElseIf dimlevel < 60 Then
                dimlevel = 9
            ElseIf dimlevel < 67 Then
                dimlevel = 10
            ElseIf dimlevel < 74 Then
                dimlevel = 11
            ElseIf dimlevel < 81 Then
                dimlevel = 12
            ElseIf dimlevel < 88 Then
                dimlevel = 13
            ElseIf dimlevel < 95 Then
                dimlevel = 14
            Else
                dimlevel = 15
            End If
            kar(LIGHTING2.level) = CByte(dimlevel)
            kar(LIGHTING2.filler) = 0

            ecrire(kar)
            If _DEBUG Then
                For Each bt As Byte In kar
                    temp = temp & VB.Right("0" & Hex(bt), 2) & " "
                Next
                WriteLog("DBG: Send Lighting2 : commande envoyée : " & temp)
            End If
        Catch ex As Exception
            WriteLog("ERR: Send Lighting2 Exception : " & ex.ToString)
        End Try
    End Sub


    ''' <summary>Gestion du protocole EMW100 / LIGHTWAVERF</summary>
    ''' <param name="adresse">Adresse du type FFFFFF-1</param>
    ''' <param name="commande">commande ON, OFF, ALL_LIGHT_ON, ALL_LIGHT_OFF</param>
    ''' <param name="type">0=EMW100 / 1=LIGHTWAVERF</param>
    ''' <param name="dimlevel">Level pour Dim de 1 à 31</param>
    ''' <remarks></remarks>
    Private Sub send_lighting5(ByVal adresse As String, ByVal commande As String, ByVal type As Integer, Optional ByVal dimlevel As Integer = 1)
        Try
            Dim kar(LIGHTING5.size) As Byte
            Dim temp As String = ""

            If Not (adresse.Length = 8) Then
                WriteLog("ERR: Send lighting5 : Adresse invalide : " & adresse)
                Exit Sub
            End If

            kar(LIGHTING5.packetlength) = LIGHTING5.size
            kar(LIGHTING5.packettype) = LIGHTING5.pTypeLighting5
            kar(LIGHTING5.subtype) = CByte(type) '0=EMW100 / 1=LIGHTWAVERF
            kar(LIGHTING2.seqnbr) = bytSeqNbr
            Try
                Dim adressetab As String() = adresse.Split(CChar("-"))
                kar(LIGHTING5.unitcode) = CByte(adressetab(1))
                kar(LIGHTING5.id1) = CByte(adressetab(0).Substring(0, 2))
                kar(LIGHTING5.id2) = CByte(Array.IndexOf(adressetoint, adressetab(0).Substring(2, 2)))
                kar(LIGHTING5.id3) = CByte(Array.IndexOf(adressetoint, adressetab(0).Substring(4, 2)))
            Catch ex As Exception
                WriteLog("ERR: Send lighting5 Exception : Adresse incorrecte")
            End Try

            Select Case commande
                Case "OFF"
                    kar(LIGHTING5.cmnd) = LIGHTING5.sOff
                    WriteRetourSend(adresse, "", "OFF")
                Case "ON"
                    kar(LIGHTING5.cmnd) = LIGHTING5.sOn
                    WriteRetourSend(adresse, "", "ON")
                Case "GROUP_OFF"
                    kar(LIGHTING5.cmnd) = LIGHTING5.sGroupOff
                    WriteRetourSend(adresse, "", "OFF")
                    'traiter tout le groupe

                    'commandes restantes à dev pour EMW100 : Mood1/Mood2/Mood3/Mood4/Mood5/Unlock/Lock/All Lock/Close relay/Stop relay/Open relay/Set Level
                    'commandes restantes à dev pour LIGHTWAVERF : Learn

                Case Else
                    WriteLog("ERR: Send lighting5 : Commande invalide : " & commande)
                    Exit Sub
            End Select

            If kar(LIGHTING5.cmnd) = LIGHTING5.sSetLevel Then
                If dimlevel = 0 Then
                    dimlevel = 0
                ElseIf dimlevel < 7 Then
                    dimlevel = 1
                ElseIf dimlevel < 14 Then
                    dimlevel = 2
                ElseIf dimlevel < 21 Then
                    dimlevel = 3
                ElseIf dimlevel < 28 Then
                    dimlevel = 4
                ElseIf dimlevel < 34 Then
                    dimlevel = 5
                ElseIf dimlevel < 40 Then
                    dimlevel = 6
                ElseIf dimlevel < 46 Then
                    dimlevel = 7
                ElseIf dimlevel < 53 Then
                    dimlevel = 8
                ElseIf dimlevel < 60 Then
                    dimlevel = 9
                ElseIf dimlevel < 67 Then
                    dimlevel = 10
                ElseIf dimlevel < 74 Then
                    dimlevel = 11
                ElseIf dimlevel < 81 Then
                    dimlevel = 12
                ElseIf dimlevel < 88 Then
                    dimlevel = 13
                ElseIf dimlevel < 95 Then
                    dimlevel = 14
                Else
                    dimlevel = 15
                End If
                kar(LIGHTING5.level) = CByte(dimlevel)
            Else
                kar(LIGHTING5.level) = 0
            End If
            kar(LIGHTING5.filler) = 0
            If kar(LIGHTING5.cmnd) = 8 Or kar(LIGHTING5.cmnd) = 9 Then 'not used commands
                Exit Sub
            End If
            If kar(LIGHTING5.id1) = 0 And kar(LIGHTING5.id2) = 0 And kar(LIGHTING5.id3) = 0 Then
                WriteLog("ERR: Send lighting5 : Adresse invalide : " & adresse)
                Exit Sub
            End If

            ecrire(kar)
            If _DEBUG Then
                For Each bt As Byte In kar
                    temp = temp & VB.Right("0" & Hex(bt), 2) & " "
                Next
                WriteLog("DBG: Send lighting5 : commande envoyée : " & temp)
            End If
        Catch ex As Exception
            WriteLog("ERR: Send lighting5 Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Gestion du protocole BLYSS</summary>
    ''' <param name="adresse">Adresse du type FFFF-A1  avec A1 A/P et 1/4</param>
    ''' <param name="commande">commande ON, OFF</param>
    ''' <param name="type">0=BLYSS</param>
    ''' <remarks></remarks>
    Private Sub send_lighting6(ByVal adresse As String, ByVal commande As String, ByVal type As Integer)
        Try
            Dim kar(LIGHTING6.size) As Byte
            Dim temp As String = ""

            If Not (adresse.Length = 7) Then
                WriteLog("ERR: Send lighting6 : Adresse invalide : " & adresse)
                Exit Sub
            End If

            kar(LIGHTING6.packetlength) = LIGHTING6.size
            kar(LIGHTING6.packettype) = LIGHTING6.pTypeLighting6
            kar(LIGHTING6.subtype) = 0 '0=BLYSS
            kar(LIGHTING6.seqnbr) = bytSeqNbr
            Try
                Dim adressetab As String() = adresse.Split(CChar("-"))
                kar(LIGHTING6.groupcode) = convert_housecode(adressetab(1).Substring(0, 1))
                kar(LIGHTING6.unitcode) = CByte(adressetab(1).Substring(1, 1))
                'on repasse pour id1/2 à l'ancienne methode car la nouvelle ne marche pas de façon trés stable à priori
                'kar(LIGHTING6.id1) = CByte(adressetab(0).Substring(0, 2))
                'kar(LIGHTING6.id2) = CByte(Array.IndexOf(adressetoint, adressetab(0).Substring(2, 2)))
                'kar(LIGHTING6.id1) = convert_id(adressetab(0).Substring(0, 2))
                'kar(LIGHTING6.id2) = convert_id(adressetab(0).Substring(2, 2))
                kar(LIGHTING6.id1) = CByte(Array.IndexOf(adressetoint, adressetab(0).Substring(0, 2)))
                kar(LIGHTING6.id2) = CByte(Array.IndexOf(adressetoint, adressetab(0).Substring(2, 2)))

            Catch ex As Exception
                WriteLog("ERR: Send lighting6 Exception : Adresse incorrecte")
            End Try

            Select Case commande
                Case "OFF"
                    kar(LIGHTING6.cmnd) = LIGHTING6.sOff
                    WriteRetourSend(adresse, "", "OFF")
                Case "ON"
                    kar(LIGHTING6.cmnd) = LIGHTING6.sOn
                    WriteRetourSend(adresse, "", "ON")
                Case "GROUP_OFF"
                    kar(LIGHTING6.cmnd) = LIGHTING6.sGroupOff
                    WriteRetourSend(adresse, "", "OFF")
                Case "GROUP_ON"
                    kar(LIGHTING6.cmnd) = LIGHTING6.sGroupOff
                    WriteRetourSend(adresse, "", "OFF")
                Case Else
                    WriteLog("ERR: Send lighting6 : Commande invalide : " & commande)
                    Exit Sub
            End Select

            ' --> bytCmndSeqNbr = "0", "1", "2", "3", "4"
            kar(LIGHTING6.cmndseqnbr) = bytCmndSeqNbr
            bytCmndSeqNbr += 1
            If bytCmndSeqNbr > 4 Then
                bytCmndSeqNbr = 0
            End If

            ' --> bytCmndSeqNbr = "0" à "145"
            kar(LIGHTING6.seqnbr2) = bytCmndSeqNbr2
            If bytCmndSeqNbr2 <> 0 Then
                If bytCmndSeqNbr2 = 145 Then
                    bytCmndSeqNbr2 = 1
                Else
                    bytCmndSeqNbr2 += 1
                End If
            End If

            kar(LIGHTING6.filler) = 0

            If kar(LIGHTING6.id1) = 0 And kar(LIGHTING6.id2) = 0 Then
                WriteLog("ERR: Send lighting6 : Adresse invalide : " & adresse)
                Exit Sub
            End If

            ecrire(kar)
            If _DEBUG Then
                For Each bt As Byte In kar
                    temp = temp & VB.Right("0" & Hex(bt), 2) & " "
                Next
                WriteLog("DBG: Send lighting6 : commande envoyée : " & temp)
            End If
        Catch ex As Exception
            WriteLog("ERR: Send lighting6 Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Gestion du protocole SECURITY</summary>
    ''' <param name="adresse">Adresse du type FFFFA1 </param>
    ''' <param name="commande">commande PANIC, PAIR, MOTION</param>
    ''' <param name="type">protocole</param>
    ''' <remarks></remarks>
    Private Sub send_security(ByVal adresse As String, ByVal commande As String, ByVal type As Integer)
        Try
            Dim kar(SECURITY1.size) As Byte
            Dim temp As String = ""

            If Not (adresse.Length = 6) Then
                WriteLog("ERR: Send Security : Adresse invalide : " & adresse)
                Exit Sub
            End If

            kar(SECURITY1.packetlength) = SECURITY1.size
            kar(SECURITY1.packettype) = SECURITY1.pTypeSecurity1
            kar(SECURITY1.subtype) = type
            'Select Case type
            '    Case 0 : kar(SECURITY1.subtype) = SECURITY1.sTypeSecX10
            '    Case 1 : kar(SECURITY1.subtype) = SECURITY1.sTypeKD101
            '    Case 2 : kar(SECURITY1.subtype) = SECURITY1.sTypeSA30
            '    Case 3 : kar(SECURITY1.subtype) = SECURITY1.sTypePowercodeSensor
            '    Case 4 : kar(SECURITY1.subtype) = SECURITY1.sTypePowercodeMotion
            '    Case 5 : kar(SECURITY1.subtype) = SECURITY1.sTypePowercodeAux
            '    Case 6 : kar(SECURITY1.subtype) = SECURITY1.sTypeMeiantech
            'End Select


            kar(SECURITY1.seqnbr) = bytSeqNbr
            kar(SECURITY1.id1) = CByte(Array.IndexOf(adressetoint, adresse.Substring(0, 2)))
            kar(SECURITY1.id2) = CByte(Array.IndexOf(adressetoint, adresse.Substring(2, 2)))
            kar(SECURITY1.id3) = CByte(Array.IndexOf(adressetoint, adresse.Substring(4, 2)))
            Select Case commande
                Case "NORMAL" : kar(SECURITY1.status) = 0
                Case "NORMAL_DELAYED" : kar(SECURITY1.status) = 1
                Case "ALARM" : kar(SECURITY1.status) = 2
                Case "ALARM_DELAYED" : kar(SECURITY1.status) = 3
                Case "MOTION" : kar(SECURITY1.status) = 4
                Case "NO_MOTION" : kar(SECURITY1.status) = 5
                Case "PANIC" : kar(SECURITY1.status) = 6
                Case "END_PANIC" : kar(SECURITY1.status) = 7
                Case "ARM_AWAY" : kar(SECURITY1.status) = 9
                Case "ARM_AWAY_DELAYED" : kar(SECURITY1.status) = &HA
                Case "ARM_HOME" : kar(SECURITY1.status) = &HB
                Case "ARM_HOME_DELAYED" : kar(SECURITY1.status) = &HC
                Case "DISARM" : kar(SECURITY1.status) = &HD
                Case "LIGHT1_OFF" : kar(SECURITY1.status) = &H10
                Case "LIGHT1_ON" : kar(SECURITY1.status) = &H11
                Case "LIGHT2_OFF" : kar(SECURITY1.status) = &H12
                Case "LIGHT2_ON" : kar(SECURITY1.status) = &H13
                Case "DARK_DETECTED" : kar(SECURITY1.status) = &H14
                Case "LIGHT_DETECTED" : kar(SECURITY1.status) = &H15
                Case "BATTERY_LOW" : kar(SECURITY1.status) = &H16
                Case "PAIR" : kar(SECURITY1.status) = &H17
                Case "NORMAL_TAMPER" : kar(SECURITY1.status) = &H80
                Case "NORMAL_DELAYED_TAMPER" : kar(SECURITY1.status) = &H81
                Case "ALARM_TAMPER" : kar(SECURITY1.status) = &H82
                Case "ALARM_DELAYED_TAMPER" : kar(SECURITY1.status) = &H83
                Case "MOTION_TAMPER" : kar(SECURITY1.status) = &H84
                Case "NO_MOTION_TAMPER" : kar(SECURITY1.status) = &H85
                Case Else
                    WriteLog("ERR: Send Security : Commande invalide : " & commande)
                    Exit Sub
            End Select
            kar(SECURITY1.filler) = 0
            ecrire(kar)

            WriteRetourSend(adresse, "", commande)

            If _DEBUG Then
                For Each bt As Byte In kar
                    temp = temp & VB.Right("0" & Hex(bt), 2) & " "
                Next
                WriteLog("DBG: Send Security : commande envoyée : " & temp)
            End If
        Catch ex As Exception
            WriteLog("ERR: Send Security Exception : " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Write"

    Private Sub WriteLog(ByVal message As String)
        Try
            'utilise la fonction de base pour loguer un event
            If STRGS.InStr(message, "DBG:") > 0 Then
                _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, "RFXtrx", STRGS.Right(message, message.Length - 5))
            ElseIf STRGS.InStr(message, "ERR:") > 0 Then
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "RFXtrx", STRGS.Right(message, message.Length - 5))
            Else
                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, "RFXtrx", message)
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, "RFXtrx WriteLog", ex.Message)
        End Try
    End Sub

    Private Sub WriteBattery(ByVal adresse As String, ByVal valeur As String)
        Try
            'Forcer le . 
            'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
            'My.Application.ChangeCulture("en-US")

            'log tous les paquets en mode debug
            If _DEBUG And (valeur = "vide" Or valeur = "0") Then WriteLog("DBG: WriteBattery : receive from " & adresse)

            If Not _IsConnect Then Exit Sub 'si on ferme le port on quitte
            If DateTime.Now < DateAdd(DateInterval.Second, 10, dateheurelancement) Then Exit Sub 'on ne traite rien pendant les 10 premieres secondes

            'Recherche si un device affecté
            Dim listedevices As New ArrayList
            'on cherche un composant de type batterie avec la même adresse que le composant, si trouvé, on modifie sa valeur
            listedevices = _Server.ReturnDeviceByAdresse1TypeDriver(_IdSrv, adresse, "BATTERIE", Me._ID, True)
            If IsNothing(listedevices) Then
                WriteLog("ERR: Communication impossible avec le serveur, l'IDsrv est peut être erroné : " & _IdSrv)
                Exit Sub
            End If
            If (listedevices.Count = 1) Then
                'listedevices.Item(0).Value = "Vide"
                listedevices.Item(0).Value = valeur
            Else
                'pas de composant Batterie trouvé avec la même adresse, on va loguer si batterie vide
                If valeur = "vide" Or valeur = "0" Then
                    listedevices = _Server.ReturnDeviceByAdresse1TypeDriver(_IdSrv, adresse, "", Me._ID, True)
                    If (listedevices.Count >= 1) Then
                        'on a trouvé un ou plusieurs composants avec cette adresse, on prend le premier
                        WriteLog("ERR: " & listedevices.Item(0).Name & " (" & adresse & ") : Battery Empty")
                    Else
                        'device pas trouvé
                        'Ajouter la gestion des composants bannis (si dans la liste des composant bannis alors on log en debug sinon onlog device non trouve empty)

                        '**************Modif JPS************************
                        ' Vérifie que le composant n'a pas déja été ajouté à la liste des nouveaux composants.
                        Dim m_newDevice As HoMIDom.HoMIDom.NewDevice =
                        (From dev In _Server.GetAllNewDevice()
                         Where dev.IdDriver = Me.ID _
                        And dev.Adresse1 = adresse _
                        And dev.Type = "BATTERIE").FirstOrDefault()

                        If (m_newDevice Is Nothing) Then  ' Non trouvé
                            'si autodiscover = true ou modedecouverte du serveur actif alors on crée le composant sinon
                            If _AutoDiscover Or _Server.GetModeDecouverte Then
                                'DBG
                                WriteLog("ERR: Device Batterie non trouvé, AutoCreation du composant : " & " " & adresse & ":" & valeur)
                                Dim unused = _Server.AddDetectNewDevice(adresse, _ID, "BATTERIE", "", valeur)
                            Else
                                WriteLog("ERR: Device Batterie non trouvé : " & " " & adresse & ":" & "Batterie Vide")
                            End If
                        Else
                            If m_newDevice.Ignore = False Then   '  Device trouvé dans New Device, afficher Erreur si ignorer n'est pas coché
                                WriteLog("ERR: Device Batterie trouvé dans la liste des nouveaux composants : " & ":" & adresse & ":" & valeur)
                            End If
                        End If
                        '*****************************************************
                    End If
                End If
            End If
            listedevices = Nothing
        Catch ex As Exception
            WriteLog("ERR: WriteBattery Exception : " & ex.Message & " --> " & adresse)
        End Try
    End Sub

    Private Sub WriteRetourSend(ByVal adresse As String, ByVal type As String, ByVal valeur As String)
        Try
            If Not _IsConnect Then Exit Sub 'si on ferme le port on quitte

            'Forcer le . 
            'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
            'My.Application.ChangeCulture("en-US")

            'log tous les paquets en mode debug
            If _DEBUG Then WriteLog("DBG: WriteRetour send to " & adresse & " (" & type & ") -> " & valeur)

            'on ne traite rien pendant les 6 premieres secondes
            If DateTime.Now > DateAdd(DateInterval.Second, 6, dateheurelancement) Then
                'Recherche si un device affecté
                Dim listedevices As New ArrayList
                listedevices = _Server.ReturnDeviceByAdresse1TypeDriver(_IdSrv, adresse, type, Me._ID, True)
                If IsNothing(listedevices) Then
                    WriteLog("ERR: Communication impossible avec le serveur, l'IDsrv est peut être erroné : " & _IdSrv)
                    Exit Sub
                End If
                If (listedevices.Count = 1) Then
                    'un device trouvé on maj la value
                    If valeur = "ON" Then
                        If TypeOf listedevices.Item(0).Value Is Boolean Then
                            listedevices.Item(0).Value = True
                        ElseIf TypeOf listedevices.Item(0).Value Is Long Or TypeOf listedevices.Item(0).Value Is Integer Then
                            listedevices.Item(0).Value = 100
                        Else
                            listedevices.Item(0).Value = "ON"
                        End If
                    ElseIf valeur = "OFF" Then
                        If TypeOf listedevices.Item(0).Value Is Boolean Then
                            listedevices.Item(0).Value = False
                        ElseIf TypeOf listedevices.Item(0).Value Is Long Or TypeOf listedevices.Item(0).Value Is Integer Then
                            listedevices.Item(0).Value = 0
                        Else
                            listedevices.Item(0).Value = "OFF"
                        End If
                    ElseIf valeur = "PANIC" Or valeur = "ALARM" Or valeur = "ALARM DELAYED" Or valeur = "ALARM + TAMPER" Or
                           valeur = "ALARM DELAYED + TAMPER" Or valeur = "MOTION" Or valeur = "MOTION + TAMPER" Then
                        If TypeOf listedevices.Item(0).Value Is Boolean Then
                            listedevices.Item(0).Value = True
                        ElseIf TypeOf listedevices.Item(0).Value Is Long Or TypeOf listedevices.Item(0).Value Is Integer Then
                            listedevices.Item(0).Value = 1
                        Else
                            listedevices.Item(0).Value = valeur
                        End If
                    ElseIf valeur = "NORMAL" Or valeur = "NORMAL + TAMPER" Or valeur = "NORMAL DELAYED + TAMPER" Or
                           valeur = "NORMAL DELAYED" Or valeur = "NO MOTION" Or valeur = "NO MOTION + TAMPER" Or
                           valeur = "PANIC END" Then
                        If TypeOf listedevices.Item(0).Value Is Boolean Then
                            listedevices.Item(0).Value = False
                        ElseIf TypeOf listedevices.Item(0).Value Is Long Or TypeOf listedevices.Item(0).Value Is Integer Then
                            listedevices.Item(0).Value = 0
                        Else
                            listedevices.Item(0).Value = valeur
                        End If
                    Else
                        listedevices.Item(0).Value = valeur
                    End If
                ElseIf (listedevices.Count > 1) Then
                    WriteLog("ERR: Plusieurs devices correspondent à : " & type & " " & adresse & ":" & valeur)
                Else
                    WriteLog("ERR: Device non trouvé : " & type & " " & adresse & ":" & valeur)
                End If
                listedevices = Nothing
            End If

        Catch ex As Exception
            WriteLog("ERR: WriteRetourSend Exception : " & adresse & " " & ex.Message)
        End Try
    End Sub

    Private Sub WriteRetour(ByVal adresse As String, ByVal type As String, ByVal valeur As String)
        Try
            'Forcer le . 
            'Thread.CurrentThread.CurrentCulture = New CultureInfo("en-US")
            'My.Application.ChangeCulture("en-US")

            'log tous les paquets en mode debug
            If _DEBUG Then WriteLog("DBG: WriteRetour : receive from " & adresse & " (" & type & ") -> " & valeur)

            If Not _IsConnect Then Exit Sub 'si on ferme le port on quitte
            If DateTime.Now < DateAdd(DateInterval.Second, 6, dateheurelancement) Then Exit Sub 'on ne traite rien pendant les 6 premieres secondes

            'Recherche si un device affecté
            Dim listedevices As New ArrayList
            listedevices = _Server.ReturnDeviceByAdresse1TypeDriver(_IdSrv, adresse, type, Me._ID, True)
            If IsNothing(listedevices) Then
                WriteLog("ERR: Communication impossible avec le serveur, l'IDsrv est peut être erroné : " & _IdSrv)
                Exit Sub
            End If
            If (listedevices.Count = 1) Then
                'un device trouvé 
                If STRGS.InStr(valeur, "CFG:") > 0 Then
                    'c'est un message de config, on log juste
                    WriteLog(listedevices.Item(0).name & " : " & valeur)
                Else
                    'on maj la value si la durée entre les deux receptions est > à 1.5s
                    If (DateTime.Now - Date.Parse(listedevices.Item(0).LastChange)).TotalMilliseconds > 1500 Then
                        If valeur = "ON" Then
                            If TypeOf listedevices.Item(0).Value Is Boolean Then
                                listedevices.Item(0).Value = True
                            ElseIf TypeOf listedevices.Item(0).Value Is Long Or TypeOf listedevices.Item(0).Value Is Integer Then
                                listedevices.Item(0).Value = 100
                            Else
                                listedevices.Item(0).Value = "ON"
                            End If
                        ElseIf valeur = "OFF" Then
                            If TypeOf listedevices.Item(0).Value Is Boolean Then
                                listedevices.Item(0).Value = False
                            ElseIf TypeOf listedevices.Item(0).Value Is Long Or TypeOf listedevices.Item(0).Value Is Integer Then
                                listedevices.Item(0).Value = 0
                            Else
                                listedevices.Item(0).Value = "OFF"
                            End If
                        ElseIf valeur = "PANIC" Or valeur = "ALARM" Or valeur = "ALARM DELAYED" Or valeur = "ALARM + TAMPER" Or
                               valeur = "ALARM DELAYED + TAMPER" Or valeur = "MOTION" Or valeur = "MOTION + TAMPER" Or
                               valeur = "GROUP_ON" Then    'JPS Ajout Group On
                            If TypeOf listedevices.Item(0).Value Is Boolean Then
                                listedevices.Item(0).Value = True
                            ElseIf TypeOf listedevices.Item(0).Value Is Long Or TypeOf listedevices.Item(0).Value Is Integer Then
                                listedevices.Item(0).Value = 1
                            Else
                                listedevices.Item(0).Value = valeur
                            End If
                        ElseIf valeur = "NORMAL" Or valeur = "NORMAL + TAMPER" Or valeur = "NORMAL DELAYED + TAMPER" Or
                               valeur = "NORMAL DELAYED" Or valeur = "NO MOTION" Or valeur = "NO MOTION + TAMPER" Or
                               valeur = "PANIC END" Or valeur = "GROUP_OFF" Then   'JPS ajout Group off
                            If TypeOf listedevices.Item(0).Value Is Boolean Then
                                listedevices.Item(0).Value = False
                            ElseIf TypeOf listedevices.Item(0).Value Is Long Or TypeOf listedevices.Item(0).Value Is Integer Then
                                listedevices.Item(0).Value = 0
                            Else
                                listedevices.Item(0).Value = valeur
                            End If
                        Else
                            listedevices.Item(0).Value = valeur
                        End If
                    Else
                        WriteLog("DBG: Reception < 1.5s de deux valeurs pour le meme composant : " & listedevices.Item(0).name & ":" & valeur)
                    End If
                End If
            ElseIf (listedevices.Count > 1) Then
                WriteLog("ERR: Plusieurs devices correspondent à : " & type & " " & adresse & ":" & valeur)
            Else

                '**************Modif JPS************************
                ' Vérifie que le composant n'a pas déja été ajouté à la liste des nouveaux composants.
                Dim m_newDevice As HoMIDom.HoMIDom.NewDevice =
                    (From dev In _Server.GetAllNewDevice() _
                        Where dev.IdDriver = Me.ID _
                        And dev.Adresse1 = adresse _
                        And dev.Type = type).FirstOrDefault()

                If (m_newDevice Is Nothing) Then  ' Non trouvé :Création du composant
                    'si autodiscover = true ou modedecouverte du serveur actif alors on crée le composant sinon
                    If _AutoDiscover Or _Server.GetModeDecouverte Then
                        'DBG
                        WriteLog("ERR: Device non trouvé, AutoCreation du composant : " & type & " " & adresse & ":" & valeur)
                        Dim unused = _Server.AddDetectNewDevice(adresse, _ID, type, "", valeur)
                    Else
                        WriteLog("ERR: Device non trouvé : " & type & " " & adresse & ":" & valeur)
                    End If
                Else
                    If m_newDevice.Ignore = False Then   '  Device trouvé dans New Device, afficher Erreur si ignorer n'est pas coché
                        WriteLog("ERR: Device trouvé dans la liste des nouveaux composants : " & type & " " & adresse & ":" & valeur)
                    End If
                End If
                '************************************

            ''si autodiscover = true alors on crée le composant sinon on logue
            'If _AutoDiscover Then
            '    If type = "" Then
            '        WriteLog("ERR: Device non trouvé, AutoCreation impossible du composant car le type ne peut etre déterminé : " & adresse & ":" & valeur)
            '    Else
            '        Try
            '            WriteLog("Device non trouvé, AutoCreation du composant : " & type & " " & adresse & ":" & valeur)
            '            _Server.SaveDevice(_IdSrv, "", "_RFXtrx_" & Date.Now.ToString("ddMMyyHHmmssf"), adresse, True, False, Me._ID, type, 0, "", "", "", "AutoDiscover RFXtrx", 0, False, "0", "", 0, 999999, -999999, 0, Nothing, "", 0, False)
            '        Catch ex As Exception
            '            WriteLog("ERR: Writeretour Exception : AutoDiscover Creation composant: " & ex.Message)
            '        End Try
            '    End If
            'Else
            '    WriteLog("ERR: Device non trouvé : " & type & " " & adresse & ":" & valeur)
            'End If
            '
            ''Ajouter la gestion des composants bannis (si dans la liste des composant bannis alors on log en debug sinon onlog device non trouve empty)

            End If
            listedevices = Nothing
        Catch ex As Exception
            WriteLog("ERR: WriteRetour Exception : " & adresse & " " & ex.Message)
        End Try
    End Sub

#End Region

End Class