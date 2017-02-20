﻿Module Variables
    Public myService As HoMIDom.HoMIDom.IHoMIDom
    Public IsConnect As Boolean
    Public IdSrv As String = "123456789"
    Public _ListElement As New List(Of uWidgetEmpty)
    Public frmMere As Window1 = Nothing
    Public Design As Boolean = False
    Public _MonRepertoire As String = System.Environment.CurrentDirectory
    Public _MonRepertoireAppData As String = System.Environment.CurrentDirectory
    Public lock_logwrite As New Object
    Public lock_dev As Boolean
    Public _ListServer As New List(Of ClServer)
    Public Media As WMedia = Nothing

    'Ecriture dans le fichier log
    Public _Debug As Boolean

    'Image des Devices/zones
    Public _AllZones As New List(Of HoMIDom.HoMIDom.Zone)
    Private _AllDevices As New List(Of HoMIDom.HoMIDom.TemplateDevice)

    'Animation des ScrollViewer
    Public m_friction As Double = 0.75
    Public m_SpeedTouch As Double = 600

    Public Property AllDevices As List(Of HoMIDom.HoMIDom.TemplateDevice)
        Get
            Return _AllDevices
        End Get
        Set(value As List(Of HoMIDom.HoMIDom.TemplateDevice))
            _AllDevices = value
        End Set
    End Property
End Module
