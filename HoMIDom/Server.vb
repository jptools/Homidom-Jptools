﻿#Region "Imports"
Imports System
Imports System.IO
Imports System.IO.Ports
Imports System.Data
'Imports System.Data.Linq
Imports System.Management
Imports System.Xml
Imports System.Xml.XPath
Imports System.Xml.Serialization
Imports System.Reflection
Imports System.ServiceModel
Imports System.ServiceModel.Description
Imports System.Security.Cryptography
Imports System.Text
Imports System.Web.HttpUtility
Imports System.Threading
Imports System.Data.SQLite
Imports System.Net
Imports System.Net.Sockets
Imports System.Net.Mail
Imports TagLib
Imports System.Text.RegularExpressions
Imports Newtonsoft

#End Region

Namespace HoMIDom

    ''' <summary>Classe Server</summary>
    ''' <remarks></remarks>
    <ServiceBehavior(InstanceContextMode:=InstanceContextMode.Single)>
    <Serializable()> Public Class Server
        Implements HoMIDom.IHoMIDom 'implémente l'interface dans cette class

        Public Shared Property Instance() As Server


#Region "Evènements"
        Public Event DeviceChanged(ByVal DeviceId As String, ByVal DeviceValue As String) Implements IHoMIDom.DeviceChanged 'Evènement lorsqu'un device change
        Public Event DeviceDeleted(ByVal DeviceId As String) Implements IHoMIDom.DeviceDeleted 'Evènement lorsqu'un device change
        Public Event NewLog(ByVal TypLog As HoMIDom.Server.TypeLog, ByVal Source As HoMIDom.Server.TypeSource, ByVal Fonction As String, ByVal Message As String) Implements IHoMIDom.NewLog  'Evènement lorsqu'un nouveau log est écrit
        Public Event MessageFromServeur(ByVal Id As String, ByVal Time As DateTime, ByVal Message As String) Implements IHoMIDom.MessageFromServeur  'Message provenant du serveur
        Public Event DriverChanged(ByVal DriverId As String) Implements IHoMIDom.DriverChanged 'Evènement lorsq'un driver est modifié
        Public Event ZoneChanged(ByVal ZoneId As String) Implements IHoMIDom.ZoneChanged 'Evènement lorsq'une zone est modifiée ou créée
        Public Event MacroChanged(ByVal MacroId As String) Implements IHoMIDom.MacroChanged 'Evènement lorsq'une macro est modifiée ou créée
        Public Event HeureSoleilChanged() Implements IHoMIDom.HeureSoleilChanged 'Evènement lorsque l'heure de lever/couché du soleil est modifié
#End Region

#Region "Declaration des variables"

        Private Shared WithEvents _ListDrivers As New ArrayList 'Liste des drivers
        Private Shared WithEvents _ListDevices As New ArrayList 'Liste des devices
        Private Shared WithEvents _ListNewDevices As New List(Of NewDevice) 'Liste des devices découverts
        Private Shared _ListImgDrivers As New List(Of Driver)

        <NonSerialized()> Private Shared _ListZones As New List(Of Zone) 'Liste des zones
        <NonSerialized()> Private Shared _ListUsers As New List(Of Users.User) 'Liste des users
        <NonSerialized()> Private Shared _ListMacros As New List(Of Macro) 'Liste des macros
        <NonSerialized()> Private Shared _ListTriggers As New List(Of Trigger) 'Liste de tous les triggers
        <NonSerialized()> Private Shared _ListGroups As New List(Of Groupes) 'Liste de tous les groupes
        <NonSerialized()> Private Shared _ListVars As New List(Of Variable) 'Liste de toutes les variables
        <NonSerialized()> Private Shared _listImages As New List(Of ImageFile) 'Liste toutes les fichiers images sur le serveur

        <NonSerialized()> Private sqlite_homidom As New Sqlite("homidom", Me) 'BDD sqlite pour Homidom
        <NonSerialized()> Private sqlite_medias As New Sqlite("medias", Me) 'BDD sqlite pour les medias
        <NonSerialized()> Shared Soleil As New Soleil 'Déclaration class Soleil
        <NonSerialized()> Shared _Longitude As Double = 2.3488  'Longitude
        <NonSerialized()> Shared _Latitude As Double = 48.85341  'latitude
        <NonSerialized()> Private Shared _HeureLeverSoleil As DateTime 'heure du levé du soleil
        <NonSerialized()> Private Shared _HeureCoucherSoleil As DateTime 'heure du couché du soleil
        <NonSerialized()> Shared _HeureLeverSoleilCorrection As Integer = 0 'correction à appliquer sur heure du levé du soleil
        <NonSerialized()> Shared _HeureCoucherSoleilCorrection As Integer = 0 'correction à appliquer sur heure du couché du soleil
        <NonSerialized()> Shared _SMTPServeur As String = "smtp.homidom.fr" 'adresse du serveur SMTP
        <NonSerialized()> Shared _SMTPLogin As String = "" 'login du serveur SMTP
        <NonSerialized()> Shared _SMTPassword As String = "" 'password du serveur SMTP
        <NonSerialized()> Shared _SMTPmailEmetteur As String = "homidom@mail.com" 'adresse mail de l'émetteur
        <NonSerialized()> Shared _SMTPPort As Integer = 587 'port smtp à utiliser
        <NonSerialized()> Shared _SMTPSSL As Boolean = True 'mail avec SSL
        <NonSerialized()> Private Shared _PortSOAP As String = "7999" 'Port IP de connexion SOAP
        <NonSerialized()> Private Shared _IPSOAP As String = "localhost" 'IP de connexion SOAP
        <NonSerialized()> Dim TimerSecond As New Timers.Timer 'Timer à la seconde
        <NonSerialized()> Shared _DateTimeLastStart As Date = Now
        Public Etat_server As Boolean 'etat du serveur : true = démarré
        '<NonSerialized()> Public Shared Etat_server As Boolean 'etat du serveur : true = démarré
        <NonSerialized()> Dim fsw As FileSystemWatcher
        <NonSerialized()> Dim _MaxMonthLog As Integer = 2
        <NonSerialized()> Dim _SaveDiffBackup As Boolean = False 'Sauvegader les backups et sauvegardes suivant différents fichiers
        <NonSerialized()> Private Shared _TypeLogEnable As New List(Of Boolean) 'True si on doit pas prendre en compte le type de log
        <NonSerialized()> Shared _LastLogs(9) As String  'Table des derniers logs
        <NonSerialized()> Shared _LastLogsError(9) As String  'Table des derniers logs en alerte ou erreur
        <NonSerialized()> Shared _DevicesNoMAJ As New List(Of String)  'Table des devices non à jour
        'Variable enregistrement de la config dans homidom.xml
        <NonSerialized()> Shared _SaveRealTime As Boolean = True 'True si on enregistre en temps réel
        <NonSerialized()> Shared _CycleSave As Integer  'Enregistrer toute les X minutes
        <NonSerialized()> Shared _NextTimeSave As DateTime  'Enregistrer toute les X minutes
        'Variable export de la config dans un dossier
        <NonSerialized()> Shared _FolderSaveFolder As String 'chemin ou exporter la config
        <NonSerialized()> Shared _CycleSaveFolder As Integer = 0  'export toutes les X heures
        <NonSerialized()> Shared _NextTimeSaveFolder As DateTime  'Enregistrer toute les X minutes

        <NonSerialized()> Shared _Finish As Boolean  'Le serveur est prêt
        <NonSerialized()> Shared _Voice As String  'Voix par défaut
        <NonSerialized()> Private Shared _OsPlatForm As String  '32 ou 64 bits
        <NonSerialized()> Private Shared _CodePays As Integer = 12
        Private Shared lock_logwrite As New Object
        <NonSerialized()> Shared _Devise As String = "€"
        <NonSerialized()> Shared _IsWeekEnd As Boolean = False

        <NonSerialized()> Shared _EnableSrvWeb As Boolean = False
        <NonSerialized()> Shared _PortSrvWeb As Integer = 8080
        <NonSerialized()> Shared _SrvWeb As ServeurWeb = Nothing
        <NonSerialized()> Shared _ModeDecouverte As Boolean = False 'Mode découverte des nouveaux devices
        <NonSerialized()> Shared _ListThread As New List(Of Thread)
        <NonSerialized()> Shared _SrvUDPIsStart As Boolean = False 'indique si le serveur UDP est actif

        'Variables Energie
        <NonSerialized()> Shared _PuissanceTotaleActuel As Integer = 0
        <NonSerialized()> Shared _PuissanceMini As Integer = 0 'Puissance mini par défaut pour connaitre la puissance de base de la maison pour les composants non gérés par le serveur
        <NonSerialized()> Shared _GererEnergie As Boolean = False
        <NonSerialized()> Shared _TarifJour As Double = 0
        <NonSerialized()> Shared _TarifNuit As Double = 0

        'Gestion des threads
        'Private Shared lock_table_TimerSecTickthread As New Object
        ' Public Shared table_TimerSecTickthread As New DataTable
        <NonSerialized()> Private ListThread As New List(Of Thread)

        'used to verify Local System separator in server.start function
        Private Declare Function GetLocaleInfoEx Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
        Private Declare Function GetThreadLocale Lib "kernel32" () As Long

#End Region

#Region "Event"
        Public Sub VarEvent(Nom As String, e As String)
            Try
                ManagerSequences.AddSequences(Sequence.TypeOfSequence.VariableChange, Nom, "", e)
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "VarEvent", "La variable: " & Nom & " a changée de valeur=" & e)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "VarEvent", "Exception : " & ex.Message)
            End Try
        End Sub


        ''' <summary>Evenement provenant des drivers </summary>
        ''' <param name="DriveName"></param>
        ''' <param name="TypeEvent"></param>
        ''' <param name="Parametre"></param>
        ''' <remarks></remarks>
        Public Sub DriversEvent(ByVal DriveName As String, ByVal TypeEvent As String, ByVal Parametre As Object)
            Try
                If Etat_server Then
                    For Each _drv In GetAllDrivers(_IdSrv)
                        If _drv.Nom = DriveName Then
                            ManagerSequences.AddSequences(Sequence.TypeOfSequence.Driver, _drv.ID, "", Nothing)
                            RaiseEvent DriverChanged(_drv.ID)
                            Exit For
                        End If
                    Next

                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DriversEvent", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Evenement provenant des devices</summary>
        ''' <param name="Device"></param>
        ''' <param name="Property"></param>
        ''' <param name="Parametres"></param>
        ''' <remarks></remarks>
        Public Sub DeviceChange(ByVal Device As Object, ByVal [Property] As String, ByVal Parametres As Object)
            Dim retour As String = ""
            Dim valeurString As String = String.Empty
            Dim genericDevice As HoMIDom.Device.DeviceGenerique = Device

            If Parametres IsNot Nothing Then
                valeurString = Parametres.ToString()
            End If

            '------------------------------------------------------------------------------------------------
            '   ENERGIE
            '------------------------------------------------------------------------------------------------
            If GererEnergie Then

                Try
                    'Si le composant permet de donner la puissance instantanée totale ou une partie on la récupère 
                    If genericDevice.Type = "ENERGIEINSTANTANEE" Then
                        PuissanceTotaleActuel = CInt(Device.Value)
                        Log(TypeLog.DEBUG, TypeSource.SERVEUR, "DeviceChange", "Calcul Energie ")
                    End If

                    If CInt(genericDevice.Puissance) > 0 Then
                        Dim _CalculVariation As Boolean = False 'True si le calcul doit prendre en compte la variation ou la suivant la valeur du device

                        Select Case genericDevice.Type
                            Case "APPAREIL"
                            Case "AUDIO"
                            Case "BAROMETRE"
                            Case "BATTERIE"
                            Case "COMPTEUR"
                            Case "CONTACT"
                            Case "DETECTEUR"
                            Case "DIRECTIONVENT"
                            Case "ENERGIEINSTANTANEE"
                            Case "ENERGIETOTALE"
                            Case "FREEBOX"
                            Case "GENERIQUEBOOLEEN"
                            Case "GENERIQUESTRING"
                            Case "GENERIQUEVALUE"
                            Case "HUMIDITE"
                            Case "LAMPE"
                                _CalculVariation = True
                            Case "LAMPERGBW"
                                _CalculVariation = True
                            Case "METEO"
                            Case "MULTIMEDIA"
                            Case "PLUIECOURANT"
                            Case "PLUIETOTAL"
                            Case "SWITCH"
                            Case "TELECOMMANDE"
                            Case "TEMPERATURE"
                            Case "TEMPERATURECONSIGNE"
                            Case "UV"
                            Case "VITESSEVENT"
                            Case "VOLET"

                        End Select

                        If IsNumeric(Device.value) And IsNumeric(genericDevice.Puissance) Then
                            If CInt(Device.Value) = 0 Then
                                PuissanceTotaleActuel -= CInt(genericDevice.Puissance)
                            Else
                                Dim _puissance As Integer = CInt(genericDevice.Puissance)
                                If _puissance < 0 Then _puissance = _puissance * -1
                                'Si prise en compte du calcul de variation
                                If _CalculVariation Then
                                    PuissanceTotaleActuel += ((_puissance * Device.Value) / 100)
                                Else
                                    PuissanceTotaleActuel += _puissance
                                End If
                            End If
                        End If

                    End If
                Catch ex As Exception
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceChange", "Calcul Energie Exception : " & ex.Message)
                End Try
            End If

            ManagerSequences.AddSequences(Sequence.TypeOfSequence.DeviceChange, genericDevice.ID, Nothing, valeurString)
            RaiseEvent DeviceChanged(genericDevice.ID, valeurString)

            Try
                If Etat_server Then
                    Dim valeur = Parametres
                    '--- on logue tout ce qui arrive en mode debug
                    'Log(TypeLog.DEBUG, TypeSource.SERVEUR, "DeviceChange", "Le device " & Device.name & " a changé : " & [Property] & " = " & valeur.ToString)

                    If Mid(valeur.ToString, 1, 4) <> "ERR:" Then 'si y a pas erreur d'acquisition

                        '------------------------------------------------------------------------------------------------
                        '    MACRO/Triggers
                        '------------------------------------------------------------------------------------------------
                        Try
                            'Parcour des triggers pour vérifier si le device déclenche des macros
                            Dim _m As Macro
                            For i As Integer = 0 To _ListTriggers.Count - 1
                                If _ListTriggers.Item(i).Enable = True Then
                                    If _ListTriggers.Item(i).Type = Trigger.TypeTrigger.DEVICE And Device.id = _ListTriggers.Item(i).ConditionDeviceId And _ListTriggers.Item(i).ConditionDeviceProperty = [Property] Then 'c'est un trigger type device + enable + device concerné
                                        Log(TypeLog.DEBUG, TypeSource.SERVEUR, "DeviceChange", " -> " & Device.name & " est associé au trigger : " & _ListTriggers.Item(i).Nom)
                                        'on lance toutes les macros associés
                                        For j As Integer = 0 To _ListTriggers.Item(i).ListMacro.Count - 1
                                            _m = ReturnMacroById(_IdSrv, _ListTriggers.Item(i).ListMacro.Item(j))
                                            If _m IsNot Nothing Then
                                                If _m.Enable Then
                                                    Log(TypeLog.DEBUG, TypeSource.SERVEUR, "DeviceChange", " --> " & _ListTriggers.Item(i).Nom & " Lance la macro : " & _m.Nom)
                                                    _m.Execute(Me)
                                                Else
                                                    Log(TypeLog.DEBUG, TypeSource.SERVEUR, "DeviceChange", " --> " & _ListTriggers.Item(i).Nom & " Macro désactivé : " & _m.Nom)
                                                End If
                                                _m = Nothing
                                            End If

                                        Next
                                    End If
                                End If
                            Next
                        Catch ex As Exception
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceChange", "Macro/Triggers Exception : " & ex.Message)
                        End Try

                        '------------------------------------------------------------------------------------------------
                        '    HISTORIQUE
                        '------------------------------------------------------------------------------------------------
                        Try
                            'Ajout dans la BDD
                            If Device.isHisto Then
                                If Device.countTempHisto = Device.RefreshHisto Or Device.RefreshHisto = 0 Then
                                    retour = sqlite_homidom.nonquery("INSERT INTO historiques (device_id,source,dateheure,valeur) VALUES (@parameter0, @parameter1, @parameter2, @parameter3)", Device.ID, [Property], Now.ToString("yyyy-MM-dd HH:mm:ss"), valeur)
                                    If Mid(retour, 1, 4) = "ERR:" Then
                                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceChange", "Erreur Requete sqlite : " & retour)
                                    Else
                                        Device.CountHisto += 1
                                        Device.countTempHisto = 1
                                        ManagerSequences.AddSequences(Sequence.TypeOfSequence.HistoryChange, Nothing, Nothing, Nothing)
                                    End If
                                Else
                                    Device.countTempHisto += 1
                                End If
                            End If
                        Catch ex As Exception
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceChange", "Historique Exception : " & ex.Message)
                        End Try

                    Else
                        'erreur d'acquisition
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceChange", "Erreur d'acquisition : " & Device.Name & " - " & valeur.ToString)
                    End If
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceChange", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Traitement à effectuer toutes les secondes/minutes/heures/minuit/midi</summary>
        ''' <remarks></remarks>
        Sub TimerSecTick()
            Dim ladate As DateTime = Now 'on récupére la date/heure

            Try
                Try
                    'verif si pas déjà trop de thread
                    If ListThread.Count > 10 Then
                        If ListThread(0).IsAlive Then
                            ListThread(0).Abort()
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "gestion thread", "Arrêt du thread " & ListThread(0).Name & " car il a pris trop de temps")
                        End If
                        ListThread.RemoveAt(0)
                    End If
                Catch ex As Exception
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "gestion thread", "Exception : " & ex.Message)
                End Try

                '---- Action à effectuer toutes les secondes ----
                Dim thr1 As New Thread(AddressOf ThreadSecond)
                thr1.Name = "ThreadSecond"
                thr1.IsBackground = True
                thr1.Priority = ThreadPriority.Highest
                thr1.Start()
                ListThread.Add(thr1)
                thr1 = Nothing

                '---- Actions à effectuer toutes les minutes ----
                If ladate.Second = 0 Then
                    Dim thr2 As New Thread(AddressOf ThreadMinute)
                    thr2.Name = "ThreadMinute"
                    thr2.IsBackground = True
                    thr2.Priority = ThreadPriority.Normal
                    thr2.Start()
                    ListThread.Add(thr2)
                End If

                '---- Actions à effectuer toutes les heures ----
                If ladate.Minute = 59 And ladate.Second = 59 Then
                    Dim thr3 As New Thread(AddressOf ThreadHour)
                    thr3.Name = "ThreadHour"
                    thr3.IsBackground = True
                    thr3.Priority = ThreadPriority.Normal
                    thr3.Start()
                    ListThread.Add(thr3)
                End If

                '---- Actions à effectuer à minuit ----
                If ladate.Hour = 0 And ladate.Minute = 0 And ladate.Second = 0 Then
                    Dim thr4 As New Thread(AddressOf ThreadMinuit)
                    thr4.Name = "ThreadMinuit"
                    thr4.IsBackground = True
                    thr4.Priority = ThreadPriority.Normal
                    thr4.Start()
                    ListThread.Add(thr4)
                End If

                '---- Actions à effectuer à 3h du mat (au cas où qu'à minuit non maj) ----
                If ladate.Hour = 3 And ladate.Minute = 0 And ladate.Second = 0 Then
                    Dim thr5 As New Thread(AddressOf Thread3h)
                    thr5.Name = "Thread3h"
                    thr5.IsBackground = True
                    thr5.Priority = ThreadPriority.AboveNormal
                    thr5.Start()
                    ListThread.Add(thr5)
                End If

                '---- Actions à effectuer à midi ----
                If ladate.Hour = 12 And ladate.Minute = 0 And ladate.Second = 0 Then
                    Dim thr6 As New Thread(AddressOf ThreadMidi)
                    thr6.Name = "ThreadMidi"
                    thr6.IsBackground = True
                    thr6.Priority = ThreadPriority.AboveNormal
                    thr6.Start()
                    ListThread.Add(thr6)
                End If

            Catch ex2 As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "TimerSecTick", "Exception 3 : " & ex2.ToString & " --> Erreur du thread TimerSecTick, suppression de tous les threads en cours")

                For Each thr In ListThread
                    If thr.IsAlive Then thr.Abort()
                Next
            End Try
            'End Try
        End Sub

#Region "Thread"


        ''' <summary>
        ''' Thread à effectuer toutes les secondes
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ThreadSecond()
            Try
                Dim thr As New Thread(AddressOf VerifTimeDevice)
                thr.IsBackground = True
                thr.Name = "VerifTimeDevice"
                thr.Start()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ThreadSecond", "Exception: " & ex.ToString)
            End Try
        End Sub

        ''' <summary>
        ''' Thread à effectuer toutes les minutes
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ThreadMinute()
            Try
                Dim thr1 As New Thread(AddressOf VerifIsJour)
                thr1.IsBackground = True
                thr1.Name = "VerifIsJour"
                thr1.Start()

                Dim thr As New Thread(AddressOf ThreadSaveConfig)
                thr.IsBackground = True
                thr.Name = "ThreadSaveConfig"
                thr.Start()

                SearchDeviceNoMaJ()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ThreadMinute", "Exception: " & ex.ToString)
            End Try
        End Sub

        ''' <summary>
        ''' Thread à effectuer toutes les heures
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ThreadHour()
            Try
                Dim thr As New Thread(AddressOf ThreadSaveConfigFolder)
                thr.IsBackground = True
                thr.Priority = ThreadPriority.Lowest
                thr.Name = "ThreadSaveConfigFolder"
                thr.Start()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ThreadHour", "Exception: " & ex.ToString)
            End Try
        End Sub

        ''' <summary>
        ''' Thread à effectuer à minuit
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ThreadMinuit()
            Try
                Dim thr As New Thread(AddressOf MaJSaint)
                thr.IsBackground = True
                thr.Priority = ThreadPriority.Lowest
                thr.Name = "ThreadMAJSaint"
                thr.Start()

                MAJ_HeuresSoleil()
                CleanLog()
                VerifIsWeekEnd()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ThreadMinuit", "Exception: " & ex.ToString)
            End Try
        End Sub

        ''' <summary>
        ''' Thread à effectuer à 3h du matin
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Thread3h()
            Try
                Dim thr As New Thread(AddressOf MaJSaint)
                thr.IsBackground = True
                thr.Priority = ThreadPriority.Lowest
                thr.Name = "ThreadMAJSaint"
                thr.Start()

                MAJ_HeuresSoleil()
                VerifIsWeekEnd()
                VerifPurge()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Thread3h", "Exception: " & ex.ToString)
            End Try
        End Sub

        ''' <summary>
        ''' Thread à effectuer à midi
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ThreadMidi()
            Try
                MAJ_HeuresSoleil()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ThreadMidi", "Exception: " & ex.ToString)
            End Try
        End Sub

        ''' <summary>
        ''' Thread permettant la sauvegarde de la config
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ThreadSaveConfig()
            Try
                Dim ladate As DateTime = Now 'on récupére la date/heure

                'on veirife si on doit enregistrer la config dans le xml
                If _CycleSave > 0 And _Finish = True Then
                    If ladate >= _NextTimeSave Then
                        _NextTimeSave = Now.AddMinutes(_CycleSave)
                        SaveConfig(_MonRepertoire & "\config\homidom.xml")
                    End If
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ThreadSaveConfig", "Exception: " & ex.ToString)
            End Try
        End Sub

        ''' <summary>
        ''' Thread permettant de sauvegarder la config dans un autre dossier
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub ThreadSaveConfigFolder()
            Try
                Dim ladate As DateTime = Now 'on récupére la date/heure

                'on verifie si on doit exporter la config vers un folder
                If _CycleSaveFolder > 0 And _Finish = True Then
                    If ladate >= _NextTimeSaveFolder Then
                        _NextTimeSaveFolder = Now.AddHours(_CycleSaveFolder)
                        SaveConfigFolder()
                    End If
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ThreadSaveConfigFolder", "Exception: " & ex.ToString)
            End Try
        End Sub
#End Region

        ''' <summary>on checke si il y a cron à faire</summary>
        ''' <remarks></remarks>
        Private Sub VerifTimeDevice()
            Try

                For Each _Trigger In _ListTriggers
                    If _Trigger.Enable Then
                        If _Trigger.Type = Trigger.TypeTrigger.TIMER Then
                            If _Trigger.Prochainedateheure <= DateAndTime.Now.ToString("yyyy-MM-dd HH:mm:ss") Then
                                _Trigger.maj_cron() 'reprogrammation du prochain shedule

                                Dim _m As Macro = Nothing
                                'lancement des macros associées
                                For j As Integer = 0 To _Trigger.ListMacro.Count - 1
                                    'on cherche la macro et on la lance en testant ces conditions
                                    _m = ReturnMacroById(_IdSrv, _Trigger.ListMacro.Item(j))
                                    If _m IsNot Nothing Then
                                        If _m.Enable Then
                                            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "TriggerTimer", "Lancement de la macro: " & _m.Nom & " ,suite au déclenchement du trigger: " & _Trigger.Nom)
                                            _m.Execute(Me)
                                        Else
                                            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "TriggerTimer", "Macro: " & _m.Nom & " non executée car Désactivée ,suite au déclenchement du trigger: " & _Trigger.Nom)
                                        End If
                                    End If
                                        _m = Nothing
                                Next

                            End If
                        End If
                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "TimerSecTick TriggerTimer", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Evenement sur UnhandleException</summary>
        ''' <remarks></remarks>
        Private Sub Server_UnhandledExceptionEvent(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "UnhandledExceptionEvent", "Exception : " & e.ExceptionObject.ToString())
        End Sub

#End Region

#Region "Fonctions/Sub propres au serveur"

#Region "Serveur"
        Public Function GetListThread() As List(Of Thread)
            Try
                Return _ListThread
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetListThread", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Permet d'envoyer un message d'un client vers le server</summary>
        ''' <param name="Message"></param>
        ''' <remarks></remarks>
        Public Sub MessageFromServer(ByVal Message As String)
            Try
                RaiseEvent MessageFromServeur(Api.GenerateGUID, Now, Message)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "MessageFromServer", "Exception : " & ex.Message)
            End Try
        End Sub


        Private Property CodePays As Integer
            Get
                Return _CodePays
            End Get
            Set(ByVal value As Integer)
                _CodePays = value
            End Set
        End Property

        ''' <summary>
        ''' Vérifie si l'IDsrv est correct, retourne True si ok
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function VerifIdSrv(ByVal Value As String) As Boolean
            Try
                If Value = _IdSrv Then
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "VerifIdSrv", "Exception : " & ex.Message)
                Return False
            End Try
        End Function
#End Region

#Region "Soleil"
        Public Sub MaJSaint()
            Try
                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).id = "saint01" Then
                        _ListDevices.Item(i).value = GetSaint()
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "MaJSaint", "Exception : " & ex.Message)
            End Try
        End Sub

        Public Function GetSaint() As String
            Try
                Dim lines() As String = My.Resources.Saint.Split(System.Environment.NewLine)
                Dim currentDate As DateTime = DateTime.Now
                If DateTime.IsLeapYear(currentDate.Year) Then
                    Return lines(Now.DayOfYear - 1).Replace(Chr(10), "")
                Else
                    If currentDate.Month > 2 Then
                        Return lines(Now.DayOfYear).Replace(Chr(10), "")
                    Else
                        Return lines(Now.DayOfYear - 1).Replace(Chr(10), "")
                    End If
                End If

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSaint", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Vérifie si c'est le weekend
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub VerifIsWeekEnd()
            Try
                If Now.DayOfWeek = DayOfWeek.Sunday Or Now.DayOfWeek = DayOfWeek.Saturday Then
                    IsWeekEnd = True
                Else
                    IsWeekEnd = False
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "VerifIsWeekEnd", "Exception : " & ex.Message)
            End Try
        End Sub

        Public Property IsWeekEnd As Boolean
            Get
                Return _IsWeekEnd
            End Get
            Set(ByVal value As Boolean)
                If _IsWeekEnd <> value Then
                    _IsWeekEnd = value
                    For i As Integer = 0 To _ListDevices.Count - 1
                        If _ListDevices.Item(i).id = "isweekend01" Then
                            _ListDevices.Item(i).value = _IsWeekEnd
                            Exit For
                        End If
                    Next
                End If
            End Set
        End Property

        Private Sub VerifIsJour()
            Try
                If _HeureLeverSoleil <= Now And _HeureCoucherSoleil >= Now Then
                    For i As Integer = 0 To _ListDevices.Count - 1
                        If _ListDevices.Item(i).id = "soleil01" Then
                            If _ListDevices.Item(i).value = False Then _ListDevices.Item(i).value = True
                            Exit For
                        End If
                    Next
                Else
                    For i As Integer = 0 To _ListDevices.Count - 1
                        If _ListDevices.Item(i).id = "soleil01" Then
                            If _ListDevices.Item(i).value = True Then _ListDevices.Item(i).value = False
                            Exit For
                        End If
                    Next
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "VerifIsJour", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Initialisation des heures du soleil</summary>
        ''' <remarks></remarks>
        Public Sub MAJ_HeuresSoleil()
            Try
                Dim dtSunrise As Date
                Dim dtSolarNoon As Date
                Dim dtSunset As Date

                Soleil.CalculateSolarTimes(_Latitude, _Longitude, Date.Now, dtSunrise, dtSolarNoon, dtSunset)
                Log(TypeLog.INFO, TypeSource.SERVEUR, "MAJ_HeuresSoleil", "Initialisation des heures du soleil")
                _HeureCoucherSoleil = DateAdd(DateInterval.Minute, _HeureCoucherSoleilCorrection, dtSunset)
                _HeureLeverSoleil = DateAdd(DateInterval.Minute, _HeureLeverSoleilCorrection, dtSunrise)

                RaiseEvent HeureSoleilChanged()

                Log(TypeLog.INFO, TypeSource.SERVEUR, "MAJ_HeuresSoleil", "Heure du lever : " & _HeureLeverSoleil)
                Log(TypeLog.INFO, TypeSource.SERVEUR, "MAJ_HeuresSoleil", "Heure du coucher : " & _HeureCoucherSoleil)

                VerifIsJour()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "MAJ_HeuresSoleil", "Exception : " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Configuration"
        Private Sub SaveRealTime()
            Try
                If _SaveRealTime Then SaveConfig(_MonRepertoire & "\config\homidom.xml")

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveRealTime", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Chargement de la config depuis le fichier XML</summary>
        ''' <param name="Fichier"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function LoadConfig(ByVal Fichier As String) As String
            Dim _file As String = Fichier & "homidom.xml"

            'Copy du fichier de config avant chargement
            Try
                If _SaveDiffBackup = False Then
                    If IO.File.Exists(_file.Replace(".xml", ".bak")) = True Then IO.File.Delete(_file.Replace(".xml", ".bak"))
                End If
                If IO.File.Exists(_file) = True Then
                    If _SaveDiffBackup = False Then
                        IO.File.Copy(_file, _file.Replace(".xml", ".bak"))
                    Else
                        Dim fich As String = _file.Replace(".xml", ".bak")
                        fich = fich.Replace(".", Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & ".")
                        IO.File.Copy(_file, fich)
                    End If
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Création du backup (.bak) du fichier de config avant chargement")
                Else
                    LoadConfig = "Fichier de configuration (" & _file & ") inexistant."
                    Exit Function
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "LoadConfig", "Erreur impossible de créer une copie de backup du fichier de config: " & ex.Message)
            End Try

            Try
                Dim dirInfo As New System.IO.DirectoryInfo(Fichier)
                Dim file As System.IO.FileInfo
                Dim files() As System.IO.FileInfo = dirInfo.GetFiles("homidom.xml")
                Dim myxml As XML
                Dim myfile As String = ""

                If (files IsNot Nothing) Then
                    For Each file In files
                        myfile = file.FullName
                        Dim list As XmlNodeList

                        myxml = New XML(myfile)

                        Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement du fichier config: " & myfile)

                        'on va tester le fichier
                        If myxml IsNot Nothing Then
                            Try
                                list = myxml.SelectNodes("/homidom/server")
                            Catch ex3 As Exception
                                Dim reponse As MsgBoxResult = MsgBox("Erreur lors de la lecture du fichier de config: " & ex3.Message & vbCrLf & vbCrLf & "Voulez-vous tenter de démarrer depuis le fichier de config sauvegardé (si Oui pensez à sauvegarder la configuration depuis l'admin si réussite)?", MsgBoxStyle.YesNo, "ERREUR SERVICE LoadConfig")
                                Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement de la configuration du serveur : " & ex3.ToString)

                                If reponse = MsgBoxResult.Yes Then
                                    Try 'Seconde chance
                                        Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement de la configuration chargement du fichier de sauvegarde")
                                        myfile = myfile.Replace(".xml", ".sav") 'on va chercher le fichier sav

                                        If IO.File.Exists(myfile) = True Then
                                            myxml = New XML(myfile)
                                            list = myxml.SelectNodes("/homidom/server")
                                        Else
                                            Return "Erreur lors du chargement de la configuration du serveur en seconde chance car le fichier de sauvegarde n'existe pas"
                                        End If

                                    Catch ex2 As Exception
                                        Return "Erreur lors du chargement de la configuration du serveur en seconde chance : " & ex2.ToString
                                    End Try
                                End If
                            End Try
                        End If

                        '******************************************
                        'on va chercher les paramètres du serveur
                        '******************************************
                        Try
                            list = myxml.SelectNodes("/homidom/server")
                            If list.Count > 0 Then 'présence des paramètres du server
                                For j As Integer = 0 To list.Item(0).Attributes.Count - 1
                                    Select Case list.Item(0).Attributes.Item(j).Name
                                        Case "savediff"
                                            _SaveDiffBackup = list.Item(0).Attributes.Item(j).Value
                                        Case "longitude"
                                            '_Longitude = list.Item(0).Attributes.Item(j).Value.Replace(".", ",")
                                            _Longitude = Regex.Replace(list.Item(0).Attributes.Item(j).Value, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
                                        Case "latitude"
                                            '_Latitude = list.Item(0).Attributes.Item(j).Value.Replace(".", ",")
                                            _Latitude = Regex.Replace(list.Item(0).Attributes.Item(j).Value, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
                                        Case "heurecorrectionlever"
                                            _HeureLeverSoleilCorrection = list.Item(0).Attributes.Item(j).Value
                                        Case "heurecorrectioncoucher"
                                            _HeureCoucherSoleilCorrection = list.Item(0).Attributes.Item(j).Value
                                        Case "ipsoap"
                                            If String.IsNullOrEmpty(list.Item(0).Attributes.Item(j).Value) Then
                                                _IPSOAP = "localhost"
                                            Else
                                                _IPSOAP = list.Item(0).Attributes.Item(j).Value
                                            End If
                                        Case "portsoap"
                                            If String.IsNullOrEmpty(list.Item(0).Attributes.Item(j).Value) = False Or IsNumeric(CInt(list.Item(0).Attributes.Item(j).Value)) = True Then
                                                _PortSOAP = list.Item(0).Attributes.Item(j).Value
                                            Else
                                                _PortSOAP = "7999"
                                            End If
                                        Case "idsrv"
                                            If String.IsNullOrEmpty(list.Item(0).Attributes.Item(j).Value) = False Then
                                                _IdSrv = list.Item(0).Attributes.Item(j).Value
                                            Else
                                                _IdSrv = "123456789"
                                            End If
                                        Case "smtpserver"
                                            _SMTPServeur = list.Item(0).Attributes.Item(j).Value
                                        Case "smtpmail"
                                            _SMTPmailEmetteur = list.Item(0).Attributes.Item(j).Value
                                        Case "smtplogin"
                                            _SMTPLogin = list.Item(0).Attributes.Item(j).Value
                                        Case "smtppassword"
                                            _SMTPassword = list.Item(0).Attributes.Item(j).Value
                                        Case "smtpport"
                                            _SMTPPort = list.Item(0).Attributes.Item(j).Value
                                        Case "smtpssl"
                                            _SMTPSSL = list.Item(0).Attributes.Item(j).Value
                                        Case "logmaxfilesize"
                                            _MaxFileSize = list.Item(0).Attributes.Item(j).Value
                                        Case "logmaxmonthlog"
                                            _MaxMonthLog = list.Item(0).Attributes.Item(j).Value
                                        Case "log0"
                                            _TypeLogEnable(0) = list.Item(0).Attributes.Item(j).Value
                                        Case "log1"
                                            _TypeLogEnable(1) = list.Item(0).Attributes.Item(j).Value
                                        Case "log2"
                                            _TypeLogEnable(2) = list.Item(0).Attributes.Item(j).Value
                                        Case "log3"
                                            _TypeLogEnable(3) = list.Item(0).Attributes.Item(j).Value
                                        Case "log4"
                                            _TypeLogEnable(4) = list.Item(0).Attributes.Item(j).Value
                                        Case "log5"
                                            _TypeLogEnable(5) = list.Item(0).Attributes.Item(j).Value
                                        Case "log6"
                                            _TypeLogEnable(6) = list.Item(0).Attributes.Item(j).Value
                                        Case "log7"
                                            _TypeLogEnable(7) = list.Item(0).Attributes.Item(j).Value
                                        Case "log8"
                                            _TypeLogEnable(8) = list.Item(0).Attributes.Item(j).Value
                                        Case "log9"
                                            _TypeLogEnable(9) = list.Item(0).Attributes.Item(j).Value
                                        Case "cyclesave"
                                            _CycleSave = list.Item(0).Attributes.Item(j).Value
                                        Case "voice"
                                            _Voice = list.Item(0).Attributes.Item(j).Value
                                            If String.IsNullOrEmpty(_Voice) = True Then
                                                If String.IsNullOrEmpty(GetFirstVoice) = False Then
                                                    _Voice = GetFirstVoice()
                                                End If
                                            End If
                                        Case "saverealtime"
                                            _SaveRealTime = list.Item(0).Attributes.Item(j).Value
                                        Case "foldersavefolder"
                                            _FolderSaveFolder = list.Item(0).Attributes.Item(j).Value
                                        Case "cyclesavefolder"
                                            _CycleSaveFolder = list.Item(0).Attributes.Item(j).Value
                                        Case "devise"
                                            _Devise = list.Item(0).Attributes.Item(j).Value
                                        Case "puissancemini"
                                            PuissanceMini = list.Item(0).Attributes.Item(j).Value
                                        Case "gererenergie"
                                            GererEnergie = list.Item(0).Attributes.Item(j).Value
                                        Case "tarifjour"
                                            _TarifJour = list.Item(0).Attributes.Item(j).Value
                                        Case "tarifnuit"
                                            _TarifNuit = list.Item(0).Attributes.Item(j).Value
                                        Case "portweb"
                                            _PortSrvWeb = list.Item(0).Attributes.Item(j).Value
                                        Case "enablesrvweb"
                                            _EnableSrvWeb = list.Item(0).Attributes.Item(j).Value
                                        Case "modedecouverte"
                                            _ModeDecouverte = list.Item(0).Attributes.Item(j).Value
                                            If _ModeDecouverte Then Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " Mode découverte activé sur tous les drivers")
                                        Case "codepays"
                                            CodePays = list.Item(0).Attributes.Item(j).Value
                                        Case Else
                                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Un attribut correspondant au serveur est inconnu: nom:" & list.Item(0).Attributes.Item(j).Name & " Valeur: " & list.Item(0).Attributes.Item(j).Value)
                                    End Select
                                Next
                            Else
                                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "LoadConfig", "Erreur : Il manque les paramètres du serveur dans le fichier de config !!")
                            End If
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Paramètres du serveur chargés")
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement de la configuraion du serveur : " & ex.Message)
                        End Try

                        '********************************
                        'on va chercher les drivers
                        '*********************************
                        Try
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement des drivers :")
                            list = Nothing
                            list = myxml.SelectNodes("/homidom/drivers/driver")

                            If list.Count > 0 Then 'présence d'un ou des driver(s)
                                For j As Integer = 0 To list.Count - 1
                                    Dim _drv As IDriver = ReturnDrvById(_IdSrv, list.Item(j).Attributes.Item(0).Value)
                                    If _drv IsNot Nothing Then
                                        _drv.Enable = list.Item(j).Attributes.GetNamedItem("enable").Value
                                        _drv.StartAuto = list.Item(j).Attributes.GetNamedItem("startauto").Value
                                        _drv.IP_TCP = list.Item(j).Attributes.GetNamedItem("iptcp").Value
                                        _drv.Port_TCP = list.Item(j).Attributes.GetNamedItem("porttcp").Value
                                        _drv.IP_UDP = list.Item(j).Attributes.GetNamedItem("ipudp").Value
                                        _drv.Port_UDP = list.Item(j).Attributes.GetNamedItem("portudp").Value
                                        _drv.COM = list.Item(j).Attributes.GetNamedItem("com").Value
                                        _drv.Refresh = list.Item(j).Attributes.GetNamedItem("refresh").Value
                                        _drv.Modele = list.Item(j).Attributes.GetNamedItem("modele").Value
                                        _drv.IdSrv = _IdSrv

                                        If Not IsNothing(list.Item(j).Attributes.GetNamedItem("autodiscover")) Then
                                            _drv.AutoDiscover = list.Item(j).Attributes.GetNamedItem("autodiscover").Value
                                        Else
                                            _drv.AutoDiscover = False
                                        End If
                                        'If _ModeDecouverte = True Then _drv.AutoDiscover = True 'si mode decouverte activé sur le serveur, on force à true sur tous les drivers

                                        'Force le driver virtuel 
                                        If _drv.ID = "DE96B466-2540-11E0-A321-65D7DFD72085" Then
                                            _drv.Enable = True
                                            _drv.StartAuto = True
                                            _drv.AutoDiscover = False
                                        End If

                                        Dim a As String
                                        Dim idx As Integer
                                        For i As Integer = 0 To list.Item(j).Attributes.Count - 1
                                            a = UCase(list.Item(j).Attributes.Item(i).Name)
                                            If a.StartsWith("PARAMETRE") Then
                                                idx = Mid(a, 10, Len(a) - 9)
                                                If idx < _drv.Parametres.Count Then
                                                    _drv.Parametres.Item(idx).valeur = list.Item(j).Attributes.Item(i).Value
                                                End If
                                            End If

                                        Next
                                        a = Nothing
                                        idx = Nothing

                                        Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " - " & _drv.Nom & " chargé")
                                        _drv = Nothing
                                    End If
                                Next
                            Else
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Aucun driver n'est enregistré dans le fichier de config")
                            End If
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement des drivers : " & ex.Message)
                        End Try

                        '******************************************
                        'on va chercher les zones
                        '******************************************
                        Try
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement des zones :")
                            list = Nothing
                            list = myxml.SelectNodes("/homidom/zones/zone")
                            If list.Count > 0 Then 'présence des zones
                                For i As Integer = 0 To list.Count - 1
                                    Dim x As New Zone
                                    For j As Integer = 0 To list.Item(i).Attributes.Count - 1
                                        Select Case list.Item(i).Attributes.Item(j).Name
                                            Case "id"
                                                x.ID = list.Item(i).Attributes.Item(j).Value
                                            Case "name"
                                                x.Name = list.Item(i).Attributes.Item(j).Value
                                            Case "icon"
                                                If list.Item(i).Attributes.Item(j).Value <> Nothing Then
                                                    If IO.File.Exists(list.Item(i).Attributes.Item(j).Value) Then
                                                        x.Icon = list.Item(i).Attributes.Item(j).Value
                                                    Else
                                                        x.Icon = _MonRepertoire & "\images\Zones\icon\defaut.png"
                                                    End If
                                                Else
                                                    x.Icon = _MonRepertoire & "\images\Zones\icon\defaut.png"
                                                End If
                                            Case "image"
                                                If list.Item(i).Attributes.Item(j).Value <> Nothing Then
                                                    If IO.File.Exists(list.Item(i).Attributes.Item(j).Value) Then
                                                        x.Image = list.Item(i).Attributes.Item(j).Value
                                                    Else
                                                        x.Image = _MonRepertoire & "\images\Zones\image\defaut.jpg"
                                                    End If
                                                Else
                                                    x.Image = _MonRepertoire & "\images\Zones\image\defaut.jpg"
                                                End If
                                            Case Else
                                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Un attribut correspondant à la zone est inconnu: nom:" & list.Item(i).Attributes.Item(j).Name & " Valeur: " & list.Item(0).Attributes.Item(j).Value)
                                        End Select
                                    Next
                                    If list.Item(i).HasChildNodes = True Then
                                        For k As Integer = 0 To list.Item(i).ChildNodes.Count - 1
                                            If list.Item(i).ChildNodes.Item(k).Name = "element" Then
                                                Dim _dev As New Zone.Element_Zone(list.Item(i).ChildNodes.Item(k).Attributes(0).Value, list.Item(i).ChildNodes.Item(k).Attributes(1).Value)
                                                x.ListElement.Add(_dev)
                                                _dev = Nothing
                                            End If
                                        Next
                                    End If
                                    _ListZones.Add(x)
                                    x = Nothing
                                Next
                            Else
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Aucune zone enregistrée dans le fichier de config")
                            End If
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> " & _ListZones.Count & " Zone(s) chargée(s)")
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement des zones : " & ex.Message)
                        End Try

                        '******************************************
                        'on va chercher les variables
                        '******************************************
                        Try
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement des variables :")
                            list = Nothing
                            list = myxml.SelectNodes("/homidom/variables/var")
                            If list.Count > 0 Then 'présence des variables
                                For i As Integer = 0 To list.Count - 1
                                    Dim x As New Variable(Me)
                                    For j As Integer = 0 To list.Item(i).Attributes.Count - 1
                                        Select Case list.Item(i).Attributes.Item(j).Name
                                            Case "id"
                                                x.ID = list.Item(i).Attributes.Item(j).Value
                                            Case "nom"
                                                x.Nom = list.Item(i).Attributes.Item(j).Value
                                            Case "enable"
                                                x.Enable = list.Item(i).Attributes.Item(j).Value
                                            Case "value"
                                                x.Value = list.Item(i).Attributes.Item(j).Value
                                            Case "description"
                                                x.Description = list.Item(i).Attributes.Item(j).Value
                                            Case Else
                                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Un attribut correspondant à la variable est inconnu: nom:" & list.Item(i).Attributes.Item(j).Name & " Valeur: " & list.Item(0).Attributes.Item(j).Value)
                                        End Select
                                    Next
                                    _ListVars.Add(x)
                                    x = Nothing
                                Next
                            End If
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> " & _ListVars.Count & " Variables(s) chargée(s)")
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement des variables : " & ex.Message)
                        End Try

                        '******************************************
                        'on va chercher les users
                        '******************************************
                        Try
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement des users :")
                            list = Nothing
                            list = myxml.SelectNodes("/homidom/users/user")
                            If list.Count > 0 Then 'présence des users
                                For i As Integer = 0 To list.Count - 1
                                    Dim x As New Users.User
                                    For j As Integer = 0 To list.Item(i).Attributes.Count - 1
                                        Select Case list.Item(i).Attributes.Item(j).Name
                                            Case "id"
                                                x.ID = list.Item(i).Attributes.Item(j).Value
                                            Case "username"
                                                x.UserName = list.Item(i).Attributes.Item(j).Value
                                            Case "nom"
                                                x.Nom = list.Item(i).Attributes.Item(j).Value
                                            Case "prenom"
                                                x.Prenom = list.Item(i).Attributes.Item(j).Value
                                            Case "profil"
                                                x.Profil = list.Item(i).Attributes.Item(j).Value
                                            Case "password"
                                                x.Password = list.Item(i).Attributes.Item(j).Value
                                            Case "numberidentification"
                                                x.NumberIdentification = list.Item(i).Attributes.Item(j).Value
                                            Case "image"
                                                If list.Item(i).Attributes.Item(j).Value <> Nothing Then
                                                    If IO.File.Exists(list.Item(i).Attributes.Item(j).Value) Then
                                                        x.Image = list.Item(i).Attributes.Item(j).Value
                                                    Else
                                                        x.Image = _MonRepertoire & "\images\icones\user_128.png"
                                                    End If
                                                Else
                                                    x.Image = _MonRepertoire & "\images\icones\user_128.png"
                                                End If
                                            Case "email"
                                                x.eMail = list.Item(i).Attributes.Item(j).Value
                                            Case "emailautre"
                                                x.eMailAutre = list.Item(i).Attributes.Item(j).Value
                                            Case "telfixe"
                                                x.TelFixe = list.Item(i).Attributes.Item(j).Value
                                            Case "telmobile"
                                                x.TelMobile = list.Item(i).Attributes.Item(j).Value
                                            Case "telautre"
                                                x.TelAutre = list.Item(i).Attributes.Item(j).Value
                                            Case "adresse"
                                                x.Adresse = list.Item(i).Attributes.Item(j).Value
                                            Case "ville"
                                                x.Ville = list.Item(i).Attributes.Item(j).Value
                                            Case "codepostal"
                                                x.CodePostal = list.Item(i).Attributes.Item(j).Value
                                            Case Else
                                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Un attribut correspondant à la zone est inconnu: nom:" & list.Item(i).Attributes.Item(j).Name & " Valeur: " & list.Item(0).Attributes.Item(j).Value)
                                        End Select
                                    Next
                                    _ListUsers.Add(x)
                                    x = Nothing
                                Next
                            Else
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Création de l'utilisateur admin par défaut !!")
                                SaveUser(_IdSrv, "", "Admin", "password", Users.TypeProfil.admin, "Administrateur", "Admin")
                            End If
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> " & _ListUsers.Count & " Utilisateur(s) chargé(s)")
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement des utilisateurs : " & ex.Message)
                        End Try


                        ''------------
                        ''on va chercher les nouveaux composants
                        ''------------
                        Try
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement des nouveaux devices détectés:")
                            list = Nothing
                            list = myxml.SelectNodes("/homidom/newdevices/newdevice")
                            If list.Count > 0 Then 'présence des devices
                                For i As Integer = 0 To list.Count - 1
                                    Dim x As New NewDevice
                                    For j As Integer = 0 To list.Item(i).Attributes.Count - 1
                                        Select Case list.Item(i).Attributes.Item(j).Name
                                            Case "id"
                                                x.ID = list.Item(i).Attributes.Item(j).Value
                                            Case "iddriver"
                                                x.IdDriver = list.Item(i).Attributes.Item(j).Value
                                            Case "adresse1"
                                                x.Adresse1 = list.Item(i).Attributes.Item(j).Value
                                            Case "adresse2"
                                                x.Adresse2 = list.Item(i).Attributes.Item(j).Value
                                            Case "name"
                                                x.Name = list.Item(i).Attributes.Item(j).Value
                                            Case "type"
                                                x.Type = list.Item(i).Attributes.Item(j).Value
                                            Case "ignore"
                                                x.Ignore = list.Item(i).Attributes.Item(j).Value
                                            Case "value"
                                                x.Value = list.Item(i).Attributes.Item(j).Value
                                            Case "datetetect"
                                                x.DateTetect = list.Item(i).Attributes.Item(j).Value
                                            Case Else
                                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Un attribut correspondant au nouveau device est inconnu: nom:" & list.Item(i).Attributes.Item(j).Name & " Valeur: " & list.Item(0).Attributes.Item(j).Value)
                                        End Select
                                    Next
                                    _ListNewDevices.Add(x)
                                    x = Nothing
                                Next
                            End If
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> " & _ListNewDevices.Count & " nouveau(x) chargé(s)")
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement des nouveaux devices : " & ex.Message)
                        End Try


                        '********************************************
                        'on va chercher les composants
                        '********************************************
                        Try
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement des composants :")
                            list = Nothing
                            list = myxml.SelectNodes("/homidom/devices/device")

                            Dim trvSoleil As Boolean = False
                            Dim trvStartSrv As Boolean = False
                            Dim trvnrjtot As Boolean = False
                            Dim trvisweekend As Boolean = False
                            Dim trvsaint As Boolean = False

                            If list.Count > 0 Then 'présence d'un composant
                                For j As Integer = 0 To list.Count - 1
                                    Dim _Dev As Object = Nothing

                                    'Suivant chaque type de device
                                    Select Case UCase(list.Item(j).Attributes.GetNamedItem("type").Value)
                                        Case "APPAREIL"
                                            Dim o As New Device.APPAREIL(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "AUDIO"
                                            Dim o As New Device.AUDIO(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "BAROMETRE"
                                            Dim o As New Device.BAROMETRE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "BATTERIE"
                                            Dim o As New Device.BATTERIE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "COMPTEUR"
                                            Dim o As New Device.COMPTEUR(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "CONTACT"
                                            Dim o As New Device.CONTACT(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "DETECTEUR"
                                            Dim o As New Device.DETECTEUR(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "DIRECTIONVENT"
                                            Dim o As New Device.DIRECTIONVENT(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "ENERGIEINSTANTANEE"
                                            Dim o As New Device.ENERGIEINSTANTANEE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "ENERGIETOTALE"
                                            Dim o As New Device.ENERGIETOTALE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "FREEBOX"
                                            Dim o As New Device.FREEBOX(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "GENERIQUEBOOLEEN"
                                            Dim o As New Device.GENERIQUEBOOLEEN(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "GENERIQUESTRING"
                                            Dim o As New Device.GENERIQUESTRING(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "GENERIQUEVALUE"
                                            Dim o As New Device.GENERIQUEVALUE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "HUMIDITE"
                                            Dim o As New Device.HUMIDITE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "LAMPE"
                                            Dim o As New Device.LAMPE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "LAMPERGBW"
                                            Dim o As New Device.LAMPERGBW(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "METEO"
                                            Dim o As New Device.METEO(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "MULTIMEDIA"
                                            Dim o As New Device.MULTIMEDIA(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "PLUIECOURANT"
                                            Dim o As New Device.PLUIECOURANT(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "PLUIETOTAL"
                                            Dim o As New Device.PLUIETOTAL(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "SWITCH"
                                            Dim o As New Device.SWITCH(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "TELECOMMANDE"
                                            Dim o As New Device.TELECOMMANDE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "TEMPERATURE"
                                            Dim o As New Device.TEMPERATURE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "TEMPERATURECONSIGNE"
                                            Dim o As New Device.TEMPERATURECONSIGNE(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "UV"
                                            Dim o As New Device.UV(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "VITESSEVENT"
                                            Dim o As New Device.VITESSEVENT(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                        Case "VOLET"
                                            Dim o As New Device.VOLET(Me)
                                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                                            _Dev = o
                                            o = Nothing
                                    End Select

                                    With _Dev
                                        'Affectation des valeurs sur les propriétés génériques
                                        If (Not list.Item(j).Attributes.GetNamedItem("id") Is Nothing) Then .ID = list.Item(j).Attributes.GetNamedItem("id").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("name") Is Nothing) Then .Name = list.Item(j).Attributes.GetNamedItem("name").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("enable") Is Nothing) Then .Enable = list.Item(j).Attributes.GetNamedItem("enable").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("driverid") Is Nothing) Then .DriverId = list.Item(j).Attributes.GetNamedItem("driverid").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("description") Is Nothing) Then .Description = list.Item(j).Attributes.GetNamedItem("description").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("adresse1") Is Nothing) Then .Adresse1 = list.Item(j).Attributes.GetNamedItem("adresse1").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("adresse2") Is Nothing) Then .Adresse2 = list.Item(j).Attributes.GetNamedItem("adresse2").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("datecreated") Is Nothing) Then .DateCreated = list.Item(j).Attributes.GetNamedItem("datecreated").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("lastchange") Is Nothing) Then .LastChange = list.Item(j).Attributes.GetNamedItem("lastchange").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("lastchangeduree") Is Nothing) Then .LastChangeDuree = list.Item(j).Attributes.GetNamedItem("lastchangeduree").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("refresh") Is Nothing) Then .Refresh = list.Item(j).Attributes.GetNamedItem("refresh").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("counthisto") Is Nothing) Then
                                            .Counthisto = list.Item(j).Attributes.GetNamedItem("counthisto").Value
                                        Else
                                            .CountHisto = DeviceAsHisto(.ID)
                                        End If
                                        If (Not list.Item(j).Attributes.GetNamedItem("modele") Is Nothing) Then .Modele = list.Item(j).Attributes.GetNamedItem("modele").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("allvalue") Is Nothing) Then .AllValue = list.Item(j).Attributes.GetNamedItem("allvalue").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("unit") Is Nothing) Then .Unit = list.Item(j).Attributes.GetNamedItem("unit").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("puissance") Is Nothing) Then .Puissance = list.Item(j).Attributes.GetNamedItem("puissance").Value
                                        If list.Item(j).Attributes.GetNamedItem("picture").Value IsNot Nothing Then
                                            If IO.File.Exists(list.Item(j).Attributes.GetNamedItem("picture").Value) Then
                                                .Picture = list.Item(j).Attributes.GetNamedItem("picture").Value
                                            Else
                                                Dim fileimg As String = _MonRepertoire & "\images\Devices\" & LCase(_Dev.type) & "-defaut.png"
                                                If IO.File.Exists(fileimg) Then
                                                    .Picture = fileimg
                                                Else
                                                    .Picture = _MonRepertoire & "\images\icones\composant_128.png"
                                                End If
                                                fileimg = Nothing
                                            End If
                                        Else
                                            .Picture = _MonRepertoire & "\images\icones\composant_128.png"
                                        End If
                                        If (Not list.Item(j).Attributes.GetNamedItem("solo") Is Nothing) Then .Solo = list.Item(j).Attributes.GetNamedItem("solo").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("lastetat") Is Nothing) Then .LastEtat = list.Item(j).Attributes.GetNamedItem("lastetat").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("ishisto") Is Nothing) Then .IsHisto = list.Item(j).Attributes.GetNamedItem("ishisto").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("refreshhisto") Is Nothing) Then .refreshhisto = list.Item(j).Attributes.GetNamedItem("refreshhisto").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("purge") Is Nothing) Then .purge = list.Item(j).Attributes.GetNamedItem("purge").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("moyjour") Is Nothing) Then .moyjour = list.Item(j).Attributes.GetNamedItem("moyjour").Value
                                        If (Not list.Item(j).Attributes.GetNamedItem("moyheure") Is Nothing) Then .moyheure = list.Item(j).Attributes.GetNamedItem("moyheure").Value

                                        'on recup les variables
                                        If list.Item(j).HasChildNodes = True Then
                                            For k As Integer = 0 To list.Item(j).ChildNodes.Count - 1
                                                If list.Item(j).ChildNodes.Item(k).Name = "var" Then
                                                    .Variables.Add(list.Item(j).ChildNodes.Item(k).Attributes(0).Value, list.Item(j).ChildNodes.Item(k).Attributes(1).Value)
                                                End If
                                            Next
                                        End If

                                        '-- propriétés generique value --
                                        If _Dev.Type = "BAROMETRE" _
                                        Or _Dev.Type = "COMPTEUR" _
                                        Or _Dev.Type = "ENERGIEINSTANTANEE" _
                                        Or _Dev.Type = "ENERGIETOTALE" _
                                        Or _Dev.Type = "GENERIQUEVALUE" _
                                        Or _Dev.Type = "HUMIDITE" _
                                        Or _Dev.Type = "LAMPE" _
                                        Or _Dev.Type = "LAMPERGBW" _
                                        Or _Dev.Type = "PLUIECOURANT" _
                                        Or _Dev.Type = "PLUIETOTAL" _
                                        Or _Dev.Type = "TEMPERATURE" _
                                        Or _Dev.Type = "TEMPERATURECONSIGNE" _
                                        Or _Dev.Type = "VITESSEVENT" _
                                        Or _Dev.Type = "UV" _
                                        Or _Dev.Type = "VOLET" _
                                        Then
                                            If (Not list.Item(j).Attributes.GetNamedItem("valuemin") Is Nothing) Then .ValueMin = Regex.Replace(list.Item(j).Attributes.GetNamedItem("valuemin").Value, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
                                            If (Not list.Item(j).Attributes.GetNamedItem("valuemax") Is Nothing) Then .ValueMax = Regex.Replace(list.Item(j).Attributes.GetNamedItem("valuemax").Value, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
                                            If (Not list.Item(j).Attributes.GetNamedItem("valuedef") Is Nothing) Then .ValueDef = Regex.Replace(list.Item(j).Attributes.GetNamedItem("valuedef").Value, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
                                            If (Not list.Item(j).Attributes.GetNamedItem("precision") Is Nothing) Then .Precision = Regex.Replace(list.Item(j).Attributes.GetNamedItem("precision").Value, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
                                            If (Not list.Item(j).Attributes.GetNamedItem("correction") Is Nothing) Then .Correction = list.Item(j).Attributes.GetNamedItem("correction").Value
                                            If (Not list.Item(j).Attributes.GetNamedItem("formatage") Is Nothing) Then .Formatage = list.Item(j).Attributes.GetNamedItem("formatage").Value
                                        End If
                                        If (Not list.Item(j).Attributes.GetNamedItem("value") Is Nothing) Then .Value = Regex.Replace(list.Item(j).Attributes.GetNamedItem("value").Value, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)
                                        If (Not list.Item(j).Attributes.GetNamedItem("valuelast") Is Nothing) Then .ValueLast = Regex.Replace(list.Item(j).Attributes.GetNamedItem("valuelast").Value, "[.,]", System.Globalization.NumberFormatInfo.CurrentInfo.NumberDecimalSeparator)

                                        If _Dev.Type = "LAMPERGBW" Then
                                            If (Not list.Item(j).Attributes.GetNamedItem("red") Is Nothing) Then .red = list.Item(j).Attributes.GetNamedItem("red").Value
                                            If (Not list.Item(j).Attributes.GetNamedItem("green") Is Nothing) Then .green = list.Item(j).Attributes.GetNamedItem("green").Value
                                            If (Not list.Item(j).Attributes.GetNamedItem("blue") Is Nothing) Then .blue = list.Item(j).Attributes.GetNamedItem("blue").Value
                                            If (Not list.Item(j).Attributes.GetNamedItem("white") Is Nothing) Then .white = list.Item(j).Attributes.GetNamedItem("white").Value
                                            If (Not list.Item(j).Attributes.GetNamedItem("temperature") Is Nothing) Then .temperature = list.Item(j).Attributes.GetNamedItem("temperature").Value
                                            If (Not list.Item(j).Attributes.GetNamedItem("speed") Is Nothing) Then .speed = list.Item(j).Attributes.GetNamedItem("speed").Value
                                            If (Not list.Item(j).Attributes.GetNamedItem("optionnal") Is Nothing) Then .Formatage = list.Item(j).Attributes.GetNamedItem("optionnal").Value
                                        End If

                                        'Verifie si prob de MaJ
                                        If .LastChangeDuree > 0 Then
                                            If DateTime.Compare(.LastChange.AddMinutes(CInt(.LastChangeDuree)), Now) < 0 Then
                                                _DevicesNoMAJ.Add(.Name)
                                            End If
                                        End If

                                        If String.IsNullOrEmpty(.ID) = False And String.IsNullOrEmpty(.Name) = False And String.IsNullOrEmpty(.Adresse1) = False And String.IsNullOrEmpty(.DriverId) = False Then
                                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " - " & .Name & " (" & .ID & " - " & .Adresse1 & " - " & .Type & ") --> " & .Value)
                                            If .ID = "soleil01" Then
                                                trvSoleil = True
                                            End If
                                            If .ID = "startsrv01" Then
                                                trvStartSrv = True
                                            End If
                                            If .ID = "energietotale01" Then
                                                trvnrjtot = True
                                            End If
                                            If .ID = "isweekend01" Then
                                                trvisweekend = True
                                            End If
                                            If .ID = "saint01" Then
                                                trvsaint = True
                                            End If
                                        Else
                                            _Dev.Enable = False
                                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "LoadConfig", " -> Erreur lors du chargement du composant (information incomplete -> Disable) " & .Name & " (" & .ID & " - " & .Adresse1 & " - " & .Type & ")")
                                        End If
                                    End With
                                    _ListDevices.Add(_Dev)

                                    _Dev = Nothing
                                Next
                            Else
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Aucun composant enregistré dans le fichier de config ceux par défaut vont être créés")
                            End If
                            If trvSoleil = False Then
                                Dim _Devs As New Device.GENERIQUEBOOLEEN(Me)
                                _Devs.ID = "soleil01"
                                _Devs.Name = "HOMI_Jour"
                                _Devs.Enable = True
                                _Devs.Adresse1 = "N/A"
                                _Devs.Description = "Levé/Couché du soleil : True si il fait jour"
                                _Devs.DriverID = "DE96B466-2540-11E0-A321-65D7DFD72085"
                                _ListDevices.Add(_Devs)
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " - " & _Devs.Name & " (" & _Devs.ID & " - " & _Devs.Adresse1 & " - " & _Devs.Type & ")")
                                _Devs = Nothing
                            End If
                            If trvStartSrv = False Then
                                Dim _Devs As New Device.GENERIQUEBOOLEEN(Me)
                                _Devs.ID = "startsrv01"
                                _Devs.Name = "HOMI_StartServeur"
                                _Devs.Enable = True
                                _Devs.Adresse1 = "N/A"
                                _Devs.AllValue = True
                                _Devs.Description = "Serveur Démarré"
                                _Devs.DriverID = "DE96B466-2540-11E0-A321-65D7DFD72085"
                                _ListDevices.Add(_Devs)
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " - " & _Devs.Name & " (" & _Devs.ID & " - " & _Devs.Adresse1 & " - " & _Devs.Type & ")")
                                _Devs = Nothing
                            End If
                            If trvnrjtot = False Then
                                Dim _Devs As New Device.GENERIQUEVALUE(Me)
                                _Devs.ID = "energietotale01"
                                _Devs.Name = "HOMI_EnergieTotaleInstantanee"
                                _Devs.Enable = True
                                _Devs.Adresse1 = "N/A"
                                _Devs.AllValue = True
                                _Devs.Description = "Energie Totale instantanee"
                                _Devs.DriverID = "DE96B466-2540-11E0-A321-65D7DFD72085"
                                _ListDevices.Add(_Devs)
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " - " & _Devs.Name & " (" & _Devs.ID & " - " & _Devs.Adresse1 & " - " & _Devs.Type & ")")
                                _Devs = Nothing
                            End If
                            If trvisweekend = False Then
                                Dim _Devs As New Device.GENERIQUEBOOLEEN(Me)
                                _Devs.ID = "isweekend01"
                                _Devs.Name = "HOMI_IsWeekend"
                                _Devs.Enable = True
                                _Devs.Adresse1 = "N/A"
                                _Devs.AllValue = True
                                _Devs.Description = "Permet de savoir si c'est le week-end (samedi/dimanche)"
                                _Devs.DriverID = "DE96B466-2540-11E0-A321-65D7DFD72085"
                                _ListDevices.Add(_Devs)
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " - " & _Devs.Name & " (" & _Devs.ID & " - " & _Devs.Adresse1 & " - " & _Devs.Type & ")")
                                _Devs = Nothing
                            End If
                            If trvsaint = False Then
                                Dim _Devs As New Device.GENERIQUESTRING(Me)
                                _Devs.ID = "saint01"
                                _Devs.Name = "HOMI_Saint"
                                _Devs.Enable = True
                                _Devs.Adresse1 = "N/A"
                                _Devs.AllValue = True
                                _Devs.Description = "Permet de connaitre le saint du jour"
                                _Devs.DriverID = "DE96B466-2540-11E0-A321-65D7DFD72085"
                                _ListDevices.Add(_Devs)
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " - " & _Devs.Name & " (" & _Devs.ID & " - " & _Devs.Adresse1 & " - " & _Devs.Type & ")")
                                _Devs = Nothing
                            End If
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> " & _ListDevices.Count & " composant(s) trouvé(s)")
                            list = Nothing
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement des Composants : " & ex.Message)
                        End Try

                        '******************************************
                        'on va chercher les triggers
                        '******************************************
                        Try
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement des triggers :")
                            list = Nothing
                            list = myxml.SelectNodes("/homidom/triggers/trigger")
                            If list.Count > 0 Then 'présence des triggers
                                For i As Integer = 0 To list.Count - 1
                                    Dim x As New Trigger
                                    x._Server = Me
                                    For j1 As Integer = 0 To list.Item(i).Attributes.Count - 1
                                        Select Case list.Item(i).Attributes.Item(j1).Name
                                            Case "id" : x.ID = list.Item(i).Attributes.Item(j1).Value
                                            Case "nom" : x.Nom = list.Item(i).Attributes.Item(j1).Value
                                            Case "enable" : x.Enable = list.Item(i).Attributes.Item(j1).Value
                                            Case "type"
                                                If list.Item(i).Attributes.Item(j1).Value = "0" Then
                                                    x.Type = Trigger.TypeTrigger.TIMER
                                                Else
                                                    x.Type = Trigger.TypeTrigger.DEVICE
                                                End If
                                            Case "description" : If list.Item(i).Attributes.Item(j1).Value <> Nothing Then x.Description = list.Item(i).Attributes.Item(j1).Value
                                            Case "conditiontime" : If list.Item(i).Attributes.Item(j1).Value <> Nothing Then x.ConditionTime = list.Item(i).Attributes.Item(j1).Value
                                            Case "conditiondeviceid" : If list.Item(i).Attributes.Item(j1).Value <> Nothing Then x.ConditionDeviceId = list.Item(i).Attributes.Item(j1).Value
                                            Case "conditiondeviceproperty" : If list.Item(i).Attributes.Item(j1).Value <> Nothing Then x.ConditionDeviceProperty = list.Item(i).Attributes.Item(j1).Value
                                            Case "prochainedateheure" : If list.Item(i).Attributes.Item(j1).Value <> Nothing Then x.Prochainedateheure = list.Item(i).Attributes.Item(j1).Value
                                            Case Else : Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Un attribut correspondant au trigger est inconnu: nom:" & list.Item(i).Attributes.Item(j1).Name & " Valeur: " & list.Item(0).Attributes.Item(j1).Value)
                                        End Select
                                    Next
                                    If list.Item(i).HasChildNodes = True Then
                                        If list.Item(i).ChildNodes.Item(0).Name = "macros" And list.Item(i).ChildNodes.Item(0).HasChildNodes Then
                                            For k9 As Integer = 0 To list.Item(i).ChildNodes.Item(0).ChildNodes.Count - 1
                                                If list.Item(i).ChildNodes.Item(0).ChildNodes.Item(k9).Name = "macro" Then
                                                    If list.Item(i).ChildNodes.Item(0).ChildNodes.Item(k9).Attributes.Count > 0 And list.Item(i).ChildNodes.Item(0).ChildNodes.Item(k9).Attributes.Item(0).Name = "id" Then
                                                        x.ListMacro.Add(list.Item(i).ChildNodes.Item(0).ChildNodes.Item(k9).Attributes.Item(0).Value)
                                                    End If
                                                End If
                                            Next
                                        End If
                                    End If
                                    _ListTriggers.Add(x)
                                    x = Nothing
                                Next
                            Else
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Aucun trigger enregistré dans le fichier de config")
                            End If
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> " & _ListTriggers.Count & " Trigger(s) chargé(s)")
                            list = Nothing
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement des triggers : " & ex.Message)
                        End Try

                        '******************************************
                        'on va chercher les macros
                        '******************************************
                        Try
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", "Chargement des macros :")
                            list = Nothing
                            list = myxml.SelectNodes("/homidom/macros/macro")
                            If list.Count > 0 Then 'présence des macros
                                For i As Integer = 0 To list.Count - 1
                                    Dim x As New Macro
                                    For j1 As Integer = 0 To list.Item(i).Attributes.Count - 1
                                        Select Case list.Item(i).Attributes.Item(j1).Name
                                            Case "id" : x.ID = list.Item(i).Attributes.Item(j1).Value
                                            Case "nom" : x.Nom = list.Item(i).Attributes.Item(j1).Value
                                            Case "enable" : x.Enable = list.Item(i).Attributes.Item(j1).Value
                                                'Case "description" : If list.Item(i).Attributes.Item(j1).Value <> Nothing Then x.Description = list.Item(0).Attributes.Item(j1).Value
                                            Case "description" : x.Description = list.Item(0).Attributes.Item(j1).Value
                                            Case Else : Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Un attribut correspondant à la macro est inconnu: nom:" & list.Item(i).Attributes.Item(j1).Name & " Valeur: " & list.Item(0).Attributes.Item(j1).Value)
                                        End Select
                                    Next
                                    LoadAction(list.Item(i), x.ListActions)
                                    _ListMacros.Add(x)
                                    x = Nothing
                                Next
                            Else
                                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> Aucune macro enregistrée dans le fichier de config")
                            End If
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", " -> " & _ListMacros.Count & " Macro(s) chargée(s)")
                            list = Nothing
                        Catch ex As Exception
                            Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "LoadConfig", "Erreur lors du chargement des macros : " & ex.Message)
                        End Try

                        Exit For
                    Next
                Else
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "LoadConfig", "Fichier de configuration non trouvé")
                End If

                _Finish = True

                'Vide les variables
                dirInfo = Nothing
                file = Nothing
                files = Nothing
                myxml = Nothing
                myfile = Nothing
                Return " Chargement de la configuration terminée"

            Catch ex As Exception
                Return " Erreur de chargement de la config: " & ex.Message
            End Try
        End Function

        Private Sub LoadAction(ByVal list As XmlNode, ByVal ListAction As ArrayList)
            Dim _typeaction As String = ""

            Try
                If list.HasChildNodes Then
                    For j2 As Integer = 0 To list.ChildNodes.Count - 1
                        If list.ChildNodes.Item(j2).Name = "action" Then
                            Dim _Act As Object = Nothing
                            Select Case list.ChildNodes.Item(j2).Attributes.Item(0).Value
                                Case "ActionDevice"
                                    Dim o As New Action.ActionDevice
                                    _Act = o
                                    o = Nothing
                                Case "ActionDriver"
                                    Dim o As New Action.ActionDriver
                                    _Act = o
                                    o = Nothing
                                Case "ActionMail"
                                    Dim o As New Action.ActionMail
                                    _Act = o
                                    o = Nothing
                                Case "ActionIf"
                                    Dim o As New Action.ActionIf
                                    _Act = o
                                    o = Nothing
                                Case "ActionMacro"
                                    Dim o As New Action.ActionMacro
                                    _Act = o
                                    o = Nothing
                                Case "ActionSpeech"
                                    Dim o As New Action.ActionSpeech
                                    _Act = o
                                    o = Nothing
                                Case "ActionHttp"
                                    Dim o As New Action.ActionHttp
                                    _Act = o
                                    o = Nothing
                                Case "ActionLogEvent"
                                    Dim o As New Action.ActionLogEvent
                                    _Act = o
                                    _typeaction = "ActionLogEvent"
                                    o = Nothing
                                Case "ActionLogEventHomidom"
                                    Dim o As New Action.ActionLogEventHomidom
                                    _Act = o
                                    _typeaction = "ActionLogEventHomidom"
                                    o = Nothing
                                Case "ActionDOS"
                                    Dim o As New Action.ActionDos
                                    _Act = o
                                    o = Nothing
                                Case "ActionVB"
                                    Dim o As New Action.ActionVB
                                    _Act = o
                                    o = Nothing
                                Case "ActionStop"
                                    Dim o As New Action.ActionSTOP
                                    _Act = o
                                    o = Nothing
                                Case "ActionVar"
                                    Dim o As New Action.ActionVar
                                    _Act = o
                                    o = Nothing
                            End Select
                            For j3 As Integer = 0 To list.ChildNodes.Item(j2).Attributes.Count - 1
                                Select Case list.ChildNodes.Item(j2).Attributes.Item(j3).Name
                                    Case "timing"
                                        _Act.timing = CDate(list.ChildNodes.Item(j2).Attributes.Item(j3).Value)
                                    Case "iddevice"
                                        _Act.iddevice = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "iddriver"
                                        _Act.iddriver = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "idmacro"
                                        _Act.idmacro = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "method"
                                        _Act.method = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "userid"
                                        _Act.userid = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "sujet"
                                        _Act.sujet = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "message"
                                        _Act.message = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "commande"
                                        _Act.Commande = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "fonction"
                                        _Act.fonction = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "fichier"
                                        _Act.Fichier = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "arguments"
                                        _Act.Arguments = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "label"
                                        _Act.label = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "script"
                                        _Act.Script = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "parametres"
                                        Dim b As String = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                        Dim a() As String = b.Split("|")
                                        Dim c As New ArrayList
                                        For cnt1 As Integer = 0 To a.Count - 1
                                            c.Add(a(cnt1))
                                        Next
                                        _Act.parametres = c
                                        b = Nothing
                                        a = Nothing
                                        c = Nothing
                                    Case "type"
                                        If _typeaction = "ActionLogEvent" Then
                                            Select Case UCase(list.ChildNodes.Item(j2).Attributes.Item(j3).Value)
                                                Case "ERREUR" : _Act.Type = 1
                                                Case "WARNING" : _Act.Type = 2
                                                Case "INFORMATION" : _Act.Type = 3
                                                Case Else : _Act.Type = 1
                                            End Select
                                        End If
                                        If _typeaction = "ActionLogEventHomidom" Then
                                            Select Case UCase(list.ChildNodes.Item(j2).Attributes.Item(j3).Value)
                                                Case "INFO" : _Act.Type = 1
                                                Case "ACTION" : _Act.Type = 2
                                                Case "MESSAGE" : _Act.Type = 3
                                                Case "VALEUR_CHANGE" : _Act.Type = 4
                                                Case "VALEUR_INCHANGE" : _Act.Type = 5
                                                Case "VALEUR_INCHANGE_PRECISION" : _Act.Type = 6
                                                Case "VALEUR_INCHANGE_LASTETAT" : _Act.Type = 7
                                                Case "ERREUR" : _Act.Type = 8
                                                Case "ERREUR_CRITIQUE" : _Act.Type = 9
                                                Case "DEBUG" : _Act.Type = 10
                                            End Select
                                        End If
                                    Case "eventid"
                                        _Act.Eventid = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "nom"
                                        _Act.Nom = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                    Case "value"
                                        _Act.Value = list.ChildNodes.Item(j2).Attributes.Item(j3).Value
                                End Select
                            Next
                            If list.ChildNodes.Item(j2).HasChildNodes Then
                                For j3 As Integer = 0 To list.ChildNodes.Item(j2).ChildNodes.Count - 1
                                    If list.ChildNodes.Item(j2).ChildNodes.Item(j3).Name = "conditions" Then
                                        For j4 As Integer = 0 To list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Count - 1
                                            Dim Condi As New Action.Condition
                                            For j5 As Integer = 0 To list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Count - 1
                                                Select Case list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Item(j5).Name
                                                    Case "typecondition"
                                                        Select Case list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Item(j5).Value
                                                            Case Action.TypeCondition.DateTime.ToString
                                                                Condi.Type = Action.TypeCondition.DateTime
                                                            Case Action.TypeCondition.Device.ToString
                                                                Condi.Type = Action.TypeCondition.Device
                                                        End Select
                                                    Case "datetime"
                                                        Condi.DateTime = list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Item(j5).Value
                                                    Case "iddevice"
                                                        Condi.IdDevice = list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Item(j5).Value
                                                    Case "propertydevice"
                                                        Condi.PropertyDevice = list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Item(j5).Value
                                                    Case "value"
                                                        Condi.Value = list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Item(j5).Value
                                                    Case "condition"
                                                        Select Case list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Item(j5).Value
                                                            Case Action.TypeSigne.Different.ToString
                                                                Condi.Condition = Action.TypeSigne.Different
                                                            Case Action.TypeSigne.Egal.ToString
                                                                Condi.Condition = Action.TypeSigne.Egal
                                                            Case Action.TypeSigne.Inferieur.ToString
                                                                Condi.Condition = Action.TypeSigne.Inferieur
                                                            Case Action.TypeSigne.InferieurEgal.ToString
                                                                Condi.Condition = Action.TypeSigne.InferieurEgal
                                                            Case Action.TypeSigne.Superieur.ToString
                                                                Condi.Condition = Action.TypeSigne.Superieur
                                                            Case Action.TypeSigne.SuperieurEgal.ToString
                                                                Condi.Condition = Action.TypeSigne.SuperieurEgal
                                                        End Select
                                                    Case "operateur"
                                                        Select Case list.ChildNodes.Item(j2).ChildNodes.Item(j3).ChildNodes.Item(j4).Attributes.Item(j5).Value
                                                            Case Action.TypeOperateur.NONE.ToString
                                                                Condi.Operateur = Action.TypeOperateur.NONE
                                                            Case Action.TypeOperateur.AND.ToString
                                                                Condi.Operateur = Action.TypeOperateur.AND
                                                            Case Action.TypeOperateur.OR.ToString
                                                                Condi.Operateur = Action.TypeOperateur.OR
                                                        End Select
                                                End Select
                                            Next
                                            _Act.Conditions.add(Condi)
                                        Next
                                    End If
                                    If list.ChildNodes.Item(j2).ChildNodes.Item(j3).Name = "then" Then
                                        LoadAction(list.ChildNodes.Item(j2).ChildNodes.Item(j3), _Act.ListTrue)
                                    End If
                                    If list.ChildNodes.Item(j2).ChildNodes.Item(j3).Name = "else" Then
                                        LoadAction(list.ChildNodes.Item(j2).ChildNodes.Item(j3), _Act.ListFalse)
                                    End If
                                Next
                            End If
                            ListAction.Add(_Act)
                        End If
                    Next
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "LoadAction", ex.message)
            End Try
        End Sub

        ''' <summary>Sauvegarde de la config dans le fichier XML</summary>
        ''' <remarks></remarks>
        Private Function SaveConfig(ByVal Fichier As String) As Boolean
            Try
                Log(TypeLog.INFO, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde de la config sous le fichier " & Fichier)

                ''Copy du fichier de config avant sauvegarde
                Try
                    If _SaveDiffBackup = False Then
                        If IO.File.Exists(Fichier.Replace(".xml", ".sav")) = True Then IO.File.Delete(Fichier.Replace(".xml", ".sav"))
                        IO.File.Copy(Fichier, Fichier.Replace(".xml", ".sav"))
                    Else
                        Dim fich As String = Fichier.Replace(".xml", ".sav")
                        fich = fich.Replace(".", Now.Year & Now.Month & Now.Day & Now.Hour & Now.Minute & Now.Second & ".")
                        IO.File.Copy(Fichier, fich)
                    End If
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "SaveConfig", "Création de sauvegarde (.sav) du fichier de config avant sauvegarde")
                Catch ex As Exception
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", "Erreur impossible de créer une copie de backup du fichier de config: " & ex.Message)
                End Try


                ''Copy du fichier de config avant sauvegarde
                'Try
                '    Dim _file As String = Fichier.Replace(".xml", "")
                '    If IO.File.Exists(_file & ".sav") = True Then IO.File.Delete(_file & ".sav")
                '    IO.File.Copy(_file & ".xml", _file & ".sav")
                '    Log(TypeLog.DEBUG, TypeSource.SERVEUR, "LoadConfig", "Création de sauvegarde (.sav) du fichier de config avant sauvegarde")
                'Catch ex As Exception
                '    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", "Erreur impossible de créer une copie de backup du fichier de config: " & ex.Message)
                'End Try

                ''Creation du fichier XML
                Dim writer As New XmlTextWriter(Fichier, System.Text.Encoding.UTF8)
                writer.WriteStartDocument(True)
                writer.Formatting = Formatting.Indented
                writer.Indentation = 2

                writer.WriteStartElement("homidom")

                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des paramètres serveur")
                ''------------ server
                writer.WriteStartElement("server")
                writer.WriteStartAttribute("ipsoap")
                writer.WriteValue(_IPSOAP)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("portsoap")
                writer.WriteValue(_PortSOAP)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("idsrv")
                writer.WriteValue(_IdSrv)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("savediff")
                writer.WriteValue(_SaveDiffBackup)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("longitude")
                writer.WriteValue(_Longitude)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("latitude")
                writer.WriteValue(_Latitude)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("heurecorrectionlever")
                writer.WriteValue(_HeureLeverSoleilCorrection)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("heurecorrectioncoucher")
                writer.WriteValue(_HeureCoucherSoleilCorrection)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("smtpserver")
                writer.WriteValue(_SMTPServeur)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("smtpmail")
                writer.WriteValue(_SMTPmailEmetteur)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("smtplogin")
                writer.WriteValue(_SMTPLogin)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("smtppassword")
                writer.WriteValue(_SMTPassword)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("smtpport")
                writer.WriteValue(_SMTPPort)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("smtpssl")
                writer.WriteValue(_SMTPSSL)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("logmaxfilesize")
                writer.WriteValue(_MaxFileSize)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("logmaxmonthlog")
                writer.WriteValue(_MaxMonthLog)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log0")
                writer.WriteValue(_TypeLogEnable(0))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log1")
                writer.WriteValue(_TypeLogEnable(1))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log2")
                writer.WriteValue(_TypeLogEnable(2))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log3")
                writer.WriteValue(_TypeLogEnable(3))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log4")
                writer.WriteValue(_TypeLogEnable(4))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log5")
                writer.WriteValue(_TypeLogEnable(5))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log6")
                writer.WriteValue(_TypeLogEnable(6))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log7")
                writer.WriteValue(_TypeLogEnable(7))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log8")
                writer.WriteValue(_TypeLogEnable(8))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("log9")
                writer.WriteValue(_TypeLogEnable(9))
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("cyclesave")
                writer.WriteValue(_CycleSave)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("voice")
                writer.WriteValue(_Voice)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("saverealtime")
                writer.WriteValue(_SaveRealTime)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("devise")
                writer.WriteValue(_Devise)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("puissancemini")
                writer.WriteValue(PuissanceMini)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("gererenergie")
                writer.WriteValue(GererEnergie)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("tarifjour")
                writer.WriteValue(_TarifJour)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("tarifnuit")
                writer.WriteValue(_TarifNuit)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("portweb")
                writer.WriteValue(_PortSrvWeb)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("enablesrvweb")
                writer.WriteValue(_EnableSrvWeb)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("modedecouverte")
                writer.WriteValue(_ModeDecouverte)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("codepays")
                writer.WriteValue(CodePays)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("foldersavefolder")
                writer.WriteValue(_FolderSaveFolder)
                writer.WriteEndAttribute()
                writer.WriteStartAttribute("cyclesavefolder")
                writer.WriteValue(_CycleSaveFolder)
                writer.WriteEndAttribute()
                writer.WriteEndElement()


                ''-------------------
                ''------------drivers
                ''------------------
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des drivers")
                writer.WriteStartElement("drivers")
                For i As Integer = 0 To _ListDrivers.Count - 1
                    Try
                        writer.WriteStartElement("driver")
                        writer.WriteStartAttribute("id")
                        writer.WriteValue(_ListDrivers.Item(i).ID)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("nom")
                        writer.WriteValue(_ListDrivers.Item(i).Nom)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("enable")
                        writer.WriteValue(_ListDrivers.Item(i).Enable)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("description")
                        writer.WriteValue(_ListDrivers.Item(i).Description)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("startauto")
                        writer.WriteValue(_ListDrivers.Item(i).StartAuto)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("protocol")
                        writer.WriteValue(_ListDrivers.Item(i).Protocol)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("iptcp")
                        writer.WriteValue(_ListDrivers.Item(i).IP_TCP)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("porttcp")
                        writer.WriteValue(_ListDrivers.Item(i).Port_TCP)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("ipudp")
                        writer.WriteValue(_ListDrivers.Item(i).IP_UDP)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("portudp")
                        writer.WriteValue(_ListDrivers.Item(i).Port_UDP)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("com")
                        writer.WriteValue(_ListDrivers.Item(i).Com)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("refresh")
                        writer.WriteValue(_ListDrivers.Item(i).Refresh)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("modele")
                        writer.WriteValue(_ListDrivers.Item(i).modele)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("autodiscover")
                        writer.WriteValue(_ListDrivers.Item(i).autodiscover)
                        writer.WriteEndAttribute()
                        If _ListDrivers.Item(i).Parametres IsNot Nothing Then
                            For j As Integer = 0 To _ListDrivers.Item(i).Parametres.count - 1
                                writer.WriteStartAttribute("parametre" & j)
                                If _ListDrivers.Item(i).Parametres.Item(j).valeur IsNot Nothing Then
                                    writer.WriteValue(_ListDrivers.Item(i).Parametres.Item(j).valeur)
                                Else
                                    writer.WriteValue(" ")
                                End If
                                writer.WriteEndAttribute()
                            Next
                        End If
                        writer.WriteEndElement()
                        'on met à jour l'ID du serveur dans les drivers au cas ou il aurait changé
                        Dim _drv As IDriver = ReturnDrvById(_IdSrv, _ListDrivers.Item(i).ID)
                        If _drv IsNot Nothing Then _drv.IdSrv = _IdSrv
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: Drivers: " & ex.ToString)
                    End Try
                Next
                writer.WriteEndElement()

                ''------------
                ''Sauvegarde des zones
                ''------------
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des zones")
                writer.WriteStartElement("zones")
                For i As Integer = 0 To _ListZones.Count - 1
                    Try
                        writer.WriteStartElement("zone")
                        writer.WriteStartAttribute("id")
                        writer.WriteValue(_ListZones.Item(i).ID)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("name")
                        writer.WriteValue(_ListZones.Item(i).Name)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("icon")
                        writer.WriteValue(_ListZones.Item(i).Icon)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("image")
                        writer.WriteValue(_ListZones.Item(i).Image)
                        writer.WriteEndAttribute()
                        If _ListZones.Item(i).ListElement IsNot Nothing Then
                            For j As Integer = 0 To _ListZones.Item(i).ListElement.Count - 1
                                writer.WriteStartElement("element")
                                writer.WriteStartAttribute("elementid")
                                writer.WriteValue(_ListZones.Item(i).ListElement.Item(j).ElementID)
                                writer.WriteEndAttribute()
                                writer.WriteStartAttribute("visible")
                                writer.WriteValue(_ListZones.Item(i).ListElement.Item(j).Visible)
                                writer.WriteEndAttribute()
                                writer.WriteEndElement()
                            Next
                        End If
                        writer.WriteEndElement()
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: Zones: " & ex.ToString)
                    End Try
                Next
                writer.WriteEndElement()

                ''------------
                ''Sauvegarde des variables
                ''------------
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des variables")
                writer.WriteStartElement("variables")
                For i As Integer = 0 To _ListVars.Count - 1
                    Try
                        writer.WriteStartElement("var")
                        writer.WriteStartAttribute("id")
                        writer.WriteValue(_ListVars.Item(i).ID)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("nom")
                        writer.WriteValue(_ListVars.Item(i).Nom)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("enable")
                        writer.WriteValue(_ListVars.Item(i).Enable)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("value")
                        writer.WriteValue(_ListVars.Item(i).Value)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("description")
                        writer.WriteValue(_ListVars.Item(i).Description)
                        writer.WriteEndAttribute()
                        writer.WriteEndElement()
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: Variables: " & ex.ToString)
                    End Try
                Next
                writer.WriteEndElement()

                ''------------
                ''Sauvegarde des users
                ''------------
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des users")
                writer.WriteStartElement("users")
                For i As Integer = 0 To _ListUsers.Count - 1
                    Try
                        writer.WriteStartElement("user")
                        writer.WriteStartAttribute("id")
                        writer.WriteValue(_ListUsers.Item(i).ID)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("username")
                        writer.WriteValue(_ListUsers.Item(i).UserName)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("nom")
                        writer.WriteValue(_ListUsers.Item(i).Nom)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("prenom")
                        writer.WriteValue(_ListUsers.Item(i).Prenom)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("profil")
                        Select Case _ListUsers.Item(i).Profil
                            Case Users.TypeProfil.invite
                                writer.WriteValue("0")
                            Case Users.TypeProfil.user
                                writer.WriteValue("1")
                            Case Users.TypeProfil.admin
                                writer.WriteValue("2")
                        End Select
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("password")
                        writer.WriteValue(_ListUsers.Item(i).Password)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("numberidentification")
                        writer.WriteValue(_ListUsers.Item(i).NumberIdentification)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("image")
                        writer.WriteValue(_ListUsers.Item(i).Image)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("email")
                        writer.WriteValue(_ListUsers.Item(i).eMail)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("emailautre")
                        writer.WriteValue(_ListUsers.Item(i).eMailAutre)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("telfixe")
                        writer.WriteValue(_ListUsers.Item(i).TelFixe)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("telmobile")
                        writer.WriteValue(_ListUsers.Item(i).TelMobile)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("telautre")
                        writer.WriteValue(_ListUsers.Item(i).TelAutre)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("adresse")
                        writer.WriteValue(_ListUsers.Item(i).Adresse)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("ville")
                        writer.WriteValue(_ListUsers.Item(i).Ville)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("codepostal")
                        writer.WriteValue(_ListUsers.Item(i).CodePostal)
                        writer.WriteEndAttribute()
                        writer.WriteEndElement()
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: Utilisateurs: " & ex.ToString)
                    End Try
                Next
                writer.WriteEndElement()


                ''------------
                ''Sauvegarde des devices
                ''------------
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des devices")
                writer.WriteStartElement("devices")
                For i As Integer = 0 To _ListDevices.Count - 1
                    Try
                        'Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", " - " & _ListDevices.Item(i).name)
                        writer.WriteStartElement("device")
                        '-- propriétés génériques --
                        writer.WriteStartAttribute("id")
                        writer.WriteValue(_ListDevices.Item(i).id)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("name")
                        writer.WriteValue(_ListDevices.Item(i).Name)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("enable")
                        writer.WriteValue(_ListDevices.Item(i).enable)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("driverid")
                        writer.WriteValue(_ListDevices.Item(i).driverid)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("description")
                        writer.WriteValue(_ListDevices.Item(i).description)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("type")
                        writer.WriteValue(_ListDevices.Item(i).type)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("adresse1")
                        writer.WriteValue(_ListDevices.Item(i).adresse1)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("adresse2")
                        writer.WriteValue(_ListDevices.Item(i).adresse2)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("datecreated")
                        writer.WriteValue(_ListDevices.Item(i).datecreated)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("lastchange")
                        writer.WriteValue(_ListDevices.Item(i).lastchange)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("lastchangeduree")
                        writer.WriteValue(_ListDevices.Item(i).LastChangeDuree)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("counthisto")
                        writer.WriteValue(_ListDevices.Item(i).CountHisto)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("refresh")
                        writer.WriteValue(_ListDevices.Item(i).refresh)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("modele")
                        writer.WriteValue(_ListDevices.Item(i).modele)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("allvalue")
                        writer.WriteValue(_ListDevices.Item(i).allvalue)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("unit")
                        writer.WriteValue(_ListDevices.Item(i).unit)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("puissance")
                        writer.WriteValue(_ListDevices.Item(i).Puissance)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("picture")
                        Dim _pict As String = _ListDevices.Item(i).picture
                        If String.IsNullOrEmpty(_pict) = True Or _pict = Nothing Then _pict = " "
                        writer.WriteValue(_pict)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("solo")
                        writer.WriteValue(_ListDevices.Item(i).solo)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("value")
                        If _ListDevices.Item(i).value IsNot Nothing Then writer.WriteValue(_ListDevices.Item(i).value)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("lastetat")
                        writer.WriteValue(_ListDevices.Item(i).lastetat)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("ishisto")
                        writer.WriteValue(_ListDevices.Item(i).isHisto)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("refreshhisto")
                        writer.WriteValue(_ListDevices.Item(i).refreshhisto)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("purge")
                        writer.WriteValue(_ListDevices.Item(i).purge)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("moyjour")
                        writer.WriteValue(_ListDevices.Item(i).moyjour)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("moyheure")
                        writer.WriteValue(_ListDevices.Item(i).moyheure)
                        writer.WriteEndAttribute()

                        '-- propriétés generique value --
                        If _ListDevices.Item(i).Type = "BAROMETRE" _
                        Or _ListDevices.Item(i).Type = "COMPTEUR" _
                        Or _ListDevices.Item(i).Type = "ENERGIEINSTANTANEE" _
                        Or _ListDevices.Item(i).Type = "ENERGIETOTALE" _
                        Or _ListDevices.Item(i).Type = "GENERIQUEVALUE" _
                        Or _ListDevices.Item(i).Type = "HUMIDITE" _
                        Or _ListDevices.Item(i).Type = "LAMPE" _
                        Or _ListDevices.Item(i).Type = "LAMPERGBW" _
                        Or _ListDevices.Item(i).Type = "PLUIECOURANT" _
                        Or _ListDevices.Item(i).Type = "PLUIETOTAL" _
                        Or _ListDevices.Item(i).Type = "TEMPERATURE" _
                        Or _ListDevices.Item(i).Type = "TEMPERATURECONSIGNE" _
                        Or _ListDevices.Item(i).Type = "VITESSEVENT" _
                        Or _ListDevices.Item(i).Type = "UV" _
                        Or _ListDevices.Item(i).Type = "VOLET" _
                        Then
                            writer.WriteStartAttribute("valuemin")
                            writer.WriteValue(_ListDevices.Item(i).valuemin)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("valuemax")
                            writer.WriteValue(_ListDevices.Item(i).valuemax)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("precision")
                            writer.WriteValue(_ListDevices.Item(i).precision)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("correction")
                            writer.WriteValue(_ListDevices.Item(i).correction)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("valuedef")
                            writer.WriteValue(_ListDevices.Item(i).valuedef)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("formatage")
                            writer.WriteValue(_ListDevices.Item(i).formatage)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("valuelast")
                            writer.WriteValue(_ListDevices.Item(i).valuelast)
                            writer.WriteEndAttribute()
                        End If
                        If _ListDevices.Item(i).Type = "LAMPERGBW" Then
                            writer.WriteStartAttribute("red")
                            writer.WriteValue(_ListDevices.Item(i).red)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("green")
                            writer.WriteValue(_ListDevices.Item(i).green)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("blue")
                            writer.WriteValue(_ListDevices.Item(i).blue)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("white")
                            writer.WriteValue(_ListDevices.Item(i).white)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("temperature")
                            writer.WriteValue(_ListDevices.Item(i).temperature)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("speed")
                            writer.WriteValue(_ListDevices.Item(i).speed)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("optionnal")
                            writer.WriteValue(_ListDevices.Item(i).optionnal)
                            writer.WriteEndAttribute()
                        End If

                        'écriture des variables
                        If _ListDevices.Item(i).Variables IsNot Nothing Then
                            For Each kvp As KeyValuePair(Of String, String) In _ListDevices.Item(i).Variables
                                writer.WriteStartElement("var")
                                writer.WriteStartAttribute("key")
                                writer.WriteValue(kvp.Key)
                                writer.WriteEndAttribute()
                                writer.WriteStartAttribute("value")
                                writer.WriteValue(kvp.Value)
                                writer.WriteEndAttribute()
                                writer.WriteEndElement()
                            Next kvp
                        End If



                        writer.WriteEndElement()
                    Catch ex As Exception
                        writer.WriteEndElement()
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: Composants: " & ex.ToString)
                    End Try
                Next
                writer.WriteEndElement()

                ''------------
                ''Sauvegarde des triggers
                ''------------
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des triggers")
                writer.WriteStartElement("triggers")
                For i As Integer = 0 To _ListTriggers.Count - 1
                    Try
                        writer.WriteStartElement("trigger")
                        writer.WriteStartAttribute("id")
                        writer.WriteValue(_ListTriggers.Item(i).ID)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("nom")
                        writer.WriteValue(_ListTriggers.Item(i).Nom)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("description")
                        writer.WriteValue(_ListTriggers.Item(i).Description)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("enable")
                        writer.WriteValue(_ListTriggers.Item(i).Enable)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("type")
                        If _ListTriggers.Item(i).Type = Trigger.TypeTrigger.TIMER Then
                            writer.WriteValue("0")
                        Else
                            writer.WriteValue("1")
                        End If
                        writer.WriteEndAttribute()
                        If _ListTriggers.Item(i).Type = Trigger.TypeTrigger.TIMER Then
                            writer.WriteStartAttribute("conditiontime")
                            writer.WriteValue(_ListTriggers.Item(i).ConditionTime)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("prochainedateheure")
                            writer.WriteValue(_ListTriggers.Item(i).Prochainedateheure)
                            writer.WriteEndAttribute()
                        End If
                        If _ListTriggers.Item(i).Type = Trigger.TypeTrigger.DEVICE Then
                            writer.WriteStartAttribute("conditiondeviceid")
                            writer.WriteValue(_ListTriggers.Item(i).ConditionDeviceId)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("conditiondeviceproperty")
                            writer.WriteValue(_ListTriggers.Item(i).ConditionDeviceProperty)
                            writer.WriteEndAttribute()
                        End If
                        writer.WriteStartElement("macros")
                        For k = 0 To _ListTriggers.Item(i).ListMacro.Count - 1
                            writer.WriteStartElement("macro")
                            writer.WriteStartAttribute("id")
                            writer.WriteValue(_ListTriggers.Item(i).ListMacro.Item(k))
                            writer.WriteEndAttribute()
                            writer.WriteEndElement()
                        Next
                        writer.WriteEndElement()
                        writer.WriteEndElement()
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: Triggers: " & ex.ToString)
                    End Try
                Next
                writer.WriteEndElement()

                ''------------
                ''Sauvegarde des macros
                ''------------
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des macros")
                writer.WriteStartElement("macros")
                For i As Integer = 0 To _ListMacros.Count - 1
                    Try
                        writer.WriteStartElement("macro")
                        writer.WriteStartAttribute("id")
                        writer.WriteValue(_ListMacros.Item(i).ID)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("nom")
                        writer.WriteValue(_ListMacros.Item(i).Nom)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("description")
                        writer.WriteValue(_ListMacros.Item(i).Description)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("enable")
                        writer.WriteValue(_ListMacros.Item(i).Enable)
                        writer.WriteEndAttribute()
                        WriteListAction(writer, _ListMacros.Item(i).ListActions)
                        writer.WriteEndElement()
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: Macros: " & ex.ToString)
                    End Try
                Next
                writer.WriteEndElement()

                ''------------
                ''Sauvegarde des nouveaux composants
                ''------------
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde des nouveaux composants")
                writer.WriteStartElement("newdevices")
                For i As Integer = 0 To _ListNewDevices.Count - 1
                    Try
                        writer.WriteStartElement("newdevice")
                        writer.WriteStartAttribute("id")
                        writer.WriteValue(_ListNewDevices.Item(i).ID)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("iddriver")
                        writer.WriteValue(_ListNewDevices.Item(i).IdDriver)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("adresse1")
                        writer.WriteValue(_ListNewDevices.Item(i).Adresse1)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("adresse2")
                        writer.WriteValue(_ListNewDevices.Item(i).Adresse2)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("name")
                        writer.WriteValue(_ListNewDevices.Item(i).Name)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("type")
                        writer.WriteValue(_ListNewDevices.Item(i).Type)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("ignore")
                        writer.WriteValue(_ListNewDevices.Item(i).Ignore)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("value")
                        writer.WriteValue(_ListNewDevices.Item(i).Value)
                        writer.WriteEndAttribute()
                        writer.WriteStartAttribute("datetetect")
                        writer.WriteValue(_ListNewDevices.Item(i).DateTetect)
                        writer.WriteEndAttribute()
                        writer.WriteEndElement()
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: Nouveau composant: " & ex.ToString)
                    End Try
                Next
                writer.WriteEndElement()
                ''FIN DES ELEMENTS------------

                writer.WriteEndDocument()
                writer.Close()

                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "SaveConfig", "Sauvegarde terminée")
                Return True
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfig", " Erreur de sauvegarde de la configuration: " & ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>Sauvegarde la config et BDD vers un dossier externe</summary>
        ''' <remarks></remarks>
        Private Function SaveConfigFolder() As Boolean
            Try
                Log(TypeLog.INFO, TypeSource.SERVEUR, "SaveConfigFolder", "Sauvegarde de la configuration vers le dossier " & _FolderSaveFolder)

                'test du chemin avec tempo au cas ou les NAS sont en veille

                Dim nbtentative = 10 'nombre de tentative pour vérifier si le répertoire de destination est ok
                Dim flagok As Boolean = False 'true si repertoire ok

                For i As Integer = 1 To nbtentative
                    If Not IO.Directory.Exists(_FolderSaveFolder) Then
                        Thread.Sleep(1000)
                    Else
                        flagok = True
                        Exit For
                    End If
                Next

                If flagok = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfigFolder", " Erreur de sauvegarde de la configuration vers le dossier externe " & _FolderSaveFolder & " car il n'est pas disponible")
                    Return False
                End If

                'If IO.Directory.Exists(_FolderSaveFolder) Then
                'homidom.xml
                If IO.File.Exists(_FolderSaveFolder & "\homidom.xml") = True Then IO.File.Delete(_FolderSaveFolder & "\homidom.xml")
                IO.File.Copy(_MonRepertoire & "\config\homidom.xml", _FolderSaveFolder & "\homidom.xml")

                'homidom.db
                If IO.File.Exists(_FolderSaveFolder & "\homidom.db") = True Then IO.File.Delete(_FolderSaveFolder & "\homidom.db")
                IO.File.Copy(_MonRepertoire & "\Bdd\homidom.db", _FolderSaveFolder & "\homidom.db")
                'Else
                'Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfigFolder", " Erreur : le dossier de sauvegarde n'existe pas : " & _FolderSaveFolder)
                'End If


            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfigFolder", " Erreur de sauvegarde de la configuration vers le dossier externe: " & ex.Message)
                Return False
            End Try

        End Function

        ''' <summary>
        ''' Ecris les actions dans le fichier de config
        ''' </summary>
        ''' <param name="writer"></param>
        ''' <param name="ListActions"></param>
        ''' <remarks></remarks>
        Private Sub WriteListAction(ByVal writer As XmlTextWriter, ByVal ListActions As ArrayList)
            Try
                For j As Integer = 0 To ListActions.Count - 1
                    writer.WriteStartElement("action")
                    writer.WriteStartAttribute("typeaction")
                    writer.WriteValue(ListActions.Item(j).TypeAction.ToString)
                    writer.WriteEndAttribute()
                    writer.WriteStartAttribute("timing")
                    writer.WriteValue(ListActions.Item(j).timing)
                    writer.WriteEndAttribute()
                    Select Case ListActions.Item(j).TypeAction
                        Case Action.TypeAction.ActionDevice
                            writer.WriteStartAttribute("iddevice")
                            writer.WriteValue(ListActions.Item(j).IdDevice)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("method")
                            writer.WriteValue(ListActions.Item(j).Method)
                            writer.WriteEndAttribute()
                            Dim a As String = ""
                            If ListActions.Item(j).parametres.count > 0 Then
                                a = a & ListActions.Item(j).parametres.item(0)
                                For k As Integer = 1 To ListActions.Item(j).parametres.count - 1
                                    a = a & "|" & ListActions.Item(j).parametres.item(k)
                                Next
                            End If
                            writer.WriteStartAttribute("parametres")
                            writer.WriteValue(a)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionDriver
                            writer.WriteStartAttribute("iddriver")
                            writer.WriteValue(ListActions.Item(j).Iddriver)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("method")
                            writer.WriteValue(ListActions.Item(j).Method)
                            writer.WriteEndAttribute()
                            Dim a As String = ""
                            If ListActions.Item(j).parametres.count > 0 Then
                                a = a & ListActions.Item(j).parametres.item(0)
                                For k As Integer = 1 To ListActions.Item(j).parametres.count - 1
                                    a = a & "|" & ListActions.Item(j).parametres.item(k)
                                Next
                            End If
                            writer.WriteStartAttribute("parametres")
                            writer.WriteValue(a)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionMacro
                            writer.WriteStartAttribute("idmacro")
                            writer.WriteValue(ListActions.Item(j).IdMacro)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionMail
                            writer.WriteStartAttribute("userid")
                            writer.WriteValue(ListActions.Item(j).UserId)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("sujet")
                            writer.WriteValue(ListActions.Item(j).Sujet)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("message")
                            writer.WriteValue(ListActions.Item(j).Message)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionSpeech
                            writer.WriteStartAttribute("message")
                            writer.WriteValue(ListActions.Item(j).Message)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionVB
                            writer.WriteStartAttribute("label")
                            writer.WriteValue(ListActions.Item(j).Label)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("script")
                            writer.WriteValue(ListActions.Item(j).Script)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionVar
                            writer.WriteStartAttribute("nom")
                            writer.WriteValue(ListActions.Item(j).Nom)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("value")
                            writer.WriteValue(ListActions.Item(j).Value)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionHttp
                            writer.WriteStartAttribute("commande")
                            writer.WriteValue(ListActions.Item(j).Commande)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionLogEvent
                            writer.WriteStartAttribute("message")
                            writer.WriteValue(ListActions.Item(j).Message)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("type")
                            Select Case UCase((ListActions.Item(j).Type.ToString))
                                Case "ERREUR"
                                    writer.WriteValue("1")
                                Case "WARNING"
                                    writer.WriteValue("2")
                                Case "INFORMATION"
                                    writer.WriteValue("3")
                                Case Else
                                    writer.WriteValue("1")
                            End Select
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("eventid")
                            writer.WriteValue(ListActions.Item(j).Eventid)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionLogEventHomidom
                            writer.WriteStartAttribute("message")
                            writer.WriteValue(ListActions.Item(j).Message)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("type")
                            Select Case UCase((ListActions.Item(j).Type.ToString))
                                Case "INFO"
                                    writer.WriteValue("1")
                                Case "ACTION"
                                    writer.WriteValue("2")
                                Case "MESSAGE"
                                    writer.WriteValue("3")
                                Case "VALEUR_CHANGE"
                                    writer.WriteValue("4")
                                Case "VALEUR_INCHANGE"
                                    writer.WriteValue("5")
                                Case "VALEUR_INCHANGE_PRECISION"
                                    writer.WriteValue("6")
                                Case "VALEUR_INCHANGE_LASTETAT"
                                    writer.WriteValue("7")
                                Case "ERREUR"
                                    writer.WriteValue("8")
                                Case "ERREUR_CRITIQUE"
                                    writer.WriteValue("9")
                                Case "DEBUG"
                                    writer.WriteValue("10")
                                Case Else
                                    writer.WriteValue("1")
                            End Select
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("fonction")
                            writer.WriteValue(ListActions.Item(j).Fonction)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionDOS
                            writer.WriteStartAttribute("fichier")
                            writer.WriteValue(ListActions.Item(j).Fichier)
                            writer.WriteEndAttribute()
                            writer.WriteStartAttribute("arguments")
                            writer.WriteValue(ListActions.Item(j).Arguments)
                            writer.WriteEndAttribute()
                        Case Action.TypeAction.ActionIf
                            writer.WriteStartElement("conditions")
                            For i2 As Integer = 0 To ListActions.Item(j).Conditions.count - 1
                                writer.WriteStartElement("condition")
                                writer.WriteStartAttribute("typecondition")
                                writer.WriteValue(ListActions.Item(j).Conditions.item(i2).Type.ToString)
                                writer.WriteEndAttribute()
                                If ListActions.Item(j).conditions.item(i2).Type = Action.TypeCondition.DateTime Then
                                    writer.WriteStartAttribute("datetime")
                                    writer.WriteValue(ListActions.Item(j).Conditions.item(i2).DateTime)
                                    writer.WriteEndAttribute()
                                End If
                                If ListActions.Item(j).conditions.item(i2).Type = Action.TypeCondition.Device Then
                                    writer.WriteStartAttribute("iddevice")
                                    writer.WriteValue(ListActions.Item(j).Conditions.item(i2).IdDevice)
                                    writer.WriteEndAttribute()
                                    writer.WriteStartAttribute("propertydevice")
                                    writer.WriteValue(ListActions.Item(j).Conditions.item(i2).propertydevice)
                                    writer.WriteEndAttribute()
                                    writer.WriteStartAttribute("value")
                                    writer.WriteValue(ListActions.Item(j).Conditions.item(i2).Value.ToString)
                                    writer.WriteEndAttribute()
                                End If
                                writer.WriteStartAttribute("condition")
                                writer.WriteValue(ListActions.Item(j).Conditions.item(i2).Condition.ToString)
                                writer.WriteEndAttribute()
                                writer.WriteStartAttribute("operateur")
                                writer.WriteValue(ListActions.Item(j).Conditions.item(i2).Operateur.ToString)
                                writer.WriteEndAttribute()
                                writer.WriteEndElement()
                            Next
                            writer.WriteEndElement()
                            writer.WriteStartElement("then")
                            If ListActions.Item(j).ListTrue IsNot Nothing Then WriteListAction(writer, ListActions.Item(j).ListTrue)
                            writer.WriteEndElement()
                            writer.WriteStartElement("else")
                            If ListActions.Item(j).ListFalse IsNot Nothing Then WriteListAction(writer, ListActions.Item(j).ListFalse)
                            writer.WriteEndElement()
                    End Select
                    writer.WriteEndElement()
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "WriteListAction", "Exception : " & ex.Message)
            End Try
        End Sub

        Public Function GetFrameworkVersionString() As String
            Try
                Dim CLRVersion As Version = System.Environment.Version

                If (CLRVersion.Major = 4 & CLRVersion.Minor = 0) Then
                    '4.0.30319.237   - 4.0
                    '4.0.30319.17020 - 4.5 (Microsoft .NET Framework 4.5 Developer Preview)
                    '4.0.30319.17379 - 4.5 (Microsoft .NET Framework 4.5 Consumer Preview)
                    '4.0.30319.17626 - 4.5 (Microsoft .NET Framework 4.5 RC)
                    '4.0.30319.17929 - 4.5 (Microsoft .NET Framework 4.5 RTM)
                    '4.0.30319.18408 - 4.5.1 (Microsoft .NET Framework 4.5.1 RTM - Windows Vista/7 - KB2858728)
                    '4.0.30319.34003 - 4.5.1 (Microsoft .NET Framework 4.5.1 RTM - Windows 8.1)
                    '4.0.30319.34209 - 4.5.2 (Microsoft .NET Framework 4.5.2 May 2014 Update)
                    '4.0.30319.42000 - 4.6 - In .NET Framework 4.6, the Environment.Version property returns the fixed version string 4.0.30319.42000

                    If (CLRVersion >= New Version(4, 0, 30319, 42000)) Then
                        Return "4.6"
                    ElseIf (CLRVersion >= New Version(4, 0, 30319, 34209)) Then
                        Return "4.5.2"
                    ElseIf (CLRVersion >= New Version(4, 0, 30319, 18408)) Then
                        Return "4.5.1"
                    ElseIf (CLRVersion >= New Version(4, 0, 30319, 17020)) Then
                        Return "4.5"
                    End If
                    Return "4.0"   '//4.0.30319.237
                ElseIf (CLRVersion.Major = 2 & CLRVersion.Minor = 0) Then
                    If (CLRVersion >= New Version(2, 0, 50727, 3521)) Then     '3.5.1
                        '2.0.50727.3521 - 3.5.1 in Windows 7 Beta 2
                        '2.0.50727.4016 - 3.5 SP1 in Windows Vista SP2 or Windows Server 2008 SP2
                        '2.0.50727.4918 - 3.5.1 in Windows 7 RC or Windows Server 2008 R2
                        Try
                            System.Reflection.Assembly.Load("System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
                            If (Environment.OSVersion.Platform = PlatformID.Win32NT & Environment.OSVersion.Version.Major >= 6 & Environment.OSVersion.Version.Minor >= 1) Then Return "3.5.1"
                            Return "3.5 SP1"
                        Catch ex As Exception
                        End Try
                        Return "2.0 SP2"
                    End If
                    If (CLRVersion >= New Version(2, 0, 50727, 3053)) Then Return "3.5 SP1"
                    If (CLRVersion >= New Version(2, 0, 50727, 1433)) Then
                        '2.0 SP1 or 3.0 SP1 or 3.5
                        Try
                            System.Reflection.Assembly.Load("System.Core, Version=3.5.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089")
                            Return "3.5"
                        Catch ex As Exception
                        End Try

                        Try
                            System.Reflection.Assembly.Load("WindowsBase, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
                            Return "3.0 SP1"
                        Catch ex As Exception
                        End Try

                        Return "2.0 SP1"
                    End If

                    If (CLRVersion = New Version(2, 0, 50727, 312)) Then Return "3.0" 'Vista RTM

                    '//2.0.50727.42 RTM - 2.0 or 3.0
                    Try
                        System.Reflection.Assembly.Load("WindowsBase, Version=3.0.0.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35")
                        Return "3.0"
                    Catch ex As Exception
                    End Try

                    Return "2.0"
                End If

                Return CLRVersion.Major.ToString(System.Globalization.CultureInfo.InvariantCulture) + "." + CLRVersion.Minor.ToString(System.Globalization.CultureInfo.InvariantCulture)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ERREUR GetFrameworkVersionString", "Erreur lors de la récupération da la version du framework .net " & ex.Message)
                Return ""
            End Try
        End Function
#End Region

#Region "Device"

        ''' <summary>Arretes les devices (Handlers)</summary>
        ''' <remarks></remarks>
        Public Sub Devices_Stop()
            Try
                'Cherche tous les devices chargés
                Log(TypeLog.INFO, TypeSource.SERVEUR, "Devices_Stop", "Arrêt des devices :")
                For Each _dev As Device.DeviceGenerique In _ListDevices
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "Devices_Stop", " - " & _dev.Name & " arrété")
                    'marche pas !!!!!
                    'Suivant chaque type de device
                    Select Case _dev.Type
                        Case "APPAREIL"
                            Dim o As Device.APPAREIL
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "AUDIO"
                            Dim o As Device.AUDIO
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "BAROMETRE"
                            Dim o As Device.BAROMETRE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "BATTERIE"
                            Dim o As Device.BATTERIE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "COMPTEUR"
                            Dim o As Device.COMPTEUR
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "CONTACT"
                            Dim o As Device.CONTACT
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "DETECTEUR"
                            Dim o As Device.DETECTEUR
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "DIRECTIONVENT"
                            Dim o As Device.DIRECTIONVENT
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "ENERGIEINSTANTANEE"
                            Dim o As Device.ENERGIEINSTANTANEE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "ENERGIETOTALE"
                            Dim o As Device.ENERGIETOTALE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "FREEBOX"
                            Dim o As Device.FREEBOX
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "GENERIQUEBOOLEEN"
                            Dim o As Device.GENERIQUEBOOLEEN
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "GENERIQUESTRING"
                            Dim o As Device.GENERIQUESTRING
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "GENERIQUEVALUE"
                            Dim o As Device.GENERIQUEVALUE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "HUMIDITE"
                            Dim o As Device.HUMIDITE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "LAMPE"
                            Dim o As Device.LAMPE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "LAMPERGBW"
                            Dim o As Device.LAMPERGBW
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "METEO"
                            Dim o As Device.METEO
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "MULTIMEDIA"
                            Dim o As Device.MULTIMEDIA
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "PLUIECOURANT"
                            Dim o As Device.PLUIECOURANT
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "PLUIETOTAL"
                            Dim o As Device.PLUIETOTAL
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "SWITCH"
                            Dim o As Device.SWITCH
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "TELECOMMANDE"
                            Dim o As Device.TELECOMMANDE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "TEMPERATURE"
                            Dim o As Device.TEMPERATURE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "TEMPERATURECONSIGNE"
                            Dim o As Device.TEMPERATURECONSIGNE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "UV"
                            Dim o As Device.UV
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "VITESSEVENT"
                            Dim o As Device.VITESSEVENT
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case "VOLET"
                            Dim o As Device.VOLET
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                        Case Else
                            Dim o As Device.GENERIQUEVALUE
                            o = _dev
                            RemoveHandler o.DeviceChanged, AddressOf DeviceChange
                            o = Nothing
                    End Select

                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Devices_Stop", " -> Erreur lors de l'arret des devices: " & ex.Message)
            End Try
        End Sub

        ''' <summary>Liste les type de devices par leur valeur d'Enum</summary>
        ''' <param name="Index"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnSringFromEnumDevice(ByVal Index As Integer) As String
            Try
                For Each value As Device.ListeDevices In [Enum].GetValues(GetType(Device.ListeDevices))
                    If value.GetHashCode = Index Then
                        Return value.ToString
                    End If
                Next
                Return ""
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnSringFromEnumDevice", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function
#End Region

#Region "Driver"
        ''' <summary>Retourne les propriétés d'un driver</summary>
        ''' <remarks></remarks>
        Public Function ReturnDriver(ByVal DriverId As String) As ArrayList
            Try
                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).ID = DriverId Then
                        Dim tabl As New ArrayList
                        tabl.Add(_ListDrivers.Item(i).nom)
                        tabl.Add(_ListDrivers.Item(i).enable)
                        tabl.Add(_ListDrivers.Item(i).description)
                        tabl.Add(_ListDrivers.Item(i).startauto)
                        tabl.Add(_ListDrivers.Item(i).protocol)
                        tabl.Add(_ListDrivers.Item(i).isconnect)
                        tabl.Add(_ListDrivers.Item(i).IP_TCP)
                        tabl.Add(_ListDrivers.Item(i).Port_TCP)
                        tabl.Add(_ListDrivers.Item(i).IP_UDP)
                        tabl.Add(_ListDrivers.Item(i).Port_UDP)
                        tabl.Add(_ListDrivers.Item(i).COM)
                        tabl.Add(_ListDrivers.Item(i).Refresh)
                        tabl.Add(_ListDrivers.Item(i).Modele)
                        tabl.Add(_ListDrivers.Item(i).Version)
                        tabl.Add(_ListDrivers.Item(i).Picture)
                        tabl.Add(_ListDrivers.Item(i).DeviceSupport)
                        tabl.Add(_ListDrivers.Item(i).Parametres)
                        'tabl.Add(_ListDrivers.Item(i).Labels)
                        tabl.Add(_ListDrivers.Item(i).OsPlatform)
                        Return tabl
                    End If
                Next
                Return Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDriver", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Ecrire ou lance propritété/Sub d'un driver</summary>
        ''' <remarks></remarks>
        Sub WriteDriver(ByVal DriverId As String, ByVal Command As String, ByVal Parametre As Object)
            Try
                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).ID = DriverId Then
                        Select Case UCase(Command)
                            Case "COM"
                                _ListDrivers.Item(i).Com = Parametre
                            Case "ENABLE"
                                _ListDrivers.Item(i).Enable = Parametre
                            Case "IP_TCP"
                                _ListDrivers.Item(i).IP_TCP = Parametre
                            Case "PORT_TCP"
                                _ListDrivers.Item(i).Port_TCP = Parametre
                            Case "IP_UDP"
                                _ListDrivers.Item(i).IP_UDP = Parametre
                            Case "PORT_UDP"
                                _ListDrivers.Item(i).Port_UDP = Parametre
                            Case "PICTURE"
                                _ListDrivers.Item(i).Picture = Parametre
                            Case "REFRESH"
                                _ListDrivers.Item(i).Refresh = Parametre
                            Case "MODELE"
                                _ListDrivers.Item(i).Modele = Parametre
                            Case "STARTAUTO"
                                _ListDrivers.Item(i).StartAuto = Parametre
                            Case "START"
                                _ListDrivers.Item(i).Start()
                            Case "STOP"
                                _ListDrivers.Item(i).Stop()
                            Case "RESTART"
                                _ListDrivers.Item(i).Restart()
                            Case "PARAMETRES"
                                For idx As Integer = 0 To Parametre.count - 1
                                    _ListDrivers.Item(i).Parametres.item(idx).valeur = Parametre(idx)
                                Next
                                'Case "LABELS"
                                '    For idx As Integer = 0 To Parametre.count - 1
                                '        _ListDrivers.Item(i).Labels.item(idx).tooltip = Parametre(idx)
                                '    Next
                            Case "DELETEDEVICE"
                                _ListDrivers.Item(i).DeleteDevice(Parametre)
                            Case "NEWDEVICE"
                                _ListDrivers.Item(i).NewDevice(Parametre)
                        End Select
                        Exit For

                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "WriteDriver", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Charge les drivers, donc toutes les dll dans le sous répertoire "drivers"</summary>
        ''' <remarks></remarks>
        Public Sub Drivers_Load()
            Try
                Dim tx As String = ""
                Dim dll As Reflection.Assembly
                Dim tp As Type
                Dim Chm As String = _MonRepertoire & "\Drivers\" 'Emplacement par défaut des drivers

                Dim strFileSize As String = ""
                Dim di As New IO.DirectoryInfo(Chm)
                Dim aryFi As IO.FileInfo() = di.GetFiles("Driver_*.dll")
                Dim fi As IO.FileInfo

                'Cherche tous les fichiers dll dans le répertoie plugin
                Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Load", "Chargement des DLL des drivers :")
                For Each fi In aryFi
                    Try
                        'chargement du plugin
                        tx = fi.FullName   'emplacement de la dll
                        'chargement de la dll
                        dll = Reflection.Assembly.LoadFrom(tx)
                        'Vérification de la présence de l'interface recherchée
                        For Each tp In dll.GetTypes
                            If tp.IsClass Then
                                If tp.GetInterface("IDriver", True) IsNot Nothing Then
                                    'création de la référence au plugin
                                    Dim i1 As IDriver
                                    i1 = DirectCast(dll.CreateInstance(tp.FullName), IDriver)
                                    i1 = CType(i1, IDriver)
                                    i1.Server = Me
                                    i1.IdSrv = _IdSrv

                                    If IO.File.Exists(_MonRepertoire & "\images\drivers\" & i1.Nom & ".png") Then
                                        i1.Picture = _MonRepertoire & "\images\drivers\" & i1.Nom & ".png"
                                    Else
                                        i1.Picture = _MonRepertoire & "\images\icones\Driver_128.png"
                                    End If

                                    'verification si un driver avec le même nom a déjà été chargé
                                    Dim driverdejacharge As Boolean = False
                                    For Each driverx As IDriver In _ListDrivers
                                        If driverx.Nom = i1.Nom Then
                                            driverdejacharge = True
                                            Exit For
                                        End If
                                    Next
                                    If driverdejacharge Then
                                        Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Load", " - " & i1.Nom & " non chargé car un autre driver avec le même nom a déjà été chargé (Version : " & i1.Version & " - " & i1.OsPlatform & " - " & fi.Name & ")")
                                    Else
                                        'Si le driver est prevu pour la plateforme de l'OS, on le charge
                                        If i1.OsPlatform.Contains(_OsPlatForm) Then
                                            Dim pt As New Driver(Me, _IdSrv, i1.ID)
                                            _ListDrivers.Add(i1)
                                            _ListImgDrivers.Add(pt)
                                            Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Load", " - " & i1.Nom & " chargé (Version : " & i1.Version & " - " & i1.OsPlatform & " - " & fi.Name & ")")
                                            pt = Nothing
                                        Else
                                            Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Load", " - " & i1.Nom & " non chargé, Platforme " & _OsPlatForm & " non géré (Version : " & i1.Version & " - " & i1.OsPlatform & " - " & fi.Name & ")")
                                        End If
                                    End If

                                    i1 = Nothing
                                End If
                            End If
                        Next
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Drivers_Load", " Erreur lors du chargement du driver: " & fi.Name & " Err: " & ex.ToString)
                    End Try
                Next

                dll = Nothing
                Chm = Nothing
                di = Nothing
                aryFi = Nothing
                fi = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Drivers_Load", " Erreur lors du chargement des drivers: " & ex.Message)
            End Try
        End Sub

        ''' <summary>Démarre tous les drivers dont la propriété StartAuto=True</summary>
        ''' <remarks></remarks>
        Public Sub Drivers_Start()
            Try
                'Cherche tous les drivers chargés
                Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Start", "Démarrage des drivers :")
                For Each driver In _ListDrivers
                    If driver.Enable And driver.StartAuto Then
                        Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Start", " - " & driver.Nom & " démarré")
                        driver.start()
                    Else
                        Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Start", " - " & driver.Nom & " non démarré car non Auto")
                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Drivers_Start", " -> Erreur lors du démarrage des drivers: " & ex.Message)
            End Try
        End Sub

        ''' <summary>Arretes les drivers démarrés</summary>
        ''' <remarks></remarks>
        Public Sub Drivers_Stop()
            Try
                'Cherche tous les drivers chargés
                Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Stop", "Arrêt des drivers :")
                For Each driver In _ListDrivers
                    If driver.Enable And driver.IsConnect Then
                        Log(TypeLog.INFO, TypeSource.SERVEUR, "Drivers_Stop", " - " & driver.Nom & " : ")
                        driver.stop()
                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Drivers_Stop", " -> Erreur lors de l'arret des drivers: " & ex.Message)
            End Try
        End Sub

#End Region

#Region "Cryptage"
        ''' <summary>
        ''' Crypter un string
        ''' </summary>
        ''' <param name="sIn"></param>
        ''' <param name="sKey"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function EncryptTripleDES(ByVal sIn As String, ByVal sKey As String) As String
            Try
                Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider
                Dim hashMD5 As New System.Security.Cryptography.MD5CryptoServiceProvider

                ' scramble the key
                sKey = ScrambleKey(sKey)
                ' Compute the MD5 hash.
                DES.Key = hashMD5.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(sKey))
                ' Set the cipher mode.
                DES.Mode = System.Security.Cryptography.CipherMode.ECB
                ' Create the encryptor.
                Dim DESEncrypt As System.Security.Cryptography.ICryptoTransform = DES.CreateEncryptor()
                ' Get a byte array of the string.
                Dim Buffer As Byte() = System.Text.ASCIIEncoding.ASCII.GetBytes(sIn)
                ' Transform and return the string.
                Return Convert.ToBase64String(DESEncrypt.TransformFinalBlock(Buffer, 0, Buffer.Length))
                DES = Nothing
                hashMD5 = Nothing
            Catch ex As Exception
                'Log(TypeLog.ERREUR, TypeSource.SERVEUR, "EncryptTripleDES", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

        ''' <summary>Décrypter un string</summary>
        ''' <param name="sOut"></param>
        ''' <param name="sKey"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DecryptTripleDES(ByVal sOut As String, ByVal sKey As String) As String
            Try
                Dim DES As New System.Security.Cryptography.TripleDESCryptoServiceProvider()
                Dim hashMD5 As New System.Security.Cryptography.MD5CryptoServiceProvider

                ' scramble the key
                sKey = ScrambleKey(sKey)
                ' Compute the MD5 hash.
                DES.Key = hashMD5.ComputeHash(System.Text.ASCIIEncoding.ASCII.GetBytes(sKey))
                ' Set the cipher mode.
                DES.Mode = System.Security.Cryptography.CipherMode.ECB
                ' Create the decryptor.
                Dim DESDecrypt As System.Security.Cryptography.ICryptoTransform = DES.CreateDecryptor()
                Dim Buffer As Byte() = Convert.FromBase64String(sOut)
                ' Transform and return the string.
                Return System.Text.ASCIIEncoding.ASCII.GetString(DESDecrypt.TransformFinalBlock(Buffer, 0, Buffer.Length))
            Catch ex As Exception
                'Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DecryptTripleDES", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

        Private Shared Function ScrambleKey(ByVal v_strKey As String) As String
            Try
                Dim sbKey As New System.Text.StringBuilder
                Dim intPtr As Integer

                For intPtr = 1 To v_strKey.Length
                    Dim intIn As Integer = v_strKey.Length - intPtr + 1
                    sbKey.Append(Mid(v_strKey, intIn, 1))
                Next
                Dim strKey As String = sbKey.ToString
                Return sbKey.ToString
            Catch ex As Exception
                'Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ScrambleKey", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function
#End Region

#Region "Log"
        'Dim _File As String = _MonRepertoire & "\logs\log.xml" 'Représente le fichier log: ex"C:\homidom\log\log.xml"
        Dim _FichierLog As String = _MonRepertoire & "\logs\log_" & DateAndTime.Now.ToString("yyyyMMdd") & ".txt"
        Dim _MaxFileSize As Long = 5120 'en Koctets

        ''' <summary>
        ''' Permet de connaître le chemin du fichier log
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property FichierLog() As String
            Get
                'Return _File
                Return _FichierLog
            End Get
        End Property

        ''' <summary>
        ''' Retourne/Fixe la Taille max du fichier log en Ko
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property MaxFileSize() As Long
            Get
                Return _MaxFileSize
            End Get
            Set(ByVal value As Long)
                _MaxFileSize = value
            End Set
        End Property

        ''' <summary>Indique le type du Log: si c'est une erreur, une info, un message...</summary>
        ''' <remarks></remarks>
        Public Enum TypeLog
            INFO = 1                    'divers
            ACTION = 2                  'action lancé par un driver/device/trigger
            MESSAGE = 3
            VALEUR_CHANGE = 4           'Valeur ayant changé
            VALEUR_INCHANGE = 5         'Valeur n'ayant pas changé
            VALEUR_INCHANGE_PRECISION = 6 'Valeur n'ayant pas changé pour cause de precision
            VALEUR_INCHANGE_LASTETAT = 7 'Valeur n'ayant pas changé pour cause de lastetat
            ERREUR = 8                   'erreur générale
            ERREUR_CRITIQUE = 9          'erreur critique demandant la fermeture du programme
            DEBUG = 10                   'visible uniquement si Homidom est en mode debug
        End Enum

        ''' <summary>Indique la source du log si c'est le serveur, un script, un device...</summary>
        ''' <remarks></remarks>
        Public Enum TypeSource
            SERVEUR = 1
            SCRIPT = 2
            TRIGGER = 3
            DEVICE = 4
            DRIVER = 5
            SOAP = 6
            CLIENT = 7
        End Enum

        ''' <summary>Indique le type d'event : info, warning ou error</summary>
        ''' <remarks></remarks>
        Public Enum TypeEventLog
            ERREUR = 1
            WARNING = 2
            INFORMATION = 3
        End Enum

        ''' <summary>
        ''' Insère dans une table le dernier log
        ''' </summary>
        ''' <param name="TypLog"></param>
        ''' <param name="Source"></param>
        ''' <param name="Fonction"></param>
        ''' <param name="Message"></param>
        ''' <remarks></remarks>
        Private Sub WriteLastLogs(ByVal TypLog As TypeLog, ByVal Source As TypeSource, ByVal Fonction As String, ByVal Message As String)
            Try
                For i = (_LastLogs.Count - 1) To 1 Step -1
                    _LastLogs(i) = _LastLogs(i - 1)
                Next
                _LastLogs(0) = Now & " - " & TypLog.ToString & " - " & Source.ToString & " - " & Fonction & " - " & Message
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "WriteLastLogs", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Insère dans une table le dernier log de type erreur ou alerte
        ''' </summary>
        ''' <param name="TypLog"></param>
        ''' <param name="Source"></param>
        ''' <param name="Fonction"></param>
        ''' <param name="Message"></param>
        ''' <remarks></remarks>
        Private Sub WriteLastLogsError(ByVal TypLog As TypeLog, ByVal Source As TypeSource, ByVal Fonction As String, ByVal Message As String)
            Try
                For i = (_LastLogsError.Count - 1) To 1 Step -1
                    _LastLogsError(i) = _LastLogsError(i - 1)
                Next
                _LastLogsError(0) = Now & " - " & TypLog.ToString & " - " & Source.ToString & " - " & Fonction & " - " & Message
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "WriteLastLogsError", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Ecrit un log dans le fichier log au format xml</summary>
        ''' <param name="TypLog"></param>
        ''' <param name="Source"></param>
        ''' <param name="Fonction"></param>
        ''' <param name="Message"></param>
        ''' <remarks></remarks>
        Public Sub Log(ByVal TypLog As TypeLog, ByVal Source As TypeSource, ByVal Fonction As String, ByVal Message As String) Implements IHoMIDom.Log
            Try

                If _TypeLogEnable(TypLog - 1) = True Then Exit Sub
                Dim _Message As String = DelRep(Message)
                Dim sourcestring As String = Source.ToString
                If Fonction = "" Then Fonction = "Divers"
                If sourcestring = "" Then sourcestring = "Divers"

                'on affiche dans la console
                If TypLog = TypeLog.ERREUR Or TypLog = TypeLog.ERREUR_CRITIQUE Then
                    Console.ForegroundColor = ConsoleColor.Red
                ElseIf TypLog = TypeLog.VALEUR_INCHANGE Or TypLog = TypeLog.VALEUR_INCHANGE_LASTETAT Or TypLog = TypeLog.VALEUR_INCHANGE_PRECISION Then
                    Console.ForegroundColor = ConsoleColor.Gray
                ElseIf TypLog = TypeLog.DEBUG Then
                    Console.ForegroundColor = ConsoleColor.DarkGray
                Else
                    Console.ForegroundColor = ConsoleColor.Black
                End If

                Console.WriteLine(Now & " " & TypLog.ToString & " " & Source.ToString & " " & Fonction & " " & _Message)


                WriteLastLogs(TypLog, Source, Fonction, _Message)

                RaiseEvent NewLog(TypLog, Source, Fonction, _Message)

                Select Case TypLog
                    Case TypeLog.ERREUR
                        WriteLastLogsError(TypLog, Source, Fonction, _Message)
                    Case TypeLog.ERREUR_CRITIQUE
                        WriteLastLogsError(TypLog, Source, Fonction, _Message)
                End Select

                'écriture dans un fichier texte
                _FichierLog = _MonRepertoire & "\logs\log_" & DateAndTime.Now.ToString("yyyyMMdd") & ".txt"
                Dim FreeF As Integer
                Try
                    FreeF = FreeFile()
                    SyncLock lock_logwrite
                        FileOpen(FreeF, FichierLog, OpenMode.Append)
                        Print(FreeF, Replace(Now & vbTab & TypLog.ToString & vbTab & sourcestring & vbTab & Fonction & vbTab & _Message, vbLf, vbCrLf) & vbCrLf)
                        'WriteLine(FreeF, Replace(Now & vbTab & TypLog.ToString & vbTab & sourcestring & vbTab & Fonction & vbTab & _Message, vbLf, vbCrLf) & vbCrLf)
                        FileClose(FreeF)
                    End SyncLock
                Catch ex As IOException
                    'wait(500)
                    Console.WriteLine(Now & " " & TypLog & " SERVER LOG ERROR IOException : " & ex.Message)
                Catch ex As Exception
                    'wait(500)
                    Console.WriteLine(Now & " " & TypLog & " SERVER LOG ERROR Exception : " & ex.Message)
                End Try
                FreeF = Nothing
            Catch ex As Exception
                Console.WriteLine("Erreur lors de l'écriture d'un log: " & ex.Message, MsgBoxStyle.Exclamation, "Erreur Serveur")
            End Try
        End Sub

        ''' <summary>
        ''' Supprime le répertoire du développeur
        ''' </summary>
        ''' <param name="Message"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function DelRep(ByVal Message As String) As String
            Try
                Dim i As Integer = 1
                Dim newmess As String = Message
                Dim start As Integer

                start = InStr(i, UCase(newmess), "C:\")
                If start = 0 Then
                    start = InStr(i, UCase(newmess), "D:\")
                End If

                Do While start > 0
                    Dim newstart As Integer = InStr(start, LCase(newmess), "\homidom\")
                    newmess = Mid(newmess, 1, start + 2) & "..." & Mid(newmess, newstart, newmess.Length - newstart + 1)
                    i = start + 5
                    start = InStr(i, UCase(newmess), "C:\")
                    If start = 0 Then
                        start = InStr(i, UCase(newmess), "D:\")
                    End If
                Loop

                Return newmess
            Catch ex As Exception
                Return Message
            End Try
        End Function


        ''' <summary>Ecrit un log dans les events Windows</summary>
        ''' <param name="message">text to write to event log</param>
        ''' <param name="type">1:erreur, 2:warning, 3:information</param>
        ''' <param name="eventid"></param>
        Public Sub LogEvent(ByVal message As String, ByVal type As TypeEventLog, Optional ByVal eventid As Integer = 0)
            Try
                Dim myEventLog = New EventLog()
                myEventLog.Source = "HoMIDoM"
                Select Case type
                    Case TypeEventLog.ERREUR : myEventLog.WriteEntry(message, EventLogEntryType.Error, eventid)
                    Case TypeEventLog.WARNING : myEventLog.WriteEntry(message, EventLogEntryType.Warning, eventid)
                    Case TypeEventLog.INFORMATION : myEventLog.WriteEntry(message, EventLogEntryType.Information, eventid)
                End Select
                'Diagnostics.EventLog.WriteEntry("HoMIDoM", message)
                myEventLog = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "LogEvent", "Exception : " & ex.Message)
            End Try
        End Sub

        'Private Function FileIsOpen(ByVal File As String) As Boolean
        '    Try
        '        'on tente d'ouvrir un stream sur le fichier, s'il est déjà utilisé, cela déclenche une erreur.
        '        Dim fs As IO.FileStream = My.Computer.FileSystem.GetFileInfo(File).Open(IO.FileMode.Open, _
        '        IO.FileAccess.Read)
        '        fs.Close()
        '        Return False
        '    Catch ex As Exception
        '        Return True
        '    End Try
        'End Function


        ' ''' <summary>Créer nouveau Fichier (donner chemin complet et nom) log</summary>
        ' ''' <param name="NewFichier"></param>
        ' ''' <remarks></remarks>
        'Public Sub CreateNewFileLog(ByVal NewFichier As String)
        '    Try
        '        Dim rw As XmlTextWriter = New XmlTextWriter(NewFichier, Nothing)
        '        rw.WriteStartDocument()
        '        rw.WriteStartElement("logs")
        '        rw.WriteStartElement("log")
        '        rw.WriteAttributeString("time", Now)
        '        rw.WriteAttributeString("type", 1)
        '        rw.WriteAttributeString("source", 1)
        '        rw.WriteAttributeString("fonction", "Log")
        '        rw.WriteAttributeString("message", "Création du nouveau fichier log")
        '        rw.WriteEndElement()
        '        rw.WriteEndElement()
        '        rw.WriteEndDocument()
        '        rw.Close()
        '    Catch ex As Exception
        '        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "CreateNewFileLog", "Erreur: " & ex.message)
        '    End Try
        'End Sub
#End Region

#Region "Declaration de la classe Server"

        ''' <summary>Déclaration de la class Server</summary>
        ''' <remarks></remarks>
        Public Sub New()
            Try
                Instance = Me
                'Check If Homidom Run in 32 or 64 bits
                If IntPtr.Size = 8 Then _OsPlatForm = "64" Else _OsPlatForm = "32"
                ManagerSequences.AddSequences(Sequence.TypeOfSequence.Server, Nothing, Nothing, Nothing)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Redémarre le service et charge la config
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Restart(ByVal IdSrv As String) Implements IHoMIDom.ReStart
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                [stop](_IdSrv)
                start()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Restart", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Démarrage du serveur</summary>
        ''' <remarks></remarks>
        Public Sub start() Implements IHoMIDom.Start
            Try
                'ajout d'un handler pour capturer les erreurs non catchés
                AddHandler AppDomain.CurrentDomain.UnhandledException, AddressOf Server_UnhandledExceptionEvent

                Dim retour As String

                'Charge les types de log
                For i As Integer = 0 To 9
                    _TypeLogEnable.Add(False)
                Next

                'Cree les sous répertoires s'ils nexistent pas
                If System.IO.Directory.Exists(_MonRepertoire & "\Logs") = False Then
                    System.IO.Directory.CreateDirectory(_MonRepertoire & "\Logs")
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Création du dossier logs")
                End If
                If System.IO.Directory.Exists(_MonRepertoire & "\Fichiers") = False Then
                    System.IO.Directory.CreateDirectory(_MonRepertoire & "\Fichiers")
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Création du dossier Fichiers")
                End If
                If System.IO.Directory.Exists(_MonRepertoire & "\Config") = False Then
                    System.IO.Directory.CreateDirectory(_MonRepertoire & "\config")
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Création du dossier config")
                End If
                If System.IO.Directory.Exists(_MonRepertoire & "\Images") = False Then
                    System.IO.Directory.CreateDirectory(_MonRepertoire & "\Images")
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Création du dossier images")
                End If
                If System.IO.Directory.Exists(_MonRepertoire & "\Images\Users") = False Then
                    System.IO.Directory.CreateDirectory(_MonRepertoire & "\Images\Users")
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Création du dossier images\Users")
                End If
                If System.IO.Directory.Exists(_MonRepertoire & "\Drivers") = False Then
                    System.IO.Directory.CreateDirectory(_MonRepertoire & "\Drivers")
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Création du dossier Drivers")
                End If
                If System.IO.Directory.Exists(_MonRepertoire & "\Templates") = False Then
                    System.IO.Directory.CreateDirectory(_MonRepertoire & "\Templates")
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Création du dossier templates")
                End If


                'Indique la version de la dll
                Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Version de la dll Homidom: " & GetServerVersion())
                If System.Environment.Is64BitOperatingSystem = True Then
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Version de l'OS: " & My.Computer.Info.OSFullName.ToString & " 64 Bits")
                Else
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Version de l'OS: " & My.Computer.Info.OSFullName.ToString & " 32 Bits")
                End If
                'Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Version du Framework: " & System.Runtime.InteropServices.RuntimeEnvironment.GetSystemVersion())
                Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Version du Framework: " & GetFrameworkVersionString() & " (" & System.Environment.Version.Major & "." & System.Environment.Version.Minor & "." & System.Environment.Version.Build & "." & System.Environment.Version.Revision & ")")
                Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Répertoire utilisé: " & My.Application.Info.DirectoryPath.ToString)

                'Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Adresse IP du serveur: " & System.Net.Dns.GetHostByName(My.Computer.Name).AddressList(0).ToString())
                '_IPSOAP = System.Net.Dns.GetHostByName(My.Computer.Name).AddressList(0).ToString()
                '_IPSOAP = System.Net.Dns.GetHostEntry(My.Computer.Name).AddressList(0).ToString()
                '_IPSOAP = Dns.GetHostEntry(Dns.GetHostName()).AddressList.Where(Function(a As IPAddress) Not a.IsIPv6LinkLocal AndAlso Not a.IsIPv6Multicast AndAlso Not a.IsIPv6SiteLocal).First().ToString()
                _IPSOAP = Dns.GetHostEntry(Dns.GetHostName()).AddressList.Where(Function(a As IPAddress) a.AddressFamily = Sockets.AddressFamily.InterNetwork).First().ToString()
                Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Adresse IP du serveur: " & _IPSOAP)
                Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Séparateur décimal CurrentCulture: '" & Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator & "'")

                'verify Local System separator
                Try
                    Const LOCALE_SDECIMAL = &HE ' Symbole décimal
                    'Const LOCALE_STHOUSAND = &HF ' Séparateur des milliers
                    Dim sBuffer As String
                    Dim nBufferLen As Long
                    nBufferLen = 255
                    sBuffer = New String(vbNullChar, nBufferLen)
                    nBufferLen = GetLocaleInfoEx(GetThreadLocale(), LOCALE_SDECIMAL, sBuffer, nBufferLen)
                    If nBufferLen > 0 Then
                        If Left$(sBuffer, nBufferLen - 1) <> Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator.ToString() Then
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "INFO", "Séparateur décimal Default User: '" & Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator & "' différent du séparateur courant, change la valeur de la clé HKU/default/International/sdecimal par le séparateur CurrentCulture")
                        End If
                    End If
                Catch ex As Exception

                End Try


                '---------- Creation table des threads ----------
                'table_TimerSecTickthread.Dispose()
                'Dim x As New DataColumn
                'x.ColumnName = "name"
                'table_TimerSecTickthread.Columns.Add(x)
                'x = New DataColumn
                'x.ColumnName = "comment"
                'table_TimerSecTickthread.Columns.Add(x)
                'x = New DataColumn
                'x.ColumnName = "datetime"
                'table_TimerSecTickthread.Columns.Add(x)
                'x = New DataColumn
                'x.ColumnName = "thread"
                'table_TimerSecTickthread.Columns.Add(x)

                'Si sauvegarde automatique
                If _CycleSave > 0 Then _NextTimeSave = Now.AddMinutes(_CycleSave)

                'Si sauvegarde automatique vers un folder
                If _CycleSaveFolder > 0 Then _NextTimeSaveFolder = Now.AddHours(_CycleSaveFolder)

                '----- Charge les drivers ----- 
                Drivers_Load()

                '----- Chargement de la config ----- 
                retour = LoadConfig(_MonRepertoire & "\Config\")
                Log(TypeLog.INFO, TypeSource.SERVEUR, "LoadConfig", retour)

                '----- Initialisation des sequences ----- 
                ManagerSequences.AddSequences(Sequence.TypeOfSequence.Device, Nothing, "", Nothing)
                ManagerSequences.AddSequences(Sequence.TypeOfSequence.Driver, Nothing, "", Nothing)
                ManagerSequences.AddSequences(Sequence.TypeOfSequence.Macro, Nothing, "", Nothing)
                ManagerSequences.AddSequences(Sequence.TypeOfSequence.Trigger, Nothing, "", Nothing)
                ManagerSequences.AddSequences(Sequence.TypeOfSequence.Zone, Nothing, "", Nothing)

                '----- Démarre les drivers ----- 
                Drivers_Start()

                '----- Calcul les heures de lever et coucher du soleil ----- 
                MAJ_HeuresSoleil()
                VerifIsJour()

                '----- Vérifie si c'est le weekend ----- 
                VerifIsWeekEnd()

                '----- Recupère le saint du jour ----- 
                MaJSaint()

                '----- Maj des triggers type CRON ----- 
                For i = 0 To _ListTriggers.Count - 1
                    'on vérifie si la condition est un cron
                    If _ListTriggers.Item(i).Type = Trigger.TypeTrigger.TIMER Then
                        _ListTriggers.Item(i).maj_cron() 'on calcule la date de prochain execution
                    End If
                Next

                '----- Démarre le Timer (triggers timers, actions programmées...) -----
                TimerSecond.Interval = 1000
                AddHandler TimerSecond.Elapsed, AddressOf TimerSecTick
                TimerSecond.Enabled = True

                '--- Démarre le serveur web si besoin
                If _EnableSrvWeb = True Then
                    _SrvWeb = New ServeurWeb(Me, _PortSrvWeb)
                    _SrvWeb.StartSrvWeb()
                    If _SrvWeb.IsStart Then Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Serveur Web démarré")
                End If

                '--- Démarre le serveur udp
                _SrvUDPIsStart = UDP.StartServerUDP(_PortSOAP - 1)
                If _SrvUDPIsStart Then Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Serveur UDP démarré sur le port:" & _PortSOAP - 1)

                '--- Scan les images
                Dim _thr As New Thread(AddressOf mImages.ScanImage)
                _thr.Priority = ThreadPriority.Lowest
                _thr.Start()

                'test biblio
                '
                ' Create a FileSystemWatcher object passing it the folder to watch.
                '
                'fsw = New FileSystemWatcher("C:\homidom")
                '
                ' Assign event procedures to the events to watch.
                '
                'AddHandler fsw.Created, AddressOf OnChanged
                'AddHandler fsw.Changed, AddressOf OnChanged
                'AddHandler fsw.Deleted, AddressOf OnChanged
                'AddHandler fsw.Renamed, AddressOf OnRenamed

                'With fsw
                '.EnableRaisingEvents = True
                '.IncludeSubdirectories = True
                '
                ' Specif the event to watch for.
                '
                '.WaitForChanged(WatcherChangeTypes.Created Or _
                '                WatcherChangeTypes.Changed Or _
                '                WatcherChangeTypes.Deleted Or _
                '                WatcherChangeTypes.Renamed)
                '
                ' Watch certain file types.
                '
                '.Filter = "*.txt"
                '
                ' Specify file change notifications.
                '
                '.NotifyFilter = (NotifyFilters.LastAccess Or _
                '                 NotifyFilters.LastWrite Or _
                '                 NotifyFilters.FileName Or _
                '                 NotifyFilters.DirectoryName)
                'End With

                'test log
                CleanLog()

                Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Serveur démarré")

            Catch ex As Exception
                Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "Start", "Exception : " & ex.Message)
            End Try

            'Manage SQLliste DB : updates
            Try
                'get sqlite moteur version
                Dim sqliteversion As String = ""
                Dim retour As String
                retour = sqlite_homidom.querysimple("SELECT SQLITE_VERSION()", sqliteversion)
                If Mid(retour, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur sqlite : " & retour)
                Else
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Version du moteur SQLlite: " & sqliteversion)
                End If

                'Get homidom BDD version
                sqliteversion = ""
                retour = sqlite_homidom.querysimple("PRAGMA user_version", sqliteversion)
                If Mid(retour, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur sqlite : " & retour)
                Else
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Version de la BDD HoMIDoM: " & sqliteversion)
                    If sqliteversion <= 1 Then
                        Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "BDD HoMIDoM Update vers la version 2")
                        retour = sqlite_homidom.nonquery("DROP TABLE 'config'")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        retour = sqlite_homidom.nonquery("CREATE TABLE 'config' ('id' INTEGER PRIMARY KEY  AUTOINCREMENT  NOT NULL  UNIQUE , 'parametre' TEXT, 'valeur' TEXT)")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        retour = sqlite_homidom.nonquery("INSERT INTO 'config' VALUES(1,'date_install','');")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        retour = sqlite_homidom.nonquery("INSERT INTO 'config' VALUES(2,'version_dll','');")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        retour = sqlite_homidom.nonquery("INSERT INTO 'config' VALUES(3,'date_register','');")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        retour = sqlite_homidom.nonquery("INSERT INTO 'config' VALUES(4,'date_maj','');")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        retour = sqlite_homidom.nonquery("INSERT INTO 'config' VALUES(5,'uid','');")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        retour = sqlite_homidom.nonquery("INSERT INTO 'config' VALUES(6,'cle_register','');")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        retour = sqlite_homidom.nonquery("PRAGMA user_version=2")
                        If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        sqliteversion = 2
                    End If
                    If sqliteversion = 2 Then
                        'Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "BDD HoMIDoM Update vers la version 3")

                        'retour = sqlite_homidom.nonquery("PRAGMA user_version=3")
                        'If Mid(retour, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLite versionning", "Erreur Update BDD : " & retour)
                        'sqliteversion = 3
                    End If
                End If
                sqliteversion = Nothing
                retour = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start SQLlite versionning", "Exception : " & ex.Message)
            End Try


            'Check Registration and Version in BDD
            Try
                Dim retoursql As String = ""
                Dim db_date_install As String = ""
                Dim db_date_register As String = ""
                Dim db_cle_register As String = ""
                Dim db_version As String = ""
                Dim db_uid As String = ""
                Dim uid As String = ""

                'on recupere les infos stockées en BDD
                Dim result As New DataTable
                retoursql = sqlite_homidom.query("SELECT valeur FROM config WHERE parametre='date_install'", result)
                If Mid(retoursql, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_date_install", "Erreur Requete sqlite : " & retoursql)
                Else
                    If result IsNot Nothing Then
                        db_date_install = result.Rows.Item(0).Item(0).ToString
                    Else
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "db_date_install date_install", "La base de donnée Homidom.db n'est pas à jour !")
                    End If
                End If
                retoursql = sqlite_homidom.query("SELECT valeur FROM config WHERE parametre='version_dll'", result)
                If Mid(retoursql, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_version", "Erreur Requete sqlite : " & retoursql)
                Else
                    If result IsNot Nothing Then
                        db_version = result.Rows.Item(0).Item(0).ToString
                    Else
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_version", "La base de donnée Homidom.db n'est pas à jour !")
                    End If
                End If
                retoursql = sqlite_homidom.query("SELECT valeur FROM config WHERE parametre='date_register'", result)
                If Mid(retoursql, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_date_register", "Erreur Requete sqlite : " & retoursql)
                Else
                    If result IsNot Nothing Then
                        db_date_register = result.Rows.Item(0).Item(0).ToString
                    Else
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_date_register", "La base de donnée Homidom.db n'est pas à jour !")
                    End If
                End If
                retoursql = sqlite_homidom.query("SELECT valeur FROM config WHERE parametre='uid'", result)
                If Mid(retoursql, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_uid", "Erreur Requete sqlite : " & retoursql)
                Else
                    If result IsNot Nothing Then
                        db_uid = result.Rows.Item(0).Item(0).ToString
                    Else
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_uid", "La base de donnée Homidom.db n'est pas à jour !")
                    End If
                End If
                retoursql = sqlite_homidom.query("SELECT valeur FROM config WHERE parametre='cle_register'", result)
                If Mid(retoursql, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_cle_register", "Erreur Requete sqlite : " & retoursql)
                Else
                    If result IsNot Nothing Then
                        db_cle_register = result.Rows.Item(0).Item(0).ToString
                    Else
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New db_cle_register", "La base de donnée Homidom.db n'est pas à jour !")
                    End If
                End If
                result = Nothing

                'on recup UID (motherboard serial an processorid)
                Dim uid1 As String = ""
                Dim uid2 As String = ""
                Dim searcher As ManagementObjectSearcher = New ManagementObjectSearcher("select * from Win32_Processor")
                For Each oReturn As ManagementObject In searcher.Get()
                    uid1 = oReturn("ProcessorId").ToString
                Next
                searcher.Dispose()
                searcher = Nothing
                searcher = New ManagementObjectSearcher("select SerialNumber from Win32_BaseBoard")
                For Each oReturn As ManagementObject In searcher.Get()
                    uid2 = oReturn("SerialNumber").ToString
                Next
                searcher.Dispose()
                searcher = Nothing
                'uid = uid1 + uid2
                'cryptage de la clé en MD5
                Dim md5Obj As New System.Security.Cryptography.MD5CryptoServiceProvider()
                Dim bytesToHash() As Byte = System.Text.Encoding.ASCII.GetBytes(uid1 + uid2)
                bytesToHash = md5Obj.ComputeHash(bytesToHash)
                Dim b As Byte
                For Each b In bytesToHash
                    uid += b.ToString("x2")
                Next

                'on recupere la version de Windows
                Dim osversion As String = My.Computer.Info.OSFullName.ToString() + "--" + Environment.OSVersion.Version.ToString() + "--" + Environment.OSVersion.ServicePack.ToString()
                If IntPtr.Size = 8 Then osversion += "--64bits" Else osversion += "--32bits"
                osversion = osversion.Replace(" ", "-")
                osversion = osversion.Replace(".", "-")

                'on recupere la resolution de l'ecran
                Dim resolution As String = Windows.Forms.Screen.PrimaryScreen.Bounds.Width.ToString() & "-" & Windows.Forms.Screen.PrimaryScreen.Bounds.Height

                'verif si premiere installation
                If String.IsNullOrEmpty(db_uid) = True Then
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Premiere Installation :Remerciements")
                    'premiere install : update uid/dateinstall/versiondll puis thanks puis register
                    'on maj la db
                    retoursql = sqlite_homidom.nonquery("UPDATE config Set valeur=@parameter0 WHERE parametre='uid'", uid)
                    If Mid(retoursql, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New Update UID", "Erreur Requete sqlite : " & retoursql)
                    retoursql = sqlite_homidom.nonquery("UPDATE config Set valeur=@parameter0 WHERE parametre='date_install'", Now.ToString("yyyMMdd HH:mm:ss"))
                    If Mid(retoursql, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New Update Date_Install", "Erreur Requete sqlite : " & retoursql)
                    retoursql = sqlite_homidom.nonquery("UPDATE config Set valeur=@parameter0 WHERE parametre='version_dll'", GetServerVersion())
                    If Mid(retoursql, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New Update version_dll", "Erreur Requete sqlite : " & retoursql)

                    'on ouvre la page web de remerciement
                    Try
                        'If (Environment.UserInteractive) Then Process.Start("http://www.homidom.com/premiereinstall_" & HtmlEncode(uid) & "_" & HtmlEncode(GetServerVersion().Replace(".", "-")) & "_" & HtmlEncode(osversion) & "_" & HtmlEncode(resolution) & ".html")
                        Dim request As HttpWebRequest = WebRequest.Create("http://www.homidom.com/premiereinstall_" & UrlEncode(uid) & "_" & UrlEncode(GetServerVersion().Replace(".", "-")) & "_" & UrlEncode(osversion) & "_" & UrlEncode(resolution) & ".html")
                        CType(request, HttpWebRequest).UserAgent = "Other"
                        Dim response As WebResponse = request.GetResponse()
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start OpenWebPage", "Exception : " & ex.Message)
                    End Try

                    'Gestion clé enregistrement

                Else
                    If db_version <> GetServerVersion() Then
                        'Homidom.dll a été mis à jour, on update la DB : version et date_maj puis remerciements
                        Log(TypeLog.INFO, TypeSource.SERVEUR, "INFO", "Mise à jour :Remerciements")

                        'on maj la db
                        retoursql = sqlite_homidom.nonquery("UPDATE config Set valeur=@parameter0 WHERE parametre='date_maj'", Now.ToString("yyyMMdd HH:mm:ss"))
                        If Mid(retoursql, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New Update date_maj", "Erreur Requete sqlite : " & retoursql)
                        retoursql = sqlite_homidom.nonquery("UPDATE config Set valeur=@parameter0 WHERE parametre='version_dll'", GetServerVersion())
                        If Mid(retoursql, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New Update version_dll", "Erreur Requete sqlite : " & retoursql)
                        retoursql = sqlite_homidom.nonquery("UPDATE config Set valeur=@parameter0 WHERE parametre='uid'", uid)
                        If Mid(retoursql, 1, 4) = "ERR:" Then Log(TypeLog.ERREUR, TypeSource.SERVEUR, "New Update uid", "Erreur Requete sqlite : " & retoursql)


                        Log(TypeLog.DEBUG, TypeSource.SERVEUR, "remerciement", " URL: http://www.homidom.com/miseajour_" & UrlEncode(uid) & "_" & UrlEncode(GetServerVersion().Replace(".", "-")) & "_" & UrlEncode(osversion) & "_" & UrlEncode(resolution) & ".html")


                        ' on ouvre la page web de remerciement
                        Try
                            'If (Environment.UserInteractive) Then Process.Start("http://www.homidom.com/miseajour_" & HtmlEncode(uid) & "_" & HtmlEncode(GetServerVersion().Replace(".", "-")) & "_" & HtmlEncode(osversion) & "_" & HtmlEncode(resolution) & ".html")
                            Dim request As HttpWebRequest = WebRequest.Create("http://www.homidom.com/miseajour_" & UrlEncode(uid) & "_" & UrlEncode(GetServerVersion().Replace(".", "-")) & "_" & UrlEncode(osversion) & "_" & UrlEncode(resolution) & ".html")
                            CType(request, HttpWebRequest).UserAgent = "Other"
                            Dim response As WebResponse = request.GetResponse()
                        Catch ex As Exception
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start OpenWebPage", "Exception : " & ex.Message)
                        End Try

                        'Dim startInfo As ProcessStartInfo = New ProcessStartInfo()
                        'startInfo.UseShellExecute = False
                        'startInfo.FileName = "http://www.homidom.com/miseajour_" & HtmlEncode(uid) & "_" & HtmlEncode(GetServerVersion().Replace(".", "-")) & "_" & HtmlEncode(osversion) & "_" & HtmlEncode(resolution) & ".html"
                        'Try
                        '    Process.Start(startInfo)
                        'Catch ex As Exception
                        '    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start OpenWebPage", "Exception : " & ex.Message)
                        'End Try

                        'Gestion clé enregistrement

                    End If

                End If


            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start Check", "Exception : " & ex.Message)
            End Try

            'Change l'etat du server
            Etat_server = True

            'changement de sequence
            ManagerSequences.AddSequences(Sequence.TypeOfSequence.ServerStart, Nothing, Nothing, Nothing)
            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Start SequenceServer", "N°: " & ManagerSequences.SequenceServer)
            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Start SequenceDriver", "N°: " & ManagerSequences.SequenceDriver)
            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Start SequenceDevice", "N°: " & ManagerSequences.SequenceDevice)
            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Start SequenceZone", "N°: " & ManagerSequences.SequenceZone)
            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Start SequenceTrigger", "N°: " & ManagerSequences.SequenceTrigger)
            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Start SequenceMacro", "N°: " & ManagerSequences.SequenceMacro)

            'passage du composant HOMI_SERVER à True
            Try
                Dim _devstart As Object = ReturnRealDeviceById("startsrv01")
                If _devstart IsNot Nothing Then _devstart.value = True
                _devstart = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Start Homi_server", "Exception : " & ex.Message)
            End Try


        End Sub

        ''' <summary>Arrêt du serveur</summary>
        ''' <remarks></remarks>
        Public Sub [stop](ByVal IdSrv As String) Implements IHoMIDom.Stop
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                'passage du composant HOMI_SERVER à False
                Dim _devstart As Object = ReturnRealDeviceById("startsrv01")
                If _devstart IsNot Nothing Then _devstart.value = False
                _devstart = Nothing

                ManagerSequences.AddSequences(Sequence.TypeOfSequence.ServerShutDown, Nothing, Nothing, Nothing)

                'on change l'etat du server pour ne plus lancer de traitement
                Etat_server = False

                If _CycleSave > 0 Or _SaveRealTime Then SaveConfig(_MonRepertoire & "\config\homidom.xml")

                TimerSecond.Enabled = False
                'RemoveHandler TimerSecond.Elapsed, AddressOf TimerSecTick
                TimerSecond.Dispose()

                '----- Arrete les devices ----- 
                Devices_Stop()
                _ListDevices.Clear()

                '----- Arrete les drivers ----- 
                Drivers_Stop()
                _ListDrivers.Clear()

                '----- Vide les variables -----
                _ListGroups.Clear()
                _ListImgDrivers.Clear()
                _ListMacros.Clear()
                _ListTriggers.Clear()
                _ListUsers.Clear()
                _ListZones.Clear()

                '----- Arrete les connexions Sqlite -----
                'retour = sqlite_homidom.disconnect("homidom")
                'If Mid(retour, 1, 4) = "ERR:" Then
                '    Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "Stop", "Erreur lors de la deconnexion de la BDD Homidom : " & retour)
                'End If
                'retour = sqlite_medias.disconnect("medias")
                'If Mid(retour, 1, 4) = "ERR:" Then
                '    Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "Stop", "Erreur lors de la deconnexion de la BDD Medias : " & retour)
                'End If

                'suprression de l'handler pour recup les erreurs non catchés
                RemoveHandler AppDomain.CurrentDomain.UnhandledException, AddressOf Server_UnhandledExceptionEvent

                Log(TypeLog.INFO, TypeSource.SERVEUR, "Stop", "Serveur Arrêté")
            Catch ex As Exception
                Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "Stop", "Exception : " & ex.Message)
            End Try
        End Sub

        Protected Overrides Sub Finalize()
            Try
                'Mettre le Code pour l'arret
                ' [stop]()
                MyBase.Finalize()
            Catch ex As Exception
                Log(TypeLog.ERREUR_CRITIQUE, TypeSource.SERVEUR, "Finalize", "Exception : " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Macro"
        Private Sub Execute(ByVal Id As String)
            Try
                Dim mymacro As New Macro

                mymacro = ReturnMacroById(_IdSrv, Id)
                If mymacro IsNot Nothing Then
                    Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Macro:Action", "Lancement de la macro " & mymacro.Nom)

                    Dim a As String = Api.GenerateGUID

                    For i = 0 To mymacro.ListActions.Count - 1
                        Dim _Action As New ThreadAction(Me, mymacro.ListActions.Item(i), mymacro.Nom, a)
                        Dim x As New Thread(AddressOf _Action.Execute)

                        If mymacro.ListActions.Item(i).typeaction = Action.TypeAction.ActionStop Then
                            x.Name &= ".STOP"
                        Else
                            x.Name = a
                        End If
                        x.Start()
                        _ListThread.Add(x)
                    Next

                    mymacro = Nothing
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Macro:Action", ex.Message)
            End Try
        End Sub
#End Region

#Region "GuideTV"
        Public MyChaine As New List(Of sChaine)
        Public MyProgramme As New List(Of sProgramme)
        Dim MyXML As HoMIDom.XML
        Dim timestart As String

        Public Structure sProgramme
            Dim DateStart As String
            Dim DateEnd As String
            Dim TimeStart As String
            Dim TimeEnd As String
            Dim IDChannel As String
            Dim Titre As String
            Dim SousTitre As String
            Dim Description As String
            Dim Duree As Integer
            Dim Categorie1 As String
            Dim Categorie2 As String
            Dim Annee As Integer
            Dim Credits As String
        End Structure

        Public Structure sChaine
            Dim Nom As String
            Dim ID As String
            Dim Ico As String
            Dim Enable As Integer
            Dim Categorie As String
            Dim Numero As Integer
        End Structure

        Public Function ConvertTextToHTML(ByVal Text As String) As String
            Try
                Text = Replace(Text, "'", "&#191;")
                Text = Replace(Text, "À", "&#192;")
                Text = Replace(Text, "Á", "&#193;")
                Text = Replace(Text, "Â", "&#194;")
                Text = Replace(Text, "Ã", "&#195;")
                Text = Replace(Text, "Ä", "&#196;")
                Text = Replace(Text, "Å", "&#197;")
                Text = Replace(Text, "Æ", "&#198;")
                Text = Replace(Text, "à", "&#224;")
                Text = Replace(Text, "á", "&#225;")
                Text = Replace(Text, "â", "&#226;")
                Text = Replace(Text, "ã", "&#227;")
                Text = Replace(Text, "ä", "&#228;")
                Text = Replace(Text, "å", "&#229;")
                Text = Replace(Text, "æ", "&#230;")
                Text = Replace(Text, "Ç", "&#199;")
                Text = Replace(Text, "ç", "&#231;")
                Text = Replace(Text, "Ð", "&#208;")
                Text = Replace(Text, "ð", "&#240;")
                Text = Replace(Text, "È", "&#200;")
                Text = Replace(Text, "É", "&#201;")
                Text = Replace(Text, "Ê", "&#202;")
                Text = Replace(Text, "Ë", "&#203;")
                Text = Replace(Text, "è", "&#232;")
                Text = Replace(Text, "é", "&#233;")
                Text = Replace(Text, "ê", "&#234;")
                Text = Replace(Text, "ë", "&#235;")
                Text = Replace(Text, "Ì", "&#204;")
                Text = Replace(Text, "Í", "&#205;")
                Text = Replace(Text, "Î", "&#206;")
                Text = Replace(Text, "Ï", "&#207;")
                Text = Replace(Text, "ì", "&#236;")
                Text = Replace(Text, "í", "&#237;")
                Text = Replace(Text, "î", "&#238;")
                Text = Replace(Text, "ï", "&#239;")
                Text = Replace(Text, "Ñ", "&#209;")
                Text = Replace(Text, "ñ", "&#241;")
                Text = Replace(Text, "Ò", "&#210;")
                Text = Replace(Text, "Ó", "&#211;")
                Text = Replace(Text, "Ô", "&#212;")
                Text = Replace(Text, "Õ", "&#213;")
                Text = Replace(Text, "Ö", "&#214;")
                Text = Replace(Text, "Ø", "&#216;")
                Text = Replace(Text, "Œ", "&#140;")
                Text = Replace(Text, "ò", "&#242;")
                Text = Replace(Text, "ó", "&#243;")
                Text = Replace(Text, "ô", "&#244;")
                Text = Replace(Text, "õ", "&#245;")
                Text = Replace(Text, "ö", "&#246;")
                Text = Replace(Text, "ø", "&#248;")
                Text = Replace(Text, "œ", "&#156;")
                Text = Replace(Text, "Š", "&#138;")
                Text = Replace(Text, "š", "&#154;")
                Text = Replace(Text, "Ù", "&#217;")
                Text = Replace(Text, "Ú", "&#218;")
                Text = Replace(Text, "Û", "&#219;")
                Text = Replace(Text, "Ü", "&#220;")
                Text = Replace(Text, "ù", "&#249;")
                Text = Replace(Text, "ú", "&#250;")
                Text = Replace(Text, "û", "&#251;")
                Text = Replace(Text, "ü", "&#252;")
                Text = Replace(Text, "Ý", "&#221;")
                Text = Replace(Text, "Ÿ", "&#159;")
                Text = Replace(Text, "ý", "&#253;")
                Text = Replace(Text, "ÿ", "&#255;")
                Text = Replace(Text, "Ž", "&#142;")
                Text = Replace(Text, "ž", "&#158;")
                Text = Replace(Text, "¢", "&#162;")
                Text = Replace(Text, "£", "&#163;")
                Text = Replace(Text, "¥", "&#165;")
                Text = Replace(Text, "™", "&#153;")
                Text = Replace(Text, "©", "&#169;")
                Text = Replace(Text, "®", "&#174;")
                Text = Replace(Text, "‰", "&#137;")
                Text = Replace(Text, "ª", "&#170;")
                Text = Replace(Text, "º", "&#186;")
                Text = Replace(Text, "¹", "&#185;")
                Text = Replace(Text, "²", "&#178;")
                Text = Replace(Text, "³", "&#179;")
                Text = Replace(Text, "¼", "&#188;")
                Text = Replace(Text, "½", "&#189;")
                Text = Replace(Text, "¾", "&#190;")
                Text = Replace(Text, "÷", "&#247;")
                Text = Replace(Text, "×", "&#215;")
                Text = Replace(Text, ">", "&#155;")
                Text = Replace(Text, "<", "&#139;")
                Text = Replace(Text, "±", "&#177;")
                Text = Replace(Text, "&", "")
                Text = Replace(Text, "‚", "&#130;")
                Text = Replace(Text, "ƒ", "&#131;")
                Text = Replace(Text, "„", "&#132;")
                Text = Replace(Text, "…", "&#133;")
                Text = Replace(Text, "†", "&#134;")
                Text = Replace(Text, "‡", "&#135;")
                Text = Replace(Text, "ˆ", "&#136;")
                Text = Replace(Text, "‘", "&#145;")
                Text = Replace(Text, "’", "&#146;")
                'Text=Replace(text,"“","&#147;")
                'Text=Replace(text,"”","&#148;")
                Text = Replace(Text, "•", "&#149;")
                Text = Replace(Text, "–", "&#150;")
                Text = Replace(Text, "—", "&#151;")
                Text = Replace(Text, "˜", "&#152;")
                Text = Replace(Text, "¿", "&#191;")
                Text = Replace(Text, "¡", "&#161;")
                Text = Replace(Text, "¤", "&#164;")
                Text = Replace(Text, "¦", "&#166;")
                Text = Replace(Text, "§", "&#167;")
                Text = Replace(Text, "¨", "&#168;")
                Text = Replace(Text, "«", "&#171;")
                Text = Replace(Text, "»", "&#187;")
                Text = Replace(Text, "¬", "&#172;")
                Text = Replace(Text, "¯", "&#175;")
                Text = Replace(Text, "´", "&#180;")
                Text = Replace(Text, "µ", "&#181;")
                Text = Replace(Text, "¶", "&#182;")
                Text = Replace(Text, "·", "&#183;")
                Text = Replace(Text, "¸", "&#184;")
                Text = Replace(Text, "þ", "&#222;")
                Text = Replace(Text, "ß", "&#223;")
                ConvertTextToHTML = Text
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ConvertTextToHTML", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

        Public Function ConvertHtmlToText(ByVal Text As String) As String
            Try
                Text = Replace(Text, "#191;", "'")
                Text = Replace(Text, "#192;", "À")
                Text = Replace(Text, "#193;", "Á")
                Text = Replace(Text, "#194;", "Â")
                Text = Replace(Text, "#195;", "Ã")
                Text = Replace(Text, "#196;", "Ä")
                Text = Replace(Text, "#197;", "Å")
                Text = Replace(Text, "#198;", "Æ")
                Text = Replace(Text, "#224;", "à")
                Text = Replace(Text, "#225;", "á")
                Text = Replace(Text, "#226;", "â")
                Text = Replace(Text, "#227;", "ã")
                Text = Replace(Text, "#228;", "ä")
                Text = Replace(Text, "#229;", "å")
                Text = Replace(Text, "#230;", "æ")
                Text = Replace(Text, "#199;", "Ç")
                Text = Replace(Text, "#231;", "ç")
                Text = Replace(Text, "#208;", "Ð")
                Text = Replace(Text, "#240;", "ð")
                Text = Replace(Text, "#200;", "È")
                Text = Replace(Text, "#201;", "É")
                Text = Replace(Text, "#202;", "Ê")
                Text = Replace(Text, "#203;", "Ë")
                Text = Replace(Text, "#232;", "è")
                Text = Replace(Text, "#233;", "é")
                Text = Replace(Text, "#234;", "ê")
                Text = Replace(Text, "#235;", "ë")
                Text = Replace(Text, "#204;", "Ì")
                Text = Replace(Text, "#205;", "Í")
                Text = Replace(Text, "#206;", "Î")
                Text = Replace(Text, "#207;", "Ï")
                Text = Replace(Text, "#236;", "ì")
                Text = Replace(Text, "#237;", "í")
                Text = Replace(Text, "#238;", "î")
                Text = Replace(Text, "#239;", "ï")
                Text = Replace(Text, "#209;", "Ñ")
                Text = Replace(Text, "#241;", "ñ")
                Text = Replace(Text, "#210;", "Ò")
                Text = Replace(Text, "#211;", "Ó")
                Text = Replace(Text, "#212;", "Ô")
                Text = Replace(Text, "#213;", "Õ")
                Text = Replace(Text, "#214;", "Ö")
                Text = Replace(Text, "#216;", "Ø")
                Text = Replace(Text, "#140;", "Œ")
                Text = Replace(Text, "#242;", "ò")
                Text = Replace(Text, "#243;", "ó")
                Text = Replace(Text, "#244;", "ô")
                Text = Replace(Text, "#245;", "õ")
                Text = Replace(Text, "#246;", "ö")
                Text = Replace(Text, "#248;", "ø")
                Text = Replace(Text, "#156;", "œ")
                Text = Replace(Text, "#138;", "Š")
                Text = Replace(Text, "#154;", "š")
                Text = Replace(Text, "#217;", "Ù")
                Text = Replace(Text, "#218;", "Ú")
                Text = Replace(Text, "#219;", "Û")
                Text = Replace(Text, "#220;", "Ü")
                Text = Replace(Text, "#249;", "ù")
                Text = Replace(Text, "#250;", "ú")
                Text = Replace(Text, "#251;", "û")
                Text = Replace(Text, "#252;", "ü")
                Text = Replace(Text, "#221;", "Ý")
                Text = Replace(Text, "#159;", "Ÿ")
                Text = Replace(Text, "#253;", "ý")
                Text = Replace(Text, "#255;", "ÿ")
                Text = Replace(Text, "#142;", "Ž")
                Text = Replace(Text, "#158;", "ž")
                Text = Replace(Text, "#162;", "¢")
                Text = Replace(Text, "#163;", "£")
                Text = Replace(Text, "#165;", "¥")
                Text = Replace(Text, "#153;", "™")
                Text = Replace(Text, "#169;", "©")
                Text = Replace(Text, "#174;", "®")
                Text = Replace(Text, "#137;", "‰")
                Text = Replace(Text, "#170;", "ª")
                Text = Replace(Text, "#186;", "º")
                Text = Replace(Text, "#185;", "¹")
                Text = Replace(Text, "#178;", "²")
                Text = Replace(Text, "#179;", "³")
                Text = Replace(Text, "#188;", "¼")
                Text = Replace(Text, "#189;", "½")
                Text = Replace(Text, "#190;", "¾")
                Text = Replace(Text, "#247;", "÷")
                Text = Replace(Text, "#215;", "×")
                Text = Replace(Text, "#155;", ">")
                Text = Replace(Text, "#139;", "<")
                Text = Replace(Text, "#177;", "±")
                Text = Replace(Text, "#130;", "‚")
                Text = Replace(Text, "#131;", "ƒ")
                Text = Replace(Text, "#132;", "„")
                Text = Replace(Text, "#133;", "…")
                Text = Replace(Text, "#134;", "†")
                Text = Replace(Text, "#135;", "‡")
                Text = Replace(Text, "#136;", "ˆ")
                Text = Replace(Text, "#145;", "‘")
                Text = Replace(Text, "#146;", "’")
                Text = Replace(Text, "#149;", "•")
                Text = Replace(Text, "#150;", "–")
                Text = Replace(Text, "#151;", "—")
                Text = Replace(Text, "#152;", "˜")
                Text = Replace(Text, "#191;", "¿")
                Text = Replace(Text, "#161;", "¡")
                Text = Replace(Text, "#164;", "¤")
                Text = Replace(Text, "#166;", "¦")
                Text = Replace(Text, "#167;", "§")
                Text = Replace(Text, "#168;", "¨")
                Text = Replace(Text, "#171;", "«")
                Text = Replace(Text, "#187;", "»")
                Text = Replace(Text, "#172;", "¬")
                Text = Replace(Text, "#175;", "¯")
                Text = Replace(Text, "#180;", "´")
                Text = Replace(Text, "#181;", "µ")
                Text = Replace(Text, "#182;", "¶")
                Text = Replace(Text, "#183;", "·")
                Text = Replace(Text, "#184;", "¸")
                Text = Replace(Text, "#222;", "þ")
                Text = Replace(Text, "#223;", "ß")
                ConvertHtmlToText = Text
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ConvertHtmlToText", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

#Region "Gestion des chaines"
        'Permet de lister les chaines dans la base de données
        Public Sub ChaineFromXMLToDB()
            Try
                MyXML = New HoMIDom.XML(_MonRepertoire & "\data\complet.xml")
                Dim liste As XmlNodeList = MyXML.SelectNodes("/tv/channel")
                Dim i As Integer
                Dim a As String
                Dim b As String
                Dim SQLconnect As New SQLiteConnection()
                Dim SQLcommand As SQLiteCommand
                SQLconnect.ConnectionString = "Data Source= " & _MonRepertoire & "\bdd\guidetv.db;"
                SQLconnect.Open()
                SQLcommand = SQLconnect.CreateCommand
                SQLcommand.CommandText = "DELETE FROM chaineTV where id<>''"
                SQLcommand.ExecuteNonQuery()
                SQLcommand = SQLconnect.CreateCommand
                Dim SQLreader As SQLiteDataReader

                'liste toute les chaines
                For i = 0 To liste.Count - 1
                    a = liste(i).Attributes.Item(0).Value
                    'Affiche l'ID et le nom
                    SQLcommand.CommandText = "SELECT * FROM chaineTV where ID='" & a & "'"
                    SQLreader = SQLcommand.ExecuteReader()
                    If SQLreader.HasRows = False Then
                        b = liste(i).ChildNodes.Item(0).ChildNodes.Item(0).Value
                        b = ConvertTextToHTML(b)
                        SQLreader.Close()
                        SQLcommand = SQLconnect.CreateCommand
                        SQLcommand.CommandText = "INSERT INTO chaineTV (id, nom,ico,enable,numero,categorie) VALUES ('" & a & "', '" & b & "','?','0','0','99')"
                        SQLcommand.ExecuteNonQuery()
                    End If
                Next
                SQLcommand.Dispose()
                SQLconnect.Close()
                ChargeChaineFromDB()
                MyXML = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ChaineFromXMLToDB", "Exception : " & ex.Message)
            End Try
        End Sub

        'Charge les chaines depuis la base de données en mémoire
        Public Sub ChargeChaineFromDB()
            Try
                Dim SQLconnect As New SQLiteConnection()
                Dim SQLcommand As SQLiteCommand
                SQLconnect.ConnectionString = "Data Source= " & _MonRepertoire & "\bdd\guidetv.db;"
                SQLconnect.Open()
                SQLcommand = SQLconnect.CreateCommand
                Dim SQLreader As SQLiteDataReader

                SQLcommand.CommandText = "SELECT * FROM chaineTV"
                SQLreader = SQLcommand.ExecuteReader()
                If SQLreader.HasRows = True Then
                    While SQLreader.Read()
                        Dim vChaine As sChaine = New sChaine
                        vChaine.Nom = SQLreader(1)
                        vChaine.ID = SQLreader(2)
                        vChaine.Ico = SQLreader(3)
                        vChaine.Enable = SQLreader(4)
                        vChaine.Numero = SQLreader(5)
                        vChaine.Categorie = SQLreader(6)
                        MyChaine.Add(vChaine)
                    End While
                Else
                    Console.WriteLine(Now & ": aucune chaine à charger depuis la DB!")
                End If
                SQLcommand.Dispose()
                SQLconnect.Close()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ChargeChaineFromDB", "Exception : " & ex.Message)
            End Try
        End Sub

#End Region

#Region "Compression"
        Public Function decompression(ByVal cheminSource As String, ByVal cheminDestination As String) As Boolean
            Try
                Dim process As Process = New Process()
                process.StartInfo.FileName = "C:\Program Files\7-zip\7z.exe"
                process.StartInfo.Arguments = " e " + cheminSource & " -aoa -o" & cheminDestination
                process.Start()
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "decompression", "Exception : " & ex.Message)
            End Try
        End Function
#End Region

#End Region

#Region "Bibliotheques"
        'Public Sub SearchTag()
        '    Try
        '        For i As Integer = 0 To _ListRepertoireAudio.Count - 1
        '            '  Dim x As New Thread(AddressOf FileTagRepload(_ListRepertoireAudio.Item(i).Repertoire))
        '        Next
        '    Catch ex As Exception
        '        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SearchTag", "Exception : " & ex.Message)
        '    End Try
        'End Sub

        'Public Class ThreadSearchTag
        '    Dim _Repertoire As String
        '    Dim _mylist As List(Of Audio.FilePlayList)

        '    Sub New(ByVal Repertoire As String)
        '        _Repertoire = Repertoire
        '    End Sub


        '    '''' <summary>Fonction de chargement des tags des fichiers audio des repertoires contenus dans la liste active </summary>
        '    '''' <remarks>Recupere les fichiers Audios selon les extensions actives</remarks>
        '    Sub FileTagRepload()

        '        ' Créér une reference du dossier
        '        Dim di As New DirectoryInfo(_Repertoire)

        '        ' Pour chacune des extensions
        '        For cpt2 = 0 To _ListExtensionAudio.Count - 1

        '            Dim _extension As String = _ListExtensionAudio.Item(cpt2).Extension
        '            Dim _extensionenable As Boolean = _ListExtensionAudio.Item(cpt2).Enable

        '            ' Recupere la liste des fichiers du repertoire si l'extension est active
        '            If _extensionenable Then ' Extension active 
        '                ' Recuperation des fichiers du repertoire
        '                Dim fiArr As FileInfo() = di.GetFiles("*" & _extension, SearchOption.TopDirectoryOnly)

        '                ' Boucle sur tous les fichiers du repertoire
        '                For i = 0 To fiArr.Length - 1
        '                    Dim ii = i
        '                    Dim Resultat = (From FileAudio In _mylist Where FileAudio.SourceWpath = fiArr(ii).FullName Select FileAudio).Count
        '                    If Resultat = 0 Then
        '                        Dim X As TagLib.File
        '                        ' Recupere les tags du fichier Audio 
        '                        X = TagLib.File.Create(fiArr(i).FullName)
        '                        Dim a As New Audio.FilePlayList(X.Tag.Title, X.Tag.FirstPerformer, X.Tag.Album, X.Tag.Year, X.Tag.Comment, X.Tag.FirstGenre,
        '                                                  System.Convert.ToString(X.Properties.Duration.Minutes) & ":" & System.Convert.ToString(Format(X.Properties.Duration.Seconds, "00")),
        '                                                  fiArr(i).Name, fiArr(i).FullName, X.Tag.Track)

        '                        _mylist.Add(a)

        '                        a = Nothing
        '                        X = Nothing
        '                    End If
        '                Next
        '            End If
        '        Next
        '    End Sub

        'End Class


        'Public Sub OnChanged(ByVal source As Object, ByVal e As FileSystemEventArgs)

        'End Sub

        'Public Sub OnRenamed(ByVal source As Object, ByVal e As RenamedEventArgs)

        'End Sub

#End Region

#Region "Voix"
        ''' <summary>
        ''' Gets a list of all the Speech Synthesis voices installed on the computer.
        ''' </summary>
        ''' <returns>A string array of all the voices installed</returns>
        ''' <remarks>Make sure you have the System.Speech namespace reference added to your project</remarks>
        Function ReturnAllSpeechSynthesisVoices() As List(Of String)
            Try
                Dim oSpeech As New System.Speech.Synthesis.SpeechSynthesizer()
                Dim installedVoices As System.Collections.ObjectModel. _
                                        ReadOnlyCollection(Of System.Speech.Synthesis.InstalledVoice) _
                                        = oSpeech.GetInstalledVoices
                Dim names As New List(Of String)

                If installedVoices IsNot Nothing Then
                    For i As Integer = 0 To installedVoices.Count - 1
                        names.Add(installedVoices(i).VoiceInfo.Name)
                    Next
                End If

                Return names
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnAllSpeechSynthesisVoices", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retourne la première voix disponible (utilisée pour la mettre par défaut si aucune voix n'est définie)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function GetFirstVoice() As String
            Try

                Dim strAllVoices As New List(Of String)
                strAllVoices = ReturnAllSpeechSynthesisVoices()
                Dim retour As String = ""

                If strAllVoices.Count > 0 Then
                    retour = strAllVoices(0)
                End If

                Return retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetFirstVoice", "Erreur : " & ex.Message)
                Return Nothing
            End Try
        End Function

#End Region

#Region "Energie"
        'ajouter la gestion save/load/param _Puissancemini
        Public Property GererEnergie As Boolean
            Get
                Return _GererEnergie
            End Get
            Set(ByVal value As Boolean)
                If value And _GererEnergie = False Then
                    PuissanceTotaleActuel = _PuissanceMini
                End If
                _GererEnergie = value
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "GererEnergie", "Activation=" & value)
            End Set
        End Property

        Public Property PuissanceMini As Integer
            Get
                Return _PuissanceMini
            End Get
            Set(ByVal value As Integer)
                If _PuissanceMini <> value Then
                    _PuissanceMini = value
                    PuissanceTotaleActuel = _PuissanceMini
                End If
            End Set
        End Property

        Public Property PuissanceTotaleActuel As Integer
            Get
                Return _PuissanceTotaleActuel
            End Get
            Set(ByVal value As Integer)
                _PuissanceTotaleActuel = value

                If _PuissanceTotaleActuel <= 0 Then _PuissanceTotaleActuel = 0

                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).id = "energietotale01" Then
                        If _ListDevices.Item(i).Value <> _PuissanceTotaleActuel Then _ListDevices.Item(i).Value = _PuissanceTotaleActuel
                        Exit For
                    End If
                Next
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "PuissanceTotaleActuelle", "Puissance Totale Actuelle=" & _PuissanceTotaleActuel)
            End Set
        End Property
#End Region

#Region "Serveur Web"

#End Region

#End Region

#Region "Interface Client via IHOMIDOM"
        '********************************************************************
        'Fonctions/Sub/Propriétés partagées en service soap pour les clients
        '********************************************************************

        '**** PROPRIETES ***************************

        Public Property Devices() As ArrayList
            Get
                Return _ListDevices
            End Get
            Set(ByVal value As ArrayList)
                _ListDevices = value
            End Set
        End Property

        Public Property Images As List(Of ImageFile)
            Get
                Return _listImages
            End Get
            Set(value As List(Of ImageFile))
                _listImages = value
            End Set
        End Property

        '*** FONCTIONS ******************************************
#Region "Serveur"
        ''' <summary>
        ''' Retourne: Sauvegader les backups et sauvegardes suivant différents fichiers
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetSaveDiffBackup() As Boolean Implements IHoMIDom.GetSaveDiffBackup
            Try
                Return _SaveDiffBackup
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSaveDiffBackup", "Erreur : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Définit: Sauvegader les backups et sauvegardes suivant différents fichiers
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SetSaveDiffBackup(Value As Boolean) Implements IHoMIDom.SetSaveDiffBackup
            Try
                _SaveDiffBackup = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetSaveDiffBackup", "Erreur : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne le répertoire courant du serveur
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetRepertoireOfServer() As String Implements IHoMIDom.GetRepertoireOfServer
            Try
                Return _MonRepertoire
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetRepertoireOfServer", "Erreur : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Indique au serveur si on sauvegarde en temps réel
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub SetSaveRealTime(ByVal Value As Boolean) Implements IHoMIDom.SetSaveRealTime
            _SaveRealTime = Value
        End Sub

        ''' <summary>
        ''' Demande au serveur si on sauvegarde en temps réel
        ''' </summary>
        ''' <remarks></remarks>
        Public Function GetSaveRealTime() As Boolean Implements IHoMIDom.GetSaveRealTime
            Return _SaveRealTime
        End Function

        ''' <summary>
        ''' Retourne le code du pays du CultureInfo
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetCodePays() As Integer Implements IHoMIDom.GetCodePays
            Try
                Return CodePays
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetCodePays", "Erreur : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Fixe le code du pays suivant CultureInfo
        ''' </summary>
        ''' <param name="Code"></param>
        ''' <remarks></remarks>
        Public Sub SetCodePays(ByVal Code As Integer) Implements IHoMIDom.SetCodePays
            Try
                CodePays = Code
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetCodePays", "Erreur : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne la liste des ports com dispo sur le serveur
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPortComDispo() As String() Implements IHoMIDom.GetPortComDispo
            Try
                Dim portNames As String() = SerialPort.GetPortNames()
                Array.Sort(portNames)
                Return portNames
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetPortComDispo", "Erreur : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Vérifie si un élément existe dans une zone, une macro, un trigger... avant de le supprimer
        ''' </summary>
        ''' <param name="IdSrv"></param>
        ''' <param name="Id"></param>
        ''' <returns>Retourne une erreur commencant par ERREUR ou la liste des noms des macros, zones...</returns>
        ''' <remarks></remarks>
        Public Function CanDelete(ByVal IdSrv As String, ByVal Id As String) As List(Of String) Implements IHoMIDom.CanDelete
            Dim retour As New List(Of String)
            Try
                Dim thr As New ThreadDelete(Me, IdSrv, Id, retour)
                Dim x As New Thread(AddressOf thr.Traite)
                x.Start()

                Do While retour.Count = 0

                Loop
                Do While retour(retour.Count - 1) <> "0"

                Loop
                Return retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "CanDelete", "Erreur : " & ex.Message)
                retour.Add("ERREUR lors de l'exécution de la fonction: " & ex.Message)
                Return retour
            End Try
        End Function

        Private Class ThreadDelete
            Dim _retour As List(Of String)
            Dim _server As Server
            Dim _id As String
            Dim _Idsrv As String

            Public Sub New(ByVal Server As Server, ByVal IdSrv As String, ByVal Id As String, ByVal Retour As List(Of String))
                _server = Server
                _retour = Retour
                _id = Id
                _Idsrv = IdSrv
            End Sub

            Public Sub Traite()
                Try
                    _server._CanDelete(_Idsrv, _id, _retour)
                Catch ex As Exception

                End Try
            End Sub
        End Class

        Private Sub _CanDelete(ByVal IdSrv As String, ByVal Id As String, ByVal retour As List(Of String))
            Try
                If VerifIdSrv(IdSrv) = False Then
                    retour.Add("ERREUR: L'Id du serveur est erronée")
                    retour.Add("0")
                    Exit Sub
                End If
                If String.IsNullOrEmpty(Id) = True Then
                    retour.Add("ERREUR: L'Id est vide")
                    retour.Add("0")
                    Exit Sub
                End If

                'va vérifier toutes les zones
                For i As Integer = 0 To _ListZones.Count - 1
                    For j As Integer = 0 To _ListZones.Item(i).ListElement.Count - 1
                        If _ListZones.Item(i).ListElement.Item(j).ElementID = Id Then
                            AddLabel(retour, "- La zone: " & _ListZones.Item(i).Name)
                            Exit For
                        End If
                    Next
                Next

                'va vérifier tous les triggers
                For i As Integer = 0 To _ListTriggers.Count - 1
                    If _ListTriggers.Item(i).ConditionDeviceId = Id Then AddLabel(retour, "- Le trigger: " & _ListTriggers.Item(i).Nom)
                Next

                'va vérifier toutes les macros
                For i As Integer = 0 To _ListMacros.Count - 1
                    VerifIdInAction(_ListMacros.Item(i).ListActions, Id, _ListMacros.Item(i).Nom, retour)
                Next

                retour.Add("0")
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "_CanDelete", "Erreur : " & ex.Message)
                retour.Add("ERREUR lors de l'exécution de la fonction: " & ex.Message)
                retour.Add("0")
            End Try
        End Sub

        Private Sub AddLabel(ByVal List As List(Of String), ByVal Message As String)
            Try
                For i As Integer = 0 To List.Count - 1
                    If List(i) = Message Then Exit Sub
                Next
                List.Add(Message)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "AddLabel", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne la liste des voix installées sur le serveur
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAllVoice() As List(Of String) Implements IHoMIDom.GetAllVoice
            Try
                Dim strAllVoices As New List(Of String)
                strAllVoices = ReturnAllSpeechSynthesisVoices()
                'Dim list As New List(Of String)
                'For Each Str As String In strAllVoices
                '    list.Add(Str)
                'Next
                Return strAllVoices
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllVoice", "Erreur : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retourne la voix par défaut
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDefautVoice() As String Implements IHoMIDom.GetDefautVoice
            Return _Voice
        End Function

        ''' <summary>
        ''' Définit la voix par défaut
        ''' </summary>
        ''' <param name="Voice"></param>
        ''' <remarks></remarks>
        Public Sub SetDefautVoice(ByVal Voice As String) Implements IHoMIDom.SetDefautVoice
            _Voice = Voice
        End Sub

        Private Sub VerifIdInAction(ByVal Actions As ArrayList, ByVal Id As String, ByVal NameMacro As String, ByVal Retour As List(Of String))
            Try
                For j As Integer = 0 To Actions.Count - 1
                    Select Case Actions.Item(j).TypeAction
                        Case Action.TypeAction.ActionDevice
                            If Actions.Item(j).IdDevice = Id Then AddLabel(Retour, "- La Macro: " & NameMacro)
                        Case Action.TypeAction.ActionIf
                            Dim x As Action.ActionIf = Actions.Item(j)
                            For k As Integer = 0 To x.Conditions.Count - 1
                                If x.Conditions.Item(k).IdDevice = Id Then AddLabel(Retour, "- La Macro: " & NameMacro)
                            Next
                            VerifIdInAction(x.ListTrue, Id, NameMacro, Retour)
                            VerifIdInAction(x.ListFalse, Id, NameMacro, Retour)
                        Case Action.TypeAction.ActionMail
                            If Actions.Item(j).UserId = Id Then AddLabel(Retour, "- La Macro: " & NameMacro)
                        Case Action.TypeAction.ActionMacro
                            If Actions.Item(j).IdMacro = Id Then AddLabel(Retour, "- La Macro: " & NameMacro)
                    End Select
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "VerifIdInAction", "Erreur : " & ex.Message)
                AddLabel(Retour, "ERREUR lors de l'exécution de la fonction: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne le paramètre de sauvegarde
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTimeSave(ByVal IdSrv As String) As Integer Implements IHoMIDom.GetTimeSave
            If VerifIdSrv(IdSrv) = False Then
                Return "-1"
            End If
            Return _CycleSave
        End Function

        ''' <summary>
        ''' Fixe le paramètre de sauvegarde
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SetTimeSave(ByVal IdSrv As String, ByVal Value As Integer) As String Implements IHoMIDom.SetTimeSave
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                If IsNumeric(Value) = False Or Value < 0 Then
                    Return "ERR: la valeur doit être numérique, positive et non nulle"
                Else
                    _CycleSave = Value
                    _NextTimeSave = Now.AddMinutes(Value)
                    Return 0
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetTimeSave", "Exception : " & ex.Message)
                Return "ERR: exception"
            End Try
        End Function

        ''' <summary>
        ''' Fixe le chemin de sauvegarde de la config vers un folder
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub SetFolderSaveFolder(ByVal Value As String) Implements IHoMIDom.SetFolderSaveFolder
            _FolderSaveFolder = Value
        End Sub

        ''' <summary>
        ''' Retourne le chemin de sauvegarde de la config vers un folder
        ''' </summary>
        ''' <remarks></remarks>
        Public Function GetFolderSaveFolder() As String Implements IHoMIDom.GetFolderSaveFolder
            Return _FolderSaveFolder
        End Function


        ''' <summary>
        ''' Retourne le cycle de sauvegarde de la config vers un folder
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTimeSaveFolder(ByVal IdSrv As String) As Integer Implements IHoMIDom.GetTimeSaveFolder
            If VerifIdSrv(IdSrv) = False Then
                Return "-1"
            End If
            Return _CycleSaveFolder
        End Function

        ''' <summary>
        ''' Fixe le cycle de sauvegarde de la config vers un folder
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SetTimeSaveFolder(ByVal IdSrv As String, ByVal Value As Integer) As String Implements IHoMIDom.SetTimeSaveFolder
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                If IsNumeric(Value) = False Or Value < 0 Then
                    Return "ERR: la valeur doit être numérique, positive et non nulle"
                Else
                    _CycleSaveFolder = Value
                    _NextTimeSaveFolder = Now.AddHours(Value)
                    Return 0
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetTimeSave", "Exception : " & ex.Message)
                Return "ERR: exception"
            End Try
        End Function


        ''' <summary>
        ''' Retourne l'Id du serveur pour SOAP
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetIdServer(ByVal IdSrv As String) As String Implements IHoMIDom.GetIdServer
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return "99"
                End If
                Return _IdSrv
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetIdServer", "Exception : " & ex.Message)
                Return "99"
            End Try
        End Function

        ''' <summary>
        ''' Fixe l'Id du serveur pour SOAP
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SetIdServer(ByVal IdSrv As String, ByVal Value As String) As String Implements IHoMIDom.SetIdServer
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                If String.IsNullOrEmpty(Value) = True Then
                    Return "ERR: l'Id ne peut être null"
                Else
                    _IdSrv = Value
                    Return 0
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetIdServer", "Exception : " & ex.Message)
                Return "ERR: Exception"
            End Try
        End Function

        ''' <summary>Retourne la date et heure du dernier démarrage du serveur</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLastStartTime() As Date Implements IHoMIDom.GetLastStartTime
            Return _DateTimeLastStart
        End Function

        ''' <summary>Retourne la version du serveur</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetServerVersion() As String Implements IHoMIDom.GetServerVersion
            Try
                Return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.ToString
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetServerVersion", "Exception : " & ex.Message)
                Return "ERR: Exception"
            End Try
        End Function

        ''' <summary>Retourne la version du framework .net du serveur</summary>
        ''' <returns>version du framework .net du serveur</returns>
        ''' <remarks>exemple: 4.5.2 (4.0.30319.18502)</remarks>
        Public Function GetFrameworkNetServerVersion() As String Implements IHoMIDom.GetFrameworkNetServerVersion
            Try
                Return GetFrameworkVersionString() & "(" & System.Environment.Version.Major & "." & System.Environment.Version.Minor & "." & System.Environment.Version.Build & "." & System.Environment.Version.Revision & ")"
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetFrameworkNetServerVersion", "Exception : " & ex.Message)
                Return "ERR: Exception"
            End Try
        End Function

        ''' <summary>Retourne l'heure du serveur</summary>
        ''' <returns>String : heure du serveur</returns>
        Public Function GetTime() As String Implements IHoMIDom.GetTime
            Try
                Return Now.ToLongTimeString
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetTime", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Permet d'envoyer un message d'un client vers le server pour les renvoyer vers tous les clients</summary>
        ''' <param name="Message"></param>
        ''' <remarks></remarks>
        Public Sub MessageToServer(ByVal Message As String) Implements IHoMIDom.MessageToServer
            Try
                'traiter le message
                ManagerSequences.AddSequences(Sequence.TypeOfSequence.Message, Nothing, Nothing, Message)
                Log(TypeLog.MESSAGE, TypeSource.SERVEUR, "MessageToServer", "Message From client : " & Message)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "MessageToServer", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Convert a file on a byte array.</summary>
        ''' <param name="file"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetByteFromImage(ByVal file As String) As Byte() Implements IHoMIDom.GetByteFromImage
            Try
                If String.IsNullOrEmpty(file) = True Then Return Nothing
                If IO.File.Exists(file) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetByteFromImage", "le fichier n'existe pas: " & file)
                    Return Nothing
                End If

                Dim array As Byte() = Nothing
                Using fs As New FileStream(file, FileMode.Open, FileAccess.Read)
                    Dim reader As New BinaryReader(fs)
                    If reader IsNot Nothing Then
                        array = reader.ReadBytes(CInt(fs.Length))
                        reader.Close()
                        reader = Nothing
                    End If
                End Using

                Return array

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetByteFromImage", ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retourne la liste de tous les fichiers image (png ou jpg) présents sur le serveur
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetListOfImage() As List(Of ImageFile) Implements IHoMIDom.GetListOfImage
            Try
                Return _listImages
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetListOfImage", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Upload d'un fichier image vers le serveur
        ''' </summary>
        ''' <param name="IdSrv"></param>
        ''' <param name="byteData"></param>
        ''' <param name="Namefile"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UploadFile(ByVal IdSrv As String, ByVal byteData As Byte(), ByVal Namefile As String) As String Implements IHoMIDom.UploadFile
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim _fichier As String = _MonRepertoire & "\images\myimages\" & Namefile

                If System.IO.Directory.Exists(_MonRepertoire & "\Images\myimages") = False Then
                    System.IO.Directory.CreateDirectory(_MonRepertoire & "\Images\myimages")
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "UploadFile", "Création du dossier myimages")
                End If

                If IO.File.Exists(_fichier) Then
                    Return "Le fichier " & Namefile & " existe déjà veuillez le renommer"
                End If

                Dim oFileStream As System.IO.FileStream
                oFileStream = New System.IO.FileStream(_fichier, System.IO.FileMode.Create)
                oFileStream.Write(byteData, 0, byteData.Length)
                oFileStream.Close()

                Dim _file As New ImageFile
                _file.FileName = Namefile
                _file.Path = _fichier
                _listImages.Add(_file)

                Return 0
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "Upload", "Exception : " & ex.Message)
                Return ex.Message
            End Try
        End Function

        ''' <summary>Retourne la version BDD sqlite</summary>
        ''' <returns>String : version BDD</returns>
        Public Function GetSqliteBddVersion() As String Implements IHoMIDom.GetSqliteBddVersion
            Try
                Dim sqliteversion As String = ""
                Dim retour As String = ""

                retour = sqlite_homidom.querysimple("PRAGMA user_version", sqliteversion)
                If Mid(retour, 1, 4) <> "ERR:" Then
                    Return sqliteversion
                Else
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSqliteBddVersion", retour)
                    Return "ERROR"
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSqliteBddVersion", "Exception : " & ex.Message)
                Return "ERROR"
            End Try
        End Function

        ''' <summary>Retourne la version du moteur sqlite</summary>
        ''' <returns>String : version Moteur sqlite</returns>
        Public Function GetSqliteVersion() As String Implements IHoMIDom.GetSqliteVersion
            Try
                Dim sqliteversion As String = ""
                Dim retour As String
                retour = sqlite_homidom.querysimple("SELECT SQLITE_VERSION()", sqliteversion)
                If Mid(retour, 1, 4) <> "ERR:" Then
                    Return sqliteversion
                Else
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSqliteBddVersion", retour)
                    Return retour
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSqliteVersion", "Exception : " & ex.Message)
                Return "ERR:Exception GetSqliteVersion"
            End Try
        End Function

        ''' <summary>
        ''' Retourne la devise du serveur
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDevise() As String Implements IHoMIDom.GetDevise
            Try
                Return _Devise
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetDevise", "Erreur: " & ex.Message)
                Return ""
            End Try
        End Function

        ''' <summary>
        ''' Fixe la devise du serveur
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub SetDevise(ByVal Value As String) Implements IHoMIDom.SetDevise
            Try
                _Devise = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetDevise", "Erreur: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne si le serveur Web est Enable
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetEnableServeurWeb() As Boolean Implements IHoMIDom.GetEnableServeurWeb
            Try
                Return _EnableSrvWeb
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetEnableServeurWeb", "Erreur: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Fixe Enable du serveur web
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub SetEnableServeurWeb(ByVal Value As Boolean) Implements IHoMIDom.SetEnableServeurWeb
            Try
                _EnableSrvWeb = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetEnableServeurWeb", "Erreur: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne le port du serveur Web 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPortServeurWeb() As Integer Implements IHoMIDom.GetPortServeurWeb
            Try
                Return _PortSrvWeb
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetPortServeurWeb", "Erreur: " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Fixe le port du serveur web
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub SetPortServeurWeb(ByVal Value As Integer) Implements IHoMIDom.SetPortServeurWeb
            Try
                _PortSrvWeb = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetPortServeurWeb", "Erreur: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Redémarre ou démarre le serveur web
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub RestartServeurWeb() Implements IHoMIDom.RestartServeurWeb
            Try
                If _EnableSrvWeb = True Then
                    _SrvWeb = New ServeurWeb(Me, _PortSrvWeb)
                    _SrvWeb.StartSrvWeb()
                    If _SrvWeb.IsStart Then Log(TypeLog.INFO, TypeSource.SERVEUR, "Start", "Serveur Web démarré")
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RestartServeurWeb", "Erreur: " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Historisation"
        ''' <summary>
        ''' Modifie un historique
        ''' </summary>
        ''' <param name="idsrv"></param>
        ''' <param name="IdDevice"></param>
        ''' <param name="DateTime"></param>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function UpdateHisto(ByVal idsrv As String, ByVal IdDevice As String, ByVal DateTime As String, ByVal Value As String, ByVal OldDateTime As String, ByVal OldValue As String, ByVal Source As String) As Integer Implements IHoMIDom.UpdateHisto
            Try
                Dim retour As String

                If VerifIdSrv(idsrv) = False Then
                    Return 99
                End If

                'correction au cas ou
                If Source = "" Then Source = "Value"
                DateTime = (Convert.ToDateTime(DateTime)).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                OldDateTime = (Convert.ToDateTime(OldDateTime)).ToString("yyyy-MM-dd HH:mm:ss") 'OldDateTime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")

                'Update de la BDD
                'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                retour = sqlite_homidom.nonquery("UPDATE historiques SET dateheure=@parameter0, valeur=@parameter1 WHERE device_id='" & IdDevice & "' AND source='" & Source & "' AND valeur='" & OldValue & "' AND dateheure='" & OldDateTime & "'", DateTime, Value)
                If Mid(retour, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "UpdateHisto", "Erreur Requete sqlite : " & retour)
                Else
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "UpdateHisto", "Mise à jour manuelle d'un relevé : " & IdDevice & " " & OldValue & "->" & Value & " (" & OldDateTime & "->" & DateTime & ")")
                End If

                Return 0
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "UpdateHisto", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Supprime un historique
        ''' </summary>
        ''' <param name="idsrv"></param>
        ''' <param name="IdDevice"></param>
        ''' <param name="DateTime"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DeleteHisto(ByVal idsrv As String, ByVal IdDevice As String, ByVal DateTime As String, ByVal Value As String, ByVal Source As String) As Integer Implements IHoMIDom.DeleteHisto
            Try
                Dim retour As String

                If VerifIdSrv(idsrv) = False Then
                    Return 99
                End If

                'correction au cas ou
                If Source = "" Then Source = "Value"
                DateTime = (Convert.ToDateTime(DateTime)).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")

                'Suppression de la BDD
                'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                retour = sqlite_homidom.nonquery("DELETE FROM historiques WHERE device_id='" & IdDevice & "' AND dateheure='" & DateTime & "' AND source='" & Source & "' AND valeur='" & Value & "'")
                If Mid(retour, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteHisto", "Erreur Requete sqlite : " & retour)
                Else
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "DeleteHisto", "Suppression manuelle d'un relevé : " & IdDevice & " " & DateTime)

                    'decrementation du nombre d'histo du composant
                    Dim _dev As Object = ReturnRealDeviceById(IdDevice)
                    If _dev IsNot Nothing Then
                        _dev.CountHisto -= 1
                        If _dev.CountHisto <= 0 Then _dev.CountHisto = 0
                    End If

                End If

                Return 0
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteHisto", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Ajoute un historique
        ''' </summary>
        ''' <param name="idsrv"></param>
        ''' <param name="IdDevice"></param>
        ''' <param name="DateTime"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AddHisto(ByVal idsrv As String, ByVal IdDevice As String, ByVal DateTime As String, ByVal Value As String, ByVal Source As String) As Integer Implements IHoMIDom.AddHisto
            Try
                Dim retour As String

                If VerifIdSrv(idsrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "AddHisto", "Erreur ID du serveur")
                    Return 99
                End If

                'correction au cas ou
                If Source = "" Then Source = "Value"
                DateTime = (Convert.ToDateTime(DateTime)).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")

                'Ajout dans la BDD
                retour = sqlite_homidom.nonquery("INSERT INTO historiques (device_id,source,dateheure,valeur) VALUES (@parameter0, @parameter1, @parameter2, @parameter3)", IdDevice, Source, DateTime, Value)
                If Mid(retour, 1, 4) = "ERR:" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "AddHisto", "Erreur Requete sqlite : " & retour)
                Else
                    Log(TypeLog.INFO, TypeSource.SERVEUR, "AddHisto", "Ajout Manuel d'un relevé : " & IdDevice & " " & Value & " (" & DateTime & ")")
                End If

                Return 0
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "AddHisto", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        Public Function GetTableDBHisto(ByVal idsrv As String) As DataTable Implements IHoMIDom.GetTableDBHisto
            Try
                If VerifIdSrv(idsrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetTableDBHisto", "Erreur ID du serveur")
                    Return Nothing
                End If

                Dim result As New DataTable
                result.TableName = "ListHisto"
                Dim retour As String
                Dim commande As String = "select * from historiques;"
                retour = sqlite_homidom.query(commande, result, "")
                If UCase(Mid(retour, 1, 3)) <> "ERR" Then
                    If result IsNot Nothing Then
                        Return result
                    Else
                        result = Nothing
                        Return Nothing
                    End If
                Else
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetTableDBHisto", retour)
                    Return Nothing
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetTableDBHisto", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function


        Public Function GetAllListHisto(ByVal idsrv As String) As List(Of Historisation) Implements IHoMIDom.GetAllListHisto
            Try
                If VerifIdSrv(idsrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllListHisto", "Erreur ID du serveur")
                    Return Nothing
                End If

                Dim result As New DataTable
                result.TableName = "ListHisto"
                Dim retour As String
                Dim commande As String = "select distinct source, device_id from historiques;"
                retour = sqlite_homidom.query(commande, result, "")
                If UCase(Mid(retour, 1, 3)) <> "ERR" Then
                    If result IsNot Nothing Then
                        Dim _list As New List(Of Historisation)
                        For i As Integer = 0 To result.Rows.Count - 1
                            Dim a As New Historisation
                            a.Nom = result.Rows.Item(i).Item(0).ToString
                            a.IdDevice = result.Rows.Item(i).Item(1).ToString
                            _list.Add(a)
                            a = Nothing
                        Next
                        result = Nothing
                        Return _list
                        _list = Nothing
                    Else
                        result = Nothing
                        Return Nothing
                    End If
                Else
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllListHisto", retour)
                    Return Nothing
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllListHisto", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne tous les historiques d'un composant</summary>
        ''' <param name="IdSrv">ID serveur</param>
        ''' <param name="IdDevice">Id du composant</param>
        ''' <param name="Source">Value ou autre champ</param>
        ''' <returns>true si le composant a un historique</returns>
        ''' <remarks></remarks>
        Public Function GetHisto(ByVal IdSrv As String, ByVal Source As String, ByVal idDevice As String) As List(Of Historisation) Implements IHoMIDom.GetHisto
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim result As New DataTable("HistoDB")
                Dim retour As String = ""
                Dim commande As String = "select * from historiques where source='" & Source & "' and device_id='" & idDevice & "' ORDER BY dateheure;"
                Dim _list As New List(Of Historisation)

                retour = sqlite_homidom.query(commande, result, "")
                If UCase(Mid(retour, 1, 3)) <> "ERR" Then
                    If result IsNot Nothing Then
                        For i As Integer = 0 To result.Rows.Count - 1
                            Dim a As New Historisation
                            a.Nom = result.Rows.Item(i).Item(2).ToString
                            a.IdDevice = result.Rows.Item(i).Item(1).ToString
                            a.DateTime = CDate(result.Rows.Item(i).Item(3).ToString)
                            a.Value = result.Rows.Item(i).Item(4).ToString
                            _list.Add(a)
                            a = Nothing
                        Next
                        result = Nothing
                        Return _list
                    Else
                        result = Nothing
                        Return Nothing
                    End If
                Else
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHisto", retour)
                    result = Nothing
                    _list = Nothing
                    Return Nothing
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHisto", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Permet d'exécuter une requete SQL sur la table histo
        ''' </summary>
        ''' <param name="IdSrv">ID du serveur</param>
        ''' <param name="Requete">Requête SQL</param>
        ''' <returns>Résultat de la requête sous un type DataTable</returns>
        ''' <remarks></remarks>
        Public Function RequeteSqLHisto(ByVal IdSrv As String, ByVal Requete As String) As DataTable Implements IHoMIDom.RequeteSqLHisto
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RequeteSqLHisto", "Erreur ID du serveur")
                    Return Nothing
                End If

                Dim result As New DataTable("SQLDB")
                Dim retour As String
                Dim commande As String = Requete '"select * from historiques where source='" & Source & "' and device_id='" & idDevice & "' ORDER BY dateheure;"

                'on vérifie que la requête n'est pas vide
                If String.IsNullOrEmpty(commande) Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RequeteSqLHisto", "Erreur la requete est vide")
                    Return Nothing
                Else
                    'on vérifie que la requête commence par select * from historiques
                    If commande.ToUpper.StartsWith("SELECT * FROM HISTORIQUES") Then
                        'on vérifie que la requête ne contient pas de ; pour que l'utilisateur en cherche pas à faire des requetes complexes cachées 
                        If InStr(Mid(commande, 1, commande.Length - 1), ";") > 0 Then
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RequeteSqLHisto", "Erreur la requete ne doit pas comporter de ; sauf à la fin")
                            Return Nothing
                        Else
                            'si tous est ok on exécute la requete
                            Log(TypeLog.DEBUG, TypeSource.SERVEUR, "RequeteSqLHisto", "Execution de la requete: " & commande)
                            retour = sqlite_homidom.query(commande, result, "")
                        End If
                    Else
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RequeteSqLHisto", "Erreur la requete doit commencer par select * from historiques")
                        Return Nothing
                    End If
                End If

                If UCase(Mid(retour, 1, 3)) <> "ERR" Then
                    If result IsNot Nothing Then
                        Return result
                    Else
                        result = Nothing
                        Return Nothing
                    End If
                Else
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RequeteSqLHisto", retour)
                    result = Nothing
                    Return Nothing
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RequeteSqLHisto", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function


        ''' <summary>
        ''' Retourne un datatable d'historique d'un device suivant sa propriété (source) puis suivant une date de début et de fin
        ''' </summary>
        ''' <param name="IdSrv">ID du serveur</param>
        ''' <param name="idDevice">ID du device</param>
        ''' <param name="Source">Source du device (ex: Value)</param>
        ''' <param name="DateStart">Date de départ</param>
        ''' <param name="DateEnd">Date de fin</param>
        ''' <param name="Moyenne">Moyenne par "heure" "jour" ou rien ""</param>
        ''' <returns>Datatable</returns>
        ''' <remarks></remarks>
        Public Function GetHistoDeviceSource(ByVal IdSrv As String, ByVal idDevice As String, ByVal Source As String, Optional ByVal DateStart As String = "", Optional ByVal DateEnd As String = "", Optional ByVal Moyenne As String = "") As List(Of Historisation) Implements IHoMIDom.GetHistoDeviceSource
            Try
                'On vérifie que l'id du serveur est correct pour lancer la fonction sinon erreur
                If VerifIdSrv(IdSrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHistoDeviceSource", "L'Id du serveur est erroné")
                    Return Nothing
                End If

                'On vérifie que datestart et dateend sont bien des dates sinon erreur
                If (IsDate(DateStart) = False And String.IsNullOrEmpty(DateStart) = True) Or (IsDate(DateEnd) = False And String.IsNullOrEmpty(DateEnd) = False) Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHistoDeviceSource", "Erreur DateStart ou DateEnd doivent être une date")
                    Return Nothing
                End If

                'Variables
                Dim result As New DataTable("HistoDB")
                Dim retour As String = ""
                Dim commande As String = ""
                Dim request_valeur As String = "*"
                Dim request_groupby As String = ""
                Dim _list As New List(Of Historisation)

                'prepare la requête suivant les moyennes
                Select Case UCase(Moyenne)
                    Case "HEURE"
                        'request_valeur = "device_id,source,DATE_FORMAT(`dateheure`,  '%Y/%m/%d %H' ) as dateheure, AVG(`valeur`)"
                        'request_groupby = " GROUP BY DATE_FORMAT(`dateheure`,  '%Y/%m/%d %H' )"
                        request_valeur = "device_id,source,strftime('%Y-%m-%d %H:%M:%S',dateheure), AVG(`valeur`)"
                        request_groupby = " GROUP BY strftime('%Y-%m-%d %H',dateheure)"
                    Case "JOUR"
                        'request_valeur = "device_id,source,DATE_FORMAT(`dateheure`,  '%Y/%m/%d' ) as dateheure, AVG(`valeur`)"
                        'request_groupby = " GROUP BY DATE_FORMAT(`dateheure`,  '%Y/%m/%d' )"
                        request_valeur = "device_id,source,date(dateheure), AVG(`valeur`)"
                        request_groupby = " GROUP BY date(dateheure)"
                    Case "", "AUCUNE"
                        request_valeur = "device_id,source,dateheure,valeur"
                        request_groupby = ""
                    Case Else
                        request_valeur = "device_id,source,dateheure,valeur"
                        request_groupby = ""
                End Select

                'Prépare la requête sql suivant datestart et dateend
                If String.IsNullOrEmpty(DateStart) = True And String.IsNullOrEmpty(DateEnd) = True Then
                    commande = "select " & request_valeur & " from historiques where source='" & Source & "' and device_id='" & idDevice & "'" & request_groupby & " ORDER BY dateheure;"
                End If
                If String.IsNullOrEmpty(DateStart) = False And String.IsNullOrEmpty(DateEnd) = True Then
                    commande = "select " & request_valeur & " from historiques where source='" & Source & "' and device_id='" & idDevice & "' and dateheure>='" & DateStart & "'" & request_groupby & " ORDER BY dateheure;"
                End If
                If String.IsNullOrEmpty(DateStart) = True And String.IsNullOrEmpty(DateEnd) = False Then
                    commande = "select " & request_valeur & " from historiques where source='" & Source & "' and device_id='" & idDevice & "' and dateheure<='" & DateEnd & "'" & request_groupby & " ORDER BY dateheure;"
                End If
                If String.IsNullOrEmpty(DateStart) = False And String.IsNullOrEmpty(DateEnd) = False Then
                    commande = "select " & request_valeur & " from historiques where source='" & Source & "' and device_id='" & idDevice & "' and dateheure between '" & DateStart & "' and '" & DateEnd & "'" & request_groupby & " ORDER BY dateheure;"
                End If

                'execute la requête sql
                retour = sqlite_homidom.query(commande, result, "")

                'Vérifie que la requête n'a pas générée d'erreur
                If UCase(Mid(retour, 1, 3)) = "ERR" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHistoDeviceSource", retour)
                Else
                    For i As Integer = 0 To result.Rows.Count - 1
                        Dim a As New Historisation
                        a.IdDevice = result.Rows.Item(i).Item(0).ToString
                        a.Nom = result.Rows.Item(i).Item(1).ToString
                        a.DateTime = CDate(result.Rows.Item(i).Item(2).ToString)
                        a.Value = result.Rows.Item(i).Item(3).ToString
                        _list.Add(a)
                        a = Nothing
                    Next
                End If

                result = Nothing
                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHistoDeviceSource", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Permet de savoir si un device a des historiques associés dans la BDD</summary>
        ''' <param name="IdDevice"></param>
        ''' <param name="Source"></param>
        ''' <returns>nombre d'historique</returns>
        ''' <remarks></remarks>
        Public Function DeviceAsHisto(ByVal IdDevice As String, Optional ByVal Source As String = "") As Long Implements IHoMIDom.DeviceAsHisto
            Try
                Dim commande As String
                Dim retour As String = ""
                Dim result As Long = 0

                If Source = "" Then
                    commande = "SELECT COUNT(*) FROM (SELECT 'rowid',* FROM 'historiques') WHERE device_id='" & IdDevice & "' ;"
                Else
                    commande = "SELECT COUNT(*) FROM historiques WHERE source='" & Source & "' and device_id='" & IdDevice & "' ;"
                End If

                retour = sqlite_homidom.count(commande, result)

                If UCase(Mid(retour, 1, 3)) = "ERR" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceAsHisto", "Erreur: " & retour)
                End If

                Return result
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceAsHisto", "Exception : " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Retourne un dictionnary retourant en clé l'id du device et la valeur True/False s'il contient des historiques
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DevicesAsHisto() As Dictionary(Of String, Long) Implements IHoMIDom.DevicesAsHisto
            Try
                Dim Retour As New Dictionary(Of String, Long)
                Dim commande As String
                Dim IdDevice As String
                Dim result As Integer = 0
                Dim retourDB As String

                For i As Integer = 0 To _ListDevices.Count - 1
                    IdDevice = _ListDevices.Item(i).ID

                    commande = "SELECT COUNT(*) FROM (SELECT 'rowid',* FROM 'historiques') WHERE device_id='" & IdDevice & "' ;"
                    retourDB = sqlite_homidom.count(commande, result)

                    If UCase(Mid(retourDB, 1, 3)) <> "ERR" Then
                        If result >= 0 Then
                            Retour.Add(IdDevice, result)
                        End If
                    Else
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DevicesAsHisto", "Erreur: " & retourDB)
                        Exit For
                    End If
                Next

                Return Retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DevicesAsHisto", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Importation d'historiques</summary>          
        ''' <param name="fichier">Fichier texte CSV</param>          
        ''' <param name="separateur">caractère de séparation, si omis point-virgule</param>          
        ''' <returns>String si OK, String "ERR:..." si erreur</returns>          
        ''' <remarks></remarks>          
        Public Function ImportHisto(ByVal fichier As String, Optional ByVal separateur As String = ";") As String Implements IHoMIDom.ImportHisto
            ' Le format d'importation est un fichier CSV codé en ANSI type export Microsoft Excel.
            ' Les fichiers produits avec la fonction d'export des historiques dans le module Admin sont compatibles.

            Dim DeleteFileAfterImport As Boolean = False
            If fichier.StartsWith("Histo\") = True Then
                ' Le fichier a été uploadé, le chemin indiqué en paramètre est relatif
                fichier = _MonRepertoire & "\Fichiers\" & fichier
                DeleteFileAfterImport = True
            End If
            Try
                If Not IO.File.Exists(fichier) Then
                    Return "ERR: fichier non trouvé."
                End If

                Dim deviceID_courant As String = ""
                Dim nom_courant As String = ""
                Dim type_courant As HoMIDom.Device.ListeDevices
                Dim _dev As TemplateDevice
                Dim lignes() = IO.File.ReadAllLines(fichier, Encoding.Default)
                Dim colonnes() As String
                Dim nb_lignes As Long = lignes.Count
                Dim ColDeviceID As Integer = -1
                Dim ColDateTime As Integer = -1
                Dim ColName As Integer = -1
                Dim ColValue As Integer = -1
                Dim ColSource As Integer = -1

                ' première ligne: on contrôle les colonnes.
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation en cours. ", "Vérification des champs.")
                colonnes = lignes(0).Split(separateur)
                For a As Integer = 0 To colonnes.Count - 1
                    If colonnes(a).ToLower = "datetime" Or colonnes(a).ToLower = "dateheure" Then
                        ColDateTime = a
                    ElseIf colonnes(a).ToLower = "name" Or colonnes(a).ToLower = "nom" Then
                        ColName = a
                    ElseIf colonnes(a).ToLower = "deviceid" Or colonnes(a).ToLower = "device_id" Then
                        ColDeviceID = a
                    ElseIf colonnes(a).ToLower = "value" Or colonnes(a).ToLower = "valeur" Then
                        ColValue = a
                    ElseIf colonnes(a).ToLower = "source" Then
                        ColSource = a
                    End If
                Next
                If ColDateTime = -1 Or ColValue = -1 Or (ColName = -1 And ColDeviceID = -1) Then
                    ' On n'a pas les colonnes nécessaires (au minimum: date, valeur et nom/device_id).
                    Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation interrompue. ", "Les champs nécessaires n'ont pas été trouvés.")
                    Return "ERR: Erreur à l'importation de " & fichier & ". Erreur: la première ligne doit contenir les noms de champs appropriés."
                    If DeleteFileAfterImport = True Then IO.File.Delete(fichier)
                    Exit Function
                End If
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation en cours. ", "Vérification des champs terminée.")

                ' Vérification de tous les enregistrements
                For i As Long = 1 To nb_lignes - 1
                    If (i + 1) Mod 1000 = 0 Then
                        Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation en cours... ", "Vérification ligne " & (i + 1).ToString & ".")
                    End If
                    colonnes = lignes(i).Split(separateur)

                    ' Contrôle de la date
                    If IsDate(colonnes(ColDateTime)) = False Then
                        Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation interrompue à la ligne " & (i + 1).ToString, ". La date est invalide.")
                        Return "ERR: Erreur à la vérification de la ligne " & (i + 1).ToString & ". La date est invalide."
                        If DeleteFileAfterImport = True Then IO.File.Delete(fichier)
                        Exit Function
                    End If

                    ' Contrôle du device_id ou du nom
                    If ColDeviceID > -1 Then
                        ' on privilégie la colonne deviceID et on contrôle que l'ID existe dans le système
                        ' deviceID_courant est utilisé pour ne pas faire lancer la boucle de recherche à chaque ligne,
                        ' mais seulement si le deviceID_courant a changé (très probablement il y a de nombreuses lignes avec le même nom)
                        If colonnes(ColDeviceID) <> deviceID_courant Then
                            ' But: accélérer l'opération en ne recontrôlant les device id que s'il y a un changement
                            ' suite du fichier
                            _dev = ReturnDeviceById(_IdSrv, colonnes(ColDeviceID))
                            If _dev.ID = colonnes(ColDeviceID) Then
                                deviceID_courant = colonnes(ColDeviceID)
                                type_courant = _dev.Type
                            Else
                                ' Le composant n'a pas été trouvé. On stoppe l'importation
                                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation interrompue à la ligne " & (i + 1).ToString, ". Le composant " & colonnes(ColDeviceID) & " n'a pas été trouvé.")
                                Return "ERR: Erreur à la vérification de la ligne " & (i + 1).ToString & ". Le composant " & colonnes(ColDeviceID) & " n'a pas été trouvé."
                                If DeleteFileAfterImport = True Then IO.File.Delete(fichier)
                                Exit Function
                            End If
                        End If
                    Else
                        ' c'est un nom de composant donc il faut retrouver le deviceid
                        ' nom_courant est utilisé pour ne pas faire lancer la boucle de recherche à chaque ligne,
                        ' mais seulement si le nom a changé (très probablement il y a de nombreuses lignes avec le même nom)
                        If colonnes(ColName) <> nom_courant Then
                            nom_courant = ""
                            For Each _dev In GetAllDevices(_IdSrv)
                                If _dev.Name = colonnes(ColName) Then
                                    nom_courant = colonnes(ColName)
                                    deviceID_courant = _dev.ID
                                    type_courant = _dev.Type
                                    ' on remplace le nom par le device ID pour l'importation des données dans la BDD
                                    colonnes(ColName) = _dev.ID
                                    Exit For
                                End If
                            Next
                            If nom_courant = "" Then
                                ' Le composant n'a pas été trouvé. On stoppe l'importation
                                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation interrompue à la ligne " & (i + 1).ToString, ". Le composant " & colonnes(ColName) & " n'a pas été trouvé.")
                                Return "ERR: Erreur à la vérification de la ligne " & (i + 1).ToString & ". Le composant " & colonnes(ColName) & " n'a pas été trouvé."
                                If DeleteFileAfterImport = True Then IO.File.Delete(fichier)
                                Exit Function
                            End If
                        Else
                            colonnes(ColName) = deviceID_courant
                        End If
                    End If
                    Select Case type_courant
                        Case Device.ListeDevices.GENERIQUEBOOLEEN
                            If colonnes(ColValue) = "True" Or "TRUE" Or "true" Or "Vrai" Or "VRAI" Or "vrai" Or "1" Or "-1" Then
                                colonnes(ColValue) = "True"
                            ElseIf colonnes(ColValue) = "False" Or "FALSE" Or "false" Or "Faux" Or "FAUX" Or "faux" Or "0" Then
                                colonnes(ColValue) = "False"
                            Else
                                ' La valeur n'est pas interprétable. On stoppe l'importation.
                                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation interrompue à la ligne " & (i + 1).ToString, ". La valeur du composant " & colonnes(ColName) & " doit être de type Vrai/Faux.")
                                Return "ERR: Erreur à la vérification de la ligne " & (i + 1).ToString & ". La valeur du composant " & colonnes(ColName) & " doit être de type Vrai/Faux."
                                If DeleteFileAfterImport = True Then IO.File.Delete(fichier)
                                Exit Function
                            End If
                    End Select
                    ' Finalement, on reconstruit la ligne avec les valeurs éventuellement modifées des colonnes
                    For a As Integer = 0 To colonnes.Count - 1
                        If a = 0 Then
                            lignes(i) = colonnes(0)
                        Else
                            lignes(i) = lignes(i) & separateur & colonnes(a)
                        End If
                    Next
                Next
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Vérification terminée. ", (nb_lignes - 1).ToString & " enregistrements vérifiés.")

                ' Importation finale dans la BDD
                Dim retour As String
                Dim source As String
                If ColDeviceID = -1 Then
                    ColDeviceID = ColName
                End If
                For i As Long = 1 To nb_lignes - 1
                    If (i + 1) Mod 100 = 0 Then
                        Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation en cours... ", "Enregistrement ligne " & (i + 1).ToString & ".")
                    End If
                    colonnes = lignes(i).Split(separateur)
                    If ColSource = -1 Then
                        source = "Value"
                    Else
                        source = colonnes(ColSource)
                    End If
                    retour = sqlite_homidom.nonquery("INSERT INTO historiques (device_id,source,dateheure,valeur) VALUES (@parameter0, @parameter1, @parameter2, @parameter3)", colonnes(ColDeviceID), source, Format(DateAndTime.DateValue(colonnes(ColDateTime)), "yyyy-MM-dd") & " " & Format(DateAndTime.TimeValue(colonnes(ColDateTime)), "hh:mm:ss"), colonnes(ColValue).Replace(",", "."))
                    If Mid(retour, 1, 4) = "ERR:" Then
                        Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation interrompue à la ligne " & (i + 1).ToString, retour)
                        Return "ERR: Erreur à la vérification de la ligne " & (i + 1).ToString & ": " & retour
                        If DeleteFileAfterImport = True Then IO.File.Delete(fichier)
                        Exit Function
                    End If
                Next
                Log(TypeLog.DEBUG, TypeSource.SERVEUR, "Importation terminée. ", (nb_lignes - 1).ToString & " enregistrements importés.")
                If DeleteFileAfterImport = True Then IO.File.Delete(fichier)
                Return (nb_lignes - 1).ToString & " enregistrements importés avec succès."

            Catch ex As Exception
                If DeleteFileAfterImport = True Then IO.File.Delete(fichier)
                Return "ERR: Erreur à l'importation de " & fichier & ". Erreur: " & ex.Message
            End Try
        End Function


        Public Function VerifPurge() As Integer Implements IHoMIDom.VerifPurge

            Try

                Dim Source As String = "Value"
                Dim result As New List(Of Historisation)
                Dim result1 As New List(Of Historisation)
                Dim DateTime As String = ""
                Dim DateStart As String = ""

                For Each _dev In _ListDevices
                    If _dev.Purge > 0 Then

                        result = New List(Of Historisation)
                        DateTime = ""
                        DateStart = ""

                        'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                        DateTime = Now.AddDays(_dev.Purge * -1).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                        DateStart = "2010-01-01 01:01:01" 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")

                        result = GetHistoDeviceSource(_IdSrv, _dev.ID, Source, DateStart, DateTime)

                        If result IsNot Nothing Then
                            For i As Integer = 0 To result.Count - 1

                                'Suppression de la BDD
                                DeleteHisto(_IdSrv, _dev.ID, result.Item(i).DateTime, result.Item(i).Value, Source)

                                'decrementation du nombre d'histo du composant
                                If _dev IsNot Nothing Then
                                    _dev.CountHisto -= 1
                                    If _dev.CountHisto <= 0 Then _dev.CountHisto = 0
                                End If
                            Next

                        End If
                    End If
                    If _dev.MoyJour > 0 And (_dev.Purge > _dev.MoyJour Or _dev.Purge = 0) Then

                        result = New List(Of Historisation)
                        result1 = New List(Of Historisation)
                        DateTime = ""
                        DateStart = ""

                        'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                        DateTime = Now.AddDays(_dev.MoyJour * -1).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                        If _dev.Purge > _dev.MoyJour Then
                            DateStart = Now.AddDays(_dev.Purge * -1).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                        Else
                            DateStart = "2010-01-01 01:01:01" 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                        End If

                        result = GetHistoDeviceSource(_IdSrv, _dev.ID, Source, DateStart, DateTime)
                        result1 = GetHistoDeviceSource(_IdSrv, _dev.ID, Source, DateStart, DateTime, "JOUR")

                        If result IsNot Nothing And result1 IsNot Nothing And DateStart <> "" Then
                            If result.Count > result1.Count Then
                                For i As Integer = 0 To result.Count - 1

                                    'Suppression de la BDD
                                    DeleteHisto(_IdSrv, _dev.ID, result.Item(i).DateTime, result.Item(i).Value, Source)

                                    'decrementation du nombre d'histo du composant
                                    If _dev IsNot Nothing Then
                                        _dev.CountHisto -= 1
                                        If _dev.CountHisto <= 0 Then _dev.CountHisto = 0
                                    End If
                                Next
                                For i As Integer = 0 To result1.Count - 1

                                    'Ajout dans la BDD
                                    AddHisto(_IdSrv, _dev.ID, result1.Item(i).DateTime, result1.Item(i).Value, Source)

                                    'incrementation du nombre d'histo du composant
                                    If _dev IsNot Nothing Then
                                        _dev.CountHisto += 1
                                    End If
                                Next
                            End If
                        End If
                    End If

                    If _dev.MoyHeure > 0 And (_dev.MoyJour > _dev.MoyHeure Or _dev.MoyJour = 0) And ((_dev.Purge > _dev.MoyHeure And _dev.Purge > 0) Or _dev.Purge = 0) Then

                        result = New List(Of Historisation)
                        result1 = New List(Of Historisation)
                        DateTime = ""
                        DateStart = ""

                        'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                        DateTime = Now.AddDays(_dev.MoyHeure * -1).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                        If _dev.MoyJour = 0 Then
                            If _dev.Purge = 0 Then
                                DateStart = "2010-01-01 01:01:01" 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                            Else
                                If _dev.Purge > _dev.MoyHeure Then
                                    DateStart = Now.AddDays(_dev.Purge * -1).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                                Else
                                    DateStart = "2010-01-01 01:01:01" 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")									
                                End If
                            End If
                        Else
                            If _dev.MoyJour > _dev.MoyHeure Then
                                DateStart = Now.AddDays(_dev.MoyJour * -1).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                            Else
                                If _dev.Purge = 0 Then
                                    DateStart = "2010-01-01 01:01:01" 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                                Else
                                    If _dev.Purge > _dev.MoyHeure Then
                                        DateStart = Now.AddDays(_dev.Purge * -1).ToString("yyyy-MM-dd HH:mm:ss") 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")
                                    Else
                                        DateStart = "2010-01-01 01:01:01" 'datetime doit etre au format ToString("yyyy-MM-dd HH:mm:ss")									
                                    End If
                                End If
                            End If
                        End If

                        result = GetHistoDeviceSource(_IdSrv, _dev.ID, Source, DateStart, DateTime)
                        result1 = GetHistoDeviceSource(_IdSrv, _dev.ID, Source, DateStart, DateTime, "HEURE")

                        If result IsNot Nothing And result1 IsNot Nothing And DateStart <> "" Then
                            If result.Count > result1.Count Then
                                For i As Integer = 0 To result.Count - 1

                                    'Suppression de la BDD
                                    DeleteHisto(_IdSrv, _dev.ID, result.Item(i).DateTime, result.Item(i).Value, Source)

                                    'decrementation du nombre d'histo du composant
                                    If _dev IsNot Nothing Then
                                        _dev.CountHisto -= 1
                                        If _dev.CountHisto <= 0 Then _dev.CountHisto = 0
                                    End If
                                Next
                                For i As Integer = 0 To result1.Count - 1

                                    'Ajout dans la BDD
                                    AddHisto(_IdSrv, _dev.ID, result1.Item(i).DateTime, result1.Item(i).Value, Source)

                                    'incrementation du nombre d'histo du composant
                                    If _dev IsNot Nothing Then
                                        _dev.CountHisto += 1
                                    End If
                                Next
                            End If
                        End If
                    End If
                Next

                Return 0
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteHisto", "Exception : " & ex.Message)
                Return -1
            End Try

        End Function


#End Region

#Region "Audio"

        Public Function Parler(ByVal Message As String) As Boolean Implements IHoMIDom.Parler
            Try
                Dim texte As String = Message
                'remplace les balises par la valeur
                'texte = texte.Replace("{time}", Now.ToShortTimeString)
                'texte = texte.Replace("{date}", Now.ToLongDateString)
                'texte = Decodestring(texte)

                Dim lamachineaparler As New Speech.Synthesis.SpeechSynthesizer
                Log(Server.TypeLog.DEBUG, Server.TypeSource.SCRIPT, "Parler", "Message: " & texte)
                With lamachineaparler
                    .SelectVoice(GetDefautVoice)
                    '.SetOutputToWaveFile("C:\tet.wav")
                    '.SetOutputToWaveFile(File)
                    .SpeakAsync(texte)
                End With

                texte = Nothing
                lamachineaparler = Nothing
            Catch ex As Exception
                Log(Server.TypeLog.ERREUR, Server.TypeSource.SCRIPT, "Parler", "Exception lors de l'annonce du message: " & Message & " : " & ex.ToString)
            End Try
        End Function

#End Region

#Region "SMTP"
        ''' <summary>
        ''' Permet de tester l'envoi de mail
        ''' </summary>
        ''' <param name="IdSrv"></param>
        ''' <param name="Adresse"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function TestSendMail(ByVal IdSrv As String, ByVal De As String, ByVal Adresse As String, ByVal smtpserveur As String, ByVal Port As Integer, ByVal SSL As Boolean, Optional ByVal Login As String = "", Optional ByVal Password As String = "") As String Implements IHoMIDom.TestSendMail
            Try
                Dim _action As New Mail(Me, De, Adresse, "Test EnvoiMail Homidom", "Test EnvoiMail homidom à " & Now.ToString, smtpserveur, Port, SSL, Login, Password)
                Dim y As New Thread(AddressOf _action.Send_email)
                y.Name = "Traitement du script"
                y.Start()
                y = Nothing

                Return "0"
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "TestSendMail", "Exception : " & ex.Message)
                Return "Erreur lors l'envoi du mail de test: " & ex.Message
            End Try
        End Function

        ''' <summary>Retourne l'adresse du serveur SMTP</summary>
        Public Function GetSMTPServeur(ByVal IdSrv As String) As String Implements IHoMIDom.GetSMTPServeur
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Return _SMTPServeur
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSMTPServeur", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Fixe l'adresse du serveur SMTP</summary>
        Public Sub SetSMTPServeur(ByVal IdSrv As String, ByVal Value As String) Implements IHoMIDom.SetSMTPServeur
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                _SMTPServeur = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetSMTPServeur", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne le login du serveur SMTP</summary>
        Public Function GetSMTPLogin(ByVal IdSrv As String) As String Implements IHoMIDom.GetSMTPLogin
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Return _SMTPLogin
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSMTPLogin", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Fixe le login du serveur SMTP</summary>
        Public Sub SetSMTPLogin(ByVal IDSrv As String, ByVal Value As String) Implements IHoMIDom.SetSMTPLogin
            Try
                If VerifIdSrv(IDSrv) = False Then
                    Exit Sub
                End If

                _SMTPLogin = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetSMTPLogin", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne le password du serveur SMTP</summary>
        Public Function GetSMTPPassword(ByVal IdSrv As String) As String Implements IHoMIDom.GetSMTPPassword
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Return _SMTPassword
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSMTPPassword", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Fixe le password du serveur SMTP</summary>
        Public Sub SetSMTPPassword(ByVal IdSrv As String, ByVal Value As String) Implements IHoMIDom.SetSMTPPassword
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                _SMTPassword = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetSMTPPassword", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne l'adresse mail du serveur</summary>
        Public Function GetSMTPMailServeur(ByVal IdSrv As String) As String Implements IHoMIDom.GetSMTPMailServeur
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Return _SMTPmailEmetteur
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSMTPMailServeur", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Fixe le password du serveur SMTP</summary>
        Public Sub SetSMTPMailServeur(ByVal IdSrv As String, ByVal Value As String) Implements IHoMIDom.SetSMTPMailServeur
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                _SMTPmailEmetteur = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetSMTPPassword", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne le port SMTP à utiliser</summary>
        Public Function GetSMTPPort(ByVal IdSrv As String) As Integer Implements IHoMIDom.GetSMTPPort
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Return _SMTPPort
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSMTPPort", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Fixe le port SMTP</summary>
        Public Sub SetSMTPPort(ByVal IdSrv As String, ByVal Value As Integer) Implements IHoMIDom.SetSMTPPort
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                _SMTPPort = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetSMTPPort", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne si on doit utiliser une connexion SLL pour SMTP</summary>
        Public Function GetSMTPSSL(ByVal IdSrv As String) As Boolean Implements IHoMIDom.GetSMTPSSL
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Return _SMTPSSL
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSMTPSSL", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Fixe si on doit utiliser une connexion SLL pour SMTP</summary>
        Public Sub SetSMTPSSL(ByVal IdSrv As String, ByVal Value As Boolean) Implements IHoMIDom.SetSMTPSSL
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                _SMTPSSL = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetSMTPSSL", "Exception : " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Gestion Soleil"
        ''' <summary>Retourne l'heure du couché du soleil</summary>
        Function GetHeureCoucherSoleil() As String Implements IHoMIDom.GetHeureCoucherSoleil
            Try
                Return _HeureCoucherSoleil
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHeureCoucherSoleil", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

        ''' <summary>Retour l'heure de lever du soleil</summary>
        Function GetHeureLeverSoleil() As String Implements IHoMIDom.GetHeureLeverSoleil
            Try
                Return _HeureLeverSoleil
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHeureLeverSoleil", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

        ''' <summary>Retourne la longitude du serveur</summary>
        Function GetLongitude() As Double Implements IHoMIDom.GetLongitude
            Try
                Return _Longitude
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetLongitude", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Applique une valeur de longitude au serveur</summary>
        ''' <param name="value"></param>
        Sub SetLongitude(ByVal IdSrv As String, ByVal value As Double) Implements IHoMIDom.SetLongitude
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                If _Longitude <> value Then
                    _Longitude = value
                    MAJ_HeuresSoleil()
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetLongitude", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne la latitude du serveur</summary>
        Function GetLatitude() As Double Implements IHoMIDom.GetLatitude
            Try
                Return _Latitude
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetLatitude", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Applique une valeur de latitude du serveur</summary>
        ''' <param name="value"></param>
        Sub SetLatitude(ByVal IdSrv As String, ByVal value As Double) Implements IHoMIDom.SetLatitude
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                If _Latitude <> value Then
                    _Latitude = value
                    MAJ_HeuresSoleil()
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetLatitude", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne la valeur de correction de l'heure de coucher du soleil</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetHeureCorrectionCoucher() As Integer Implements IHoMIDom.GetHeureCorrectionCoucher
            Try

                Return _HeureCoucherSoleilCorrection
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHeureCorrectionCoucher", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Applique la valeur de correction de l'heure de coucher du soleil</summary>
        ''' <param name="value"></param>
        ''' <remarks></remarks>
        Sub SetHeureCorrectionCoucher(ByVal IdSrv As String, ByVal value As Integer) Implements IHoMIDom.SetHeureCorrectionCoucher
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                If _HeureCoucherSoleilCorrection <> value Then
                    _HeureCoucherSoleilCorrection = value
                    MAJ_HeuresSoleil()
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetHeureCorrectionCoucher", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne la valeur de correction de l'heure de lever du soleil</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetHeureCorrectionLever() As Integer Implements IHoMIDom.GetHeureCorrectionLever
            Try
                Return _HeureLeverSoleilCorrection
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetHeureCorrectionLever", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Applique la valeur de correction de l'heure de coucher du soleil</summary>
        ''' <param name="value"></param>
        ''' <remarks></remarks>
        Sub SetHeureCorrectionLever(ByVal IdSrv As String, ByVal value As Integer) Implements IHoMIDom.SetHeureCorrectionLever
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                If _HeureLeverSoleilCorrection <> value Then
                    _HeureLeverSoleilCorrection = value
                    MAJ_HeuresSoleil()
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetHeureCorrectionLever", "Exception : " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Driver"
        ''' <summary>Supprimer un driver de la config</summary>
        ''' <param name="driverId"></param>
        Public Function DeleteDriver(ByVal IdSrv As String, ByVal driverId As String) As Integer Implements IHoMIDom.DeleteDriver
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If


                If driverId = "DE96B466-2540-11E0-A321-65D7DFD72085" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDriver", "La suppression du driver Virtuel est impossible")
                    Return -1
                End If

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).Id = driverId Then
                        _ListDrivers.RemoveAt(i)
                        SaveRealTime()
                        Return 0
                        Exit For
                    End If
                Next

                Return -1
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDriver", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Arrête un driver par son Id</summary>
        ''' <param name="DriverId"></param>
        ''' <remarks></remarks>
        Public Sub StopDriver(ByVal IdSrv As String, ByVal DriverId As String) Implements IHoMIDom.StopDriver
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).id = DriverId Then
                        _ListDrivers.Item(i).stop()
                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "StopDriver", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Démarre un driver par son id</summary>
        ''' <param name="DriverId"></param>
        ''' <remarks></remarks>
        Public Sub StartDriver(ByVal IdSrv As String, ByVal DriverId As String) Implements IHoMIDom.StartDriver
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).id = DriverId Then
                        _ListDrivers.Item(i).start()
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "StartDriver", "Exception : " & ex.Message)
            End Try
        End Sub

        Public Function VerifChamp(ByVal Idsrv As String, ByVal DriverId As String, ByVal Champ As String, ByVal Value As Object) As String Implements IHoMIDom.VerifChamp
            Try
                If VerifIdSrv(Idsrv) = False Then
                    Return "L'ID du serveur est erroné pour pouvoir exécuter cette fonction VERIFCHAMP"
                End If

                Dim retour As String = "0"

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).id = DriverId Then
                        retour = _ListDrivers.Item(i).VerifChamp(Champ, Value)
                        Exit For
                    End If
                Next
                Return retour
            Catch ex As Exception
                Return "Une erreur est apparue lors de la vérification du champ " & Champ & ": " & ex.Message
            End Try
        End Function

        ''' <summary>Retourne la liste de tous les drivers</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetAllDrivers(ByVal IdSrv As String) As List(Of TemplateDriver) Implements IHoMIDom.GetAllDrivers
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim _list As New List(Of TemplateDriver)

                For i As Integer = 0 To _ListDrivers.Count - 1
                    Dim x As New TemplateDriver
                    With x
                        .Nom = _ListDrivers.Item(i).nom
                        .ID = _ListDrivers.Item(i).id
                        .COM = _ListDrivers.Item(i).com
                        .Description = _ListDrivers.Item(i).description
                        .Enable = _ListDrivers.Item(i).enable
                        .IP_TCP = _ListDrivers.Item(i).ip_tcp
                        .IP_UDP = _ListDrivers.Item(i).ip_udp
                        .IsConnect = _ListDrivers.Item(i).isconnect
                        .Modele = _ListDrivers.Item(i).modele
                        .Picture = _ListDrivers.Item(i).picture
                        .Port_TCP = _ListDrivers.Item(i).port_tcp
                        .Port_UDP = _ListDrivers.Item(i).port_udp
                        .Protocol = _ListDrivers.Item(i).protocol
                        .Refresh = _ListDrivers.Item(i).refresh
                        .StartAuto = _ListDrivers.Item(i).startauto
                        .AutoDiscover = _ListDrivers.Item(i).autoDiscover
                        .Version = _ListDrivers.Item(i).version
                        For j As Integer = 0 To _ListDrivers.Item(i).DeviceSupport.count - 1
                            .DeviceSupport.Add(_ListDrivers.Item(i).devicesupport.item(j).ToString)
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).Parametres.count - 1
                            Dim y As New Driver.Parametre
                            y.Nom = _ListDrivers.Item(i).Parametres.item(j).nom
                            y.Description = _ListDrivers.Item(i).Parametres.item(j).description
                            y.Valeur = _ListDrivers.Item(i).Parametres.item(j).valeur
                            .Parametres.Add(y)
                            y = Nothing
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).LabelsDriver.count - 1
                            Dim y As New Driver.cLabels
                            y.NomChamp = _ListDrivers.Item(i).LabelsDriver.item(j).NomChamp
                            y.LabelChamp = _ListDrivers.Item(i).LabelsDriver.item(j).LabelChamp
                            y.Tooltip = _ListDrivers.Item(i).LabelsDriver.item(j).Tooltip
                            y.Parametre = _ListDrivers.Item(i).LabelsDriver.item(j).Parametre
                            .LabelsDriver.Add(y)
                            y = Nothing
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).LabelsDevice.count - 1
                            Dim y As New Driver.cLabels
                            y.NomChamp = _ListDrivers.Item(i).LabelsDevice.item(j).NomChamp
                            y.LabelChamp = _ListDrivers.Item(i).LabelsDevice.item(j).LabelChamp
                            y.Tooltip = _ListDrivers.Item(i).LabelsDevice.item(j).Tooltip
                            y.Parametre = _ListDrivers.Item(i).LabelsDevice.item(j).Parametre
                            .LabelsDevice.Add(y)
                            y = Nothing
                        Next
                        'Dim _listactdrv As New ArrayList
                        Dim _listactd As New List(Of String)
                        For j As Integer = 0 To Api.ListMethod(_ListDrivers.Item(i)).Count - 1
                            _listactd.Add(Api.ListMethod(_ListDrivers.Item(i)).Item(j).ToString)
                        Next
                        If _listactd.Count > 0 Then
                            For n As Integer = 0 To _listactd.Count - 1
                                Dim a() As String = _listactd.Item(n).Split("|")
                                Dim p As New DeviceAction
                                With p
                                    .Nom = a(0)
                                    If a.Length > 1 Then
                                        For t As Integer = 1 To a.Length - 1
                                            Dim pr As New DeviceAction.Parametre
                                            Dim b() As String = a(t).Split(":")
                                            With pr
                                                .Nom = b(0)
                                                .Type = b(1)
                                            End With
                                            p.Parametres.Add(pr)
                                        Next
                                    End If
                                End With
                                .DeviceAction.Add(p)
                                p = Nothing
                            Next
                        End If

                        _listactd = Nothing
                        '_listactdrv = Nothing
                    End With
                    _list.Add(x)
                    x = Nothing
                Next
                _list.Sort(AddressOf sortDriver)
                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllDrivers", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function sortDriver(ByVal x As TemplateDriver, ByVal y As TemplateDriver) As Integer
            Try
                Return x.Nom.CompareTo(y.Nom)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "sortDriver", "Exception : " & ex.Message)
                Return 0
            End Try
        End Function

        ''' <summary>Sauvegarde ou créer un driver dans la config</summary>
        ''' <param name="driverId"></param>
        ''' <param name="name"></param>
        ''' <param name="enable"></param>
        ''' <param name="startauto"></param>
        ''' <param name="iptcp"></param>
        ''' <param name="porttcp"></param>
        ''' <param name="ipudp"></param>
        ''' <param name="portudp"></param>
        ''' <param name="com"></param>
        ''' <param name="refresh"></param>
        ''' <param name="picture"></param>
        ''' <param name="modele"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SaveDriver(ByVal IdSrv As String, ByVal driverId As String, ByVal name As String, ByVal enable As Boolean, ByVal startauto As Boolean, ByVal iptcp As String, ByVal porttcp As String, ByVal ipudp As String, ByVal portudp As String, ByVal com As String, ByVal refresh As Integer, ByVal picture As String, ByVal modele As String, ByVal autodiscover As Boolean, Optional ByVal Parametres As ArrayList = Nothing) As String Implements IHoMIDom.SaveDriver
            Try

                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim myID As String
                'Driver Existant
                myID = driverId

                'verification pour ne pas modifier le driver virtuel
                If driverId = "DE96B466-2540-11E0-A321-65D7DFD72085" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveDriver", "La modification du driver Virtuel est impossible")
                    Return -1
                End If

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).id = driverId Then
                        _ListDrivers.Item(i).Enable = enable
                        _ListDrivers.Item(i).StartAuto = startauto
                        If _ListDrivers.Item(i).IP_TCP <> "@" Then _ListDrivers.Item(i).IP_TCP = iptcp
                        If _ListDrivers.Item(i).Port_TCP <> "@" Then _ListDrivers.Item(i).Port_TCP = porttcp
                        If _ListDrivers.Item(i).IP_UDP <> "@" Then _ListDrivers.Item(i).IP_UDP = ipudp
                        If _ListDrivers.Item(i).Port_UDP <> "@" Then _ListDrivers.Item(i).Port_UDP = portudp
                        If _ListDrivers.Item(i).Com <> "@" Then _ListDrivers.Item(i).Com = com
                        _ListDrivers.Item(i).Refresh = refresh
                        _ListDrivers.Item(i).Picture = picture
                        _ListDrivers.Item(i).Modele = modele
                        _ListDrivers.Item(i).Autodiscover = autodiscover
                        If Parametres IsNot Nothing Then
                            For j As Integer = 0 To Parametres.Count - 1
                                _ListDrivers.Item(i).parametres.item(j).valeur = Parametres.Item(j)
                            Next
                        End If
                        SaveRealTime()
                    End If
                Next


                ManagerSequences.AddSequences(Sequence.TypeOfSequence.Driver, myID, Nothing, Nothing)
                'génération de l'event
                Return myID
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveDriver", "Exception : " & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>Retourne un driver par son ID</summary>
        ''' <param name="DriverId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnDriverById(ByVal IdSrv As String, ByVal DriverId As String) As TemplateDriver Implements IHoMIDom.ReturnDriverByID
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim retour As New TemplateDriver

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).ID = DriverId Then
                        retour.Nom = _ListDrivers.Item(i).nom
                        retour.ID = _ListDrivers.Item(i).id
                        retour.COM = _ListDrivers.Item(i).com
                        retour.Description = _ListDrivers.Item(i).description
                        retour.Enable = _ListDrivers.Item(i).enable
                        retour.IP_TCP = _ListDrivers.Item(i).ip_tcp
                        retour.IP_UDP = _ListDrivers.Item(i).ip_udp
                        retour.IsConnect = _ListDrivers.Item(i).isconnect
                        retour.Modele = _ListDrivers.Item(i).modele
                        retour.Picture = _ListDrivers.Item(i).picture
                        retour.Port_TCP = _ListDrivers.Item(i).port_tcp
                        retour.Port_UDP = _ListDrivers.Item(i).port_udp
                        retour.Protocol = _ListDrivers.Item(i).protocol
                        retour.Refresh = _ListDrivers.Item(i).refresh
                        retour.StartAuto = _ListDrivers.Item(i).startauto
                        retour.Version = _ListDrivers.Item(i).version
                        retour.AutoDiscover = _ListDrivers.Item(i).autoDiscover
                        For j As Integer = 0 To _ListDrivers.Item(i).DeviceSupport.count - 1
                            retour.DeviceSupport.Add(_ListDrivers.Item(i).devicesupport.item(j).ToString)
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).Parametres.count - 1
                            Dim y As New Driver.Parametre
                            y.Nom = _ListDrivers.Item(i).Parametres.item(j).nom
                            y.Description = _ListDrivers.Item(i).Parametres.item(j).description
                            y.Valeur = _ListDrivers.Item(i).Parametres.item(j).valeur
                            retour.Parametres.Add(y)
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).LabelsDriver.count - 1
                            Dim y As New Driver.cLabels
                            y.NomChamp = _ListDrivers.Item(i).LabelsDriver.item(j).NomChamp
                            y.LabelChamp = _ListDrivers.Item(i).LabelsDriver.item(j).LabelChamp
                            y.Tooltip = _ListDrivers.Item(i).LabelsDriver.item(j).Tooltip
                            y.Parametre = _ListDrivers.Item(i).LabelsDriver.item(j).Parametre
                            retour.LabelsDriver.Add(y)
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).LabelsDevice.count - 1
                            Dim y As New Driver.cLabels
                            y.NomChamp = _ListDrivers.Item(i).LabelsDevice.item(j).NomChamp
                            y.LabelChamp = _ListDrivers.Item(i).LabelsDevice.item(j).LabelChamp
                            y.Tooltip = _ListDrivers.Item(i).LabelsDevice.item(j).Tooltip
                            y.Parametre = _ListDrivers.Item(i).LabelsDevice.item(j).Parametre
                            retour.LabelsDevice.Add(y)
                        Next
                        Dim _listactdrv As New ArrayList
                        Dim _listactd As New List(Of String)
                        For j As Integer = 0 To Api.ListMethod(_ListDrivers.Item(i)).Count - 1
                            _listactd.Add(Api.ListMethod(_ListDrivers.Item(i)).Item(j).ToString)
                        Next
                        If _listactd.Count > 0 Then
                            For n As Integer = 0 To _listactd.Count - 1
                                Dim a() As String = _listactd.Item(n).Split("|")
                                Dim p As New DeviceAction
                                With p
                                    .Nom = a(0)
                                    If a.Length > 1 Then
                                        For t As Integer = 1 To a.Length - 1
                                            Dim pr As New DeviceAction.Parametre
                                            Dim b() As String = a(t).Split(":")
                                            With pr
                                                .Nom = b(0)
                                                .Type = b(1)
                                            End With
                                            p.Parametres.Add(pr)
                                        Next
                                    End If
                                End With
                                retour.DeviceAction.Add(p)
                            Next
                        End If

                        _listactd = Nothing
                        _listactdrv = Nothing
                        Exit For
                    End If
                Next
                Return retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDriverById", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne un driver par son ID</summary>
        ''' <param name="DriverId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnDrvById(ByVal IdSrv As String, ByVal DriverId As String) As Object
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim retour As Object = Nothing

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).ID = DriverId Then
                        retour = _ListDrivers.Item(i)
                        Exit For
                    End If
                Next

                Return retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDrvById", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne le driver par son Nom</summary>
        ''' <param name="DriverNom"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnDriverByNom(ByVal IdSrv As String, ByVal DriverNom As String) As Object Implements IHoMIDom.ReturnDriverByNom
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim retour As New TemplateDriver

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).Nom = DriverNom.ToUpper() Then
                        retour.Nom = _ListDrivers.Item(i).nom
                        retour.ID = _ListDrivers.Item(i).id
                        retour.COM = _ListDrivers.Item(i).com
                        retour.Description = _ListDrivers.Item(i).description
                        retour.Enable = _ListDrivers.Item(i).enable
                        retour.IP_TCP = _ListDrivers.Item(i).ip_tcp
                        retour.IP_UDP = _ListDrivers.Item(i).ip_udp
                        retour.IsConnect = _ListDrivers.Item(i).isconnect
                        retour.Modele = _ListDrivers.Item(i).modele
                        retour.Picture = _ListDrivers.Item(i).picture
                        retour.Port_TCP = _ListDrivers.Item(i).port_tcp
                        retour.Port_UDP = _ListDrivers.Item(i).port_udp
                        retour.Protocol = _ListDrivers.Item(i).protocol
                        retour.Refresh = _ListDrivers.Item(i).refresh
                        retour.StartAuto = _ListDrivers.Item(i).startauto
                        retour.Version = _ListDrivers.Item(i).version
                        retour.AutoDiscover = _ListDrivers.Item(i).autoDiscover

                        For j As Integer = 0 To _ListDrivers.Item(i).DeviceSupport.count - 1
                            retour.DeviceSupport.Add(_ListDrivers.Item(i).devicesupport.item(j).ToString)
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).Parametres.count - 1
                            Dim y As New Driver.Parametre
                            y.Nom = _ListDrivers.Item(i).Parametres.item(j).nom
                            y.Description = _ListDrivers.Item(i).Parametres.item(j).description
                            y.Valeur = _ListDrivers.Item(i).Parametres.item(j).valeur
                            retour.Parametres.Add(y)
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).LabelsDriver.count - 1
                            Dim y As New Driver.cLabels
                            y.NomChamp = _ListDrivers.Item(i).LabelsDriver.item(j).NomChamp
                            y.LabelChamp = _ListDrivers.Item(i).LabelsDriver.item(j).LabelChamp
                            y.Tooltip = _ListDrivers.Item(i).LabelsDriver.item(j).Tooltip
                            y.Parametre = _ListDrivers.Item(i).LabelsDriver.item(j).Parametre
                            retour.LabelsDriver.Add(y)
                        Next
                        For j As Integer = 0 To _ListDrivers.Item(i).LabelsDevice.count - 1
                            Dim y As New Driver.cLabels
                            y.NomChamp = _ListDrivers.Item(i).LabelsDevice.item(j).NomChamp
                            y.LabelChamp = _ListDrivers.Item(i).LabelsDevice.item(j).LabelChamp
                            y.Tooltip = _ListDrivers.Item(i).LabelsDevice.item(j).Tooltip
                            y.Parametre = _ListDrivers.Item(i).LabelsDevice.item(j).Parametre
                            retour.LabelsDevice.Add(y)
                        Next
                        Dim _listactdrv As New ArrayList
                        Dim _listactd As New List(Of String)
                        For j As Integer = 0 To Api.ListMethod(_ListDrivers.Item(i)).Count - 1
                            _listactd.Add(Api.ListMethod(_ListDrivers.Item(i)).Item(j).ToString)
                        Next
                        If _listactd.Count > 0 Then
                            For n As Integer = 0 To _listactd.Count - 1
                                Dim a() As String = _listactd.Item(n).Split("|")
                                Dim p As New DeviceAction
                                With p
                                    .Nom = a(0)
                                    If a.Length > 1 Then
                                        For t As Integer = 1 To a.Length - 1
                                            Dim pr As New DeviceAction.Parametre
                                            Dim b() As String = a(t).Split(":")
                                            With pr
                                                .Nom = b(0)
                                                .Type = b(1)
                                            End With
                                            p.Parametres.Add(pr)
                                        Next
                                    End If
                                End With
                                retour.DeviceAction.Add(p)
                            Next
                        End If

                        _listactd = Nothing
                        _listactdrv = Nothing
                        Return retour
                        Exit For

                    End If
                Next
                Return Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDriverByNom", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Permet d'exécuter une commande Sub d'un Driver</summary>
        ''' <param name="DriverId"></param>
        ''' <param name="Action"></param>
        ''' <remarks></remarks>
        Sub ExecuteDriverCommand(ByVal IdSrv As String, ByVal DriverId As String, ByVal Action As DeviceAction) Implements IHoMIDom.ExecuteDriverCommand
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                Dim _retour As Object
                Dim x As Object = Nothing

                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).id = DriverId Then
                        x = _ListDrivers.Item(i)
                        Exit For
                    End If
                Next

                If x IsNot Nothing Then

                    If Action.Parametres.Count > 0 Then
                        Select Case Action.Parametres.Count
                            Case 1
                                _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value)
                            Case 2
                                _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value, Action.Parametres.Item(1).Value)
                            Case 3
                                _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value, Action.Parametres.Item(1).Value, Action.Parametres.Item(2).Value)
                            Case 4
                                _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value, Action.Parametres.Item(1).Value, Action.Parametres.Item(2).Value, Action.Parametres.Item(3).Value)
                            Case 5
                                _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value, Action.Parametres.Item(1).Value, Action.Parametres.Item(2).Value, Action.Parametres.Item(3).Value, Action.Parametres.Item(4).Value)
                        End Select
                    Else
                        CallByName(x, Action.Nom, CallType.Method)
                    End If
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ExecuteDriverCommand", "Erreur lors du traitement de la commande ExecuteDriverCommand: " & ex.Message)
            End Try
        End Sub

#End Region

#Region "Device"

        ''' <summary>
        ''' Permet de changer la valeur d'un device
        ''' </summary>
        ''' <param name="idsrv"></param>
        ''' <param name="IdDevice"></param>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ChangeValueOfDevice(ByVal idsrv As String, ByVal IdDevice As String, ByVal Value As Object) As Integer Implements IHoMIDom.ChangeValueOfDevice
            Try
                If VerifIdSrv(idsrv) = False Then
                    Return 99
                End If

                Dim result As Integer = -1
                Dim _dev As Object = ReturnRealDeviceById(IdDevice)

                If _dev IsNot Nothing Then
                    _dev.value = Value
                    result = 0
                End If

                Return result
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ChangeValueOfDevice", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Permet de changer la valeur d'un device
        ''' </summary>
        ''' <param name="idsrv"></param>
        ''' <param name="IdDevice"></param>
        ''' <param name="Value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ChangeValueOfDeviceSimple(ByVal idsrv As String, ByVal IdDevice As String, ByVal Value As String) As Integer Implements IHoMIDom.ChangeValueOfDeviceSimple
            Try
                If VerifIdSrv(idsrv) = False Then
                    Return 99
                End If

                Dim result As Integer = -1
                Dim _dev As Object = ReturnRealDeviceById(IdDevice)

                If _dev IsNot Nothing Then
                    _dev.value = Value
                    result = 0
                End If

                Return result
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ChangeValueOfDeviceSimple", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Retourne la liste des devices non à jour
        ''' </summary>
        ''' <param name="idsrv"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetDeviceNoMaJ(ByVal idsrv) As List(Of String) Implements IHoMIDom.GetDeviceNoMaJ
            Try
                Return _DevicesNoMAJ
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetDeviceNoMaJ", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Cherche les devices non à jour (lancé toutes les heures)
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub SearchDeviceNoMaJ()
            Try
                _DevicesNoMAJ.Clear()

                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).LastChangeDuree > 0 Then
                        If DateTime.Compare(_ListDevices.Item(i).LastChange.AddMinutes(CInt(_ListDevices.Item(i).LastChangeDuree)), Now) < 0 Then
                            _DevicesNoMAJ.Add(_ListDevices.Item(i).Name)
                        End If
                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SearchDeviceNoMaJ", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Supprimer un device de la config
        ''' </summary>
        ''' <param name="IdSrv"></param>
        ''' <param name="deviceId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DeleteDevice(ByVal IdSrv As String, ByVal deviceId As String) As Integer Implements IHoMIDom.DeleteDevice
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).Id = deviceId Then
                        'on teste si c'est un device systeme pour ne pas le supprimer
                        If Left(_ListDevices.Item(i).Name, 5) = "HOMI_" Then
                            Return -2
                        End If

                        'on arrete le timer en forçant le refresh à 0
                        _ListDevices.Item(i).refresh = 0
                        _ListDevices.Item(i).driver.deletedevice(deviceId)
                        _ListDevices.RemoveAt(i)

                        'va vérifier toutes les zones
                        For j As Integer = 0 To _ListZones.Count - 1
                            DeleteDeviceToZone(IdSrv, _ListZones.Item(j).ID, deviceId)
                        Next
                        'va vérifier tous les triggers
                        For j As Integer = 0 To _ListTriggers.Count - 1
                            DeleteDeviceToTrigger(IdSrv, _ListTriggers.Item(j).ID, deviceId)
                        Next
                        'va vérifier toutes les actions des macros
                        For j As Integer = 0 To _ListMacros.Count - 1
                            DeleteIDToAction(IdSrv, _ListMacros.Item(j).ListActions, deviceId)
                        Next
                        'Supprime l'historique du device dans la bdd
                        DeleteDeviceToBD(deviceId)

                        SaveRealTime()
                        Return 0
                    End If
                Next

                Return -1
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDevice", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Supprime un device dans la bdd de l'historique
        ''' </summary>
        ''' <param name="DeviceId"></param>
        ''' <remarks></remarks>
        Private Sub DeleteDeviceToBD(ByVal DeviceId As String)
            Try
                Dim commande As String
                Dim retourDB As String

                commande = "delete from historiques where device_id='" & DeviceId & "' ;"
                retourDB = sqlite_homidom.nonquery(commande, Nothing)

                If UCase(Mid(retourDB, 1, 3)) = "ERR" Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDeviceToBD", "Erreur: " & retourDB)
                End If

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDeviceToBD", "Erreur : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retoune la position du device dans une des zones sinon "-1"
        ''' </summary>
        ''' <param name="IdDevice"></param>
        ''' <remarks></remarks>
        Private Function DeviceInZone(ByVal IdDevice As Integer) As Integer
            Try
                Dim retour As Integer = -1

                For i As Integer = 0 To _ListZones.Count - 1
                    For j As Integer = 0 To _ListZones.Item(i).ListElement.Count - 1
                        If _ListZones.Item(i).ListElement.Item(j).ElementID = IdDevice Then
                            retour = j
                            Exit For
                        End If
                    Next
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceInZone", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne la liste de tous les devices</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAllDevices(ByVal IdSrv As String) As List(Of TemplateDevice) Implements IHoMIDom.GetAllDevices
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim _list As New List(Of TemplateDevice)

                For i As Integer = 0 To _ListDevices.Count - 1
                    Dim x As New TemplateDevice
                    Dim _listact As New List(Of String)

                    With x
                        .Name = _ListDevices.Item(i).name
                        .ID = _ListDevices.Item(i).id
                        .Enable = _ListDevices.Item(i).enable
                        .LastEtat = _ListDevices.Item(i).LastEtat
                        Select Case UCase(_ListDevices.Item(i).type)
                            Case "APPAREIL" : .Type = Device.ListeDevices.APPAREIL  'modules pour diriger un appareil  ON/OFF
                            Case "AUDIO" : .Type = Device.ListeDevices.AUDIO
                            Case "BAROMETRE" : .Type = Device.ListeDevices.BAROMETRE  'pour stocker les valeur issu d'un barometre meteo ou web
                            Case "BATTERIE" : .Type = Device.ListeDevices.BATTERIE
                            Case "COMPTEUR" : .Type = Device.ListeDevices.COMPTEUR  'compteur DS2423, RFXPower...
                            Case "CONTACT" : .Type = Device.ListeDevices.CONTACT  'detecteur de contact : switch 1-wire
                            Case "DETECTEUR" : .Type = Device.ListeDevices.DETECTEUR  'tous detecteurs : mouvement, obscurite...
                            Case "DIRECTIONVENT" : .Type = Device.ListeDevices.DIRECTIONVENT
                            Case "ENERGIEINSTANTANEE" : .Type = Device.ListeDevices.ENERGIEINSTANTANEE
                            Case "ENERGIETOTALE" : .Type = Device.ListeDevices.ENERGIETOTALE
                            Case "FREEBOX" : .Type = Device.ListeDevices.FREEBOX
                            Case "GENERIQUEBOOLEEN" : .Type = Device.ListeDevices.GENERIQUEBOOLEEN
                            Case "GENERIQUESTRING" : .Type = Device.ListeDevices.GENERIQUESTRING
                            Case "GENERIQUEVALUE" : .Type = Device.ListeDevices.GENERIQUEVALUE
                            Case "HUMIDITE" : .Type = Device.ListeDevices.HUMIDITE
                            Case "LAMPE" : .Type = Device.ListeDevices.LAMPE
                            Case "LAMPERGBW" : .Type = Device.ListeDevices.LAMPERGBW
                            Case "METEO" : .Type = Device.ListeDevices.METEO
                            Case "MULTIMEDIA" : .Type = Device.ListeDevices.MULTIMEDIA
                            Case "PLUIECOURANT" : .Type = Device.ListeDevices.PLUIECOURANT
                            Case "PLUIETOTAL" : .Type = Device.ListeDevices.PLUIETOTAL
                            Case "SWITCH" : .Type = Device.ListeDevices.SWITCH
                            Case "TELECOMMANDE" : .Type = Device.ListeDevices.TELECOMMANDE
                            Case "TEMPERATURE" : .Type = Device.ListeDevices.TEMPERATURE
                            Case "TEMPERATURECONSIGNE" : .Type = Device.ListeDevices.TEMPERATURECONSIGNE
                            Case "UV" : .Type = Device.ListeDevices.UV
                            Case "VITESSEVENT" : .Type = Device.ListeDevices.VITESSEVENT
                            Case "VOLET" : .Type = Device.ListeDevices.VOLET
                        End Select

                        .Description = _ListDevices.Item(i).description
                        .Adresse1 = _ListDevices.Item(i).adresse1
                        .Adresse2 = _ListDevices.Item(i).adresse2
                        .DriverID = _ListDevices.Item(i).driverid
                        .Picture = _ListDevices.Item(i).picture
                        .Solo = _ListDevices.Item(i).solo
                        .Refresh = _ListDevices.Item(i).refresh
                        .Modele = _ListDevices.Item(i).modele
                        .GetDeviceCommandePlus = _ListDevices.Item(i).GetCommandPlus
                        .Value = _ListDevices.Item(i).value
                        .DateCreated = _ListDevices.Item(i).DateCreated
                        .LastChange = _ListDevices.Item(i).LastChange
                        .LastChangeDuree = _ListDevices.Item(i).LastChangeDuree
                        .Unit = _ListDevices.Item(i).Unit
                        .AllValue = _ListDevices.Item(i).AllValue
                        .VariablesOfDevice = _ListDevices.Item(i).Variables
                        .CountHisto = _ListDevices.Item(i).CountHisto
                        .IsHisto = _ListDevices.Item(i).isHisto
                        .RefreshHisto = _ListDevices.Item(i).RefreshHisto
                        .Purge = _ListDevices.Item(i).purge
                        .MoyHeure = _ListDevices.Item(i).moyheure
                        .MoyJour = _ListDevices.Item(i).moyjour

                        If IsNumeric(_ListDevices.Item(i).valuelast) Then .ValueLast = _ListDevices.Item(i).valuelast

                        _listact = ListMethod(_ListDevices.Item(i).id)
                        If _listact.Count > 0 Then
                            For Each n In _listact
                                Dim a() As String = n.Split("|")
                                Dim p As New DeviceAction
                                With p
                                    .Nom = a(0)
                                    If a.Length > 1 Then
                                        For t As Integer = 1 To a.Length - 1
                                            Dim pr As New DeviceAction.Parametre
                                            Dim b() As String = a(t).Split(":")
                                            With pr
                                                .Nom = b(0)
                                                .Type = b(1)
                                            End With
                                            p.Parametres.Add(pr)
                                        Next
                                    End If
                                End With
                                .DeviceAction.Add(p)
                                a = Nothing
                                p = Nothing
                            Next
                        End If
                        _listact = Nothing

                        Dim _flag As Boolean = True
                        Select Case .Type
                            Case Device.ListeDevices.BAROMETRE
                            Case Device.ListeDevices.COMPTEUR
                            Case Device.ListeDevices.ENERGIEINSTANTANEE
                            Case Device.ListeDevices.ENERGIETOTALE
                            Case Device.ListeDevices.GENERIQUEVALUE
                            Case Device.ListeDevices.HUMIDITE
                            Case Device.ListeDevices.LAMPERGBW
                                .red = _ListDevices.Item(i).red
                                .green = _ListDevices.Item(i).green
                                .blue = _ListDevices.Item(i).blue
                                .white = _ListDevices.Item(i).white
                                .temperature = _ListDevices.Item(i).temperature
                                .speed = _ListDevices.Item(i).speed
                                .optionnal = _ListDevices.Item(i).optionnal
                            Case Device.ListeDevices.PLUIECOURANT
                            Case Device.ListeDevices.PLUIETOTAL
                            Case Device.ListeDevices.TEMPERATURE
                            Case Device.ListeDevices.TEMPERATURECONSIGNE
                            Case Device.ListeDevices.VITESSEVENT
                            Case Device.ListeDevices.UV
                            Case Device.ListeDevices.VITESSEVENT
                            Case Device.ListeDevices.METEO
                                .ConditionActuel = _ListDevices.Item(i).ConditionActuel
                                .ConditionJ1 = _ListDevices.Item(i).ConditionJ1
                                .ConditionJ2 = _ListDevices.Item(i).ConditionActuel
                                .ConditionJ3 = _ListDevices.Item(i).ConditionJ3
                                .ConditionToday = _ListDevices.Item(i).ConditionToday
                                .HumiditeActuel = _ListDevices.Item(i).HumiditeActuel
                                .IconActuel = _ListDevices.Item(i).IconActuel
                                .IconJ1 = _ListDevices.Item(i).IconJ1
                                .IconJ2 = _ListDevices.Item(i).IconJ2
                                .IconJ3 = _ListDevices.Item(i).IconJ3
                                .IconToday = _ListDevices.Item(i).IconToday
                                .JourJ1 = _ListDevices.Item(i).JourJ1
                                .JourJ2 = _ListDevices.Item(i).JourJ2
                                .JourJ3 = _ListDevices.Item(i).JourJ3
                                .JourToday = _ListDevices.Item(i).JourToday
                                .MaxJ1 = _ListDevices.Item(i).MaxJ1
                                .MaxJ2 = _ListDevices.Item(i).MaxJ2
                                .MaxJ3 = _ListDevices.Item(i).MaxJ3
                                .MaxToday = _ListDevices.Item(i).MaxToday
                                .MinJ1 = _ListDevices.Item(i).MinJ1
                                .MinJ2 = _ListDevices.Item(i).MinJ2
                                .MinJ3 = _ListDevices.Item(i).MinJ3
                                .MinToday = _ListDevices.Item(i).MinToday
                                .TemperatureActuel = _ListDevices.Item(i).TemperatureActuel
                                .VentActuel = _ListDevices.Item(i).VentActuel
                                _flag = False
                            Case Device.ListeDevices.MULTIMEDIA
                                .Commandes = _ListDevices.Item(i).Commandes
                                _flag = False
                            Case Else
                                _flag = False
                        End Select

                        If _flag Then
                            .Correction = _ListDevices.Item(i).correction
                            .Precision = _ListDevices.Item(i).precision
                            .Formatage = _ListDevices.Item(i).formatage
                            .ValueDef = _ListDevices.Item(i).valuedef
                            .ValueMax = _ListDevices.Item(i).valuemax
                            .ValueMin = _ListDevices.Item(i).valuemin
                        End If
                    End With

                    _list.Add(x)
                    x = Nothing
                Next

                _list.Sort(AddressOf sortDevice)
                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllDevices", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function sortDevice(ByVal x As TemplateDevice, ByVal y As TemplateDevice) As Integer
            Try
                Return x.Name.CompareTo(y.Name)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeviceInZone", "Exception : " & ex.Message)
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' Sauvegarder ou créer un device
        ''' </summary>
        ''' <param name="IdSrv"></param>
        ''' <param name="deviceId"></param>
        ''' <param name="name"></param>
        ''' <param name="address1"></param>
        ''' <param name="enable"></param>
        ''' <param name="solo"></param>
        ''' <param name="driverid"></param>
        ''' <param name="type"></param>
        ''' <param name="refresh"></param>
        ''' <param name="address2"></param>
        ''' <param name="image"></param>
        ''' <param name="modele"></param>
        ''' <param name="description"></param>
        ''' <param name="lastchangeduree"></param>
        ''' <param name="lastEtat"></param>
        ''' <param name="correction"></param>
        ''' <param name="formatage"></param>
        ''' <param name="precision"></param>
        ''' <param name="valuemax"></param>
        ''' <param name="valuemin"></param>
        ''' <param name="valuedef"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SaveDevice(ByVal IdSrv As String, ByVal deviceId As String, ByVal name As String, ByVal address1 As String, ByVal enable As Boolean, ByVal solo As Boolean, ByVal driverid As String, ByVal type As String, ByVal refresh As Integer, ByVal Historisation As Boolean, ByVal RefreshHisto As Double, ByVal purge As Double, ByVal moyjour As Double, ByVal moyheure As Double, Optional ByVal address2 As String = "", Optional ByVal image As String = "", Optional ByVal modele As String = "", Optional ByVal description As String = "", Optional ByVal lastchangeduree As Integer = 0, Optional ByVal lastEtat As Boolean = True, Optional ByVal correction As String = "0", Optional ByVal formatage As String = "", Optional ByVal precision As Double = 0, Optional ByVal valuemax As Double = 9999, Optional ByVal valuemin As Double = -9999, Optional ByVal valuedef As Double = 0, Optional ByVal Commandes As List(Of Telecommande.Commandes) = Nothing, Optional ByVal Unit As String = "", Optional ByVal Puissance As Integer = 0, Optional ByVal AllValue As Boolean = False, Optional ByVal Variables As Dictionary(Of String, String) = Nothing, Optional ByVal Proprietes As Dictionary(Of String, String) = Nothing) As String Implements IHoMIDom.SaveDevice
            Try
                'Vérification de l'Id du serveur pour accepter le traitement
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                'Déclaration des variables
                Dim myID As String = ""

                'Test si c'est un nouveau device
                If String.IsNullOrEmpty(deviceId) = True Then 'C'est un nouveau device

                    For i1 As Integer = 0 To _ListDevices.Count - 1
                        If LCase(_ListDevices.Item(i1).name) = LCase(name) Then
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveDevice", "Le nom du device: " & name & " existe déjà impossible de l'enregister")
                            Return 98
                        End If
                    Next

                    myID = System.Guid.NewGuid.ToString()

                    Dim MyNewObj As Object = Nothing

                    'Suivant chaque type de device
                    Select Case UCase(type)
                        Case "TEMPERATURE"
                            Dim o As New Device.TEMPERATURE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "HUMIDITE"
                            Dim o As New Device.HUMIDITE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "BATTERIE"
                            Dim o As New Device.BATTERIE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "TEMPERATURECONSIGNE"
                            Dim o As New Device.TEMPERATURECONSIGNE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "ENERGIETOTALE"
                            Dim o As New Device.ENERGIETOTALE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "ENERGIEINSTANTANEE"
                            Dim o As New Device.ENERGIEINSTANTANEE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "PLUIETOTAL"
                            Dim o As New Device.PLUIETOTAL(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "PLUIECOURANT"
                            Dim o As New Device.PLUIECOURANT(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "VITESSEVENT"
                            Dim o As New Device.VITESSEVENT(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "DIRECTIONVENT"
                            Dim o As New Device.DIRECTIONVENT(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "UV"
                            Dim o As New Device.UV(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "APPAREIL"
                            Dim o As New Device.APPAREIL(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "LAMPE"
                            Dim o As New Device.LAMPE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "LAMPERGBW"
                            Dim o As New Device.LAMPERGBW(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "CONTACT"
                            Dim o As New Device.CONTACT(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "METEO"
                            Dim o As New Device.METEO(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "AUDIO"
                            Dim o As New Device.AUDIO(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "MULTIMEDIA"
                            Dim o As New Device.MULTIMEDIA(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "FREEBOX"
                            Dim o As New Device.FREEBOX(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "VOLET"
                            Dim o As New Device.VOLET(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "BAROMETRE"
                            Dim o As New Device.BAROMETRE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "COMPTEUR"
                            Dim o As New Device.COMPTEUR(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "DETECTEUR"
                            Dim o As New Device.DETECTEUR(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "GENERIQUEBOOLEEN"
                            Dim o As New Device.GENERIQUEBOOLEEN(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "GENERIQUESTRING"
                            Dim o As New Device.GENERIQUESTRING(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "GENERIQUEVALUE"
                            Dim o As New Device.GENERIQUEVALUE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "SWITCH"
                            Dim o As New Device.SWITCH(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            MyNewObj = o
                        Case "TELECOMMANDE"
                            Dim o As New Device.TELECOMMANDE(Me)
                            AddHandler o.DeviceChanged, AddressOf DeviceChange
                            ' o.Driver.newdevice(deviceId)
                            MyNewObj = o
                    End Select

                    'Propriétés communes
                    With MyNewObj
                        .ID = myID
                        .Name = name
                        .DateCreated = Now
                        .Picture = image
                        .Adresse1 = address1
                        .Adresse2 = address2
                        .Enable = enable
                        .DriverID = driverid
                        .Modele = modele
                        .Refresh = refresh
                        .Solo = solo
                        .Description = description
                        .LastChangeDuree = lastchangeduree
                        .LastEtat = lastEtat
                        .Unit = Unit
                        .Puissance = Puissance
                        .AllValue = AllValue
                        .Variables = Variables
                        .CountHisto = 0
                        .isHisto = Historisation
                        .RefreshHisto = RefreshHisto
                        .Purge = purge
                        .MoyJour = moyjour
                        .MoyHeure = moyheure
                    End With

                    Select Case UCase(type)
                        Case "BAROMETRE", "COMPTEUR", "ENERGIEINSTANTANEE", "ENERGIETOTALE", "GENERIQUEVALUE", "HUMIDITE", "LAMPE", "PLUIECOURANT", "PLUIETOTAL", "TEMPERATURE", "TEMPERATURECONSIGNE", "VITESSEVENT", "UV", "VOLET"
                            With MyNewObj
                                .Correction = correction
                                .Formatage = formatage
                                .Precision = precision
                                .ValueMax = valuemax
                                .ValueMin = valuemin
                                .ValueDef = valuedef
                            End With
                        Case "LAMPERGBW"
                            With MyNewObj
                                .Correction = correction
                                .Formatage = formatage
                                .Precision = precision
                                .ValueMax = valuemax
                                .ValueMin = valuemin
                                .ValueDef = valuedef
                                If Proprietes.ContainsKey("red") Then .red = Proprietes("red") Else .red = 0
                                If Proprietes.ContainsKey("green") Then .green = Proprietes("green") Else .green = 0
                                If Proprietes.ContainsKey("blue") Then .blue = Proprietes("blue") Else .blue = 0
                                If Proprietes.ContainsKey("white") Then .white = Proprietes("white") Else .white = 0
                                If Proprietes.ContainsKey("temperature") Then .temperature = Proprietes("temperature") Else .temperature = 0
                                If Proprietes.ContainsKey("speed") Then .speed = Proprietes("speed") Else .speed = 0
                                If Proprietes.ContainsKey("optionnal") Then .optionnal = Proprietes("optionnal") Else .optionnal = 0
                            End With
                    End Select

                    _ListDevices.Add(MyNewObj)
                    SaveRealTime()

                    'Libére la mémoire de la variable
                    MyNewObj = Nothing

                Else 'Device Existant
                    myID = deviceId
                    For i As Integer = 0 To _ListDevices.Count - 1
                        If _ListDevices.Item(i).ID = deviceId Then

                            'propriétés modifiables pour tout device
                            _ListDevices.Item(i).description = description
                            _ListDevices.Item(i).picture = image
                            _ListDevices.Item(i).Puissance = Puissance
                            _ListDevices.Item(i).Variables = Variables

                            'on teste si c'est un device systeme pour ne pas le modifier
                            If Left(_ListDevices.Item(i).Name, 5) = "HOMI_" Then
                                Return -2
                            End If

                            'sauvegarde des propriétés
                            _ListDevices.Item(i).name = name
                            _ListDevices.Item(i).adresse1 = address1
                            _ListDevices.Item(i).adresse2 = address2
                            _ListDevices.Item(i).enable = enable
                            _ListDevices.Item(i).driverid = driverid
                            _ListDevices.Item(i).modele = modele
                            _ListDevices.Item(i).refresh = refresh
                            _ListDevices.Item(i).solo = solo
                            _ListDevices.Item(i).LastChangeDuree = lastchangeduree
                            _ListDevices.Item(i).LastEtat = lastEtat
                            _ListDevices.Item(i).Driver.newdevice(deviceId)
                            _ListDevices.Item(i).Unit = Unit
                            _ListDevices.Item(i).AllValue = AllValue
                            _ListDevices.Item(i).isHisto = Historisation
                            _ListDevices.Item(i).RefreshHisto = RefreshHisto
                            _ListDevices.Item(i).Purge = purge
                            _ListDevices.Item(i).MoyJour = moyjour
                            _ListDevices.Item(i).MoyHeure = moyheure

                            'si c'est un device de type double ou integer
                            If _ListDevices.Item(i).type = "BAROMETRE" _
                                Or _ListDevices.Item(i).type = "COMPTEUR" _
                                Or _ListDevices.Item(i).type = "ENERGIEINSTANTANEE" _
                                Or _ListDevices.Item(i).type = "ENERGIETOTALE" _
                                Or _ListDevices.Item(i).Type = "GENERIQUEVALUE" _
                                Or _ListDevices.Item(i).Type = "HUMIDITE" _
                                Or _ListDevices.Item(i).Type = "LAMPE" _
                                Or _ListDevices.Item(i).Type = "PLUIECOURANT" _
                                Or _ListDevices.Item(i).Type = "PLUIETOTAL" _
                                Or _ListDevices.Item(i).Type = "TEMPERATURE" _
                                Or _ListDevices.Item(i).Type = "TEMPERATURECONSIGNE" _
                                Or _ListDevices.Item(i).Type = "VITESSEVENT" _
                                Or _ListDevices.Item(i).Type = "UV" _
                                Or _ListDevices.Item(i).Type = "VOLET" Then
                                _ListDevices.Item(i).Correction = correction
                                _ListDevices.Item(i).Formatage = formatage
                                _ListDevices.Item(i).Precision = precision
                                _ListDevices.Item(i).ValueMax = valuemax
                                _ListDevices.Item(i).ValueMin = valuemin
                                _ListDevices.Item(i).ValueDef = valuedef
                            ElseIf _ListDevices.Item(i).Type = "LAMPERGBW" Then
                                _ListDevices.Item(i).Correction = correction
                                _ListDevices.Item(i).Formatage = formatage
                                _ListDevices.Item(i).Precision = precision
                                _ListDevices.Item(i).ValueMax = valuemax
                                _ListDevices.Item(i).ValueMin = valuemin
                                _ListDevices.Item(i).ValueDef = valuedef
                                If Proprietes.ContainsKey("red") Then _ListDevices.Item(i).red = Proprietes("red") Else _ListDevices.Item(i).red = 0
                                If Proprietes.ContainsKey("green") Then _ListDevices.Item(i).green = Proprietes("green") Else _ListDevices.Item(i).green = 0
                                If Proprietes.ContainsKey("blue") Then _ListDevices.Item(i).blue = Proprietes("blue") Else _ListDevices.Item(i).blue = 0
                                If Proprietes.ContainsKey("white") Then _ListDevices.Item(i).white = Proprietes("white") Else _ListDevices.Item(i).white = 0
                                If Proprietes.ContainsKey("temperature") Then _ListDevices.Item(i).temperature = Proprietes("temperature") Else _ListDevices.Item(i).temperature = 0
                                If Proprietes.ContainsKey("speed") Then _ListDevices.Item(i).speed = Proprietes("speed") Else _ListDevices.Item(i).speed = 0
                                If Proprietes.ContainsKey("optionnal") Then _ListDevices.Item(i).optionnal = Proprietes("optionnal") Else _ListDevices.Item(i).optionnal = 0
                            End If
                            SaveRealTime()
                            Exit For 'on a trouvé le device, on arrete donc de le chercher.
                        End If
                    Next
                End If

                ManagerSequences.AddSequences(Sequence.TypeOfSequence.DeviceChange, myID, Nothing, Nothing)
                RaiseEvent DeviceChanged(myID, "")

                Return myID
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveDevice", "Exception : " & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>Supprime une commande IR d'un device</summary>
        ''' <param name="deviceId"></param>
        ''' <param name="CmdName"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DeleteDeviceCommandIR(ByVal IdSrv As String, ByVal deviceId As String, ByVal CmdName As String) As Integer Implements IHoMIDom.DeleteDeviceCommandIR
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).Id = deviceId Then
                        For j As Integer = 0 To _ListDevices.Item(i).ListCommandname.count - 1
                            If _ListDevices.Item(i).ListCommandname(j) = CmdName Then
                                _ListDevices.Item(i).ListCommandname.removeat(j)
                                _ListDevices.Item(i).ListCommanddata.removeat(j)
                                _ListDevices.Item(i).ListCommandrepeat.removeat(j)
                                Return 0
                            End If
                        Next
                    End If
                Next

                Return -1
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDeviceCommandIR", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Ajoute ou modifie une commande IR à un device</summary>
        ''' <param name="deviceId"></param>
        ''' <param name="CmdName"></param>
        ''' <param name="CmdData"></param>
        ''' <param name="CmdRepeat"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SaveDeviceCommandIR(ByVal IdSrv As String, ByVal deviceId As String, ByVal CmdName As String, ByVal CmdData As String, ByVal CmdRepeat As String) As String Implements IHoMIDom.SaveDeviceCommandIR
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim flag As Boolean

                'On vérifie avant que si la commande existe on la modifie
                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).id = deviceId Then
                        For j As Integer = 0 To _ListDevices.Item(i).listcommandName.count - 1
                            If _ListDevices.Item(i).listcommandname(j) = CmdName Then
                                _ListDevices.Item(i).listcommanddata(j) = CmdData
                                _ListDevices.Item(i).listcommandrepeat(j) = CmdRepeat
                                flag = True
                            End If
                        Next
                        'sinon on la crée
                        If flag = False Then
                            _ListDevices.Item(i).listcommandname.add(CmdName)
                            _ListDevices.Item(i).listcommanddata.add(CmdData)
                            _ListDevices.Item(i).listcommandrepeat.add(CmdRepeat)
                        End If
                        SaveRealTime()
                    End If
                Next
                Return 0
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveDeviceCommandIR", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Commencer un apprentissage IR</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function StartIrLearning(ByVal IdSrv As String) As String Implements IHoMIDom.StartIrLearning
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim retour As String = ""
                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).protocol = "IR" Then
                        Dim x As Object = _ListDrivers.Item(i)
                        retour = x.LearnCodeIR()
                        Log(TypeLog.INFO, TypeSource.SERVEUR, "SERVEUR", "Apprentissage IR: " & retour)
                    End If
                Next
                Return retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "StartIrLearning", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

        ''' <summary>Vérifie si un device par son ID existe</summary>
        ''' <param name="DeviceId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ExistDeviceById(ByVal DeviceId As String) As Boolean Implements IHoMIDom.ExistDeviceById
            Try
                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).ID = DeviceId Then
                        Return True
                    End If
                Next

                Return False
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ExistDeviceById", "Exception : " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>Retourne un device par son ID</summary>
        ''' <param name="DeviceId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnDeviceById(ByVal IdSrv As String, ByVal DeviceId As String) As TemplateDevice Implements IHoMIDom.ReturnDeviceByID
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim retour As New TemplateDevice
                Dim _listact As New List(Of String)

                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).ID = DeviceId Then
                        retour.ID = _ListDevices.Item(i).id
                        retour.Name = _ListDevices.Item(i).name
                        retour.Enable = _ListDevices.Item(i).enable
                        retour.GetDeviceCommandePlus = _ListDevices.Item(i).GetCommandPlus
                        retour.VariablesOfDevice = _ListDevices.Item(i).Variables
                        retour.CountHisto = _ListDevices.Item(i).CountHisto

                        Select Case UCase(_ListDevices.Item(i).type)
                            Case "APPAREIL" : retour.Type = Device.ListeDevices.APPAREIL  'modules pour diriger un appareil  ON/OFF
                            Case "AUDIO" : retour.Type = Device.ListeDevices.AUDIO
                            Case "BAROMETRE" : retour.Type = Device.ListeDevices.BAROMETRE  'pour stocker les valeur issu d'un barometre meteo ou web
                            Case "BATTERIE" : retour.Type = Device.ListeDevices.BATTERIE
                            Case "COMPTEUR" : retour.Type = Device.ListeDevices.COMPTEUR  'compteur DS2423, RFXPower...
                            Case "CONTACT" : retour.Type = Device.ListeDevices.CONTACT  'detecteur de contact : switch 1-wire
                            Case "DETECTEUR" : retour.Type = Device.ListeDevices.DETECTEUR  'tous detecteurs : mouvement, obscurite...
                            Case "DIRECTIONVENT" : retour.Type = Device.ListeDevices.DIRECTIONVENT
                            Case "ENERGIEINSTANTANEE" : retour.Type = Device.ListeDevices.ENERGIEINSTANTANEE
                            Case "ENERGIETOTALE" : retour.Type = Device.ListeDevices.ENERGIETOTALE
                            Case "FREEBOX" : retour.Type = Device.ListeDevices.FREEBOX
                            Case "GENERIQUEBOOLEEN" : retour.Type = Device.ListeDevices.GENERIQUEBOOLEEN
                            Case "GENERIQUESTRING" : retour.Type = Device.ListeDevices.GENERIQUESTRING
                            Case "GENERIQUEVALUE" : retour.Type = Device.ListeDevices.GENERIQUEVALUE
                            Case "HUMIDITE" : retour.Type = Device.ListeDevices.HUMIDITE
                            Case "LAMPE" : retour.Type = Device.ListeDevices.LAMPE
                            Case "LAMPERGBW" : retour.Type = Device.ListeDevices.LAMPERGBW
                            Case "METEO" : retour.Type = Device.ListeDevices.METEO
                            Case "MULTIMEDIA" : retour.Type = Device.ListeDevices.MULTIMEDIA
                            Case "PLUIECOURANT" : retour.Type = Device.ListeDevices.PLUIECOURANT
                            Case "PLUIETOTAL" : retour.Type = Device.ListeDevices.PLUIETOTAL
                            Case "SWITCH" : retour.Type = Device.ListeDevices.SWITCH
                            Case "TELECOMMANDE" : retour.Type = Device.ListeDevices.TELECOMMANDE
                            Case "TEMPERATURE" : retour.Type = Device.ListeDevices.TEMPERATURE
                            Case "TEMPERATURECONSIGNE" : retour.Type = Device.ListeDevices.TEMPERATURECONSIGNE
                            Case "UV" : retour.Type = Device.ListeDevices.UV
                            Case "VITESSEVENT" : retour.Type = Device.ListeDevices.VITESSEVENT
                            Case "VOLET" : retour.Type = Device.ListeDevices.VOLET
                        End Select

                        retour.Description = _ListDevices.Item(i).description
                        retour.Adresse1 = _ListDevices.Item(i).adresse1
                        retour.Adresse2 = _ListDevices.Item(i).adresse2
                        retour.DriverID = _ListDevices.Item(i).driverid
                        retour.Picture = _ListDevices.Item(i).picture
                        retour.Solo = _ListDevices.Item(i).solo
                        retour.Refresh = _ListDevices.Item(i).refresh
                        retour.Modele = _ListDevices.Item(i).modele
                        retour.LastEtat = _ListDevices.Item(i).LastEtat
                        retour.DateCreated = _ListDevices.Item(i).DateCreated
                        retour.LastChange = _ListDevices.Item(i).LastChange
                        retour.LastChangeDuree = _ListDevices.Item(i).LastChangeDuree
                        retour.Unit = _ListDevices.Item(i).Unit
                        retour.Puissance = _ListDevices.Item(i).Puissance
                        retour.AllValue = _ListDevices.Item(i).AllValue
                        retour.IsHisto = _ListDevices.Item(i).isHisto
                        retour.RefreshHisto = _ListDevices.Item(i).RefreshHisto
                        retour.Purge = _ListDevices.Item(i).purge
                        retour.MoyJour = _ListDevices.Item(i).moyjour
                        retour.MoyHeure = _ListDevices.Item(i).moyheure

                        Try
                            retour.Value = _ListDevices.Item(i).Value
                        Catch ex As Exception

                        End Try

                        _listact = ListMethod(_ListDevices.Item(i).id)

                        If _listact.Count > 0 Then
                            For Each n In _listact
                                Dim a() As String = n.Split("|")
                                Dim p As New DeviceAction
                                With p
                                    .Nom = a(0)
                                    If a.Length > 1 Then
                                        For t As Integer = 1 To a.Length - 1
                                            Dim pr As New DeviceAction.Parametre
                                            Dim b() As String = a(t).Split(":")
                                            With pr
                                                .Nom = b(0)
                                                .Type = b(1)
                                            End With
                                            p.Parametres.Add(pr)
                                        Next
                                    End If
                                End With
                                retour.DeviceAction.Add(p)
                                p = Nothing
                                a = Nothing
                            Next
                        End If

                        Dim _retour1 As Boolean = True
                        Select Case retour.Type
                            Case Device.ListeDevices.METEO
                                _retour1 = False
                                retour.ConditionActuel = _ListDevices.Item(i).ConditionActuel
                                retour.ConditionJ1 = _ListDevices.Item(i).ConditionJ1
                                retour.ConditionJ2 = _ListDevices.Item(i).ConditionActuel
                                retour.ConditionJ3 = _ListDevices.Item(i).ConditionJ3
                                retour.ConditionToday = _ListDevices.Item(i).ConditionToday
                                retour.HumiditeActuel = _ListDevices.Item(i).HumiditeActuel
                                retour.IconActuel = _ListDevices.Item(i).IconActuel
                                retour.IconJ1 = _ListDevices.Item(i).IconJ1
                                retour.IconJ2 = _ListDevices.Item(i).IconJ2
                                retour.IconJ3 = _ListDevices.Item(i).IconJ3
                                retour.IconToday = _ListDevices.Item(i).IconToday
                                retour.JourJ1 = _ListDevices.Item(i).JourJ1
                                retour.JourJ2 = _ListDevices.Item(i).JourJ2
                                retour.JourJ3 = _ListDevices.Item(i).JourJ3
                                retour.JourToday = _ListDevices.Item(i).JourToday
                                retour.MaxJ1 = _ListDevices.Item(i).MaxJ1
                                retour.MaxJ2 = _ListDevices.Item(i).MaxJ2
                                retour.MaxJ3 = _ListDevices.Item(i).MaxJ3
                                retour.MaxToday = _ListDevices.Item(i).MaxToday
                                retour.MinJ1 = _ListDevices.Item(i).MinJ1
                                retour.MinJ2 = _ListDevices.Item(i).MinJ2
                                retour.MinJ3 = _ListDevices.Item(i).MinJ3
                                retour.MinToday = _ListDevices.Item(i).MinToday
                                retour.TemperatureActuel = _ListDevices.Item(i).TemperatureActuel
                                retour.VentActuel = _ListDevices.Item(i).VentActuel
                                _retour1 = False
                            Case Device.ListeDevices.MULTIMEDIA
                                retour.Commandes = _ListDevices.Item(i).Commandes
                                _retour1 = False
                            Case Device.ListeDevices.BAROMETRE
                            Case Device.ListeDevices.COMPTEUR
                            Case Device.ListeDevices.ENERGIEINSTANTANEE
                            Case Device.ListeDevices.ENERGIETOTALE
                            Case Device.ListeDevices.GENERIQUEVALUE
                            Case Device.ListeDevices.HUMIDITE
                            Case Device.ListeDevices.LAMPE
                            Case Device.ListeDevices.LAMPERGBW
                                retour.red = _ListDevices.Item(i).red
                                retour.green = _ListDevices.Item(i).green
                                retour.blue = _ListDevices.Item(i).blue
                                retour.white = _ListDevices.Item(i).white
                                retour.temperature = _ListDevices.Item(i).temperature
                                retour.speed = _ListDevices.Item(i).speed
                                retour.optionnal = _ListDevices.Item(i).optionnal
                            Case Device.ListeDevices.PLUIECOURANT
                            Case Device.ListeDevices.PLUIETOTAL
                            Case Device.ListeDevices.TEMPERATURE
                            Case Device.ListeDevices.TEMPERATURECONSIGNE
                            Case Device.ListeDevices.VITESSEVENT
                            Case Device.ListeDevices.UV
                            Case Device.ListeDevices.VOLET
                            Case Else
                                _retour1 = False
                        End Select

                        If _retour1 Then
                            retour.Correction = _ListDevices.Item(i).correction
                            retour.Precision = _ListDevices.Item(i).precision
                            retour.Formatage = _ListDevices.Item(i).formatage
                            retour.Value = _ListDevices.Item(i).value
                            retour.ValueDef = _ListDevices.Item(i).valuedef
                            retour.ValueLast = _ListDevices.Item(i).valuelast
                            retour.ValueMax = _ListDevices.Item(i).valuemax
                            retour.ValueMin = _ListDevices.Item(i).valuemin
                        End If

                        Exit For
                    End If
                Next

                If String.IsNullOrEmpty(retour.ID) = False Then
                    Return retour
                Else
                    Return Nothing
                End If

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDeviceById", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne un device par son ID</summary>
        ''' <param name="DeviceId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnRealDeviceById(ByVal DeviceId As String) As Object
            Try
                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).ID = DeviceId Then
                        Return _ListDevices.Item(i)
                    End If
                Next

                'Si pas trouvé on retourne rien
                Return Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnRealDeviceById", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retourne le type d'une propriété d'un device (boolean, string, double...)
        ''' </summary>
        ''' <param name="DeviceId">ID du device</param>
        ''' <param name="Property">Nom de la propriété</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function TypeOfPropertyOfDevice(ByVal DeviceId As String, ByVal [Property] As String) As String Implements IHoMIDom.TypeOfPropertyOfDevice
            Try
                Dim _dev As Object = ReturnRealDeviceById(DeviceId)
                Dim _result As String = ""

                If _dev IsNot Nothing Then
                    _result = TypeOfProperty(_dev, [Property])
                End If

                Return _result
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "TypeOfPropertyOfDevice", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

        ''' <summary>Retourne un device par son nom</summary>
        ''' <param name="Name"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnRealDeviceByName(ByVal Name As String) As Object
            Try
                For i As Integer = 0 To _ListDevices.Count - 1
                    If UCase(_ListDevices.Item(i).Name) = UCase(Name) Then
                        Return _ListDevices.Item(i)
                    End If
                Next

                'Si pas trouvé on retourne rien
                Return Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnRealDeviceByName", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>liste les méthodes d'un device depuis son ID</summary>
        ''' <param name="DeviceId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function ListMethod(ByVal DeviceId As String) As List(Of String) Implements IHoMIDom.ListMethod
            Try
                Dim _list As New List(Of String)

                'recherche du device
                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).ID = DeviceId Then
                        'device trouvé : _ListDevices.Item(i)

                        'on récupéres les méthodes classiques
                        Dim _listmethod As ArrayList
                        _listmethod = Api.ListMethod(_ListDevices.Item(i))
                        For j As Integer = 0 To _listmethod.Count - 1
                            _list.Add(_listmethod.Item(j).ToString)
                        Next

                        Exit For
                    End If
                Next

                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ListMethod", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne une liste de device par son driver</summary>
        ''' <param name="DriverID"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnDeviceByDriver(ByVal IdSrv As String, ByVal DriverID As String, ByVal Enable As Boolean) As ArrayList Implements IHoMIDom.ReturnDeviceByDriver
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim listresultat As New ArrayList

                For i As Integer = 0 To _ListDevices.Count - 1
                    If (_ListDevices.Item(i).DriverID = DriverID.ToUpper()) And _ListDevices.Item(i).Enable = Enable Then
                        listresultat.Add(_ListDevices.Item(i))
                    End If
                Next

                Return listresultat
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDeviceByDriver", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne une liste de device par son Adresse1 et/ou type et/ou son driver, ex: "A1" "TEMPERATURE" "RFXCOM_RECEIVER"</summary>
        ''' <param name="DeviceAdresse"></param>
        ''' <param name="DeviceType"></param>
        ''' <param name="DriverID"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnDeviceByAdresse1TypeDriver(ByVal IdSrv As String, ByVal DeviceAdresse As String, ByVal DeviceType As String, ByVal DriverID As String, ByVal Enable As Boolean) As ArrayList Implements IHoMIDom.ReturnDeviceByAdresse1TypeDriver
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim listresultat As New ArrayList

                For i As Integer = 0 To _ListDevices.Count - 1
                    If (String.IsNullOrEmpty(DeviceAdresse) = True Or _ListDevices.Item(i).Adresse1.ToUpper() = DeviceAdresse.ToUpper()) And (DeviceType = "" Or _ListDevices.Item(i).type = DeviceType.ToUpper()) And (DriverID = "" Or _ListDevices.Item(i).DriverID = DriverID.ToUpper()) And _ListDevices.Item(i).Enable = Enable Then
                        listresultat.Add(_ListDevices.Item(i))
                    End If
                Next

                Return listresultat
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDeviceByAdresse1TypeDriver", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne une liste de device par son Adresse2 et/ou type et/ou son driver, ex: "A1" "TEMPERATURE" "RFXCOM_RECEIVER"</summary>
        ''' <param name="DeviceAdresse"></param>
        ''' <param name="DeviceType"></param>
        ''' <param name="DriverID"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnDeviceByAdresse2TypeDriver(ByVal IdSrv As String, ByVal DeviceAdresse As String, ByVal DeviceType As String, ByVal DriverID As String, ByVal Enable As Boolean) As ArrayList Implements IHoMIDom.ReturnDeviceByAdresse2TypeDriver
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim listresultat As New ArrayList

                For i As Integer = 0 To _ListDevices.Count - 1
                    If (String.IsNullOrEmpty(DeviceAdresse) = True Or _ListDevices.Item(i).Adresse2.ToUpper() = DeviceAdresse.ToUpper()) And (DeviceType = "" Or _ListDevices.Item(i).type = DeviceType.ToUpper()) And (DriverID = "" Or _ListDevices.Item(i).DriverID = DriverID.ToUpper()) And _ListDevices.Item(i).Enable = Enable Then
                        listresultat.Add(_ListDevices.Item(i))
                    End If
                Next

                Return listresultat
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDeviceByAdresse2TypeDriver", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne une liste de device par sa Zone et/ou type et/ou son driver, ex: "Salle" "Volet" </summary>
        ''' <param name="ZoneID"></param>
        ''' <param name="DeviceType"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnDeviceByZoneType(ByVal IdSrv As String, ByVal ZoneID As String, ByVal DeviceType As String, ByVal Enable As Boolean) As ArrayList Implements IHoMIDom.ReturnDeviceByZoneType
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim listresultat As New ArrayList

                For Each dev In _ListDevices
                    If ((String.IsNullOrEmpty(ZoneID)) Or dev.ID = ZoneID) And _
                        ((String.IsNullOrEmpty(DeviceType)) Or dev.type = DeviceType.ToUpper()) And _
                        dev.Enable = Enable Then
                        listresultat.Add(dev)
                    End If
                Next

                Return listresultat
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnDeviceByZoneType", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Permet d'exécuter une commande Sub d'un Device</summary>
        ''' <param name="DeviceId"></param>
        ''' <param name="Action"></param>
        ''' <remarks></remarks>
        Sub ExecuteDeviceCommand(ByVal IdSrv As String, ByVal DeviceId As String, ByVal Action As DeviceAction) Implements IHoMIDom.ExecuteDeviceCommand
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                Dim _retour As Object
                Dim x As Object = Nothing
                Dim verifaction As Boolean = False

                x = ReturnDeviceById(IdSrv, DeviceId)

                If x IsNot Nothing Then
                    'On vérifie si l'action existe avant de la lancer
                    Try
                        'on verifie les methodes classiques
                        For i As Integer = 0 To x.DeviceAction.Count - 1
                            If (x.DeviceAction.Item(i).Nom.ToString.StartsWith(Action.Nom, StringComparison.CurrentCultureIgnoreCase)) Then
                                verifaction = True
                            End If
                        Next
                        'on verifie les methodes avancées (du driver)
                        For i As Integer = 0 To x.GetDeviceCommandePlus.Count - 1
                            If (x.GetDeviceCommandePlus.Item(i).NameCommand.ToString.StartsWith(Action.Nom, StringComparison.CurrentCultureIgnoreCase)) Then
                                verifaction = True
                            End If
                        Next
                        'si c'est ExecuteCommand on laisse passer
                        If Action.Nom = "ExecuteCommand" Then verifaction = True

                        If verifaction = False Then
                            Log(Server.TypeLog.ERREUR, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "ExecuteDeviceCommand non effectué car la commande " & Action.Nom & " n'existe pas pour le composant : " & x.Name)
                            Exit Sub
                        End If
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ExecuteDeviceCommand", "Verification Action Exception : " & ex.Message)
                    End Try

                    'on lance l'action
                    For i As Integer = 0 To _ListDevices.Count - 1
                        If _ListDevices.Item(i).ID = DeviceId Then
                            'device trouvé : _ListDevices.Item(i)
                            x = _ListDevices.Item(i)
                            Exit For
                        End If
                    Next

                    If Action.Parametres IsNot Nothing Then
                        Log(Server.TypeLog.DEBUG, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "parametres count: " & Action.Parametres.Count)

                        If Action.Parametres.Count > 0 Then
                            Select Case Action.Parametres.Count
                                Case 1
                                    _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value)
                                    Log(Server.TypeLog.INFO, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "effectué: " & x.Name & " Command: " & Action.Nom & " Parametre: " & Action.Parametres.Item(0).Value)
                                Case 2
                                    _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value, Action.Parametres.Item(1).Value)
                                    Log(Server.TypeLog.INFO, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "effectué: " & x.Name & " Command: " & Action.Nom & " Parametre0: " & Action.Parametres.Item(0).Value & " Parametre1: " & Action.Parametres.Item(1).Value)
                                Case 3
                                    _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value, Action.Parametres.Item(1).Value, Action.Parametres.Item(2).Value)
                                    Log(Server.TypeLog.INFO, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "effectué: " & x.Name & " Command: " & Action.Nom & " Parametre0: " & Action.Parametres.Item(0).Value & " Parametre1: " & Action.Parametres.Item(1).Value & " Parametre2: " & Action.Parametres.Item(2).Value)
                                Case 4
                                    _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value, Action.Parametres.Item(1).Value, Action.Parametres.Item(2).Value, Action.Parametres.Item(3).Value)
                                Case 5
                                    _retour = CallByName(x, Action.Nom, CallType.Method, Action.Parametres.Item(0).Value, Action.Parametres.Item(1).Value, Action.Parametres.Item(2).Value, Action.Parametres.Item(3).Value, Action.Parametres.Item(4).Value)
                            End Select
                        Else
                            CallByName(x, Action.Nom, CallType.Method)
                            Log(Server.TypeLog.INFO, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "effectué: " & x.Name & " Command: " & Action.Nom)
                        End If
                    Else
                        CallByName(x, Action.Nom, CallType.Method)
                        Log(Server.TypeLog.INFO, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "effectué: " & x.Name & " Command: " & Action.Nom & " Aucun paramètre")
                    End If
                Else
                    Log(Server.TypeLog.ERREUR, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "non effectué car le composant n'a pas été trouvé : " & DeviceId)
                End If
            Catch ex As Exception
                Log(Server.TypeLog.ERREUR, Server.TypeSource.SERVEUR, "ExecuteDevicecommand", "Erreur lors du traitement : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Permet d'exécuter une commande Sub d'un Device</summary>
        ''' <param name="DeviceId"></param>
        ''' <param name="Action"></param>
        ''' <remarks></remarks>
        Sub ExecuteDeviceCommandSimple(ByVal IdSrv As String, ByVal DeviceId As String, ByVal Action As DeviceActionSimple) Implements IHoMIDom.ExecuteDeviceCommandSimple
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                Dim _retour As Object = Nothing
                Dim x As Object = Nothing
                Dim verifaction As Boolean = False

                x = ReturnDeviceById(IdSrv, DeviceId)

                If x IsNot Nothing Then

                    'On vérifie si l'action existe avant de la lancer
                    Try
                        'on verifie les methodes classiques
                        For i As Integer = 0 To x.DeviceAction.Count - 1
                            If (x.DeviceAction.Item(i).Nom.ToString.StartsWith(Action.Nom, StringComparison.CurrentCultureIgnoreCase)) Then
                                verifaction = True
                            End If
                        Next
                        'on verifie les methodes avancées (du driver)
                        For i As Integer = 0 To x.GetDeviceCommandePlus.Count - 1
                            If (x.GetDeviceCommandePlus.Item(i).NameCommand.ToString.StartsWith(Action.Nom, StringComparison.CurrentCultureIgnoreCase)) Then
                                verifaction = True
                            End If
                        Next
                        If verifaction = False Then
                            Log(Server.TypeLog.ERREUR, Server.TypeSource.SERVEUR, "ExecuteDeviceCommandSimple", "Non effectué car la commande " & Action.Nom & " n'existe pas pour le composant : " & x.Name)
                            Exit Sub
                        End If
                    Catch ex As Exception
                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ExecuteDeviceCommandSimple", "Verification Action Exception : " & ex.Message)
                    End Try

                    'on lance l'action
                    For i As Integer = 0 To _ListDevices.Count - 1
                        If _ListDevices.Item(i).ID = DeviceId Then
                            'device trouvé : _ListDevices.Item(i)
                            x = _ListDevices.Item(i)
                            Exit For
                        End If
                    Next
                    If String.IsNullOrEmpty(Action.Param2) = False Then
                        _retour = CallByName(x, Action.Nom, CallType.Method, Action.Param1, Action.Param2)
                        Log(Server.TypeLog.INFO, Server.TypeSource.SERVEUR, "ExecuteDeviceCommandSimple", "Effectué: " & x.Name & " Command: " & Action.Nom & " Parametre1/2: " & Action.Param1 & "/" & Action.Param2)
                    ElseIf String.IsNullOrEmpty(Action.Param1) = False Then
                        _retour = CallByName(x, Action.Nom, CallType.Method, Action.Param1)
                        Log(Server.TypeLog.INFO, Server.TypeSource.SERVEUR, "ExecuteDeviceCommandSimple", "Effectué: " & x.Name & " Command: " & Action.Nom & " Parametre1: " & Action.Param1)
                    Else
                        _retour = CallByName(x, Action.Nom, CallType.Method)
                        Log(Server.TypeLog.INFO, Server.TypeSource.SERVEUR, "ExecuteDeviceCommandSimple", "Effectué: " & x.Name & " Command: " & Action.Nom & " sans parametre")
                    End If
                Else
                    Log(Server.TypeLog.ERREUR, Server.TypeSource.SERVEUR, "ExecuteDeviceCommandSimple", "ExecuteDeviceCommandSimple non effectué car le composant n'a pas été trouvé : " & DeviceId)
                End If
            Catch ex As Exception
                Log(Server.TypeLog.ERREUR, Server.TypeSource.SERVEUR, "ExecuteDeviceCommandSimple", "Erreur lors du traitement : " & ex.Message)
            End Try
        End Sub

#End Region

#Region "Zone"
        ''' <summary>Supprimer une zone de la config</summary>
        ''' <param name="zoneId"></param>
        Public Function DeleteZone(ByVal IdSrv As String, ByVal zoneId As String) As Integer Implements IHoMIDom.DeleteZone
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                For i As Integer = 0 To _ListZones.Count - 1
                    If _ListZones.Item(i).ID = zoneId Then
                        _ListZones.RemoveAt(i)
                        SaveRealTime()
                        ManagerSequences.AddSequences(Sequence.TypeOfSequence.ZoneDelete, zoneId, Nothing, Nothing)
                        Return 0
                    End If
                Next
                Return -1
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteZone", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Retourne la liste de toutes les zones</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetAllZones(ByVal IdSrv As String) As List(Of Zone) Implements IHoMIDom.GetAllZones
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim _list As New List(Of Zone)

                For i As Integer = 0 To _ListZones.Count - 1
                    Dim x As New Zone
                    With x
                        .Name = _ListZones.Item(i).Name
                        .ID = _ListZones.Item(i).ID
                        .Icon = _ListZones.Item(i).Icon
                        .Image = _ListZones.Item(i).Image
                        For j As Integer = 0 To _ListZones.Item(i).ListElement.Count - 1
                            .ListElement.Add(_ListZones.Item(i).ListElement.Item(j))
                        Next
                    End With
                    _list.Add(x)
                    x = Nothing
                Next

                _list.Sort(AddressOf sortZone)
                Return _list

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllZones", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function sortZone(ByVal x As Zone, ByVal y As Zone) As Integer
            Try
                Return x.Name.CompareTo(y.Name)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "sortZone", "Exception : " & ex.Message)
                Return 0
            End Try
        End Function

        ''' <summary>ajouter un device à une zone</summary>
        ''' <param name="ZoneId"></param>
        ''' <param name="DeviceId"></param>
        ''' <param name="Visible"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function AddDeviceToZone(ByVal IdSrv As String, ByVal ZoneId As String, ByVal DeviceId As String, ByVal Visible As Boolean) As String Implements IHoMIDom.AddDeviceToZone
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim _zone As Zone = ReturnZoneById(_IdSrv, ZoneId)
                Dim _retour As String = -1

                If _zone IsNot Nothing Then
                    For i As Integer = 0 To _zone.ListElement.Count - 1
                        If _zone.ListElement.Item(i).ElementID = DeviceId Then
                            _zone.ListElement.Item(i).Visible = Visible
                            Return 0
                        End If
                    Next

                    Dim _dev As New Zone.Element_Zone(DeviceId, Visible)
                    _zone.ListElement.Add(_dev)
                    _retour = 0
                End If

                ManagerSequences.AddSequences(Sequence.TypeOfSequence.ZoneChange, ZoneId, Nothing, Nothing)

                Return _retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "AddDeviceToZone", "Exception : " & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>supprimer un device à une zone</summary>
        ''' <param name="ZoneId"></param>
        ''' <param name="DeviceId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function DeleteDeviceToZone(ByVal IdSrv As String, ByVal ZoneId As String, ByVal DeviceId As String) As String Implements IHoMIDom.DeleteDeviceToZone
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim _zone As Zone = ReturnZoneById(_IdSrv, ZoneId)
                Dim _retour As String = -1
                If _zone IsNot Nothing Then
                    For i As Integer = 0 To _zone.ListElement.Count - 1
                        If _zone.ListElement.Item(i).ElementID = DeviceId Then
                            _zone.ListElement.RemoveAt(i)
                            Exit For
                        End If
                    Next
                    _retour = 0
                End If

                ManagerSequences.AddSequences(Sequence.TypeOfSequence.ZoneChange, ZoneId, Nothing, Nothing)

                Return _retour

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDeviceToZone", "Exception : " & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>sauvegarde ou créer une zone dans la config</summary>
        ''' <param name="zoneId"></param>
        ''' <param name="name"></param>
        ''' <param name="ListElement"></param>
        ''' <param name="icon"></param>
        ''' <param name="image"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function SaveZone(ByVal IdSrv As String, ByVal zoneId As String, ByVal name As String, Optional ByVal ListElement As List(Of Zone.Element_Zone) = Nothing, Optional ByVal icon As String = "", Optional ByVal image As String = "") As String Implements IHoMIDom.SaveZone
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim myID As String = ""

                If String.IsNullOrEmpty(icon) = True Then icon = _MonRepertoire & "\images\icones\Zone_128.png"
                If String.IsNullOrEmpty(image) = True Then image = _MonRepertoire & "\images\icones\Zone_Image.png"
                If String.IsNullOrEmpty(zoneId) = True Then
                    Dim x As New Zone
                    With x
                        x.ID = System.Guid.NewGuid.ToString()
                        x.Name = name
                        x.Icon = icon
                        x.Image = image
                        x.ListElement = ListElement
                    End With
                    myID = x.ID
                    _ListZones.Add(x)
                    SaveRealTime()
                    ManagerSequences.AddSequences(Sequence.TypeOfSequence.ZoneAdd, myID, Nothing, Nothing)
                Else
                    'zone Existante
                    myID = zoneId
                    For i As Integer = 0 To _ListZones.Count - 1
                        If _ListZones.Item(i).ID = zoneId Then
                            _ListZones.Item(i).Name = name
                            _ListZones.Item(i).Icon = icon
                            _ListZones.Item(i).Image = image
                            _ListZones.Item(i).ListElement = ListElement
                            SaveRealTime()
                            ManagerSequences.AddSequences(Sequence.TypeOfSequence.ZoneChange, myID, Nothing, Nothing)
                        End If
                    Next
                End If

                RaiseEvent ZoneChanged(myID)

                Return myID
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveZone", "Exception : " & ex.Message)
                Return ""
            End Try
        End Function

        ''' <summary>Retourne la liste des devices d'une zone depuis son ID</summary>
        ''' <param name="ZoneId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetDeviceInZone(ByVal IdSrv As String, ByVal zoneId As String) As List(Of String) Implements IHoMIDom.GetDeviceInZone
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim x As Zone = ReturnZoneById(_IdSrv, zoneId)
                Dim y As New List(Of String) '(Of TemplateDevice)

                If x IsNot Nothing Then
                    If x.ListElement.Count > 0 Then
                        For i As Integer = 0 To x.ListElement.Count - 1
                            'renvoie l'objet device
                            'Dim z As TemplateDevice = ReturnDeviceById(IdSrv, x.ListElement.Item(i).ElementID)
                            'If z IsNot Nothing Then y.Add(z)

                            'renvoie uniquement l'ID
                            'on verifie si c'est un device (ou une zone ou une macro)
                            For j As Integer = 0 To _ListDevices.Count - 1
                                If _ListDevices.Item(j).ID = x.ListElement.Item(i).ElementID Then y.Add(x.ListElement.Item(i).ElementID)
                            Next
                        Next
                    End If
                End If
                Return y

                y = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetDeviceInZone", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne la liste des zones d'une zone depuis son ID</summary>
        ''' <param name="ZoneId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetZoneInZone(ByVal IdSrv As String, ByVal zoneId As String) As List(Of String) Implements IHoMIDom.GetZoneInZone
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim x As Zone = ReturnZoneById(_IdSrv, zoneId)
                Dim y As New List(Of String) 'New List(Of Zone)
                If x IsNot Nothing Then
                    If x.ListElement.Count > 0 Then
                        For i As Integer = 0 To x.ListElement.Count - 1
                            'Dim z As Zone = ReturnZoneById(IdSrv, x.ListElement.Item(i).ElementID)
                            'If z IsNot Nothing Then y.Add(z)


                            'renvoie uniquement l'ID
                            'on verifie si c'est une zone
                            For j As Integer = 0 To _ListZones.Count - 1
                                If _ListZones.Item(j).ID = x.ListElement.Item(i).ElementID Then y.Add(x.ListElement.Item(i).ElementID)
                            Next


                        Next
                    End If
                End If

                Return y
                y = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetZoneInZone", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Retourne la liste des macros d'une zone depuis son ID</summary>
        ''' <param name="ZoneId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetMacroInZone(ByVal IdSrv As String, ByVal zoneId As String) As List(Of String) Implements IHoMIDom.GetMacroInZone
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim x As Zone = ReturnZoneById(_IdSrv, zoneId)
                Dim y As New List(Of String) 'New List(Of Macro)

                If x IsNot Nothing Then
                    If x.ListElement.Count > 0 Then
                        For i As Integer = 0 To x.ListElement.Count - 1
                            'Dim z As Macro = ReturnMacroById(IdSrv, x.ListElement.Item(i).ElementID)
                            'If z IsNot Nothing Then y.Add(z)

                            'renvoie uniquement l'ID
                            'on verifie si c'est une zone
                            For j As Integer = 0 To _ListMacros.Count - 1
                                If _ListMacros.Item(j).ID = x.ListElement.Item(i).ElementID Then y.Add(x.ListElement.Item(i).ElementID)
                            Next

                        Next
                    End If
                End If

                Return y
                y = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetMacroInZone", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Indique si la zone ne contient aucun device
        ''' </summary>
        ''' <param name="zoneId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ZoneIsEmpty(ByVal IdSrv As String, ByVal zoneId As String) As Boolean Implements IHoMIDom.ZoneIsEmpty
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return False
                End If

                Dim retour As Boolean = True
                Dim x As Zone = ReturnZoneById(_IdSrv, zoneId)

                If x IsNot Nothing Then
                    If x.ListElement.Count > 0 Then
                        retour = False
                    End If
                End If

                Return retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ZoneIsEmpty", "Exception : " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>Retourne la zone par son ID</summary>
        ''' <param name="ZoneId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnZoneById(ByVal IdSrv As String, ByVal ZoneId As String) As Zone Implements IHoMIDom.ReturnZoneByID
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                If (From Zone In _ListZones Where Zone.ID = ZoneId Select Zone).Count > 0 Then
                    Dim Resultat = (From Zone In _ListZones Where Zone.ID = ZoneId Select Zone).First
                    Return Resultat
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnZoneById", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function


#End Region

#Region "Macro"
        ''' <summary>supprimer un id dans une liste d'actions</summary>
        ''' <param name="Actions"></param>
        ''' <param name="Id"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function DeleteIDToAction(ByVal IdSrv As String, ByVal Actions As ArrayList, ByVal Id As String) As String
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteIDToAction", "Idsrv incorrect, Action will not be deleted in actions list : " & Id)
                    Return 99
                End If

                Dim _retour As String = -1
                For j As Integer = 0 To Actions.Count - 1
                    Select Case Actions.Item(j).TypeAction
                        Case Action.TypeAction.ActionDevice
                            If Actions.Item(j).IdDevice = Id Then
                                Actions.Item(j).IdDevice = ""
                                _retour = 0
                            End If

                        Case Action.TypeAction.ActionIf
                            Dim x As Action.ActionIf = Actions.Item(j)
                            For k As Integer = 0 To x.Conditions.Count - 1
                                If x.Conditions.Item(k).IdDevice = Id Then
                                    x.Conditions.Item(k).IdDevice = ""
                                    _retour = 0
                                End If
                            Next
                            DeleteIDToAction(IdSrv, x.ListTrue, Id)
                            DeleteIDToAction(IdSrv, x.ListFalse, Id)

                        Case Action.TypeAction.ActionMail
                            If Actions.Item(j).UserId = Id Then
                                Actions.Item(j).UserId = ""
                                _retour = 0
                            End If

                        Case Action.TypeAction.ActionMacro
                            If Actions.Item(j).IdMacro = Id Then
                                Actions.Item(j).IdMacro = ""
                                _retour = 0
                            End If
                    End Select
                Next

                Return _retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDeviceToZone", "Exception : " & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>Supprimer une macro de la config</summary>
        ''' <param name="macroId"></param>
        Public Function DeleteMacro(ByVal IdSrv As String, ByVal macroId As String) As Integer Implements IHoMIDom.DeleteMacro
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteMacro", "Idsrv incorrect, macros will not be deleted : " & macroId)
                    Return 99
                End If

                For i As Integer = 0 To _ListMacros.Count - 1
                    If _ListMacros.Item(i).ID = macroId Then
                        _ListMacros.RemoveAt(i)
                        SaveRealTime()
                        ManagerSequences.AddSequences(Sequence.TypeOfSequence.MacroDelete, macroId, Nothing, Nothing)

                        Return 0
                    End If
                Next
                Return -1
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteMacro", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Retourne la liste de toutes les macros</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetAllMacros(ByVal IdSrv As String) As List(Of Macro) Implements IHoMIDom.GetAllMacros
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllMacros", "Idsrv incorrect, macros will not be returned")
                    Return Nothing
                End If

                Dim _list As New List(Of Macro)
                For i As Integer = 0 To _ListMacros.Count - 1
                    Dim x As New Macro
                    With x
                        .Nom = _ListMacros.Item(i).Nom
                        .ID = _ListMacros.Item(i).ID
                        .Description = _ListMacros.Item(i).Description
                        .Enable = _ListMacros.Item(i).Enable
                        .ListActions = _ListMacros.Item(i).ListActions
                    End With
                    _list.Add(x)
                    x = Nothing
                Next

                _list.Sort(AddressOf sortMacro)
                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllMacros", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function sortMacro(ByVal x As Macro, ByVal y As Macro) As Integer
            Try
                Return x.Nom.CompareTo(y.Nom)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RunMacro", "Exception : " & ex.Message)
                Return 0
            End Try
        End Function

        ''' <summary>
        ''' Permet de créer ou modifier une macro
        ''' </summary>
        ''' <param name="macroId"></param>
        ''' <param name="nom"></param>
        ''' <param name="enable"></param>
        ''' <param name="description"></param>
        ''' <param name="listactions"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SaveMacro(ByVal IdSrv As String, ByVal macroId As String, ByVal nom As String, ByVal enable As Boolean, Optional ByVal description As String = "", Optional ByVal listactions As ArrayList = Nothing) As String Implements IHoMIDom.SaveMacro
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveMacro", "Idsrv incorrect macro will not be saved : " & nom)
                    Return 99
                End If

                Dim myID As String = ""
                If String.IsNullOrEmpty(macroId) = True Then
                    Dim x As New Macro
                    With x
                        x._Server = Me
                        x.ID = System.Guid.NewGuid.ToString()
                        x.Nom = nom
                        x.Enable = enable
                        x.Description = description
                        x.ListActions = listactions
                    End With
                    myID = x.ID
                    _ListMacros.Add(x)
                    SaveRealTime()
                Else
                    'macro Existante
                    myID = macroId
                    For i As Integer = 0 To _ListMacros.Count - 1
                        If _ListMacros.Item(i).ID = macroId Then
                            _ListMacros.Item(i).Nom = nom
                            _ListMacros.Item(i).Enable = enable
                            _ListMacros.Item(i).Description = description
                            _ListMacros.Item(i).ListActions = listactions
                            SaveRealTime()
                            Exit For
                        End If
                    Next
                End If

                ManagerSequences.AddSequences(Sequence.TypeOfSequence.MacroChange, myID, Nothing, Nothing)
                RaiseEvent MacroChanged(myID)

                Return myID
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveMacro", "Exception : " & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>Retourne la macro par son ID</summary>
        ''' <param name="MacroId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnMacroById(ByVal IdSrv As String, ByVal MacroId As String) As Macro Implements IHoMIDom.ReturnMacroById
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnMacroById", "Idsrv incorrect, macro will not be returnned : " & MacroId)
                    Return Nothing
                End If

                If (From Macro In _ListMacros Where Macro.ID = MacroId Select Macro).Count > 0 Then
                    Dim Resultat = (From Macro In _ListMacros Where Macro.ID = MacroId Select Macro).First
                    Return Resultat
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnMacroById", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        Public Sub RunMacro(ByVal IDSrv As String, ByVal Id As String) Implements IHoMIDom.RunMacro
            Try
                If VerifIdSrv(IDSrv) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RunMacro", "Idsrv incorrect, macro will not be runned : " & Id)
                    Exit Sub
                End If

                Execute(Id)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "RunMacro", "Exception : " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Trigger"
        ''' <summary>supprimer un device à une zone</summary>
        ''' <param name="TriggerId"></param>
        ''' <param name="DeviceId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function DeleteDeviceToTrigger(ByVal IdSrv As String, ByVal TriggerId As String, ByVal DeviceId As String) As String
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim _Trigger As Trigger = ReturnTriggerById(_IdSrv, TriggerId)
                Dim _retour As String = -1

                If _Trigger IsNot Nothing Then
                    If _Trigger.ConditionDeviceId = DeviceId Then
                        _Trigger.ConditionDeviceId = ""
                    End If
                    _retour = 0
                End If

                ManagerSequences.AddSequences(Sequence.TypeOfSequence.TriggerChange, TriggerId, Nothing, Nothing)

                Return _retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteDeviceToZone", "Exception : " & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>Supprimer un trigger de la config</summary>
        ''' <param name="triggerId"></param>
        Public Function DeleteTrigger(ByVal IdSrv As String, ByVal triggerId As String) As Integer Implements IHoMIDom.DeleteTrigger
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                For i As Integer = 0 To _ListTriggers.Count - 1
                    If _ListTriggers.Item(i).ID = triggerId Then
                        _ListTriggers.RemoveAt(i)
                        SaveRealTime()
                        ManagerSequences.AddSequences(Sequence.TypeOfSequence.TriggerDelete, triggerId, Nothing, Nothing)
                        Return 0
                    End If
                Next

                Return -1
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteTrigger", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Retourne la liste de toutes les macros</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetAllTriggers(ByVal IdSrv As String) As List(Of Trigger) Implements IHoMIDom.GetAllTriggers
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim _list As New List(Of Trigger)

                For i As Integer = 0 To _ListTriggers.Count - 1
                    Dim x As New Trigger
                    With x
                        .Nom = _ListTriggers.Item(i).Nom
                        .ID = _ListTriggers.Item(i).ID
                        .Description = _ListTriggers.Item(i).Description
                        .Enable = _ListTriggers.Item(i).Enable
                        .Type = _ListTriggers.Item(i).Type
                        Try
                            If Not IsNothing(_ListTriggers.Item(i).Prochainedateheure) Then .Prochainedateheure = _ListTriggers.Item(i).Prochainedateheure
                            If Not IsNothing(_ListTriggers.Item(i).ConditionTime) Then .ConditionTime = _ListTriggers.Item(i).ConditionTime
                            If Not IsNothing(_ListTriggers.Item(i).ConditionDeviceId) Then .ConditionDeviceId = _ListTriggers.Item(i).ConditionDeviceId
                            If Not IsNothing(_ListTriggers.Item(i).ConditionDeviceProperty) Then .ConditionDeviceProperty = _ListTriggers.Item(i).ConditionDeviceProperty
                            If Not IsNothing(_ListTriggers.Item(i).ListMacro) Then .ListMacro = _ListTriggers.Item(i).ListMacro
                        Catch ex As Exception
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllTriggers", "Exception part2 : " & .Nom & " ---" & ex.Message)
                        End Try

                    End With
                    _list.Add(x)
                Next

                _list.Sort(AddressOf sortTrigger)
                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllTriggers", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function sortTrigger(ByVal x As Trigger, ByVal y As Trigger) As Integer
            Try
                Return x.Nom.CompareTo(y.Nom)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "sortTrigger", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Permet de créer ou modifier un trigger
        ''' </summary>
        ''' <param name="triggerId"></param>
        ''' <param name="nom"></param>
        ''' <param name="enable"></param>
        ''' <param name="description"></param>
        ''' <param name="conditiontimer"></param>
        ''' <param name="deviceid"></param>
        ''' <param name="deviceproperty"></param>
        ''' <param name="TypeTrigger"></param>
        ''' <param name="macro"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function SaveTrigger(ByVal IdSrv As String, ByVal triggerId As String, ByVal nom As String, ByVal enable As Boolean, ByVal TypeTrigger As Trigger.TypeTrigger, Optional ByVal description As String = "", Optional ByVal conditiontimer As String = "", Optional ByVal deviceid As String = "", Optional ByVal deviceproperty As String = "", Optional ByVal macro As List(Of String) = Nothing) As String Implements IHoMIDom.SaveTrigger
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim myID As String = ""

                If String.IsNullOrEmpty(triggerId) = True Then
                    Dim x As New Trigger
                    With x
                        x._Server = Me
                        x.ID = System.Guid.NewGuid.ToString()
                        x.Nom = nom
                        x.Enable = enable
                        Select Case TypeTrigger
                            Case Trigger.TypeTrigger.TIMER
                                x.Type = Trigger.TypeTrigger.TIMER
                                x.ConditionTime = conditiontimer
                            Case Trigger.TypeTrigger.DEVICE
                                x.Type = Trigger.TypeTrigger.DEVICE
                                x.ConditionDeviceId = deviceid
                                x.ConditionDeviceProperty = deviceproperty
                        End Select
                        If String.IsNullOrEmpty(description) = False Then x.Description = description
                        If macro IsNot Nothing Then
                            If macro.Count > 0 Then
                                x.ListMacro = macro
                            End If
                        End If
                        If TypeTrigger = Trigger.TypeTrigger.TIMER Then x.maj_cron()
                    End With
                    myID = x.ID
                    _ListTriggers.Add(x)
                    SaveRealTime()
                    x = Nothing
                    ManagerSequences.AddSequences(Sequence.TypeOfSequence.TriggerAdd, myID, Nothing, Nothing)
                Else
                    'trigger Existante
                    myID = triggerId
                    For i As Integer = 0 To _ListTriggers.Count - 1
                        If _ListTriggers.Item(i).ID = triggerId Then
                            _ListTriggers.Item(i).Nom = nom
                            _ListTriggers.Item(i).Enable = enable
                            _ListTriggers.Item(i).Description = description
                            Select Case TypeTrigger
                                Case Trigger.TypeTrigger.TIMER
                                    _ListTriggers.Item(i).Type = HoMIDom.Trigger.TypeTrigger.TIMER
                                    _ListTriggers.Item(i).ConditionTime = conditiontimer
                                Case Trigger.TypeTrigger.DEVICE
                                    _ListTriggers.Item(i).Type = HoMIDom.Trigger.TypeTrigger.DEVICE
                                    _ListTriggers.Item(i).ConditionDeviceId = deviceid
                                    _ListTriggers.Item(i).ConditionDeviceProperty = deviceproperty
                            End Select
                            If macro IsNot Nothing Then _ListTriggers.Item(i).ListMacro = macro
                            If TypeTrigger = Trigger.TypeTrigger.TIMER Then _ListTriggers.Item(i).maj_cron()
                            SaveRealTime()
                            ManagerSequences.AddSequences(Sequence.TypeOfSequence.TriggerChange, triggerId, Nothing, Nothing)
                        End If
                    Next
                End If

                Return myID
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveTrigger", "Exception : " & ex.Message & vbCrLf & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>Retourne le trigger par son ID</summary>
        ''' <param name="TriggerId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnTriggerById(ByVal IdSrv As String, ByVal TriggerId As String) As Trigger Implements IHoMIDom.ReturnTriggerById
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim Resultat = (From Trigger In _ListTriggers Where Trigger.ID = TriggerId Select Trigger).First
                Return Resultat

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnTriggerById", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

#End Region

#Region "User"
        ''' <summary>Supprime un user</summary>
        ''' <param name="userId"></param>
        Public Function DeleteUser(ByVal IdSrv As String, ByVal userId As String) As Integer Implements IHoMIDom.DeleteUser
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                For i As Integer = 0 To _ListUsers.Count - 1
                    If _ListUsers.Item(i).ID = userId Then
                        _ListUsers.RemoveAt(i)
                        SaveRealTime()
                        ManagerSequences.AddSequences(Sequence.TypeOfSequence.UserDelete, userId, Nothing, Nothing)
                        Return 0
                    End If
                Next

                Return -1
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteUser", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Retourne la liste de tous les users</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function GetAllUsers(ByVal IdSrv As String) As List(Of Users.User) Implements IHoMIDom.GetAllUsers
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim _list As New List(Of Users.User)

                For i As Integer = 0 To _ListUsers.Count - 1
                    Dim x As New Users.User
                    With x
                        .Adresse = _ListUsers.Item(i).Adresse
                        .CodePostal = _ListUsers.Item(i).CodePostal
                        .eMail = _ListUsers.Item(i).eMail
                        .eMailAutre = _ListUsers.Item(i).eMailAutre
                        .ID = _ListUsers.Item(i).ID
                        .Image = _ListUsers.Item(i).Image
                        .Nom = _ListUsers.Item(i).Nom
                        .NumberIdentification = _ListUsers.Item(i).NumberIdentification
                        .Password = _ListUsers.Item(i).Password
                        .Prenom = _ListUsers.Item(i).Prenom
                        .Profil = _ListUsers.Item(i).Profil
                        .TelAutre = _ListUsers.Item(i).TelAutre
                        .TelFixe = _ListUsers.Item(i).TelFixe
                        .TelMobile = _ListUsers.Item(i).TelMobile
                        .UserName = _ListUsers.Item(i).UserName
                        .Ville = _ListUsers.Item(i).Ville
                    End With
                    _list.Add(x)
                Next

                _list.Sort(AddressOf sortUser)
                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllUsers", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        Private Function sortUser(ByVal x As Users.User, ByVal y As Users.User) As Integer
            Try
                Return x.UserName.CompareTo(y.UserName)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "sortUser", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Retourne un user par son username
        ''' </summary>
        ''' <param name="Username"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnUserByUsername(ByVal IdSrv As String, ByVal Username As String) As Users.User Implements IHoMIDom.ReturnUserByUsername
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim Resultat = (From User In _ListUsers Where User.UserName = Username Select User).First
                Return Resultat
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnUserByUsername", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Créer ou modifie un user par son ID
        ''' </summary>
        ''' <param name="userId"></param>
        ''' <param name="UserName"></param>
        ''' <param name="Password"></param>
        ''' <param name="Profil"></param>
        ''' <param name="Nom"></param>
        ''' <param name="Prenom"></param>
        ''' <param name="NumberIdentification"></param>
        ''' <param name="Image"></param>
        ''' <param name="eMail"></param>
        ''' <param name="eMailAutre"></param>
        ''' <param name="TelFixe"></param>
        ''' <param name="TelMobile"></param>
        ''' <param name="TelAutre"></param>
        ''' <param name="Adresse"></param>
        ''' <param name="Ville"></param>
        ''' <param name="CodePostal"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function SaveUser(ByVal IdSrv As String, ByVal userId As String, ByVal UserName As String, ByVal Password As String, ByVal Profil As Users.TypeProfil, ByVal Nom As String, ByVal Prenom As String, Optional ByVal NumberIdentification As String = "", Optional ByVal Image As String = "", Optional ByVal eMail As String = "", Optional ByVal eMailAutre As String = "", Optional ByVal TelFixe As String = "", Optional ByVal TelMobile As String = "", Optional ByVal TelAutre As String = "", Optional ByVal Adresse As String = "", Optional ByVal Ville As String = "", Optional ByVal CodePostal As String = "") As String Implements IHoMIDom.SaveUser
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim myID As String = ""
                If String.IsNullOrEmpty(userId) = True Then
                    For i As Integer = 0 To _ListUsers.Count - 1
                        If _ListUsers.Item(i).UserName = UserName Then
                            myID = "ERROR Username déjà utlisé"
                            Return myID
                        End If
                    Next
                    Dim x As New Users.User
                    With x
                        x.ID = System.Guid.NewGuid.ToString()
                        x.Adresse = Adresse
                        x.CodePostal = CodePostal
                        x.eMail = eMail
                        x.eMailAutre = eMailAutre
                        x.Image = Image
                        x.Nom = Nom
                        x.NumberIdentification = NumberIdentification
                        x.Password = EncryptTripleDES(Password, "homidom")
                        x.Prenom = Prenom
                        x.Profil = Profil
                        x.TelAutre = TelAutre
                        x.TelFixe = TelFixe
                        x.TelMobile = TelMobile
                        x.UserName = UserName
                        x.Ville = Ville
                    End With
                    myID = x.ID
                    _ListUsers.Add(x)
                    SaveRealTime()
                    ManagerSequences.AddSequences(Sequence.TypeOfSequence.UserAdd, myID, Nothing, Nothing)
                Else
                    'user Existant
                    myID = userId
                    For i As Integer = 0 To _ListUsers.Count - 1
                        If _ListUsers.Item(i).ID = userId Then
                            _ListUsers.Item(i).Adresse = Adresse
                            _ListUsers.Item(i).CodePostal = CodePostal
                            _ListUsers.Item(i).eMail = eMail
                            _ListUsers.Item(i).eMailAutre = eMailAutre
                            _ListUsers.Item(i).Image = Image
                            _ListUsers.Item(i).Nom = Nom
                            _ListUsers.Item(i).NumberIdentification = NumberIdentification
                            _ListUsers.Item(i).Password = EncryptTripleDES(Password, "homidom")
                            _ListUsers.Item(i).Prenom = Prenom
                            _ListUsers.Item(i).Profil = Profil
                            _ListUsers.Item(i).TelAutre = TelAutre
                            _ListUsers.Item(i).TelFixe = TelFixe
                            _ListUsers.Item(i).TelMobile = TelMobile
                            _ListUsers.Item(i).UserName = UserName
                            _ListUsers.Item(i).Ville = Ville
                            SaveRealTime()
                            ManagerSequences.AddSequences(Sequence.TypeOfSequence.VariableChange, userId, Nothing, Nothing)
                        End If
                    Next
                End If

                'génération de l'event
                Return myID
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveUser", "Exception : " & ex.Message)
                Return "-1"
            End Try
        End Function

        ''' <summary>Vérifie le couple username password</summary>
        ''' <param name="Username"></param>
        ''' <param name="Password"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function VerifLogin(ByVal Username As String, ByVal Password As String) As Boolean Implements IHoMIDom.VerifLogin
            Try
                Dim retour As Boolean = False

                For i As Integer = 0 To _ListUsers.Count - 1
                    If _ListUsers.Item(i).UserName = Username Then
                        Dim a As String = EncryptTripleDES(Password, "homidom")
                        If a = _ListUsers.Item(i).Password Then
                            Return True
                        End If
                    End If
                Next

                Return retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "VerifLogin", "Exception : " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>Permet de changer de Password sur un user</summary>
        ''' <param name="Username"></param>
        ''' <param name="OldPassword"></param>
        ''' <param name="Password"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function ChangePassword(ByVal IdSrv As String, ByVal Username As String, ByVal OldPassword As String, ByVal ConfirmNewPassword As String, ByVal Password As String) As Boolean Implements IHoMIDom.ChangePassword
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return False
                End If

                Dim retour As Boolean = False
                For i As Integer = 0 To _ListUsers.Count - 1
                    If _ListUsers.Item(i).UserName = Username Then
                        If _ListUsers.Item(i).Password = OldPassword Then
                            If ConfirmNewPassword = Password Then
                                _ListUsers.Item(i).Password = EncryptTripleDES(Password, "homidom")
                                retour = True
                                Exit For
                            End If
                        End If
                    End If
                Next
                Return retour
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ChangePassword", "Exception : " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>Retourne un user par son ID</summary>
        ''' <param name="UserId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnUserById(ByVal IdSrv As String, ByVal UserId As String) As Users.User Implements IHoMIDom.ReturnUserById
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return Nothing
                End If

                Dim Resultat = (From User In _ListUsers Where User.ID = UserId Select User).First
                Return Resultat

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnUserById", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

#End Region

#Region "Telecommande"
        ''' <summary>
        ''' Retourne la liste des templates télécommande (fichier xml), présents dans le répertoire templates
        ''' </summary>
        ''' <returns>List of Templates</returns>
        ''' <remarks></remarks>
        Public Function GetListOfTemplate() As List(Of Telecommande.Template) Implements IHoMIDom.GetListOfTemplate
            Try
                Dim Tabl As New List(Of Telecommande.Template)
                Dim MyPath As String = _MonRepertoire & "\templates\"
                Dim xml As XML = Nothing

                Dim dirInfo As New System.IO.DirectoryInfo(MyPath)
                Dim file As System.IO.FileInfo
                Dim files() As System.IO.FileInfo = dirInfo.GetFiles("*.xml", System.IO.SearchOption.AllDirectories)

                If (files IsNot Nothing) Then
                    For Each file In files
                        Try
                            'Deserialize text file to a new object.
                            Dim x As New XmlSerializer(GetType(Telecommande.Template))
                            Dim objStreamReader As New StreamReader(file.FullName)
                            Dim _template As Telecommande.Template
                            _template = x.Deserialize(objStreamReader)
                            objStreamReader.Close()
                            Tabl.Add(_template)
                        Catch ex As Exception
                            Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetListOfTemplate", "Erreur lors du chargement du template " & file.FullName & " : " & ex.Message)
                        End Try

                    Next
                End If

                Return Tabl
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetListOfTemplate", "Erreur : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retourne un template via son ID
        ''' </summary>
        ''' <returns>List of Templates</returns>
        ''' <remarks></remarks>
        Public Function GetTemplateFromID(ByVal ID As String) As Telecommande.Template Implements IHoMIDom.GetTemplateFromID
            Try
                Dim Tabl As New List(Of Telecommande.Template)
                Dim _template As Telecommande.Template = Nothing
                Tabl = Me.GetListOfTemplate

                For Each _tpl In Tabl
                    If _tpl.ID = ID Then
                        _template = _tpl
                        Exit For
                    End If
                Next

                Return _template
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetTemplateFromID", "Erreur : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Crée un nouveau template dans le répertoire templates
        ''' </summary>
        ''' <returns>0 si ok, sinon message d'erreur</returns>
        ''' <remarks></remarks>
        Public Function CreateNewTemplate(ByVal Template As Telecommande.Template) As String Implements IHoMIDom.CreateNewTemplate
            Try
                Dim MyPath As String = _MonRepertoire & "\templates\"
                Dim _Fichier As String = MyPath & Template.Name & ".xml"

                If IO.File.Exists(_Fichier) Then
                    Log(TypeLog.DEBUG, TypeSource.SERVEUR, "CreateNewTemplate", "Le template existe déjà pour ce même nom!")
                    Return "Le template existe déjà pour ce même nom!"
                End If

                If Template.Variables.Count = 0 Then
                    'variable "command" qui sert à envoyer une commande
                    Dim x As New Telecommande.TemplateVar
                    x.Name = "command"
                    x.Type = Telecommande.TypeOfVar.String
                    Template.Variables.Add(x)

                    'variable "ip" qui contient l'adresse ip du template
                    Dim y As New Telecommande.TemplateVar
                    y.Name = "ip"
                    y.Type = Telecommande.TypeOfVar.String
                    Template.Variables.Add(y)

                    'variable "trame" qui contient le retour de trame
                    Dim z As New Telecommande.TemplateVar
                    z.Name = "trame"
                    z.Type = Telecommande.TypeOfVar.String
                    Template.Variables.Add(z)
                End If

                Template.ID = Api.GenerateGUID

                Template.InitCmd()

                Template = CreateDefautTemplateGrafic(Template)

                Dim streamIO As StreamWriter = Nothing
                Try
                    Dim serialXML As New System.Xml.Serialization.XmlSerializer(GetType(Telecommande.Template))
                    ' Ouverture d'un flux en écriture sur le fichier XML des contacts
                    streamIO = New StreamWriter(_Fichier, False)
                    ' Sérialisation de la liste des contacts
                    serialXML.Serialize(streamIO, Template)
                Catch ex As Exception
                    ' Propagrer l'exception
                    Throw ex
                Finally
                    ' En cas d'erreur, n'oublier pas de fermer le flux en écriture si ce dernier est toujours ouvert
                    If streamIO IsNot Nothing Then
                        streamIO.Close()
                    End If
                End Try

                Log(TypeLog.INFO, TypeSource.SERVEUR, "CreateNewTemplate", "Nouveau template créé: " & Template.Name)

                Return "0"
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "CreateNewTemplate", "Erreur : " & ex.Message)
                Return ex.Message
            End Try
        End Function

        '''' <param name="Fabricant"></param>
        '''' <param name="Modele"></param>
        '''' <param name="Driver"></param>
        ''' <summary>
        ''' Recupère la liste des commandes d'un template donné
        ''' </summary>
        ''' <returns>Liste de commandes</returns>
        ''' <remarks></remarks>
        Public Function ReadTemplate(ByVal Name As String) As List(Of Telecommande.Commandes)
            Try
                Dim _list As New List(Of Telecommande.Commandes)

                Dim MyPath As String = _MonRepertoire & "\templates\" & Name & ".xml"

                If IO.File.Exists(MyPath) = False Then
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReadTemplate", "Erreur le fichier n'existe pas: " & MyPath)
                    Return Nothing
                End If

                Try
                    'Deserialize text file to a new object.
                    Dim x As New XmlSerializer(GetType(Telecommande.Template))
                    Dim objStreamReader As New StreamReader(MyPath)
                    Dim _template As Telecommande.Template
                    _template = x.Deserialize(objStreamReader)
                    objStreamReader.Close()
                    If _template IsNot Nothing Then _list = _template.Commandes
                Catch ex As Exception
                    Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReadTemplate", "Erreur lors du chargement du template " & Name & " : " & ex.Message)
                End Try

                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReadTemplate", "Erreur : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>Demander un apprentissage à un driver</summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function StartLearning(ByVal IdSrv As String, ByVal DriverId As String) As String Implements IHoMIDom.StartLearning
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return "ERREUR: l'Id du serveur est erroné"
                End If

                Dim retour As String = ""
                For i As Integer = 0 To _ListDrivers.Count - 1
                    If _ListDrivers.Item(i).ID = DriverId Then
                        Dim x As Object = _ListDrivers.Item(i)
                        If x.GetType().GetMethod("LearnCode") = Nothing Then
                            Return "ERR: Ce driver ne possede de fonction LearnCode"
                        Else
                            retour = x.LearnCode()
                            Log(TypeLog.INFO, TypeSource.SERVEUR, "SERVEUR", "StartLearning: " & retour)
                            If String.IsNullOrEmpty(retour) Then
                                Return "ERR: Erreur lors de l'apprentissage du code (retour vide), veuillez consulter les logs du serveur pour en savoir plus"
                            Else
                                Return retour
                            End If
                        End If
                    End If
                Next
                Return ""

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "StartLearning", "Erreur : " & ex.Message)
                Return ("ERREUR: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Sauvegarde un template donné
        ''' </summary>
        ''' <param name="IdSrv">Id du Serveur</param>
        ''' <param name="Template">Nom du template</param>
        ''' <returns>O si ok sinon message d'erreur</returns>
        ''' <remarks></remarks>
        Public Function SaveTemplate(ByVal IdSrv As String, ByVal Template As Telecommande.Template) As String Implements IHoMIDom.SaveTemplate
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim MyPath As String = _MonRepertoire & "\templates\"
                Dim _Fichier As String = MyPath & Template.Name & ".xml"

                If IO.File.Exists(_Fichier) = False Then
                    Return "Le template " & Template.Name & ".xml n'existe pas!"
                End If

                Dim streamIO As StreamWriter = Nothing
                Try
                    Dim serialXML As New System.Xml.Serialization.XmlSerializer(GetType(Telecommande.Template))
                    ' Ouverture d'un flux en écriture sur le fichier XML des contacts
                    streamIO = New StreamWriter(_Fichier, False)
                    ' Sérialisation de la liste des contacts
                    serialXML.Serialize(streamIO, Template)
                Catch ex As Exception
                    ' Propagrer l'exception
                    Throw ex
                Finally
                    ' En cas d'erreur, n'oublier pas de fermer le flux en écriture si ce dernier est toujours ouvert
                    If streamIO IsNot Nothing Then
                        streamIO.Close()
                    End If
                End Try

                Return "0"
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveTemplate", "Erreur : " & ex.Message)
                Return ex.Message
            End Try
        End Function

        ''' <summary>
        ''' Supprimer un template donné
        ''' </summary>
        ''' <param name="IdSrv">Id du Serveur</param>
        ''' <param name="Template">Nom du template</param>
        ''' <returns>O si ok sinon message d'erreur</returns>
        ''' <remarks></remarks>
        Public Function DeleteTemplate(ByVal IdSrv As String, ByVal Template As Telecommande.Template) As String Implements IHoMIDom.DeleteTemplate
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim MyPath As String = _MonRepertoire & "\templates\"
                Dim _Fichier As String = MyPath & Template.Name & ".xml"

                If IO.File.Exists(_Fichier) = False Then
                    Return "Le template " & Template.Name & ".xml n'existe pas!"
                Else
                    IO.File.Delete(_Fichier)
                End If

                Return "0"
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveTemplate", "Erreur : " & ex.Message)
                Return ex.Message
            End Try
        End Function


        ''' <summary>
        ''' Efface la partie graphique du template
        ''' </summary>
        ''' <param name="Template"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function EffaceTemplateGraphic(ByVal Template As Telecommande.Template) As String
            Try
                Template.GraphicTemplate = Nothing

                Return "0"
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "EffaceTemplate", "Erreur : " & ex.Message)
                Return ex.Message
            End Try
        End Function

        ''' <summary>
        ''' Cree un template graphique par defaut
        ''' </summary>
        ''' <param name="Template"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function CreateDefautTemplateGrafic(ByVal Template As Telecommande.Template) As Telecommande.Template
            Try
                Dim _NewTemplate As Telecommande.Template = Template

                _NewTemplate.GraphicTemplate.Width = 800
                _NewTemplate.GraphicTemplate.Height = 600
                _NewTemplate.GraphicTemplate.BackGroundPicture = _MonRepertoire & "\Images\Telecommande\front_multimedia.png"
                _NewTemplate.GraphicTemplate.Widgets.Clear()

                If _NewTemplate.IsAudioVideo And _NewTemplate.Commandes.Count > 0 Then
                    For Each cmd In _NewTemplate.Commandes
                        Dim _widget As New Widget
                        Dim _output As New Widget.Output
                        Dim _picture As New Widget.Picture
                        Dim _margin As Integer = 7

                        _widget.Label = cmd.Name
                        _widget.Height = 56
                        _widget.Width = 56

                        _output.Commande = cmd.Name
                        _output.TemplateID = _NewTemplate.ID
                        _widget.Outputs.Add(_output)

                        _picture.Path = _MonRepertoire & "\images\telecommande\" & cmd.Name & ".png"
                        _widget.Pictures.Add(_picture)

                        Dim _col As Byte = 0
                        Dim _row As Byte = 0

                        Select Case cmd.Name.ToUpper
                            Case "0"
                                _row = 4
                                _col = 2
                            Case "1"
                                _row = 1
                                _col = 1
                            Case "2"
                                _row = 1
                                _col = 2
                            Case "3"
                                _row = 1
                                _col = 3
                            Case "4"
                                _row = 2
                                _col = 1
                            Case "5"
                                _row = 2
                                _col = 2
                            Case "6"
                                _row = 2
                                _col = 3
                            Case "7"
                                _row = 3
                                _col = 1
                            Case "8"
                                _row = 3
                                _col = 2
                            Case "9"
                                _row = 3
                                _col = 3
                            Case "PLAY"
                                _row = 8
                                _col = 9
                            Case "PAUSE"
                                _row = 8
                                _col = 10
                            Case "STOP"
                                _row = 8
                                _col = 8
                            Case "POWER"
                                _row = 0
                                _col = 0
                            Case "AVANCE"
                                _row = 9
                                _col = 9
                            Case "RECUL"
                                _row = 9
                                _col = 8
                            Case "NEXTCHAPITRE"
                                _row = 9
                                _col = 10
                            Case "PREVIOUSCHAPITRE"
                                _row = 9
                                _col = 7
                            Case "OK"
                                _row = 6
                                _col = 2
                            Case "VOLUMEUP"
                                _row = 5
                                _col = 0
                            Case "VOLUMEDOWN"
                                _row = 7
                                _col = 0
                            Case "MUTE"
                                _row = 6
                                _col = 0
                            Case "FLECHEHAUT"
                                _row = 5
                                _col = 2
                            Case "FLECHEBAS"
                                _row = 7
                                _col = 2
                            Case "FLECHEGAUCHE"
                                _row = 6
                                _col = 1
                            Case "FLECHEDROITE"
                                _row = 6
                                _col = 3
                            Case "ENREGISTRER"
                                _row = 8
                                _col = 7
                            Case "BLUE"
                                _row = 9
                                _col = 0
                            Case "RED"
                                _row = 9
                                _col = 1
                            Case "GREEN"
                                _row = 9
                                _col = 2
                            Case "YELLOW"
                                _row = 9
                                _col = 3
                            Case "CHANNELUP"
                                _row = 5
                                _col = 4
                            Case "CHANNELDOWN"
                                _row = 7
                                _col = 4
                        End Select
                        _widget.X = (_col * (_widget.Width + 5))
                        _widget.Y = (_row * _widget.Height)

                        _NewTemplate.GraphicTemplate.Widgets.Add(_widget)
                    Next
                End If

                Return _NewTemplate
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "CreateDefautTemplate", "Erreur : " & ex.Message)
                Return Template
            End Try
        End Function

        ''' <summary>
        ''' Demande au device d'envoyer une commande (telecommande) à son driver
        ''' </summary>
        ''' <param name="IdSrv">Id du serveur, retourne 99 si non OK</param>
        ''' <param name="IdDevice">Id du device concerné</param>
        ''' <param name="Commande">Nom de la Commande à envoyée</param>
        ''' <returns>0 si Ok sinon erreur</returns>
        ''' <remarks></remarks>
        Public Function TelecommandeSendCommand(ByVal IdSrv As String, ByVal IdDevice As String, ByVal Commande As String) As String Implements IHoMIDom.TelecommandeSendCommand
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim x As Object = Nothing

                For i As Integer = 0 To _ListDevices.Count - 1
                    If _ListDevices.Item(i).ID = IdDevice Then
                        _ListDevices.Item(i).EnvoyerCommande(Commande)
                        Return "0"
                    End If
                Next

                Return "Erreur: le device n'a pas été trouvé"
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "TelecommandeSendCommand", "Erreur : " & ex.Message)
                Return "Erreur lors du traitement de la fonction TelecommandeSendCommand: " & ex.Message
            End Try
        End Function

        Public Function DriverTelecommandeSendCommand(ByVal IdSrv As String, ByVal IdDriver As String, ByVal Type As Integer, ByVal Parametres As List(Of String)) As String Implements IHoMIDom.DriverTelecommandeSendCommand
            Try
                Dim _driver As Object = Nothing
                _driver = ReturnDrvById(IdSrv, IdDriver)

                Select Case Type
                    Case 0 'http
                        _driver.EnvoyerCode(Parametres(0), Parametres(1)) 'EnvoyerCode(code, cmd.Repeat)
                    Case 1 'ir
                        _driver.EnvoyerCode(Parametres(0), Parametres(1), Parametres(2)) 'EnvoyerCode(code, cmd.Repeat, cmd.Format)
                    Case 2 'rs232
                        _driver.EnvoyerCode(Parametres(0), Parametres(1), Parametres(2), Parametres(3), Parametres(4), Parametres(5), Parametres(6)) 'EnvoyerCode(code, Me.Port, 9600, 8, 0, 1, cmd.Repeat)
                End Select

                Return "0"
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DriverTelecommandeSendCommand", "Erreur : " & ex.Message)
                Return "Erreur lors du traitement de la fonction DriverTelecommandeSendCommand: " & ex.Message
            End Try
        End Function

        ''' <summary>
        ''' Exporte le template vers une destination
        ''' </summary>
        ''' <param name="IdSrv">Id du serveur</param>
        ''' <returns>le fichier sous format text si ok sinon message d'erreur commençant par ERREUR</returns>
        ''' <remarks></remarks>
        Public Function ExportTemplateMultimedia(ByVal IdSrv As String, ByVal Template As Telecommande.Template) As String Implements IHoMIDom.ExportTemplateMultimedia
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                Dim retour As String
                Dim SR As New StreamReader(Template.File)
                retour = SR.ReadToEnd()
                SR.Close()
                Return retour
            Catch ex As Exception
                Return "ERREUR lors de l'exportation du fichier de template: " & ex.Message
            End Try
        End Function

        ''' <summary>
        ''' Importe un fichier de Template depuis une source
        ''' </summary>
        ''' <returns>"0" si ok sinon message d'erreur</returns>
        ''' <remarks></remarks>
        Public Function ImportTemplateMultimedia(ByVal IdSrv As String, ByVal Template As Telecommande.Template) As String Implements IHoMIDom.ImportTemplateMultimedia
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return "L'Id du serveur est incorrect!"
                End If

                If IO.File.Exists(Template.File) = False Then
                    Return "Le serveur n'a pas trouvé le fichier de template !"
                End If

                'sauvegarde de l'ancien fichier sous .old
                IO.File.Copy(Template.File, Template.File.Replace("xml", "old"), True)
                IO.File.Copy(Template.File, _MonRepertoire & "\templates, True")

                Return "0"
            Catch ex As Exception
                Return "Erreur lors de l'importation du fichier de config: " & ex.Message
            End Try
        End Function


#End Region

#Region "Log"

        ''' <summary>
        ''' Retourne pour chaque type de log s'il doit être pris en compte ou non
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTypeLogEnable() As List(Of Boolean) Implements IHoMIDom.GetTypeLogEnable
            Try
                Dim _list As New List(Of Boolean)
                For i As Integer = 0 To _TypeLogEnable.Count - 1
                    If _TypeLogEnable.Item(i) = True Then
                        _list.Add(False)
                    Else
                        _list.Add(True)
                    End If
                Next
                Return _list
                _list = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetTypeLogEnable", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Fixe si chaque type de log doit être pris en compte ou non
        ''' </summary>
        ''' <param name="ListTypeLogEnable"></param>
        ''' <remarks></remarks>
        Public Sub SetTypeLogEnable(ByVal ListTypeLogEnable As List(Of Boolean)) Implements IHoMIDom.SetTypeLogEnable
            Try
                For i As Integer = 0 To ListTypeLogEnable.Count - 1
                    If ListTypeLogEnable(i) = True Then
                        _TypeLogEnable.Item(i) = False
                    Else
                        _TypeLogEnable.Item(i) = True
                    End If
                Next
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetTypeLogEnable", "Erreur : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne les derniers logs les plus récents (du plus récent au plus ancien)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLastLogs() As List(Of String) Implements IHoMIDom.GetLastLogs
            Try
                Dim list As New List(Of String)
                For i = 0 To (_LastLogs.Count - 1)
                    list.Add(_LastLogs(i))
                Next
                Return list
                list = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetLastLogs", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retourne les logs en erreur les plus récents (du plus récent au plus ancien)
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetLastLogsError() As List(Of String) Implements IHoMIDom.GetLastLogsError
            Try
                Dim list As New List(Of String)
                For i = 0 To (_LastLogsError.Count - 1)
                    list.Add(_LastLogsError(i))
                Next
                Return list
                list = Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetLastLogsError", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Retourne le nombre de mois à conserver une archive de log avant de le supprimer
        ''' </summary>
        ''' <param name="Month"></param>
        ''' <remarks></remarks>
        Public Sub SetMaxMonthLog(ByVal Month As Integer) Implements IHoMIDom.SetMaxMonthLog
            Try
                If IsNumeric(Month) Then
                    _MaxMonthLog = Month
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetMaxMonthLog", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Définit le nombre de mois à conserver une archive de log avant de le supprimer
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxMonthLog() As Integer Implements IHoMIDom.GetMaxMonthLog
            Try
                Return _MaxMonthLog
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetMaxMonthLog", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>renvoi le fichier log suivant une requête xml si besoin</summary>
        ''' <param name="Requete">X représentant le nombre des derniers historiques à renvoyer, si non définit renvoi tout</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Function ReturnLog(Optional ByVal Requete As String = "") As String Implements IHoMIDom.ReturnLog
            Try
                Dim retour As String = ""
                If String.IsNullOrEmpty(Requete) = True Then
                    If System.IO.File.Exists(_MonRepertoire & "\logs\log_" & DateAndTime.Now.ToString("yyyyMMdd") & ".txt") Then
                        Dim SR As New StreamReader(_MonRepertoire & "\logs\log_" & DateAndTime.Now.ToString("yyyyMMdd") & ".txt", Encoding.GetEncoding("ISO-8859-1"))
                        retour = SR.ReadToEnd()
                        retour = HtmlDecode(retour)
                        SR.Close()
                    Else
                        retour = ""
                    End If
                Else
                    If IsNumeric(Requete) Then
                        Dim cnt As Integer = CInt(Requete)
                        If cnt < 8 Then cnt = 0

                        Dim lignes() As String = IO.File.ReadAllLines(_MonRepertoire & "\logs\log_" & DateAndTime.Now.ToString("yyyyMMdd") & ".txt")
                        Dim cnt1 As Integer = lignes.Length - cnt

                        If cnt1 <= 0 Then cnt1 = 0

                        For i As Integer = 0 To lignes.Length - 1
                            If i >= cnt1 Then
                                retour &= lignes(i)
                            End If
                        Next

                        retour = HtmlDecode(retour)
                    Else
                        'creation d'une nouvelle instance du membre xmldocument
                        Dim XmlDoc As XmlDocument = New XmlDocument()
                        XmlDoc.Load(_MonRepertoire & "\logs\log.xml")
                    End If

                End If
                If retour.Length > 1000000 Then
                    Dim retour2 As String = Mid(retour, retour.Length - 1000001, 1000000)
                    retour = Now & vbTab & "ERREUR" & vbTab & "SERVER" & vbTab & "ReturnLog" & vbTab & "Trop de lignes à traiter dans le log du jour, seules les dernières lignes seront affichées, merci de consulter le fichier sur le serveur par en avoir la totalité." & vbCrLf & vbCrLf & retour2
                    Return retour
                End If
                Return retour
            Catch ex As Exception
                ReturnLog = "Erreur lors de la récupération du log: " & ex.Message
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnLog", "Exception : " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Fixe la taille max du fichier log en Ko avant d'en créer un nouveau
        ''' </summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub SetMaxFileSizeLog(ByVal Value As Long) Implements IHoMIDom.SetMaxFileSizeLog
            Try
                MaxFileSize = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetMaxFileSizeLog", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne la taille max du fichier log en Ko 
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetMaxFileSizeLog() As Long Implements IHoMIDom.GetMaxFileSizeLog
            Try
                Return MaxFileSize
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetMaxFileSizeLog", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function
#End Region

#Region "Configuration"
        ''' <summary>
        ''' Exporte le fichier de config vers une destination
        ''' </summary>
        ''' <param name="IdSrv">Id du serveur</param>
        ''' <returns>le fichier sous format text si ok sinon message d'erreur commençant par ERREUR</returns>
        ''' <remarks></remarks>
        Public Function ExportConfig(ByVal IdSrv As String) As String Implements IHoMIDom.ExportConfig
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return "ERREUR: L'Id du serveur est incorrect"
                End If

                Dim retour As String
                Dim SR As New StreamReader(_MonRepertoire & "\Config\homidom.xml")
                retour = SR.ReadToEnd()
                SR.Close()
                Return retour
            Catch ex As Exception
                Return "ERREUR lors de l'exportation du fichier de config: " & ex.Message
            End Try
        End Function

        ''' <summary>
        ''' Importe un fichier de config depuis une source
        ''' </summary>
        ''' <param name="Source">chemin + fichier (homidom.xml)</param>
        ''' <returns>"0" si ok sinon message d'erreur</returns>
        ''' <remarks></remarks>
        Public Function ImportConfig(ByVal IdSrv As String, ByVal Source As String) As String Implements IHoMIDom.ImportConfig
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return "L'Id du serveur est incorrect!"
                End If

                If IO.File.Exists(Source) = False Then
                    Return "Le serveur n'a pas trouvé le fichier de configuration !"
                End If

                'sauvegarde de l'ancien fichier sous .old
                IO.File.Copy(_MonRepertoire & "\config\homidom.xml", _MonRepertoire & "\config\homidom.old", True)
                IO.File.Copy(Source, _MonRepertoire & "\config\homidom.xml", True)

                Return "0"
            Catch ex As Exception
                Return "Erreur lors de l'importation du fichier de config: " & ex.Message
            End Try
        End Function

        ''' <summary>Sauvegarder la configuration</summary>
        ''' <remarks></remarks>
        Public Function SaveConfiguration(ByVal IdSrv As String) As String Implements IHoMIDom.SaveConfig
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return "L'Id du serveur est incorrect"
                End If

                If SaveConfig(_MonRepertoire & "\config\homidom.xml") = True Then
                    Return "0"
                Else
                    Return "Erreur lors de l'enregistrement veuillez consulter le log"
                End If

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveConfiguration", "Exception : " & ex.Message)
                Return "Erreur lors de l'enregistrement veuillez consulter le log"
            End Try
        End Function
#End Region

#Region "SOAP"
        ''' <summary>Fixer la valeur du port SOAP</summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub SetPortSOAP(ByVal IdSrv As String, ByVal Value As Double) Implements IHoMIDom.SetPortSOAP
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                _PortSOAP = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetPortSOAP", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne la valeur du port SOAP</summary>
        ''' <returns>Numero du port ou -1 si erreur</returns>
        ''' <remarks></remarks>
        Public Function GetPortSOAP() As Double Implements IHoMIDom.GetPortSOAP
            Try
                Return _PortSOAP
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetPortSOAP", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>Fixer la valeur IP SOAP</summary>
        ''' <param name="Value"></param>
        ''' <remarks></remarks>
        Public Sub SetIPSOAP(ByVal IdSrv As String, ByVal Value As String) Implements IHoMIDom.SetIPSOAP
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Exit Sub
                End If

                _IPSOAP = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetIPSOAP", "Erreur: " & ex.Message)
            End Try
        End Sub

        ''' <summary>Retourne la valeur IP SOAP</summary>
        ''' <returns>Numero du port ou -1 si erreur</returns>
        ''' <remarks></remarks>
        Public Function GetIPSOAP() As String Implements IHoMIDom.GetIPSOAP
            Try
                Return _IPSOAP
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetIPSOAP", "Erreur: " & ex.Message)
                Return -1
            End Try
        End Function

#End Region

#Region "Maintenance"
        Public Sub CleanLog()
            Try

                Dim dirInfo As New System.IO.DirectoryInfo(_MonRepertoire & "\logs\")
                Dim files() As System.IO.FileInfo = dirInfo.GetFiles("*.txt", System.IO.SearchOption.AllDirectories)
                Dim DateRef As DateTime = Now.AddMonths(-1 * _MaxMonthLog)
                Dim cnt As Integer = 0

                If (files IsNot Nothing) Then
                    For Each file In files
                        If InStr(file.Name, "_") > 0 Then
                            Dim a() As String = file.Name.Split("_")
                            If a IsNot Nothing Then
                                If a.Count = 2 And a(0) = "log" Then
                                    Try
                                        If CDate(Mid(a(1), 1, 4) & "-" & Mid(a(1), 5, 2) & "-" & Mid(a(1), 7, 2)) < DateRef Then
                                            file.Delete()
                                            cnt += 1
                                        End If
                                    Catch ex As Exception
                                        Log(TypeLog.ERREUR, TypeSource.SERVEUR, "CleanLog", "  - Erreur: " & ex.Message)
                                    End Try
                                End If
                            End If
                        End If
                    Next
                End If
                Log(TypeLog.INFO, TypeSource.SERVEUR, "CleanLog", cnt & " Fichier(s) log supprimé(s) ( < " & _MaxMonthLog & " mois = " & DateRef.ToString("dd/MM/yyyy") & ")")
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "CleanLog", "Erreur: " & ex.Message)
            End Try
        End Sub
#End Region

#Region "Energie"
        ''' <summary>
        ''' True si on doit gérer l'energie
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SetGererEnergie(ByVal Value As Boolean) Implements IHoMIDom.SetGererEnergie
            Try
                GererEnergie = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetGererEnergie", "Erreur: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' True si on doit gérer l'energie
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetGererEnergie() As Boolean Implements IHoMIDom.GetGererEnergie
            Try
                Return GererEnergie
            Catch ex As Exception
                Return False
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetGererEnergie", "Erreur: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Set Tarif jour
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SetTarifJour(ByVal Value As Double) Implements IHoMIDom.SetTarifJour
            Try
                _TarifJour = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetTarifJour", "Erreur: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Get Tarif joure
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTarifJour() As Double Implements IHoMIDom.GetTarifJour
            Try
                Return _TarifJour
            Catch ex As Exception
                Return 0
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetTarifJour", "Erreur: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Set Tarif nuit
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SetTarifNuit(ByVal Value As Double) Implements IHoMIDom.SetTarifNuit
            Try
                _TarifNuit = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetTarifNuit", "Erreur: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Get Tarif nuit
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetTarifNuit() As Double Implements IHoMIDom.GetTarifNuit
            Try
                Return _TarifNuit
            Catch ex As Exception
                Return 0
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetTarifNuit", "Erreur: " & ex.Message)
            End Try
        End Function

        ''' <summary>
        ''' Set Puissance mini
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SetPuissanceMini(ByVal Value As Integer) Implements IHoMIDom.SetPuissanceMini
            Try
                PuissanceMini = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetPuissanceMini", "Erreur: " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Get Puissance Mini
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetPuissanceMini() As Integer Implements IHoMIDom.GetPuissanceMini
            Try
                Return PuissanceMini
            Catch ex As Exception
                Return 0
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetPuissanceMini", "Erreur: " & ex.Message)
            End Try
        End Function
#End Region

#Region "Decouverte"
        ''' <summary>
        ''' Retourne True/False si le mode découverte des nouveaux devices est activé
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetModeDecouverte() As Boolean Implements IHoMIDom.GetModeDecouverte
            Try
                Return _ModeDecouverte
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetModeDecouverte", "Exception : " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Fixe (True/False) si le mode découverte des nouveaux devices doit être activé
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SetModeDecouverte(ByVal Value As Boolean) Implements IHoMIDom.SetModeDecouverte
            Try
                _ModeDecouverte = Value
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetModeDecouverte", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Ajoute un nouveau device dans la liste découverte
        ''' </summary>
        ''' <param name="Adresse1"></param>
        ''' <param name="DriverId"></param>
        ''' <param name="Type"></param>
        ''' <param name="Adresse2"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function AddDetectNewDevice(ByVal Adresse1 As String, ByVal DriverId As String, Optional ByVal Type As String = "", Optional ByVal Adresse2 As String = "", Optional ByVal Value As String = "") As String Implements IHoMIDom.AddDetectNewDevice
            Try
                Dim flag As Boolean = False
                Dim _return As String = ""

                'get driver name
                Dim drivernom As String = DriverId
                Try
                    For i As Integer = 0 To _ListDrivers.Count - 1
                        If _ListDrivers.Item(i).ID = DriverId Then
                            drivernom = _ListDrivers.Item(i).nom
                            Exit For
                        End If
                    Next
                Catch ex As Exception

                End Try

                'check si le composant existe déjà (meme adresse1, adresse2 et driver_id)
                For Each _dev As NewDevice In _ListNewDevices
                    If _dev.Adresse1 = Adresse1 And _dev.IdDriver = DriverId And _dev.Adresse2 = Adresse2 And (Type = "" Or Type = _dev.Type) Then
                        _dev.DateTetect = Now
                        _dev.Value = Value
                        'If (Type = "" Or _dev.Type = Type) Then
                        Log(TypeLog.DEBUG, TypeSource.SERVEUR, "AddDetectNewDevice", "Le composant du driver " & drivernom & " existe déjà : " & Adresse1 & ":" & Adresse2)
                        'End If
                        flag = True
                        _return = "Message: Composant déjà existant"
                        Exit For
                    End If
                Next

                If flag = False Then
                    Dim x As New NewDevice
                    x.ID = System.Guid.NewGuid.ToString()
                    x.Adresse1 = Adresse1
                    x.Adresse2 = Adresse2
                    x.IdDriver = DriverId
                    x.Type = Type
                    x.Ignore = False
                    x.DateTetect = Now
                    x.Value = Value
                    x.Name = "NouveauComposant" & _ListNewDevices.Count

                    _ListNewDevices.Add(x)
                    x = Nothing
                    Log(TypeLog.DEBUG, TypeSource.SERVEUR, "AddDetectNewDevice", "Le composant du driver " & drivernom & " a été ajouté : " & Type & " - " & Adresse1 & ":" & Adresse2)
                    _return = 0
                End If

                Return _return
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "AddDetectNewDevice", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

        ''' <summary>
        ''' Retourne un NewDevice suivant son ID
        ''' </summary>
        ''' <param name="Id"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ReturnNewDevice(ByVal Id As String) As NewDevice Implements IHoMIDom.ReturnNewDevice
            Try
                If (From Dev As NewDevice In _ListNewDevices Where Dev.ID = Id Select Dev).Count > 0 Then
                    Dim Resultat As NewDevice = (From Dev In _ListNewDevices Where Dev.ID = Id Select Dev).First
                    Return Resultat
                Else
                    Return Nothing
                End If
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnNewDevice", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Set un NewDevice suivant son ID
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub SaveNewDevice(ByVal NewDevice As NewDevice) Implements IHoMIDom.SaveNewDevice
            Try

                For Each _dev As NewDevice In _ListNewDevices
                    If _dev.ID = NewDevice.ID Then
                        _dev.Adresse1 = NewDevice.Adresse1
                        _dev.Adresse2 = NewDevice.Adresse2
                        _dev.DateTetect = NewDevice.DateTetect
                        _dev.ID = NewDevice.ID
                        _dev.IdDriver = NewDevice.IdDriver
                        _dev.Name = NewDevice.Name
                        _dev.Ignore = NewDevice.Ignore
                        _dev.Type = NewDevice.Type
                        _dev.Value = NewDevice.Value
                        Exit For
                    End If
                Next

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveNewDevice", "Exception : " & ex.Message)
            End Try
        End Sub

        ''' <summary>
        ''' Retourne tous les nouveaux devices détectés
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAllNewDevice() As List(Of NewDevice) Implements IHoMIDom.GetAllNewDevice
            Try
                Return _ListNewDevices
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllNewDevice", "Exception : " & ex.Message)
                Return Nothing
            End Try
        End Function

        ''' <summary>
        ''' Existe t-il des nouveaux devices non ignorés?
        ''' </summary>
        ''' <returns>True si nouveau device</returns>
        ''' <remarks></remarks>
        Public Function AsNewDevice() As Boolean Implements IHoMIDom.AsNewDevice
            Try
                Dim flag As Boolean = False

                For Each _dev As NewDevice In _ListNewDevices
                    If _dev.Ignore = False Then flag = True
                Next

                Return flag
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "AsNewDevice", "Exception : " & ex.Message)
                Return False
            End Try
        End Function

        ''' <summary>
        ''' Supprimer un newdevice de la liste
        ''' </summary>
        ''' <param name="IdSrv"></param>
        ''' <param name="NewDeviceId"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function DeleteNewDevice(ByVal IdSrv As String, ByVal NewDeviceId As String) As Integer Implements IHoMIDom.DeleteNewDevice
            Try
                If VerifIdSrv(IdSrv) = False Then
                    Return 99
                End If

                For i As Integer = 0 To _ListNewDevices.Count - 1
                    If _ListNewDevices.Item(i).ID = NewDeviceId Then
                        _ListNewDevices.RemoveAt(i)

                        Return 0
                    End If
                Next

                Return -1
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteNewDevice", "Exception : " & ex.Message)
                Return -1
            End Try
        End Function

#End Region

#Region "Variable"
        Public Function GetAllVariables(idsrv As String) As List(Of Variable) Implements IHoMIDom.GetAllVariables
            Try
                If VerifIdSrv(idsrv) = False Then
                    Return Nothing
                End If

                Dim _list As New List(Of Variable)
                _list = _ListVars

                Return _list
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetAllVariables: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function AddVariable(ByVal idsrv As String, Name As String, Optional Enable As Boolean = True, Optional Value As String = "", Optional Description As String = "") As String Implements IHoMIDom.AddVariable
            Try
                If VerifIdSrv(idsrv) = False Then
                    Return "ID du serveur erroné!!"
                End If

                If String.IsNullOrEmpty(Name) Then
                    Return "Erreur le nom de la variable doit être définit!!"
                End If

                For Each _var In _ListVars
                    If _var.Nom.ToLower = Name.ToLower Then
                        Return "Erreur cette variable existe déjà!!"
                    End If
                Next

                Dim _newvar As New Variable
                With _newvar
                    .ID = GenerateGUID()
                    .Nom = Name
                    .Enable = Enable
                    .Value = Value
                    .Description = Description
                End With
                _ListVars.Add(_newvar)

                Return Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "AddVariable: ", ex.ToString)
                Return ex.Message
            End Try
        End Function

        Public Function SaveVariable(ByVal idsrv As String, Name As String, Optional Enable As Boolean = True, Optional Value As String = "", Optional Description As String = "") As String Implements IHoMIDom.SaveVariable
            Try
                If VerifIdSrv(idsrv) = False Then
                    Return "ID du serveur erroné!!"
                End If

                If String.IsNullOrEmpty(Name) Then
                    Return "Erreur le nom de la variable doit être définit!!"
                End If

                Dim trv As Boolean

                For Each _var In _ListVars
                    If _var.Nom.ToLower = Name.ToLower Then
                        _var.Enable = Enable
                        _var.Description = Description
                        _var.Value = Value
                        trv = True
                        Exit For
                    End If
                Next

                If trv Then
                    Return Nothing
                Else
                    Return "La variable " & Name & " n'a pas pu être modifiée car non trouvée!!"
                End If

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SaveVariable: ", ex.Message)
                Return ex.Message
            End Try
        End Function


        Public Function GetValueOfVariable(ByVal idsrv As String, Name As String) As String Implements IHoMIDom.GetValueOfVariable
            Try
                If VerifIdSrv(idsrv) = False Then
                    Return Nothing
                End If

                For Each _var In _ListVars
                    If _var.Nom.ToLower = Name.ToLower Then
                        Return _var.Value
                    End If
                Next

                Return Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetValueOfVariable: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function SetValueOfVariable(ByVal idsrv As String, Name As String, Value As String) As String Implements IHoMIDom.SetValueOfVariable
            Try
                If VerifIdSrv(idsrv) = False Then
                    Return 99
                End If

                For Each _var In _ListVars
                    If _var.Nom.ToLower = Name.ToLower Then
                        _var.Value = Value
                    End If
                Next

                Return Nothing
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "SetValueOfVariable: ", ex.Message)
                Return ex.Message
            End Try
        End Function

        Public Function DeleteVariable(ByVal idsrv As String, Name As String) As String Implements IHoMIDom.DeleteVariable
            Try
                If VerifIdSrv(idsrv) = False Then
                    Return "ID du serveur erroné!!"
                End If

                Dim i As Integer = 0
                Dim trv As Boolean

                For Each _var In _ListVars
                    If _var.Nom.ToLower = Name.ToLower Then
                        _ListVars.RemoveAt(i)
                        trv = True
                        Exit For
                    End If
                    i += 1
                Next

                If trv Then
                    Return Nothing
                Else
                    Return "La variable " & Name & " n'a pas pu être supprimée car non trouvée!!"
                End If

            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "DeleteVariable: ", ex.Message)
                Return ex.Message
            End Try
        End Function
#End Region

#Region "Sequences"

        Public Function ReturnSequences() As List(Of Sequence) Implements IHoMIDom.ReturnSequences
            Try
                Return ManagerSequences.Sequences
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnSequences: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function ReturnSequenceFromNumero(Numero As String) As Sequence Implements IHoMIDom.ReturnSequenceFromNumero
            Try
                Return ManagerSequences.ReturnSequenceFromNumero(Numero)
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "ReturnSequenceFromNumero: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetSequenceDriver() As String Implements IHoMIDom.GetSequenceDriver
            Try
                Return ManagerSequences.SequenceDriver
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSequenceDriver: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetSequenceDevice() As String Implements IHoMIDom.GetSequenceDevice
            Try
                Return ManagerSequences.SequenceDevice
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSequenceDevice: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetSequenceTrigger() As String Implements IHoMIDom.GetSequenceTrigger
            Try
                Return ManagerSequences.SequenceTrigger
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSequenceTrigger: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetSequenceZone() As String Implements IHoMIDom.GetSequenceZone
            Try
                Return ManagerSequences.SequenceZone
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSequenceZone: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetSequenceMacro() As String Implements IHoMIDom.GetSequenceMacro
            Try
                Return ManagerSequences.SequenceMacro
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSequenceMacro: ", ex.Message)
                Return Nothing
            End Try
        End Function

        Public Function GetSequenceServer() As String Implements IHoMIDom.GetSequenceServer
            Try
                Return ManagerSequences.SequenceServer
            Catch ex As Exception
                Log(TypeLog.ERREUR, TypeSource.SERVEUR, "GetSequenceServer: ", ex.Message)
                Return Nothing
            End Try
        End Function
#End Region

#End Region

    End Class

End Namespace