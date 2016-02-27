Imports HoMIDom
Imports HoMIDom.HoMIDom.Server
Imports HoMIDom.HoMIDom.Device
Imports System.IO
Imports System.Xml
Imports OfficeOpenXml

Imports STRGS = Microsoft.VisualBasic.Strings


'Imports System.Reflection
'Imports System.Runtime.CompilerServices


' Auteur : Jptools
' Date : 8/02/2015

''' <summary>Driver Excel</summary>
''' <remarks></remarks>
<Serializable()> Public Class Driver_GestionChauffage
    Implements HoMIDom.HoMIDom.IDriver

#Region "Variable Driver"
    '!!!Attention les variables ci-dessous doivent avoir une valeur par défaut obligatoirement
    'aller sur l'adresse http://www.somacon.com/p113.php pour avoir un ID
    Dim _ID As String = "3BAB5794-CBEE-11E4-A3B7-91811E5D46B0"
    Dim _Nom As String = "GestionChauffage"
    Dim _Enable As Boolean = False
    Dim _Description As String = "Driver GestionChauffage"
    Dim _StartAuto As Boolean = False
    Dim _Protocol As String = "Fichier"
    Dim _IsConnect As Boolean = False
    Dim _IP_TCP As String = "@"
    Dim _Port_TCP As String = "@"
    Dim _IP_UDP As String = "@"
    Dim _Port_UDP As String = "@"
    Dim _Com As String = "@"
    Dim _Refresh As Integer = 0
    Dim _Modele As String = "EPPlus"
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

    'parametres avancés
    Dim _DEBUG As Boolean = False

    'A ajouter dans les ppt du driver
    Dim _tempsentrereponse As Integer = 1500
    Dim _ignoreadresse As Boolean = False
    Dim _lastetat As Boolean = True
#End Region

#Region "Variables internes"


#End Region

#Region "Declaration"
    Dim ModeChauffage As String = ""
    Dim ModeSemaine As String = ""

#End Region

#Region "Fonctions génériques"
    Shared GestionTemp As Object

    ''' <summary>
    ''' Retourne la liste des Commandes avancées
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetCommandPlus() As List(Of DeviceCommande)
        Return _DeviceCommandPlus
    End Function

    ''' <summary>Execute une commande avancée</summary>
    ''' <param name="MyDevice">Objet représentant le Device </param>
    ''' <param name="Command">Nom de la commande avancée à éxécuter</param>
    ''' <param name="Param">tableau de paramétres</param>
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
                    WriteLog("DBG: ExecuteCommandXX1 : " & Command)
                    Select Case UCase(Command)
                        Case "ON", "OFF"
                            Write(MyDevice, Command, Param(0), Param(1))
                            WriteLog("DBG: ExecuteCommandXX2 : " & Command)
                        Case Else

                    End Select
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


    Public WriteOnly Property IdSrv As String Implements HoMIDom.HoMIDom.IDriver.IdSrv
        Set(ByVal value As String)
            _IdSrv = value
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

    Public ReadOnly Property DeviceSupport() As System.Collections.ArrayList Implements HoMIDom.HoMIDom.IDriver.DeviceSupport
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

    Public Sub Read(ByVal Objet As Object) Implements HoMIDom.HoMIDom.IDriver.Read
        Try
            If _Enable = False Then Exit Sub

            Select Case Objet.Type

                Case "GENERIQUESTRING"
                    Select Case Objet.adresse1.toUpper

                        Case "NUMEROSEMAINE"
                            Objet.Value = LireNumSemaine(Now)          ' Rechercher le numero de la semaine

                        Case "NUMEROJOUR"
                            ' Objet.Value = Weekday(Now, vbMonday)     'déterminer le num du jour 
                            Objet.Value = DatePart(DateInterval.DayOfYear, Now, FirstDayOfWeek.Monday)  'déterminer le num du jour de l'année

                        Case "MODESEMAINE"
                            Objet.Value = ModeSemaine        'Affiche le type de semaine

                        Case "MODECHAUFFAGE"               'Affiche le mode de chauffage en cours
                            If ModeChauffage = "ECC" Then
                                Objet.Value = "Economique"
                            ElseIf ModeChauffage = "EC" Then
                                Objet.Value = "1/2 Economique"
                            ElseIf ModeChauffage = "PRE" Then
                                Objet.Value = "Confort"
                            ElseIf ModeChauffage = "PRE+" Then
                                Objet.Value = "Confort +"
                            End If

                        Case Else
                            WriteLog("ERR: Nom Valeur n'existe pas")
                    End Select

                Case "SWITCH"
                    WriteLog(" Mode : " & Objet.Name & " : " & Objet.Value)

                Case "GENERIQUEVALUE"
                    Select Case Objet.adresse1.toUpper
                        Case "TEMPERATURECONSIGNE"
                            If _Parametres.Item(6).Valeur = False Then        ' Si pas Hors Gele
                                If _Parametres.Item(5).Valeur = True Then      'Si on ne désire pas lire le fichier (on garde l'ancienne consigne)
                                    LireFichier(Objet)          'Lire la consigne de température
                                End If
                            Else
                                Objet.Value = _Parametres.Item(7).Valeur    'Valeur Consigne Hors Gele
                            End If
                    End Select
                Case Else
                    WriteLog("ERR: Type non conforme")
            End Select

        Catch ex As Exception
            WriteLog("ERR: Lecture des données" & ex.ToString)
        End Try

    End Sub

    Public Property Refresh() As Integer Implements HoMIDom.HoMIDom.IDriver.Refresh
        Get
            Return _Refresh
        End Get
        Set(ByVal value As Integer)
            _Refresh = value
        End Set
    End Property

    Public Sub Restart() Implements HoMIDom.HoMIDom.IDriver.Restart
        [Stop]()
        Start()
    End Sub

    Public Property Server() As HoMIDom.HoMIDom.Server Implements HoMIDom.HoMIDom.IDriver.Server
        Get
            Return _Server
        End Get
        Set(ByVal value As HoMIDom.HoMIDom.Server)
            _Server = value
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

    ''' <summary>
    ''' Permet de vérifier si un champ est valide
    ''' </summary>
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

    Public Sub Start() Implements HoMIDom.HoMIDom.IDriver.Start
        Try
            _DEBUG = _Parametres.Item(0).Valeur

            If TesterFichier() = True Then
                _IsConnect = True
                WriteLog("Driver démarré")
            Else
                _IsConnect = False
                WriteLog("ERR: Driver non démarré car le fichier n'est pas trouvé")
            End If
        Catch ex As Exception
            _IsConnect = False
            WriteLog("ERR: Driver en erreur lors du démarrage: " & ex.Message)
        End Try
    End Sub

    Public Property StartAuto() As Boolean Implements HoMIDom.HoMIDom.IDriver.StartAuto
        Get
            Return _StartAuto
        End Get
        Set(ByVal value As Boolean)
            _StartAuto = value
        End Set
    End Property

    Public Sub [Stop]() Implements HoMIDom.HoMIDom.IDriver.Stop
        Try
            _IsConnect = False
            WriteLog("Driver arrêté")
        Catch ex As Exception
            WriteLog("ERR: Erreur lors de l'arrêt du Driver: " & ex.Message)
        End Try
    End Sub

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

    Public Sub Write(ByVal Objet As Object, ByVal Commande As String, Optional ByVal Parametre1 As Object = Nothing, Optional ByVal Parametre2 As Object = Nothing) Implements HoMIDom.HoMIDom.IDriver.Write
        Try
            If _Enable = False Then Exit Sub

            If Objet.type = "SWITCH" Then
                WriteLog("DBG:  WriteXX2" & Objet.Type & " : " & Commande & " sur le noeud : " & Objet.Adresse1.ToString)
                Select Case UCase(Commande)
                    Case "ON"          'active ou désactive la lecture du fichier Excel
                        If Objet.adresse1.toUpper = "ACTIVERLECTURE" Then
                            _Parametres.Item(5).Valeur = True
                            Objet.Value = 100

                        ElseIf Objet.adresse1.toUpper = "ACTIVERHORSGELE" Then
                            _Parametres.Item(6).Valeur = True
                            Objet.Value = 100
                        End If
                        WriteLog("DBG: WriteXX3" & Objet.Type & " : " & Commande & " sur le noeud : " & Objet.Adresse1.ToString)
                    Case "OFF"
                        If Objet.adresse1.toUpper = "ACTIVERLECTURE" Then
                            _Parametres.Item(5).Valeur = False
                            Objet.Value = 0

                        ElseIf Objet.adresse1.toUpper = "ACTIVERHORSGELE" Then
                            _Parametres.Item(6).Valeur = False
                            Objet.Value = 0
                        End If
                        WriteLog("DBG: WriteXX4" & Objet.Type & " : " & Commande & " sur le noeud : " & Objet.Adresse1.ToString)
                    Case Else
                        WriteLog("ERR: Erreur la commande du device n'est pas supporté par ce driver - device: " & Objet.name & " commande:" & Commande)
                End Select
            Else
                WriteLog("ERR:Erreur le type de device n'est pas supporté par ce driver - device: " & Objet.name & " type:" & Objet.type.ToString)
            End If

        Catch ex As Exception
            WriteLog("ERR: Erreur lors de la traitement de la commande - device: " & Objet.name & " commande:" & Commande & " erreur: " & ex.Message)
        End Try
    End Sub

    Public Sub DeleteDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.DeleteDevice

    End Sub

    Public Sub NewDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.NewDevice

    End Sub

    ''' <summary>ajout des commandes avancées pour les devices</summary>
    ''' <param name="nom">Nom de la commande avancée</param>
    ''' <param name="description">Description qui sera affichée dans l'admin</param>
    ''' <param name="nbparam">Nombre de parametres attendus</param>
    ''' <remarks></remarks>
    Private Sub Add_DeviceCommande(ByVal Nom As String, ByVal Description As String, ByVal NbParam As Integer)
        Try
            Dim x As New DeviceCommande
            x.NameCommand = Nom
            x.DescriptionCommand = Description
            x.CountParam = NbParam
            _DeviceCommandPlus.Add(x)
        Catch ex As Exception
            WriteLog("ERR: add_devicecommande : Exception : " & ex.Message)
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
            WriteLog("ERR: add_devicecommande : Exception : " & ex.Message)
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
            WriteLog("ERR: add_devicecommande : Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>ajout de parametre avancés</summary>
    ''' <param name="nom">Nom du parametre (sans espace)</param>
    ''' <param name="description">Description du parametre</param>
    ''' <param name="valeur">Sa valeur</param>
    ''' <remarks></remarks>
    Private Sub Add_ParamAvance(ByVal nom As String, ByVal description As String, ByVal valeur As Object)
        Try
            Dim x As New HoMIDom.HoMIDom.Driver.Parametre
            x.Nom = nom
            x.Description = description
            x.Valeur = valeur
            _Parametres.Add(x)
        Catch ex As Exception
            WriteLog("ERR: add_devicecommande : Exception : " & ex.Message)
        End Try
    End Sub

    Public Sub New()
        Try
            _Version = Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString

            'liste des devices compatibles
            _DeviceSupport.Add(ListeDevices.GENERIQUESTRING)
            _DeviceSupport.Add(ListeDevices.SWITCH.ToString)
            _DeviceSupport.Add(ListeDevices.GENERIQUEVALUE)

            'Parametres avancés
            Add_ParamAvance("Debug", "Activer le Debug complet (True/False)", False)
            Add_ParamAvance("Fichier Excel", "Emplacement du fichier", "C:\Program Files\HoMIDoM\Config\Calendrier.xlsx")
            Add_ParamAvance("Ecconomique", "Température du mode ecconomique : ECC", 18)
            Add_ParamAvance("DemiEcconomique", "Température du mode 1/2 Ecconomique : EC", 19)
            Add_ParamAvance("Confort", "Température du mode Confort : PRE", 20)
            Add_ParamAvance("ActiverLecture", "Activer la lecture (True/False)", True)
            Add_ParamAvance("ActiverHorsGele", "Activer le Mode Hors Gêle (True/False)", False)
            Add_ParamAvance("Hors Gêle", "Température du mode Hors Gêle", 8)
            Add_ParamAvance("Confort +", "Température du mode Confort + : PRE+", 21)

            'ajout des commandes avancées pour les devices          
            'Add_DeviceCommande("COMMANDE", "DESCRIPTION", nbparametre)
            'Add_DeviceCommande("Tester fichier", "Teste la présence du fichier", 0)
            'Add_DeviceCommande("RECEIVE", "Recevoir un Message", 1)

            'Libellé Driver
            Add_LibelleDriver("HELP", "Aide...", "Pas d'aide actuellement...")

            'Libellé Device
            Add_LibelleDevice("ADRESSE1", "Nom Valeur", "Nom pour TemperatureConsigne, NumeroSemaine et NumeroJour")
            Add_LibelleDevice("ADRESSE2", "@", "")
            Add_LibelleDevice("SOLO", "@", "")
            Add_LibelleDevice("REFRESH", "Refresh", "")
            Add_LibelleDevice("LASTCHANGEDUREE", "LastChange Durée", "")

        Catch ex As Exception
            WriteLog("ERR: New : " & ex.Message)
        End Try
    End Sub
#End Region


    '---------------------------------------------------------------------------------------------------------
#Region "Fonctions internes"

    'Insérer ci-dessous les fonctions propres au driver et nom communes (ex: start)
    'Ajout d'une fonction spéciale
    Public Sub TesterPresenceFichier()
        WriteLog("DBG: Tester la présence du fichier : Calendrier.xlsx")
        Try
            If TesterFichier() = True Then
                WriteLog("Le fichier Calendrier.xlsx est trouvé")
            Else
                WriteLog("ERR: Le fichier Calendrier.xlsx n'est pas trouvé")
            End If
        Catch ex As Exception
            WriteLog("ERR: Tester présence du fichier, Exception : " & ex.Message)
        End Try
    End Sub


    Public Sub SauverCopieFichier()

        Dim zipFile, newFile As FileInfo
        Dim NomFichier As String

        WriteLog("DBG: Créer une copie du fichier : Calendrier.xlsx")
        Try
            'Sauver une copie du fichier Excel "Calendrier.bak"
            newFile = New FileInfo(_Parametres.Item(1).Valeur) 'Ouverture du fichier Excel
            If newFile.Exists Then
                NomFichier = _Parametres.Item(1).Valeur
                Dim ParaFichier = Split(NomFichier, ".")
                ParaFichier(1) = "bak"
                zipFile = New FileInfo(ParaFichier(0) + "." + ParaFichier(1))
                If (zipFile.Exists) Then
                    zipFile.Delete()
                End If
                newFile.CopyTo(zipFile.FullName)
                WriteLog("Le fichier Calendrier.bak est créé")
            Else
                WriteLog("ERR: Le fichier Calendrier.xlsx n'est pas trouvé")
            End If
        Catch ex As Exception
            WriteLog("ERR: Créer une copie, Exception : " & ex.Message)
        End Try
    End Sub


    Public Function TesterFichier() As Boolean

        Try
            Dim newFile As FileInfo
            Dim worksheet As ExcelWorksheet
            Dim pck As ExcelPackage
            Dim Mode As String
            Dim CellValue As String = ""

            newFile = New FileInfo(_Parametres.Item(1).Valeur)
            If newFile.Exists Then
                pck = New ExcelPackage(newFile)
                Mode = "Calendrier"
                worksheet = pck.Workbook.Worksheets.Item(Mode)
                CellValue = worksheet.Cells(1, 1).Value         'Lire HOMIDOM
                WriteLog("La valeur recherchée : " & CellValue)
            Else
                WriteLog("ERR: Lecture Fichier :  non trouvé")
            End If

            If CellValue = "HoMIDoM" Then     'Vérifier si le fichier est compatible
                Return 1
            Else
                Return 0
            End If
        Catch ex As Exception
            WriteLog("ERR: Lecture Fichier" & ex.ToString)
            Return 0
        End Try
    End Function


    'Permet de lire le fichier 
    Public Function LireFichier(ByVal Objet As Object) As Boolean

        Try

            Dim newFile As FileInfo
            Dim worksheet As ExcelWorksheet
            Dim pck As ExcelPackage
            Dim Mode As String

            Dim NumJour As Integer
            Dim Ligne, NumSemaine As Integer

            NumSemaine = LireNumSemaine(Now)          ' Rechercher le numero de la semaine
            If NumSemaine = 0 Then
                WriteLog("ERR: Numéro de semaine erroné")
                Return 0             'Erreur du numero de semaine
            End If

            'Lecture du mode de Chauffage
            newFile = New FileInfo(_Parametres.Item(1).Valeur)
            If newFile.Exists Then
                pck = New ExcelPackage(newFile)
                Mode = "Calendrier"
                worksheet = pck.Workbook.Worksheets.Item(Mode)
                ModeSemaine = worksheet.Cells(NumSemaine + 1, 2).Value   'Recherche le mode  de chauffage          
                NumJour = Weekday(Now, vbMonday)     'déterminer le num du jour       

                'Num ligne a calculer en fonction de l'Heure (par tranche 1/2h)
                Ligne = ((Timer / 3600) * 2) + 1
                WriteLog(" Mode: " & ModeSemaine & " N° Lignes: " & CStr(Ligne + 1))
                If ModeSemaine = "Conger" Or ModeSemaine = "Normal" Or ModeSemaine = "Reduit" Or ModeSemaine = "Charger" Or ModeSemaine = "Absence" Then
                    worksheet = pck.Workbook.Worksheets.Item(ModeSemaine)
                    ModeChauffage = worksheet.Cells(Ligne + 1, NumJour + 1).Value
                    WriteLog(" Mode : " & ModeChauffage & " N° Jour: " & CStr(NumJour) & " N° Sem: " & CStr(NumSemaine))

                    Select Case ModeChauffage
                        Case "ECC"
                            Objet.Value = _Parametres.Item(2).Valeur
                        Case "EC"
                            Objet.Value = _Parametres.Item(3).Valeur
                        Case "PRE"
                            Objet.Value = _Parametres.Item(4).Valeur
                        Case "PRE+"
                            Objet.Value = _Parametres.Item(8).Valeur
                        Case Else
                            WriteLog("ERR: Erreur Mode Chauffage")
                            Objet.Value = 18
                    End Select
                Else
                    WriteLog("ERR: Erreur Mode Chauffage")
                    Objet.Value = 18
                End If
                Return 1
            Else
                WriteLog("ERR: Fichier non trouvé")
                Return 0
            End If
        Catch ex As Exception
            WriteLog("ERR: Lecture Fichier" & ex.ToString)
            Return 0
        End Try
    End Function


    'Calcul le numero de la semaine
    Public Function LireNumSemaine(ByVal dat As Date) As Integer
        If IsDate(dat) Then
            Dim semaine As Integer
            Dim semain As Integer

            semain = Weekday(dat)
            If semain = 2 Then
                dat = dat.AddDays(6)
            End If
            If semain = 3 Then
                dat = dat.AddDays(5)
            End If
            If semain = 4 Then
                dat = dat.AddDays(4)
            End If
            If semain = 5 Then
                dat = dat.AddDays(3)
            End If
            If semain = 6 Then
                dat = dat.AddDays(2)
            End If
            If semain = 7 Then
                dat = dat.AddDays(1)
            End If
            semaine = DatePart("ww", dat, vbMonday, vbFirstFourDays)

            Return semaine
        End If
        Return Nothing
    End Function

    Private Sub WriteLog(ByVal message As String)
        Try
            'utilise la fonction de base pour loguer un event
            If STRGS.InStr(message, "DBG:") > 0 Then
                If _DEBUG Then
                    _Server.Log(TypeLog.DEBUG, TypeSource.DRIVER, Me.Nom, STRGS.Right(message, message.Length - 5))
                End If
            ElseIf STRGS.InStr(message, "ERR:") > 0 Then
                _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom, STRGS.Right(message, message.Length - 5))
            Else
                _Server.Log(TypeLog.INFO, TypeSource.DRIVER, Me.Nom, message)
            End If
        Catch ex As Exception
            _Server.Log(TypeLog.ERREUR, TypeSource.DRIVER, Me.Nom & " WriteLog", ex.Message)
        End Try
    End Sub

#End Region


End Class


