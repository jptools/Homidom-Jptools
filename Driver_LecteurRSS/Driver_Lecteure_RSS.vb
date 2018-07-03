Imports HoMIDom
Imports HoMIDom.HoMIDom.Server
Imports HoMIDom.HoMIDom.Device
Imports System.Net
Imports System.IO
Imports System.Text

Imports System.Text.RegularExpressions
Imports STRGS = Microsoft.VisualBasic.Strings
Imports System.Xml
Imports System.Web


' Auteur : jptools
' Date : 06/06/2018

''' <summary>Driver lecteur RSS (Alsace, Le Monde etc...)</summary>
''' <remarks></remarks>
<Serializable()> Public Class Driver_Lecteur_RSS
    Implements HoMIDom.HoMIDom.IDriver

#Region "Variables génériques"
    '!!!Attention les variables ci-dessous doivent avoir une valeur par défaut obligatoirement
    'aller sur l'adresse http://www.somacon.com/p113.php pour avoir un ID
    Dim _ID As String = "6E4C3C4E-6962-11E8-90CB-122A19C6383E"
    Dim _Nom As String = "Lecteur_RSS"
    Dim _Enable As Boolean = False
    Dim _Description As String = "Lecteur de données rss"
    Dim _StartAuto As Boolean = False
    Dim _Protocol As String = "WEB"
    Dim _IsConnect As Boolean = False
    Dim _IP_TCP As String = "@"
    Dim _Port_TCP As String = "@"
    Dim _IP_UDP As String = "@"
    Dim _Port_UDP As String = "@"
    Dim _Com As String = "@"
    Dim _Refresh As Integer = 0
    Dim _Modele As String = "Lecteur_rss"
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

    'A ajouter dans les ppt du driver
    ' Dim _tempsentrereponse As Integer = 1500
    ' Dim _ignoreadresse As Boolean = False
    ' Dim _lastetat As Boolean = True

    'param avancé
    Dim _DEBUG As Boolean = False
    Dim _NbLIGNES As Integer = 10
    Dim _NbCARACT As Integer = 50
    Dim _SAUTLIGNE As Boolean = False
    Dim _AFFDESCRIPT As Boolean = False
    Dim _AFFTITRES As Boolean = True
    Dim _LIGNESEPARAT As Boolean = False

#End Region

#Region "Variables internes"

    'cree un hashtable avec toutes les possibilités
    Dim NametoNum As New Hashtable

    Public Sub GetCodeHtml()
        Try
            NametoNum.Clear()
            NametoNum.Add("&quot;", "34")
            NametoNum.Add("&num;", "35")
            NametoNum.Add("&apos;", "39")
            NametoNum.Add("&#39;", "39")
            NametoNum.Add("&amp;", "38")
            NametoNum.Add("&lt;", "60")
            NametoNum.Add("&rlt;", "62")
            NametoNum.Add("&lsqb;", "91")
            NametoNum.Add("&rsqb;", "92")
            NametoNum.Add("&grave;", "96")
            NametoNum.Add("&cub;", "123")
            NametoNum.Add("&rcub;", "125")
            NametoNum.Add("&euro;", "128")
            NametoNum.Add("&nbsp;", "160")
            NametoNum.Add("&laquo;", "171")
            NametoNum.Add("&acute;", "180")
            NametoNum.Add("&agrave;", "224")
            NametoNum.Add("&raquo;", "187")
            NametoNum.Add("&Agrave;", "192")
            NametoNum.Add("&Egrave;", "200")
            NametoNum.Add("&Eacute;", "201")
            NametoNum.Add("&Igrave;", "204")
            NametoNum.Add("&Ograve;", "210")
            NametoNum.Add("&Ugrave;", "217")
            NametoNum.Add("&aacute;", "225")
            NametoNum.Add("&acirc;", "226")
            NametoNum.Add("&ccedil;", "231")
            NametoNum.Add("&egrave;", "232")
            NametoNum.Add("&eacute;", "233")
            NametoNum.Add("&ecirc;", "234")
            NametoNum.Add("&euml;", "235")
            NametoNum.Add("&igrave;", "236")
            NametoNum.Add("&iacute;", "237")
            NametoNum.Add("&icirc;", "238")
            NametoNum.Add("&iuml;", "239")
            NametoNum.Add("&ocirc;", "244")
            NametoNum.Add("&ouml;", "246")
            NametoNum.Add("&oslash;", "248")
            NametoNum.Add("&ugrave;", "249")
            NametoNum.Add("&ucirc;", "251")
            NametoNum.Add("&uuml;", "252")
            NametoNum.Add("&Scaron;", "352")
            NametoNum.Add("&Yuml;", "376")
            NametoNum.Add("&lsquo;", "8216")
            NametoNum.Add("&rsquo;", "39") '8217
            NametoNum.Add("&sbquo;", "8218")
            NametoNum.Add("&dagger;", "8224")
            NametoNum.Add("&Dagger;", "8225")
            NametoNum.Add("&hellip;", "8230")
            NametoNum.Add("&rsaquo;", "8250")
            WriteLog("DBG: GetCodeHtml, " & NametoNum.Count & " Code HTML chargés")
        Catch ex As Exception
            WriteLog("ERR: GetCodeHtml, Exception : " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Propriétés génériques"
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
    Public Property Refresh() As Integer Implements HoMIDom.HoMIDom.IDriver.Refresh
        Get
            Return _Refresh
        End Get
        Set(ByVal value As Integer)
            _Refresh = value
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
                    'Write(deviceobject, Command, Param(0), Param(1))
                    Select Case UCase(Command)
                        Case ""
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
                Case "ADRESSE1"
                    If Value IsNot Nothing Then
                        If String.IsNullOrEmpty(Value) Then
                            retour = vbCrLf & "Veuillez saisir le site web" & vbCrLf & "ex : LeMonde, LeFigaro, InfoWeb1, InfoWeb2, ..."
                        End If
                        Dim Temp As String = UCase(Value)
                        Select Case True
                            Case Temp = "LEMONDE"
                            Case Temp = "LEFIGARO"
                            Case Temp = "INFOWEB1"
                            Case Temp = "INFOWEB2"
                            Case Temp = "INFOWEB3"
                            Case Temp = "INFOWEB4"
                            Case Temp = "INFOWEB5"
                            Case Else
                                retour = vbCrLf & "Veuillez saisir le site web" & vbCrLf & "ex : LeMonde, LeFigaro, InfoWeb1, InfoWeb2, ..."
                        End Select
                    End If
                Case "ADRESSE2"
                    'Tester Adresse2 si valide
            End Select
            Return retour
        Catch ex As Exception
            Return "Une erreur est apparue lors de la vérification du champ " & Champ & ": " & ex.ToString
        End Try
    End Function

    ''' <summary>Démarrer le driver</summary>
    ''' <remarks></remarks>
    Public Sub Start() Implements HoMIDom.HoMIDom.IDriver.Start
        Try
            'récupération des paramétres avancés
            Try
                _DEBUG = _Parametres.Item(0).Valeur
                _NbLIGNES = _Parametres.Item(1).Valeur
                _NbCARACT = _Parametres.Item(7).Valeur
                _SAUTLIGNE = _Parametres.Item(8).Valeur
                _Description = _Parametres.Item(9).Valeur
                _AFFTITRES = _Parametres.Item(10).Valeur
                _LIGNESEPARAT = _Parametres.Item(11).Valeur
            Catch ex As Exception
                _DEBUG = False
                _Parametres.Item(0).Valeur = False
                WriteLog("ERR: Erreur dans les paramétres avancés. utilisation des valeur par défaut : " & ex.Message)
            End Try

            'Si internet n'est pas disponible on ne mets pas à jour les informations
            If My.Computer.Network.IsAvailable = False Then
                WriteLog("ERR: Pas de connection réseau, une connexion sera nécessaire : ")
            End If
            _IsConnect = True

            'charge les codes Html
            GetCodeHtml()

        Catch ex As Exception
            _IsConnect = False
            WriteLog("ERR: Driver " & Me.Nom & " Erreur démarrage " & ex.Message)
        End Try
    End Sub

    ''' <summary>Arrêter le du driver</summary>
    ''' <remarks></remarks>
    Public Sub [Stop]() Implements HoMIDom.HoMIDom.IDriver.Stop
        Try
            _IsConnect = False
            NametoNum.Clear()
            WriteLog("Driver " & Me.Nom & " arrêté")
        Catch ex As Exception
            WriteLog("ERR: Driver " & Me.Nom & " Erreur arrêt " & ex.Message)
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
        Try
            Try ' lecture de la variable, permet de rafraichir la variable sans redemarrer le service
                _DEBUG = _Parametres.Item(0).Valeur
                _NbLIGNES = _Parametres.Item(1).Valeur
                _NbCARACT = _Parametres.Item(7).Valeur
                _SAUTLIGNE = _Parametres.Item(8).Valeur
                _AFFDESCRIPT = _Parametres.Item(9).Valeur
                _AFFTITRES = _Parametres.Item(10).Valeur
                _LIGNESEPARAT = _Parametres.Item(11).Valeur
            Catch ex As Exception
                _DEBUG = False
                _Parametres.Item(0).Valeur = False
                _Parametres.Item(1).Valeur = 5
                _Parametres.Item(7).Valeur = 50
                _Parametres.Item(8).Valeur = False
                _Parametres.Item(9).Valeur = False
                _Parametres.Item(10).Valeur = True
                _Parametres.Item(11).Valeur = False
                WriteLog("ERR: Erreur de lecture paramètre : " & ex.Message)
            End Try

            If _Enable = False Then Exit Sub

            If _IsConnect = False Then
                WriteLog("ERR: READ, Le driver n'est pas démarré, impossible d'écrire sur le port")
                Exit Sub
            End If

            'Si internet n'est pas disponible on ne mets pas à jour les informations
            If My.Computer.Network.IsAvailable = False Then
                WriteLog("ERR: Pas de connexion réseau, lecture des alertes impossible")
                Exit Sub
            End If

            'Vérification du nombre de caratères par ligne
            If _NbCARACT < 20 Or _NbCARACT > 150 Then
                WriteLog("ERR: Nombre de caratères par ligne hors limite (20 à 150)")
                Exit Sub
            End If

            ''Vérification du nombre de ligne d'infos
            If _NbLIGNES < 1 Or _NbLIGNES > 30 Then
                WriteLog("ERR: Nombre de lignes hors limite (0 à 25)")
                Exit Sub
            End If

            Dim typeSource As String = Objet.Adresse1
            Dim Info As String = Objet.Adresse2
            Dim Temp As String = ""

            Select Case UCase(typeSource)
                Case "LEMONDE"
                    If Info = "" Then
                        Temp = "https://www.lemonde.fr/rss/une.xml"
                    Else
                        Temp = "https://www.lemonde.fr/" & Info & "/rss_full.xml"
                    End If
                    Dim alrt As String = GetInfosWeb(Temp)
                    If alrt <> "" Then Objet.Value = alrt

                Case "LEFIGARO"
                    If Info = "" Then
                        Temp = "http://www.lefigaro.fr/rss/figaro_flash-actu.xml"
                    Else
                        Temp = "http://www.lefigaro.fr/rss/figaro_" & Info & ".xml"
                    End If
                    Dim alrt As String = GetInfosWeb(Temp)
                    If alrt <> "" Then Objet.Value = alrt

                Case "INFOWEB1"
                    If _Parametres.Item(2).Valeur <> "" Then
                        Dim alrt As String = GetInfosWeb(_Parametres.Item(2).Valeur)
                        If alrt <> "" Then Objet.Value = alrt
                    Else
                        WriteLog("ERR: Read INFOWEB_1, Adresse Web Vide")
                    End If

                Case "INFOWEB2"
                    If _Parametres.Item(3).Valeur <> "" Then
                        Dim alrt As String = GetInfosWeb(_Parametres.Item(3).Valeur)
                        If alrt <> "" Then Objet.Value = alrt
                    Else
                        WriteLog("ERR: Read INFOWEB_2, Adresse Web Vide")
                    End If

                Case "INFOWEB3"
                    If _Parametres.Item(4).Valeur <> "" Then
                        Dim alrt As String = GetInfosWeb(_Parametres.Item(4).Valeur)
                        If alrt <> "" Then Objet.Value = alrt
                    Else
                        WriteLog("ERR: Read INFOWEB_3, Adresse Web Vide")
                    End If

                Case "INFOWEB4"
                    If _Parametres.Item(5).Valeur <> "" Then
                        Dim alrt As String = GetInfosWeb(_Parametres.Item(5).Valeur)
                        If alrt <> "" Then Objet.Value = alrt
                    Else
                        WriteLog("ERR: Read INFOWEB_4, Adresse Web Vide")
                    End If

                Case "INFOWEB5"
                    If _Parametres.Item(6).Valeur <> "" Then
                        Dim alrt As String = GetInfosWeb(_Parametres.Item(6).Valeur)
                        If alrt <> "" Then Objet.Value = alrt
                    Else
                        WriteLog("ERR: Read INFOWEB_5, Adresse Web Vide")
                    End If
            End Select

        Catch ex As Exception
            WriteLog("ERR: Read, Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Commander un device</summary>
    ''' <param name="Objet">Objet représentant le device à interroger</param>
    ''' <param name="Command">La commande à passer</param>
    ''' <param name="Parametre1"></param>
    ''' <param name="Parametre2"></param>
    ''' <remarks></remarks>
    Public Sub Write(ByVal Objet As Object, ByVal Command As String, Optional ByVal Parametre1 As Object = Nothing, Optional ByVal Parametre2 As Object = Nothing) Implements HoMIDom.HoMIDom.IDriver.Write
        Try
            If _Enable = False Then Exit Sub

            If _IsConnect = False Then
                WriteLog("ERR: READ, Le driver n'est pas démarré, impossible d'écrire sur le port")
                Exit Sub
            End If

        Catch ex As Exception
            WriteLog("ERR: Write, Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Fonction lancée lors de la suppression d'un device</summary>
    ''' <param name="DeviceId">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub DeleteDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.DeleteDevice
        Try

        Catch ex As Exception
            WriteLog("ERR: DeleteDevice, Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Fonction lancée lors de l'ajout d'un device</summary>
    ''' <param name="DeviceId">Objet représetant le device à interroger</param>
    ''' <remarks></remarks>
    Public Sub NewDevice(ByVal DeviceId As String) Implements HoMIDom.HoMIDom.IDriver.NewDevice
        Try

        Catch ex As Exception
            WriteLog("ERR: NewDevice, Exception : " & ex.Message)
        End Try
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
            WriteLog("ERR: Add_DeviceCommande, Exception :" & ex.Message)
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
            WriteLog("ERR: Add_LibelleDriver, Exception : " & ex.Message)
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
            WriteLog("ERR: Add_LibelleDevice, Exception : " & ex.Message)
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
            WriteLog("ERR: Add_ParamAvance, Exception : " & ex.Message)
        End Try
    End Sub

    ''' <summary>Creation d'un objet de type</summary>
    ''' <remarks></remarks>
    Public Sub New()
        Try
            _Version = Reflection.Assembly.GetExecutingAssembly.GetName.Version.ToString

            'liste des devices compatibles
            _DeviceSupport.Add(ListeDevices.GENERIQUESTRING)

            'Parametres avancés
            Add_ParamAvance("Debug", "Activer le Debug complet (True/False)", False)
            Add_ParamAvance("Nombre d'infos", "Nombre d'infos à Afficher (1 à 30)", 5)
            Add_ParamAvance("INFOWEB1", "Adresse des infos (Ex: https://www.lemonde.fr/rss/une.xml)", "")
            Add_ParamAvance("INFOWEB2", "Adresse des infos (Ex: https://www.lemonde.fr/rss/une.xml)", "")
            Add_ParamAvance("INFOWEB3", "Adresse des infos (Ex: https://www.lemonde.fr/rss/une.xml)", "")
            Add_ParamAvance("INFOWEB4", "Adresse des infos (Ex: https://www.lemonde.fr/rss/une.xml)", "")
            Add_ParamAvance("INFOWEB5", "Adresse des infos (Ex: https://www.lemonde.fr/rss/une.xml)", "")
            Add_ParamAvance("Nombre de caractères", "Nombre de caractères par ligne (20 / 150)", 50)
            Add_ParamAvance("Saut de ligne", "Sauter une ligne après chaque titre (True/False)", False)
            Add_ParamAvance("Descriptions", "Activer l'affichage des déscriptions (True/False)", False)
            Add_ParamAvance("Titre Info", "Activer l'affichage des titres (True/False)", True)
            Add_ParamAvance("Séparation", "Activer l'affichage d'une ligne de séparation (True/False)", False)

            'Libellé Driver
            Add_LibelleDriver("HELP", "Aide...", "Pas d'aide actuellement...")

            'Libellé Device
            Add_LibelleDevice("ADRESSE1", "Type Source: Actualité", "Le Monde, Lefigaro, InfoWeb_1, InfoWeb_2, ...", "")
            Add_LibelleDevice("ADRESSE2", "Nom info", "Nom de l'info pour le Monde (culture, m-actu, en-bref, paris, ... )", "")
            Add_LibelleDevice("REFRESH", "Refresh en sec", "Minimum 600, valeur rafraicissement station", "600")
            ' Libellés Device inutiles
            Add_LibelleDevice("SOLO", "@", "")
            Add_LibelleDevice("MODELE", "@", "")
            Add_LibelleDevice("LASTCHANGEDUREE", "@", "")
        Catch ex As Exception
            WriteLog("ERR: New, Exception : " & ex.Message)
        End Try
    End Sub

#End Region

#Region "Fonctions internes"

    Private Function GetInfosWeb(ByVal AdresseWeb As String)
        Try
            Dim doc As New XmlDocument
            Dim nodes As XmlNodeList
            Dim Donnees As String = ""
            Dim stringurl As String = ""
            ' Dim Infos As String = ""
            Dim NbInfos As Integer = 0
            Dim Temp As String = ""
            Dim NbCaract As Integer = 0
            Dim Temp1 As String = ""

            doc = New XmlDocument()
            stringurl = AdresseWeb
            If stringurl <> "" Then
                WriteLog("DBG: GetInfosWeb, Chargement de " & stringurl)
                Dim url As New Uri(stringurl)
                Dim Request As HttpWebRequest = CType(HttpWebRequest.Create(url), System.Net.HttpWebRequest)
                Dim response As Net.HttpWebResponse = CType(Request.GetResponse(), Net.HttpWebResponse)
                doc.Load(response.GetResponseStream)
                Dim version As String = doc.DocumentElement.GetAttribute("version")
                If version = "2.0" Then
                    nodes = doc.SelectNodes("/rss/channel/item")
                    For Each node As XmlNode In nodes
                        For Each _child As XmlNode In node
                            Select Case _child.Name
                                Case "description"
                                    If _AFFDESCRIPT And _child.FirstChild IsNot Nothing Then
                                        Temp = _child.FirstChild.Value
                                        WriteLog("DBG: GetInfosWeb, Infos Description => " & Temp)
                                        Temp = Remplace_html_entities(Temp)
                                        Temp = Replace(Temp, vbCrLf, "")   'effacer les retours à la ligne
                                        Do
                                            NbCaract = InStr(_NbCARACT, Temp, " ")
                                            If NbCaract = 0 Then
                                                Donnees = Donnees & Temp & vbCrLf
                                                Exit Do
                                            Else
                                                NbCaract = InStr(_NbCARACT, Temp, " ")
                                                If NbCaract Then
                                                    Select Case True
                                                        Case InStr(_NbCARACT, Temp, ":") = NbCaract + 1 : NbCaract = NbCaract + 2
                                                        Case InStr(_NbCARACT, Temp, "?") = NbCaract + 1 : NbCaract = NbCaract + 2
                                                        Case InStr(_NbCARACT, Temp, "!") = NbCaract + 1 : NbCaract = NbCaract + 2
                                                    End Select
                                                    Temp1 = Left(Temp, NbCaract)
                                                    Donnees = Donnees & Temp1 & vbCrLf
                                                    Temp = Mid(Temp, NbCaract + 1, Len(Temp))
                                                End If
                                            End If
                                        Loop Until Len(Temp) = 0
                                    End If

                      '  Case "guid" : desc = _child.FirstChild.Value
                      '  Case "pubDate" : desc = _child.FirstChild.Value
                     '   Case "category" : desc = _child.FirstChild.Value
                     '   Case "link" : desc = _child.FirstChild.Value

                                Case "title"
                                    If _AFFTITRES And _child.FirstChild IsNot Nothing Then
                                        Temp = _child.FirstChild.Value
                                        WriteLog("DBG: GetInfosWeb, Infos Title => " & Temp)
                                        Temp = Replace(Temp, vbCrLf, "")   'effacer les retours à la ligne
                                        Do
                                            NbCaract = InStr(_NbCARACT, Temp, " ")
                                            If NbCaract = 0 Then
                                                Donnees = Donnees & Temp & vbCrLf
                                                Exit Do
                                            Else
                                                NbCaract = InStr(_NbCARACT, Temp, " ")
                                                If NbCaract Then
                                                    Select Case True
                                                        Case InStr(_NbCARACT, Temp, ":") = NbCaract + 1 : NbCaract = NbCaract + 2
                                                        Case InStr(_NbCARACT, Temp, "?") = NbCaract + 1 : NbCaract = NbCaract + 2
                                                        Case InStr(_NbCARACT, Temp, "!") = NbCaract + 1 : NbCaract = NbCaract + 2
                                                    End Select
                                                    Temp1 = Left(Temp, NbCaract)
                                                    Donnees = Donnees & Temp1 & vbCrLf
                                                    Temp = Mid(Temp, NbCaract + 1, Len(Temp))
                                                End If
                                            End If
                                        Loop Until Len(Temp) = 0
                                    End If
                            End Select
                        Next
                        NbInfos = NbInfos + 1
                        If _LIGNESEPARAT And (NbInfos < _NbLIGNES) Then
                            Dim Separat As New String("-", _NbCARACT + 15)
                            Donnees = Donnees & Separat & vbCrLf
                        End If
                        If _SAUTLIGNE And (NbInfos < _NbLIGNES) Then
                            Donnees = Donnees & vbCrLf
                        End If
                        If NbInfos >= _NbLIGNES Then
                            Return Donnees
                        End If
                    Next
                End If
            Else
                WriteLog("ERR: GetInfosWeb, Adresse Web Vide")
            End If
            If Donnees = "" Then
                Donnees = "Pas de données"
            End If
            WriteLog("DBG: GetInfosWeb : Données => " & Donnees)
            Return Donnees
        Catch ex As Exception
            WriteLog("ERR: GetInfosWeb, Exception : " & ex.Message)
            Return ""
        End Try

    End Function


    Public Function Remplace_html_entities(ByVal e_txt As String) As String
        Try
            'recherche les entités via une expression réguliere
            Dim myregexp As New Regex("&[#a-zA-Z0-9]{2,6}\;", RegexOptions.IgnoreCase)
            If myregexp.IsMatch(e_txt) Then
                Dim elt As Match
                For Each elt In myregexp.Matches(e_txt)
                    If NametoNum.ContainsKey(elt.Value) Then
                        'remplace l'occurence par la valeur associée
                        WriteLog("DBG: Remplace_html_entities, Code : " & elt.Value)
                        e_txt = Replace(e_txt, elt.Value, Chr(NametoNum(elt.Value)))
                    Else
                        WriteLog("ERR: Remplace_html_entities, Erreur code non traité : " & elt.Value)
                    End If
                Next
            End If
            WriteLog("DBG: Remplace_html_entities, Texte remplacé : " & e_txt)
            Return e_txt
        Catch ex As Exception
            WriteLog("ERR: Remplace_html_entities, Exception : " & ex.Message)
            Return e_txt
        End Try
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

