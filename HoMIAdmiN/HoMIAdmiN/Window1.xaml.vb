﻿Imports System.Windows.Threading
'Imports System.Runtime.Serialization.Formatters.Soap
Imports HoMIDom.HoMIDom
Imports System.ServiceModel
Imports System.IO
Imports System.Xml
Imports System.Data
Imports System.Xml.Serialization
Imports System.Threading
Imports System.Reflection.Assembly
Imports System.Windows.Media.Animation
Imports System.ComponentModel
Imports System.Net
Imports System.Text


Class Window1
#Region "Variables"
    Public Event menu_gerer(ByVal IndexMenu As Integer)

    Public Shared CanvasUser As Canvas
    Public Shared ListServer As New List(Of ClServer)

    Dim Myfile As String
    Dim myChannelFactory As ServiceModel.ChannelFactory(Of HoMIDom.HoMIDom.IHoMIDom) = Nothing
    Dim FlagStart As Boolean = False
    ' Dim MyRep As String = System.Environment.CurrentDirectory
    Dim MyRep As String = My.Application.Info.DirectoryPath
    Dim MyRepAppData As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData) & "\HoMIAdmiN"
    Dim mypush As PushNotification

    Dim MainMenu As uMainMenu
    Dim WMainMenu As Double = 0
    Dim HMainMenu As Double = 0
    Dim _MainMenuAction As Integer = -1 'NEW,MODIF,SUPP -1 si aucun
    Dim myBrushVert As New RadialGradientBrush()
    Dim myBrushRouge As New RadialGradientBrush()
    Dim flagTreeV As Boolean = False
    Dim flagShowMainMenu As Boolean = False
    Dim FlagOK As Boolean = False
    Dim _IsConnect As Boolean = False
    Dim _ShowNbHistoInDevice As Boolean = False


#End Region

#Region "Fonctions de base"

    Public Sub New()
        Try
            ' Cet appel est requis par le Concepteur Windows Form.
            InitializeComponent()

            Me.Title = "HoMIAdmiN v" & My.Application.Info.Version.ToString & " - HoMIDoM"


            Dim spl As Window2 = New Window2
            spl.Show()
            Thread.Sleep(1000)

            Me.Window1.WindowState = My.Settings.WindowState
            If Me.Window1.WindowState = Windows.WindowState.Normal Then
                Me.Window1.Top = My.Settings.Top
                Me.Window1.Left = My.Settings.Left
                Me.Window1.Width = My.Settings.Width
                Me.Window1.Height = My.Settings.Height
            End If

            myBrushVert.GradientOrigin = New Point(0.75, 0.25)
            myBrushVert.GradientStops.Add(New GradientStop(Colors.LightGreen, 0.0))
            myBrushVert.GradientStops.Add(New GradientStop(Colors.Green, 0.5))
            myBrushVert.GradientStops.Add(New GradientStop(Colors.DarkGreen, 1.0))

            myBrushRouge.GradientOrigin = New Point(0.75, 0.25)
            myBrushRouge.GradientStops.Add(New GradientStop(Colors.Yellow, 0.0))
            myBrushRouge.GradientStops.Add(New GradientStop(Colors.Red, 0.5))
            myBrushRouge.GradientStops.Add(New GradientStop(Colors.DarkRed, 1.0))

            flagTreeV = True

            spl.Close()
            spl = Nothing

            'AddHandler Window1.menu_gerer, AddressOf MainMenuGerer
            ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
            AddHandler Timer_Sec.Tick, AddressOf dispatcherTimer_Tick
            Timer_Sec.Interval = New TimeSpan(0, 0, 1)
            Timer_Sec.Start()

            'si le repertoire appdata n'existe pas on le crée et copie la config depuis le repertoire courant
            If Not System.IO.Directory.Exists(System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData) & "\HoMIAdmiN") Then
                System.IO.Directory.CreateDirectory(System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData) & "\HoMIAdmiN")
            End If
            MyRepAppData = System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData) & "\HoMIAdmiN"
            'Myfile = MyRep & "\Config\HoMIAdmiN.xml"
            Myfile = MyRepAppData & "\HoMIAdmiN.xml"
            If Not System.IO.File.Exists(Myfile) Then
                System.IO.File.Copy(MyRep & "\Config\HoMIAdmiN.xml", Myfile, False)
            End If

            AffIsDisconnect()

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub New: " & ex.Message, "ERREUR", "")
        End Try

    End Sub

    'Affiche la date et heure, heures levé et couché du soleil
    Public Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        Timer_Sec.Stop()
        Try
            If IsConnect = True Then
                Try
                    Dim _string As String
                    Dim list As New List(Of String)

                    'Modifie la date et Heure
                    LblStatus.Content = Now.ToLongDateString & " " & Now.ToLongTimeString

                    If Now.Second = 0 Then
                        'Modifie les heures du soleil
                        LHS.Content = CDate(myService.GetHeureLeverSoleil).ToShortTimeString
                        LCS.Content = CDate(myService.GetHeureCoucherSoleil).ToShortTimeString

                        'Vérifie si nouveau device
                        If myService.AsNewDevice Then
                            ImgNewDevice.ToolTip = "Afficher les nouveaux composants du mode découverte"
                            ImgNewDevice.Visibility = Windows.Visibility.Visible
                        Else
                            If ImgNewDevice.Visibility = Windows.Visibility.Visible Then ImgNewDevice.Visibility = Windows.Visibility.Collapsed
                        End If

                        _string = "Liste des composants non à jour : " & vbCrLf
                        list = myService.GetDeviceNoMaJ(IdSrv)
                        If list.Count > 0 Then
                            For Each logmsg As String In list
                                If String.IsNullOrEmpty(logmsg) = False Then
                                    _string &= logmsg & vbCrLf
                                    If ImgDeviceNoMaj.Visibility <> Windows.Visibility.Visible Then ImgDeviceNoMaj.Visibility = Windows.Visibility.Visible
                                    ImgDeviceNoMaj.ToolTip = _string
                                End If
                            Next
                            _string = Nothing
                        Else
                            ImgDeviceNoMaj.Visibility = Windows.Visibility.Collapsed
                        End If
                        list.Clear()
                    End If

                    'met a jour le treeview si ajout-retrait device
                    If refreshtreeviewdevice And Tabcontrol1.SelectedIndex = 1 Then
                        Me.AffDevice()
                        refreshtreeviewdevice = False
                    End If

                    'Modifie les LOG
                    list = myService.GetLastLogs
                    If list IsNot Nothing Then
                        If list.Count > 0 Then
                            _string = "Liste des Logs du serveur : " & vbCrLf
                            For Each logmsg In list
                                If String.IsNullOrEmpty(logmsg) = False Then _string &= logmsg & vbCrLf
                            Next
                            LOG.ToolTip = _string
                            ImgLog.ToolTip = _string
                            LOG.Content = Mid(list(0), 1, 100) & "..."
                        Else
                            LOG.ToolTip = "Pas encore de logs..."
                            ImgLog.ToolTip = "Pas encore de logs..."
                            LOG.Content = "Pas encore de logs..."
                        End If
                        list.Clear()
                    End If

                    list = myService.GetLastLogsError
                    If list IsNot Nothing Then
                        If list.Count > 0 Then
                            _string = "Liste des Erreurs du serveur : " & vbCrLf
                            For Each logerror As String In list
                                If String.IsNullOrEmpty(logerror) = False Then
                                    _string &= logerror & vbCrLf
                                    If ImgError.Visibility <> Windows.Visibility.Visible Then ImgError.Visibility = Windows.Visibility.Visible
                                    ImgError.ToolTip = _string
                                End If
                            Next
                        Else
                            ImgError.Visibility = Windows.Visibility.Collapsed
                        End If
                        list.Clear()
                    End If

                    _string = Nothing

                Catch ex As Exception
                    IsConnect = False
                    AffIsDisconnect()
                    LblStatus.Content = Now.ToLongDateString & " " & Now.ToLongTimeString & " "

                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "La communication a été perdue avec le serveur, veuillez vérifier que celui-ci est toujours actif", "ERREUR", "")
                    ClearAllTreeview() 'vide le treeview
                    CanvasRight.Children.Clear()  'ferme le menu principal
                    PageConnexion() 'affiche la fenetre de connexion

                End Try
            Else
                LblStatus.Content = Now.ToLongDateString & " " & Now.ToLongTimeString & "      "
            End If

            If _IsConnect <> IsConnect Then
                _IsConnect = IsConnect
                If _IsConnect = True Then
                    AffIsConnect()
                Else
                    AffIsDisconnect()
                End If
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR dispatcherTimer_Tick: " & ex.Message, "ERREUR", "")
        End Try

        Timer_Sec.Start()
    End Sub

    Private Sub AffIsConnect()
        'Modifie l'adresse/port du serveur connecté
        Dim serveurcomplet As String = myChannelFactory.Endpoint.Address.ToString()
        serveurcomplet = Mid(serveurcomplet, 8, serveurcomplet.Length - 8)
        serveurcomplet = Split(serveurcomplet, "/", 2)(0)
        Dim srblbl As String = Split(serveurcomplet, ":", 2)(0) & ":" & Split(serveurcomplet, ":", 2)(1)
        LblConnect.ToolTip = "Serveur connecté adresse utilisée: " & srblbl
        LblConnect.Content = srblbl

        'Modifie les heures du soleil
        LHS.Content = CDate(myService.GetHeureLeverSoleil).ToShortTimeString
        LCS.Content = CDate(myService.GetHeureCoucherSoleil).ToShortTimeString
        serveurcomplet = Nothing
        srblbl = ""

        Ellipse1.Fill = myBrushVert
    End Sub

    Private Sub AffIsDisconnect()
        Ellipse1.Fill = myBrushRouge
        LblConnect.Content = "Serveur non connecté"
        LblConnect.ToolTip = Nothing
    End Sub

    Private Sub StoryBoardFinish(ByVal sender As Object, ByVal e As System.EventArgs)
        CanvasRight.Children.Clear()
        CanvasRight.UpdateLayout()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

        ShowMainMenu()
    End Sub

    Protected Overrides Sub Finalize()
        Try
            'mypush.Close()
            If ClientUDP IsNot Nothing Then ClientUDP.Disconnect()

            If ListServer.Count > 0 Then
                'Serialize object to a text file.
                Dim objStreamWriter As New StreamWriter(Myfile)
                Dim x As New XmlSerializer(ListServer.GetType)
                x.Serialize(objStreamWriter, ListServer)
                objStreamWriter.Close()
                x = Nothing
            End If

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'enregistrement du fichier de config xml de l'application Admin : " & ex.Message, "Erreur", "")
        End Try
        MyBase.Finalize()
    End Sub

    ''' <summary>
    ''' Sélection d'un serveur pour connexion dans le sous menu connexion ou au démarrage
    ''' </summary>
    ''' <param name="Name"></param>
    ''' <param name="IP"></param>
    ''' <param name="Port"></param>
    ''' <remarks></remarks>
    Public Function Connect_Srv(ByVal Name As String, ByVal IP As String, ByVal Port As String) As Integer
        Try
            Me.Cursor = Cursors.Wait
            myadress = "http://" & IP & ":" & Port & "/service"
            MyPort = Port

            LOG.Content = "Connexion au serveur...."
            If UrlIsValid(myadress) = False Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la connexion au serveur sélectionné: " & Chr(10) & Name & " - " & IP & ":" & Port & vbCrLf & "Veuillez vérifier que celui-ci est démarré", "Erreur Admin", "")
                Return -1
            End If

            Dim binding As New ServiceModel.BasicHttpBinding
            binding.MaxBufferPoolSize = 250000000
            binding.MaxReceivedMessageSize = 250000000
            binding.MaxBufferSize = 250000000
            binding.ReaderQuotas.MaxArrayLength = 250000000
            binding.ReaderQuotas.MaxNameTableCharCount = 250000000
            binding.ReaderQuotas.MaxBytesPerRead = 250000000
            binding.ReaderQuotas.MaxStringContentLength = 250000000
            binding.SendTimeout = TimeSpan.FromMinutes(60)
            binding.CloseTimeout = TimeSpan.FromMinutes(60)
            binding.OpenTimeout = TimeSpan.FromMinutes(60)
            binding.ReceiveTimeout = TimeSpan.FromMinutes(60)

            myChannelFactory = New ServiceModel.ChannelFactory(Of HoMIDom.HoMIDom.IHoMIDom)(binding, New System.ServiceModel.EndpointAddress(myadress))
            myService = myChannelFactory.CreateChannel()


            If myChannelFactory.State <> CommunicationState.Opened Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la connexion au serveur", "Erreur Admin", "")
                IsConnect = False
                Return -1
                Exit Function
            Else
                Try
                    myService.GetTime()
                Catch ex As Exception
                    myChannelFactory.Abort()
                    IsConnect = False
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la connexion au serveur sélectionné: " & Chr(10) & Name & " - " & IP & ":" & Port & vbCrLf & "Erreur: " & ex.ToString, "Erreur Admin", "")
                    Return -1
                End Try

                If myService.GetIdServer(IdSrv) = "99" Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "L'ID du serveur est erroné, impossible de communiquer avec celui-ci", "Erreur", "")
                    Return -1
                End If

                IsConnect = True
                LOG.Content = "Connexion au serveur effectuée...."

                'LOG.Content = "Connexion au serveur live...."
                'mypush = New PushNotification("http://" & IP & ":" & Port & "/live", IdSrv)
                'AddHandler mypush.DeviceChanged, AddressOf DeviceChanged
                'mypush.Open()

                LOG.Content = "Connexion au serveur UDP...."
                ClientUDP = New HoMIDom.HoMIDom.UDPClient(myService)

                'ClientUDP.Connect("Application Admin", System.Net.Dns.GetHostEntry(IP).AddressList(0).ToString(), CInt(Port) - 1)
                'ClientUDP.Connect("Application Admin", System.Net.Dns.GetHostEntry(Dns.GetHostName()).AddressList.Where(Function(a As IPAddress) Not a.IsIPv6LinkLocal AndAlso Not a.IsIPv6Multicast AndAlso Not a.IsIPv6SiteLocal).First().ToString(), CInt(Port) - 1)
                ClientUDP.Connect("Application Admin", System.Net.Dns.GetHostEntry(Dns.GetHostName()).AddressList.Where(Function(a As IPAddress) a.AddressFamily = Sockets.AddressFamily.InterNetwork).First().ToString(), CInt(Port) - 1)

                AddHandler ClientUDP.OnMessageReceive, AddressOf OnMessageReceiveUDP

                Tabcontrol1.SelectedIndex = 0
                My.Settings.SaveRealTime = myService.GetSaveRealTime
                My.Settings.Save()

                AffDriver()
                ShowMainMenu()

                LOG.Content = "Chargement des drivers...."
                ManagerDrivers.LoadDrivers()
                LOG.Content = "Chargement des zones...."
                ManagerZones.LoadZones()
                LOG.Content = "Chargement des devices...."
                ManagerDevices.LoadDevices()

                '_TableDBHisto = myService.GetTableDBHisto(IdSrv)
            End If

            LblSrv.Content = Name
            LblSrv.ToolTip = "Serveur courant : " & Name
            For i As Integer = 0 To ListServer.Count - 1
                If ListServer.Item(i).Nom = Name Then
                    If File.Exists(ListServer.Item(i).Icon) Then
                        ImgSrv.Source = ImageFromUri(ListServer.Item(i).Icon)
                        Exit For
                    End If
                End If
            Next
            CanvasUser = CanvasRight
            Me.Cursor = Nothing

            Return 0
            binding = Nothing
        Catch ex As Exception
            Me.Cursor = Nothing
            myChannelFactory.Abort()
            IsConnect = False
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la connexion au serveur sélectionné: " & ex.ToString, "Erreur Admin", "")
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Message UDP reçu
    ''' </summary>
    ''' <param name="Message"></param>
    ''' <remarks></remarks>
    Public Sub OnMessageReceiveUDP(Message As String)
        Try
            If String.IsNullOrEmpty(Message) = False Then
                Dim a() As String
                a = Message.Split("|")
                If a.Length > 0 Then
                    Select Case a(0).ToUpper
                        Case HoMIDom.HoMIDom.Sequence.TypeOfSequence.Device
                            Dim seq As Sequence = myService.ReturnSequenceFromNumero(a(1))
                            DeviceChanged(seq.ID, Nothing)
                        Case HoMIDom.HoMIDom.Sequence.TypeOfSequence.DeviceChange
                            Dim seq As Sequence = myService.ReturnSequenceFromNumero(a(1))
                            DeviceChanged(seq.ID, Nothing)
                        Case HoMIDom.HoMIDom.Sequence.TypeOfSequence.Driver

                        Case HoMIDom.HoMIDom.Sequence.TypeOfSequence.Server

                        Case HoMIDom.HoMIDom.Sequence.TypeOfSequence.Macro

                        Case HoMIDom.HoMIDom.Sequence.TypeOfSequence.Trigger

                        Case HoMIDom.HoMIDom.Sequence.TypeOfSequence.Zone

                        Case Else
                    End Select
                End If
            End If
        Catch ex As CommunicationException
            'nothing to do, server is disconnected
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur OnMessageReceiveUDP: " & ex.ToString, "Erreur", " OnMessageReceiveUDP")
        End Try
    End Sub

    'Event lorsqu'un device change
    Private Sub DeviceChanged(ByVal deviceid As String, ByVal DeviceValue As String)
        Try
            For Each _dev In _ListeDevices
                If _dev.ID = deviceid Then
                    _dev = myService.ReturnDeviceByID(IdSrv, deviceid)
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.DEBUG, "Le device " & _dev.Name & " a été mis à jour", "Debug Admin", "")
                End If
            Next
        Catch ex As EndpointNotFoundException
            'nothing to do, server is disconnected
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur Event DeviceChanged: " & ex.ToString, "Erreur", " Event DeviceChanged")
        End Try
    End Sub

    Private Sub Window1_Closing(ByVal sender As Object, ByVal e As System.ComponentModel.CancelEventArgs) Handles Me.Closing
        Try
            If IsConnect = True And FlagChange Then

                If MessageBox.Show("Voulez-vous enregistrer la configuration avant de quitter?", "HomIAdmin", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                    Try
                        If IsConnect = False Then
                            MessageBox.Show("Impossible d'enregistrer la configuration car le serveur n'est pas connecté !!", "Erreur", MessageBoxButton.OK, MessageBoxImage.Asterisk)
                        Else
                            Dim retour As String = myService.SaveConfig(IdSrv)
                            If retour <> "0" Then
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'enregistrement veuillez consulter le log", "HomIAdmin", "")
                            Else
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Enregistrement effectué", "HomIAdmin", "")
                            End If
                        End If
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Unloaded: " & ex.Message, "ERREUR", "")
                    End Try

                End If
                FlagChange = False
            End If
            My.Settings.WindowState = Me.Window1.WindowState
            If Me.Window1.WindowState = Windows.WindowState.Normal Then
                My.Settings.Top = Me.Window1.Top
                My.Settings.Left = Me.Window1.Left
                My.Settings.Width = Me.Window1.Width
                My.Settings.Height = Me.Window1.Height
            End If
            My.Settings.Save()

            End
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Unloaded Quitter: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub Window1_Loaded(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        Try
            'Connexion au serveur web
            If File.Exists(Myfile) Then
                Try
                    'Deserialize text file to a new object.
                    Dim objStreamReader As New StreamReader(Myfile)
                    Dim x As New XmlSerializer(ListServer.GetType)
                    ListServer = x.Deserialize(objStreamReader)
                    x = Nothing
                    objStreamReader.Close()
                    objStreamReader = Nothing
                Catch ex As Exception
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'ouverture du fichier de config xml (" & Myfile & "), vérifiez que toutes les balises requisent soient présentes: " & ex.Message, "Erreur Admin", "")
                End Try

                If ListServer.Count = 0 Then 'Aucun serveur trouvé dans le fichier de config
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Aucun serveur n'a été trouvé dans le fichier de config (" & Myfile & "), veuillez saisir manuellement les paramètres du serveur", "Info Admin", "")
                End If
                'on ajoute la liste des serveurs dans le menu Connexion
            Else
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le fichier de config xml de l'admin (" & Myfile & ") est absent, veuillez utiliser la connexion manuelle", "Admin", "")
            End If

            'Page de connexion
            Do While FlagStart = False
                PageConnexion()
            Loop

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Window1_Loaded: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    ''' <summary>
    ''' Affiche la fenêtre de connexion
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub PageConnexion()
        Try
            Dim frm As New Window3
            frm.Owner = Me
            frm.ShowDialog()

            If frm.DialogResult.HasValue And frm.DialogResult.Value Then
                LOG.Content = "Connexion au serveur...."
                If Connect_Srv(frm.TxtName.Text, frm.TxtIP.Text, frm.TxtPort.Text) <> 0 Then
                    frm.Close()
                    PageConnexion()
                    Me.Cursor = Nothing
                    Exit Sub
                Else
                    If myService.GetIdServer(IdSrv) = "99" Then
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "L'ID du serveur est erroné, impossible de communiquer avec celui-ci", "Erreur", "")
                        frm.Close()
                        Me.Cursor = Nothing
                        Exit Sub
                    End If
                    'If myService.VerifLogin(frm.TxtUsername.Text, frm.TxtPassword.Password) = False Then
                    '    AfficheMessageAndLog (HoMIDom.HoMIDom.Server.TypeLog.ERREUR,"Le username ou le password sont erroné, impossible veuillez réessayer", "Erreur","")
                    'Else
                    frm.Close()
                    FlagStart = True
                    frm = Nothing
                    LOG.Content = "Connexion...."
                    'End If
                End If
            Else
                End
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Connexion: " & ex.Message, "ERREUR", "")
        End Try

        Me.Cursor = Nothing
    End Sub

    'action à faire quand le serveur n'est pas connecté
    Private Sub Serveur_notconnected_action()
        MessageBox.Show("Action Impossible, le serveur n'est pas connecté !", "Erreur", MessageBoxButton.OK, MessageBoxImage.Asterisk)
        ClearAllTreeview() 'vide le treeview
        CanvasRight.Children.Clear()  'ferme le menu principal
        PageConnexion() 'affiche la fenetre de connexion
    End Sub

#End Region

#Region "Treeview"

    Private Sub ShowTreeView()
        Try
            If flagTreeV = True Then Exit Sub
            DKpanel.Width = Double.NaN
            Dim da3 As DoubleAnimation = New DoubleAnimation
            da3.From = 0
            da3.To = 1
            da3.Duration = New Duration(TimeSpan.FromMilliseconds(600))
            Dim sc As ScaleTransform = New ScaleTransform()
            DKpanel.RenderTransform = sc
            sc.BeginAnimation(ScaleTransform.ScaleXProperty, da3)
            flagTreeV = True
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub ShowTreeView: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub ClearAllTreeview()
        Try
            TreeViewDriver.Items.Clear()
            TreeViewDriver.UpdateLayout()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            TreeViewDevice.Items.Clear()
            TreeViewDevice.UpdateLayout()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            TreeViewZone.Items.Clear()
            TreeViewZone.UpdateLayout()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            TreeViewUser.Items.Clear()
            TreeViewUser.UpdateLayout()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            TreeViewTrigger.Items.Clear()
            TreeViewTrigger.UpdateLayout()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            TreeViewMacro.Items.Clear()
            TreeViewMacro.UpdateLayout()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()
            TreeViewHisto.Items.Clear()
            TreeViewHisto.UpdateLayout()
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            Me.UpdateLayout()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()


        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub ClearAllTreeview: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub RefreshTreeView()
        Try
            If IsConnect = False Then
                'Serveur_notconnected_action()
                Exit Sub
            End If

            Me.Cursor = Cursors.Wait
            ClearAllTreeview()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            Select Case Tabcontrol1.SelectedIndex
                Case 0
                    AffDriver()
                Case 1
                    AffDevice()
                Case 2
                    AffZone()
                Case 3
                    AffUser()
                Case 4
                    AffTrigger()
                Case 5
                    AffScene()
                Case 6
                    AffHisto()
            End Select
            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub RefreshTreeView: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Afficher la liste des zones
    Public Sub AffZone()
        Try
            TreeViewZone.Items.Clear()
            TreeViewZone.UpdateLayout()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            If IsConnect = False Then Exit Sub

            Dim ListeZones As List(Of Zone) = myService.GetAllZones(IdSrv)
            CntZone.Content = ListeZones.Count & " Zone(s)"

            For Each zon As Zone In ListeZones
                Dim newchild As New TreeViewItem
                Dim stack As New StackPanel
                Dim img As New Image
                Dim uri As String = ""
                stack.Orientation = Orientation.Horizontal

                img.Height = 20
                img.Width = 20

                If String.IsNullOrEmpty(Trim(zon.Icon)) = False Then
                    img.Source = ConvertArrayToImage(myService.GetByteFromImage(zon.Icon), 20)
                Else
                    uri = MyRep & "\Images\icones\Zone_32.png"
                    img.Source = ImageFromUri(uri)
                End If

                Dim label As New Label
                label.Foreground = New SolidColorBrush(Colors.White)
                label.Content = zon.Name & " {" & zon.ListElement.Count & " éléments}"

                Dim ctxMenu As New ContextMenu
                ctxMenu.Foreground = System.Windows.Media.Brushes.Black
                ctxMenu.Background = System.Windows.Media.Brushes.LightGray
                ctxMenu.BorderBrush = System.Windows.Media.Brushes.Black
                Dim mnu0 As New MenuItem
                mnu0.Header = "Modifier"
                mnu0.Tag = 0
                mnu0.Uid = zon.ID
                AddHandler mnu0.Click, AddressOf MnuitemZon_Click
                ctxMenu.Items.Add(mnu0)
                Dim mnu1 As New MenuItem
                mnu1.Header = "Supprimer"
                mnu1.Tag = 1
                mnu1.Uid = zon.ID
                AddHandler mnu1.Click, AddressOf MnuitemZon_Click
                ctxMenu.Items.Add(mnu1)
                label.ContextMenu = ctxMenu

                stack.Children.Add(img)
                stack.Children.Add(label)

                newchild.Foreground = New SolidColorBrush(Colors.White)
                newchild.Header = stack
                newchild.Uid = zon.ID
                Dim marg As New Thickness(-12, 0, 0, 0)
                newchild.Margin = marg
                TreeViewZone.Items.Add(newchild)
            Next

            ListeZones = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffZone: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Gère les menus click droit sur les zones
    Private Sub MnuitemZon_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Select Case sender.tag
                Case 0 'Modifier
                    Dim x As New uZone(Classe.EAction.Modifier, sender.uid)
                    AddHandler x.CloseMe, AddressOf UnloadControl
                    AddHandler x.Loaded, AddressOf ControlLoaded
                    AffControlPage(x)
                Case 1 'Supprimer
                    DeleteElement(sender.uid, 2)
            End Select

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MnuitemZon_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Afficher la liste des users
    Public Sub AffUser()
        Try
            TreeViewUser.Items.Clear()
            TreeViewUser.UpdateLayout()

            If IsConnect = False Then Exit Sub

            Dim ListeUsers As List(Of Users.User) = myService.GetAllUsers(IdSrv)
            CntUser.Content = ListeUsers.Count & " User(s)"

            For Each Usr As Users.User In ListeUsers
                Dim newchild As New TreeViewItem
                Dim stack As New StackPanel
                Dim img As New Image
                Dim uri As String = ""

                stack.Orientation = Orientation.Horizontal

                img.Height = 20
                img.Width = 20

                If String.IsNullOrEmpty(Trim(Usr.Image)) = False Then
                    img.Source = ConvertArrayToImage(myService.GetByteFromImage(Usr.Image))
                Else
                    uri = MyRep & "\Images\icones\User_32.png"
                    img.Source = ImageFromUri(uri)
                End If

                Dim label As New Label
                label.Foreground = New SolidColorBrush(Colors.White)
                label.Content = Usr.UserName

                Dim ctxMenu As New ContextMenu
                ctxMenu.Foreground = System.Windows.Media.Brushes.Black
                ctxMenu.Background = System.Windows.Media.Brushes.LightGray
                ctxMenu.BorderBrush = System.Windows.Media.Brushes.Black
                Dim mnu0 As New MenuItem
                mnu0.Header = "Modifier"
                mnu0.Tag = 0
                mnu0.Uid = Usr.ID
                AddHandler mnu0.Click, AddressOf MnuitemUsr_Click
                ctxMenu.Items.Add(mnu0)
                Dim mnu1 As New MenuItem
                mnu1.Header = "Supprimer"
                mnu1.Tag = 1
                mnu1.Uid = Usr.ID
                AddHandler mnu1.Click, AddressOf MnuitemUsr_Click
                ctxMenu.Items.Add(mnu1)
                label.ContextMenu = ctxMenu

                stack.Children.Add(img)
                stack.Children.Add(label)

                newchild.Foreground = New SolidColorBrush(Colors.White)
                newchild.Header = stack
                newchild.Uid = Usr.ID
                Dim marg As New Thickness(-12, 0, 0, 0)
                newchild.Margin = marg
                TreeViewUser.Items.Add(newchild)
            Next

            ListeUsers = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffUser: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Gère les menus click droit sur les users
    Private Sub MnuitemUsr_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Select Case sender.tag
                Case 0 'Modifier
                    Dim x As New uUser(Classe.EAction.Modifier, sender.uid)
                    AddHandler x.CloseMe, AddressOf UnloadControl
                    AddHandler x.Loaded, AddressOf ControlLoaded
                    AffControlPage(x)
                Case 1 'Supprimer
                    DeleteElement(sender.uid, 3)
            End Select

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MnuitemUsr_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Afficher la liste des drivers
    Public Sub AffDriver()
        Try
            TreeViewDriver.Items.Clear()
            TreeViewDriver.UpdateLayout()

            If IsConnect = False Then Exit Sub

            Dim ListeDrivers As List(Of TemplateDriver) = myService.GetAllDrivers(IdSrv)
            CntDriver.Content = ListeDrivers.Count & " Driver(s)"

            Dim Drv As TemplateDriver
            For Each Drv In ListeDrivers
                Dim newchild As New TreeViewItem
                Dim stack As New StackPanel
                stack.Orientation = Orientation.Horizontal
                stack.HorizontalAlignment = HorizontalAlignment.Left
                Dim Graph As Shape
                Dim img As New Image
                Dim tool As New Label

                img.Height = 20
                img.Width = 20
                img.Margin = New Thickness(2)
                If String.IsNullOrEmpty(Drv.Picture) = False Then
                    img.Source = ConvertArrayToImage(myService.GetByteFromImage(Drv.Picture), 20)
                End If

                If Drv.IsConnect = True Then
                    Dim Rect As New Polygon
                    Dim myPointCollection As PointCollection = New PointCollection
                    myPointCollection.Add(New Point(0, 0))
                    myPointCollection.Add(New Point(10, 5))
                    myPointCollection.Add(New Point(0, 10))
                    Rect.Points = myPointCollection
                    myPointCollection = Nothing

                    Rect.Width = 9
                    Rect.Height = 9

                    Rect.Fill = Brushes.DarkGreen 'myBrush
                    Rect.Tag = Drv.ID
                    Rect.ToolTip = "Arrêter le driver"
                    AddHandler Rect.MouseDown, AddressOf StopDriver
                    Graph = Rect
                    Rect = Nothing
                Else
                    Dim Rect As New Rectangle
                    Rect.Width = 9
                    Rect.Height = 9

                    Rect.Fill = Brushes.Red 'myBrush
                    Rect.ToolTip = "Démarrer le driver"
                    Rect.Tag = Drv.ID
                    AddHandler Rect.MouseDown, AddressOf StartDriver
                    Graph = Rect
                    Rect = Nothing
                End If


                '**************************** POPUP ***************************
                Dim label As New Label
                If Drv.Enable = True Then
                    label.Foreground = New SolidColorBrush(Colors.White)
                Else
                    label.Foreground = New SolidColorBrush(Colors.Black)
                End If
                label.Content = Drv.Nom
                tool.Content = "Nom: " & Drv.Nom & vbCrLf
                tool.Content &= "Enable " & Drv.Enable & vbCrLf
                tool.Content &= "Description: " & Drv.Description & vbCrLf
                tool.Content &= "Version: " & Drv.Version & vbCrLf
                If Drv.Modele <> "" And Drv.Modele <> "@" Then tool.Content &= "Modele: " & Drv.Modele & vbCrLf

                Dim tl As New ToolTip
                Dim imgpopup As New Image
                Dim stkpopup As New StackPanel
                tl.Foreground = System.Windows.Media.Brushes.White
                tl.Background = System.Windows.Media.Brushes.WhiteSmoke
                tl.BorderBrush = System.Windows.Media.Brushes.Black
                imgpopup.Width = 45
                imgpopup.Height = 45
                imgpopup.Source = img.Source
                stkpopup.Children.Add(imgpopup)
                stkpopup.Children.Add(tool)
                tl.Content = stkpopup
                label.ToolTip = tl
                tl = Nothing
                imgpopup = Nothing
                stkpopup = Nothing

                '*************************** CLIC DROIT **************************
                Dim ctxMenu As New ContextMenu
                ctxMenu.Foreground = System.Windows.Media.Brushes.Black
                ctxMenu.Background = System.Windows.Media.Brushes.LightGray
                ctxMenu.BorderBrush = System.Windows.Media.Brushes.Black
                If Drv.Enable = False Then
                    Dim mnu1 As New MenuItem
                    mnu1.Header = "Enable"
                    mnu1.Tag = 3
                    mnu1.Uid = Drv.ID
                    AddHandler mnu1.Click, AddressOf MnuitemDrv_Click
                    ctxMenu.Items.Add(mnu1)
                    mnu1 = Nothing
                Else
                    Dim mnu2 As New MenuItem
                    mnu2.Header = "Disable"
                    mnu2.Tag = 4
                    mnu2.Uid = Drv.ID
                    AddHandler mnu2.Click, AddressOf MnuitemDrv_Click
                    ctxMenu.Items.Add(mnu2)
                    mnu2 = Nothing
                End If
                If Drv.IsConnect = False Then
                    Dim mnu3 As New MenuItem
                    mnu3.Header = "Démarrer"
                    mnu3.Tag = 0
                    mnu3.Uid = Drv.ID
                    AddHandler mnu3.Click, AddressOf MnuitemDrv_Click
                    ctxMenu.Items.Add(mnu3)
                    If Drv.Enable = False Then mnu3.IsEnabled = False
                    mnu3 = Nothing
                Else
                    Dim mnu4 As New MenuItem
                    mnu4.Header = "Arrêter"
                    mnu4.Tag = 1
                    mnu4.Uid = Drv.ID
                    AddHandler mnu4.Click, AddressOf MnuitemDrv_Click
                    ctxMenu.Items.Add(mnu4)
                    mnu4 = Nothing
                End If
                Dim mnu5 As New MenuItem
                mnu5.Header = "Modifier"
                mnu5.Tag = 2
                mnu5.Uid = Drv.ID
                AddHandler mnu5.Click, AddressOf MnuitemDrv_Click
                ctxMenu.Items.Add(mnu5)
                mnu5 = Nothing
                label.ContextMenu = ctxMenu

                stack.Children.Add(img)
                stack.Children.Add(Graph)
                stack.Children.Add(label)

                newchild.Header = stack
                newchild.Uid = Drv.ID

                Dim marg As New Thickness(-12, 0, 0, 0)
                newchild.Margin = marg
                TreeViewDriver.Items.Add(newchild)
                marg = Nothing

                newchild = Nothing
                stack = Nothing
                ctxMenu = Nothing
                label = Nothing
                Graph = Nothing
            Next

            Drv = Nothing
            ListeDrivers = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffDriver: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Gère les menus click droit sur les drivers
    Private Sub MnuitemDrv_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Select Case sender.tag
                Case 0 'Démarrer
                    If myService.ReturnDriverByID(IdSrv, sender.uid).Enable = True Then
                        'myService.StartDriver(IdSrv, sender.uid)
                        'AffDriver()

                        Me.Cursor = Cursors.Wait
                        myService.StartDriver(IdSrv, sender.uid)
                        Dim OK As Boolean = False
                        Dim t As DateTime = DateTime.Now
                        Do While DateTime.Now < t.AddSeconds(10) And OK = False
                            OK = myService.ReturnDriverByID(IdSrv, sender.uid).IsConnect
                            Thread.Sleep(1000)
                        Loop

                        If OK = False Then
                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le driver n'a pas pu être démarré, veuillez consulter le log pour en connaitre la raison", "INFO", "")
                        Else
                            AffDriver()
                        End If

                        Me.Cursor = Nothing
                    Else
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le driver ne peut être démarré car sa propriété Enable est à False!", "Avertissement", "")
                    End If
                Case 1 'Arrêter
                    If sender.uid = "DE96B466-2540-11E0-A321-65D7DFD72085" Then
                        MessageBox.Show("Vous ne pouvez pas arrêter ce driver", "Information", MessageBoxButton.OK, MessageBoxImage.Asterisk)
                        Exit Sub
                    End If

                    myService.StopDriver(IdSrv, sender.uid)
                    AffDriver()
                Case 2 'Modifier
                    Dim x As New uDriver(sender.uid)
                    AddHandler x.CloseMe, AddressOf UnloadControl
                    AddHandler x.Loaded, AddressOf ControlLoaded
                    AffControlPage(x)
                Case 3 'Enable
                    Dim x As TemplateDriver = myService.ReturnDriverByID(IdSrv, sender.uid)
                    myService.SaveDriver(IdSrv, sender.uid, x.Nom, True, x.StartAuto, x.IP_TCP, x.Port_TCP, x.IP_UDP, x.Port_UDP, x.COM, x.Refresh, x.Picture, x.Modele, x.AutoDiscover)
                    If IsConnect Then ManagerDrivers.EnableDriver(sender.uid)
                    AffDriver()
                Case 4 'Disable
                    ' si driver démarré, arrete serveur avant
                    If IsConnect Then
                        If MessageBox.Show("Le driver est démarré. Vous ne pouvez pas faire cette action." & vbCrLf & " Voulez-vous l'arrêter pour continuer ?", "Information", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                            myService.StopDriver(IdSrv, sender.uid)
                            AffDriver()
                        Else
                            Exit Sub
                        End If
                    End If
                    Dim x As TemplateDriver = myService.ReturnDriverByID(IdSrv, sender.uid)
                    myService.SaveDriver(IdSrv, sender.uid, x.Nom, False, x.StartAuto, x.IP_TCP, x.Port_TCP, x.IP_UDP, x.Port_UDP, x.COM, x.Refresh, x.Picture, x.Modele, x.AutoDiscover)
                    If IsConnect Then ManagerDrivers.DisableDriver(sender.uid)
                    AffDriver()
            End Select

        Catch ex As Exception
            Me.Cursor = Nothing
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MnuitemDrv_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Afficher la liste des composants
    Public Sub AffDevice()
        Try
            TreeViewDevice.Items.Clear()
            TreeViewDevice.UpdateLayout()

            If IsConnect = False Then Exit Sub

            'ajout David pour le bug de rafraichissement du treeview composant suite aux modifs de seb
            ManagerDevices.LoadDevices()
            If _ListeDevices IsNot Nothing Then
                CntDevice.Content = _ListeDevices.Count & " Composant(s)"
                For Each Dev As TemplateDevice In _ListeDevices
                    Dim newchild As New TreeViewItem
                    Dim stack As New StackPanel
                    Dim img As New Image
                    Dim FlagZone As Boolean = False
                    Dim nomdriver As String = ManagerDrivers.ReturnDriverByID(Dev.DriverID).Nom
                    Dim _nbhisto As Long = 0
                    Dim uri2 As String = MyRep & "\Images\Icones\ZoneNo_32.png"
                    Dim default_imgzone = ImageFromUri(uri2)
                    stack.Orientation = Orientation.Horizontal

                    'If _DevicesAsHisto.ContainsKey(Dev.ID) Then
                    _nbhisto = Dev.CountHisto  'myService.DeviceAsHisto(Dev.ID)
                    'End If

                    'gestion de l'image du composant dans le menu
                    img.Height = 20
                    img.Width = 20
                    If String.IsNullOrEmpty(Trim(Dev.Picture)) = False Then
                        img.Source = ConvertArrayToImage(myService.GetByteFromImage(Dev.Picture), 20)
                    Else
                        Dim bmpImage As New BitmapImage()
                        Dim uri As String = MyRep & "\Images\Icones\Composant_32.png"
                        img.Source = ImageFromUri(uri)
                    End If
                    stack.Children.Add(img)

                    'creation du label
                    Dim label As New Label
                    If Dev.Enable = True Then
                        label.Foreground = New SolidColorBrush(Colors.White)
                        'on verifie si la composant n'est pas à jour depuis au moins lastchangeduree
                        If Dev.LastChangeDuree > 0 Then
                            If DateTime.Compare(Dev.LastChange.AddMinutes(CInt(Dev.LastChangeDuree)), Now) < 0 Then
                                label.Foreground = New SolidColorBrush(Colors.Red)
                            End If
                        End If
                    Else
                        label.Foreground = New SolidColorBrush(Colors.Black)
                    End If
                    label.Content = Dev.Name & " (" & nomdriver & ")"

                    'verification si le device fait parti d'une zone
                    Dim Imagezone As String = ""
                    Dim nomzone As String = ""
                    For Each Zon In _ListeZones
                        If FlagZone = False Then
                            For Each elemnt In Zon.ListElement
                                If elemnt.ElementID = Dev.ID Then
                                    FlagZone = True
                                    Imagezone = Zon.Image
                                    nomzone = Zon.Name
                                    Exit For
                                End If
                            Next
                        Else
                            Exit For
                        End If
                    Next

                    If FlagZone = False Then
                        Dim imgNoZone As New Image
                        imgNoZone.Height = 20
                        imgNoZone.Width = 20
                        imgNoZone.Source = default_imgzone
                        imgNoZone.ToolTip = "Ce composant ne fait pas partie d'une zone"
                        stack.Children.Add(imgNoZone)
                        imgNoZone = Nothing
                    Else
                        Dim imgZone As New Image
                        imgZone.Height = 20
                        imgZone.Width = 20
                        imgZone.ToolTip = nomzone
                        If String.IsNullOrEmpty(Imagezone) = False Then
                            imgZone.Source = ConvertArrayToImage(myService.GetByteFromImage(Imagezone), 20)
                        Else
                            imgZone.Source = ImageFromUri(MyRep & "\Images\Icones\Zone_32.png")
                        End If
                        stack.Children.Add(imgZone)
                    End If

                    '*************************** TOOL TIP **************************
                    Dim tl As New ToolTip
                    tl.Foreground = System.Windows.Media.Brushes.White
                    tl.Background = System.Windows.Media.Brushes.WhiteSmoke
                    tl.BorderBrush = System.Windows.Media.Brushes.Black
                    tl.Tag = Dev.ID

                    Dim stkpopup As New StackPanel

                    Dim imgpopup As New Image
                    imgpopup.Width = 45
                    imgpopup.Height = 45
                    imgpopup.Source = img.Source
                    stkpopup.Children.Add(imgpopup)
                    imgpopup = Nothing

                    Dim tool As New Label
                    tool.Content = "Nom: " & Dev.Name & vbCrLf
                    tool.Content &= "Enable: " & Dev.Enable & vbCrLf
                    tool.Content &= "Description: " & Dev.Description & vbCrLf
                    tool.Content &= "Type: " & Dev.Type.ToString & vbCrLf
                    tool.Content &= "Driver: " & nomdriver & vbCrLf
                    tool.Content &= "Date MAJ: " & Dev.LastChange & vbCrLf
                    tool.Content &= "Nb Histo: " & _nbhisto & vbCrLf
                    tool.Content &= "Value: " & Dev.Value
                    stkpopup.Children.Add(tool)
                    tool = Nothing

                    tl.Content = stkpopup
                    stkpopup = Nothing

                    label.ToolTip = tl

                    '*************************** CLIC DROIT **************************
                    Dim ctxMenu As New ContextMenu
                    ctxMenu.Foreground = System.Windows.Media.Brushes.Black
                    ctxMenu.Background = System.Windows.Media.Brushes.LightGray
                    ctxMenu.BorderBrush = System.Windows.Media.Brushes.Black
                    Dim mnu0 As New MenuItem
                    mnu0.Header = "Modifier"
                    mnu0.Tag = 0
                    mnu0.Uid = Dev.ID
                    AddHandler mnu0.Click, AddressOf MnuitemDev_Click
                    ctxMenu.Items.Add(mnu0)
                    mnu0 = Nothing
                    If Dev.Enable = False Then
                        Dim mnu1 As New MenuItem
                        mnu1.Header = "Enable"
                        mnu1.Tag = 1
                        mnu1.Uid = Dev.ID
                        AddHandler mnu1.Click, AddressOf MnuitemDev_Click
                        ctxMenu.Items.Add(mnu1)
                        mnu1 = Nothing
                    Else
                        Dim mnu2 As New MenuItem
                        mnu2.Header = "Disable"
                        mnu2.Tag = 2
                        mnu2.Uid = Dev.ID
                        AddHandler mnu2.Click, AddressOf MnuitemDev_Click
                        ctxMenu.Items.Add(mnu2)
                        mnu2 = Nothing
                    End If
                    Dim mnu4 As New MenuItem
                    mnu4.Header = "Historique"
                    mnu4.Tag = 4
                    mnu4.Uid = Dev.ID
                    'If _nbhisto <= 0 Then mnu4.IsEnabled = False
                    If _nbhisto <= 0 Then mnu4.FontStyle = System.Windows.FontStyles.Italic
                    AddHandler mnu4.Click, AddressOf MnuitemDev_Click
                    ctxMenu.Items.Add(mnu4)
                    mnu4 = Nothing
                    Dim mnu5 As New MenuItem
                    mnu5.Header = "Supprimer"
                    mnu5.Tag = 5
                    mnu5.Uid = Dev.ID
                    AddHandler mnu5.Click, AddressOf MnuitemDev_Click
                    ctxMenu.Items.Add(mnu5)
                    mnu5 = Nothing
                    Dim mnu6 As New MenuItem
                    mnu6.Header = "Tester"
                    mnu6.Tag = 6
                    mnu6.Uid = Dev.ID
                    AddHandler mnu6.Click, AddressOf MnuitemDev_Click
                    ctxMenu.Items.Add(mnu6)
                    mnu6 = Nothing

                    label.ContextMenu = ctxMenu


                    '********************** Ajout du label au stackpanel puis treeview *********************
                    stack.Children.Add(label)
                    newchild.Header = stack
                    newchild.Foreground = New SolidColorBrush(Colors.White)
                    newchild.Uid = Dev.ID
                    newchild.Margin = New Thickness(-12, 0, 0, 0)
                    TreeViewDevice.Items.Add(newchild)

                    img = Nothing
                    ctxMenu = Nothing
                    label = Nothing
                    newchild = Nothing
                    stack = Nothing
                    tl = Nothing
                Next

                ' ListeDevices = Nothing
            Else
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffDevice: Le serveur n'a renvoyé aucun composant...", "ERREUR", "")
            End If
            'ListeZones = Nothing
            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffDevice: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    'Retourne True si le device a plusieurs histo sur des propriétés différents
    Private Function AsMultiHisto(ByVal Deviceid As String, ByVal x As List(Of HoMIDom.HoMIDom.Historisation)) As Boolean
        Dim retour As Boolean = False
        Try
            Dim tmp As Boolean = False
            Dim _name As String = ""

            For Each _histo As HoMIDom.HoMIDom.Historisation In x
                If _histo.IdDevice = Deviceid And _name = "" Then
                    _name = _histo.Nom
                ElseIf _histo.IdDevice = Deviceid And _name <> _histo.Nom Then
                    Return True
                End If
            Next
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AsMultiHisto: " & ex.Message, "ERREUR", "")
        End Try
        Return retour
    End Function

    'Gère les menus click droit sur les devices
    Private Sub MnuitemDev_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Select Case sender.tag
                Case 0 'Modifier
                    Dim x As New uDevice(Classe.EAction.Modifier, sender.uid)
                    AddHandler x.CloseMe, AddressOf UnloadControl
                    AddHandler x.Loaded, AddressOf ControlLoaded
                    AffControlPage(x)
                Case 1 'Enable
                    Dim x As TemplateDevice = myService.ReturnDeviceByID(IdSrv, sender.uid)
                    '                    myService.SaveDevice(IdSrv, sender.uid, x.Name, x.Adresse1, True, x.Solo, x.DriverID, x.Type.ToString, x.Refresh, x.IsHisto, x.RefreshHisto, x.Purge, x.MoyJour, x.MoyHeure)
                    myService.SaveDevice(IdSrv, sender.uid, x.Name, x.Adresse1, True, x.Solo, x.DriverID, x.Type.ToString, x.Refresh, x.IsHisto, x.RefreshHisto, x.Purge, x.MoyJour, x.MoyHeure, x.Adresse2, x.Picture, x.Modele, x.Description, x.LastChangeDuree, x.LastEtat, x.Correction, x.Formatage, x.Precision, x.ValueMax, x.ValueMin, x.ValueDef, x.Commandes, x.Unit, x.Puissance, x.AllValue, x.VariablesOfDevice)
                    If IsConnect Then ManagerDevices.EnableDevice(sender.uid)
                    AffDevice()
                Case 2 'Disable
                    Dim x As TemplateDevice = myService.ReturnDeviceByID(IdSrv, sender.uid)
                    '                    myService.SaveDevice(IdSrv, sender.uid, x.Name, x.Adresse1, False, x.Solo, x.DriverID, x.Type.ToString, x.Refresh, x.IsHisto, x.RefreshHisto, x.Purge, x.MoyJour, x.MoyHeure)
                    myService.SaveDevice(IdSrv, sender.uid, x.Name, x.Adresse1, False, x.Solo, x.DriverID, x.Type.ToString, x.Refresh, x.IsHisto, x.RefreshHisto, x.Purge, x.MoyJour, x.MoyHeure, x.Adresse2, x.Picture, x.Modele, x.Description, x.LastChangeDuree, x.LastEtat, x.Correction, x.Formatage, x.Precision, x.ValueMax, x.ValueMin, x.ValueDef, x.Commandes, x.Unit, x.Puissance, x.AllValue, x.VariablesOfDevice)
                    If IsConnect Then ManagerDevices.DisableDevice(sender.uid)
                    AffDevice()
                Case 4 'Graphe
                    Dim Devices As New List(Of Dictionary(Of String, String))
                    Dim y As New Dictionary(Of String, String)
                    y.Add(sender.uid, "Value")
                    Devices.Add(y)

                    Dim x As New uHisto(Devices)
                    x.Uid = System.Guid.NewGuid.ToString()
                    x.Width = CanvasRight.ActualWidth - 20
                    x.Height = CanvasRight.ActualHeight - 20
                    x._with = CanvasRight.ActualWidth - 20
                    AddHandler x.CloseMe, AddressOf UnloadControl
                    CanvasRight.Children.Clear()
                    CanvasRight.Children.Add(x)

                Case 5 'Supprimer
                    DeleteElement(sender.uid, 1)
                    AffDevice()
                Case 6 'Tester
                    Try
                        If myService.ReturnDeviceByID(IdSrv, sender.uid).Enable = False Then
                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Vous ne pouvez pas exécuter de commandes car le device n'est pas activé (propriété Enable)!", "Erreur", "")
                            Exit Sub
                        End If

                        Dim y As New uTestDevice(sender.uid)
                        y.Uid = System.Guid.NewGuid.ToString()
                        AddHandler y.CloseMe, AddressOf UnloadControl
                        Window1.CanvasUser.Children.Add(y)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur Tester: " & ex.Message, "Erreur", "")
                    End Try

            End Select

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MnuitemDev_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Afficher la liste des macros
    Public Sub AffScene()
        Try
            TreeViewMacro.Items.Clear()
            TreeViewMacro.UpdateLayout()

            If IsConnect = False Then Exit Sub

            Dim ListeMacros As List(Of Macro) = myService.GetAllMacros(IdSrv)
            CntMacro.Content = ListeMacros.Count & " Macro(s)"

            For Each Mac As Macro In ListeMacros
                Dim newchild As New TreeViewItem
                Dim stack As New StackPanel
                Dim img As New Image
                Dim uri As String = ""

                stack.Orientation = Orientation.Horizontal

                img.Height = 20
                img.Width = 20

                Dim label As New Label
                If Mac.Enable = True Then
                    label.Foreground = New SolidColorBrush(Colors.White)
                Else
                    label.Foreground = New SolidColorBrush(Colors.Black)
                End If
                label.Content = Mac.Nom

                Dim ctxMenu As New ContextMenu
                ctxMenu.Foreground = System.Windows.Media.Brushes.Black
                ctxMenu.Background = System.Windows.Media.Brushes.LightGray
                ctxMenu.BorderBrush = System.Windows.Media.Brushes.Black
                Dim mnu0 As New MenuItem
                mnu0.Header = "Modifier"
                mnu0.Tag = 0
                mnu0.Uid = Mac.ID
                AddHandler mnu0.Click, AddressOf MnuitemMac_Click
                ctxMenu.Items.Add(mnu0)
                Dim mnu1 As New MenuItem
                mnu1.Header = "Supprimer"
                mnu1.Tag = 1
                mnu1.Uid = Mac.ID
                AddHandler mnu1.Click, AddressOf MnuitemMac_Click
                ctxMenu.Items.Add(mnu1)
                Dim mnu2 As New MenuItem
                mnu2.Header = "Executer"
                mnu2.Tag = 2
                mnu2.Uid = Mac.ID
                AddHandler mnu2.Click, AddressOf MnuitemMac_Click
                ctxMenu.Items.Add(mnu2)
                label.ContextMenu = ctxMenu

                uri = MyRep & "\Images\Icones\Macro_32.png"
                img.Source = ImageFromUri(uri)

                stack.Children.Add(img)
                stack.Children.Add(label)

                newchild.Foreground = New SolidColorBrush(Colors.White)
                newchild.Header = stack
                newchild.Uid = Mac.ID
                Dim marg As New Thickness(-12, 0, 0, 0)
                newchild.Margin = marg
                TreeViewMacro.Items.Add(newchild)
            Next

            ListeMacros = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffScene: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Gère les menus click droit sur les macros
    Private Sub MnuitemMac_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Select Case sender.tag
                Case 0 'Modifier
                    Dim x As New uMacro(Classe.EAction.Modifier, sender.uid)
                    x.Width = CanvasRight.ActualWidth - 20
                    x._Width = CanvasRight.ActualWidth - 20
                    x.Height = CanvasRight.ActualHeight - 20
                    x._Height = CanvasRight.ActualHeight - 20
                    AddHandler x.CloseMe, AddressOf UnloadControl
                    AddHandler x.Loaded, AddressOf ControlLoaded
                    AffControlPage(x)
                Case 1 'Supprimer
                    DeleteElement(sender.uid, 5)
                Case 2 'Executer
                    myService.RunMacro(IdSrv, sender.uid)
            End Select

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MnuitemMac_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Afficher la liste des triggers
    Public Sub AffTrigger()
        Try
            TreeViewTrigger.Items.Clear()
            TreeViewTrigger.UpdateLayout()

            If IsConnect = False Then Exit Sub

            Dim ListeTriggers As List(Of Trigger) = myService.GetAllTriggers(IdSrv)
            CntTrigger.Content = ListeTriggers.Count & " Trigger(s)"

            For Each Trig In ListeTriggers
                Dim newchild As New TreeViewItem
                Dim stack As New StackPanel
                Dim img As New Image
                Dim uri As String = ""

                stack.Orientation = Orientation.Horizontal

                img.Height = 20
                img.Width = 20

                Dim label As New Label
                If Trig.Enable = True Then
                    label.Foreground = New SolidColorBrush(Colors.White)
                Else
                    label.Foreground = New SolidColorBrush(Colors.Black)
                End If
                label.Content = Trig.Nom

                Dim ctxMenu As New ContextMenu
                ctxMenu.Foreground = System.Windows.Media.Brushes.Black
                ctxMenu.Background = System.Windows.Media.Brushes.LightGray
                ctxMenu.BorderBrush = System.Windows.Media.Brushes.Black
                Dim mnu0 As New MenuItem
                mnu0.Header = "Modifier"
                mnu0.Tag = 0
                mnu0.Uid = Trig.ID
                AddHandler mnu0.Click, AddressOf MnuitemTrig_Click
                ctxMenu.Items.Add(mnu0)
                Dim mnu1 As New MenuItem
                mnu1.Header = "Supprimer"
                mnu1.Tag = 1
                mnu1.Uid = Trig.ID
                AddHandler mnu1.Click, AddressOf MnuitemTrig_Click
                ctxMenu.Items.Add(mnu1)
                label.ContextMenu = ctxMenu

                If Trig.Type = 0 Then
                    uri = MyRep & "\Images\Icones\Trigger_clock_32.png"
                Else
                    uri = MyRep & "\Images\Icones\Trigger_device_32.png"
                End If

                img.Source = ImageFromUri(uri)

                stack.Children.Add(img)
                stack.Children.Add(label)

                newchild.Foreground = New SolidColorBrush(Colors.White)
                newchild.Header = stack
                newchild.Uid = Trig.ID
                Dim marg As New Thickness(-12, 0, 0, 0)
                newchild.Margin = marg
                TreeViewTrigger.Items.Add(newchild)
            Next

            ListeTriggers = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Afftrigger: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    'Gère les menus click droit sur les triggers
    Private Sub MnuitemTrig_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Select Case sender.tag
                Case 0 'Modifier
                    Dim _Trig As Trigger = myService.ReturnTriggerById(IdSrv, sender.uid)

                    If _Trig IsNot Nothing Then
                        If _Trig.Type = Trigger.TypeTrigger.TIMER Then
                            Dim x As New uTriggerTimer(Classe.EAction.Modifier, sender.uid)
                            AddHandler x.CloseMe, AddressOf UnloadControl
                            AddHandler x.Loaded, AddressOf ControlLoaded
                            AffControlPage(x)
                        Else
                            Dim x As New uTriggerDevice(Classe.EAction.Modifier, sender.uid)
                            AddHandler x.CloseMe, AddressOf UnloadControl
                            AddHandler x.Loaded, AddressOf ControlLoaded
                            AffControlPage(x)
                        End If
                        _Trig = Nothing
                    End If
                Case 1 'Supprimer
                    DeleteElement(sender.uid, 4)
            End Select

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MnuitemTrig_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Afficher la liste des historisations
    Public Sub AffHisto()
        Try
            TreeViewHisto.Items.Clear()
            TreeViewHisto.UpdateLayout()

            Dim x As New List(Of HoMIDom.HoMIDom.Historisation)
            'Dim ListeDevices As List(Of TemplateDevice) = myService.GetAllDevices(IdSrv)

            x = myService.GetAllListHisto(IdSrv)

            If _ListeDevices IsNot Nothing Then
                For Each _dev As TemplateDevice In _ListeDevices ' ListeDevices
                    Dim IsNode As Boolean = False
                    Dim parent = New TreeViewItem
                    Dim y As New CheckBox
                    Dim trv As Boolean = False

                    If AsMultiHisto(_dev.ID, x) Then 'le device a différents historique
                        parent.Header = _dev.Name
                        parent.Foreground = New SolidColorBrush(Colors.White)
                        IsNode = True
                    Else
                        y.Content = _dev.Name
                        y.Foreground = New SolidColorBrush(Colors.Black)
                        y.Background = New SolidColorBrush(Colors.DarkGray)
                        y.BorderBrush = New SolidColorBrush(Colors.Black)
                        y.Margin = New Thickness(-15, 1, 0, 0)
                        'y.IsEnabled = False
                        y.Uid = _dev.ID
                        AddHandler y.MouseDoubleClick, AddressOf MnuitemHisto_Click
                    End If

                    If x IsNot Nothing Then
                        For i As Integer = 0 To x.Count - 1

                            Dim a As Historisation = x.Item(i)

                            'AddHandler y.MouseDoubleClick, AddressOf MnuitemHisto_Click

                            If _dev.ID = a.IdDevice And IsNode = False Then
                                y.Tag = a.Nom
                                y.Foreground = New SolidColorBrush(Colors.White)
                                y.IsEnabled = True

                                '*************************** CLIC DROIT **************************
                                Dim ctxMenu As New ContextMenu
                                ctxMenu.Foreground = System.Windows.Media.Brushes.Black
                                ctxMenu.Background = System.Windows.Media.Brushes.LightGray
                                ctxMenu.BorderBrush = System.Windows.Media.Brushes.Black
                                Dim mnu0 As New MenuItem
                                mnu0.Header = "Afficher"
                                mnu0.Tag = 0
                                mnu0.Uid = _dev.ID
                                AddHandler mnu0.Click, AddressOf MnuitemHisto_Click
                                ctxMenu.Items.Add(mnu0)
                                y.ContextMenu = ctxMenu

                                '*************************** TOOL TIP **************************
                                Dim tl As New ToolTip
                                tl.Foreground = System.Windows.Media.Brushes.White
                                tl.Background = System.Windows.Media.Brushes.WhiteSmoke
                                tl.BorderBrush = System.Windows.Media.Brushes.Black
                                Dim _nbhisto As Long = _dev.CountHisto
                                Dim nomdriver As String = myService.ReturnDriverByID(IdSrv, _dev.DriverID).Nom
                                Dim stkpopup As New StackPanel
                                Dim tool As New Label
                                tool.Content = "Nom: " & _dev.Name & vbCrLf
                                tool.Content &= "Enable " & _dev.Enable & vbCrLf
                                tool.Content &= "Description: " & _dev.Description & vbCrLf
                                tool.Content &= "Type: " & _dev.Type.ToString & vbCrLf
                                tool.Content &= "Driver: " & nomdriver & vbCrLf
                                tool.Content &= "Date MAJ: " & _dev.LastChange & vbCrLf
                                tool.Content &= "Nb Histo: " & _nbhisto & vbCrLf
                                tool.Content &= "Value: " & _dev.Value
                                stkpopup.Children.Add(tool)
                                tool = Nothing
                                tl.Content = stkpopup
                                stkpopup = Nothing
                                y.ToolTip = tl

                                AddHandler y.MouseDoubleClick, AddressOf MnuitemHisto_Click

                                TreeViewHisto.Items.Add(y)
                                trv = True
                                mnu0 = Nothing
                                ctxMenu = Nothing
                            ElseIf _dev.ID = a.IdDevice And IsNode = True Then
                                Dim y1 As New CheckBox
                                y1.Content = _dev.Name
                                y1.Foreground = New SolidColorBrush(Colors.White)
                                y1.Background = New SolidColorBrush(Colors.DarkGray)
                                y1.BorderBrush = New SolidColorBrush(Colors.Black)
                                y1.Margin = New Thickness(-15, 1, 0, 0)
                                y1.Uid = _dev.ID
                                y1.Content = a.Nom
                                y1.Tag = a.Nom

                                '*************************** CLIC DROIT **************************
                                Dim ctxMenu As New ContextMenu
                                ctxMenu.Foreground = System.Windows.Media.Brushes.Black
                                ctxMenu.Background = System.Windows.Media.Brushes.LightGray
                                ctxMenu.BorderBrush = System.Windows.Media.Brushes.Black
                                Dim mnu0 As New MenuItem
                                mnu0.Header = "Afficher"
                                mnu0.Tag = 0
                                mnu0.Uid = _dev.ID
                                AddHandler mnu0.Click, AddressOf MnuitemHisto_Click
                                ctxMenu.Items.Add(mnu0)
                                y1.ContextMenu = ctxMenu

                                AddHandler y1.MouseDoubleClick, AddressOf MnuitemHisto_Click

                                parent.Items.Add(y1)
                                trv = True

                                mnu0 = Nothing
                                ctxMenu = Nothing
                                y1 = Nothing
                            End If
                            a = Nothing
                        Next
                        If IsNode Then
                            TreeViewHisto.Items.Add(parent)
                        End If
                    End If
                    If trv = False Then TreeViewHisto.Items.Add(y)
                    parent = Nothing
                    y = Nothing
                Next

                x = Nothing
                'ListeDevices = Nothing
            End If

            TreeViewHisto.Items.SortDescriptions.Clear()
            TreeViewHisto.Items.SortDescriptions.Add(New SortDescription("Content", ListSortDirection.Ascending))
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffHisto: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    'Gère les menus click droit sur les histo
    Private Sub MnuitemHisto_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If


            Dim Devices As New List(Of Dictionary(Of String, String))
            Dim y As New Dictionary(Of String, String)
            y.Add(sender.uid, "Value")
            Devices.Add(y)

            Dim x As New uHisto(Devices)
            x.Uid = System.Guid.NewGuid.ToString()
            x.Width = CanvasRight.ActualWidth - 20
            x.Height = CanvasRight.ActualHeight - 20
            x._with = CanvasRight.ActualWidth - 20
            AddHandler x.CloseMe, AddressOf UnloadControl
            CanvasRight.Children.Clear()
            CanvasRight.Children.Add(x)

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MnuitemHisto_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Double Clic sur le treeview
    Private Sub TreeView_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles TreeViewDriver.MouseDoubleClick, TreeViewDevice.MouseDoubleClick, TreeViewZone.MouseDoubleClick, TreeViewUser.MouseDoubleClick, TreeViewTrigger.MouseDoubleClick, TreeViewMacro.MouseDoubleClick, TreeViewHisto.MouseDoubleClick
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            If sender.SelectedItem IsNot Nothing Then
                If sender.SelectedItem.uid Is Nothing Then Exit Sub

                Me.Cursor = Cursors.Wait
                Select Case sender.Name
                    Case "TreeViewDriver"  'driver
                        Dim x As New uDriver(sender.SelectedItem.uid)
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AddHandler x.Loaded, AddressOf ControlLoaded
                        AffControlPage(x)
                    Case "TreeViewDevice"  'device
                        Dim x As New uDevice(Classe.EAction.Modifier, sender.SelectedItem.uid)
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AddHandler x.Loaded, AddressOf ControlLoaded
                        AffControlPage(x)
                    Case "TreeViewZone"  'zone
                        Dim x As New uZone(Classe.EAction.Modifier, sender.SelectedItem.uid)
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AddHandler x.Loaded, AddressOf ControlLoaded
                        AffControlPage(x)
                    Case "TreeViewUser"  'user
                        Dim x As New uUser(Classe.EAction.Modifier, sender.SelectedItem.uid)
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AddHandler x.Loaded, AddressOf ControlLoaded
                        AffControlPage(x)
                    Case "TreeViewTrigger"  'trigger
                        Dim _Trig As Trigger = myService.ReturnTriggerById(IdSrv, sender.SelectedItem.uid)

                        If _Trig IsNot Nothing Then
                            If _Trig.Type = Trigger.TypeTrigger.TIMER Then
                                Dim x As New uTriggerTimer(Classe.EAction.Modifier, sender.SelectedItem.uid)
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AddHandler x.Loaded, AddressOf ControlLoaded
                                AffControlPage(x)
                            Else
                                Dim x As New uTriggerDevice(Classe.EAction.Modifier, sender.SelectedItem.uid)
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AddHandler x.Loaded, AddressOf ControlLoaded
                                AffControlPage(x)
                            End If
                            _Trig = Nothing
                        End If
                    Case "TreeViewMacro"  'macros
                        Dim x As New uMacro(Classe.EAction.Modifier, sender.SelectedItem.uid)
                        x.Width = CanvasRight.ActualWidth - 20
                        x._Width = CanvasRight.ActualWidth - 20
                        x.Height = CanvasRight.ActualHeight - 20
                        x._Height = CanvasRight.ActualHeight - 20
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AddHandler x.Loaded, AddressOf ControlLoaded
                        AffControlPage(x)
                    Case "TreeViewHisto" 'histo
                        Me.Cursor = Nothing
                End Select

            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub TreeView_MouseDoubleClick : " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub TreeView_SelectedItemChanged(ByVal sender As System.Object, ByVal e As System.Windows.RoutedPropertyChangedEventArgs(Of System.Object)) Handles TreeViewDriver.SelectedItemChanged, TreeViewDevice.SelectedItemChanged, TreeViewZone.SelectedItemChanged, TreeViewUser.SelectedItemChanged, TreeViewTrigger.SelectedItemChanged, TreeViewMacro.SelectedItemChanged, TreeViewHisto.SelectedItemChanged
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            If Mouse.LeftButton = MouseButtonState.Pressed Then
                If sender.SelectedItem IsNot Nothing Then
                    Dim effects As DragDropEffects
                    Dim obj As New DataObject()
                    obj.SetData(GetType(String), sender.SelectedItem.uid)
                    effects = DragDrop.DoDragDrop(sender, obj, DragDropEffects.Copy Or DragDropEffects.Move)
                End If
            End If

            If sender.SelectedItem IsNot Nothing Then
                If sender.SelectedItem.uid Is Nothing Then Exit Sub
            End If

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub TreeViewG_SelectedItemChanged: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

#End Region

#Region "Drivers"

    Private Sub StopDriver(ByVal sender As Object, ByVal e As Input.MouseEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            If sender.tag = "DE96B466-2540-11E0-A321-65D7DFD72085" Then
                MessageBox.Show("Vous ne pouvez pas arrêter ce driver", "Information", MessageBoxButton.OK, MessageBoxImage.Asterisk)
                Exit Sub
            End If

            Me.Cursor = Cursors.Wait

            myService.StopDriver(IdSrv, sender.tag)

            Dim OK As Boolean = True
            Dim t As DateTime = DateTime.Now
            Do While DateTime.Now < t.AddSeconds(10) And OK = True
                OK = myService.ReturnDriverByID(IdSrv, sender.tag).IsConnect
                Thread.Sleep(1000)
            Loop

            If OK = True Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le driver n'a pas pu être arrêté, veuillez consulter le log pour en connaitre la raison", "INFO", "")
            Else
                AffDriver()
            End If

            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub StopDriver: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub StartDriver(ByVal sender As Object, ByVal e As Input.MouseEventArgs)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            If myService.ReturnDriverByID(IdSrv, sender.tag).Enable = True Then
                Me.Cursor = Cursors.Wait

                myService.StartDriver(IdSrv, sender.tag)

                Dim OK As Boolean = False
                Dim t As DateTime = DateTime.Now
                Do While DateTime.Now < t.AddSeconds(10) And OK = False
                    OK = myService.ReturnDriverByID(IdSrv, sender.tag).IsConnect
                    Thread.Sleep(1000)
                Loop

                If OK = False Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le driver n'a pas pu être démarré, veuillez consulter le log pour en connaitre la raison", "INFO", "")
                Else
                    AffDriver()
                End If

                Me.Cursor = Nothing
            Else
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le driver ne peut être démarré car sa propriété Enable est à False!", "Avertissement", "")
            End If
        Catch ex As Exception
            Me.Cursor = Nothing
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub BtnStart_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

#End Region

#Region "MainMenu"

    Private Sub ShowMainMenu()
        Try
            MainMenu = New uMainMenu

            MainMenu.Uid = "MAINMENU"
            AddHandler MainMenu.menu_contextmenu, AddressOf MainMenuContextmenu
            AddHandler MainMenu.menu_gerer, AddressOf MainMenuGerer
            AddHandler MainMenu.menu_delete, AddressOf MainMenuDelete
            AddHandler MainMenu.menu_edit, AddressOf MainMenuEdit
            AddHandler MainMenu.menu_create, AddressOf MainMenuNew
            AddHandler MainMenu.menu_autre, AddressOf MainMenuAutre
            CanvasRight.Children.Add(MainMenu)
            WMainMenu = CanvasRight.ActualWidth / 2 - (MainMenu.Width / 2)
            HMainMenu = CanvasRight.ActualHeight / 2 - (MainMenu.Height / 2)
            Canvas.SetLeft(MainMenu, WMainMenu)
            Canvas.SetTop(MainMenu, HMainMenu)

            AnimationApparition(MainMenu)
            flagShowMainMenu = True
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub ShowMainMenu: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub MainMenuContextmenu(ByVal index As String)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If
            Me.Cursor = Cursors.Wait

            'on affiche le context menu correspondant
            Select Case index.Substring(0, 3)
                Case "drv"
                    Tabcontrol1.SelectedIndex = 0
                    Select Case index
                        Case "drv_gerer"
                        Case "drv_modifier"
                            Try
                                _MainMenuAction = 1
                                Dim x As New uSelectElmt("Choisir {TITLE} à éditer", "tag_driver")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu Driver Modifier: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "cpt"
                    Tabcontrol1.SelectedIndex = 1
                    Select Case index
                        Case "cpt_gerer"
                        Case "cpt_modifier"
                            Try
                                _MainMenuAction = 1
                                Dim x As New uSelectElmt("Choisir {TITLE} à éditer", "tag_composant")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu COMPOSANTS Modifier: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "cpt_ajouter"
                            Try
                                Dim x As New uDevice(Classe.EAction.Nouveau, "")
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu COMPOSANTS Ajouter: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "cpt_supprimer"
                            Try
                                _MainMenuAction = 2
                                Dim x As New uSelectElmt("Choisir {TITLE} à supprimer", "tag_composant")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu COMPOSANTS Supprimer: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "zon"
                    Tabcontrol1.SelectedIndex = 2
                    Select Case index
                        Case "zon_gerer"
                        Case "zon_modifier"
                            Try
                                _MainMenuAction = 1
                                Dim x As New uSelectElmt("Choisir {TITLE} à éditer", "tag_zone")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu ZONE Modifier: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "zon_ajouter"
                            Try
                                Tabcontrol1.SelectedIndex = 2
                                Dim x As New uZone(Classe.EAction.Nouveau, "")
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu ZONE Ajouter: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "zon_supprimer"
                            Try
                                _MainMenuAction = 2
                                Dim x As New uSelectElmt("Choisir {TITLE} à supprimer", "tag_zone")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu ZONE Supprimer: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "usr"
                    Tabcontrol1.SelectedIndex = 3
                    Select Case index
                        Case "usr_gerer"
                        Case "usr_modifier"
                            Try
                                _MainMenuAction = 1
                                Dim x As New uSelectElmt("Choisir {TITLE} à éditer", "tag_user")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu USER Modifier: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "usr_ajouter"
                            Try
                                Tabcontrol1.SelectedIndex = 3
                                Dim x As New uZone(Classe.EAction.Nouveau, "")
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu USER Ajouter: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "usr_supprimer"
                            Try
                                _MainMenuAction = 2
                                Dim x As New uSelectElmt("Choisir {TITLE} à supprimer", "tag_user")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu USER Supprimer: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "trg"
                    Tabcontrol1.SelectedIndex = 4
                    Select Case index
                        Case "trg_gerer"
                        Case "trg_modifier"
                            Try
                                _MainMenuAction = 1
                                Dim x As New uSelectElmt("Choisir {TITLE} à éditer", "tag_trigger")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu TRIGGER Modifier: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "trg_ajouter_timer"
                            Try
                                Tabcontrol1.SelectedIndex = 3
                                Dim x As New uTriggerTimer(0, "")
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu TRIGGER Ajouer Timer: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "trg_ajouter_composant"
                            Try
                                Tabcontrol1.SelectedIndex = 3
                                Dim x As New uTriggerDevice(0, "")
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu TRIGGER Ajouter Composant: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "trg_supprimer"
                            Try
                                _MainMenuAction = 2
                                Dim x As New uSelectElmt("Choisir {TITLE} à supprimer", "tag_trigger")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu TRIGGER Supprimer: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "mac"
                    Tabcontrol1.SelectedIndex = 5
                    Select Case index
                        Case "mac_gerer"
                        Case "mac_modifier"
                            Try
                                _MainMenuAction = 1
                                Dim x As New uSelectElmt("Choisir {TITLE} à éditer", "tag_macro")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu MACRO Modifier: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "mac_ajouter"
                            Try
                                Tabcontrol1.SelectedIndex = 3
                                Dim x As New uMacro(Classe.EAction.Nouveau, "")
                                x.Width = CanvasRight.ActualWidth - 20
                                x._Width = CanvasRight.ActualWidth - 20
                                x.Height = CanvasRight.ActualHeight - 20
                                x._Height = CanvasRight.ActualHeight - 20
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu MACRO ajouter: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "mac_supprimer"
                            Try
                                _MainMenuAction = 2
                                Dim x As New uSelectElmt("Choisir {TITLE} à supprimer", "tag_macro")
                                AddHandler x.CloseMe, AddressOf UnloadSelectElmt
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu MACRO Supprimer: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "cfg"
                    Select Case index
                        Case "cfg_log"
                            Try
                                Dim x As New uLog
                                x.Uid = System.Guid.NewGuid.ToString()
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                x.Width = CanvasRight.ActualWidth - 100
                                x.Height = CanvasRight.ActualHeight - 50
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu LOG: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "cfg_configurer"
                            Try
                                Dim x As New uConfigServer
                                x.Uid = System.Guid.NewGuid.ToString()
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu Configurer: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "cfg_importer"
                            Try
                                ' Configure open file dialog box
                                Dim dlg As New Microsoft.Win32.OpenFileDialog()
                                dlg.FileName = "Homidom" ' Default file name
                                dlg.DefaultExt = ".xml" ' Default file extension
                                dlg.Filter = "Fichier de configuration (.xml)|*.xml" ' Filter files by extension
                                ' Show open file dialog box
                                Dim result As Boolean = dlg.ShowDialog()
                                ' Process open file dialog box results
                                If result = True Then
                                    ' Open document
                                    Dim filename As String = dlg.FileName
                                    If MessageBox.Show("Etes vous sur que le serveur puisse accéder au fichier " & filename & " que ce soit en local ou via le réseau, sinon il ne pourra pas l'importer!", "Import Config", MessageBoxButton.OKCancel, MessageBoxImage.Question) = MessageBoxResult.Cancel Then Exit Sub
                                    Dim retour As String = myService.ImportConfig(IdSrv, filename)
                                    If retour <> "0" Then
                                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, retour, "Erreur import config", "")
                                    Else
                                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "L'import du fichier de configuration a été effectué, l'ancien fichier a été renommé en .old, veuillez redémarrer le serveur pour prendre en compte cette nouvelle configuration", "Import config", "")
                                    End If
                                End If
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre config_importer: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "cfg_exporter"
                            Try
                                ' Configure open file dialog box
                                Dim dlg As New Microsoft.Win32.SaveFileDialog()
                                dlg.FileName = "" ' Default file name
                                dlg.DefaultExt = ".xml" ' Default file extension
                                dlg.Filter = "Fichier de configuration (.xml)|*.xml" ' Filter files by extension
                                ' Show open file dialog box
                                Dim result As Boolean = dlg.ShowDialog()
                                ' Process open file dialog box results
                                If result = True Then
                                    ' Open document
                                    Dim filename As String = dlg.FileName
                                    Dim retour As String = myService.ExportConfig(IdSrv)
                                    If retour.StartsWith("ERREUR") Then
                                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, retour, "Erreur export config", "")
                                    Else
                                        Dim TargetFile As StreamWriter
                                        TargetFile = New StreamWriter(filename, False)
                                        TargetFile.Write(retour)
                                        TargetFile.Close()
                                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "L'export du fichier de configuration a été effectué", "Export config", "")
                                    End If
                                End If
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu config_exporter: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "cfg_sauvegarder"
                            Try
                                If IsConnect = False Then
                                    MessageBox.Show("Impossible d'enregistrer la config car le serveur n'est pas connecté !!", "Erreur", MessageBoxButton.OK, MessageBoxImage.Asterisk)
                                    Exit Sub
                                End If

                                Dim retour As String = myService.SaveConfig(IdSrv)
                                If retour <> "0" Then
                                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'enregistrement veuillez consulter le log", "HomIAdmin", "")
                                Else
                                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Enregistrement effectué", "HomIAdmin", "")
                                    FlagChange = False
                                End If
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu config_sauvegarder: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "qit"
                    Select Case index
                        Case "qit_quitter"
                            Try
                                If IsConnect = True And FlagChange Then
                                    If MessageBox.Show("Voulez-vous enregistrer la configuration avant de quitter?", "HomIAdmin", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                                        If IsConnect = False Then
                                            MessageBox.Show("Impossible d'enregistrer la configuration car le serveur n'est pas connecté !!", "Erreur", MessageBoxButton.OK, MessageBoxImage.Asterisk)
                                        Else
                                            Dim retour As String = myService.SaveConfig(IdSrv)
                                            If retour <> "0" Then
                                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'enregistrement veuillez consulter le log", "HomIAdmin", "")
                                            Else
                                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Enregistrement effectué", "HomIAdmin", "")
                                                FlagChange = False
                                            End If
                                        End If
                                    End If

                                End If
                                Me.Close()
                                End
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu Quitter: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "qit_stop"
                            Try
                                If MessageBox.Show("Confirmer vous l'arrêt du serveur?", "HomIAdmin", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                                    myService.Stop(IdSrv)
                                    RefreshTreeView()
                                End If
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu STOP: " & ex.Message, "ERREUR", "")
                            End Try
                        Case "qit_start"
                            Try
                                If MessageBox.Show("Confirmer vous le démarrage du serveur?", "HomIAdmin", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                                    myService.Start()
                                    RefreshTreeView()
                                End If
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu START: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "var"
                    Try
                        Dim x As New uVariables
                        x.Uid = System.Guid.NewGuid.ToString()
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu variables gérer: " & ex.Message, "ERREUR", "")
                    End Try
                Case "hst"
                    Tabcontrol1.SelectedIndex = 6
                    Select Case index
                        Case "hst_gérer"
                    End Select
                Case "aid"
                    Select Case index
                        Case "aid_gérer"
                            Try
                                Dim x As New uHelp
                                x.Uid = System.Guid.NewGuid.ToString()
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu aide gérer: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select
                Case "mod"
                    Select Case index
                        Case "mod_ajouter"
                            Try
                                Dim x As New uModuleSimple()
                                AddHandler x.CloseMe, AddressOf UnloadControl
                                AffControlPage(x)
                            Catch ex As Exception
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuContextmenu Module ajouter: " & ex.Message, "ERREUR", "")
                            End Try
                    End Select


            End Select
            ShowTreeView()
            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuGerer: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub MainMenuGerer(ByVal index As String)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Me.Cursor = Cursors.Wait

            'affichage du treeview correspondant
            Select Case index
                Case "tag_driver" : Tabcontrol1.SelectedIndex = 0
                Case "tag_composant" : Tabcontrol1.SelectedIndex = 1
                Case "tag_zone" : Tabcontrol1.SelectedIndex = 2
                Case "tag_user" : Tabcontrol1.SelectedIndex = 3
                Case "tag_trigger" : Tabcontrol1.SelectedIndex = 4
                Case "tag_macro" : Tabcontrol1.SelectedIndex = 5
            End Select
            ShowTreeView()
            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuGerer: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub MainMenuNew(ByVal index As String)
        Try
            _MainMenuAction = 0

            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Select Case index
                Case "tag_composant"
                    Try
                        Tabcontrol1.SelectedIndex = 1
                        Dim x As New uDevice(Classe.EAction.Nouveau, "")
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuNew Composants: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_zone"
                    Try
                        Tabcontrol1.SelectedIndex = 2
                        Dim x As New uZone(Classe.EAction.Nouveau, "")
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuNew Zone: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_user"
                    Try
                        Tabcontrol1.SelectedIndex = 3
                        Dim x As New uUser(Classe.EAction.Nouveau, "")
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuNew User: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_triggertimer"
                    Try
                        Tabcontrol1.SelectedIndex = 4
                        Dim x As New uTriggerTimer(0, "")
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuNew TriggerTime: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_triggerdevice"
                    Try
                        Tabcontrol1.SelectedIndex = 4
                        Dim x As New uTriggerDevice(0, "")
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuNew TriggerDevice: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_macro"
                    Try
                        Tabcontrol1.SelectedIndex = 5
                        Dim x As New uMacro(Classe.EAction.Nouveau, "")
                        x.Width = CanvasRight.ActualWidth - 20
                        x._Width = CanvasRight.ActualWidth - 20
                        x.Height = CanvasRight.ActualHeight - 20
                        x._Height = CanvasRight.ActualHeight - 20
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuNew Macro: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_module"
                    Try
                        Dim x As New uModuleSimple()
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuNew Module: " & ex.Message, "ERREUR", "")
                    End Try
            End Select
            ShowTreeView()
            _MainMenuAction = -1
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'exécution de MainMenuNew: " & ex.ToString, "Erreur Admin", "")
        End Try
    End Sub

    Private Sub MainMenuDelete(ByVal index As String)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            _MainMenuAction = 2

            Dim x As New uSelectElmt("Choisir {TITLE} à supprimer", index)
            AddHandler x.CloseMe, AddressOf UnloadSelectElmt
            AffControlPage(x)
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'exécution de MainMenuDelete: " & ex.ToString, "Erreur Admin", "")
        End Try
    End Sub

    Private Sub MainMenuEdit(ByVal index As String)
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            _MainMenuAction = 1

            Dim x As New uSelectElmt("Choisir {TITLE} à éditer", index)
            AddHandler x.CloseMe, AddressOf UnloadSelectElmt
            AffControlPage(x)
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'exécution de MainMenuEdit: " & ex.ToString, "Erreur Admin", "")
        End Try
    End Sub

    Private Sub MainMenuAutre(ByVal index As String)
        Try
            If IsConnect = False And index <> "tag_quitter" Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Me.Cursor = Cursors.Wait
            Select Case index.ToLower
                Case "tag_histo" : Tabcontrol1.SelectedIndex = 6
                Case "tag_histo_import"
                    Try
                        Dim result As String
                        Dim openFileDialog1 As New Microsoft.Win32.OpenFileDialog()
                        openFileDialog1.Filter = "Fichier texte CSV (*.csv)|*.csv"
                        openFileDialog1.FilterIndex = 2
                        openFileDialog1.RestoreDirectory = True
                        Dim DlgResult As Forms.DialogResult
                        DlgResult = openFileDialog1.ShowDialog()
                        If DlgResult = Forms.DialogResult.OK Or DlgResult = -1 Then
                            If openFileDialog1.FileName.EndsWith("csv") Then
                                ' Le fichier est apparemment ok.
                                ' On regarde si l'admin est connecté en local, auquel cas
                                ' pas besoin d'uploader le fichier sur le serveur
                                Dim serveur As String = myChannelFactory.Endpoint.Address.ToString()
                                serveur = Mid(serveur, 8, serveur.Length - 8)
                                serveur = Split(serveur, "/", 2)(0)
                                serveur = Split(serveur, ":", 2)(0)
                                If serveur = "localhost" Then
                                    Mouse.OverrideCursor = Cursors.Wait
                                    result = myService.ImportHisto(openFileDialog1.FileName)
                                    Mouse.OverrideCursor = Nothing
                                Else
                                    Dim fileInfo = New FileInfo(openFileDialog1.FileName)
                                    Upload(IdSrv, openFileDialog1.FileName, "Histo\\" & fileInfo.Name)
                                    Mouse.OverrideCursor = Cursors.Wait
                                    result = myService.ImportHisto("Histo\" & fileInfo.Name)
                                    Mouse.OverrideCursor = Nothing
                                End If
                            Else
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Format non supporté!", "Erreur", "")
                                Exit Try
                            End If
                            If result.Contains("ERR") = True Then
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, result, "Erreur", "")
                            Else
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Importation terminée.", "HoMIAdmiN", "")
                            End If
                        End If
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre import_histo: " & ex.Message, "Erreur", "")
                    End Try
                Case "tag_config_log"
                    Try
                        If IsConnect = False Then
                            MessageBox.Show("Impossible d'afficher le log car le serveur n'est pas connecté !!", "Erreur", MessageBoxButton.OK, MessageBoxImage.Asterisk)
                            Exit Sub
                        End If
                        Dim x As New uLog
                        x.Uid = System.Guid.NewGuid.ToString()
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        x.Width = CanvasRight.ActualWidth - 100
                        x.Height = CanvasRight.ActualHeight - 50
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre log: " & ex.Message, "ERREUR", "")
                    End Try

                Case "tag_multimedia" 'Multimedia Playlist
                    Try
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Désolé cette fonctionnalité n'est pas encore disponible...", "INFO", "")
                        Exit Select
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre Playlist: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_multimedia_template"
                    Try
                        Dim frm As New WTelecommandeNew()
                        frm.ShowDialog()
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_config" 'Configurer le serveur
                    Try
                        Dim x As New uConfigServer
                        x.Uid = System.Guid.NewGuid.ToString()
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre config: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_var" 'Configurer le serveur
                    Try
                        Dim x As New uVariables
                        x.Uid = System.Guid.NewGuid.ToString()
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre var: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_aide" 'Aide
                    Try
                        Dim x As New uHelp
                        x.Uid = System.Guid.NewGuid.ToString()
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre aide: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_composant_nvx" 'nvx composants
                    Try
                        Dim x As New uNewDevice()
                        AddHandler x.CloseMe, AddressOf UnloadControl
                        AddHandler x.CreateNewDevice, AddressOf CreateNewDevice
                        AffControlPage(x)
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre nvx composant: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_quitter"
                    Try
                        If IsConnect = True And FlagChange Then

                            If MessageBox.Show("Voulez-vous enregistrer la configuration avant de quitter?", "HomIAdmin", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                                Try
                                    If IsConnect = False Then
                                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.MESSAGE, "Impossible d'enregistrer la configuration car le serveur n'est pas connecté !!", "Erreur", "")
                                    Else
                                        Dim retour As String = myService.SaveConfig(IdSrv)
                                        If retour <> "0" Then
                                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'enregistrement veuillez consulter le log", "HomIAdmin", "")
                                        Else
                                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Enregistrement effectué", "HomIAdmin", "")
                                        End If
                                    End If
                                Catch ex As Exception
                                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Quitter: " & ex.Message, "ERREUR", "")
                                End Try

                            End If
                            FlagChange = False
                        End If
                        Me.Close()
                        End
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre Quitter: " & ex.Message, "ERREUR", "")
                    End Try
                Case "tag_quitter_stop"
                    If MessageBox.Show("Confirmer vous l'arrêt du serveur?", "HomIAdmin", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                        myService.Stop(IdSrv)
                        RefreshTreeView()
                    End If
                Case "tag_quitter_start"
                    If MessageBox.Show("Confirmer vous le démarrage du serveur?", "HomIAdmin", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.Yes Then
                        myService.Start()
                        RefreshTreeView()
                    End If
                Case "tag_config_exporter" 'Exporter le fichier de config
                    Try
                        ' Configure open file dialog box
                        Dim dlg As New Microsoft.Win32.SaveFileDialog()
                        dlg.FileName = "" ' Default file name
                        dlg.DefaultExt = ".xml" ' Default file extension
                        dlg.Filter = "Fichier de configuration (.xml)|*.xml" ' Filter files by extension

                        ' Show open file dialog box
                        Dim result As Boolean = dlg.ShowDialog()

                        ' Process open file dialog box results
                        If result = True Then
                            ' Open document
                            Dim filename As String = dlg.FileName
                            Dim retour As String = myService.ExportConfig(IdSrv)
                            If retour.StartsWith("ERREUR") Then
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, retour, "Erreur export config", "")
                            Else
                                Dim TargetFile As StreamWriter
                                TargetFile = New StreamWriter(filename, False)
                                TargetFile.Write(retour)
                                TargetFile.Close()
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "L'export du fichier de configuration a été effectué", "Export config", "")
                            End If
                        End If
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre config_exporter: " & ex.Message, "ERREUR", "")
                    End Try

                Case "tag_config_importer" 'Importer le fichier de config
                    Try
                        ' Configure open file dialog box
                        Dim dlg As New Microsoft.Win32.OpenFileDialog()
                        dlg.FileName = "Homidom" ' Default file name
                        dlg.DefaultExt = ".xml" ' Default file extension
                        dlg.Filter = "Fichier de configuration (.xml)|*.xml" ' Filter files by extension

                        ' Show open file dialog box
                        Dim result As Boolean = dlg.ShowDialog()

                        ' Process open file dialog box results
                        If result = True Then
                            ' Open document
                            Dim filename As String = dlg.FileName

                            If MessageBox.Show("Etes vous sur que le serveur puisse accéder au fichier " & filename & " que ce soit en local ou via le réseau, sinon il ne pourra pas l'importer!", "Import Config", MessageBoxButton.OKCancel, MessageBoxImage.Question) = MessageBoxResult.Cancel Then
                                Exit Sub
                            End If

                            Dim retour As String = myService.ImportConfig(IdSrv, filename)
                            If retour <> "0" Then
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, retour, "Erreur import config", "")
                            Else
                                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "L'import du fichier de configuration a été effectué, l'ancien fichier a été renommé en .old, veuillez redémarrer le serveur pour prendre en compte cette nouvelle configuration", "Import config", "")
                            End If
                        End If
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre config_importer: " & ex.Message, "ERREUR", "")
                    End Try

                Case "tag_config_sauvegarder"
                    Try
                        If IsConnect = False Then
                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Impossible d'enregistrer la config car le serveur n'est pas connecté !!", "Erreur", "Window1.MainMenuAutre")
                            Exit Sub
                        End If

                        Dim retour As String = myService.SaveConfig(IdSrv)
                        If retour <> "0" Then
                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'enregistrement veuillez consulter le log", "HomIAdmin", "")
                        Else
                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Enregistrement effectué", "HomIAdmin", "")
                            FlagChange = False
                        End If
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre config_sauvegarder: " & ex.Message, "ERREUR", "")
                    End Try

            End Select
            ShowTreeView()
            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub MainMenuAutre: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

#End Region

#Region "Fenêtres et actions autres"

    Private Sub UnloadSelectElmt(ByVal Objet As Object)
        Try
            flagShowMainMenu = False

            If Objet.retour = "CANCEL" Then
                CanvasRight.Children.Clear()
                CanvasRight.UpdateLayout()

                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()

                Me.UpdateLayout()
                ShowMainMenu()
            Else
                If IsConnect = False Then
                    Serveur_notconnected_action()
                    Exit Sub
                End If

                Select Case _MainMenuAction
                    Case 1 'Edit
                        Try
                            If Objet.retour IsNot Nothing Then

                                Me.Cursor = Cursors.Wait
                                Select Case Objet.Type
                                    Case 0 'driver
                                        Dim x As New uDriver(Objet.retour)
                                        x.Uid = System.Guid.NewGuid.ToString()
                                        AddHandler x.CloseMe, AddressOf UnloadControl
                                        CanvasRight.Children.Clear()
                                        CanvasRight.UpdateLayout()

                                        GC.Collect()
                                        GC.WaitForPendingFinalizers()
                                        GC.Collect()

                                        Me.UpdateLayout()
                                        CanvasRight.Children.Add(x)

                                        AnimationApparition(x)
                                    Case 1 'device
                                        Dim x As New uDevice(Classe.EAction.Modifier, Objet.retour)
                                        x.Uid = System.Guid.NewGuid.ToString()
                                        AddHandler x.CloseMe, AddressOf UnloadControl
                                        CanvasRight.Children.Clear()
                                        CanvasRight.UpdateLayout()

                                        GC.Collect()
                                        GC.WaitForPendingFinalizers()
                                        GC.Collect()

                                        Me.UpdateLayout()
                                        CanvasRight.Children.Add(x)

                                        AnimationApparition(x)
                                    Case 2 'zone
                                        Dim x As New uZone(Classe.EAction.Modifier, Objet.retour)
                                        x.Uid = System.Guid.NewGuid.ToString()
                                        AddHandler x.CloseMe, AddressOf UnloadControl
                                        CanvasRight.Children.Clear()
                                        CanvasRight.UpdateLayout()

                                        GC.Collect()
                                        GC.WaitForPendingFinalizers()
                                        GC.Collect()

                                        Me.UpdateLayout()
                                        CanvasRight.Children.Add(x)

                                        AnimationApparition(x)
                                    Case 3 'user
                                        Dim x As New uUser(Classe.EAction.Modifier, Objet.retour)
                                        x.Uid = System.Guid.NewGuid.ToString()
                                        AddHandler x.CloseMe, AddressOf UnloadControl
                                        CanvasRight.Children.Clear()
                                        CanvasRight.UpdateLayout()

                                        GC.Collect()
                                        GC.WaitForPendingFinalizers()
                                        GC.Collect()

                                        Me.UpdateLayout()
                                        CanvasRight.Children.Add(x)

                                        AnimationApparition(x)

                                    Case 4 'trigger
                                        Dim _Trig As Trigger = myService.ReturnTriggerById(IdSrv, Objet.retour)

                                        If _Trig IsNot Nothing Then
                                            If _Trig.Type = Trigger.TypeTrigger.TIMER Then
                                                Dim x As New uTriggerTimer(Classe.EAction.Modifier, Objet.retour)
                                                x.Uid = System.Guid.NewGuid.ToString()
                                                AddHandler x.CloseMe, AddressOf UnloadControl
                                                CanvasRight.Children.Clear()
                                                CanvasRight.UpdateLayout()

                                                GC.Collect()
                                                GC.WaitForPendingFinalizers()
                                                GC.Collect()

                                                Me.UpdateLayout()
                                                CanvasRight.Children.Add(x)

                                                AnimationApparition(x)

                                            Else
                                                Dim x As New uTriggerDevice(Classe.EAction.Modifier, Objet.retour)
                                                x.Uid = System.Guid.NewGuid.ToString()
                                                AddHandler x.CloseMe, AddressOf UnloadControl
                                                CanvasRight.Children.Clear()
                                                CanvasRight.UpdateLayout()

                                                GC.Collect()
                                                GC.WaitForPendingFinalizers()
                                                GC.Collect()

                                                Me.UpdateLayout()
                                                CanvasRight.Children.Add(x)

                                                AnimationApparition(x)
                                            End If
                                            _Trig = Nothing
                                        End If
                                    Case 5 'macros

                                        Dim x As New uMacro(Classe.EAction.Modifier, Objet.retour)
                                        x.Uid = System.Guid.NewGuid.ToString()
                                        x.Width = CanvasRight.ActualWidth - 20
                                        x._Width = CanvasRight.ActualWidth - 20
                                        x.Height = CanvasRight.ActualHeight - 20
                                        x._Height = CanvasRight.ActualHeight - 20
                                        AddHandler x.CloseMe, AddressOf UnloadControl
                                        CanvasRight.Children.Clear()
                                        CanvasRight.UpdateLayout()

                                        GC.Collect()
                                        GC.WaitForPendingFinalizers()
                                        GC.Collect()

                                        Me.UpdateLayout()
                                        CanvasRight.Children.Add(x)

                                        AnimationApparition(x)
                                    Case 6 'histo

                                End Select

                                ShowTreeView()

                                Me.Cursor = Nothing
                            End If
                        Catch ex As Exception
                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub TreeView_MouseDoubleClick: " & ex.Message, "ERREUR", "")
                        End Try
                    Case 2 'Supprimer
                        DeleteElement(Objet.retour, Objet.Type)
                End Select
                Me.UpdateLayout()
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub UnloadSelectElmt: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub DeleteElement(ByVal ID As String, ByVal Type As Integer)
        Try
            If ID IsNot Nothing Then
                Me.Cursor = Cursors.Wait
                Dim retour As Integer
                Dim _retour As New List(Of String)
                _retour = myService.CanDelete(IdSrv, ID)

                If _retour(0).StartsWith("ERREUR") Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, _retour(0), "Erreur CanDelete", "")
                    Exit Sub
                Else
                    If _retour(0) <> "0" Then
                        Dim a As String = "Attention !! Confirmez vous de supprimer cet élément car il est utilisé dans: " & vbCrLf
                        For i As Integer = 0 To _retour.Count - 2
                            a = a & _retour(i) & vbCrLf
                        Next
                        If myService.DeviceAsHisto(ID) Then
                            a = a & "- Historiques" & vbCrLf
                        End If
                        If MessageBox.Show(a, "Suppression", MessageBoxButton.YesNo, MessageBoxImage.Question) = MessageBoxResult.No Then
                            Me.Cursor = Nothing
                            Exit Sub
                        End If
                    End If
                End If


                Select Case Type
                    Case 0
                    Case 1
                        retour = myService.DeleteDevice(IdSrv, ID)
                        AffDevice()
                    Case 2
                        retour = myService.DeleteZone(IdSrv, ID)
                        AffZone()
                    Case 3
                        retour = myService.DeleteUser(IdSrv, ID)
                        AffUser()
                    Case 4
                        retour = myService.DeleteTrigger(IdSrv, ID)
                        AffTrigger()
                    Case 5
                        retour = myService.DeleteMacro(IdSrv, ID)
                        AffScene()
                End Select

                If retour = -2 Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Vous ne pouvez pas supprimer cet élément!", "Erreur", "")
                    Me.Cursor = Nothing
                    Exit Sub
                Else
                    FlagChange = True
                End If

            Else
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Veuillez sélectionner un élément à supprimer!")
                Me.Cursor = Nothing
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR de la suppression: " & ex.ToString, "ERREUR", "")
        End Try
        CanvasRight.Children.Clear()
        ShowMainMenu()
        Me.Cursor = Nothing
    End Sub

    Private Sub ControlLoaded()
        Try
            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub ControlLoaded: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    ''' <summary>
    ''' Affiche le usercontrol et l'anime
    ''' </summary>
    ''' <param name="Objet"></param>
    ''' <remarks></remarks>
    Private Sub AffControlPage(ByVal Objet As Object)
        Try
            If CanvasRight.Children.Count > 0 And Objet IsNot Nothing Then

                Objet3 = Objet

                For i As Integer = 0 To CanvasRight.Children.Count - 1
                    If My.Settings.AnimationDuration > 0 Then
                        Dim myDoubleAnimation As DoubleAnimation = New DoubleAnimation()
                        myDoubleAnimation.From = 1.0
                        myDoubleAnimation.To = 0.0
                        myDoubleAnimation.Duration = New Duration(TimeSpan.FromMilliseconds(My.Settings.AnimationDuration))
                        Dim myStoryboard As Storyboard = New Storyboard()
                        myStoryboard.Children.Add(myDoubleAnimation)
                        AddHandler myStoryboard.Completed, AddressOf AffControlPageSuite

                        Storyboard.SetTarget(myDoubleAnimation, CanvasRight.Children.Item(i))
                        Storyboard.SetTargetProperty(myDoubleAnimation, New PropertyPath(UserControl.OpacityProperty))
                        myStoryboard.Begin()

                        myStoryboard = Nothing
                    Else
                        AffControlPageSuite()
                    End If
                Next
            End If

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffControlPage: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Dim Objet3 As Object = Nothing

    Private Sub AffControlPageSuite()
        Try
            If Objet3 IsNot Nothing Then
                CanvasRight.Children.Clear()
                CanvasRight.UpdateLayout()

                GC.Collect()
                GC.WaitForPendingFinalizers()
                GC.Collect()

                Me.UpdateLayout()

                CanvasRight.Children.Add(Objet3)
                AnimationApparition(Objet3)
                Objet3 = Nothing
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AffControlPageSuite: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Décharger une fenêtre suivant son Id
    Public Sub UnloadControl(ByVal MyControl As Object, Optional ByVal Cancel As Boolean = False)
        Try
            Me.Cursor = Cursors.Wait
            flagShowMainMenu = False
            If Cancel = False Then RefreshTreeView()

            Dim myDoubleAnimation As DoubleAnimation = New DoubleAnimation()
            myDoubleAnimation.From = 1.0
            myDoubleAnimation.To = 0.0
            myDoubleAnimation.Duration = New Duration(TimeSpan.FromMilliseconds(My.Settings.AnimationDuration))
            Dim myStoryboard As Storyboard = New Storyboard()
            myStoryboard.Children.Add(myDoubleAnimation)
            AddHandler myStoryboard.Completed, AddressOf StoryBoardFinish

            Storyboard.SetTarget(myDoubleAnimation, MyControl)
            Storyboard.SetTargetProperty(myDoubleAnimation, New PropertyPath(UserControl.OpacityProperty))
            myStoryboard.Begin()

            myStoryboard = Nothing
            myDoubleAnimation = Nothing
            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub UnloadControl: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub AnimationApparition(ByVal Objet As Object)
        Try
            If Objet IsNot Nothing Then
                Dim myDoubleAnimation As DoubleAnimation = New DoubleAnimation()
                myDoubleAnimation.From = 0.0
                myDoubleAnimation.To = 1.0
                myDoubleAnimation.Duration = New Duration(TimeSpan.FromMilliseconds(My.Settings.AnimationDuration))
                Dim myStoryboard As Storyboard = New Storyboard()
                myStoryboard.Children.Add(myDoubleAnimation)

                Storyboard.SetTarget(myDoubleAnimation, Objet)
                Storyboard.SetTargetProperty(myDoubleAnimation, New PropertyPath(UserControl.OpacityProperty))
                myStoryboard.Begin()

                myDoubleAnimation = Nothing
                myStoryboard = Nothing
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub AnimationApparition: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub BtnGenereGraph_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnGenereGraph.Click
        Try
            If IsConnect = False Then
                Serveur_notconnected_action()
                Exit Sub
            End If

            Me.Cursor = Cursors.Wait

            Dim Devices As New List(Of Dictionary(Of String, String))

            For i As Integer = 0 To TreeViewHisto.Items.Count - 1
                Dim chk As CheckBox

                If TreeViewHisto.Items(i).GetType.ToString.Contains("CheckBox") Then
                    chk = TreeViewHisto.Items(i)

                    If chk.IsChecked = True Then
                        Dim y As New Dictionary(Of String, String)
                        y.Add(chk.Uid, chk.Tag)
                        Devices.Add(y)
                    End If

                Else 'il a des enfants
                    Dim trv1 As TreeViewItem = TreeViewHisto.Items(i)
                    For j1 As Integer = 0 To trv1.Items.Count - 1
                        chk = trv1.Items(j1)

                        If chk.IsChecked = True Then
                            Dim y As New Dictionary(Of String, String)
                            y.Add(chk.Uid, chk.Tag)
                            Devices.Add(y)
                        End If

                    Next
                End If
            Next

            Dim x As New uHisto(Devices)
            x.Uid = System.Guid.NewGuid.ToString()
            x.Width = CanvasRight.ActualWidth - 20
            x.Height = CanvasRight.ActualHeight - 20
            x._with = CanvasRight.ActualWidth - 20
            AddHandler x.CloseMe, AddressOf UnloadControl
            CanvasRight.Children.Clear()

            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

            Me.UpdateLayout()

            CanvasRight.Children.Add(x)

            Devices = Nothing
            Me.Cursor = Nothing
        Catch ex As Exception
            Me.Cursor = Nothing
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub BtnGenereGraph_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub Ellipse1_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles Ellipse1.MouseDown
        Try
            If IsConnect = False Then PageConnexion()
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Ellipse1_MouseDown: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub Tabcontrol1_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles Tabcontrol1.SelectionChanged
        RefreshTreeView()
    End Sub

    Private Sub LOG_PreviewMouseMove(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseEventArgs) Handles LOG.PreviewMouseMove, ImgLog.PreviewMouseMove
        Try
            If IsConnect = True Then
                Dim list As List(Of String) = myService.GetLastLogs
                Dim a As String = ""
                For i As Integer = 0 To list.Count - 1
                    a &= list(i)
                    If (i <> list.Count - 1) Then a &= vbCrLf
                Next
                LOG.ToolTip = a
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub LOG_PreviewMouseMove: " & ex.Message, "ERREUR", "")
        End Try
    End Sub


#End Region

    Private Sub ImgNewDevice_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgNewDevice.MouseDown
        Try
            If e.ClickCount = 1 Then
                Dim x As New uNewDevice()
                AddHandler x.CloseMe, AddressOf UnloadControl
                AddHandler x.CreateNewDevice, AddressOf CreateNewDevice
                AffControlPage(x)
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub ImgNewDevice_MouseDown: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub CreateNewDevice(ByVal MyObject As Object)
        Try
            UnloadControl(MyObject)

            Tabcontrol1.SelectedIndex = 1
            Dim x As New uDevice(Classe.EAction.Nouveau, "")
            AddHandler x.CloseMe, AddressOf UnloadControl
            AffControlPage(x)
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub CreateNewDevice: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub LHS_MouseDoubleClick(sender As System.Object, e As System.Windows.Input.MouseButtonEventArgs) Handles LHS.MouseDoubleClick
        Try
            Dim frm As New WTelecommandeNew
            frm.ShowDialog()
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
        End Try
    End Sub

    Private Sub Window1_SizeChanged(sender As Object, e As SizeChangedEventArgs) Handles Window1.SizeChanged

        If flagShowMainMenu = True Then
            WMainMenu = CanvasRight.ActualWidth / 2 - (MainMenu.Width / 2)
            HMainMenu = CanvasRight.ActualHeight / 2 - (MainMenu.Height / 2)
            Canvas.SetLeft(MainMenu, WMainMenu)
            Canvas.SetTop(MainMenu, HMainMenu)
        End If

    End Sub

    Private Sub RfrDev_MouseDown(sender As System.Object, e As System.Windows.Input.MouseButtonEventArgs) Handles RfrDev.MouseDown
        AffDevice()
    End Sub
End Class



