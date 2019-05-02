Imports System.IO

Public Class uSelectElmt
    Dim _retour As String
    Dim _Type As Integer
    Dim _TypeElement As String

    Public Event CloseMe(ByVal MyObject As Object)

    Public ReadOnly Property Retour As String
        Get
            Return _retour
        End Get
    End Property
    Public ReadOnly Property Type As Integer
        Get
            Return _Type
        End Get
    End Property

    Public Sub New(ByVal Title As String, ByVal TypeElement As String)

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        Try
            ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
            _TypeElement = TypeElement
            Select Case TypeElement
                Case "tag_driver"
                    Title = Replace(Title, "{TITLE}", "un Driver")
                    _Type = 0
                    For Each driver As HoMIDom.HoMIDom.TemplateDriver In myService.GetAllDrivers(IdSrv)
                        Dim x As New uElement
                        x.ID = driver.ID
                        If IO.File.Exists(driver.Picture) Then
                            x.Image = driver.Picture
                        Else
                            x.Image = ".\images\icones\Driver_32.png"
                        End If
                        x.Title = driver.Nom
                        x.Width = 300
                        ListBox1.Items.Add(x)
                        x = Nothing
                    Next
                Case "tag_composant"
                    Title = Replace(Title, "{TITLE}", "un Composant")
                    _Type = 1
                    For Each device In myService.GetAllDevices(IdSrv)
                        Dim x As New uElement
                        x.ID = device.ID
                        If IO.File.Exists(device.Picture) Then
                            x.Image = device.Picture
                        Else
                            x.Image = ".\images\icones\Composant_32.png"
                        End If
                        x.Title = device.Name
                        x.Width = 300
                        ListBox1.Items.Add(x)
                        x = Nothing
                    Next
                Case "tag_zone"
                    Title = Replace(Title, "{TITLE}", "une Zone")
                    _Type = 2
                    For Each zone In myService.GetAllZones(IdSrv)
                        Dim x As New uElement
                        x.ID = zone.ID
                        x.Image = zone.Icon
                        x.Title = zone.Name
                        x.Width = 300
                        ListBox1.Items.Add(x)
                        x = Nothing
                    Next
                Case "tag_user"
                    Title = Replace(Title, "{TITLE}", "un Utilisateur")
                    _Type = 3
                    For Each user In myService.GetAllUsers(IdSrv)
                        Dim x As New uElement
                        x.ID = user.ID
                        x.Image = user.Image
                        x.Title = user.Nom
                        x.Width = 300
                        ListBox1.Items.Add(x)
                        x = Nothing
                        BtnArchiver.Visibility = Visibility.Hidden   'Ajout JPS
                    Next
                Case "tag_trigger"
                    Title = Replace(Title, "{TITLE}", "un Trigger")
                    _Type = 4
                    For Each trigger In myService.GetAllTriggers(IdSrv)
                        Dim x As New uElement
                        x.ID = trigger.ID
                        x.Title = trigger.Nom
                        x.Width = 300
                        ListBox1.Items.Add(x)
                        x = Nothing
                    Next
                Case "tag_macro"
                    Title = Replace(Title, "{TITLE}", "une Macro")
                    _Type = 5
                    For Each macro In myService.GetAllMacros(IdSrv)
                        Dim x As New uElement
                        x.ID = macro.ID
                        x.Title = macro.Nom
                        x.Width = 300
                        ListBox1.Items.Add(x)
                        x = Nothing
                    Next
            End Select

            LblTitle.Content = Title
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'exécution de NewSelectElement: " & ex.ToString, "Erreur Admin", "")
        End Try
    End Sub

    Private Sub BtnCancel_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnCancel.Click
        _retour = "CANCEL"
        RaiseEvent CloseMe(Me)
    End Sub

    Private Sub BtnOK_Click(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnOK.Click
        Try
            If ListBox1.SelectedItem IsNot Nothing Then
                Dim stk As uElement = ListBox1.SelectedItem
                _retour = stk.ID
                RaiseEvent CloseMe(Me)
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uSelectElmt BtnOK_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub ListBox1_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ListBox1.MouseDoubleClick
        Try
            If ListBox1.SelectedItem IsNot Nothing Then
                Dim stk As uElement = ListBox1.SelectedItem
                _retour = stk.ID
                RaiseEvent CloseMe(Me)
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uSelectElmt ListBox1_MouseDoubleClick: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub ListBox1_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListBox1.SelectionChanged
        Try
            For Each Objet As uElement In e.RemovedItems
                Objet.IsSelect = False
            Next
            For Each Objet As uElement In e.AddedItems
                Objet.IsSelect = True
            Next
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uSelectElmt ListBox1_SelectionChanged: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub BtnArchiver_Click(sender As Object, e As RoutedEventArgs) Handles BtnArchiver.Click
        Try
            Dim newFile As FileInfo
            Dim NomFichier As String = ""

            Select Case _TypeElement
                Case "tag_driver"
                    NomFichier = My.Application.Info.DirectoryPath & "\Config\DiversList.csv"
                Case "tag_composant"
                    NomFichier = My.Application.Info.DirectoryPath & "\Config\DevicesList.csv"
                Case "tag_zone"
                    NomFichier = My.Application.Info.DirectoryPath & "\Config\ZoneList.csv"
                Case "tag_trigger"
                    NomFichier = My.Application.Info.DirectoryPath & "\Config\TriggerList.csv"
                Case "tag_macro"
                    NomFichier = My.Application.Info.DirectoryPath & "\Config\MacroList.csv"
            End Select
            newFile = New FileInfo(NomFichier)
            If newFile.Exists Then
                SauverCopiefichier(NomFichier)
                newFile.Delete()
            End If
            Select Case _TypeElement
                Case "tag_driver"
                    SauverDiverList(NomFichier)
                Case "tag_composant"
                    SauverDeviceList(NomFichier)
                Case "tag_zone"
                    SauverZoneList(NomFichier)
                Case "tag_trigger"
                    SauverTriggerList(NomFichier)
                Case "tag_macro"
                    SauverMacroList(NomFichier)
            End Select
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur  BtnArchiver_Click : " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub SauverDiverList(ByVal NomFichier As String)
        'Sauver la liste des Drivers dans un fichier "DiversList.csv"  Ajouter par JPS
        Try
            Dim ListDiver As New List(Of HoMIDom.HoMIDom.TemplateDriver)
            Dim zipFile As FileInfo

            ListDiver.Clear()
            ListDiver = myService.GetAllDrivers(IdSrv)
            zipFile = New FileInfo(NomFichier)
            Dim Fichier As StreamWriter = zipFile.CreateText()
            Fichier.WriteLine("Nom" & ";" & "Enable" & ";" & "IsConnect" & ";" & "StartAuto" & ";" & "ID" & ";" & "Description")
            For i = 0 To ListDiver.Count - 1
                Fichier.WriteLine(ListDiver(i).Nom & ";" &
                ListDiver(i).Enable & ";" &
                ListDiver(i).IsConnect & ";" &
                ListDiver(i).StartAuto & ";" &
                ListDiver(i).ID & ";" &
                ListDiver(i).Description)
            Next
            Fichier.Flush()
            Fichier.Close()
            ListDiver.Clear()
            'newFile = Nothing
            'zipFile = Nothing
            'ListDevice = Nothing
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Sauve DiverList : ", "Réussit", "")

        Catch ex As IO.IOException
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur DiverList.csv : " & ex.ToString, "ERREUR", "")

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur SauverDiverList : " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub SauverZoneList(ByVal NomFichier As String)
        'Sauver la liste des Drivers dans un fichier "ZoneList.csv"  Ajouter par JPS
        Try
            Dim ListZone As New List(Of HoMIDom.HoMIDom.Zone)
            Dim Listdevices As HoMIDom.HoMIDom.TemplateDevice
            Dim Device As New List(Of String)
            Dim zipFile As FileInfo

            ListZone.Clear()
            ListZone = myService.GetAllZones(IdSrv)
            zipFile = New FileInfo(NomFichier)
            Dim Fichier As StreamWriter = zipFile.CreateText()
            Fichier.WriteLine("Nom" & ";" & "ID" & ";" & "Nb Elements" & ";" & "Nom Devices" & ";" & "Adresse1" & ";" &
                              "Type" & ";" & "Enable")
            For i = 0 To ListZone.Count - 1
                Fichier.WriteLine(ListZone(i).Name & ";" &
                ListZone(i).ID & ";" &
                ListZone(i).ListElement.Count)
                Device.Clear()
                Device = myService.GetDeviceInZone(IdSrv, ListZone(i).ID)
                For j = 0 To Device.Count - 1
                    Listdevices = myService.ReturnDeviceByID(IdSrv, Device(j))
                    Fichier.WriteLine(";;;" & Listdevices.Name & ";" &
                                  Listdevices.Adresse1 & ";" &
                                  Listdevices.Type.ToString & ";" &
                                  Listdevices.Enable)
                Next
            Next
            Fichier.Flush()
            Fichier.Close()
            ListZone.Clear()
            Device.Clear()
            'newFile = Nothing
            'zipFile = Nothing
            'ListDevice = Nothing
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Sauve ZoneList : ", "Réussit", "")

        Catch ex As IO.IOException
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur ZoneList.csv : " & ex.ToString, "ERREUR", "")

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur SauverZoneList : " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub SauverTriggerList(ByVal NomFichier As String)
        'Sauver la liste des Triggers dans un fichier "TriggerList.csv"  Ajouter par JPS
        Try
            Dim ListTrigger As New List(Of HoMIDom.HoMIDom.Trigger)
            Dim ListMacro As New List(Of String)
            Dim Macro As HoMIDom.HoMIDom.Macro
            Dim zipFile As FileInfo

            ListTrigger.Clear()
            ListTrigger = myService.GetAllTriggers(IdSrv)
            zipFile = New FileInfo(NomFichier)
            Dim Fichier As StreamWriter = zipFile.CreateText()
            Fichier.WriteLine("Nom Trigger" & ";" & "Type" & ";" & "Enable" & ";" & "ID Trigger" & ";" & "Nom Macro" & ";" &
                              "Enable" & ";" & "ID Macro")
            For i = 0 To ListTrigger.Count - 1
                ListMacro.Clear()
                ListMacro = ListTrigger(i).ListMacro
                For j = 0 To ListMacro.Count - 1
                    Macro = myService.ReturnMacroById(IdSrv, ListMacro(j))
                    If j = 0 Then
                        Fichier.WriteLine(ListTrigger(i).Nom & ";" &
                        ListTrigger(i).Type.ToString & ";" &
                        ListTrigger(i).Enable & ";" &
                        ListTrigger(i).ID & ";" & Macro.Nom & ";" &
                        Macro.Enable & ";" &
                        Macro.ID)
                    Else
                        Fichier.WriteLine(";;;;" & Macro.Nom & ";" &
                        Macro.Enable & ";" &
                        Macro.ID)
                    End If
                Next
            Next
            Fichier.Flush()
            Fichier.Close()
            ListTrigger.Clear()
            ListMacro.Clear()
            'newFile = Nothing
            'zipFile = Nothing
            'ListDevice = Nothing
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Sauve TriggerList : ", "Réussit", "")

        Catch ex As IO.IOException
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur TriggerList.csv : " & ex.ToString, "ERREUR", "")

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur SauverTriggerList : " & ex.ToString, "ERREUR", "")
        End Try


    End Sub

    Private Sub SauverMacroList(ByVal NomFichier As String)
        'Sauver la liste des Macro dans un fichier "MacroList.csv"  Ajouter par JPS
        Try
            Dim ListMacro As New List(Of HoMIDom.HoMIDom.Macro)
            Dim zipFile As FileInfo
            Dim Compt As Integer = -1

            ListMacro.Clear()
            ListMacro = myService.GetAllMacros(IdSrv)
            zipFile = New FileInfo(NomFichier)
            Dim Fichier As StreamWriter = zipFile.CreateText()
            Fichier.WriteLine("Nom" & ";" & "ID" & ";" & "Enable" & ";" & "Then/Else" & ";" & "Type Action" & ";" &
                              "Devices/Variables" & ";" & "Actions/Test" & ";" & "Operateur" & ";" & "Parametres")
            For i = 0 To ListMacro.Count - 1
                Compt = -1
                Fichier.WriteLine(ListMacro(i).Nom & ";" &
                  ListMacro(i).ID & ";" &
                  ListMacro(i).Enable) ' & ";" &               

                For j = 0 To ListMacro(i).ListActions.Count - 1
                    Dim _typ As HoMIDom.HoMIDom.Action.TypeAction = ListMacro(i).ListActions(j).TypeAction
                    Select Case _typ
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionDevice
                            If String.IsNullOrEmpty(ListMacro(i).ListActions(j).IdDevice) = False Then
                                Dim devices As HoMIDom.HoMIDom.TemplateDevice = myService.ReturnDeviceByID(IdSrv, ListMacro(i).ListActions(j).IdDevice)
                                Fichier.WriteLine(";;;;ActionDevice;" & devices.Name & ";" & ListMacro(i).ListActions(j).Method _
                               & ";;" & ListMacro(i).ListActions(j).Parametres(0))
                            Else
                                Fichier.WriteLine(";;;;ActionDevice;" & ";;" & ListMacro(i).ListActions(j).Method _
                                & ";;" & ListMacro(i).ListActions(j).Parametres(0))
                            End If
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionDriver
                            Dim driver As HoMIDom.HoMIDom.TemplateDriver = myService.ReturnDriverByID(IdSrv, ListMacro(i).ListActions(j).IdDriver)
                            Fichier.WriteLine(";;;;ActionDriver;" & ListMacro(i).ListActions(j).Method & ";" & driver.Nom)
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionMail
                            Fichier.WriteLine(";;;;ActionMail;;;" & ListMacro(i).ListActions(j).Sujet & ";" & ListMacro(i).ListActions(j).Message)
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionSpeech
                            Fichier.WriteLine(";;;;ActionSpeech;;;;" & ListMacro(i).ListActions(j).Message)
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionHttp
                            Fichier.WriteLine(";;;;ActionHttp;" & ListMacro(i).ListActions(j).Commande)
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionIf
                            Compt += 1
                            For g = 0 To ListMacro(i).ListActions(j).Conditions.Count - 1
                                If String.IsNullOrEmpty(ListMacro(i).ListActions(j).Conditions(g).IdDevice) = False Then
                                    Dim Tex As String = ListMacro(i).ListActions(j).Conditions(g).Operateur.ToString
                                    If Tex = "NONE" Then
                                        Tex = ""
                                    End If
                                    Dim devices As HoMIDom.HoMIDom.TemplateDevice = myService.ReturnDeviceByID(IdSrv, ListMacro(i).ListActions(j).Conditions(g).IdDevice)
                                    Fichier.WriteLine(";;;;ActionIf;" & devices.Name & ";" & Tex & ";" &
                                                       ListMacro(i).ListActions(j).Conditions(g).Condition.ToString & ";" &
                                                       ListMacro(i).ListActions(j).Conditions(g).Value)
                                Else
                                    Fichier.WriteLine(";;;;ActionIf;;" & ListMacro(i).ListActions(j).Conditions(g).Condition.ToString & ";" &
                                                      ListMacro(i).ListActions(j).Conditions(g).Type.ToString & ";" &
                                                      ListMacro(i).ListActions(j).Conditions(g).DateTime)
                                End If
                            Next

                            'Si True
                            For f = 0 To ListMacro(i).ListActions(j).ListTrue.Count - 1
                                Dim _type As HoMIDom.HoMIDom.Action.TypeAction = ListMacro(i).ListActions(j).ListTrue(f).TypeAction
                                Select Case _type
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDevice
                                        If String.IsNullOrEmpty(ListMacro(i).ListActions(j).ListTrue(f).IdDevice) = False Then
                                            Dim devices As HoMIDom.HoMIDom.TemplateDevice = myService.ReturnDeviceByID(IdSrv, ListMacro(i).ListActions(j).ListTrue(f).IdDevice)
                                            Fichier.WriteLine(";;;Then;ActionDevice;" & devices.Name & ";" & ListMacro(i).ListActions(j).ListTrue(f).Method _
                                             & ";;" & ListMacro(i).ListActions(j).ListTrue(f).Parametres(0))
                                        Else
                                            Fichier.WriteLine(";;;Then;;ActionDevice;;;" & ListMacro(i).ListActions(j).ListTrue(f).Method _
                                            & ";;" & ListMacro(i).ListActions(j).ListTrue(f).Parametres(0))
                                        End If
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionIf
                                        RechercheList(Fichier, ListMacro(i).ListActions(j).ListTrue, f, Compt, "Then")
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDriver
                                        Dim driver As HoMIDom.HoMIDom.TemplateDriver = myService.ReturnDriverByID(IdSrv, ListMacro(i).ListActions(j).ListTrue(f).IdDriver)
                                        Fichier.WriteLine(";;;Then;ActionDriver;" & ListMacro(i).ListActions(j).ListTrue(f).Method & ";" & driver.Nom)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionMail
                                        Fichier.WriteLine(";;;Then;ActionMail;;;" & ListMacro(i).ListActions(j).ListTrue(f).Sujet & ";" & ListMacro(i).ListActions(j).ListTrue(f).Message)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionSpeech
                                        Fichier.WriteLine(";;;Then;ActionSpeech;;;;" & ListMacro(i).ListActions(j).ListTrue(f).Message)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionHttp
                                        Fichier.WriteLine(";;;Then;ActionHttp;" & ListMacro(i).ListActions(j).ListTrue(f).Commande)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionMacro
                                        Dim Macr As HoMIDom.HoMIDom.Macro = myService.ReturnMacroById(IdSrv, ListMacro(i).ListActions(j).ListTrue(f).IdMacro)
                                        Fichier.WriteLine(";;;Then;ActionMacro;" & Macr.Nom)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEvent
                                        Fichier.WriteLine(";;;Then;ActionLogEvent;" & ListMacro(i).ListActions(j).ListTrue(f).Type.ToString & ";;" &
                                                          ListMacro(i).ListActions(j).ListTrue(f).Eventid & ";" & ListMacro(i).ListActions(j).ListTrue(f).Message)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEventHomidom
                                        Fichier.WriteLine(";;;Then;ActionLogEventHomidom;" & ListMacro(i).ListActions(j).ListTrue(f).Type.ToString & ";;" &
                                                          ListMacro(i).ListActions(j).ListTrue(f).Fonction & ";" & ListMacro(i).ListActions(j).ListTrue(f).Message)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDOS
                                        Fichier.WriteLine(";;;Then;ActionDOS;" & ListMacro(i).ListActions(j).ListTrue(f).Fichier & ";" & ListMacro(i).ListActions(j).ListTrue(f).Arguments)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionVB
                                        Fichier.WriteLine(";;;Then;ActionVB;" & ListMacro(i).ListActions(j).ListTrue(f).Label & ";Executer")
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionStop
                                        Fichier.WriteLine(";;;Then;ActionStop;")
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionVar
                                        Fichier.WriteLine(";;;Then;ActionVar;" & ListMacro(i).ListActions(j).ListTrue(f).Nom & ";;;" & ListMacro(i).ListActions(j).ListTrue(f).Value)
                                End Select
                            Next

                            'Si False
                            For e = 0 To ListMacro(i).ListActions(j).ListFalse.Count - 1
                                Dim _type As HoMIDom.HoMIDom.Action.TypeAction = ListMacro(i).ListActions(j).ListFalse(e).TypeAction
                                Select Case _type
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDevice
                                        If String.IsNullOrEmpty(ListMacro(i).ListActions(j).ListFalse(e).IdDevice) = False Then
                                            Dim devices As HoMIDom.HoMIDom.TemplateDevice = myService.ReturnDeviceByID(IdSrv, ListMacro(i).ListActions(j).ListFalse(e).IdDevice)
                                            Fichier.WriteLine(";;;Else;ActionDevice;" & devices.Name & ";" & ListMacro(i).ListActions(j).ListFalse(e).Method _
                                             & ";;" & ListMacro(i).ListActions(j).ListFalse(e).Parametres(0))
                                        Else
                                            Fichier.WriteLine(";;;Else;ActionDevice;" & ";;" & ListMacro(i).ListActions(j).ListFalse(e).Method _
                                              & ";;" & ListMacro(i).ListActions(j).ListFalse(e).Parametres(0))
                                        End If
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionIf
                                        RechercheList(Fichier, ListMacro(i).ListActions(j).ListFalse, e, Compt, "Else")
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDriver
                                        Dim driver As HoMIDom.HoMIDom.TemplateDriver = myService.ReturnDriverByID(IdSrv, ListMacro(i).ListActions(j).ListFalse(e).IdDriver)
                                        Fichier.WriteLine(";;;Else;ActionDriver;" & ListMacro(i).ListActions(j).ListFalse(e).Method & ";" & driver.Nom)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionMail
                                        Fichier.WriteLine(";;;Else;ActionMail;;;" & ListMacro(i).ListActions(j).ListFalse(e).Sujet & ";" & ListMacro(i).ListActions(j).ListFalse(e).Message)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionSpeech
                                        Fichier.WriteLine(";;;Else;ActionSpeech;;;;" & ListMacro(i).ListActions(j).ListFalse(e).Message)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionHttp
                                        Fichier.WriteLine(";;;Else;ActionHttp;" & ListMacro(i).ListActions(j).ListFalse(e).Commande)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionMacro
                                        Dim Macr As HoMIDom.HoMIDom.Macro = myService.ReturnMacroById(IdSrv, ListMacro(i).ListActions(j).ListFalse(e).IdMacro)
                                        Fichier.WriteLine(";;;Else;ActionMacro;" & Macr.Nom)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEvent
                                        Fichier.WriteLine(";;;Else;ActionLogEvent;" & ListMacro(i).ListActions(j).ListFalse(e).Type.ToString & ";;" &
                                                          ListMacro(i).ListActions(j).ListFalse(e).Eventid & ";" & ListMacro(i).ListActions(j).ListFalse(e).Message)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEventHomidom
                                        Fichier.WriteLine(";;;Else;ActionLogEventHomidom;" & ListMacro(i).ListActions(j).ListFalse(e).Type.ToString & ";;" &
                                                          ListMacro(i).ListActions(j).ListFalse(e).Fonction & ";" & ListMacro(i).ListActions(j).ListFalse(e).Message)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDOS
                                        Fichier.WriteLine(";;;Else;ActionDOS;" & ListMacro(i).ListActions(j).ListFalse(e).Fichier & ";" & ListMacro(i).ListActions(j).ListFalse(e).Arguments)
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionVB
                                        Fichier.WriteLine(";;;Else;ActionVB;" & ListMacro(i).ListActions(j).ListFalse(e).Label & ";Executer")
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionStop
                                        Fichier.WriteLine(";;;Else;ActionStop;")
                                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionVar
                                        Fichier.WriteLine(";;;Else;ActionVar;" & ListMacro(i).ListActions(j).ListFalse(e).Nom & ";;;" & ListMacro(i).ListActions(j).ListFalse(e).Value)
                                End Select
                            Next
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionMacro
                            Dim Macr As HoMIDom.HoMIDom.Macro = myService.ReturnMacroById(IdSrv, ListMacro(i).ListActions(j).IdMacro)
                            Fichier.WriteLine(";;;;ActionMacro;" & Macr.Nom)
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEvent
                            Fichier.WriteLine(";;;;ActionLogEvent;" & ListMacro(i).ListActions(j).Type.ToString & ";;" &
                                              ListMacro(i).ListActions(j).Eventid & ";" & ListMacro(i).ListActions(j).Message)
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEventHomidom
                            Fichier.WriteLine(";;;;ActionLogEventHomidom;" & ListMacro(i).ListActions(j).Type.ToString & ";;" &
                                              ListMacro(i).ListActions(j).Fonction & ";" & ListMacro(i).ListActions(j).Message)
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionDOS
                            Fichier.WriteLine(";;;;ActionDOS;" & ListMacro(i).ListActions(j).Fichier & ";" & ListMacro(i).ListActions(j).Arguments)
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionVB
                            Fichier.WriteLine(";;;;ActionVB;" & ListMacro(i).ListActions(j).Label & ";Executer")
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionStop
                            Fichier.WriteLine(";;;;ActionStop;")
                        Case HoMIDom.HoMIDom.Action.TypeAction.ActionVar
                            Fichier.WriteLine(";;;;ActionVar;" & ListMacro(i).ListActions(j).Nom & ";;;" & ListMacro(i).ListActions(j).Value)
                    End Select
                Next
            Next
            Fichier.Flush()
            Fichier.Close()
            ListMacro.Clear()

            'newFile = Nothing
            'zipFile = Nothing
            'ListDevice = Nothing
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Sauve MacroList : ", "Réussit", "")

        Catch ex As IO.IOException
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur MacroList.csv : " & ex.ToString, "ERREUR", "")

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur SauverMacroList : " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub RechercheList(ByVal Fichier As StreamWriter, ByVal Pointeur As ArrayList, ByVal Index As Integer, ByRef Cpt As Integer, ByVal Test As String)
        Try
            Dim Espace1 As String = ""
            Dim Espace As String = ""
            Dim Ide As String = ""
            Dim Compt As Integer = Cpt + 1
            Espace = Espace.PadLeft(Compt * 5, " ")
            Espace1 = Espace1.PadLeft(Cpt * 5, " ")
            Dim Texte As String = "Then (" & Compt.ToString & ")"
            Dim Texte1 As String = "False (" & Compt.ToString & ")"

            If Cpt = 0 Then
                Ide = ""
            Else
                Ide = "(" & Cpt.ToString & ")"
            End If
            For f = 0 To Pointeur(Index).Conditions.Count - 1
                Dim _type As HoMIDom.HoMIDom.Action.TypeAction = Pointeur(Index).TypeAction
                Select Case _type
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionIf
                        If String.IsNullOrEmpty(Pointeur(Index).Conditions(f).IdDevice) = False Then
                            Dim Tex As String = Pointeur(Index).Conditions(f).Operateur.ToString
                            If Tex = "NONE" Then
                                Tex = ""
                            End If
                            Dim devices As HoMIDom.HoMIDom.TemplateDevice = myService.ReturnDeviceByID(IdSrv, Pointeur(Index).Conditions(f).IdDevice)
                            Fichier.WriteLine(";;;" & Espace1 & Test & Ide & ";" & Espace & "ActionIf (" & Compt.ToString & ");" & devices.Name & ";" & Tex & ";" &
                                                           Pointeur(Index).Conditions(f).Condition.ToString & ";" &
                                                            Pointeur(Index).Conditions(f).Value)
                        Else
                            Fichier.WriteLine(";;;" & Espace1 & Test & Ide & ";" & Espace & "ActionIf (" & Compt.ToString & ");;" & Pointeur(Index).Conditions(f).Condition.ToString & ";" &
                                                          Pointeur(Index).Conditions(f).Type.ToString & ";" &
                                                          Pointeur(Index).Conditions(f).DateTime)
                        End If
                End Select
            Next

            'Si True
            For f = 0 To Pointeur(Index).ListTrue.Count - 1
                Dim _type As HoMIDom.HoMIDom.Action.TypeAction = Pointeur(Index).ListTrue(f).TypeAction
                Select Case _type
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDevice
                        If String.IsNullOrEmpty(Pointeur(Index).ListTrue(f).IdDevice) = False Then
                            Dim devices As HoMIDom.HoMIDom.TemplateDevice = myService.ReturnDeviceByID(IdSrv, Pointeur(Index).ListTrue(f).IdDevice)
                            Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionDevice;" & devices.Name & ";" & Pointeur(Index).ListTrue(f).Method _
                             & ";;" & Pointeur(Index).ListTrue(f).Parametres(0))
                        Else
                            Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionDevice;;;" & Pointeur(Index).ListTrue(f).Method _
                            & ";;" & Pointeur(Index).ListTrue(f).Parametres(0))
                        End If
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionIf
                        RechercheList(Fichier, Pointeur(Index).ListTrue, Index, Compt, "Then")
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDriver
                        Dim driver As HoMIDom.HoMIDom.TemplateDriver = myService.ReturnDriverByID(IdSrv, Pointeur(0).ListTrue(f).IdDriver)
                        Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionDriver;" & Pointeur(Index).ListTrue(f).Method & ";" & driver.Nom)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionMail
                        Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionMail;;;" & Pointeur(Index).ListTrue(f).Sujet & ";" & Pointeur(0).ListTrue(f).Message)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionSpeech
                        Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionSpeech;;;;" & Pointeur(Index).ListTrue(f).Message)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionHttp
                        Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionHttp;" & Pointeur(Index).ListTrue(f).Commande)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionMacro
                        Dim Macr As HoMIDom.HoMIDom.Macro = myService.ReturnMacroById(IdSrv, Pointeur(Index).ListTrue(f).IdMacro)
                        Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionMacro;" & Macr.Nom)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEvent
                        Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionLogEvent;" & Pointeur(Index).ListTrue(f).Type.ToString & ";;" &
                                          Pointeur(0).ListTrue(f).Eventid & ";" & Pointeur(Index).ListTrue(f).Message)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEventHomidom
                        Fichier.WriteLine(";;;;" & Espace & Texte & ";" & Espace & "ActionLogEventHomidom;" & Pointeur(Index).ListTrue(f).Type.ToString & ";;" &
                                          Pointeur(0).ListTrue(f).Fonction & ";" & Pointeur(Index).ListTrue(f).Message)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDOS
                        Fichier.WriteLine(";;;;" & Espace & Texte & ";" & Espace & "ActionDOS;" & Pointeur(Index).ListTrue(f).Fichier & ";" & Pointeur(Index).ListTrue(f).Arguments)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionVB
                        Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionVB;" & Pointeur(Index).ListTrue(f).Label & ";Executer")
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionStop
                        Fichier.WriteLine(";;;;" & Espace & Texte & ";" & Espace & "ActionStop;")
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionVar
                        Fichier.WriteLine(";;;" & Espace & Texte & ";" & Espace & "ActionVar;" & Pointeur(Index).ListTrue(f).Nom & ";;;" & Pointeur(Index).ListTrue(f).Value)
                End Select
            Next

            'Si False
            For e = 0 To Pointeur(Index).ListFalse.Count - 1
                Dim _type As HoMIDom.HoMIDom.Action.TypeAction = Pointeur(Index).ListFalse(e).TypeAction
                Select Case _type
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDevice
                        If String.IsNullOrEmpty(Pointeur(Index).ListFalse(e).IdDevice) = False Then
                            Dim devices As HoMIDom.HoMIDom.TemplateDevice = myService.ReturnDeviceByID(IdSrv, Pointeur(Index).ListFalse(e).IdDevice)
                            Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionDevice;" & devices.Name & ";" & Pointeur(Index).ListFalse(e).Method _
                            & ";;" & Pointeur(Index).ListFalse(e).Parametres(0))
                        Else
                            Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionDevice;" & ";;" & Pointeur(Index).ListFalse(e).Method _
                             & ";;" & Pointeur(Index).ListFalse(e).Parametres(0))
                        End If
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionIf
                        RechercheList(Fichier, Pointeur(Index).ListFalse, Index, Compt, "Else")
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDriver
                        Dim driver As HoMIDom.HoMIDom.TemplateDriver = myService.ReturnDriverByID(IdSrv, Pointeur(Index).ListFalse(e).IdDriver)
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionDriver;" & Pointeur(Index).ListFalse(e).Method & ";" & driver.Nom)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionMail
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionMail;;;" & Pointeur(Index).ListFalse(e).Sujet & ";" & Pointeur(Index).ListFalse(e).Message)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionSpeech
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionSpeech;;;;" & Pointeur(Index).ListFalse(e).Message)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionHttp
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionHttp;" & Pointeur(Index).ListFalse(e).Commande)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionMacro
                        Dim Macr As HoMIDom.HoMIDom.Macro = myService.ReturnMacroById(IdSrv, Pointeur(Index).ListFalse(e).IdMacro)
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionMacro;" & Macr.Nom)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEvent
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionLogEvent;" & Pointeur(Index).ListFalse(e).Type.ToString & ";;" &
                                          Pointeur(0).ListFalse(e).Eventid & ";" & Pointeur(Index).ListFalse(e).Message)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionLogEventHomidom
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionLogEventHomidom;" & Pointeur(Index).ListFalse(e).Type.ToString & ";;" &
                                          Pointeur(0).ListFalse(e).Fonction & ";" & Pointeur(Index).ListFalse(e).Message)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionDOS
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionDOS;" & Pointeur(Index).ListFalse(e).Fichier & ";" & Pointeur(Index).ListFalse(e).Arguments)
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionVB
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionVB;" & Pointeur(Index).ListFalse(e).Label & ";Executer")
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionStop
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionStop;")
                    Case HoMIDom.HoMIDom.Action.TypeAction.ActionVar
                        Fichier.WriteLine(";;;" & Espace & Texte1 & ";" & Espace & "ActionVar;" & Pointeur(Index).ListFalse(e).Nom & ";;;" & Pointeur(Index).ListFalse(e).Value)
                End Select
            Next

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur RechercheList : " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub SauverDeviceList(ByVal NomFichier As String)
        'Sauver la liste des New Devices dans un fichier "DevicesList.csv"  Ajouter par JPS
        Try
            Dim Listdevices As New List(Of HoMIDom.HoMIDom.TemplateDevice)
            Dim zipFile As FileInfo
            Dim DriverNam As String = ""
            Dim IdDriver As String = ""

            Listdevices.Clear()
            Listdevices = myService.GetAllDevices(IdSrv)
            zipFile = New FileInfo(NomFichier)
            Dim Fichier As StreamWriter = zipFile.CreateText()
            Fichier.WriteLine("Adresse1" & ";" & "Adresse2" & ";" & "Nom" & ";" & "Type" & ";" & "Enable" & ";" & "ID" & ";" &
                              "Date/Heure" & ";" & "Driver")
            For i = 0 To Listdevices.Count - 1
                If String.IsNullOrEmpty(Listdevices(i).DriverID) = False Then
                    If Listdevices(i).DriverID <> IdDriver Then
                        Dim x As HoMIDom.HoMIDom.TemplateDriver = myService.ReturnDriverByID(IdSrv, Listdevices(i).DriverID)
                        If x IsNot Nothing Then DriverNam = x.Nom
                        IdDriver = Listdevices(i).DriverID
                    End If
                    Fichier.WriteLine(Listdevices(i).Adresse1 & ";" &
                                  Listdevices(i).Adresse2 & ";" &
                                  Listdevices(i).Name & ";" &
                                  Listdevices(i).Type.ToString & ";" &
                                  Listdevices(i).Enable & ";" &
                                  Listdevices(i).ID & ";" &
                                  Listdevices(i).DateCreated & ";" &
                                  DriverNam)
                End If
            Next
            Fichier.Flush()
            Fichier.Close()
            Listdevices.Clear()
            'newFile = Nothing
            'zipFile = Nothing
            'ListNewDevice = Nothing
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "Sauve DeviceList : ", "Réussit", "")
        Catch ex As IO.IOException
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur DeviceList.csv : " & ex.ToString, "ERREUR", "")

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur SauverDeviceList : " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub SauverCopiefichier(ByVal FileTitle As String)    'Ajouter par JPS
        Try
            Dim zipFile As FileInfo
            Dim newFile As FileInfo

            'Sauver une copie du fichier en "*.bak"            
            Dim ParaFichier = Split(FileTitle, ".")
            ParaFichier(1) = "bak"
            zipFile = New FileInfo(ParaFichier(0) + "." + ParaFichier(1))
            If (zipFile.Exists) Then
                zipFile.Delete()
            End If
            newFile = New FileInfo(FileTitle)
            newFile.CopyTo(zipFile.FullName)
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur SauverCopiefichier: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

End Class
