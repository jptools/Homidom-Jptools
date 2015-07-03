﻿Imports HoMIDom.HoMIDom.Telecommande
Imports HoMIDom.HoMIDom.Api
Imports System.IO

Public Class WTelecommandeNew

#Region "Variables"
    Dim FlagNewCmd As Boolean
    Dim _DevId As String = ""
    'Dim x As HoMIDom.HoMIDom.TemplateDevice = Nothing
    'Dim _SelectDriverIndex As Integer 'Index du driver sélectionné
    'Dim ListButton As New List(Of ImageButton)
    Dim _Row As Integer
    Dim _Col As Integer
    Dim _MyTemplate As HoMIDom.HoMIDom.Telecommande.Template = Nothing
    Dim _CurrentTemplate As HoMIDom.HoMIDom.Telecommande.Template = Nothing
#End Region

#Region "Gestion des commandes"
    ''' <summary>
    ''' Nouvelle commande
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnNewCmd_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnNewCmd.Click
        TxtCmdName.Text = ""
        TxtCmdRepeat.Text = "0"
        TxtCmdData.Text = ""
        CbFormat.SelectedIndex = 0
        'ImgCommande2.Source = Nothing
        ' ImgCommande2.Tag = ""

        BtnNewCmd.Visibility = Windows.Visibility.Hidden
        BtnSaveCmd.Visibility = Windows.Visibility.Visible
        FlagNewCmd = True
    End Sub

    ''' <summary>
    ''' Sauvegarder commande
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnSaveCmd_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnSaveCmd.Click
        Try
            If IsNumeric(TxtCmdRepeat.Text) = False Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Numérique obligatoire pour repeat !!", "Erreur", "")
                Exit Sub
            End If
            If String.IsNullOrEmpty(TxtCmdName.Text) Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le nom de la commande est obligatoire !!", "Erreur", "")
                Exit Sub
            End If

            If _CurrentTemplate IsNot Nothing Then

                If FlagNewCmd = True Then 'nouvelle commande
                    Dim _cmd As New HoMIDom.HoMIDom.Telecommande.Commandes
                    With _cmd
                        .Name = TxtCmdName.Text
                        .Code = TxtCmdData.Text
                        .Repeat = TxtCmdRepeat.Text
                        .Format = CbFormat.SelectedIndex
                        '.Picture = ImgCommande2.Tag
                    End With
                    _CurrentTemplate.Commandes.Add(_cmd)
                Else 'modifier commande
                    Dim idx As Integer = ListCmd.SelectedIndex
                    If idx < 0 Then Exit Sub

                    With _CurrentTemplate.Commandes.Item(idx)
                        .Name = TxtCmdName.Text
                        .Code = TxtCmdData.Text
                        .Repeat = TxtCmdRepeat.Text
                        .Format = CbFormat.SelectedIndex
                        '.Picture = ImgCommande2.Tag
                    End With
                End If
                TxtCmdName.Text = ""
                TxtCmdData.Text = ""
                TxtCmdRepeat.Text = ""
                CbFormat.SelectedIndex = 0

                ListCmd.Items.Clear()
                For i2 As Integer = 0 To _CurrentTemplate.Commandes.Count - 1
                    ListCmd.Items.Add(_CurrentTemplate.Commandes.Item(i2).Name)
                Next
            End If

            BtnNewCmd.Visibility = Windows.Visibility.Visible
            FlagNewCmd = False
        Catch Ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur BtnSaveCmd: " & Ex.Message, "Erreur", "")
        End Try
    End Sub

    ''' <summary>
    ''' Supprimer une commande
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnDelCmd_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnDelCmd.Click
        Try
            If ListCmd.SelectedIndex >= 0 Then
                _CurrentTemplate.Commandes.RemoveAt(ListCmd.SelectedIndex)
                Dim retour As String = myService.SaveTemplate(IdSrv, _CurrentTemplate)
                If retour <> "0" Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'enregistrement de la commande dans le template: " & retour, "Erreur", "")
                    Exit Sub
                Else
                    ListCmd.Items.Clear()
                    For i2 As Integer = 0 To _CurrentTemplate.Commandes.Count - 1
                        ListCmd.Items.Add(_CurrentTemplate.Commandes.Item(i2).Name)
                    Next

                End If

                BtnDelCmd.Visibility = Windows.Visibility.Hidden
                TxtCmdData.Text = ""
                TxtCmdName.Text = ""
                TxtCmdRepeat.Text = ""
                CbFormat.SelectedIndex = 0
                'ImgCommande2.Source = Nothing
                'ImgCommande2.Tag = Nothing
            End If
        Catch Ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur BtnDelCmd: " & Ex.Message, "Erreur", "")
        End Try
    End Sub

    ''' <summary>
    ''' Tester une commande
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnTstCmd_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnTstCmd.Click
        Try
            If String.IsNullOrEmpty(TxtCmdName.Text) = False And _CurrentTemplate IsNot Nothing Then
                If String.IsNullOrEmpty(_DevId) = False Then
                    Dim x As New HoMIDom.HoMIDom.DeviceAction
                    x.Nom = "EnvoyerCommande"
                    Dim param As New HoMIDom.HoMIDom.DeviceAction.Parametre
                    param.Value = TxtCmdName.Text
                    x.Parametres.Add(param)
                    myService.ExecuteDeviceCommand(IdSrv, _DevId, x)
                End If

                '_CurrentTemplate.ExecuteCommand(IdSrv, TxtCmdName.Text, myService)
            End If
            'Dim retour As String = myService.TelecommandeSendCommand(IdSrv, _DeviceId, TxtCmdName.Text)
            'If retour <> 0 Then
            '    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, retour, "Erreur", "")
            'End If
        Catch Ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur BtnTstCmd: " & Ex.Message, "Erreur", "")
        End Try
    End Sub

    ''' <summary>
    ''' Apprendre une commande IR
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnLearn_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnLearn.Click
        Try
            Dim _idDrv As String = "0"

            'On envoi la commande
            Select Case _CurrentTemplate.Type
                Case 0 'Template de type http
                    _IdDrv = "D04010DA-5E22-11E1-A742-4E4A4824019B"
                Case 1 'Template de type IR
                    _idDrv = "74FD4E7C-34ED-11E0-8AC4-70CEDED72085"
                Case 2 'Template de type rs232
                    _idDrv = "7631FA52-31C2-11E3-8BE1-6DD36088709B"
            End Select


            TxtCmdData.Text = myService.StartLearning(IdSrv, _idDrv)
        Catch Ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur BtnLearnCmd: " & Ex.Message, "Erreur", "")
        End Try
    End Sub

    'Private Sub BtnImg_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnImg.Click
    '    Try
    '        Dim frm As New WindowImg
    '        frm.ShowDialog()
    '        If frm.DialogResult.HasValue And frm.DialogResult.Value Then
    '            Dim retour As String = frm.FileName
    '            If String.IsNullOrEmpty(retour) = False Then
    '                ImgCommande2.Source = ConvertArrayToImage(myService.GetByteFromImage(retour))
    '                ImgCommande2.Tag = retour
    '                BtnSaveCmd.Visibility = Windows.Visibility.Visible
    '            End If
    '            frm.Close()
    '        Else
    '            frm.Close()
    '        End If
    '        frm = Nothing
    '    Catch ex As Exception
    '        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub ImgCommande_MouseLeftButtonDown: " & ex.Message, "ERREUR", "")
    '    End Try
    'End Sub

    ''' <summary>
    ''' Changement de sélection d'une commande dans la liste
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub ListCmd_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles ListCmd.SelectionChanged
        Try
            Dim i As Integer = ListCmd.SelectedIndex
            If i < 0 Then Exit Sub

            If _CurrentTemplate IsNot Nothing Then

                BtnSaveCmd.Visibility = Windows.Visibility.Collapsed
                BtnDelCmd.Visibility = Windows.Visibility.Visible
                If String.IsNullOrEmpty(_DevId) = False Then BtnTstCmd.Visibility = Windows.Visibility.Visible
                TxtCmdName.Text = _CurrentTemplate.Commandes.Item(i).Name
                TxtCmdData.Text = _CurrentTemplate.Commandes.Item(i).Code
                TxtCmdRepeat.Text = _CurrentTemplate.Commandes.Item(i).Repeat
                CbFormat.SelectedIndex = _CurrentTemplate.Commandes.Item(i).Format
                'If String.IsNullOrEmpty(_CurrentTemplate.Commandes.Item(i).Picture) = False Then
                '    ImgCommande2.Source = ConvertArrayToImage(myService.GetByteFromImage(_CurrentTemplate.Commandes.Item(i).Picture))
                '    ImgCommande2.Tag = _CurrentTemplate.Commandes.Item(i).Picture
                '    ImgCommande2.Command = TxtCmdName.Text
                'End If

            End If
        Catch Ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur ListCmd_SelectionChanged: " & Ex.Message, "Erreur", "")
        End Try
    End Sub

    Private Sub TxtCmdData_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.TextChangedEventArgs) Handles TxtCmdData.TextChanged
        BtnSaveCmd.Visibility = Windows.Visibility.Visible
    End Sub

    ''' <summary>
    ''' Valeur de data de la commande a changée
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub TxtCmdData_TextInput(ByVal sender As Object, ByVal e As System.Windows.Input.TextCompositionEventArgs) Handles TxtCmdData.TextInput
        If String.IsNullOrEmpty(_DevId) = False Then BtnTstCmd.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub TxtCmdName_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles TxtCmdName.MouseDown
        BtnSaveCmd.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub TxtCmdRepeat_TextChanged(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles TxtCmdRepeat.MouseDown
        BtnSaveCmd.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub CbFormat_MouseUp(sender As Object, e As System.Windows.Input.MouseButtonEventArgs) Handles CbFormat.MouseUp
        BtnSaveCmd.Visibility = Windows.Visibility.Visible
    End Sub

    Private Sub CbFormat_SelectionChanged(sender As System.Object, e As System.Windows.Controls.SelectionChangedEventArgs) Handles CbFormat.SelectionChanged
        If CbFormat.Visibility = Windows.Visibility.Visible Then BtnSaveCmd.Visibility = Windows.Visibility.Visible
    End Sub

    '''' <summary>
    '''' Clic sur image de la commande
    '''' </summary>
    '''' <param name="sender"></param>
    '''' <param name="e"></param>
    '''' <remarks></remarks>
    'Private Sub ImgCommande_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgCommande2.MouseLeftButtonDown ', Canvas2.MouseLeftButtonDown
    '    For i As Integer = 0 To grid_Telecommande.Children.Count - 1
    '        Dim cvs As Canvas = grid_Telecommande.Children.Item(i)

    '        If cvs IsNot Nothing Then
    '            If cvs.Children.Count <> 0 Then
    '                Dim y As ImageButton = cvs.Children.Item(0)
    '                If ListCmd.SelectedValue = y.Command Then
    '                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Vous ne pouvez pas inclure cette commande dans la grille car elle y est déjà utilisée!", "Information", "")
    '                    Exit Sub
    '                End If

    '            End If
    '        End If
    '    Next

    '    If ImgCommande2.Source IsNot Nothing Then
    '        Dim effects As DragDropEffects
    '        Dim obj As New DataObject()
    '        obj.SetData(GetType(ImageButton), sender)
    '        effects = DragDrop.DoDragDrop(sender, obj, DragDropEffects.Copy Or DragDropEffects.Move)
    '    Else
    '        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Vous ne pouvez pas inclure cette commande dans la grille car elle ne comporte aucune image!", "Information", "")
    '    End If
    'End Sub

#End Region


#Region "Gestion des Templates"
    ''' <summary>
    ''' Nouveau template
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnNewTemplate_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnNewTemplate.Click
        StkNewTemplate.Visibility = Windows.Visibility.Visible
        TxtTplName.Focus()
        BtnSaveTemplate.Visibility = Windows.Visibility.Visible
        StkTplBase.Visibility = Windows.Visibility.Visible
        TxtTplFab.IsReadOnly = False
        TxtTplMod.IsReadOnly = False
        TxtTplName.IsReadOnly = False
        TxtCharEndReceive.IsReadOnly = False
        TxtTrameInit.IsReadOnly = False

        TxtTplFab.Text = ""
        TxtTplName.Text = ""
        TxtTplMod.Text = ""
        TxtCharEndReceive.Text = ""
        TxtTrameInit.Text = ""
        cbTemplate.Text = ""
        ChkMultimedia.IsChecked = True

        RdHttp.IsEnabled = True
        RdIR.IsEnabled = True
        RdRS232.IsEnabled = True
    End Sub

    ''' <summary>
    ''' Sauvegarder Template
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub BtnSaveTemplate_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnSaveTemplate.Click
        Try

            If String.IsNullOrEmpty(TxtTplFab.Text) Or HaveCaractSpecial(TxtTplFab.Text) Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le nom du fabricant est obligatoire et ne doit pas comporter de caractère spécial (%,-,!...) ", "Erreur", "")
                TxtTplFab.Focus()
                Exit Sub
            End If
            If String.IsNullOrEmpty(TxtTplMod.Text) Or HaveCaractSpecial(TxtTplMod.Text) Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le nom du modèle est obligatoire et ne doit pas comporter de caractère spécial (%,-,!...)", "Erreur", "")
                TxtTplMod.Focus()
                Exit Sub
            End If
            If String.IsNullOrEmpty(TxtTplName.Text) Or HaveCaractSpecial(TxtTplName.Text) Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le nom du template est obligatoire et ne doit pas comporter de caractère spécial (%,-,!...)", "Erreur", "")
                TxtTplName.Focus()
                Exit Sub
            End If

            _CurrentTemplate = New HoMIDom.HoMIDom.Telecommande.Template

            With _CurrentTemplate
                .Name = TxtTplName.Text
                .Modele = TxtTplMod.Text
                .Fabricant = TxtTplFab.Text
                .TrameInit = TxtTrameInit.Text
                .CharEndReceive = TxtCharEndReceive.Text
                .IsAudioVideo = ChkMultimedia.IsChecked
            End With

            'If _Driver IsNot Nothing Then
            '    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "OK", "", "")
            'Else
            '    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "VIDE", "", "")
            'End If

            Dim _type As String = ""
            If RdHttp.IsChecked Then
                _CurrentTemplate.Type = 0
            ElseIf RdIR.IsChecked Then
                _CurrentTemplate.Type = 1
            ElseIf RdRS232.IsChecked Then
                _CurrentTemplate.Type = 2
            End If

            Dim retour As String = myService.CreateNewTemplate(_CurrentTemplate)
            If retour <> "0" Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la création du nouveau template: " & retour, "Erreur", "")
                Exit Sub
            Else

                StkNewTemplate.Visibility = Windows.Visibility.Collapsed
                BtnSaveTemplate.Visibility = Windows.Visibility.Hidden
                StkTplBase.Visibility = Windows.Visibility.Collapsed

                cbTemplate.Items.Clear()
                Dim _list As New List(Of HoMIDom.HoMIDom.Telecommande.Template)
                _list = myService.GetListOfTemplate
                For i As Integer = 0 To _list.Count - 1
                    cbTemplate.Items.Add(_list(i).Name)
                Next
                cbTemplate.SelectedValue = _CurrentTemplate.Name


                TxtTplFab.Text = ""
                TxtTplName.Text = ""
                TxtTplMod.Text = ""
                TxtCharEndReceive.Text = ""
                TxtTrameInit.Text = ""
                cbTemplate.Text = ""

                TxtTplFab.IsReadOnly = True
                TxtTplMod.IsReadOnly = True
                TxtTplName.IsReadOnly = True
                TxtCharEndReceive.IsReadOnly = True
                TxtTrameInit.IsReadOnly = True
                RdHttp.IsEnabled = False
                RdIR.IsEnabled = False
                RdRS232.IsEnabled = False
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & ex.ToString)
        End Try
    End Sub


#End Region


    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="TemplateName">Nom du template à afficher</param>
    ''' <remarks></remarks>
    Public Sub New(Optional TemplateName As String = "", Optional DeviceId As String = "")

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        Try
            _DevId = DeviceId
            Call Refresh(TemplateName)

        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'ouverture de la fenêtre d'édition:" & ex.ToString, "Erreur", "")
        End Try
    End Sub

#Region "Designer"

    Private Sub DeleteButton(ByVal sender As Object)
        Try
            sender.parent.children.clear()

            For i As Integer = 0 To _CurrentTemplate.Commandes.Count - 1
                If _CurrentTemplate.Commandes(i).Name = sender.command Then

                End If
            Next
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Telecommande DeleteButton: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub Img_MouseLeftButtonDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs)
        Try
            Dim effects As DragDropEffects
            Dim obj As New DataObject()
            obj.SetData(GetType(ImageButton), sender)
            effects = DragDrop.DoDragDrop(sender, obj, DragDropEffects.Copy Or DragDropEffects.Move)
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Telecommande Img_MouseLeftButtonDown: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    'Fermer
    Private Sub button_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles button.Click
        Try
            DialogResult = True
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub Telecommande button_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

#End Region


    Private Sub Refresh(Optional TemplateName As String = "")
        Try
            'Récupère la liste des template
            Dim _list As New List(Of HoMIDom.HoMIDom.Telecommande.Template)
            _list = myService.GetListOfTemplate

            Dim idx As Integer = -1
            cbTemplate.Items.Clear()
            For i As Integer = 0 To _list.Count - 1
                cbTemplate.Items.Add(_list(i).Name) 'ajoute le nom du template dans la liste
                If String.IsNullOrEmpty(TemplateName) = False Then
                    If TemplateName = _list(i).Name Then idx = i
                End If
            Next
            cbTemplate.SelectedIndex = idx

            BtnSaveCmd.Visibility = Windows.Visibility.Collapsed
            BtnDelCmd.Visibility = Windows.Visibility.Collapsed
            BtnDeleteTemplate.Visibility = Windows.Visibility.Collapsed

            TxtTplFab.Text = ""
            TxtTplName.Text = ""
            TxtTplMod.Text = ""
            TxtCharEndReceive.Text = ""
            TxtTrameInit.Text = ""
            cbTemplate.Text = ""
        Catch ex As Exception
            MessageBox.Show(ex.ToString)
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'ouverture de la fenêtre d'édition:" & ex.ToString, "Erreur", "")
        End Try
    End Sub

    ''' <summary>
    ''' Sélection d'un template
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub cbTemplate_SelectionChanged(ByVal sender As System.Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles cbTemplate.SelectionChanged
        Try
            If cbTemplate.SelectedIndex < 0 Then
            Else
                'Récupère la liste des templates
                Dim _list As New List(Of HoMIDom.HoMIDom.Telecommande.Template)
                _list = myService.GetListOfTemplate

                _CurrentTemplate = _list.Item(cbTemplate.SelectedIndex)
                ChargeCmd()
                ChargeVar()

                TxtTplName.Text = _list.Item(cbTemplate.SelectedIndex).Name
                TxtTplFab.Text = _list.Item(cbTemplate.SelectedIndex).Fabricant
                TxtTplMod.Text = _list.Item(cbTemplate.SelectedIndex).Modele
                TxtTrameInit.Text = _list.Item(cbTemplate.SelectedIndex).TrameInit
                TxtCharEndReceive.Text = _list.Item(cbTemplate.SelectedIndex).CharEndReceive
                ChkMultimedia.IsChecked = _list.Item(cbTemplate.SelectedIndex).IsAudioVideo

                Select Case _list.Item(cbTemplate.SelectedIndex).Type
                    Case 0 'http
                        RdHttp.IsChecked = True
                        BtnLearn.Visibility = Windows.Visibility.Collapsed
                        Label27.Visibility = Windows.Visibility.Collapsed
                        CbFormat.Visibility = Windows.Visibility.Collapsed
                    Case 1 'IR
                        RdIR.IsChecked = True
                        BtnLearn.Visibility = Windows.Visibility.Visible
                        Label27.Visibility = Windows.Visibility.Visible
                        CbFormat.Visibility = Windows.Visibility.Visible
                    Case 2 'RS232
                        RdRS232.IsChecked = True
                        BtnLearn.Visibility = Windows.Visibility.Visible
                        Label27.Visibility = Windows.Visibility.Collapsed
                        CbFormat.Visibility = Windows.Visibility.Collapsed
                End Select

                RdHttp.IsEnabled = False
                RdIR.IsEnabled = False
                RdRS232.IsEnabled = False

                StkCmd.Visibility = Windows.Visibility.Visible
                StkVar.Visibility = Windows.Visibility.Visible
                BtnDeleteTemplate.Visibility = Windows.Visibility.Visible
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & ex.ToString)
        End Try
    End Sub

    Private Sub ChargeCmd()
        Try
            ListCmd.Items.Clear()
            If _CurrentTemplate IsNot Nothing Then
                For Each cmd In _CurrentTemplate.Commandes
                    ListCmd.Items.Add(cmd.Name)
                Next
            End If

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur ChargeCmd: " & ex.Message)
        End Try
    End Sub

    Private Sub ChargeVar()
        Try
            ListVar.Items.Clear()

            If _CurrentTemplate IsNot Nothing Then
                For Each var In _CurrentTemplate.Variables
                    ListVar.Items.Add(var.Name)
                Next
            End If

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur ChargeVar: " & ex.Message)
        End Try
    End Sub


#Region "Variable"
    Dim FlagNewVar As Boolean

    Private Sub BtnNewVar_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnNewVar.Click
        StkParamVar.Visibility = Windows.Visibility.Visible
        TxtVarName.IsEnabled = True
        TxtVarName.Text = ""
        TxtVarVal.Text = ""
        CbVarType.SelectedIndex = 0
        FlagNewVar = True
    End Sub

    Private Sub ListVar_SelectionChanged(ByVal sender As System.Object, ByVal e As Object) Handles ListVar.SelectionChanged, ListVar.MouseLeftButtonDown
        Try
            FlagNewVar = False
            If _CurrentTemplate IsNot Nothing Then
                For Each _var In _CurrentTemplate.Variables
                    If _var.Name = ListVar.SelectedValue Then
                        TxtVarName.Text = _var.Name
                        TxtVarName.IsEnabled = False
                        TxtVarVal.Text = _var.Value
                        CbVarType.SelectedIndex = _var.Type
                        StkParamVar.Visibility = Windows.Visibility.Visible
                        Exit For
                    End If
                Next
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur ListVar_SelectionChanged: " & ex.Message, "Erreur", "")
        End Try
    End Sub

    Private Sub BtnAnnulVar_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnAnnulVar.Click
        Try
            FlagNewVar = False
            StkParamVar.Visibility = Windows.Visibility.Collapsed
            TxtVarName.Text = ""
            TxtVarVal.Text = ""
            CbVarType.SelectedIndex = 0
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur BtnAnnulVar_Click: " & ex.Message, "Erreur", "")
        End Try
    End Sub

    Private Sub BtnDeleteVar_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnDeleteVar.Click
        Try
            FlagNewVar = False

            If ListVar.SelectedIndex < 0 Then
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Veuillez sélectionner une variable à supprimer!", "Erreur", "")
                Exit Sub
            End If

            If _CurrentTemplate IsNot Nothing Then
                Dim idx As Integer = 0
                For Each var In _CurrentTemplate.Variables
                    If var.Name = ListVar.SelectedValue Then
                        _CurrentTemplate.Variables.RemoveAt(idx)
                        Exit For
                    End If
                    idx += 1
                Next

                ChargeVar()
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur BtnDeleteVar_Click: " & ex.Message, "Erreur", "")
        End Try
    End Sub

    Private Sub BtnModifVar_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnModifVar.Click
        Try
            If FlagNewVar Then
                TxtVarName.IsEnabled = True
            Else
                TxtVarName.IsEnabled = False

                If ListVar.SelectedIndex < 0 Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Veuillez sélectionner une variable à modifier!", "Erreur", "")
                    Exit Sub
                End If
            End If

            If _CurrentTemplate IsNot Nothing Then
                If FlagNewVar Then
                    For Each var In _CurrentTemplate.Variables
                        If var.Name.ToUpper = TxtVarName.Text.ToUpper Then
                            MessageBox.Show("Vous ne pouvez pas utiliser ce nom de variable car il existe déjà !", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)
                            Exit For
                        End If
                    Next

                    Dim v As New TemplateVar
                    With v
                        .Name = TxtVarName.Text
                        .Type = CbVarType.SelectedIndex
                        .Value = TxtVarVal.Text
                    End With
                    _CurrentTemplate.Variables.Add(v)

                Else
                    Dim idx As Integer = 0
                    For Each var In _CurrentTemplate.Variables
                        If var.Name = TxtVarName.Text Then
                            var.Type = CbVarType.SelectedIndex
                            var.Value = TxtVarVal.Text
                            Exit For
                        End If
                        idx += 1
                    Next
                End If


                ChargeVar()
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur BtnModifVar_Click: " & ex.Message, "Erreur", "")
        End Try
    End Sub
#End Region

    Private Sub buttonOk_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles buttonOk.Click
        Try
            If _CurrentTemplate IsNot Nothing Then

                Dim retour As String = myService.SaveTemplate(IdSrv, _CurrentTemplate)

                If retour <> "0" Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de l'enregistrement de la commande dans le template: " & retour, "Erreur", "")
                End If
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur buttonOk_Click: " & ex.Message, "Erreur", "buttonOk_Click")
        End Try
    End Sub

    Private Sub BtnDeleteTemplate_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnDeleteTemplate.Click
        Try
            If cbTemplate.SelectedIndex < 0 Then
                MessageBox.Show("Veuillez sélectionner un template à supprimer!", "Erreur", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            Else
                'Récupère la liste des templates
                Dim _list As New List(Of HoMIDom.HoMIDom.Telecommande.Template)
                _list = myService.GetListOfTemplate

                'Recupère le template courant
                _CurrentTemplate = _list.Item(cbTemplate.SelectedIndex)

                'Supprime le template
                Dim retour As String = myService.DeleteTemplate(IdSrv, _CurrentTemplate)

                If retour <> "0" Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & retour, "Erreur", "DeleteTemplate")
                Else
                    Call Refresh()
                End If
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & ex.ToString)
        End Try
    End Sub

    Private Sub BtnImportTemplate_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnImportTemplate.Click
        Try
            MessageBox.Show("Fonction non disponible pour le moment!", "Export", MessageBoxButton.OK, MessageBoxImage.Information)
            Exit Sub

            If IsConnect = False Then
                Exit Sub
            End If

        '    Me.Cursor = Cursors.Wait
        '    'Exporter le fichier de config

        '    ' Configure open file dialog box
        '    Dim dlg As New Microsoft.Win32.OpenFileDialog
        '    dlg.FileName = "" ' Default file name
        '    dlg.DefaultExt = ".xml" ' Default file extension
        '    dlg.Filter = "Fichier de template (.xml)|*.xml" ' Filter files by extension

        '    ' Show open file dialog box
        '    Dim result As Boolean = dlg.ShowDialog()

        '    ' Process open file dialog box results
        '    If result = True Then
        '        ' Open document
        '        Dim filename As String = dlg.FileName
        '        Dim retour As String = myService.ExportConfig(IdSrv)
        '        If retour.StartsWith("ERREUR") Then
        '            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, retour, "Erreur export template", "")
        '        Else
        '            Dim TargetFile As StreamWriter
        '            TargetFile = New StreamWriter(filename, False)
        '            TargetFile.Write(retour)
        '            TargetFile.Close()
        '            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "L'export du fichier de template a été effectué", "Export Template", "")
        '        End If
        '    End If

        '    Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub BtnExportTemplate_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub BtnExportTemplate_Click(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles BtnExportTemplate.Click
        Try
            If IsConnect = False Then
                Exit Sub
            End If

            If cbTemplate.SelectedIndex < 0 Then
                MessageBox.Show("Veuillez sélectionner un template à exporter", "Erreur", MessageBoxButton.OK, MessageBoxImage.Exclamation)
                Exit Sub
            End If

            Me.Cursor = Cursors.Wait
            'Exporter le fichier de config

            ' Configure open file dialog box
            Dim dlg As New Microsoft.Win32.SaveFileDialog()
            dlg.FileName = "" ' Default file name
            dlg.DefaultExt = ".xml" ' Default file extension
            dlg.Filter = "Fichier de template (.xml)|*.xml" ' Filter files by extension

            ' Show open file dialog box
            Dim result As Boolean = dlg.ShowDialog()

            ' Process open file dialog box results
            If result = True And _CurrentTemplate IsNot Nothing Then
                ' Open document
                Dim filename As String = dlg.FileName
                Dim retour As String = myService.ExportTemplateMultimedia(IdSrv, _CurrentTemplate)
                If retour.StartsWith("ERREUR") Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, retour, "Erreur export template", "")
                Else
                    Dim TargetFile As StreamWriter
                    TargetFile = New StreamWriter(filename, False)
                    TargetFile.Write(retour)
                    TargetFile.Close()
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, "L'export du fichier de template a été effectué", "Export Template", "")
                End If
            End If

            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub BtnExportTemplate_Click: " & ex.Message, "ERREUR", "")
        End Try
    End Sub


    Private Sub RdIR_Checked(sender As Object, e As System.Windows.RoutedEventArgs) Handles RdIR.Checked, RdIR.Click
        If RdIR.IsChecked Then
            Label27.Visibility = Windows.Visibility.Visible
            CbFormat.Visibility = Windows.Visibility.Visible
        Else
            Label27.Visibility = Windows.Visibility.Collapsed
            CbFormat.Visibility = Windows.Visibility.Collapsed
        End If
    End Sub


End Class
