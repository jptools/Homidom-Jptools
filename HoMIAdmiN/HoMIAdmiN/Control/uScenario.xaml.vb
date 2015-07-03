﻿Imports HoMIDom.HoMIDom
Imports System.Threading

Public Class uScenario

    Dim Span As Integer = 60 / 5 'Espacement correspondant à 1 seconde default 60/5
    Dim _Duree As Integer = 10 'default 2
    Dim _Zoom As Integer = 1 'default 1
    Dim _ListAction As New ArrayList 'liste des actions
    Dim t As Double
    Dim _Width As Double = 0

    Public Event AsChange()

    'Duree max du timeline en minutes
    Public Property Duree As Integer
        Get
            Return _Duree
        End Get
        Set(ByVal value As Integer)
            _Duree = value
        End Set
    End Property

    'Zoom du timeline
    Public Property Zoom As Integer
        Get
            Return _Zoom
        End Get
        Set(ByVal value As Integer)
            _Zoom = value
        End Set
    End Property

    'Items (actions du timeline)

    Public Property Items As ArrayList
        Get
            Return _ListAction
        End Get
        Set(ByVal value As ArrayList)
            Try
                For i As Integer = 0 To value.Count - 1
                    Dim x As New uAction

                    x.ObjAction = value.Item(i)
                    x.Uid = HoMIDom.HoMIDom.Api.GenerateGUID
                    x.Span = Span
                    x.Zoom = _Zoom

                    AddHandler x.DeleteAction, AddressOf DeleteAction
                    AddHandler x.ChangeAction, AddressOf ChangeAction

                    Dim j As Double = x.ObjAction.Timing.Minute + (x.ObjAction.Timing.Hour * 60)

                    If j >= Duree Then
                        Duree = j + 1
                        StckPnlLib.Dispatcher.BeginInvoke(New Affiche_Label2(AddressOf Affiche_Label))
                        StckPnlLibTr.Dispatcher.BeginInvoke(New Affiche_Trait2(AddressOf Affiche_Trait))
                    End If

                    _ListAction.Add(value.Item(i))
                    StackPanel1.Children.Add(x)

                    Me.Dispatcher.BeginInvoke(New Affiche_Action2(AddressOf Affiche_Action))
                    x = Nothing
                Next

            Catch ex As Exception
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario Items: " & ex.Message, "ERREUR", "")
            End Try

        End Set
    End Property



    Public Sub New()
        ' Cet appel est requis par le concepteur.

        InitializeComponent()

        Try
            Me.Cursor = Cursors.Wait

            ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().

            Me.Dispatcher.BeginInvoke(New Affiche_Label2(AddressOf Affiche_Label))
            Me.Dispatcher.BeginInvoke(New Affiche_Trait2(AddressOf Affiche_Trait))
            Me.Dispatcher.BeginInvoke(New Affiche_Action2(AddressOf Affiche_Action))

            Me.Cursor = Nothing

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario New: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub ScrollViewer1_DragOver(ByVal sender As System.Object, ByVal e As System.Windows.DragEventArgs) Handles ScrollViewer2.DragOver
        Try
            If e.Data.GetDataPresent(GetType(String)) Then
                e.Effects = DragDropEffects.Copy
            Else
                e.Effects = DragDropEffects.None
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario ScrollViewer1_DragOver: " & ex.Message, "ERREUR", "")
        End Try
    End Sub

    Private Sub ScrollViewer1_Drop(ByVal sender As System.Object, ByVal e As System.Windows.DragEventArgs) Handles ScrollViewer2.Drop
        Try
            If e.Data.GetDataPresent(GetType(String)) Then
                e.Effects = DragDropEffects.Copy

                Dim uri As String = e.Data.GetData(GetType(String)).ToString

                Dim x As New uAction
                x.Uid = HoMIDom.HoMIDom.Api.GenerateGUID

                AddHandler x.DeleteAction, AddressOf DeleteAction
                AddHandler x.ChangeAction, AddressOf ChangeAction

                x.Width = _Width '(_Duree * 3600) + 100 'StckPnlLib.ActualWidth

                Select Case UCase(uri)
                    Case "ACTIONDEVICE"
                        Dim y As New Action.ActionDevice
                        x.ObjAction = y
                    Case "ACTIONDRIVER"
                        Dim y As New Action.ActionDriver
                        x.ObjAction = y
                    Case "ACTIONMAIL"
                        Dim y As New Action.ActionMail
                        x.ObjAction = y
                    Case "ACTIONSPEECH"
                        Dim y As New Action.ActionSpeech
                        x.ObjAction = y
                    Case "ACTIONHTTP"
                        Dim y As New Action.ActionHttp
                        x.ObjAction = y
                    Case "ACTIONIF"
                        Dim y As New Action.ActionIf
                        x.ObjAction = y
                    Case "ACTIONMACRO"
                        Dim y As New Action.ActionMacro
                        x.ObjAction = y
                    Case "ACTIONLOGEVENT"
                        Dim y As New Action.ActionLogEvent
                        x.ObjAction = y
                    Case "ACTIONLOGEVENTHOMIDOM"
                        Dim y As New Action.ActionLogEventHomidom
                        x.ObjAction = y
                    Case "ACTIONDOS"
                        Dim y As New Action.ActionDos
                        x.ObjAction = y
                    Case "ACTIONVB"
                        Dim y As New Action.ActionVB
                        y.Script = My.Resources.ExempleVBS
                        x.ObjAction = y
                    Case "ACTIONSTOP"
                        Dim y As New Action.ActionSTOP
                        x.ObjAction = y
                    Case "ACTIONVAR"
                        Dim y As New Action.ActionVar
                        x.ObjAction = y
                End Select

                x.Span = Span
                x.Zoom = _Zoom

                _ListAction.Add(x.ObjAction)

                StackPanel1.Children.Add(x)

                RaiseEvent AsChange()
            Else
                e.Effects = DragDropEffects.None
            End If

            ScrollViewer2.UpdateLayout()
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & ex.ToString, "ERREUR", "")
        End Try

    End Sub



    'Supprimer une action

    Private Sub DeleteAction(ByVal Id As String)
        Try
            For i As Integer = 0 To StackPanel1.Children.Count - 1
                If StackPanel1.Children.Item(i).Uid = Id Then
                    StackPanel1.Children.RemoveAt(i)
                    _ListAction.RemoveAt(i)
                    Exit For
                End If
            Next

            RaiseEvent AsChange()
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario DeleteAction: " & ex.Message, "ERREUR", "")
        End Try

    End Sub



    'Mise à jour d'une action

    Private Sub ChangeAction(ByVal Id As String)
        Try

            For i As Integer = 0 To StackPanel1.Children.Count - 1
                Dim x As uAction = StackPanel1.Children.Item(i)
                Dim j As Double = x.ObjAction.Timing.Minute + (x.ObjAction.Timing.Hour * 60)

                If j >= Duree Then
                    Duree = j + 1
                    StckPnlLib.Dispatcher.BeginInvoke(New Affiche_Label2(AddressOf Affiche_Label))
                    StckPnlLibTr.Dispatcher.BeginInvoke(New Affiche_Trait2(AddressOf Affiche_Trait))
                End If

                x = Nothing
                j = 0
            Next

            Me.Dispatcher.BeginInvoke(New Affiche_Action2(AddressOf Affiche_Action))

            RaiseEvent AsChange()
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario ChangeAction: " & ex.Message, "ERREUR", "")
        End Try

    End Sub


    'Afficher les éléments du timeline
    Private Sub Afficher()

        Try
            Me.Dispatcher.BeginInvoke(New Affiche_Label2(AddressOf Affiche_Label))
            Me.Dispatcher.BeginInvoke(New Affiche_Trait2(AddressOf Affiche_Trait))

            'On affecte la valeur du Zoom à chaque action
            Me.Dispatcher.BeginInvoke(New Affiche_Action2(AddressOf Affiche_Action))

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur Affiche: " & ex.ToString)
        End Try

    End Sub

    Private Delegate Sub Affiche_Action2()

    Private Sub Affiche_Action()
        Try
            If StackPanel1 IsNot Nothing Then
                For i As Integer = 0 To StackPanel1.Children.Count - 1
                    Dim x As uAction = StackPanel1.Children.Item(i)
                    x.Width = _Width '(_Duree * 60) * 60 '+ 100
                    x.Zoom = _Zoom
                    x = Nothing
                Next

            End If

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario Affiche_Action: " & ex.Message, "ERREUR", "")
        End Try

    End Sub



    Private Delegate Sub Affiche_Label2()

    Private Sub Affiche_Label()

        Try

            If StckPnlLib IsNot Nothing Then StckPnlLib.Children.Clear()

            t = (_Duree * 60) / _Zoom

            '_Width = 130

            Dim _time As DateTime

            For j As Integer = 0 To (t / 5) 'ajout des labels temps

                Dim x As New Label

                With x

                    x.FontSize = 10

                    x.HorizontalContentAlignment = HorizontalAlignment.Left

                    x.Width = 60

                    '_Width += 60

                    x.Foreground = New SolidColorBrush(Colors.White)

                    x.Content = _time.ToLongTimeString

                End With

                StckPnlLib.Children.Add(x)

                _time = _time.AddSeconds(5 * _Zoom)

            Next

            Dim x2 As New Label 'ajout d'un label vide à la fin du timeline pour avoir un espace 

            x2.Width = 60

            StckPnlLib.Children.Add(x2)



            x2 = Nothing

        Catch ex As Exception

            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur Affiche_Label: " & ex.ToString, "ERREUR", "")

        End Try

    End Sub

    Private Delegate Sub Affiche_Trait2()

    Private Sub Affiche_Trait()
        Try
            If StckPnlLibTr IsNot Nothing Then StckPnlLibTr.Children.Clear()
            t = _Duree * 60 / _Zoom

            Dim flag As Boolean

            For j As Integer = 0 To (t / 5)
                Dim y1 As New Canvas
                y1.Width = 60

                If flag = False Then
                    y1.Background = New SolidColorBrush(Colors.Transparent)
                    flag = True
                Else
                    y1.Background = New SolidColorBrush(Colors.Transparent)
                    flag = False
                End If

                Dim R1 As New Rectangle

                R1.Fill = New SolidColorBrush(Colors.White)
                R1.Width = 3
                R1.Height = 16
                y1.Children.Add(R1)

                For k = 1 To 4

                    Dim R2 As New Rectangle

                    R2.Fill = New SolidColorBrush(Colors.White)
                    R2.Width = 1
                    R2.Height = 8
                    y1.Children.Add(R2)

                    Canvas.SetLeft(R2, k * Span)
                Next

                StckPnlLibTr.Children.Add(y1)
                y1 = Nothing

            Next

            StckPnlLibTr.UpdateLayout()
            ScrollViewer2.UpdateLayout()

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur Affiche_Trait: " & ex.ToString)
        End Try

    End Sub

    'Zoom avant
    Private Sub ZoomPlus_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ZoomPlus.MouseDown

        Try

            _Zoom -= 5

            If _Zoom <= 0 Then _Zoom = 1

            Afficher()

        Catch ex As Exception

            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario ZoomPlus_MouseDown: " & ex.Message, "ERREUR", "")

        End Try

    End Sub

    'Zoom arrière
    Private Sub ZoomMoins_MouseDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ZoomMoins.MouseDown

        Try

            _Zoom += 5

            Afficher()

        Catch ex As Exception

            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario ZoomMoins_MouseDown: " & ex.Message, "ERREUR", "")

        End Try

    End Sub

    Private Sub ScrollViewer2_ScrollChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.ScrollChangedEventArgs) Handles ScrollViewer2.ScrollChanged
        Try
            ScrollViewer2.ScrollToHorizontalOffset(e.HorizontalOffset)
            ScrollViewer2.ScrollToVerticalOffset(e.VerticalOffset)
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "ERREUR Sub uScenario ScrollViewer1_ScrollChanged: " & ex.Message, "ERREUR", "")
        End Try
    End Sub



#Region "Ajout Action"

    'Ajouter action device

    Private Sub Image1_MouseLeftButtonDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActDevice.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONDEVICE")

        effects = DragDrop.DoDragDrop(Me.ImgActDevice, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action driver

    Private Sub ImgActDriver_MouseLeftButtonDown(ByVal sender As System.Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActDriver.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONDRIVER")

        effects = DragDrop.DoDragDrop(Me.ImgActDevice, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action mail

    Private Sub ImgActMail_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActMail.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONMAIL")

        effects = DragDrop.DoDragDrop(Me.ImgActMail, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action if

    Private Sub ImgActIf_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActIf.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONIF")

        effects = DragDrop.DoDragDrop(Me.ImgActIf, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action macro

    Private Sub ImgActMacro_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActMacro.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONMACRO")

        effects = DragDrop.DoDragDrop(Me.ImgActMacro, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action PArler

    Private Sub ImgActSpeech_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActSpeech.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONSPEECH")

        effects = DragDrop.DoDragDrop(Me.ImgActSpeech, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action http

    Private Sub ImgActHttp_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActHttp.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONHTTP")

        effects = DragDrop.DoDragDrop(Me.ImgActHttp, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action log event Windows

    Private Sub ImgActLog_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActLog.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONLOGEVENT")

        effects = DragDrop.DoDragDrop(Me.ImgActLog, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action log event Homidom

    Private Sub ImgActLogHomidom_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActLogHomidom.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONLOGEVENTHOMIDOM")

        effects = DragDrop.DoDragDrop(Me.ImgActLogHomidom, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action Dos

    Private Sub ImgActDos_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActDos.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONDOS")

        effects = DragDrop.DoDragDrop(Me.ImgActDos, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action Var

    Private Sub ImgActVar_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActVAR.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONVAR")

        effects = DragDrop.DoDragDrop(Me.ImgActVAR, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action VB

    Private Sub ImgActVB_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActVB.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONVB")

        effects = DragDrop.DoDragDrop(Me.ImgActVB, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub



    'Ajouter action STOP

    Private Sub ImgActSTOP_MouseLeftButtonDown(ByVal sender As Object, ByVal e As System.Windows.Input.MouseButtonEventArgs) Handles ImgActSTOP.MouseLeftButtonDown

        Dim effects As DragDropEffects

        Dim obj As New DataObject()

        obj.SetData(GetType(String), "ACTIONSTOP")

        effects = DragDrop.DoDragDrop(Me.ImgActSTOP, obj, DragDropEffects.Copy Or DragDropEffects.Move)

    End Sub

#End Region


    Private Sub uScenario_Loaded(sender As System.Object, e As System.Windows.RoutedEventArgs) Handles MyBase.Loaded
        'StackPanel1.Height = uScenario.Height - 110
    End Sub

End Class
