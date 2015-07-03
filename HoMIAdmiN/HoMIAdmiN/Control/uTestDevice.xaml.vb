﻿Imports HoMIDom.HoMIDom.Api

Partial Public Class uTestDevice
    Dim _DeviceId As String
    Dim _Device As HoMIDom.HoMIDom.TemplateDevice

    Public Event CloseMe(ByVal MyObject As Object)

    Private Sub BtnTest_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnTest.Click
        Try
            If CbCmd.SelectedItem IsNot Nothing Then
                Dim y As Label = CbCmd.SelectedItem
                Dim x As New HoMIDom.HoMIDom.DeviceAction

                If y.Tag = 0 Then 'c une fonction de base
                    x.Nom = y.Content
                    For i As Integer = 0 To StkParam.Children.Count - 1
                        Dim stk As StackPanel = StkParam.Children.Item(i)
                        Dim txt As TextBox = stk.Children.Item(1)
                        Dim param As New HoMIDom.HoMIDom.DeviceAction.Parametre
                        param.Value = txt.Text
                        x.Parametres.Add(param)
                    Next
                    myservice.ExecuteDeviceCommand(IdSrv, _DeviceId, x)
                Else 'c une fonction avancée
                    x.Nom = "ExecuteCommand"

                    Dim param As New HoMIDom.HoMIDom.DeviceAction.Parametre
                    param.Value = y.Content
                    x.Parametres.Add(param)

                    For i As Integer = 0 To StkParam.Children.Count - 1
                        Dim stk As StackPanel = StkParam.Children.Item(i)
                        Dim txt As TextBox = stk.Children.Item(1)
                        Dim param1 As New HoMIDom.HoMIDom.DeviceAction.Parametre
                        param1.Value = txt.Text
                        x.Parametres.Add(param1)
                    Next

                    myService.ExecuteDeviceCommand(IdSrv, _DeviceId, x)
                End If
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors du test: " & ex.Message, "Erreur", "")
        End Try
    End Sub

    Public Sub New(ByVal DeviceId As String)

        ' Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        Try

            ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
            _DeviceId = DeviceId

            _Device = myService.ReturnDeviceByID(IdSrv, _DeviceId)

            If _Device IsNot Nothing Then

                For i As Integer = 0 To _Device.DeviceAction.Count - 1
                    Dim x As New Label
                    x.Content = _Device.DeviceAction.Item(i).Nom
                    x.Tag = 0 'c une fonction de base
                    CbCmd.Items.Add(x)
                Next

                For i As Integer = 0 To _Device.GetDeviceCommandePlus.Count - 1
                    Dim x As New Label
                    x.Content = _Device.GetDeviceCommandePlus.Item(i).NameCommand
                    x.ToolTip = _Device.GetDeviceCommandePlus.Item(i).DescriptionCommand
                    x.Tag = 1 'c une fonction avancée
                    CbCmd.Items.Add(x)
                Next

            Else
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le Device est inconnu !!", "Erreur", "")
                RaiseEvent CloseMe(Me)
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & ex.Message, "Erreur", "")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnClose.Click
        RaiseEvent CloseMe(Me)
    End Sub

    Private Sub CbCmd_SelectionChanged(ByVal sender As Object, ByVal e As System.Windows.Controls.SelectionChangedEventArgs) Handles CbCmd.SelectionChanged
        Try
            If CbCmd.SelectedIndex < 0 Then Exit Sub
            StkParam.Children.Clear()

            Dim x As Label = CbCmd.SelectedItem

            If x.Tag = 0 Then 'c une fonction de base
                For i As Integer = 0 To _Device.DeviceAction.Count - 1
                    If _Device.DeviceAction.Item(i).Nom = x.Content Then
                        For j As Integer = 0 To _Device.DeviceAction.Item(i).Parametres.Count - 1
                            Dim stk As New StackPanel
                            stk.Orientation = Orientation.Horizontal

                            Dim lbl As New Label
                            lbl.Content = _Device.DeviceAction.Item(i).Parametres.Item(j).Nom
                            lbl.Width = 76
                            lbl.Foreground = Brushes.White
                            stk.Children.Add(lbl)

                            Dim txt As New TextBox
                            txt.Height = 20
                            txt.Width = 250
                            txt.ToolTip = _Device.DeviceAction.Item(i).Parametres.Item(j).Type
                            txt.Background = Brushes.DarkGray
                            txt.Foreground = Brushes.White
                            txt.BorderBrush = Brushes.Black
                            stk.Children.Add(txt)

                            StkParam.Children.Add(stk)
                        Next
                        Exit Sub
                    End If
                Next
            Else 'c une fonction avancée
                For i As Integer = 0 To _Device.GetDeviceCommandePlus.Count - 1
                    If _Device.GetDeviceCommandePlus.Item(i).NameCommand = x.Content Then
                        For j As Integer = 1 To _Device.GetDeviceCommandePlus.Item(i).CountParam
                            Dim stk As New StackPanel
                            stk.Orientation = Orientation.Horizontal

                            Dim lbl As New Label
                            lbl.Content = "Parametre " & j & ":"
                            lbl.Width = 76
                            lbl.Foreground = Brushes.White
                            stk.Children.Add(lbl)

                            Dim txt As New TextBox
                            txt.Height = 20
                            txt.Width = 250
                            stk.Children.Add(txt)

                            StkParam.Children.Add(stk)
                        Next
                        Exit Sub
                    End If
                Next
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur Sub CbCmd_SelectionChanged: " & ex.Message, "Erreur", "")
        End Try
    End Sub
End Class
