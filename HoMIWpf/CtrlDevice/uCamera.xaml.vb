﻿
Imports System.Windows.Threading
Imports System.Reflection
Imports System.Threading
Imports System.Net
Imports MjpegProcessor
Imports System.IO

Public Class uCamera
    Dim _URL As String = ""
    Dim _ListButton As New List(Of uHttp.ButtonHttp)
    Dim _AsError As Boolean
    Private _mjpeg As MjpegDecoder

    Public Property URL As String
        Get
            Return _URL
        End Get
        Set(ByVal value As String)
            Try
                _URL = value
                If String.IsNullOrEmpty(_URL) = False Then 
                    lbl.Visibility = Windows.Visibility.Collapsed
                    lbl.Text = ""
                    _mjpeg.ParseStream(New Uri(_URL))
                End If
            Catch ex As Exception
                '                lbl.Text = "Erreur: " & ex.Message
                AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uCamera.URL : " & _URL & vbCr & ex.Message, "Erreur", "uCamera.URL")
                lbl.Visibility = Windows.Visibility.Visible
            End Try
        End Set
    End Property

    Public Property ListButton As List(Of uHttp.ButtonHttp)
        Get
            Return _ListButton
        End Get
        Set(ByVal value As List(Of uHttp.ButtonHttp))
            _ListButton = value

            Try

                If _ListButton IsNot Nothing Then
                    StkButton.Children.Clear()
                    For Each _button As uHttp.ButtonHttp In _ListButton
                        Dim x As New uHttp.ButtonHttp
                        x.Foreground = Brushes.White
                        x.Margin = New Thickness(5)
                        x.Height = _button.Height
                        x.Width = _button.Width
                        x.Content = _button.Content
                        x.URL = _button.URL
                        x.SetResourceReference(Control.TemplateProperty, "GlassButton")
                        AddHandler x.Click, AddressOf Button_Click
                        StkButton.Children.Add(x)
                    Next
                End If
            Catch ex As Exception
                AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uCamera.ListButton.set: " & ex.Message, "Erreur", "uCamera.ListButton.set")
            End Try
        End Set
    End Property

    Private Sub Button_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
        Try
            Dim x As uHttp.ButtonHttp = sender
            'If My.Computer.Network.IsAvailable = True And String.IsNullOrEmpty(x.URL) = False Then
            If String.IsNullOrEmpty(x.URL) = False Then
                Dim reader As StreamReader = Nothing
                Dim str As String = ""
                Dim request As WebRequest = WebRequest.Create(x.URL)
                Dim response As WebResponse = request.GetResponse()

                reader = New StreamReader(response.GetResponseStream())
                str = reader.ReadToEnd
                reader.Close()
            Else
                MessageBox.Show("Erreur l'url: " & x.URL & " n'est pas valide", "Erreur", MessageBoxButton.OK, MessageBoxImage.Error)
            End If
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uCamera.Button_Click: " & ex.Message, "Erreur", "uCamera.Button_Click")
        End Try
    End Sub


    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        Try
            ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
            _mjpeg = New MjpegDecoder
            AddHandler _mjpeg.FrameReady, AddressOf mjpeg_FrameReady

        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uCamera.New: " & ex.Message, "Erreur", "uCamera.New")
        End Try
    End Sub

    Private Sub uCamera_SizeChanged(ByVal sender As Object, ByVal e As System.Windows.SizeChangedEventArgs) Handles Me.SizeChanged
        Try

            Dim _size As Size = e.NewSize
            Dim y As Double = _size.Height

            If _ListButton.Count > 0 Then
                y = y - StkButton.ActualHeight - 30
            End If

            image.Width = _size.Width
            image.Height = y

        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uCamera.uCamera_SizeChanged: " & ex.ToString, "Erreur", "uCamera.uCamera_SizeChanged")
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        Try
            If _mjpeg IsNot Nothing Then
                _mjpeg.StopStream()
                _mjpeg = Nothing
            End If
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uCamera.Finalize: " & ex.ToString, "Erreur", "uCamera.Finalize")
        End Try
        MyBase.Finalize()
    End Sub

    Private Sub uCamera_Unloaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Unloaded
        Try
            If _mjpeg IsNot Nothing Then
                _mjpeg.StopStream()
                _mjpeg = Nothing
            End If
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uCamera.uCamera_Unloaded: " & ex.ToString, "Erreur", "uCamera.uCamera_Unloaded")
        End Try
    End Sub

    Private Sub mjpeg_FrameReady(ByVal sender As Object, ByVal e As FrameReadyEventArgs)
        Try
            image.Source = e.BitmapImage
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uCamera.mjpeg_FrameReady: " & ex.ToString, "Erreur", "uCamera.mjpeg_FrameReady")
        End Try
    End Sub

End Class
