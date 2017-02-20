﻿Imports System.IO
Imports System.Windows.Media.Animation
Imports System.Net
Imports System.Xml
Imports System.Xml.XPath
Imports HoMIDom.HoMIDom
Imports System.Reflection
Imports System.Threading

Module Fonctions
    ''' <summary>
    ''' Permet de vérifier si 2 objets sont identiques au niveau type et valeur(s)
    ''' </summary>
    ''' <param name="objet1"></param>
    ''' <param name="objet2"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsDiff(ByVal objet1 As Object, ByVal objet2 As Object) As Boolean
        Try
            If objet2 Is Nothing Or objet1 Is Nothing Then
                Return False
            ElseIf objet1.GetType <> objet2.GetType Then
                Return False
            End If

            For Each pi As PropertyInfo In objet1.GetType.GetProperties()
                Try
                    If pi.GetValue(objet1, Nothing) <> pi.GetValue(objet2, Nothing) Then
                        Return True
                    End If
                Catch ex As Exception
                    Return True
                End Try
            Next
            Return False

        Catch ex As Exception
            MessageBox.Show("Erreur IsDiff: " & ex.Message)
        End Try
    End Function

    ' ''' <summary>
    ' ''' Permet de recopier les valeurs propriétés d'un objet vers un autre
    ' ''' </summary>
    ' ''' <param name="objet1"></param>
    ' ''' <param name="objet2"></param>
    ' ''' <returns></returns>
    ' ''' <remarks></remarks>
    'Public Function Duplicate(ByVal source As Object, ByVal destination As Object) As Boolean
    '    Try
    '        If source Is Nothing Then
    '            Return False
    '        ElseIf source.GetType <> destination.GetType Then
    '            Return False
    '        End If

    '        For Each pi As PropertyInfo In source.GetType.GetProperties()
    '            Try
    '                pi.SetValue (source,pi.GetValue (source,
    '                If pi.GetValue(source, Nothing) <> pi.GetValue(objet2, Nothing) Then
    '                    Return True
    '                End If
    '            Catch ex As Exception
    '                Return True
    '            End Try
    '        Next
    '        Return False

    '    Catch ex As Exception
    '        MessageBox.Show("Erreur IsDiff: " & ex.Message)
    '    End Try
    'End Function


    Public Function ConvertArrayToImage(ByVal value As Byte()) As BitmapImage
        Try
            Dim ImgSource As BitmapImage = Nothing
            Dim array As Byte() = TryCast(value, Byte())

            If array IsNot Nothing Then
                ImgSource = New BitmapImage()
                ImgSource.BeginInit()
                ImgSource.CacheOption = BitmapCacheOption.OnLoad
                ImgSource.CreateOptions = BitmapCreateOptions.DelayCreation
                ImgSource.StreamSource = New MemoryStream(array)
                array = Nothing
                ImgSource.EndInit()
                If ImgSource.CanFreeze Then ImgSource.Freeze()
            End If

            Return ImgSource
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "ERREUR Sub ConvertArrayToImage: " & ex.Message, "Erreur", "ConvertArrayToImage")
            Return Nothing
        End Try
    End Function

    Public Sub ScrollToPosition(ByVal ScrollViewer As UAniScrollViewer.AniScrollViewer, ByVal x As Double, ByVal y As Double, ByVal Duree As Double)
        Dim vertAnim As New DoubleAnimation()
        vertAnim.From = ScrollViewer.VerticalOffset
        vertAnim.[To] = y
        vertAnim.DecelerationRatio = 0.99
        vertAnim.Duration = New Duration(TimeSpan.FromMilliseconds(Duree))

        Dim horzAnim As New DoubleAnimation()
        horzAnim.From = ScrollViewer.HorizontalOffset
        horzAnim.[To] = x
        horzAnim.DecelerationRatio = 0.99
        horzAnim.Duration = New Duration(TimeSpan.FromMilliseconds(Duree))

        Dim sb As New Storyboard()
        sb.Children.Add(vertAnim)
        sb.Children.Add(horzAnim)

        Storyboard.SetTarget(vertAnim, ScrollViewer)
        Storyboard.SetTargetProperty(vertAnim, New PropertyPath(UAniScrollViewer.AniScrollViewer.CurrentVerticalOffsetProperty))
        Storyboard.SetTarget(horzAnim, ScrollViewer)
        Storyboard.SetTargetProperty(horzAnim, New PropertyPath(UAniScrollViewer.AniScrollViewer.CurrentHorizontalOffsetProperty))

        sb.Begin()

    End Sub

    Public Function UrlIsValid(ByVal url As String) As Boolean
        Dim is_valid As Boolean = False
        If url.ToLower().StartsWith("www.") Then url = _
            "http://" & url

        Dim web_response As HttpWebResponse = Nothing
        Try
            Dim web_request As HttpWebRequest = _
                HttpWebRequest.Create(url)
            web_response = _
                DirectCast(web_request.GetResponse(),  _
                HttpWebResponse)
            Return True
            web_request = Nothing
            web_response = Nothing
        Catch ex As Exception
            Return False
        Finally
            If Not (web_response Is Nothing) Then _
                web_response.Close()
        End Try
    End Function

    ''' <summary>
    ''' Vérifie si la valeur est un boolean
    ''' </summary>
    ''' <param name="value"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function IsBoolean(ByVal value As Object) As Boolean
        Try
            Dim x As Boolean = value
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Fonction permettant de charger une image 
    ''' </summary>
    ''' <param name="FileChm">chemin du fichier</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function LoadBitmapImage(ByVal FileChm As String) As BitmapImage
        Dim bmpImage As New BitmapImage

        Try
            If File.Exists(FileChm) Then
                bmpImage.BeginInit()
                bmpImage.CacheOption = BitmapCacheOption.OnLoad
                bmpImage.CreateOptions = BitmapCreateOptions.DelayCreation
                bmpImage.UriSource = New Uri(FileChm, UriKind.Absolute)
                bmpImage.EndInit()
                If bmpImage.CanFreeze Then bmpImage.Freeze()
            End If
            Return bmpImage
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "ERREUR Sub LoadBitmapImage (FileChm= " & FileChm & "): " & ex.Message, "Erreur", "LoadBitmapImage")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Retourne la valeur d'une variable entourée par des balises inferieur et superieur
    ''' </summary>
    ''' <param name="ValueTxt"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function TraiteBalise(ByVal ValueTxt As String) As String
        Try
            If String.IsNullOrEmpty(ValueTxt) Then Return ""
            Dim x As String = ValueTxt.ToUpper.Trim(" ")

            If x.StartsWith("<") And x.EndsWith(">") Then
                Dim _val As String = Mid(x, 2, Len(x) - 2)
                Select Case _val
                    Case "SYSTEM_DATE"
                        Return Now.Date.ToShortDateString
                    Case "SYSTEM_LONG_DATE"
                        Return Now.Date.ToLongDateString
                    Case "SYSTEM_TIME"
                        Return Now.ToShortTimeString
                    Case "SYSTEM_LONG_TIME"
                        Return Now.ToLongTimeString
                    Case "SYSTEM_SOLEIL_COUCHE"
                        If IsConnect Then
                            Dim _date As Date = myService.GetHeureCoucherSoleil
                            Return _date.ToShortTimeString
                        Else
                            Return ""
                        End If
                    Case "SYSTEM_SOLEIL_LEVE"
                        If IsConnect Then
                            Dim _date As Date = myService.GetHeureLeverSoleil
                            Return _date.ToShortTimeString
                        Else
                            Return ""
                        End If
                    Case "SYSTEM_CONDITION"
                        If AllDevices IsNot Nothing And String.IsNullOrEmpty(frmMere.Ville) = False Then
                            For Each ObjMeteo As HoMIDom.HoMIDom.TemplateDevice In AllDevices
                                If ObjMeteo.Type = HoMIDom.HoMIDom.Device.ListeDevices.METEO And ObjMeteo.Enable = True And ObjMeteo.Name.ToUpper = frmMere.Ville.ToUpper Then
                                    If frmMere.MaJWidgetFromServer And IsConnect Then
                                        Return myService.ReturnDeviceByID(IdSrv, ObjMeteo.ID).ConditionActuel
                                    Else
                                        Return ObjMeteo.ConditionActuel
                                    End If
                                End If
                            Next
                        Else
                            Return ""
                        End If
                    Case "SYSTEM_TEMP_ACTUELLE"
                        If AllDevices IsNot Nothing And String.IsNullOrEmpty(frmMere.Ville) = False Then
                            For Each ObjMeteo As HoMIDom.HoMIDom.TemplateDevice In AllDevices
                                If ObjMeteo.Type = HoMIDom.HoMIDom.Device.ListeDevices.METEO And ObjMeteo.Enable = True And ObjMeteo.Name.ToUpper = frmMere.Ville.ToUpper Then
                                    If frmMere.MaJWidgetFromServer And IsConnect Then
                                        Return myService.ReturnDeviceByID(IdSrv, ObjMeteo.ID).TemperatureActuel & " °C"
                                    Else
                                        Return ObjMeteo.TemperatureActuel & " °C"
                                    End If
                                End If
                            Next
                        Else
                            Return "# °C"
                        End If
                    Case "SYSTEM_ICO_METEO"
                        If AllDevices IsNot Nothing And String.IsNullOrEmpty(frmMere.Ville) = False Then
                            For Each ObjMeteo As HoMIDom.HoMIDom.TemplateDevice In AllDevices
                                If ObjMeteo.Type = HoMIDom.HoMIDom.Device.ListeDevices.METEO And ObjMeteo.Enable = True And ObjMeteo.Name.ToUpper = frmMere.Ville.ToUpper Then
                                    If frmMere.MaJWidgetFromServer And IsConnect Then
                                        Return myService.ReturnDeviceByID(IdSrv, ObjMeteo.ID).IconActuel
                                    Else
                                        Return ObjMeteo.IconActuel
                                    End If
                                End If
                            Next
                        Else
                            Return ""
                        End If
                    Case Else
                        Dim a As String = myService.GetValueOfVariable(IdSrv, _val)
                        If String.IsNullOrEmpty(a) = False Then
                            Return a
                        Else
                            Return " "
                        End If
                End Select
            Else
                Return ValueTxt
            End If
            Return ""
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur Fonctions.TraiteBalise: " & ex.ToString, "Erreur", " Fonctions.TraiteBalise")
            Return "Erreur"
        End Try
    End Function

    Public Sub Refresh()
        Try
            If IsConnect Then
                Do While lock_dev
                    Thread.Sleep(100)
                Loop

                lock_dev = True
                AllDevices = myService.GetAllDevices(IdSrv)
                lock_dev = False
            End If
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur Fonctions.Thread_MAJ.Refresh: " & ex.ToString, "Erreur", " Fonctions.Thread_MAJ.Refresh")
        End Try
    End Sub

    Public Sub Refresh_Zone()
        Try
            If IsConnect Then
                _AllZones = myService.GetAllZones(IdSrv) 'recup l'image des zones
            End If
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur Fonctions.Thread_MAJ.Refresh: " & ex.ToString, "Erreur", " Fonctions.Thread_MAJ.Refresh")
        End Try
    End Sub

    Public Function ReturnDeviceById(DeviceId As String) As TemplateDevice
        Try
            If AllDevices IsNot Nothing Then
                Dim retour As TemplateDevice = Nothing

                Do While lock_dev
                    Thread.Sleep(100)
                Loop

                lock_dev = True
                For Each _dev In AllDevices
                    If _dev.ID = DeviceId Then
                        retour = _dev
                        Exit For
                    End If
                Next

                lock_dev = False
                Return retour
            End If
            Return Nothing
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur Fonctions.ReturnDeviceById: " & ex.ToString, "Erreur", " Fonctions.ReturnDeviceById")
            Return Nothing
        End Try
    End Function

End Module
