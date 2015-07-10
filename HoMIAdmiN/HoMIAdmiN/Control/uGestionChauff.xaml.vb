
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports System.Xml
Imports HoMIDom.HoMIDom.Api
Imports OfficeOpenXml
Imports System.Collections.Generic
Imports System.Data


' Auteur : JPTOOLS
' Date : 17/06/2015

'Class permettant d'afficher le mode du chauffage en lisant un fichier Excel.xlsx
'Nécessite l'installation de la librairie EPPlus.dll
#Region "Class pour List"
'Class Calendrier pour affichage
Public Class Calendrier
    Public Property Semaine() As String
        Get
            Return _Semaine
        End Get
        Set(value As String)
            _Semaine = value
        End Set
    End Property
    Private _Semaine As String

    Public Property Mode() As String
        Get
            Return _Mode
        End Get
        Set(value As String)
            _Mode = value
        End Set
    End Property
    Public _Mode As String
End Class


'Class Semaine pour affichage
Public Class Semaine
    Public Property Heure() As String
        Get
            Return _Heure
        End Get
        Set(value As String)
            _Heure = value
        End Set
    End Property
    Private _Heure As String

    Public Property Lundi() As String
        Get
            Return _Lundi
        End Get
        Set(value As String)
            _Lundi = value
        End Set
    End Property
    Private _Lundi As String

    Public Property Mardi() As String
        Get
            Return _Mardi
        End Get
        Set(value As String)
            _Mardi = value
        End Set
    End Property
    Private _Mardi As String

    Public Property Mercredi() As String
        Get
            Return _Mercredi
        End Get
        Set(value As String)
            _Mercredi = value
        End Set
    End Property
    Private _Mercredi As String

    Public Property Jeudi() As String
        Get
            Return _Jeudi
        End Get
        Set(value As String)
            _Jeudi = value
        End Set
    End Property
    Private _Jeudi As String

    Public Property Vendredi() As String
        Get
            Return _Vendredi
        End Get
        Set(value As String)
            _Vendredi = value
        End Set
    End Property
    Private _Vendredi As String

    Public Property Samedi() As String
        Get
            Return _Samedi
        End Get
        Set(value As String)
            _Samedi = value
        End Set
    End Property
    Private _Samedi As String

    Public Property Dimanche() As String
        Get
            Return _Dimanche
        End Get
        Set(value As String)
            _Dimanche = value
        End Set
    End Property
    Private _Dimanche As String

End Class

#End Region

Public Class uGestionChauff
    Dim _DeviceId As String
    Dim _Device As HoMIDom.HoMIDom.TemplateDevice
    Dim newFile As FileInfo
    Dim pck As ExcelPackage
    Dim _listCalendar As New List(Of Calendrier)()
    Dim _listSemaineConger As New List(Of Semaine)()
    Dim _listSemaineNormal As New List(Of Semaine)()
    Dim _listSemaineReduit As New List(Of Semaine)()
    Dim _listSemaineCharger As New List(Of Semaine)()
    Dim _listSemaineAbsence As New List(Of Semaine)()
    Dim _Driver As HoMIDom.HoMIDom.TemplateDriver
    Dim IsActiveLecture As Boolean
    Dim _Mode As String
    Dim _IsModifier As Boolean = False
    Dim _Valeur As String = ""
    Dim Remplir_Fond As Boolean = True
    Dim _Colonne As Integer
    Dim _Ligne As Integer
    Dim FichierOk As Boolean = False

    Public Event CloseMe(ByVal MyObject As Object)


    Public Sub New(ByVal DeviceId As String, ByVal Driver As HoMIDom.HoMIDom.TemplateDriver)

        ' Cet appel est requis par le Concepteur Windows Form.
        InitializeComponent()

        Try

            _DeviceId = DeviceId
            _Device = myService.ReturnDeviceByID(IdSrv, _DeviceId)
            _Driver = Driver

            If _Device IsNot Nothing And _Driver IsNot Nothing Then
                IsActiveLecture = _Driver.Parametres.Item(5).Valeur
                _Driver.Parametres.Item(5).Valeur = False

                If _Driver.Parametres.Item(1).Valeur <> "" Then
                    newFile = New FileInfo(_Driver.Parametres.Item(1).Valeur) 'Ouverture du fichier Excel
                    If newFile.Exists Then
                        FichierOk = True
                        pck = New ExcelPackage(newFile)
                        LireCalendrier()
                        AddHandler DataGrid1.Loaded, AddressOf DataGridOk
                    Else
                        FichierOk = False
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Fichier non trouvé :" & _Driver.Parametres.Item(1).Valeur, "Erreur", "")                        
                    End If
                Else
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Fichier Excel non définit !!", "Erreur", "")
                    RaiseEvent CloseMe(Me)
                End If
            Else
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Le Device ou le Driver est inconnu !!", "Erreur", "")
                RaiseEvent CloseMe(Me)
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & ex.Message, "Erreur", "")
        End Try
    End Sub

    Private Sub DataGridOk(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs)
        Try
            If DataGrid1.Columns.Count > 0 Then
                DataGrid1.Columns(0).IsReadOnly = True
            End If

            If DataGrid1.Columns.Count > 2 Then
                DataGrid1.Columns(0).Width = 45
                DataGrid1.Columns(1).Width = 63
                DataGrid1.Columns(2).Width = 63
                DataGrid1.Columns(3).Width = 65
                DataGrid1.Columns(4).Width = 65
                DataGrid1.Columns(5).Width = 65
                DataGrid1.Columns(6).Width = 65
                DataGrid1.Columns(7).Width = 65

            ElseIf DataGrid1.Columns.Count = 2 Then
                DataGrid1.Columns(0).Width = 65
                DataGrid1.Columns(1).Width = 65
            End If
            Colorier()
            If Remplir_Fond = True Then
                DataGrid1.AlternatingRowBackground = New SolidColorBrush(Colors.WhiteSmoke)
            Else
                DataGrid1.AlternatingRowBackground = New SolidColorBrush(Colors.White)
            End If

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur uGestionChauff GridOK: " & ex.ToString, "Erreur Admin", "")
        End Try
    End Sub


    Public Sub LireCalendrier()

        _Mode = "Calendrier"
        ModeChauffage.Text = _Mode        
        ValeursAdmis.Text = "Conger" & vbCrLf & "Normal" & vbCrLf & "Reduit" & vbCrLf & "Charger" & vbCrLf & "Absence"
        DataGrid1.ItemsSource = _listCalendar
        DataGrid1.Items.Refresh()
        DataGridOk(Me, Nothing)

    End Sub


    Public Sub LireModeChauffage(ByRef ListSemaine As List(Of Semaine), ByVal Mode As String)

        _Mode = Mode
        ModeChauffage.Text = _Mode
        ValeursAdmis.Text = "ECC" & vbCrLf & "EC" & vbCrLf & "PRE"
        DataGrid1.ItemsSource = ListSemaine
        DataGrid1.Items.Refresh()
        DataGridOk(Me, Nothing)

    End Sub

    Public Sub Colorier()

        Dim i As Integer
        Dim j As Integer
        '   Dim dataGridCell As DataGridCell

        For i = 0 To i < DataGrid1.Items.Count
            For j = 0 To j < DataGrid1.Columns.Count

                '  DataGrid1.Columns(j).Items(i).Background = ....
                ' DataGrid1.Items(i).Columns(j).Background = New SolidColorBrush(Colors.Red)
                ' dataGridCell = DataGrid1.Cell(i, j)
                'dataGridCell.Background = New SolidColorBrush(Colors.Red)

                '   DataGrid1.RowBackground = New SolidColorBrush(Colors.Red)

                'DataGrid1.RowHeaderStyle = Colors.Red)
   

             

            Next
        Next


    'DirectCast



    End Sub


    Private Sub BtnOk_Click(sender As Object, e As RoutedEventArgs) Handles BtnOk.Click
        If _IsModifier = True Then
            If MessageBox.Show("Voulez-vous sauver une copie ?", "Sauvegarde", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = MessageBoxResult.OK Then
                SauverCopiefichier()
            End If
            Sauverfichier()
        End If
        _IsModifier = False
        RaiseEvent CloseMe(Me)
    End Sub


    Private Sub BtnFermer_Click(sender As Object, e As RoutedEventArgs) Handles BtnFermer.Click
        RaiseEvent CloseMe(Me)
    End Sub

    Private Sub BtnCharger1_Click(sender As Object, e As RoutedEventArgs) Handles BtnCharger1.Click
        LireModeChauffage(_listSemaineCharger, "Charger")
    End Sub

    Private Sub BtnNormal1_Click(sender As Object, e As RoutedEventArgs) Handles BtnNormal1.Click
        LireModeChauffage(_listSemaineNormal, "Normal")
    End Sub

    Private Sub BtnConger1_Click(sender As Object, e As RoutedEventArgs) Handles BtnConger1.Click
        LireModeChauffage(_listSemaineConger, "Conger")
    End Sub

    Private Sub BtnAbsence1_Click(sender As Object, e As RoutedEventArgs) Handles BtnAbsence1.Click
        LireModeChauffage(_listSemaineAbsence, "Absence")
    End Sub

    Private Sub BtnReduit1_Click(sender As Object, e As RoutedEventArgs) Handles BtnReduit1.Click
        LireModeChauffage(_listSemaineReduit, "Reduit")
    End Sub

    Private Sub BtnCalendrier1_Click(sender As Object, e As RoutedEventArgs) Handles BtnCalendrier1.Click
        LireCalendrier()
    End Sub

    Protected Overrides Sub Finalize()

        _Driver.Parametres.Item(5).Valeur = IsActiveLecture
        MyBase.Finalize()
        _listSemaineConger.Clear()
        _listSemaineNormal.Clear()
        _listSemaineReduit.Clear()
        _listSemaineCharger.Clear()
        _listSemaineAbsence.Clear()
        _listCalendar.Clear()

    End Sub


    Private Sub SauverCopiefichier()

        Dim zipFile As FileInfo
        Dim NomFichier As String

        'Sauver une copie du fichier Excel "Calendrier.bak"
        NomFichier = _Driver.Parametres.Item(1).Valeur
        Dim ParaFichier = Split(NomFichier, ".")
        ParaFichier(1) = "Bak"
        zipFile = New FileInfo(ParaFichier(0) + "." + ParaFichier(1))
        If (zipFile.Exists) Then
            zipFile.Delete()
        End If
        newFile.CopyTo(zipFile.FullName)
    End Sub



    Private Sub Sauverfichier()

        SauverCalendrier(_listCalendar, "Calendrier")
        SauverSemaine(_listSemaineConger, "Conger")
        SauverSemaine(_listSemaineNormal, "Normal")
        SauverSemaine(_listSemaineReduit, "Reduit")
        SauverSemaine(_listSemaineCharger, "Charger")
        SauverSemaine(_listSemaineAbsence, "Absence")

        pck.Save()

    End Sub


    Private Sub SauverCalendrier(ByRef ListCalendrier As List(Of Calendrier), ByVal Mode As String)

        Dim Ligne As Integer
        Dim worksheet As ExcelWorksheet
        Dim Donnee As Calendrier

        worksheet = pck.Workbook.Worksheets.Item(Mode)
        For Ligne = 0 To 51
            Donnee = ListCalendrier.Item(Ligne)
            worksheet.Cells(Ligne + 2, 2).Value = Donnee.Mode
        Next

    End Sub


    Private Sub SauverSemaine(ByRef ListSemaine As List(Of Semaine), ByVal Mode As String)

        Dim Ligne As Integer
        Dim worksheet As ExcelWorksheet
        Dim Donnee As Semaine

        worksheet = pck.Workbook.Worksheets.Item(Mode)
        For Ligne = 0 To 48
            Donnee = ListSemaine.Item(Ligne)

            worksheet.Cells(Ligne + 2, 2).Value = Donnee.Lundi
            worksheet.Cells(Ligne + 2, 3).Value = Donnee.Mardi
            worksheet.Cells(Ligne + 2, 4).Value = Donnee.Mercredi
            worksheet.Cells(Ligne + 2, 5).Value = Donnee.Jeudi
            worksheet.Cells(Ligne + 2, 6).Value = Donnee.Vendredi
            worksheet.Cells(Ligne + 2, 7).Value = Donnee.Samedi
            worksheet.Cells(Ligne + 2, 8).Value = Donnee.Dimanche
        Next

    End Sub



    Private Sub DataGrid1_BeginningEdit(sender As Object, e As DataGridBeginningEditEventArgs) Handles DataGrid1.BeginningEdit

        Dim worksheet As ExcelWorksheet

        'Sauver la valeur avant la modifier       
        _Colonne = DataGrid1.CurrentCell.Column.DisplayIndex    'Colonne
        _Ligne = DataGrid1.SelectedIndex      'Ligne   
        worksheet = pck.Workbook.Worksheets.Item(_Mode)
        _IsModifier = True
        _Valeur = worksheet.Cells(_Ligne + 2, _Colonne + 1).Value

    End Sub


    Private Sub ModifierSemaine(ByRef ListSemaine As List(Of Semaine))

        Dim Cellule As String = ""
        Dim Donnee As Semaine
        Dim Ligne As Single
        Dim Colog As String = ""

        If _Valeur <> "" Then
            Donnee = ListSemaine.Item(_Ligne)
            Select Case _Colonne
                Case 1
                    Cellule = Donnee.Lundi
                    Colog = "Lundi"
                Case 2
                    Cellule = Donnee.Mardi
                    Colog = "Mardi"
                Case 3
                    Cellule = Donnee.Mercredi
                    Colog = "Mercredi"
                Case 4
                    Cellule = Donnee.Jeudi
                    Colog = "Jeudi"
                Case 5
                    Cellule = Donnee.Vendredi
                    Colog = "Venderdi"
                Case 6
                    Cellule = Donnee.Samedi
                    Colog = "Samedi"
                Case 7
                    Cellule = Donnee.Dimanche
                    Colog = "Dimanche"
            End Select
            If Cellule = "ECC" Or Cellule = "EC" Or Cellule = "PRE" Then
                'Bonne valeur
            Else
                Ligne = _Ligne / 2
                MessageBox.Show("Valeurs erronée : " & "'" & Cellule & "'" & " à la Ligne : " & CStr(Ligne) & "h, Colonne : " & Colog, "Erreur de saisie", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Select Case _Colonne
                    Case 1
                        Donnee.Lundi = _Valeur
                    Case 2
                        Donnee.Mardi = _Valeur
                    Case 3
                        Donnee.Mercredi = _Valeur
                    Case 4
                        Donnee.Jeudi = _Valeur
                    Case 5
                        Donnee.Vendredi = _Valeur
                    Case 6
                        Donnee.Samedi = _Valeur
                    Case 7
                        Donnee.Dimanche = _Valeur
                End Select
                DataGrid1.Items.Refresh()
                DataGridOk(Me, Nothing)
                DataGrid1.UnselectAllCells()
            End If
            _Valeur = ""
        End If

    End Sub


    Private Sub ModifierCalendrier()

        Dim Cellule As String = ""
        Dim Donnee As Calendrier

        If _Valeur <> "" Then
            Donnee = _listCalendar.Item(_Ligne)
            Select Case _Colonne
                Case 1
                    Cellule = Donnee.Mode
                    If Cellule = "Conger" Or Cellule = "Normal" Or Cellule = "Reduit" Or Cellule = "Charger" Or Cellule = "Absence" Then

                    Else
                        MessageBox.Show("Valeurs erronée : " & "'" & Cellule & "'" & " à la Ligne : " & CStr(_Ligne + 1), "Erreur de saisie", MessageBoxButtons.OK, MessageBoxIcon.Error)
                        Donnee.Mode = _Valeur
                        DataGrid1.Items.Refresh()
                        DataGridOk(Me, Nothing)
                        DataGrid1.UnselectAllCells()
                    End If
            End Select
            _Valeur = ""
        End If

    End Sub



    Private Sub DataGrid1_Loaded(sender As Object, e As RoutedEventArgs) Handles DataGrid1.Loaded

        'Preparer les pages  
        If FichierOk = True Then
            PrepareListSemaine(_listSemaineConger, "Conger")
            PrepareListSemaine(_listSemaineNormal, "Normal")
            PrepareListSemaine(_listSemaineReduit, "Reduit")
            PrepareListSemaine(_listSemaineCharger, "Charger")
            PrepareListSemaine(_listSemaineAbsence, "Absence")
            PrepareCalendrier("Calendrier")
        End If

    End Sub

    Private Sub PrepareCalendrier(ByVal Mode As String)

        Dim worksheet As ExcelWorksheet
        Dim Text1, Text2 As String

        _Mode = Mode
        worksheet = pck.Workbook.Worksheets.Item(_Mode)
        _listCalendar.Clear()
        For Ligne = 2 To 53
            Text1 = worksheet.Cells(Ligne, 1).Value
            Text2 = worksheet.Cells(Ligne, 2).Value
            _listCalendar.Add(New Calendrier() With {.Semaine = Text1, .Mode = Text2})
        Next
        DataGrid1.ItemsSource = _listCalendar

    End Sub


    Private Sub PrepareListSemaine(ByRef ListSemaine As List(Of Semaine), ByVal Mode As String)
        Dim worksheet As ExcelWorksheet

        _Mode = Mode
        worksheet = pck.Workbook.Worksheets.Item(_Mode)
        ListSemaine.Clear()
        Dim Text1, Text2, Text3, Text4, Text5, Text6, Text7, Text8 As String
        For Ligne = 2 To 50
            Text1 = worksheet.Cells(Ligne, 1).Value
            Text2 = worksheet.Cells(Ligne, 2).Value
            Text3 = worksheet.Cells(Ligne, 3).Value
            Text4 = worksheet.Cells(Ligne, 4).Value
            Text5 = worksheet.Cells(Ligne, 5).Value
            Text6 = worksheet.Cells(Ligne, 6).Value
            Text7 = worksheet.Cells(Ligne, 7).Value
            Text8 = worksheet.Cells(Ligne, 8).Value
            ListSemaine.Add(New Semaine() With {.Heure = Text1, .Lundi = Text2, .Mardi = Text3, .Mercredi = Text4, .Jeudi = Text5, .Vendredi = Text6, .Samedi = Text7, .Dimanche = Text8})
        Next
        DataGrid1.ItemsSource = ListSemaine

    End Sub


    Private Sub RemplirFond_Checked(sender As Object, e As RoutedEventArgs) Handles RemplirFond.Checked

        Remplir_Fond = True
        DataGridOk(Me, Nothing)

    End Sub


    Private Sub RemplirFond_Unchecked(sender As Object, e As RoutedEventArgs) Handles RemplirFond.Unchecked

        Remplir_Fond = False
        DataGridOk(Me, Nothing)

    End Sub



    Private Sub DataGrid1_SelectedCellsChanged(sender As Object, e As SelectedCellsChangedEventArgs) Handles DataGrid1.SelectedCellsChanged
        If _Mode = "Calendrier" Then
            ModifierCalendrier()
        Else
            Select Case _Mode
                Case "Conger"
                    ModifierSemaine(_listSemaineConger)
                Case "Normal"
                    ModifierSemaine(_listSemaineNormal)
                Case "Reduit"
                    ModifierSemaine(_listSemaineReduit)
                Case "Charger"
                    ModifierSemaine(_listSemaineCharger)
                Case "Absence"
                    ModifierSemaine(_listSemaineAbsence)
            End Select
        End If
    End Sub
End Class
