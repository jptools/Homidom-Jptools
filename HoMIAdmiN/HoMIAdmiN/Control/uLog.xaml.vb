﻿Imports System.Data
Imports System.IO
Imports System.Collections.ObjectModel

Partial Public Class uLog
    Public Event CloseMe(ByVal MyObject As Object)
    Public ligneLog As New ObservableCollection(Of Dictionary(Of String, Object))
    Dim keys As New List(Of String)
    Dim headers As String() = {"datetime", "typesource", "source", "fonction", "message"}

    Private Sub BtnRefresh_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnRefresh.Click
        RefreshLog()
    End Sub

    Private Sub RefreshLog()
        Try
            'Variables
            Dim _LigneIgnorees As Integer = 0
            Me.Cursor = Cursors.Wait

            ligneLog.Clear()

            If IsConnect = True Then
                Dim MyRepAppData As String = System.Environment.GetFolderPath(System.Environment.SpecialFolder.CommonApplicationData) & "\HoMIAdmiN"
                Dim TargetFile As StreamWriter
                TargetFile = New StreamWriter(MyRepAppData & "\log.txt", False)
                TargetFile.Write(myService.ReturnLog)
                TargetFile.Close()

                Dim tr As TextReader = New StreamReader(MyRepAppData & "\log.txt")
                'Dim lineCount As Integer = 1

                'lecture de la premiere ligne souvent incomplete ou avec un message du serveur
                Dim line As String = tr.ReadLine()
                If line <> "" Then
                    Dim tmp As String() = line.Trim.Split(vbTab)
                    If tmp.Length > 3 Then
                        If tmp(3) = "ReturnLog" Then
                            MessageBox.Show(tmp(4), "Message du serveur", MessageBoxButton.OK, MessageBoxImage.Information)
                            line = tr.ReadLine()
                            line = tr.ReadLine()
                        End If
                    End If
                End If

                While tr.Peek() >= 0
                    Try
                        line = tr.ReadLine()

                        If line <> "" Then
                            Dim tmp As String() = line.Trim.Split(vbTab)

                            If tmp.Length < 6 And tmp.Length > 3 Then
                                If tmp(4).Length > 255 Then tmp(4) = Mid(tmp(4), 1, 255)
                                Dim sensorData As New Dictionary(Of String, Object) ' creates a dictionary where column name is the key and data is the value
                                For i As Integer = 0 To tmp.Length - 1
                                    sensorData(keys(i)) = tmp(i)
                                Next
                                'ligneLog.Add(sensorData)
                                ligneLog.Insert(0, sensorData)
                                sensorData = Nothing
                            Else
                                'ligne au format incorrect 
                                Dim sensorData As New Dictionary(Of String, Object) ' creates a dictionary where column name is the key and data is the value
                                sensorData(keys(0)) = ""
                                sensorData(keys(1)) = ""
                                sensorData(keys(2)) = ""
                                sensorData(keys(3)) = ""
                                sensorData(keys(4)) = line.Trim.ToString
                                ligneLog.Insert(0, sensorData)
                                sensorData = Nothing
                            End If
                            'lineCount += 1
                        End If
                    Catch ex As Exception
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la ligne du fichier log: " & ex.ToString, "ERREUR", "")
                    End Try
                End While
                tr.Close()

                Try
                    If File.Exists(MyRepAppData & "\log.txt") Then
                        File.Delete(MyRepAppData & "\log.txt")
                    End If
                Catch ex As Exception
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la suppression du fichier log temporaire: " & ex.ToString, "ERREUR", "")
                End Try
            End If

            Try
                'If _LigneIgnorees > 0 Then
                '    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, _LigneIgnorees & " ligne(s) du log ne seront pas prises en compte car elles ne respectent pas le format attendu, veuillez consultez le fichier log sur le serveur pour avoir la totalité", "INFO", "")
                'End If

                DGW.DataContext = ligneLog
            Catch ex As Exception
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & ex.ToString)
            End Try


            Me.Cursor = Nothing
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la récuppération du fichier log: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub CreateGridColumn(ByVal headerText As String)
        Try
            Dim col1 As DataGridTextColumn = New DataGridTextColumn()
            col1.Header = headerText
            col1.Binding = New Binding(String.Format("[{0}]", headerText))
            DGW.Columns.Add(col1)
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur Sub CreateGridColumn: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Public Sub New()
        Try

            ' Cet appel est requis par le Concepteur Windows Form.
            InitializeComponent()

            ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
            For Each h As String In headers
                CreateGridColumn(h)
                keys.Add(h)
            Next

            RefreshLog()
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors sur la fonction New de uLog: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub BtnClose_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnClose.Click
        RaiseEvent CloseMe(Me)
    End Sub

    Private Sub DGW_LoadingRow(ByVal sender As Object, ByVal e As System.Windows.Controls.DataGridRowEventArgs) Handles DGW.LoadingRow
        Try
            Dim RowDataContaxt As Dictionary(Of String, Object) = TryCast(e.Row.DataContext, Dictionary(Of String, Object))
            If RowDataContaxt IsNot Nothing Then
                Select Case RowDataContaxt(keys(1))
                    Case "INFO"
                        e.Row.Background = Brushes.White
                    Case "ACTION"
                        e.Row.Background = Brushes.White
                    Case "MESSAGE"
                        e.Row.Background = Brushes.White
                    Case "VALEUR CHANGE"
                        e.Row.Background = Brushes.White
                    Case "VALEUR INCHANGE"
                        e.Row.Background = Brushes.White
                    Case "VALEUR INCHANGE PRECISION"
                        e.Row.Background = Brushes.White
                    Case "VALEUR INCHANGE LASTETAT"
                        e.Row.Background = Brushes.White
                    Case "ERREUR"
                        e.Row.Background = Brushes.Red
                    Case "ERREUR CRITIQUE"
                        e.Row.Background = Brushes.Red
                    Case "DEBUG"
                        e.Row.Background = Brushes.Yellow
                    Case Else
                        e.Row.Background = Brushes.White
                End Select
            End If

        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur DGW_LoadingRow: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub ExportLog()
        Try
            ' Configure open file dialog box
            Dim dlg As New Microsoft.Win32.SaveFileDialog()
            dlg.FileName = "Log" ' Default file name
            dlg.DefaultExt = ".txt" ' Default file extension
            dlg.Filter = "Fichier log (.txt)|*.txt" ' Filter files by extension
            ' Show open file dialog box
            Dim result As Boolean = dlg.ShowDialog()
            ' Process open file dialog box results
            If result = True Then
                ' Open document
                Dim filename As String = dlg.FileName
                Dim retour As String = myService.ReturnLog
                If retour.StartsWith("ERREUR") Then
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, retour, "Erreur ReturnLog", "")
                Else
                    Dim TargetFile As StreamWriter
                    TargetFile = New StreamWriter(filename, False)
                    TargetFile.Write(retour)
                    TargetFile.Close()
                End If
            End If
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur ExportLog: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub

    Private Sub BtnExportLog_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnExportLog.Click
        ExportLog()
    End Sub

    Private Sub BtnRefreshLocal_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnRefreshLocal.Click
        Try
            'Variables
            Dim _LigneIgnorees As Integer = 0
            Me.Cursor = Cursors.Wait

            If System.IO.File.Exists(My.Application.Info.DirectoryPath & "\logs\log_" & DateAndTime.Now.ToString("yyyyMMdd") & ".txt") Then
                ligneLog.Clear()

                If IsConnect = True Then
                    Dim tr As TextReader = New StreamReader(My.Application.Info.DirectoryPath & "\logs\log_" & DateAndTime.Now.ToString("yyyyMMdd") & ".txt")
                    'Dim lineCount As Integer = 1

                    Dim line As String
                    While tr.Peek() >= 0
                        Try
                            line = tr.ReadLine()

                            If line <> "" Then
                                Dim tmp As String() = line.Trim.Split(vbTab)

                                If tmp.Length < 6 And tmp.Length > 3 Then
                                    If tmp(4).Length > 255 Then tmp(4) = Mid(tmp(4), 1, 255)
                                    Dim sensorData As New Dictionary(Of String, Object) ' creates a dictionary where column name is the key and data is the value
                                    For i As Integer = 0 To tmp.Length - 1
                                        sensorData(keys(i)) = tmp(i)
                                    Next
                                    'ligneLog.Add(sensorData)
                                    ligneLog.Insert(0, sensorData)
                                    sensorData = Nothing
                                Else
                                    _LigneIgnorees += 1

                                    ''test david ligne ignorée contenu
                                    'Dim strintemp As String = ""
                                    'For i As Integer = 0 To tmp.Length - 1
                                    '    strintemp = strintemp & "---" & tmp(i)
                                    'Next
                                    'MessageBox.Show("ligne ignorée: " & strintemp, "Test David", MessageBoxButton.OK, MessageBoxImage.Exclamation)

                                End If
                                'lineCount += 1
                            End If
                        Catch ex As Exception
                            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la ligne du fichier log: " & ex.ToString, "ERREUR", "")
                        End Try
                    End While
                    tr.Close()
                End If

                Try
                    If _LigneIgnorees > 0 Then
                        AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.INFO, _LigneIgnorees & " ligne(s) du log ne seront pas prises en compte car elles ne respectent pas le format attendu, veuillez consultez le fichier log sur le serveur pour avoir la totalité", "INFO", "")
                    End If

                    DGW.DataContext = ligneLog
                Catch ex As Exception
                    AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur: " & ex.ToString)
                End Try


                Me.Cursor = Nothing
            Else
                AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Fichier log local non trouvé : " & My.Application.Info.DirectoryPath & "\logs\log_" & DateAndTime.Now.ToString("yyyyMMdd") & ".txt")
            End If
            
        Catch ex As Exception
            AfficheMessageAndLog(HoMIDom.HoMIDom.Server.TypeLog.ERREUR, "Erreur lors de la récuppération du fichier log: " & ex.ToString, "ERREUR", "")
        End Try
    End Sub
End Class
