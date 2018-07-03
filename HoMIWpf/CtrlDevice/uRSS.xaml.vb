Imports System.Net
Imports System.Xml
Imports System.Xml.XmlReader
Imports System.IO
Imports System.ServiceModel.Syndication
Imports System.Threading
Imports System.ComponentModel
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.Windows.Threading

Public Class uRSS
    Dim _Uri As String
    Dim dt As DispatcherTimer
    Private _SizeStatus As Double = 10
    Private _CouleurStatus As Windows.Media.Brush = Nothing
    Private WithEvents RssReader As RSS.RssReader = Nothing
    Private RssFeed As RSS.RssFeed = Nothing

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
        Catch ex As Exception

        End Try
    End Sub

    Public Sub New()

        ' Cet appel est requis par le concepteur.
        InitializeComponent()
        GetCodeHtml()
        RssReader = New RSS.RssReader

        'Lancer le Timer_Tick
        dt = New DispatcherTimer()
        AddHandler dt.Tick, AddressOf dispatcherTimer_Tick
        dt.Interval = New TimeSpan(0, 30, 0) ' 30 minutes
        dt.Start()
    End Sub

    Private Sub URSS_Unloaded(ByVal sender As Object, ByVal e As System.Windows.RoutedEventArgs) Handles Me.Unloaded
        Try
            If dt IsNot Nothing Then
                dt.Stop()
                RemoveHandler dt.Tick, AddressOf dispatcherTimer_Tick
                dt = Nothing
            End If
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur uRSS_Unloaded: " & ex.ToString, "Erreur")
        End Try
    End Sub

    Public Sub dispatcherTimer_Tick(ByVal sender As Object, ByVal e As EventArgs)
        RefreshChannel()
    End Sub

    Public Property URIRss As String
        Get
            Return _Uri
        End Get
        Set(ByVal value As String)
            _Uri = value
            If String.IsNullOrEmpty(_Uri) = False Then RefreshChannel()
        End Set
    End Property

    Public Property SizeStatus As Double
        Get
            Return _SizeStatus
        End Get
        Set(value As Double)
            _SizeStatus = value
        End Set
    End Property

    Public Property CouleurStatus As Windows.Media.Brush
        Get
            Return _CouleurStatus
        End Get
        Set(value As Windows.Media.Brush)
            _CouleurStatus = value
        End Set
    End Property

    Private Sub Charger_Click(sender As Object, e As RoutedEventArgs) Handles Charger.Click
        Try
            If String.IsNullOrEmpty(_Uri) = False Then
                RefreshChannel()
            End If

        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur Charger_Click: " & ex.ToString, "Erreur")
        End Try
    End Sub

    Private Sub RefreshChannel()
        Try
            'Si internet n'est pas disponible on ne mets pas à jour les informations
            If My.Computer.Network.IsAvailable = False Then
                'Log(FctLog.TypeLog.ERREUR, "Erreur RSS -> RefreshChannel: ", "Pas de connexion réseau, lecture des infos impossible", "")
                Exit Sub
            End If
            LstRssItems.Items.Clear()
            LstRssItems.FontSize = SizeStatus
            LstRssItems.Foreground = CouleurStatus
            LstRssItems.Background = Background

            ' Téléchargement du flux           
            'RssReader.Load(New Uri("http://lemonde.fr/rss/une.xml"))
            'RssReader.Load(New Uri("https://france3-regions.francetvinfo.fr/grand-est/haut-rhin/rss"))
            RssReader.Load(New Uri(URIRss))
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur RefreshChannel: " & ex.ToString, "Erreur")
        End Try
    End Sub

    Private Sub Reader_Completed(ByVal Sender As Object, ByVal e As RSS.RssReader.CompletedEventArgs) Handles RssReader.Completed
        RssFeed = e.Result
        ShowInfos()
    End Sub


#Region " Infos "

    Private Sub ShowInfos()
        Dim Temp As String = ""
        Dim NbCaract As Integer = 0
        Dim Temp1 As String = ""
        Dim _NbCARACT As Integer = (Width / ((LstRssItems.FontSize) / 2)) - 3
        Dim Separat As New String("-", _NbCARACT + 17)
        Dim NbInfos As Integer = 0

        Try
            If RssFeed IsNot Nothing Then
                PicImage.Visibility = Windows.Visibility.Hidden
                If Not String.IsNullOrEmpty(RssFeed.ImageUrl) Then
                    PicImage.Source = New BitmapImage(New Uri(RssFeed.ImageUrl))
                End If
                If PicImage.IsLoaded = True Then
                    PicImage.Visibility = Windows.Visibility.Visible
                End If
                LabelTitre.Text = RssFeed.Title
                For Each item As RSS.RssItem In RssFeed.Items

                    ' LstRssItems.Items.Add(item.Title)
                    Temp = item.Title
                    Temp = Replace(Temp, vbCrLf, "")
                    Do
                        NbCaract = InStr(_NbCARACT, Temp, " ")
                        If NbCaract = 0 Then
                            LstRssItems.Items.Add(Temp)
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
                                LstRssItems.Items.Add(Temp1)
                                Temp = Mid(Temp, NbCaract + 1, Len(Temp))
                            End If
                        End If
                    Loop Until Len(Temp) = 0

                    'LstRssItems.Items.Add(item.Description)
                    Temp = item.Description
                    Temp = Remplace_html_entities(Temp)
                    Temp = Replace(Temp, vbCrLf, "")
                    Do
                        NbCaract = InStr(_NbCARACT, Temp, " ")
                        If NbCaract = 0 Then
                            LstRssItems.Items.Add(Temp)
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
                                LstRssItems.Items.Add(Temp1)
                                Temp = Mid(Temp, NbCaract + 1, Len(Temp))
                            End If
                        End If
                    Loop Until Len(Temp) = 0
                    LstRssItems.Items.Add(NbInfos.ToString & ": Lire la suite ->")  'item.Link
                    LstRssItems.Items.Add(Separat)
                    NbInfos = NbInfos + 1
                Next
            End If
        Catch ex As Exception
            AfficheMessageAndLog(FctLog.TypeLog.ERREUR, "Erreur ShowInfos: " & ex.ToString, "Erreur")
        End Try
    End Sub

    Public Function Remplace_html_entities(ByVal e_txt As String) As String
        Try
            'recherche les entités via une expression réguliere
            Dim myregexp As New Regex("&[#a-zA-Z0-9]{2,6}\;", RegexOptions.IgnoreCase)
            If myregexp.IsMatch(e_txt) Then
                Dim elt As Match
                For Each elt In myregexp.Matches(e_txt)
                    If NametoNum.ContainsKey(elt.Value) Then
                        'remplace l'occurence par la valeur associée                        
                        e_txt = Replace(e_txt, elt.Value, Chr(NametoNum(elt.Value)))
                    End If
                Next
            End If
            Return e_txt
        Catch ex As Exception
            Return e_txt
        End Try
    End Function

    Private Sub DClickPage(sender As Object, e As MouseButtonEventArgs) Handles LstRssItems.MouseDoubleClick
        If e.Source.SelectedIndex > -1 Then
            Dim Debut As Integer = InStr(e.Source.SelectedValue, "Lire la suite ->")
            If Debut Then
                Dim val As String = Left(e.Source.SelectedValue, Debut - 3)
                Dim item As RSS.RssItem = RssFeed.Items(CByte(val))
                If Not String.IsNullOrEmpty(item.Link) Then
                    'Lancer internet avec l'adresse
                    Process.Start(item.Link)
                End If

            End If
        End If
    End Sub


#End Region

End Class


Namespace RSS

    ''' <summary>
    ''' Classe permettant de télécharger un flux RSS de façon asynchrone
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RssReader

        ''' <summary>
        ''' Permet de connaitre l'avance du téléchargement du flux RSS
        ''' </summary>
        ''' <param name="Sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Event ProgressChanged(ByVal Sender As Object, ByVal e As System.Net.DownloadProgressChangedEventArgs)
        ''' <summary>
        ''' Signal la fin du téléchargement du flux
        ''' </summary>
        ''' <param name="Sender"></param>
        ''' <param name="e"></param>
        ''' <remarks></remarks>
        Public Event Completed(ByVal Sender As Object, ByVal e As CompletedEventArgs)

        ''' <summary>
        ''' Télécharge un flux RSS à partir de l'Uri spécifiée
        ''' </summary>
        ''' <param name="Uri">Adresse du flux</param>
        ''' <remarks></remarks>
        Public Sub Load(ByVal Uri As Uri)
            Me.Load(Uri, Nothing)
        End Sub

        ''' <summary>
        ''' Télécharge un flux RSS à partir de l'Uri spécifiée
        ''' </summary>
        ''' <param name="Uri">Adresse du flux</param>
        ''' <param name="UserState">Objet défini par l'utilisateur retourné à la fin de l'opération asynchrone </param>
        ''' <remarks></remarks>
        Public Sub Load(ByVal Uri As Uri, ByVal UserState As Object)
            If Client IsNot Nothing Then Cancel()
            Client = New Net.WebClient
            _UserState = UserState
            Client.DownloadDataAsync(Uri, UserState)
        End Sub

        Private _UserState As Object
        ''' <summary>
        ''' Annule le téléchargement du flux
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Cancel()
            Client.Dispose()
            Client = Nothing
            Dim args As New CompletedEventArgs(Nothing, Nothing, True, _UserState)
            RaiseEvent Completed(Me, args)
        End Sub

        ''' <summary>
        ''' Détermine si un téléchargement est en cours
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property IsBusy() As Boolean
            Get
                If Client Is Nothing Then
                    Return False
                Else
                    Return Client.IsBusy
                End If
            End Get
        End Property

        Private WithEvents Client As Net.WebClient

        Private Sub Client_DownloadProgressChanged(ByVal sender As Object, ByVal e As System.Net.DownloadProgressChangedEventArgs) Handles Client.DownloadProgressChanged
            RaiseEvent ProgressChanged(Me, e)
        End Sub

        Private Sub Client_DownloadDataCompleted(ByVal sender As Object, ByVal e As System.Net.DownloadDataCompletedEventArgs) Handles Client.DownloadDataCompleted
            Client = Nothing
            Dim res As RssFeed = Nothing
            If e.Error Is Nothing And e.Cancelled = False Then
                ' Créer le RssFeed
                res = New RssFeed(e.Result)
            End If
            Dim args As New CompletedEventArgs(res, e.Error, e.Cancelled, e.UserState)
            RaiseEvent Completed(Me, args)
        End Sub

        Public Class CompletedEventArgs
            Inherits EventArgs

            Friend Sub New(ByVal Result As RssFeed, ByVal [Error] As Exception, ByVal Canceled As Boolean, ByVal UserState As Object)
                _Result = Result
                _Error = [Error]
                _Canceled = Canceled
                _UserState = UserState
            End Sub

            Public ReadOnly Property Result() As RssFeed
                Get
                    Return _Result
                End Get
            End Property

            Private ReadOnly _Result As RssFeed

            Public ReadOnly Property [Error]() As System.Exception
                Get
                    Return _Error
                End Get
            End Property

            Private ReadOnly _Error As Exception

            Public ReadOnly Property UserState() As Object
                Get
                    Return _UserState
                End Get
            End Property

            Private ReadOnly _UserState As Exception

            Public ReadOnly Property Canceled() As Boolean
                Get
                    Return _Canceled
                End Get
            End Property

            Private ReadOnly _Canceled As Boolean

        End Class


    End Class

    ''' <summary>
    ''' Représente un flux RSS
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RssFeed


#Region " Contructeurs "

        Public Sub New(ByVal XmlDocument As Xml.XmlDocument)
            LoadXml(XmlDocument)
        End Sub

        Public Sub New(ByVal Stream As IO.Stream)
            Dim doc As New XmlDocument
            doc.Load(Stream)
            LoadXml(doc)
        End Sub

        Public Sub New(ByVal Datas() As Byte)
            Dim Stream As New IO.MemoryStream(Datas)
            Dim doc As New XmlDocument
            doc.Load(Stream)
            LoadXml(doc)
        End Sub

        Private Sub LoadXml(ByVal doc As Xml.XmlDocument)
            tags.Clear()
            If doc.DocumentElement.Name.ToLower = "rss" Then
                Dim version As String = doc.DocumentElement.GetAttribute("version")
                tags.Add(New RssTag("version", version))

                If version = "2.0" Then
                    Dim channel As XmlElement = doc.DocumentElement.GetElementsByTagName("channel")(0)
                    For Each elem As XmlElement In channel
                        If elem.Name.ToLower = "item" Then
                            ' Un item
                            _Items.Add(New RssItem(elem))
                        Else
                            ' Un tag du channel
                            LoadElem("", elem)
                        End If
                    Next
                End If
            End If
        End Sub

        Private Sub LoadElem(ByVal BaseName As String, ByVal Elem As XmlElement)
            For Each attr As XmlNode In Elem.Attributes
                tags.Add(New RssTag(BaseName & attr.Name, attr.InnerText))
            Next
            For Each subElem As XmlNode In Elem.ChildNodes
                If subElem.NodeType = XmlNodeType.Element Then
                    LoadElem(BaseName & Elem.Name & ".", subElem)
                Else
                    tags.Add(New RssTag(BaseName & Elem.Name, subElem.InnerText))
                End If
            Next
        End Sub

#End Region

#Region " Propriétés "

        ' All Tags
        Private tags As New List(Of RssTag)
        Private ReadOnly _Elements As New RssTagsCollection(tags)
        ''' <summary>
        ''' Liste de tous les éléments d'un flux. Accessible par index ou par nom
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Elements() As RssTagsCollection
            Get
                Return _Elements
            End Get
        End Property

        ' Items
        Private _Items As New List(Of RssItem)
        Private ReadOnly _RoItems As New RssItemsCollection(_Items)
        ''' <summary>
        ''' Liste des items contenus dans ce flux.
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Items() As RssItemsCollection
            Get
                Return _RoItems
            End Get
        End Property

        ' Eléments principaux
        Public ReadOnly Property Title() As String
            Get
                Return Elements.GetValueByName("title")
            End Get
        End Property
        Public ReadOnly Property Link() As String
            Get
                Return Elements.GetValueByName("link")
            End Get
        End Property
        Public ReadOnly Property Description() As String
            Get
                Return Elements.GetValueByName("description")
            End Get
        End Property
        Public ReadOnly Property Language() As String
            Get
                Return Elements.GetValueByName("language")
            End Get
        End Property
        Public ReadOnly Property Copyright() As String
            Get
                Return Elements.GetValueByName("copyright")
            End Get
        End Property
        Public ReadOnly Property ManagingEditor() As String
            Get
                Return Elements.GetValueByName("managingEditor")
            End Get
        End Property
        Public ReadOnly Property WebMaster() As String
            Get
                Return Elements.GetValueByName("webMaster")
            End Get
        End Property
        Public ReadOnly Property PubDate() As String
            Get
                Return Elements.GetValueByName("pubDate")
            End Get
        End Property
        Public ReadOnly Property LastBuildDate() As String
            Get
                Return Elements.GetValueByName("lastBuildDate")
            End Get
        End Property
        Public ReadOnly Property Category() As String
            Get
                Return Elements.GetValueByName("category")
            End Get
        End Property
        Public ReadOnly Property Generator() As String
            Get
                Return Elements.GetValueByName("generator")
            End Get
        End Property
        Public ReadOnly Property Docs() As String
            Get
                Return Elements.GetValueByName("docs")
            End Get
        End Property
        Public ReadOnly Property Ttl() As String
            Get
                Return Elements.GetValueByName("ttl", "0")
            End Get
        End Property
        Public ReadOnly Property Rating() As String
            Get
                Return Elements.GetValueByName("rating")
            End Get
        End Property
        Public ReadOnly Property SkipHours() As String
            Get
                Return Elements.GetValueByName("skipHours")
            End Get
        End Property
        Public ReadOnly Property SkipDays() As String
            Get
                Return Elements.GetValueByName("skipDays")
            End Get
        End Property

        Public ReadOnly Property ImageTitle() As String
            Get
                Return Elements.GetValueByName("image.title")
            End Get
        End Property
        Public ReadOnly Property ImageUrl() As String
            Get
                Return Elements.GetValueByName("image.url")
            End Get
        End Property
        Public ReadOnly Property ImageLink() As String
            Get
                Return Elements.GetValueByName("image.link")
            End Get
        End Property
        Public ReadOnly Property ImageWidth() As Integer
            Get
                Return Elements.GetValueByName("image.width", "0")
            End Get
        End Property
        Public ReadOnly Property ImageHeight() As Integer
            Get
                Return Elements.GetValueByName("image.height", "0")
            End Get
        End Property
        Public ReadOnly Property ImageDescription() As String
            Get
                Return Elements.GetValueByName("image.description")
            End Get
        End Property

#End Region

    End Class

    ''' <summary>
    ''' Item contenu dans un flux RSS
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RssItem

#Region " Contructeur "

        Public Sub New(ByVal XmlElement As Xml.XmlElement)
            For Each subElem As XmlNode In XmlElement.ChildNodes
                LoadElem("", subElem)
            Next
        End Sub


        Private Sub LoadElem(ByVal BaseName As String, ByVal Elem As XmlElement)
            For Each attr As XmlNode In Elem.Attributes
                tags.Add(New RssTag(BaseName & attr.Name, attr.InnerText))
            Next
            For Each subElem As XmlNode In Elem.ChildNodes
                If subElem.NodeType = XmlNodeType.Element Then
                    LoadElem(BaseName & Elem.Name & ".", subElem)
                Else
                    tags.Add(New RssTag(BaseName & Elem.Name, subElem.InnerText))
                End If
            Next
        End Sub


#End Region

#Region " Propriétés "

        Private tags As New List(Of RssTag)
        Private _Elements As New RssTagsCollection(tags)
        ''' <summary>
        ''' Liste de tous les éléments d'un item. Accessible par index ou par nom
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property Elements() As RssTagsCollection
            Get
                Return _Elements
            End Get
        End Property

        Public ReadOnly Property Title() As String
            Get
                Return _Elements.GetValueByName("title")
            End Get
        End Property

        Public ReadOnly Property Description() As String
            Get
                Return Elements.GetValueByName("description")
            End Get
        End Property

        Public ReadOnly Property Link() As String
            Get
                Return _Elements.GetValueByName("link")
            End Get
        End Property

        Public ReadOnly Property Author() As String
            Get
                Return Elements.GetValueByName("author")
            End Get
        End Property

        Public ReadOnly Property Category() As String
            Get
                Return Elements.GetValueByName("category")
            End Get
        End Property

        Public ReadOnly Property Comments() As String
            Get
                Return Elements.GetValueByName("comments")
            End Get
        End Property

        Public ReadOnly Property Enclosure() As String
            Get
                Return Elements.GetValueByName("enclosure")
            End Get
        End Property

        Public ReadOnly Property Guid() As String
            Get
                Return Elements.GetValueByName("guid")
            End Get
        End Property

        Public ReadOnly Property PubDate() As String
            Get
                Return Elements.GetValueByName("pubDate")
            End Get
        End Property

        Public ReadOnly Property Source() As String
            Get
                Return Elements.GetValueByName("source")
            End Get
        End Property
#End Region

    End Class

    Public Class RssItemsCollection
        Inherits Collections.ObjectModel.ReadOnlyCollection(Of RssItem)

        Public Sub New(ByVal BaseList As List(Of RssItem))
            MyBase.New(BaseList)
        End Sub
    End Class

    ''' <summary>
    ''' Elément (propriété) d'un flux ou d'un item RSS
    ''' </summary>
    ''' <remarks></remarks>
    Public Class RssTag
        Private ReadOnly _Name As String
        Private ReadOnly _Value As String

        Public Sub New(ByVal Name As String, ByVal Value As String)
            _Name = Name
            _Value = Value
        End Sub

        Public ReadOnly Property Name() As String
            Get
                Return _Name
            End Get
        End Property

        Public ReadOnly Property Value() As String
            Get
                Return _Value
            End Get
        End Property

    End Class
    Public Class RssTagsCollection
        Inherits Collections.ObjectModel.ReadOnlyCollection(Of RssTag)

        Public Sub New(ByVal BaseList As List(Of RssTag))
            MyBase.New(BaseList)
        End Sub

        Public Function GetTagByName(ByVal TagName As String) As RssTag
            For Each It As RssTag In Me
                If It.Name.ToLower = TagName.ToLower Then Return It
            Next
            Return Nothing
        End Function

        Public Function GetValueByName(ByVal TagName As String, ByVal DefaultValue As String) As String
            For Each It As RssTag In Me
                If It.Name.ToLower = TagName.ToLower Then Return It.Value
            Next
            Return DefaultValue
        End Function

        Public Function GetValueByName(ByVal TagName As String) As String
            For Each It As RssTag In Me
                If It.Name.ToLower = TagName.ToLower Then Return It.Value
            Next
            Return ""
        End Function
    End Class

End Namespace