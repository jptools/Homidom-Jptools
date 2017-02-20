

'Imports System.ComponentModel


Public Class uThermostat
    Dim _Value As Integer = 0

    Public Event ChangeValueThermo(ByVal Value As Integer)


    Public Property Value As Integer
        Get
            Return _Value
        End Get
        Set(value2 As Integer)
            _Value = value2
            Dim mycolor As System.Windows.Media.Color
            '  Dim R, G, B As Byte           
            mycolor = System.Windows.Media.Color.FromRgb(_Value * 8, 0, (31 - _Value) * 8)
            Dim mybrush = New SolidColorBrush(mycolor)
            ProgressBar1.Foreground = mybrush
            ProgressBar1.Value = (_Value * 4) - 20
            Temp.Content = _Value.ToString
        End Set
    End Property


    Public Sub New(Title As String)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
      

    End Sub

    Private Sub BtnUp_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnUp.Click        
        Dim mycolor As System.Windows.Media.Color
        _Value += 1
        If _Value > 30 Then _Value = 30            
        mycolor = System.Windows.Media.Color.FromRgb(_Value * 8, 0, (31 - _Value) * 8)
        Dim mybrush = New SolidColorBrush(mycolor)
        ProgressBar1.Foreground = mybrush
        ProgressBar1.Value = (_Value * 4) - 20
        Temp.Content = _Value.ToString
        RaiseEvent ChangeValueThermo(_Value)
    End Sub

    Private Sub BtnDown_Click(ByVal sender As System.Object, ByVal e As System.Windows.RoutedEventArgs) Handles BtnDown.Click
        Dim mycolor As System.Windows.Media.Color
        _Value -= 1
        If _Value < 5 Then _Value = 5             
        mycolor = System.Windows.Media.Color.FromRgb(_Value * 8, 0, (31 - _Value) * 8)
        Dim mybrush = New SolidColorBrush(mycolor)
        ProgressBar1.Foreground = mybrush
        ProgressBar1.Value = (_Value * 4) - 20
        Temp.Content = _Value.ToString
        RaiseEvent ChangeValueThermo(_Value)
    End Sub


End Class

