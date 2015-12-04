Imports ViewModels
Imports System.Globalization

Class Application

    ' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
    ' can be handled in this file.

#Region "Class Variables"

    ''' <summary>
    ''' Payroll Window Object.
    ''' </summary>
    Private prview As PayrollView

    ''' <summary>
    ''' Payroll Window Data Context Object.
    ''' </summary>
    Private prviewmodel As PayrollViewModel

    ''' <summary>
    ''' Login Window Object.
    ''' </summary>
    Private logview As LoginView

    ''' <summary>
    ''' Login Window Data Context Object.
    ''' </summary>
    Private logviewmodel As LoginViewModel

    ''' <summary>
    ''' Splash Window Object.
    ''' </summary>
    Private splashScreen As SplashScreen

    ''' <summary>
    ''' Timer That Sets The Length Of Time To Display The Splash Screen.
    ''' </summary>
    ''' <remarks></remarks>
    Private timer As Windows.Forms.Timer

#End Region

#Region "Startup Members"

    Private Sub Application_Startup(ByVal sender As Object, ByVal e As System.Windows.StartupEventArgs) Handles Me.Startup

        ' Initialises Global Settings For This App.
        Me.Init()

        ' Display The Splash Screen
        splashScreen = New SplashScreen("Images/PayProSplashScreen24.png")
        splashScreen.Show(False)

        '' Wait For 3 Seconds And Then Call 'CloseSplash' Method To Close The Splash Screen And Launch Login Window.
        timer = New Windows.Forms.Timer
        timer.Interval = 3000
        AddHandler timer.Tick, AddressOf CloseSplash
        timer.Start()
    End Sub

    ''' <summary>
    ''' Extends The 'Application_Startup' Method By Closing The Splash Screen And Opening The Login Window.
    ''' </summary>
    Private Sub CloseSplash(ByVal sender As Object, ByVal e As EventArgs)
        Dim span As TimeSpan = New TimeSpan(0, 0, 0, 1, 0)
        splashScreen.Close(span)
        timer.Dispose()
        Me.RunLogin()
    End Sub

    ''' <summary>
    ''' Opens The Login Window.
    ''' </summary>
    Private Sub RunLogin()
        logview = New LoginView
        logviewmodel = New LoginViewModel

        logview.DataContext = logviewmodel
        logview.ShowDialog()

        If logviewmodel.LoginCredentials.IsAuthentic = True Then
            Me.RunPayroll()
        Else
            My.Application.Shutdown()
        End If
    End Sub

    ''' <summary>
    ''' Opens The Main Payroll Window.
    ''' </summary>
    Private Sub RunPayroll()
        prview = New PayrollView
        prviewmodel = New PayrollViewModel

        prview.DataContext = prviewmodel
        prview.Show()
    End Sub

#End Region

#Region "Global Settings"

    ''' <summary>
    ''' Initialise Global Settings.
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub Init()

        ' HACK - Ensure The Current Culture Passed Into Bindings Is The OS Culture. 
        ' By Default, WPF Uses en-US As The Culture, Regardless Of The System Settings. 
        FrameworkElement.LanguageProperty.OverrideMetadata(GetType(FrameworkElement), New FrameworkPropertyMetadata(System.Windows.Markup.XmlLanguage.GetLanguage(CultureInfo.CurrentCulture.IetfLanguageTag)))

        ' Register This Routed Command So That 'ALL' The Text Within A Textbox Is Automatically Selected. This Will Affect 'ALL' Textboxes
        ' Throughout The Entire App.
        EventManager.RegisterClassHandler(GetType(TextBox), TextBox.GotFocusEvent, New RoutedEventHandler(AddressOf TextBox_GotFocus))

    End Sub

    ''' <summary>
    ''' Automatically Selects 'All' Text Within 'All' Textboxes.
    ''' </summary>
    Private Sub TextBox_GotFocus(ByVal sender As Object, ByVal e As EventArgs)
        Dim tbx As TextBox = CType(sender, TextBox)
        tbx.SelectAll()
    End Sub

#End Region

End Class


