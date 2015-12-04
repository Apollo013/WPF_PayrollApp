
Namespace ViewModels.Base

    Public MustInherit Class ViewModelBase
        Inherits UINotificationBase

#Region "Comments"

        ' This Class Provides ICommand Members For Basic Window Control Such As Minimising, Maximising, Normalising.
        ' This Is Achieved By Binding The 'currentWindowState' Property To A Views 'WindowState' Property - XAML = 'WindowState="{Binding currentWindowState}"'.
        ' Any Button Controls Should Also Be Bound To The Relevant ICommand Properties - XAML = 'Command="{Binding Path=MinimiseCommand}"'...

        ' This Class Also Handles The Window 'Close' Event & The Window 'Title' Can Also Be Set Here.

        ' The Class Inherits The 'UINotificationBase' Class Which Implements The 'INotifyPropertyChanged' Interface Used For updating The UI.

#End Region

#Region "Properties"

        Private _currentWindowState As WindowState = WindowState.Normal
        ''' <summary>
        ''' Gets Or Sets The WindowState Of The Current View.
        ''' </summary>
        ''' '''<remarks>By Default This Is WindowState.Normal</remarks>
        Public Property currentWindowState() As WindowState
            Get
                Return _currentWindowState
            End Get
            Set(ByVal value As WindowState)
                _currentWindowState = value
                ' Notify The UI
                ReportPropertyChanged("currentWindowState")
            End Set
        End Property

        Private _windowTitle As String = "Window Title Goes Here"
        ''' <summary>
        ''' Provides Binding For Specifying A Windows Title.
        ''' </summary>
        Public Property windowTitle() As String
            Get
                Return _windowTitle
            End Get
            Set(ByVal value As String)
                _windowTitle = value
                ReportPropertyChanged("windowTitle")
            End Set
        End Property

        Private _windowMaxHeight As Double = SystemParameters.WorkArea.Height + 10
        ''' <summary>
        ''' Get / Sets The Maximum Height Of The Current Window
        ''' </summary>
        ''' <remarks>By Default This Is SystemParameters.WorkArea.Height + 10</remarks>
        Public Property windowMaxHeight() As Double
            Get
                Return _windowMaxHeight
            End Get
            Set(ByVal value As Double)
                _windowMaxHeight = value
            End Set
        End Property

#End Region

#Region "Constructors"

        Public Sub New()
            ' Register 'DragWindow' Method To Handle 'All' Windows 'MouseLeftButtonDownEvent' So That We Can Drag The Window Across The Screen With The Mouse.
            EventManager.RegisterClassHandler(GetType(Window), Window.MouseLeftButtonDownEvent, New RoutedEventHandler(AddressOf DragWindow))
            EventManager.RegisterClassHandler(GetType(DataGridCell), Window.MouseDoubleClickEvent, New RoutedEventHandler(AddressOf MouseDoubleClickDataGrid))
        End Sub

#End Region

#Region "Minimise"

        Private _MinimiseCommand As ICommand
        ''' <summary>
        ''' Provides ICommand Binding For A View's 'Minimise' Control.
        ''' </summary>
        Public ReadOnly Property MinimiseCommand() As ICommand
            Get
                If _MinimiseCommand Is Nothing Then
                    _MinimiseCommand = New RelayCommand(AddressOf Me.MinimiseExecute)
                End If
                Return _MinimiseCommand
            End Get
        End Property

        ''' <summary>
        ''' Execution Logic For Minimising The Current View.
        ''' </summary>
        Public Sub MinimiseExecute()
            Me.currentWindowState = WindowState.Minimized
        End Sub

#End Region

#Region "Maximise"

        Private _MaximiseCommand As ICommand
        ''' <summary>
        ''' Provides ICommand Binding For A View's 'Maximise' Control.
        ''' </summary>
        Public ReadOnly Property MaximiseCommand() As ICommand
            Get
                If _MaximiseCommand Is Nothing Then
                    _MaximiseCommand = New RelayCommand(AddressOf Me.MaximiseExecute)
                End If
                Return _MaximiseCommand
            End Get
        End Property

        ''' <summary>
        ''' Execution Logic For Maximising The Current View.
        ''' </summary>
        Public Sub MaximiseExecute()
            If _currentWindowState <> WindowState.Maximized Then
                Me.currentWindowState = WindowState.Maximized
            Else
                Me.currentWindowState = WindowState.Normal
            End If
        End Sub

#End Region

#Region "Normalise"

        Private _NormaliseCommand As ICommand
        ''' <summary>
        ''' Provides ICommand Binding For A View's 'Restore' Control.
        ''' </summary>
        Public ReadOnly Property NormaliseCommand() As ICommand
            Get
                If _NormaliseCommand Is Nothing Then
                    _NormaliseCommand = New RelayCommand(AddressOf Me.NormaliseExecute)
                End If
                Return _NormaliseCommand
            End Get
        End Property

        ''' <summary>
        ''' Execution Logic For Restoring The Current View.
        ''' </summary>
        Public Sub NormaliseExecute()
            If _currentWindowState <> WindowState.Normal Then
                Me.currentWindowState = WindowState.Normal
            Else
                Me.currentWindowState = WindowState.Maximized
            End If
        End Sub

#End Region

#Region "Close"

        Private _CloseCommand As ICommand
        ''' <summary>
        ''' Provides ICommand Binding For A View's 'Close' Control.
        ''' </summary>
        ''' <remarks>This Is NOT Designed To Shut Down The Application</remarks>
        Public ReadOnly Property CloseCommand() As ICommand
            Get
                If _CloseCommand Is Nothing Then
                    _CloseCommand = New RelayCommand(AddressOf Me.Close)
                End If
                Return _CloseCommand
            End Get
        End Property

        ''' <summary>
        ''' Execution Logic Fot Closing Down The Current Active Window.
        ''' </summary>
        Public Sub Close()
            Me.CurrentWindow.Close()
        End Sub

#End Region

#Region "ShutDown"

        Private _ShutdownCommand As ICommand
        ''' <summary>
        ''' Provides ICommand Binding For A View's 'Close' Control.
        ''' </summary>
        Public ReadOnly Property ShutdownCommand() As ICommand
            Get
                If _ShutdownCommand Is Nothing Then
                    _ShutdownCommand = New RelayCommand(AddressOf Me.ApplicationShutdown)
                End If
                Return _ShutdownCommand
            End Get
        End Property

        ''' <summary>
        ''' Execution Logic Fot Shutting Down The Current Application.
        ''' </summary>
        Public Sub ApplicationShutdown()
            My.Application.Shutdown()
        End Sub

#End Region

#Region "Window Drag"

        ''' <summary>
        ''' Allows A Window To Be Dragged Across The Screen.
        ''' </summary>
        ''' <param name="sender">The Current Window</param>
        ''' <param name="e"></param>
        Private Sub DragWindow(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            If e.ButtonState = MouseButtonState.Pressed Then
                Dim win As Window = CType(sender, Window)
                win.DragMove()
            End If
        End Sub

#End Region

#Region "Current Window"

        ''' <summary>
        ''' Returns The Current 'Active' Window.
        ''' </summary>
        Public Function CurrentWindow() As Window
            Dim win As New Window
            For Each w As Window In Application.Current.Windows
                If w.IsActive Then
                    win = w
                End If
            Next
            Return win
        End Function

#End Region

#Region "Data Grid Mouse Double Click"

        Public Overridable Sub MouseDoubleClickDataGrid(ByVal sender As Object, ByVal e As MouseButtonEventArgs)
            'MsgBox("dbl clk")
        End Sub

#End Region

    End Class

End Namespace
