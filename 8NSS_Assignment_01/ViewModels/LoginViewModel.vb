Imports Models
Imports Helpers.Login
Imports ViewModels.Base

Namespace ViewModels

    Public Class LoginViewModel
        Inherits ViewModelBase

#Region "Comments"

        ' This Class Binds Our 'LoginModel' MODEL To The 'LoginView' VIEW As Part Of Our MVVM Architecture Design.
        ' Credentials Entered By The User On The UI Are Pushed To The MODEL Through The 'LoginCredentials' Property Where The Validation Occurs.
        ' The 'LoginCommand' & 'CancelCommand' Properties Are Of Type ICommand So That We Can Bind Them To Our 'Login' or 'Cancel' Button Control's 'Command' Property In XAML.
        ' We Use The 'RelayCommand' Class To Add The Address Of The Execution Logic's Sub Routine For Each Command.

#End Region

#Region "Consructor"

        Public Sub New()
            ' Instantiate A New 'LoginModel' Object.
            _LoginCredentials = New LoginModel
        End Sub

#End Region

#Region "Data Binding Properties"

        Private _LoginCredentials As LoginModel
        ''' <summary>
        ''' Property Used For Binding Our Login Model Data To The View (and vice versa).
        ''' </summary>
        Public Property LoginCredentials() As LoginModel
            Get
                If _LoginCredentials Is Nothing Then
                    _LoginCredentials = New LoginModel
                End If
                Return _LoginCredentials
            End Get
            Set(value As LoginModel)
                _LoginCredentials = value
            End Set
        End Property

#End Region

#Region "Login ICommand Property & Methods"

        Private _LoginCommand As ICommand
        ''' <summary>
        ''' Property That Provides Command Binding For The View's 'Login' Button Control.
        ''' </summary>
        Public ReadOnly Property LoginCommand() As ICommand
            Get
                If _LoginCommand Is Nothing Then
                    _LoginCommand = New RelayCommand(AddressOf LoginExecute)
                End If
                Return _LoginCommand
            End Get
        End Property

        ''' <summary>
        ''' Attempts Login, Checks First That The Login Object Is Authentic.
        ''' </summary>
        ''' <remarks>Execution Logic For Attempting A Login.</remarks>
        Private Sub LoginExecute()
            If _LoginCredentials.IsAuthentic Then
                Me.Close()
            Else
                MessageBox.Show("User Name and Password Do Not Match.", "Invalid Login", MessageBoxButton.OK, MessageBoxImage.Exclamation)
            End If
        End Sub

#End Region

#Region "Cancel ICommand Property & Methods"

        Private _CancelCommand As ICommand
        ''' <summary>
        ''' Property That Provides Command Binding For The View's 'Close' or 'Cancel' Button Control.
        ''' </summary>
        Public ReadOnly Property CancelCommand() As ICommand
            Get
                If _CancelCommand Is Nothing Then
                    _CancelCommand = New RelayCommand(AddressOf Me.Close)
                End If
                Return _CancelCommand
            End Get
        End Property

#End Region

    End Class

End Namespace
