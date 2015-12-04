Imports Models.Base
Imports Helpers.Login
Imports System.Configuration
Imports System.Security

Namespace Models

    Public Class LoginModel
        Inherits DataNotificationBase

#Region "Comments"

        ' This Class Simply Allows Us To Compare The User Credentials Entered And With The Credentials Stored In Our Configuration File.
        ' This Class Only Inherits From The 'DataNotificationBase' Class (and not the 'ModelBase' Class like the others do).
        ' The 'DataNotificationBase' Class Implements The INotify Interface Used For Updating The UI.
        ' Because The Only Implementation We Need Is To Update The UI When The User Enters In User Credentials On The 'LoginView' Window.

#End Region

#Region "Constructor"

        Public Sub New()
        End Sub

#End Region

#Region "Properties"

        Private _Input_UserName As String
        ''' <summary>
        ''' Gets / Sets The 'UserName' As Entered By The User On The UI.
        ''' </summary>
        Public Property Input_UserName() As String
            Get
                Return _Input_UserName
            End Get
            Set(ByVal value As String)
                _Input_UserName = value
                ReportPropertyChanged("Input_UserName")
            End Set
        End Property

        Private _Input_Password As String
        ''' <summary>
        ''' Gets / Sets The 'Password' As Entered By The User On The UI.
        ''' </summary>
        Public Property Input_Password() As String
            Get
                Return _Input_Password
            End Get
            Set(ByVal value As String)
                _Input_Password = value
                ReportPropertyChanged("Input_Password")
            End Set
        End Property

        Private _IsAuthentic As Boolean = False
        ''' <summary>
        ''' Determines If The Login Credentials Are Valid.
        ''' </summary>
        ''' <returns>TRUE If Valid, Otherwise FALSE</returns>
        Public ReadOnly Property IsAuthentic() As Boolean
            Get
                '_IsAuthentic = Me.Authenticate
                Return Me.Authenticate '_IsAuthentic
            End Get
        End Property

#End Region

#Region "Methods"

        ''' <summary>
        ''' Compares The Credentials Entered By The User With Those Stored In Our Configuration File (Case Sensitive).
        ''' </summary>
        ''' <returns>TRUE If Both Sets Of Credentials Match, FALSE Otherwise.</returns>
        Private Function Authenticate() As Boolean
            Const Authentic As Integer = 0

            Dim resultA As Integer = String.Compare(Me.Input_UserName, My.Settings.UserName)
            Dim resultB As Integer = String.Compare(Me._Input_Password, My.Settings.Password)

            Return ((resultA = Authentic) And (resultB = Authentic))
        End Function

#End Region

    End Class

End Namespace
