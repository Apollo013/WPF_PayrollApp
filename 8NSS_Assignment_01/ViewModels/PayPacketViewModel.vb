Imports Models
Imports ViewModels.Base
Imports DataAccessLayer

Namespace ViewModels

    Public Class PayPacketViewModel
        Inherits ViewModelBase

#Region "Comments"

        ' View Model For Updating Individual Pay Packets.

#End Region

#Region "Declarations"

        ''' <summary>
        ''' An Instance Of Our 'TextFileAccess' Object.
        ''' </summary>
        Private _dataAccess As TextFileAccess

#End Region

#Region "Properties"

        Private _payPacket As PayPacketModel
        ''' <summary>
        ''' Pay Packet Object Reference.
        ''' </summary>
        Public Property payPacket() As PayPacketModel
            Get
                Return _payPacket
            End Get
            Set(ByVal value As PayPacketModel)
                _payPacket = value
                ReportPropertyChanged("payPacket")
            End Set
        End Property

        Private _WasSuccessful As Boolean
        ''' <summary>
        ''' Indicates If The Record Was Successfully Saved.
        ''' </summary>
        ''' <returns>True If Successful, False Otherwise.</returns>
        Public ReadOnly Property WasSuccessful() As Boolean
            Get
                Return _WasSuccessful
            End Get
        End Property


#End Region

#Region "Constructor"

        ''' <summary>
        ''' Pay Packet Object Constructor.
        ''' </summary>
        ''' <param name="_payPack">The Pay Packet Object To Edit</param>
        ''' <param name="dataAccess">The Data Access Object For Persisting Our Changes.</param>
        ''' <remarks></remarks>
        Public Sub New(ByRef _payPack As PayPacketModel, ByRef dataAccess As TextFileAccess)
            _dataAccess = dataAccess
            _payPack.BeginEdit()
            _WasSuccessful = False
            _payPacket = _payPack
        End Sub

#End Region

#Region "'Save' ICommand Property & Methods"

        Private _SaveCommand As ICommand
        ''' <summary>
        ''' Property That Provides Command Binding For The View's 'Save' Button Control.
        ''' </summary>
        Public ReadOnly Property SaveCommand() As ICommand
            Get
                If _SaveCommand Is Nothing Then
                    _SaveCommand = New RelayCommand(AddressOf SaveExecute, AddressOf CanSaveExecute)
                End If
                Return _SaveCommand
            End Get
        End Property

        ''' <summary>
        ''' Saves The Pay Packet.
        ''' </summary>
        ''' <remarks>Execution Logic For Saving Settings</remarks>
        Private Sub SaveExecute()
            Try
                ' If This Is A New Record, Then Add It, Otherwise Just Change It.
                If _payPacket.IsNew Then
                    _dataAccess.Append(_payPacket.Name, CStr(_payPacket.Hours), CStr(_payPacket.Rate))
                Else
                    _dataAccess.Change(_payPacket.Index, _payPacket.Name, CStr(_payPacket.Hours), CStr(_payPacket.Rate))
                End If
                _WasSuccessful = True
                _payPacket.EndEdit()
            Catch ex As Exception
                MessageBox.Show(ex.ToString, "Error Saving Record", MessageBoxButton.OK, MessageBoxImage.Error)
                _WasSuccessful = False
            Finally
                Me.Close()
            End Try
        End Sub

        ''' <summary>
        ''' Execution Status Logic For Save Command.
        ''' </summary>
        ''' <returns>A Boolean Value Indicating Whether Or Not The Pay Packet Object Is Valid.</returns>
        ''' <remarks>If The Object Is Not Valid Then The Save Button Should Disable.</remarks>
        Private Function CanSaveExecute() As Boolean
            Return _payPacket.IsValid
        End Function

#End Region

#Region "'Cancel' ICommand Property & Methods"

        Private _CancelCommand As ICommand
        ''' <summary>
        ''' Property That Provides Command Binding For The View's 'Close' or 'Cancel' Button Control.
        ''' </summary>
        Public ReadOnly Property CancelCommand() As ICommand
            Get
                If _CancelCommand Is Nothing Then
                    _CancelCommand = New RelayCommand(AddressOf CancelExecute, AddressOf CanCancelExecute)
                End If
                Return _CancelCommand
            End Get
        End Property

        ''' <summary>
        ''' Closes The Current View.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub CancelExecute()
            ' Return Object To It's Original State
            _payPacket.CancelEdit()
            Me.Close()
        End Sub

        ''' <summary>
        ''' Execution Status Logic For Close Command
        ''' </summary>
        ''' <returns>TRUE - Always</returns>
        ''' <remarks></remarks>
        Private Function CanCancelExecute() As Boolean
            Return True
        End Function

#End Region

    End Class

End Namespace
