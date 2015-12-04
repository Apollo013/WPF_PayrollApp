Imports System.Collections.ObjectModel
Imports System.ComponentModel
Imports DataAccessLayer
Imports ViewModels.Base
Imports Models

Namespace ViewModels

    Public Class PayrollViewModel
        Inherits ViewModelBase

#Region "Variables"

        ''' <summary>
        ''' Handles The 'Navigation' Of Our Payroll Observable Collection.
        ''' </summary>
        ''' <remarks></remarks>
        Private _payrollView As ICollectionView

#End Region

#Region "Properties"

        Private _paypacketCollection As ObservableCollection(Of PayPacketModel)
        ''' <summary>
        ''' Stores A Collection Of PayPacketModel Onjects
        ''' </summary>
        ''' <remarks>Tracks Changes To All Objects In The Collection.</remarks>
        Public ReadOnly Property paypacketCollection() As ObservableCollection(Of PayPacketModel)
            Get
                If _paypacketCollection Is Nothing Then
                    _paypacketCollection = New ObservableCollection(Of PayPacketModel)

                    _payrollView = CollectionViewSource.GetDefaultView(_paypacketCollection)
                    ' Delegate A Method To Handle The 'ChangedEvent' Event Of '_payrollView' Collection View Source.
                    AddHandler _payrollView.CurrentChanged, AddressOf ViewChanged

                End If
                Return _paypacketCollection
            End Get
        End Property

        ''' <summary>
        ''' Returns The Number Of Pay Packets Currenty In The Payroll Collection.
        ''' </summary>
        ''' <returns>An Integer Value Specifying The Number Of Pay Packets.</returns>
        Public ReadOnly Property Count() As Integer
            Get
                Return Me.paypacketCollection.Count
            End Get
        End Property

        Private _CurrentPayPacket As PayPacketModel
        ''' <summary>
        ''' Returns The Currently Selected Pay Packet Object In Our Observable Collection.
        ''' </summary>
        Public Property CurrentPayPacket() As PayPacketModel
            Get
                Return _CurrentPayPacket
            End Get
            Set(ByVal value As PayPacketModel)
                _CurrentPayPacket = value
                ReportPropertyChanged("CurrentPayPacket")
                ' _CurrentPayPacket.NotifyUI()
            End Set
        End Property

        Private _dataAccess As TextFileAccess
        ''' <summary>
        ''' An Instance Of Our 'TextFileAccess' Object.
        ''' </summary>
        ''' <remarks>The Save / OpenFileDialog Title , Default File Extension and Filter Are Passed As Arguments.</remarks>
        Public ReadOnly Property DataAccess() As TextFileAccess
            Get
                If _dataAccess Is Nothing Then
                    _dataAccess = New TextFileAccess("Payroll File", ".prl", "Payroll Files|*.prl")
                End If
                Return _dataAccess
            End Get
        End Property

        Private _PayTotalCollection As ObservableCollection(Of PayTotalModel)
        Public ReadOnly Property PayTotalCollection() As ObservableCollection(Of PayTotalModel)
            Get
                If _PayTotalCollection Is Nothing Then
                    _PayTotalCollection = New ObservableCollection(Of PayTotalModel)
                End If
                Return _PayTotalCollection
            End Get
        End Property

#End Region

#Region "Constructor"

        Public Sub New()
            'Maximise The Current Window When Opened.
            Me.currentWindowState = WindowState.Maximized
        End Sub

#End Region

#Region "Load Methods"

        ''' <summary>
        ''' Populates The _payrollCollection Object
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub Load()

            ' _dataAccess.FileContent Is A Strongly Typed List<of String>.
            ' Each Item In _dataAccess.FileContent Corresponds To A PayPacket Property, i.e Either the Name, Hours or Rate.
            ' Therefore, We Must Read 3 Lines At A Time In Order To Make A PayPacket Object.

            Const itemsPerObject As Byte = 3

            ' Clear All Current Pay Packets.
            _paypacketCollection.Clear()

            ' Index Counter.
            Dim index As Integer = 0

            ' Iterate Through The List (Step = itemsPerObject or 3 In This Case).
            While index <= Me.DataAccess.FileContent.Count() - itemsPerObject
                ' Add A New Pay Packet Object.
                _paypacketCollection.Add(New PayPacketModel With {.Index = index,
                                                                .Name = Me.DataAccess.FileContent.Item(index),
                                                                .Hours = CDbl(Me.DataAccess.FileContent.Item(index + 1)),
                                                                .Rate = CDbl(Me.DataAccess.FileContent.Item(index + 2))})
                index += itemsPerObject
            End While

            ' Move To the First Pay Packet Object
            _payrollView.MoveCurrentToFirst()
            '_payrollView.
            ' Show The Count
            ReportPropertyChanged("Count")

            ' Refresh Totals
            Me.RefreshTotals()
        End Sub

        ''' <summary>
        ''' Refreshes Payroll Totals.
        ''' </summary>
        Private Sub RefreshTotals()
            ' Clear The Paytotal Collection.
            Me.PayTotalCollection.Clear()

            ' Only Reload It If There Are Currently Pay Packets Being Viewed.
            If Me.Count Then
                _PayTotalCollection.Add(New PayTotalModel With {.TotalsTitle = "Basic", .Hours = Me.CalculateBasicHours, .Earnings = Me.CalculateBasicEarnings})
                _PayTotalCollection.Add(New PayTotalModel With {.TotalsTitle = "Time And Half", .Hours = Me.CalculateTimeAndHalfHours, .Earnings = Me.CalculateTimeAndHalfEarnings})
                _PayTotalCollection.Add(New PayTotalModel With {.TotalsTitle = "Double Time", .Hours = Me.CalculateDoubleHours, .Earnings = Me.CalculateDoubleEarnings})
                _PayTotalCollection.Add(New PayTotalModel With {.TotalsTitle = "Gross Total", .Hours = Me.CalculateTotalHours, .Earnings = Me.CalculateTotalEarnings})
                _PayTotalCollection.Add(New PayTotalModel With {.TotalsTitle = "Average", .Hours = Me.CalculateAverageHours, .Earnings = Me.CalculateAverageEarnings})
                _PayTotalCollection.Add(New PayTotalModel With {.TotalsTitle = "Minimum", .Hours = Me.CalculateMinimumHours, .Earnings = Me.CalculateMinimumEarnings})
                _PayTotalCollection.Add(New PayTotalModel With {.TotalsTitle = "Maximum", .Hours = Me.CalculateMaximumHours, .Earnings = Me.CalculateMaximumEarnings})
            End If

            ' Update The UI.
            ReportPropertyChanged("PayTotalCollection")
        End Sub

#End Region

#Region "Event Handlers"

        ''' <summary>
        ''' Handles The CurrentChanged Event Of Our Observable Object View.
        ''' </summary>
        ''' <remarks>Used To Trap The Current Pay Packet Object Selected By The User.</remarks>
        Private Sub ViewChanged(ByVal sender As Object, ByVal e As EventArgs)
            Me._CurrentPayPacket = Me._payrollView.CurrentItem
        End Sub

#End Region

#Region "Payroll Total Calculations"

        ' The Following Members Calculate Total Basic, Time And Half, Double Time,
        ' Overall Totals and Averages For Hours And Earnings.
        ' LINQ Queries Are Used For All.

        Private p As PayPacketModel

        Private Function CalculateBasicHours() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.BasicHours
                   Into Sum()
        End Function

        Private Function CalculateBasicEarnings() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.BasicEarnings
                   Into Sum()
        End Function

        Private Function CalculateTimeAndHalfHours() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.TimeAndHalfHours
                   Into Sum()
        End Function

        Private Function CalculateTimeAndHalfEarnings() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.TimeAndHalfEarnings
                   Into Sum()
        End Function

        Private Function CalculateDoubleHours() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.DoubleTimeHours
                   Into Sum()
        End Function

        Private Function CalculateDoubleEarnings() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.DoubleTimeEarnings
                   Into Sum()
        End Function

        Private Function CalculateTotalHours() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.Hours
                   Into Sum()
        End Function

        Private Function CalculateTotalEarnings() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.TotalEarnings
                   Into Sum()
        End Function

        Private Function CalculateAverageHours() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.Hours
                   Into Average()
        End Function

        Private Function CalculateAverageEarnings() As Double
            Return Aggregate p In _paypacketCollection
                   Select p.TotalEarnings
                   Into Average()
        End Function

        Private Function CalculateMinimumHours() As Double
            Return Aggregate p In _paypacketCollection
                   Into Min(p.Hours)
        End Function

        Private Function CalculateMinimumEarnings() As Double
            Return Aggregate p In _paypacketCollection
                   Into Min(p.TotalEarnings)
        End Function

        Private Function CalculateMaximumHours() As Double
            Return Aggregate p In _paypacketCollection
                   Into Max(p.Hours)
        End Function

        Private Function CalculateMaximumEarnings() As Double
            Return Aggregate p In _paypacketCollection
                   Into Max(p.TotalEarnings)
        End Function

#End Region

#Region "'New File' ICommand Property & Methods"

        ' This Will Always Be Available, i.e. It Will Never Disable.

        Private _NewFileCommand As ICommand
        ''' <summary>
        ''' Provides Command Binding For The View's 'Create File' Control.
        ''' </summary>
        Public ReadOnly Property NewFileCommand() As ICommand
            Get
                If _NewFileCommand Is Nothing Then
                    _NewFileCommand = New RelayCommand(AddressOf Me.NewFileExecute)
                End If
                Return _NewFileCommand
            End Get
        End Property

        ''' <summary>
        ''' Execution Logic For Creating A New File.
        ''' </summary>
        Private Sub NewFileExecute()
            If Me.DataAccess.CreateFile Then
                Me.Load()
            End If
        End Sub

#End Region

#Region "'Open File' ICommand Property & Methods"

        ' This Will Always Be Available, i.e. It Will Never Disable.

        Private _FileOpenCommand As ICommand
        ''' <summary>
        ''' Provides Command Binding For This View's 'Open' Control (Button).
        ''' </summary>
        Public ReadOnly Property FileOpenCommand() As ICommand
            Get
                If _FileOpenCommand Is Nothing Then
                    _FileOpenCommand = New RelayCommand(AddressOf FileOpenExecute)
                End If
                Return _FileOpenCommand
            End Get
        End Property

        ''' <summary>
        ''' Calls The File Open Dialog.
        ''' </summary>
        ''' <remarks>Execution Logic For File Open Dialog</remarks>
        Private Sub FileOpenExecute()
            Me.DataAccess.OpenFile()
            Me.Load()
        End Sub

#End Region

#Region "'Delete File' ICommand Property & Methods"

        ' This Will Disable, Depending On Whether The File Exists.

        Private _DeleteFileCommand As ICommand
        ''' <summary>
        ''' Provides Command Binding For The View's 'Delete File' Control.
        ''' </summary>
        Public ReadOnly Property DeleteFileCommand() As ICommand
            Get
                If _DeleteFileCommand Is Nothing Then
                    _DeleteFileCommand = New RelayCommand(AddressOf Me.DeleteFileExecute, AddressOf CanDeleteFileExecute)
                End If
                Return _DeleteFileCommand
            End Get
        End Property

        ''' <summary>
        ''' Execution Logic For Deleting A File.
        ''' </summary>
        Private Sub DeleteFileExecute()
            Dim result As Integer = MessageBox.Show("Do You Want To Delete The Current Payroll File ?", "Please Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Question)
            If result = MessageBoxResult.No Then Return
            If Me.DataAccess.DeleteFile Then
                Me.Load()
            End If
        End Sub

        ''' <summary>
        ''' Determines Whether We Can Delete A File Or Not.
        ''' </summary>
        ''' <returns>True If We Can, False Otherwise</returns>
        ''' <remarks>Condition Depends On Whether The File Exists.</remarks>
        Private Function CanDeleteFileExecute() As Boolean
            Return Me.DataAccess.FileExists
        End Function

#End Region

#Region "'New Pay Packet' ICommand Property & Methods"

        ' This Will Disable, Depending On The Number Of Records Being Displayed.

        Private _InsertCommand As ICommand
        ''' <summary>
        '''  Provides Command Binding For This View's 'Add' Control (Button).
        ''' </summary>
        Public ReadOnly Property InsertCommand() As ICommand
            Get
                If _InsertCommand Is Nothing Then
                    _InsertCommand = New RelayCommand(AddressOf InsertExecute, AddressOf CanInsertExecute)
                End If
                Return _InsertCommand
            End Get
        End Property

        ''' <summary>
        ''' Adds A New Pay Packet.
        ''' </summary>
        Private Sub InsertExecute()
            ' Create A New Pay Packet Object
            Dim payObj As New PayPacketModel(True)

            ' Create Instances Of Our View And View Model Objects.
            Dim vPay As New PayPacketView

            ' Pass In References Concerning The Pay Packet Object And Our DataAccess Property Object
            Dim vmPay As New PayPacketViewModel(payObj, _dataAccess)

            ' Set The DataContext Of Our View And Open It.
            vPay.DataContext = vmPay
            vPay.ShowDialog()

            ' Add The New Record If It Was Saved Successfully.
            If vmPay.WasSuccessful Then
                Me.Load()
            End If

            ' Clear Objects
            vPay = Nothing
            vmPay = Nothing
        End Sub

        ''' <summary>
        ''' Indicates If Our Data Access Object Is Valid.
        ''' </summary>
        ''' <returns>True If Valid, False Otherwise.</returns>
        ''' <remarks></remarks>
        Private Function CanInsertExecute() As Boolean
            Return Me.DataAccess.FileExists
        End Function

#End Region

#Region "'Edit Pay Packet' ICommand Property & Methods"

        ' This Will Disable, Depending On The Number Of Records Being Displayed
        ' And If A Pay Packet Is Selected.

        Private _EditCommand As ICommand
        ''' <summary>
        '''  Provides Command Binding For This View's 'Edit' Control (Button).
        ''' </summary>
        Public ReadOnly Property EditCommand() As ICommand
            Get
                If _EditCommand Is Nothing Then
                    _EditCommand = New RelayCommand(AddressOf EditExecute, AddressOf CanEditExecute)
                End If
                Return _EditCommand
            End Get
        End Property

        ''' <summary>
        ''' Opens The 'PayPacketView' Window To Allow Edit Of Pay Packet..
        ''' </summary>
        ''' <remarks>Execution Logic For Opening Settings</remarks>
        Private Sub EditExecute()
            ' Create Instances Of Our View And View Model Objects (Passing In The Currently Selected Pay Packet).
            Dim vPay As New PayPacketView
            Dim vmUpdate As New PayPacketViewModel(_CurrentPayPacket, _dataAccess)

            ' Set The DataContext Of Our View And Open It.
            vPay.DataContext = vmUpdate
            vPay.ShowDialog()

            ' Refresh Totals If The  Record Was Saved Successfully.
            If vmUpdate.WasSuccessful Then
                Me.RefreshTotals()
            End If

            ' Clear Objects
            vPay = Nothing
            vmUpdate = Nothing
        End Sub

        ''' <summary>
        ''' Determines Whether We Can Edit Or Not.
        ''' </summary>
        ''' <returns>True If We Can, False Otherwise</returns>
        ''' <remarks>Condition Depends On  Whether There Are Records Present AND If One Is Actually Selected.</remarks>
        Private Function CanEditExecute() As Boolean
            If (Not _CurrentPayPacket Is Nothing) And (Not _paypacketCollection Is Nothing) Then
                Return _paypacketCollection.Count > 0 And (Not _CurrentPayPacket Is Nothing)
            Else
                Return False
            End If
        End Function

#End Region

#Region "'Delete Pay Packet' ICommand Property & Methods"

        ' This Will Disable, Depending On The Number Of Records Being Displayed
        ' And If A Pay Packet Is Selected.

        Private _DeleteCommand As ICommand
        ''' <summary>
        '''  Provides Command Binding For This View's 'Delete' Control (Button).
        ''' </summary>
        Public ReadOnly Property DeleteCommand() As ICommand
            Get
                If _DeleteCommand Is Nothing Then
                    _DeleteCommand = New RelayCommand(AddressOf DeleteExecute, AddressOf CanDeleteExecute)
                End If
                Return _DeleteCommand
            End Get
        End Property

        ''' <summary>
        ''' Deletes The Current Pay Packet Object.
        ''' </summary>
        Private Sub DeleteExecute()
            Dim result As Integer
            result = MessageBox.Show("The Pay Packet For " & _CurrentPayPacket.Name & " Will Be Permanently Lost." & Environment.NewLine & "Are You Sure You Want To Delete ?", "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning)
            If result = MessageBoxResult.Yes Then
                Me.DataAccess.Remove(_CurrentPayPacket.Index, 3)
                Me.Load()
            End If
        End Sub

        ''' <summary>
        ''' Determines Whether We Can Perform A Delete Or Not.
        ''' </summary>
        ''' <returns>True If We Can, False Otherwise</returns>
        ''' <remarks>Condition Depends On Whether There Are Records Present AND If One Is Actually Selected.</remarks>
        Private Function CanDeleteExecute() As Boolean
            If (Not _CurrentPayPacket Is Nothing) And (Not _paypacketCollection Is Nothing) Then
                Return _paypacketCollection.Count > 0 And (Not _CurrentPayPacket Is Nothing)
            Else
                Return False
            End If
        End Function

#End Region

#Region "'Settings' ICommand Property & Methods"

        'This Will Always Be Available, i.e. It Will Never Disable.

        Private _settingsCommand As ICommand
        ''' <summary>
        ''' Provides Command Binding For This View's 'Settings' Control (Button).
        ''' </summary>
        Public ReadOnly Property SettingsCommand() As ICommand
            Get
                If _settingsCommand Is Nothing Then
                    _settingsCommand = New RelayCommand(AddressOf SettingsExecute)
                End If
                Return _settingsCommand
            End Get
        End Property

        ''' <summary>
        ''' Opens The Settings View.
        ''' </summary>
        ''' <remarks>Execution Logic For Opening Settings</remarks>
        Private Sub SettingsExecute()
            Dim view As New SettingsView
            Dim viewmodel As New SettingsViewModel
            view.DataContext = viewmodel
            view.ShowDialog()

            ' Refresh Totals
            Me.Load()
        End Sub

#End Region

    End Class

End Namespace
