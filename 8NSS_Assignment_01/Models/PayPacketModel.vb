Imports Models.Base
Imports Helpers

Namespace Models

    ''' <summary>
    ''' Representation Of An Employees Pay Packet
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PayPacketModel
        Inherits ModelBase

#Region "Comments"

        ' This Class Stores All The Pay Packs For Each Employee.
        ' It Inherits From the 'ModelBase' Class.
        ' Only 3 Properties Will Be Changed Directly By The User; Name, Hours & Rate.
        ' The Remaining Properties Will Automatically Recalculate Based On The Hours & Rate Properties.

        ' Because We Are Also Going To Allow Edits & Deletes Aswell As Inserts,
        ' We'll Need To Keep Track Of The Starting Index Position Of Each Record Within Our 'Data' Text File.
        ' We'll Do This With The 'Index' Property Which Will Point To The Position Of The 'Name' Property.        

#End Region

#Region "Constructor"

        ''' <summary>
        ''' Pay Packet Constructor.
        ''' </summary>
        ''' <param name="isNew">If A New 'Record' Is Being Created, Then Pass True As The Arguement.</param>
        ''' <remarks></remarks>
        Public Sub New(Optional ByVal isNew As Boolean = False)

            ' Initialising Variables So That Validation Will Kick In Immediately.
            Me.Name = ""
            Me.Hours = 0
            Me.Rate = 0

            ' Flag Object As A New When 'Inserting'.
            _IsNew = isNew

        End Sub

#End Region

#Region "Properties"

        Private _Index As Integer
        ''' <summary>
        ''' Represents The Position Within The Text File For The Beginning Of Each Record.
        ''' </summary>
        ''' <remarks>This Will Allow Us To Edit and Delete Records Within Our Text File</remarks>
        Public Property Index() As Integer
            Get
                Return _Index
            End Get
            Set(ByVal value As Integer)
                _Index = value
                ' Once This Has Been Assigned An Index Number, This Is No Longer Considered A New Record.
                _IsNew = False
            End Set
        End Property

        Private _Name As String
        ''' <summary>
        ''' The Name of The Employee.
        ''' </summary>
        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)

                ' Property Has Changed So Remove The Previous Error (If It Exists).
                Me.RemoveError("Name")

                ' Validate.
                If String.IsNullOrEmpty(value) Then
                    Me.AddError("Name", "A Name For The Employee Must Be Provided.")
                End If

                ' Assign New Value.
                _Name = value

                ' Notify The UI.
                ReportPropertyChanged("Name")

            End Set
        End Property

        Private _Hours As Double = 0
        ''' <summary>
        ''' The Total Number Of Hours Worked By The Employee.
        ''' </summary>
        Public Property Hours() As Double
            Get
                Return FormatNumber(_Hours)
            End Get
            Set(ByVal value As Double)

                If value <> _Hours Then

                    ' Property Has Changed So Remove The Previous Error (If It Exists).
                    Me.RemoveError("Hours")

                    ' Validate.
                    If value < 0 Or value > My.Settings.Hours_DoubleTimeCeiling Then
                        Me.AddError("Hours", "Hours Must Be Between 0 and " & My.Settings.Hours_DoubleTimeCeiling)
                    End If

                    ' Assign New Value.
                    _Hours = value

                    ' Notify The UI.
                    ReportPropertyChanged("Hours")

                    ' Force A Recalculation Of All Rates, Hours & Earnings.
                    NotifyUI()

                End If

            End Set
        End Property

        Private _Rate As Double = 0
        ''' <summary>
        ''' The Standard Hourly Rate.
        ''' </summary>
        Public Property Rate() As Double
            Get
                Return FormatNumber(_Rate)
            End Get
            Set(ByVal value As Double)

                If value <> _Rate Then

                    ' Property Has Changed So Remove The Previous Error (If It Exists).
                    Me.RemoveError("Rate")

                    ' Validate.
                    If value <= 0 Then
                        AddError("Rate", "Rate Must Be Greater Than Zero.")
                    ElseIf value < My.Settings.MinimumWage Then
                        AddError("Rate", "Rate Cannot Be Less Than The Minimum Wage Allowed.")
                    ElseIf value > My.Settings.MaximumWage Then
                        AddError("Rate", "Rate Cannot Be Greater Than The Maximum Wage Allowed.")
                    End If

                    ' Assign New Value.
                    _Rate = value

                    ' Notify The UI.
                    ReportPropertyChanged("Rate")

                    ' Force A Recalculation Of All Rates, Hours & Earnings.
                    NotifyUI()
                End If

            End Set
        End Property

        Private _IsNew As Boolean
        ''' <summary>
        ''' Determines If This Is An Insert Rather Than And An Edit.
        ''' </summary>
        Public ReadOnly Property IsNew() As Boolean
            Get
                Return _IsNew
            End Get
        End Property

#End Region

#Region "Hours Breakdown Properties"

        Private _BasicHours As Double = 0
        ''' <summary>
        ''' The Basic Hours Worked For A Flat Week.
        ''' </summary>
        Public ReadOnly Property BasicHours() As Double
            Get
                ' Calculate
                If _Hours = 0 Then                                              ' No Hours Worked
                    _BasicHours = 0
                ElseIf _Hours < My.Settings.Hours_FlatWeekCeiling Then          ' Some Basic Hours Worked
                    _BasicHours = _Hours
                Else                                                            ' Full Criteria For Basic Hours Worked
                    _BasicHours = My.Settings.Hours_FlatWeekCeiling
                End If
                Return _BasicHours
            End Get
        End Property

        Private _TimeAndHalfHours As Double = 0
        ''' <summary>
        ''' The Hours Worked For Time And A Half.
        ''' </summary>
        Public ReadOnly Property TimeAndHalfHours() As Double
            Get
                ' Calculate
                If _Hours <= My.Settings.Hours_FlatWeekCeiling Then             ' No Time And A Half Hours Worked
                    _TimeAndHalfHours = 0
                ElseIf _Hours < My.Settings.Hours_TimeAndHalfCeiling Then       ' Some Time And A Half Hours Worked
                    _TimeAndHalfHours = _Hours - My.Settings.Hours_FlatWeekCeiling
                Else                                                            ' Full Criteria For Time And A Half Hours Worked
                    _TimeAndHalfHours = My.Settings.Hours_TimeAndHalfCeiling - My.Settings.Hours_FlatWeekCeiling
                End If
                Return _TimeAndHalfHours
            End Get
        End Property

        Private _DoubleTimeHours As Double = 0
        ''' <summary>
        ''' The Hours Worked For Double Time.
        ''' </summary>
        Public ReadOnly Property DoubleTimeHours() As Double
            Get
                ' Calculate
                If _Hours <= My.Settings.Hours_TimeAndHalfCeiling Then          ' No Double Time Hours Worked
                    _DoubleTimeHours = 0
                ElseIf _Hours < My.Settings.Hours_DoubleTimeCeiling Then        ' Some Double Time Hours Worked
                    _DoubleTimeHours = _Hours - My.Settings.Hours_TimeAndHalfCeiling
                Else                                                            ' Full Criteria For Double Time Hours Worked
                    _DoubleTimeHours = My.Settings.Hours_DoubleTimeCeiling - My.Settings.Hours_TimeAndHalfCeiling
                End If
                Return _DoubleTimeHours
            End Get
        End Property

#End Region

#Region "Rates Breakdown"

        Private _TimeAndHalfRate As Double
        Public ReadOnly Property TimeAndHalfRate() As Double
            Get
                Return Me.Rate * 1.5
            End Get
        End Property

        Private _DoubleRate As Double
        Public ReadOnly Property DoubleRate() As Double
            Get
                Return Me.Rate * 2
            End Get
        End Property

#End Region

#Region "Earnings Breakdown Properties"

        Private _BasicEarnings As Double = 0
        ''' <summary>
        ''' Earnings Excluding Any Overtime.
        ''' </summary>
        Public ReadOnly Property BasicEarnings() As Double
            Get
                Return Me.Rate * Me.BasicHours
            End Get
        End Property

        Private _TimeAndHalfEarnings As Double = 0
        ''' <summary>
        ''' Overtime Earnings For Time And A Half (Excluding Double Time).
        ''' </summary>
        Public ReadOnly Property TimeAndHalfEarnings() As Double
            Get
                Return Me.TimeAndHalfRate * Me.TimeAndHalfHours
            End Get
        End Property

        Private _DoubleTimeEarnings As Double = 0
        ''' <summary>
        ''' Overtime Earnings For Double Time (Excluding Time And A Half)
        ''' </summary>
        Public ReadOnly Property DoubleTimeEarnings() As Double
            Get
                Return Me.DoubleRate * Me.DoubleTimeHours
            End Get
        End Property

        Private _OvertimeEarnings As Double
        ''' <summary>
        ''' Total Overtime Earnings.
        ''' </summary>
        Public ReadOnly Property OvertimeEarnings() As Double
            Get
                Return Me.TimeAndHalfEarnings + Me.DoubleTimeEarnings
            End Get
        End Property

        Private _TotalEarnings As Double = 0
        ''' <summary>
        ''' Total Earnings Including Basic Earnings And Any Overtime.
        ''' </summary>
        Public ReadOnly Property TotalEarnings() As Double
            Get
                Return Me.BasicEarnings + Me.TimeAndHalfEarnings + Me.DoubleTimeEarnings
            End Get
        End Property

#End Region

#Region "UI Methods"

        ''' <summary>
        ''' Forces The UI To Re-Display (and re-calculate) Overtime Hours And Earnings.
        ''' </summary>
        ''' <remarks>Called By Both 'Set' Methods Of The Hours and Rate Properties</remarks>
        Public Sub NotifyUI()
            ReportPropertyChanged("BasicHours")
            ReportPropertyChanged("TimeAndHalfRate")
            ReportPropertyChanged("DoubleRate")
            ReportPropertyChanged("TimeAndHalfHours")
            ReportPropertyChanged("DoubleTimeHours")
            ReportPropertyChanged("BasicEarnings")
            ReportPropertyChanged("TimeAndHalfEarnings")
            ReportPropertyChanged("DoubleTimeEarnings")
            ReportPropertyChanged("TotalEarnings")
            ReportPropertyChanged("EarningsToText")
            ReportPropertyChanged("NumToText")
            ReportPropertyChanged("OvertimeEarnings")
        End Sub

#End Region

#Region "Backup & Restore"

        ''' <summary>
        ''' Structure Used To Make A Copy Of Pay Packet Settings.
        ''' </summary>
        ''' <remarks>This Will Allow The End User To Cancel Or Restore An Update (When Using A Form).</remarks>
        Private Structure PayPacketBackup
            Dim Name As String
            Dim Hours As Double
            Dim Rate As Double
        End Structure

        ''' <summary>
        ''' Create A New Instance Of The Backup Structure.
        ''' </summary>
        Private Backup As New PayPacketBackup

        ''' <summary>
        ''' Clear The Contents Of The Backup Structure.
        ''' </summary>
        Protected Overrides Sub BackupClear()
            Backup = Nothing
        End Sub

        ''' <summary>
        ''' Make A Copy Of Our Pay Packet Object.
        ''' </summary>
        Protected Overrides Sub BackupData()
            Backup.Name = Me.Name
            Backup.Hours = Me.Hours
            Backup.Rate = Me.Rate
        End Sub

        ''' <summary>
        ''' Restore The Object Back To It's Original State
        ''' </summary>
        Protected Overrides Sub BackupRestore()
            Me.Name = Backup.Name
            Me.Hours = Backup.Hours
            Me.Rate = Backup.Rate
        End Sub

#End Region

    End Class

End Namespace

