Namespace Models

    Public Class SettingsModel
        Inherits Models.Base.ModelBase

#Region "Comments"

        ' The Purpose Of This Class Is To Retrieve & Write Values From & To The Applications Settings File.
        ' These Settings Relate To Min / Max Wage Limits & The Max Hours Ceiling For A Flat Week, Time & A Half, & Double Time.
        ' This Class Inherits From The ModelBase Class To Handle Any Property Changes & Data Validation Errors So That The UI Can Be Notified.

#End Region

#Region "Properties"

        Private _MinumumWage As Double = 0
        ''' <summary>
        ''' Gets / Sets The Minumum Wage.
        ''' </summary>
        ''' <value>A Double Value Specifying The Minimum Wage.</value>
        ''' <returns>A Double Value Containing The Minimum Wage.</returns>
        Public Property MinimumWage() As Double
            Get
                Return _MinumumWage
            End Get
            Set(ByVal value As Double)

                ' Property Has Changed So Remove The Previous Error (If It Exists).
                RemoveError("MinimumWage")

                ' Validate.
                If value <= 0 Then
                    AddError("MinimumWage", "Minimum Wage Must Be Greater Than Zero.")
                ElseIf value > Me.MaximumWage Then
                    AddError("MinimumWage", "Minimum Wage Cannot Be Greater Than The Maximum Wage.")
                ElseIf Not IsNumeric(value) Then
                    AddError("MinimumWage", "Only Numeric Values Are Permitted For Minimum Wage.")
                End If

                ' Assign The New Value (Even If Does Not Meet The Validation Criteria - This Will Force The UI To Display The Error).
                _MinumumWage = value

                ' Notify The UI.
                Me.UpdateUI()
            End Set
        End Property

        Private _MaximumWage As Double = 0
        ''' <summary>
        ''' Gets / Sets The Maximum Wage.
        ''' </summary>
        ''' <value>A Double Value Specifying The Maximum Wage.</value>
        ''' <returns>A Double Value Containing The Maximum Wage.</returns>
        Public Property MaximumWage() As Double
            Get
                Return _MaximumWage
            End Get
            Set(ByVal value As Double)

                ' Property Has Changed So Remove The Previous Error (If It Exists).
                RemoveError("MaximumWage")

                ' Validate.
                If value <= 0 Then
                    AddError("MaximumWage", "Maximum Wage Must Be Greater Than Zero.")
                ElseIf value < Me.MinimumWage Then
                    AddError("MaximumWage", "Maximum Wage Cannot Be Less Than The Minimum Wage.")
                ElseIf Not IsNumeric(value) Then
                    AddError("MaximumWage", "Only Numeric Values Are Permitted For Maximum Wage.")
                End If

                ' Assign The New Value.
                _MaximumWage = value

                ' Notify The UI.
                Me.UpdateUI()
            End Set
        End Property

        Private _FlatWeekCeiling As Double
        ''' <summary>
        ''' Gets / Sets The Maximum Ceiling Value For The Hours Worked During A Flat Week.
        ''' </summary>
        ''' <value>An Integer Value Specifying The Maximum Hours Worked During A Flat Week.</value>
        ''' <returns>An Integer Value Containing The Maximum Hours Worked During A Flat Week.</returns>
        Public Property FlatWeekCeiling() As Double
            Get
                Return _FlatWeekCeiling
            End Get
            Set(ByVal value As Double)

                ' Property Has Changed So Remove The Previous Error (If It Exists).
                RemoveError("FlatWeekCeiling")

                ' Validate
                If value < 0 Then
                    AddError("FlatWeekCeiling", "The Ceiling Value For A Flat Week Cannot Contain A Negative Value.")
                ElseIf value >= Me.TimeAndHalfCeiling Then
                    AddError("FlatWeekCeiling", "The Ceiling Value For A Flat Week Must Be Less Than That For Time And A Half.")
                ElseIf Not IsNumeric(value) Then
                    AddError("FlatWeekCeiling", "Only Numeric Values Are Permitted For The Number Of Hours Worked In A Flat Week.")
                End If

                ' Assign The New Value.
                _FlatWeekCeiling = value

                ' Notify The UI.
                Me.UpdateUI()
            End Set
        End Property

        Private _TimeAndHalfCeiling As Double
        ''' <summary>
        ''' Gets / Sets The Maximum Ceiling Value For The Hours Worked For Time And A Half.
        ''' </summary>
        ''' <value>An Integer Value Specifying The Maximum Hours Worked For Time And A Half.</value>
        ''' <returns>An Integer Value Containing The Maximum Hours Worked For Time And A Half.</returns>
        Public Property TimeAndHalfCeiling() As Double
            Get
                Return _TimeAndHalfCeiling
            End Get
            Set(ByVal value As Double)

                ' Property Has Changed So Remove The Previous Error (If It Exists).
                RemoveError("TimeAndHalfCeiling")

                ' Validate.
                If value < 0 Then
                    AddError("TimeAndHalfCeiling", "The Ceiling Value For Time And A Half Hours Cannot Contain A Negative Value.")
                ElseIf value <= Me.FlatWeekCeiling Then
                    AddError("TimeAndHalfCeiling", "The Ceiling Value For Time And A Half Hours Must Be Greater Than That For A Flat Week.")
                ElseIf value >= Me.DoubleTimeCeiling Then
                    AddError("TimeAndHalfCeiling", "The Ceiling Value For Time And A Half Hours Must Be Less Than That For Double Time.")
                ElseIf Not IsNumeric(value) Then
                    AddError("TimeAndHalfCeiling", "Only Numeric Values Are Permitted For The Number Of Hours Worked For Time And A Half.")
                End If

                ' Assign The New Value.
                _TimeAndHalfCeiling = value

                ' Notify The UI.
                Me.UpdateUI()

            End Set
        End Property

        Private _DoubleTimeCeiling As Double
        ''' <summary>
        ''' Gets / Sets The Maximum Ceiling Value For The Hours Worked For Double Time.
        ''' </summary>
        ''' <value>An Integer Value Specifying The Maximum Hours Worked For Double Time.</value>
        ''' <returns>An Integer Value Containing The Maximum Hours Worked For Double Time.</returns>
        Public Property DoubleTimeCeiling() As Double
            Get
                Return _DoubleTimeCeiling
            End Get
            Set(ByVal value As Double)

                ' Property Has Changed So Remove The Previous Error (If It Exists).
                RemoveError("DoubleTimeCeiling")

                ' Validate
                If value < 0 Then
                    AddError("DoubleTimeCeiling", "The Ceiling Value For Double Time Hours Cannot Contain A Negative Value.")
                ElseIf value <= Me.TimeAndHalfCeiling Then
                    AddError("DoubleTimeCeiling", "The Ceiling Value For Double Time Hours Must Be Greater Than That For Time And A Half.")
                ElseIf Not IsNumeric(value) Then
                    AddError("DoubleTimeCeiling", "Only Numeric Values Are Permitted For The Number Of Hours Worked For Double Time.")
                End If

                ' Assign The New Value.
                _DoubleTimeCeiling = value

                ' Notify The UI
                Me.UpdateUI()
            End Set
        End Property

#End Region

#Region "Methods"

        ''' <summary>
        ''' Initialise Our Validation Object, Properties And Make A Backup.
        ''' </summary>
        Public Sub New()
            MyBase.New()
            Me.Read()
            Me.BackupData()
        End Sub

        ''' <summary>
        ''' Initialises Setting Properties.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub Read()
            _MinumumWage = My.Settings.MinimumWage
            _MaximumWage = My.Settings.MaximumWage
            _FlatWeekCeiling = My.Settings.Hours_FlatWeekCeiling
            _TimeAndHalfCeiling = My.Settings.Hours_TimeAndHalfCeiling
            _DoubleTimeCeiling = My.Settings.Hours_DoubleTimeCeiling
        End Sub

        ''' <summary>
        ''' Save Our Settings And Do Some House Keeping. 
        ''' </summary>
        Public Sub Save()
            My.Settings.MinimumWage = _MinumumWage
            My.Settings.MaximumWage = _MaximumWage
            My.Settings.Hours_FlatWeekCeiling = _FlatWeekCeiling
            My.Settings.Hours_TimeAndHalfCeiling = _TimeAndHalfCeiling
            My.Settings.Hours_DoubleTimeCeiling = _DoubleTimeCeiling
            My.Settings.Save()
            Me.EndEdit()
        End Sub

        ''' <summary>
        ''' Notifies The UI That Changes Have Been Made To one Or More Properties That May Affect Other Properties.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub UpdateUI()
            ReportPropertyChanged("MinimumWage")
            ReportPropertyChanged("MaximumWage")
            ReportPropertyChanged("FlatWeekCeiling")
            ReportPropertyChanged("TimeAndHalfCeiling")
            ReportPropertyChanged("DoubleTimeCeiling")
        End Sub

#End Region

#Region "Backp & Restore"

        ''' <summary>
        ''' Structure Used To Make A Copy Of Wage Settings.
        ''' </summary>
        ''' <remarks>This Will Allow The End User To Cancel Or Restore An Update (When Using A Form).</remarks>
        Private Structure SettingsBackup
            Dim MinWage As Double
            Dim MaxWage As Double
            Dim FlatWeekHours As Double
            Dim TimeandHalfHours As Double
            Dim DoubleTimeHours As Double
        End Structure

        ''' <summary>
        ''' Create A New Instance Of The Backup Structure.
        ''' </summary>
        Private Backup As New SettingsBackup

        ''' <summary>
        ''' Clear The Contents Of The Backup Structure.
        ''' </summary>
        Protected Overrides Sub BackupClear()
            Backup = New SettingsBackup
        End Sub

        ''' <summary>
        ''' Make A Copy Of Our Wage Settings.
        ''' </summary>
        Protected Overrides Sub BackupData()
            Backup.MinWage = Me.MinimumWage
            Backup.MaxWage = Me.MaximumWage
            Backup.FlatWeekHours = Me.FlatWeekCeiling
            Backup.TimeandHalfHours = Me.TimeAndHalfCeiling
            Backup.DoubleTimeHours = Me.DoubleTimeCeiling
        End Sub

        ''' <summary>
        ''' Restore Previous Settings.
        ''' </summary>
        Protected Overrides Sub BackupRestore()
            Me.MinimumWage = Backup.MinWage
            Me.MaximumWage = Backup.MaxWage
            Me.FlatWeekCeiling = Backup.FlatWeekHours
            Me.TimeAndHalfCeiling = Backup.TimeandHalfHours
            Me.DoubleTimeCeiling = Backup.DoubleTimeHours
        End Sub

#End Region

    End Class

End Namespace
