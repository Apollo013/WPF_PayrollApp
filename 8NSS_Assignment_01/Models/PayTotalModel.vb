Imports Models.Base

Namespace Models

    Public Class PayTotalModel
        Inherits DataNotificationBase

#Region "Comments"

        ' This Model Will Be Used To Store The The Different Totals Of All Pay Packets In Each Pay Roll File.
        ' Examples Will Include Total Hours & Earnings For Basic Time Worked, Time And A Half And Also Double Time.
        ' We Will Also Total the Entire File And Get The Average Hours Worked And Average Earnings.

#End Region

#Region "Properties"

        Private _TotalsTitle As String
        ''' <summary>
        ''' Title For Each Total Type, Basic, Double Time, Etc ...
        ''' </summary>
        Public Property TotalsTitle() As String
            Get
                Return _TotalsTitle
            End Get
            Set(ByVal value As String)
                _TotalsTitle = value
                ReportPropertyChanged("TotalsTitle")
            End Set
        End Property

        Private _Hours As Double
        ''' <summary>
        ''' Total Hours For Each Total Type.
        ''' </summary>
        Public Property Hours() As Double
            Get
                Return _Hours
            End Get
            Set(ByVal value As Double)
                _Hours = value
                ReportPropertyChanged("Hours")
            End Set
        End Property

        Private _Earnings As Double
        ''' <summary>
        ''' Total Earnings For Each Total Type.
        ''' </summary>
        Public Property Earnings() As Double
            Get
                Return _Earnings
            End Get
            Set(ByVal value As Double)
                _Earnings = value
                ReportPropertyChanged("Earnings")
            End Set
        End Property

#End Region

    End Class

End Namespace
