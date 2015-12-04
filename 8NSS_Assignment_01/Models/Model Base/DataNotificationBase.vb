Imports System.ComponentModel

Namespace Models.Base

    Public MustInherit Class DataNotificationBase
        Implements INotifyPropertyChanged

#Region "Comments"

        ' This Class Implements The INotifyPropertyChanged Interface That Will Notify The UI That A Data Bound Property Has Changed, 
        ' And Needs To Update It’s Data.

#End Region

#Region "INotifyPropertyChanged"

        ''' <summary>
        ''' Raises This Objects PropertyChanged Event.
        ''' </summary>
        ''' <param name="propertyname">The Name Of The Property Whose Value Has Changed.</param>
        Protected Sub ReportPropertyChanged(ByVal propertyname As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyname))
        End Sub

        ''' <summary>
        ''' Raised When A Property On This Object Has A New Value.
        ''' </summary>
        Public Event PropertyChanged As PropertyChangedEventHandler Implements System.ComponentModel.INotifyPropertyChanged.PropertyChanged

#End Region

#Region "Dubugging Aides"

        ''' <summary>
        ''' Warns The Developer If This Object Does Not Have A Public Property With The Specified Name.
        ''' This Method Does Not Exist In A Release Build.
        ''' </summary>
        <Conditional("DEBUG")> _
        <DebuggerStepThrough()> _
        Public Sub VerifyPropertyName(ByVal propertyname As String)

            ' If We Raise PropertyChanged And Do Not Specify A Property Name,
            ' All Properties On The Object Are Considered Changed By The Binding System.
            If [String].IsNullOrEmpty(propertyname) Then
                Return
            End If

            'Verify That The Property Name Matches A Real Public Instance Property In This Object.
            If TypeDescriptor.GetProperties(Me)(propertyname) Is Nothing Then
                Dim msg As String = "Invalid property name: " & propertyname

                If Me.ThrowOnInvalidPropertyName Then
                    Throw New ArgumentException(msg)
                Else
                    Debug.Fail(msg)
                End If
            End If
        End Sub

        ''' <summary>
        ''' Returns Whether An Exception Is Thrown, or If A DEBUG.Fail() Is Used
        ''' When An Invalid Property Name Is Passed To The VerifyPropertyName Method.
        ''' The Default Value Is FALSE, But Sub-Classes Can Override This.
        ''' </summary>
        Private _ThrowOnInvalidPropertyName As Boolean = False
        Protected Property ThrowOnInvalidPropertyName() As Boolean
            Get
                Return _ThrowOnInvalidPropertyName
            End Get
            Private Set(ByVal value As Boolean)
                _ThrowOnInvalidPropertyName = value
            End Set
        End Property

#End Region

    End Class

End Namespace


