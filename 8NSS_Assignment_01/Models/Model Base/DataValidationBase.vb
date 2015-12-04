Imports System.ComponentModel
Imports System.Collections.Generic

Namespace Models.Base

    Public MustInherit Class DataValidationBase
        Inherits DataNotificationBase
        Implements IDataErrorInfo

#Region "Comments"

        ' This Interface Allows Us To Display Data Errors In The View’s Data Bound Properties. 
        ' The Data Errors Notify The User That There Is A Problem With The Data Entered, 
        ' And It Must Be Corrected Before The Data Is Accepted.

        ' Thanks To Beth Massi For This - http://blogs.msdn.com/b/bethmassi/archive/2008/06/27/displaying-data-validation-messages-in-wpf.aspx

#End Region

#Region "IDataErrorInfo"

        ''' <summary>
        ''' Raised When The Error Count Has Changed.
        ''' </summary>
        ''' <param name="HasErrors">Returns A Boolean Value Indicating Whether Or Not The Object Has Data Errors.</param>
        Public Event ErrorListChanged(ByVal HasErrors As Boolean)

        ''' <summary>
        ''' This Dictionary Contains A List Of Our Validation Errors For Each Field/Property In The Object.
        ''' </summary>
        Private _ValidationErrorsList As New Dictionary(Of String, String)
        Public ReadOnly Property ValidationErrorsList As Dictionary(Of String, String)
            Get
                Return _ValidationErrorsList
            End Get
        End Property

        ''' <summary>
        ''' Adds An Error To The Dictionary.
        ''' </summary>
        ''' <param name="columnName" >The Name Of The Property That Has The Error.</param>
        ''' <param name="msg" >The Message To Display To The User.</param>
        Protected Sub AddError(ByVal columnName As String, ByVal msg As String)
            If Not ValidationErrorsList.ContainsKey(columnName) Then
                ValidationErrorsList.Add(columnName, msg)
            End If
        End Sub

        ''' <summary>
        ''' Removes A Specific Error From The Dictionary (If One Exists).
        ''' </summary>
        ''' <param name="columnname">The Property That No Longer Has The Error.</param>
        Protected Sub RemoveError(ByVal columnname As String)
            If ValidationErrorsList.ContainsKey(columnname) Then
                ValidationErrorsList.Remove(columnname)
            End If
        End Sub

        ''' <summary>
        ''' Returns Whether Or Not The Object Has Errors.
        ''' </summary>
        ''' <returns>A Boolean Value Indicating If The Object Has Errors.</returns>
        Public Overridable ReadOnly Property HasErrors() As Boolean
            Get
                Return (_ValidationErrorsList.Count > 0)
            End Get
        End Property

        ''' <summary>
        ''' Returns Whether Or Not The Object Is Valid.
        ''' </summary>
        ''' <returns>A Boolean Value Indicating Whether The Object Is Valid Or Not.</returns>
        Public ReadOnly Property IsValid() As Boolean
            Get
                Return (_ValidationErrorsList.Count = 0)
            End Get
        End Property

        ''' <summary>
        ''' Returns An Error Message Indicating What Is Wrong With The Object.
        ''' </summary>
        ''' <returns>A String Message When There Is A Problem OR NOTHING When No Problem Exists.</returns>
        Public ReadOnly Property [Error] As String Implements System.ComponentModel.IDataErrorInfo.Error
            Get
                If ValidationErrorsList.Count > 0 Then
                    Return String.Format("{0} data is invalid.", TypeName(Me))
                Else
                    Return Nothing
                End If
            End Get
        End Property

        ''' <summary>
        ''' Returns the Error Message For The Property With The Given Name.
        ''' </summary>
        ''' <param name="columnName">The Property That Has A Data Error.</param>
        ''' <returns>A String Message When There Is A Problem OR NOTHING When No Problem Exists</returns>
        Default Public ReadOnly Property Item(ByVal columnName As String) As String Implements System.ComponentModel.IDataErrorInfo.Item
            Get
                If ValidationErrorsList.ContainsKey(columnName) Then
                    Return ValidationErrorsList(columnName).ToString
                Else
                    Return Nothing
                End If
            End Get
        End Property

#End Region

    End Class

End Namespace
