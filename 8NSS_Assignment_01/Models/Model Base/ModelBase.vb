Imports System.ComponentModel

Namespace Models.Base

    Public MustInherit Class ModelBase
        Inherits DataValidationBase
        Implements IEditableObject

#Region "Comments"

        ' The ModelBase Is An Abstract Class Which Can Be Inherited By Any Model Class. 
        ' This Model Inherits The DataValidationBase Object And Also Implements The IEditableObject Interface. 
        ' With This Interface, We Can Allow The User To Accept The Valid Data, Reset Data, Or Cancel The Edit Altogether.
        ' We Can Also Track Any Changes To An Objects Properties.

        ' Backup & Restore
        ' Any Class That Inherits From This Class Will Automatically Have The Following Overrideable Methods Added To It's Structure.
        ' 1) BackupClear()
        ' 2) BackupRestore()
        ' 3) BackupData()
        ' Whether You Choose To Implement Or Add Any Implementation To These Methods Is Simply Up To The Developer And Application Requirements.
        ' See Comments Below In The 'Backup & Restore' Section For More Information.

#End Region

#Region "Constructor"

        ''' <summary>
        '''Reset the properties so that validation will be executed when the model is initialized.
        ''' </summary>
        Public Sub New()
            'BeginEdit()
        End Sub

#End Region

#Region "IEditableObject"

        ' Allows Us To Safely Handle Our Data When Updating It Either Through A Detail Form Or A DataGridView.

        ''' <summary>
        ''' Gets Or Sets Whether An Object Is Currently Being Edited.
        ''' </summary>
        ''' <remarks></remarks>
        Private _IsInEdit As Boolean
        Public Property IsInEdit() As Boolean
            Get
                Return _IsInEdit
            End Get
            Set(ByVal value As Boolean)
                _IsInEdit = value
                ReportPropertyChanged("IsInEdit")
            End Set
        End Property

        ''' <summary>
        ''' Begins The Editing Session.
        ''' </summary>
        Public Sub BeginEdit() Implements System.ComponentModel.IEditableObject.BeginEdit
            BackupData()
            IsInEdit = True
        End Sub

        ''' <summary>
        ''' Cancels The Editing Session.
        ''' </summary>
        Public Sub CancelEdit() Implements System.ComponentModel.IEditableObject.CancelEdit
            If Me.IsInEdit Then
                BackupRestore()
                ModificationsClear()
                IsInEdit = False
                BackupClear()
            End If
        End Sub

        ''' <summary>
        ''' Saves The Data And Ends The Editing Session.
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub EndEdit() Implements System.ComponentModel.IEditableObject.EndEdit
            ModificationsClear()
            IsInEdit = False
            BackupClear()
        End Sub

#End Region

#Region "Modifications"

        ' Track The Changes Of Any Of An Objects Properties.

        ''' <summary>
        ''' List That Tracks Any Changes In An Object.
        ''' </summary>
        Private _ModifiedList As New List(Of String)
        Public ReadOnly Property ModifiedList As List(Of String)
            Get
                Return _ModifiedList
            End Get
        End Property

        Public Event ModifiedListChanged(ByVal HasModifications As Boolean)

        ''' <summary>
        ''' Determines If ChangesList Has Any Changes.
        ''' </summary>
        ''' <value></value>
        ''' <returns>TRUE If There Are Changes, FALSE If There Are No Changes.</returns>
        Public Overridable ReadOnly Property HasModifications() As Boolean
            Get
                Return (_ModifiedList.Count > 0)
            End Get
        End Property

        ''' <summary>
        ''' Adds The Name Of A Property To ChangesList If The Property's Value Has Changed.
        ''' </summary>
        ''' <param name="propertyName">The Property Name Whose Value HAS Changed</param>
        Protected Sub ModificationAdd(ByVal propertyName As String)
            If Not _ModifiedList.Contains(propertyName) Then
                _ModifiedList.Add(propertyName)
                RaiseEvent ModifiedListChanged(_ModifiedList.Count > 0)
            End If
        End Sub

        ''' <summary>
        ''' Removes The Name Of A Property From ChangesList If The Property's Value Has Not Changed.
        ''' </summary>
        ''' <param name="propertyName">The Property Name Whose Value HAS NOT Changed</param>
        Protected Sub ModificationRemove(ByVal propertyName As String)
            If Not _ModifiedList.Contains(propertyName) Then
                _ModifiedList.Remove(propertyName)
                RaiseEvent ModifiedListChanged(_ModifiedList.Count > 0)
            End If
        End Sub

        ''' <summary>
        ''' Clears The ChangesList Of All Objects.
        ''' </summary>
        Protected Sub ModificationsClear()
            _ModifiedList.Clear()
            RaiseEvent ModifiedListChanged(_ModifiedList.Count > 0)
        End Sub

#End Region

#Region "Backup & Restore"

        ' The Sole Purpose Of These Methods Is To Backup & Restore An Objects Property Values.
        ' In Order For This To Work, A Structure (Of Some Sort) Must Me Created Within The Actual Object.
        ' The Variables Within That Structure Must Have The Same Data Types As That Of The Objects Properties.
        ' These Methods Are Called From All Methods Within The 'IEditableObject' Section Above.

        ''' <summary>
        '''  Resets The Backup Structure To It's Original State.
        ''' </summary>
        Protected MustOverride Sub BackupClear()

        ''' <summary>
        ''' Restores Properties From Backup.
        ''' </summary>
        Protected MustOverride Sub BackupRestore()

        ''' <summary>
        ''' Fills Backup From Properties.
        ''' </summary>
        Protected MustOverride Sub BackupData()

#End Region

    End Class

End Namespace

