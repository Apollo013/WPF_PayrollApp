Imports Models
'Imports System.Windows.Input
Imports ViewModels.Base

Namespace ViewModels

    Public Class SettingsViewModel
        Inherits ViewModelBase

#Region "Comments"

        ' This Object Serves To Bind Our 'SettingsModel' To The 'Settings' View As Part Of Our MVVM Architecture Design.
        ' All Data Entered By The User On The UI Is Pushed To The Data Model Through The 'WageSettings' Property Via The SaveExecute() Sub Routine.
        ' The 'Save' & 'Close' Properties Are Of Type ICommand So That We Can Bind Them To Our XAML Button Control's 'Command' Property.
        ' We Use The 'RelayCommand' Class To Add The Address Of The Execution Logic's Sub Routine For Each Command.

#End Region

#Region "Data Binding Properties"

        Private _WageSettings As New SettingsModel
        ''' <summary>
        ''' Property Used For Binding Our Wage Settings Data To The View.
        ''' </summary>
        Public ReadOnly Property WageSettings() As SettingsModel
            Get
                Return _WageSettings
            End Get
        End Property

#End Region

#Region "Constructor"

        Public Sub New()
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
        ''' Saves The Settings.
        ''' </summary>
        ''' <remarks>Execution Logic For Saving Settings</remarks>
        Private Sub SaveExecute()
            _WageSettings.Save()
            Me.Close()
        End Sub

        ''' <summary>
        ''' Execution Status Logic For Save Command.
        ''' </summary>
        ''' <returns>A Boolean Value Indicating Whether Or Not The WageSetting Object Is Valid.</returns>
        ''' <remarks>If The Object Is Not Valid Then The Save Button Should Disable.</remarks>
        Private Function CanSaveExecute() As Boolean
            Return _WageSettings.IsValid
        End Function

#End Region

#Region "'Close' ICommand Property & Methods"

        Private _CancelCommand As ICommand
        ''' <summary>
        ''' Property That Provides Command Binding For The View's 'Close' or 'Cancel' Button Control.
        ''' </summary>
        Public ReadOnly Property CancelCommand() As ICommand
            Get
                If _CancelCommand Is Nothing Then
                    _CancelCommand = New RelayCommand(AddressOf CancelExecute, AddressOf CanCloseExecute)
                End If
                Return _CancelCommand
            End Get
        End Property

        ''' <summary>
        ''' Closes The Current View.
        ''' </summary>
        ''' <remarks></remarks>
        Private Sub CancelExecute()
            _WageSettings.CancelEdit()
            Me.Close()
        End Sub

        ''' <summary>
        ''' Execution Status Logic For Close Command
        ''' </summary>
        ''' <returns>TRUE - Always</returns>
        ''' <remarks></remarks>
        Private Function CanCloseExecute() As Boolean
            Return True
        End Function

#End Region

    End Class

End Namespace
