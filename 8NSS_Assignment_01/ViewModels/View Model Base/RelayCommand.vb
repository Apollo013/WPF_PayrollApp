Imports System.Windows.Input

Namespace ViewModels.Base

    Public Class RelayCommand
        Implements ICommand

#Region "Comments"

        ' The RelayCommand Class Will Be Used To Push The Command Of Our Buttons On Our Views To The ViewModels. 
        ' Since The ViewModels And The Views Are In Separate Layers, They Can Only Communicate Though Data Binding.
        ' The Default Return Value For The CanExecute Method Is 'True'.

#End Region

#Region "Declarations"

        Private ReadOnly _execute As Action
        Private ReadOnly _canexecute As Func(Of Boolean)

#End Region

#Region "Constructors"

        ''' <summary>
        ''' Creates A New Command That Can Always Execute.
        ''' </summary>
        ''' <param name="pExecute">The Execution Logic.</param>
        Public Sub New(ByVal pExecute As Action)
            Me.New(pExecute, Nothing)
        End Sub

        ''' <summary>
        ''' Creates A New Command.
        ''' </summary>
        ''' <param name="pExecute">The Execution Logic.</param>
        ''' <param name="pCanExecute">The Execution Status Logic.</param>
        Public Sub New(ByVal pExecute As Action, ByVal pCanExecute As Func(Of Boolean))
            If pExecute Is Nothing Then
                Throw New ArgumentNullException("execute")
            End If

            _execute = pExecute
            _canexecute = pCanExecute
        End Sub

#End Region

#Region "ICommand Members"

        <DebuggerStepThrough()> _
        Public Function CanExecute(ByVal parameter As Object) As Boolean Implements System.Windows.Input.ICommand.CanExecute
            Return If(_canexecute Is Nothing, True, _canexecute())
        End Function

        Public Custom Event CanExecuteChanged As EventHandler Implements ICommand.CanExecuteChanged

            AddHandler(ByVal value As EventHandler)
                AddHandler CommandManager.RequerySuggested, value
            End AddHandler

            RemoveHandler(ByVal value As EventHandler)
                RemoveHandler CommandManager.RequerySuggested, value
            End RemoveHandler

            RaiseEvent(ByVal sender As System.Object, ByVal e As System.EventArgs)
            End RaiseEvent
        End Event

        Public Sub Execute(ByVal parameter As Object) Implements System.Windows.Input.ICommand.Execute
            _execute()
        End Sub

#End Region

    End Class

End Namespace
