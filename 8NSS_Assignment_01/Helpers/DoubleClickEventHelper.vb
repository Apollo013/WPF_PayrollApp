Imports System.Windows.Controls.Primitives
Imports System.Windows.Input
Imports System.Collections.Generic
Imports System.Windows.Controls
Imports System.Windows

Namespace Helpers.Events

    Public NotInheritable Class DoubleClickEventHelper

#Region "Comments"

        ' Thanks To http://sachabarber.net/?p=532 For The C# Version Which I Ported To VB.

        ' This Is An Attached Behaviour Dependancy Property.
        ' This Class Allows Us To Bind A MouseDoubleClick 'Event' On An ItemsControl (Datagrid, ListView), To A Command So That We Can Execute Whatever Logic We Want (Usually To Edit A Record).
        ' This Is Only Required If We Are Implementing A 'Pure' MVVM Architecture In A WPF Solution.
        ' This Is A NotInheritable / Static Class.

        ' *** XAML Implementation ***
        ' 1) Create An XML Namespace To Reference This Class  -  
        '               xmlns:local="clr-namespace:Helpers.Events"
        ' 2) Attach The Property To The ItemsControl
        '               local:DoubleClickEventHelper.HandleDoubleClick="True"
        '               local:DoubleClickEventHelper.TheCommandToRun="{Binding Path=EditCommand}">

#End Region

#Region "Double Click Dependancy Property Members"

        ''' <summary>
        ''' Attached Double Click Dependancy Property.
        ''' </summary>
        Public Shared ReadOnly DoubleClickProperty As DependencyProperty = DependencyProperty.RegisterAttached _
                                                                            ("HandleDoubleClick", _
                                                                            GetType(Boolean), _
                                                                            GetType(DoubleClickEventHelper), _
                                                                            New FrameworkPropertyMetadata(False, New PropertyChangedCallback(AddressOf OnHandledDoubleClickChanged)))

        ''' <summary>
        ''' Get The HandleDoubleClick Property
        ''' </summary>
        Public Shared Function GetHandleDoubleClick(ByVal dp As DependencyObject) As Boolean
            Return CBool(dp.GetValue(DoubleClickProperty))
        End Function

        ''' <summary>
        ''' Set The HandleDoubleClick Property
        ''' </summary>
        Public Shared Sub SetHandleDoubleClick(ByVal dp As DependencyObject, ByVal value As Boolean)
            dp.SetValue(DoubleClickProperty, value)
        End Sub

        ''' <summary>
        ''' Hooks Up A Weak Event Against A Source's 'MouseDoubleClick'.
        ''' </summary>
        Private Shared Sub OnHandledDoubleClickChanged(ByVal sender As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim selector As Selector = TryCast(sender, Selector) 'TryCast

            If Not selector Is Nothing Then
                If CBool(e.NewValue) Then
                    RemoveHandler selector.MouseDoubleClick, AddressOf OnMouseDoubleClick
                    AddHandler selector.MouseDoubleClick, AddressOf OnMouseDoubleClick
                End If
            End If
        End Sub

        ''' <summary>
        ''' Handles MouseDoubleClick
        ''' </summary>
        Private Shared Sub OnMouseDoubleClick(ByVal sender As Object, e As MouseButtonEventArgs)
            ' Will Only Fire If The MouseDoubleClick Occurred Over An Actual
            ' ItemsControl Item. This Is nessecary In Case We May Have Clicked The Headers Which Are Not Items.

            Dim listView As ItemsControl = TryCast(sender, ItemsControl)

            Dim originalSender As DependencyObject = TryCast(e.OriginalSource, DependencyObject)
            If listView Is Nothing OrElse originalSender Is Nothing Then Return

            Dim container As DependencyObject = ItemsControl.ContainerFromElement(CType(sender, ItemsControl), CType(e.OriginalSource, DependencyObject))

            If container Is Nothing OrElse container.Equals(DependencyProperty.UnsetValue) Then Return

            ' We Have A Container At This Stage, Now Get The Item.
            Dim activatedItem As Object = listView.ItemContainerGenerator.ItemFromContainer(container)

            If Not activatedItem Is Nothing Then
                Dim s As DependencyObject = TryCast(sender, DependencyObject)
                Dim command As ICommand = s.GetValue(TheCommandToRunProperty)
                If Not command Is Nothing Then
                    If command.CanExecute(Nothing) Then
                        command.Execute(Nothing)
                    End If
                End If

            End If
        End Sub

#End Region

#Region "Command Dependancy Property Members"

        ''' <summary>
        ''' The Actual Command To Run.
        ''' </summary>
        Public Shared ReadOnly TheCommandToRunProperty As DependencyProperty = DependencyProperty.RegisterAttached _
                                                                            ("TheCommandToRun", _
                                                                            GetType(ICommand), _
                                                                            GetType(DoubleClickEventHelper), _
                                                                            New FrameworkPropertyMetadata(CType(Nothing, ICommand)))


        ''' <summary>
        ''' Gets The 'CommandToRunProperty'
        ''' </summary>
        Public Shared Function GetTheCommandToRun(ByVal dp As DependencyObject) As ICommand 'GetTheCommandToRun
            Return DirectCast(dp.GetValue(TheCommandToRunProperty), ICommand)
        End Function

        ''' <summary>
        ''' Set The CommandToRunProperty
        ''' </summary>
        Public Shared Sub SetTheCommandToRun(ByVal dp As DependencyObject, ByVal value As ICommand)
            dp.SetValue(TheCommandToRunProperty, value)
        End Sub

#End Region

    End Class

End Namespace
