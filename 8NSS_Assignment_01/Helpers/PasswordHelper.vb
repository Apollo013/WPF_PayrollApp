Namespace Helpers.Login

    Public NotInheritable Class PasswordHelper

#Region "Comments"

        ' Thanks To Mark Henrikson For This. http://mhenrikson.blogspot.com/ 

        ' Password Controls Cannot Be Directly Binded To Which Means That We Cannot Validate Unless We Push
        ' The Value Back Directly To The Server DataBase (Which Is Fine When You're Using a db). 
        ' To Address This, This Class Creates A Dependancy Property That Can Be Attached To The Password Control.
        ' This Allows Us To Nominate Another Control (TextBlock or TextBox) That Will Reflect What Has Been Typed 
        ' Into The Password Control By The User, And We Bind To This Control Instead.

        ' This Work Around Is Fine Where Security Is Not A Top Priority.

#End Region

#Region "Dependancy Objects"

        Public Shared ReadOnly PasswordProperty As DependencyProperty = DependencyProperty.RegisterAttached("Password", GetType(String), GetType(PasswordHelper), New FrameworkPropertyMetadata(String.Empty, New PropertyChangedCallback(AddressOf OnPasswordPropertyChanged)))

        Public Shared ReadOnly AttachProperty As DependencyProperty = DependencyProperty.RegisterAttached("Attach", GetType(Boolean), GetType(PasswordHelper), New PropertyMetadata(False, New PropertyChangedCallback(AddressOf Attach)))

        Private Shared ReadOnly IsUpdatingProperty As DependencyProperty = DependencyProperty.RegisterAttached("IsUpdating", GetType(Boolean), GetType(PasswordHelper))

#End Region

#Region "Methods"

        Private Sub New()
        End Sub

        Public Shared Sub SetAttach(ByVal dp As DependencyObject, ByVal value As Boolean)
            dp.SetValue(AttachProperty, value)
        End Sub

        Public Shared Function GetAttach(ByVal dp As DependencyObject) As Boolean
            Return CBool(dp.GetValue(AttachProperty))
        End Function

        Public Shared Function GetPassword(ByVal dp As DependencyObject) As String
            Return DirectCast(dp.GetValue(PasswordProperty), String)
        End Function

        Public Shared Sub SetPassword(ByVal dp As DependencyObject, ByVal value As String)
            dp.SetValue(PasswordProperty, value)
        End Sub

        Private Shared Function GetIsUpdating(ByVal dp As DependencyObject) As Boolean
            Return CBool(dp.GetValue(IsUpdatingProperty))
        End Function

        Private Shared Sub SetIsUpdating(ByVal dp As DependencyObject, ByVal value As Boolean)
            dp.SetValue(IsUpdatingProperty, value)
        End Sub

        Private Shared Sub OnPasswordPropertyChanged(ByVal sender As System.Windows.DependencyObject, ByVal e As System.Windows.DependencyPropertyChangedEventArgs)
            Dim passwordBox As PasswordBox = TryCast(sender, PasswordBox)
            RemoveHandler passwordBox.PasswordChanged, AddressOf PasswordChanged

            If Not CBool(GetIsUpdating(passwordBox)) Then
                passwordBox.Password = DirectCast(e.NewValue, String)
            End If
            AddHandler passwordBox.PasswordChanged, AddressOf PasswordChanged
        End Sub

        Private Shared Sub Attach(ByVal sender As DependencyObject, ByVal e As DependencyPropertyChangedEventArgs)
            Dim passwordBox As PasswordBox = TryCast(sender, PasswordBox)

            If passwordBox Is Nothing Then
                Return
            End If

            If CBool(e.OldValue) Then
                RemoveHandler passwordBox.PasswordChanged, AddressOf PasswordChanged
            End If

            If CBool(e.NewValue) Then
                AddHandler passwordBox.PasswordChanged, AddressOf PasswordChanged
            End If
        End Sub

        Private Shared Sub PasswordChanged(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim passwordBox As PasswordBox = TryCast(sender, PasswordBox)
            SetIsUpdating(passwordBox, True)
            SetPassword(passwordBox, passwordBox.Password)
            SetIsUpdating(passwordBox, False)
        End Sub

#End Region

    End Class

End Namespace
