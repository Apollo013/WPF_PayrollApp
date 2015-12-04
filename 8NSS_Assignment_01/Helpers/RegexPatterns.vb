Namespace Helpers

    Public NotInheritable Class RegexPatterns

        Public Shared Function DoublePrecision2() As String
            Return "^([1-9]{1}[\d]{0,2}(\,[\d]{3})*(\.[\d]{0,2})?|[1-9]{1}[\d]{0,}(\.[\d]{0,2})?|0(\.[\d]{0,2})?|(\.[\d]{1,2})?)$"

            ' Matches       132,123,123.23 | 123456.23  | 0.23
            ' Non Matches           123,12 | 123.123    | 1322,132.23
        End Function

    End Class

End Namespace
