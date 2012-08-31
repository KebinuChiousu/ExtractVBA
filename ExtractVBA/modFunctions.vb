Module modFunctions

    Function LastPos(ByVal source As String, ByVal search As String) As Integer

        Dim ret As Integer
        Dim idx As Integer = 1
        Dim check As Integer = 1

        Do Until check = 0

            check = InStr(idx, source, search)

            If check > idx Then
                ret = check
            End If

            idx = check + 1

            If check + 1 > source.Length Then
                Exit Do
            End If

        Loop

        ret += 1

        Return ret

    End Function

End Module
