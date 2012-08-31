Imports System
Imports System.Collections.Generic
Imports System.Text
Imports System.IO
Imports System.Security.Cryptography.X509Certificates
Imports System.Drawing
'EPPlus Namespaces
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports OfficeOpenXml.Drawing.Chart

Module modMain

    Sub Main()

        Dim args() As String
        Dim file As FileInfo = Nothing

        args = Environment.GetCommandLineArgs

        If UBound(args) = 0 Then
            GoTo InvalidFile
        End If

        file = New FileInfo(args(1))

        If file.Exists Then
            If file.Extension <> ".xlsm" Then
                GoTo InvalidFile
            End If
        Else
            Console.WriteLine("File supplied does not exist... Exiting.")
            Exit Sub
        End If

        ExtractCode(file)

        Exit Sub

InvalidFile:

        Console.WriteLine("Must Supply Excel Macro-Enabled Workbook (*.xlsm)")
        Console.WriteLine("ex: ExtractVBA.exe Sample.xlsm")

    End Sub

    Sub ExtractCode(ByRef file As FileInfo)

        Dim pck As ExcelPackage
        Dim wkbk As OfficeOpenXml.ExcelWorkbook
        Dim vba As OfficeOpenXml.VBA.ExcelVbaProject
        '
        Dim ref As OfficeOpenXml.VBA.ExcelVbaReference
        Dim code As OfficeOpenXml.VBA.ExcelVBAModule
        Dim type As OfficeOpenXml.VBA.eModuleType
        Dim attr() As String
        Dim attr_name(4) As String
        Dim source As String

        Dim path As String = ""

        ReDim attr(0)

        Dim sb As StringBuilder
        Dim name As String
        Dim filename As String = ""

        pck = New ExcelPackage(file)
        wkbk = pck.Workbook
        vba = pck.Workbook.VbaProject

        Dim idx As Integer
        Dim idx2 As Integer
        Dim attr_idx As Integer

        path = Environment.CurrentDirectory & "\" & file.Name & ".vba"

        attr_name(0) = "VB_Name"
        attr_name(1) = "VB_GlobalNameSpace"
        attr_name(2) = "VB_Creatable"
        attr_name(3) = "VB_PredeclaredId"
        attr_name(4) = "VB_Exposed"

        If Not Directory.Exists(path) Then
            Directory.CreateDirectory(path)
        End If

        For idx = 0 To vba.Modules.Count - 1
            code = vba.Modules(idx)
            name = code.Name
            type = code.Type
            source = code.Code

            attr_idx = 0

            For idx2 = 0 To code.Attributes.Count - 1

                If Array.IndexOf(attr_name, code.Attributes(idx2).Name) >= 0 Then

                    ReDim Preserve attr(attr_idx)

                    attr(attr_idx) = "Attribute"
                    attr(attr_idx) += " "
                    attr(attr_idx) += code.Attributes(idx2).Name
                    attr(attr_idx) += " = "
                    attr(attr_idx) += code.Attributes(idx2).Value

                    attr_idx += 1

                End If

            Next

            If source <> "" Then

                sb = New StringBuilder()

                Select Case type
                    Case OfficeOpenXml.VBA.eModuleType.Document, OfficeOpenXml.VBA.eModuleType.Class

                        sb.AppendLine("VERSION 1.0 CLASS")
                        sb.AppendLine("BEGIN")
                        sb.AppendLine("  MultiUse = -1  'True")
                        sb.AppendLine("END")

                        For attr_idx = 0 To UBound(attr)
                            sb.AppendLine(attr(attr_idx))
                        Next

                        sb.Append(source)

                        filename = name & ".cls"

                    Case OfficeOpenXml.VBA.eModuleType.Module

                        For attr_idx = 0 To UBound(attr)
                            sb.AppendLine(attr(attr_idx))
                        Next

                        sb.Append(source)

                        filename = name & ".bas"
                    Case OfficeOpenXml.VBA.eModuleType.Designer

                        For attr_idx = 0 To UBound(attr)
                            sb.AppendLine(attr(attr_idx))
                        Next

                        sb.Append(source)

                        filename = name & ".txt"
                End Select


                Using outfile As New StreamWriter(path & "\" & filename)
                    outfile.Write(sb.ToString())
                End Using

                sb = Nothing

            End If

        Next

        sb = New StringBuilder()


        For idx = 0 To vba.References.Count - 1
            ref = vba.References(idx)
            sb.AppendLine(Mid(ref.Libid, LastPos(ref.Libid, "#")))
            sb.AppendLine(ref.Libid)
            sb.AppendLine("------------------------------------------------------------------------")
        Next

        Using outfile As New StreamWriter(path & "\references.txt")
            outfile.Write(sb.ToString())
        End Using

        sb = Nothing

    End Sub

End Module
