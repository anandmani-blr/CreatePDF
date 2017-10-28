Imports Spire.Xls
Imports Spire.Pdf
Imports Spire.Xls.Converter
Public Class PDFClass
    Public Function fnCreatePDF(ByVal strKey As String, ByVal strSourcePath As String, ByRef strDestPath As String) As Boolean
        'strSourcePath = replace(strSourcePath,"")
        'Anand Added new code to test
        Try
            If LCase(strKey) <> "" Then
                'Changed to key
                If UCase(strKey) = "VS1961A" Then
                    fnCreatePDF = False
                    Dim workbook As New Workbook()
                    workbook.LoadFromFile(strSourcePath)
                    strDestPath = strSourcePath
                    strDestPath = Replace(strDestPath, ".xls", ".pdf")
                    strDestPath = Replace(strDestPath, ".xlsx", ".pdf")
                    workbook.SaveToFile(strDestPath, Spire.Xls.FileFormat.PDF)
                    fnCreatePDF = True
                Else
                    fnCreatePDF = False
                    strDestPath = "Incorrect Key"
                End If
            Else
                fnCreatePDF = False
                strDestPath = Err.Description
            End If
        Catch EX As Exception
            fnCreatePDF = False
            strDestPath = Err.Description

        End Try
    End Function
End Class
