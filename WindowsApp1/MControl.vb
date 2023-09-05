Imports Excel = Microsoft.Office.Interop.Excel
Imports Word = Microsoft.Office.Interop.Word

Module MControl
    Public AppExcel As Excel.Application
    Public Libro As Excel.Workbook
    Public Hoja As Excel.Worksheet
    Dim wordApp As New Word.Application
    Dim doc As Word.Document
    Public Sub Crear()
        AppExcel = CreateObject("Excel.Application")
        Libro = AppExcel.Workbooks.Add
        Hoja = Libro.ActiveSheet
        AppExcel.Visible = True
        AppExcel.Application.DisplayAlerts = False
        Libro.AppExcel.Workbooks.Open(Filename:="C:\Users\Documents.Xlsx")
        Hoja.Libro.activesheet
        Libro.Save()
    End Sub


    Public Sub Guardarcomo()
        Try
            If Not FileIO.FileSystem.DirectoryExists("C:\Users\Fancisco\Desktop\Tecnologia de la Informacion") Then
                MkDir("C:\Users\Fancisco\Desktop\Tecnologia de la Informacion")
            End If

            Libro.SaveAs(Filename:="C:\Users\Fancisco.Xlsx")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub


    Public Sub Cerrar()
        Try
            Libro.Close()
            AppExcel = Nothing
            Libro = Nothing
            Hoja = Nothing
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub



    Private Sub CrearWord(sender As Object, e As EventArgs)
        wordApp.Visible = True
        doc = wordApp.Documents.Add()
    End Sub
End Module
