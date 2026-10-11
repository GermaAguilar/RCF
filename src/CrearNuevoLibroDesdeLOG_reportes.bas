Attribute VB_Name = "Módulo1"
Sub CrearNuevoLibroDesdeLOG_reportes()
    Dim wbOrigen As Workbook
    Dim wbNuevo As Workbook
    Dim strRuta As String

    ' Obtener la ruta del libro origen (asumiendo que está abierto)
    Set wbOrigen = ThisWorkbook

    ' Obtener la ruta de la carpeta de descargas
    strRuta = Environ("USERPROFILE") & "\Downloads\"

    ' Crear el nuevo libro y eliminar la hoja 1
    Set wbNuevo = Workbooks.Add

    ' Copiar las hojas especificadas
    With wbOrigen
        .Sheets("RCF").Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
        .Sheets("Anexo N13").Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
        .Sheets("Base Impuesto Diferido").Copy After:=wbNuevo.Sheets(wbNuevo.Sheets.Count)
    End With

    ' Guardar el libro en la carpeta de descargas
    wbNuevo.SaveAs Filename:=strRuta & "Extracto_LOG_reportes.xlsx"
    wbNuevo.Close SaveChanges:=False
End Sub
