Attribute VB_Name = "Módulo5"
Sub Generar_estado_situacion_financiera()

    Dim wsBalance As Worksheet
    Dim wsRCF As Worksheet
    Dim wsAnexo As Worksheet
    Dim wsBCInicial As Worksheet
    Dim ws2a As Worksheet
    Dim ws2b As Worksheet
    Dim celda As Range
    Dim celdaBuscar As Range
    Dim saldoTotal As Double
    Dim cuentas1 As Variant
    Dim cuentas2 As Variant
    Dim cuentas3 As Variant
    Dim cuentas4 As Variant
    Dim cuentas5 As Variant
    Dim cuentas6 As Variant
    Dim cuentas7 As Variant
    Dim cuentas8 As Variant
    Dim cuentas9 As Variant
    Dim i As Integer
    Dim rngOutput As Range
    Dim wsMenu As Worksheet
    Dim rangoBC As Range
    Dim total As Double
    Dim primerosDigitos As String

    ' Establecer la hoja de balance y la hoja de RCF
    Set wsBalance = ThisWorkbook.Sheets("BC")
    Set wsRCF = ThisWorkbook.Sheets("RCF")
    Set wsAnexo = ThisWorkbook.Sheets("Anexo N13")
    Set wsBCInicial = ThisWorkbook.Sheets("BC inicial")
    Set wsBaseID = ThisWorkbook.Sheets("Base impuesto Diferido")
    Set wsMenu = ThisWorkbook.Sheets("Menu")


    ' Definir los grupos de cuentas a buscar
    cuentas1 = Array(611011, 611012, 611021, 611022, 611031, 611032, 611041, 611042, 611051, 611052, 611061, 611062, 611071, 611072, 611081, 611091, 611092, 611101, 611111, 611112, 611121, 611122, 611131, 611132, 611151, 611152, 611161, 611162, 611163, 611171, 611172, 611181, 611182, 611191, 611192, 611201, 611202, 611211, 611212, 611221)
    cuentas2 = Array(611121)
    cuentas3 = Array(611141, 611142)
    cuentas4 = Array(711011, 711021, 711031, 711041, 711051, 711052, 711061, 711071, 711081, 711091, 711092, 711093, 711101, 712011, 712012, 712021, 712031, 713011, 713021, 713031, 713041, 713042, 713043, 713051, 713061, 713071, 713081, 713091, 713101, 713102, 713111, 713112, 713113, 713121, 713131, 713141, 713142, 713143, 713181, 713183, 713191, 713201, 713211, 713221, 713231, 713241, 713251, 713261, 713271, 713281, 713291, 713301, 713302, 713303, 713311, 713321, 713331, 713341, 713351, 713361, 713371, 713381, 713391, 713401, 713411, 713421, 713431, 713441, 713451, 713461, 713471, 713481, 713491, 713501, 713511, 713521, 713531, 713541, 713551, 713561, 713571, 713581, 713591, 713601, 713611, 713621, 713631, 713981, 713991, 714011)
    cuentas5 = Array(714011)
    cuentas6 = Array(612011)
    cuentas7 = Array(714013004)
    cuentas8 = Array(714013006)
    cuentas9 = Array(714013007)
    cuentas10 = Array(714013008)

    ' Definir las celdas de salida para cada grupo
    Set rngOutput1 = wsRCF.Range("D12")
    Set rngOutput2 = wsRCF.Range("D14")
    Set rngOutput3 = wsRCF.Range("D17")
    Set rngOutput4 = wsRCF.Range("G28")
    Set rngOutput5 = wsRCF.Range("G34")
    Set rngOutput6 = wsRCF.Range("G36")
    Set rngOutput7 = wsRCF.Range("I49")
    Set rngOutput8 = wsRCF.Range("I50")
    Set rngOutput9 = wsRCF.Range("I47")
    Set rngOutput10 = wsRCF.Range("I48")

    ' Procesar el primer grupo de cuentas
    saldoTotal = 0
    For Each celda In wsBalance.Range("A1:A" & wsBalance.Cells(wsBalance.Rows.Count, "A").End(xlUp).Row)
        For i = LBound(cuentas1) To UBound(cuentas1)
            If celda.Value = cuentas1(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 3).Value ' Cambio a Offset(0, 3) para la columna D
            End If
        Next i
    Next celda
    rngOutput1.Value = saldoTotal

    ' Procesar el segundo grupo de cuentas
    saldoTotal = 0
    For Each celda In wsBalance.Range("A1:A" & wsBalance.Cells(wsBalance.Rows.Count, "A").End(xlUp).Row)
        For i = LBound(cuentas2) To UBound(cuentas2)
            If celda.Value = cuentas2(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 3).Value ' Cambio a Offset(0, 3) para la columna D
            End If
        Next i
    Next celda
    rngOutput2.Value = saldoTotal

    ' Procesar el tercer grupo de cuentas
    saldoTotal = 0
    For Each celda In wsBalance.Range("A1:A" & wsBalance.Cells(wsBalance.Rows.Count, "A").End(xlUp).Row)
        For i = LBound(cuentas3) To UBound(cuentas3)
            If celda.Value = cuentas3(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 3).Value ' Cambio a Offset(0, 3) para la columna D
            End If
        Next i
    Next celda
    rngOutput3.Value = saldoTotal

    ' Procesar el cuarto grupo de cuentas
    saldoTotal = 0
    For Each celda In wsBalance.Range("A1:A" & wsBalance.Cells(wsBalance.Rows.Count, "A").End(xlUp).Row)
        For i = LBound(cuentas4) To UBound(cuentas4)
            If celda.Value = cuentas4(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 3).Value ' Cambio a Offset(0, 3) para la columna D
            End If
        Next i
    Next celda
    rngOutput4.Value = saldoTotal

        ' Procesar el quinta grupo de cuentas
    saldoTotal = 0
    For Each celda In wsBalance.Range("A1:A" & wsBalance.Cells(wsBalance.Rows.Count, "A").End(xlUp).Row)
        For i = LBound(cuentas5) To UBound(cuentas5)
            If celda.Value = cuentas5(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 3).Value ' Cambio a Offset(0, 3) para la columna D
            End If
        Next i
    Next celda

    ' Convertir el saldoTotal a su opuesto
    saldoTotal = saldoTotal * -1

    rngOutput5.Value = saldoTotal

            ' Procesar el sexta grupo de cuentas
    saldoTotal = 0
    For Each celda In wsBalance.Range("A1:A" & wsBalance.Cells(wsBalance.Rows.Count, "A").End(xlUp).Row)
        For i = LBound(cuentas6) To UBound(cuentas6)
            If celda.Value = cuentas6(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 3).Value ' Cambio a Offset(0, 3) para la columna D
            End If
        Next i
    Next celda
    rngOutput6.Value = saldoTotal


    ' Procesar APARTADO PARA IMPUESTO SOBRE LA RENTA

    saldoTotal = 0
    For Each celda In wsBalance.Range("B1:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)
        For i = LBound(cuentas7) To UBound(cuentas7)
            If celda.Value = cuentas7(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 2).Value ' Cambio a Offset(0, 2) para la columna D
            End If
        Next i
    Next celda
    rngOutput7.Value = saldoTotal


                    ' Procesar APARTADO PARA IMPUESTO SOBRE LA RENTA
    saldoTotal = 0
    For Each celda In wsBalance.Range("B1:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)
        For i = LBound(cuentas8) To UBound(cuentas8)
            If celda.Value = cuentas8(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 2).Value ' Cambio a Offset(0, 2) para la columna D
            End If
        Next i
    Next celda
    rngOutput8.Value = saldoTotal


    ' Procesar APTDO. P/IMP.S/RENTA (DIF.UT/SUSTITUTIVA)

    saldoTotal = 0
    For Each celda In wsBalance.Range("B1:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)
        For i = LBound(cuentas9) To UBound(cuentas9)
            If celda.Value = cuentas9(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 2).Value ' Cambio a Offset(0, 2) para la columna D
            End If
        Next i
    Next celda
    rngOutput9.Value = saldoTotal

     ' Procesar Apartado para Impuesto Sobre la Renta - Estimado

        saldoTotal = 0
    For Each celda In wsBalance.Range("B1:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)
        For i = LBound(cuentas10) To UBound(cuentas10)
            If celda.Value = cuentas10(i) Then
                saldoTotal = saldoTotal + celda.Offset(0, 2).Value ' Cambio a Offset(0, 2) para la columna D
            End If
        Next i
    Next celda
    rngOutput10.Value = saldoTotal



    '--------------------------- Anexo N13------------------------------------------

    ' Saldo de Tributos Pagados y No Causados y/o Tributos Causados en el Ejercicio Actual y Pagados en el Ejercicio Anterior

    For Each celda In wsAnexo.Range("A8:A9")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        For Each celda In wsAnexo.Range("A8:A9")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda


    ' Saldo de Tributos Causados y No Pagados y/o Tributos Pagados en el Ejercicio Actual y Causados en el Ejercicio Anterior

    For Each celda In wsAnexo.Range("A13:A21")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        For Each celda In wsAnexo.Range("A13:A21")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        ' Saldo de Diferencia en Cambio No Realizada y/o Realizada

    For Each celda In wsAnexo.Range("A25:A141")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        For Each celda In wsAnexo.Range("A24:A141")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Saldo de Remuneraciones Causadas y No Pagadas y/o Pagadas en el Ejercicio Actual y Causadas en el Ejercicio Anterior

    For Each celda In wsAnexo.Range("A146:A150")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        For Each celda In wsAnexo.Range("A146:A150")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        'Saldo de Arrendamientos Devengados y No Cobrados y/o Cobrados en el Ejercicio Actual y Devengados en el Ejercicio Anterior

    For Each celda In wsAnexo.Range("A154")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        For Each celda In wsAnexo.Range("A154")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda


        'Saldo Gastos de Administración Causados y No Pagados y/o Pagados en el Ejercicio Actual y Causados en el Ejercicio Anterior

    For Each celda In wsAnexo.Range("A158:A162")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        For Each celda In wsAnexo.Range("A158:A162")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Saldo Gastos de Conservación Causados y No Pagados y/o Pagados en el Ejercicio Actual y Causados en el Ejercicio Anterior


    For Each celda In wsAnexo.Range("A166:A170")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

        For Each celda In wsAnexo.Range("A166:A170")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

     'Gastos Generales No Deducibles

        For Each celda In wsAnexo.Range("A174:A201")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

            For Each celda In wsAnexo.Range("A174:A200")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 5).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 5).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

         'Reservas

        For Each celda In wsAnexo.Range("A206:A217")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

            For Each celda In wsAnexo.Range("A206:A217")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 5).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 5).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

             'Liberaciones de Reservas

        For Each celda In wsAnexo.Range("A221:A224")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 4).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 4).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

            For Each celda In wsAnexo.Range("A221:A224")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 5).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13

     Else
         celda.Offset(0, 5).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    ' Tributos Pagados y No Causados y/o Causados en el Ej. Actual y Pagados en el Ej. Anterior


    wsRCF.Range("G30").Value = wsAnexo.Range("F22").Value
    Debug.Print "Valor copiado de Anexo n13 F143 a RCF G29: " & wsAnexo.Range("F143").Value

    ' Diferencial Cambiario No Realizado o (Realizado)

    wsRCF.Range("G20").Value = -wsAnexo.Range("F143").Value

    ' Gastos No Deducibles, según Anexo N°13

    wsRCF.Range("G29").Value = -wsAnexo.Range("F203").Value

    ' Remuneraciones Causadas y No Pagadas y/o Pagadas en el Ejercicio Actual y Causadas en el Ejercicio Anterior, según Anexo N°13

    wsRCF.Range("G32").Value = -wsAnexo.Range("F151").Value


    ' ----------------------------- Base de impuesto diferido ------------------------------------

    'Saldo de Tributos Pagados y No Causados y/o Tributos Causados en el Ejercicio Actual y Pagados en el Ejercicio Anterior


            For Each celda In wsBaseID.Range("A8:A9")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Saldo de Tributos Causados y No Pagados y/o Tributos Pagados en el Ejercicio Actual y Causados en el Ejercicio Anterior

              For Each celda In wsBaseID.Range("A13:A21")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Saldo de Diferencia en Cambio No Realizada y/o Realizada

                  For Each celda In wsBaseID.Range("A25:A142")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Saldo de Remuneraciones Causadas y No Pagadas y/o Pagadas en el Ejercicio Actual y Causadas en el Ejercicio Anterior

                  For Each celda In wsBaseID.Range("A146:A150")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Saldo de Arrendamientos Devengados y No Cobrados y/o Cobrados en el Ejercicio Actual y Devengados en el Ejercicio Anterior

                  For Each celda In wsBaseID.Range("A154")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Saldo Gastos de Administración Causados y No Pagados y/o Pagados en el Ejercicio Actual y Causados en el Ejercicio Anterior

                  For Each celda In wsBaseID.Range("A158:A162")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Saldo Gastos de Conservación Causados y No Pagados y/o Pagados en el Ejercicio Actual y Causados en el Ejercicio Anterior


                      For Each celda In wsBaseID.Range("A166:A170")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Provisiones

                      For Each celda In wsBaseID.Range("A174:A177")
     Set celdaBuscar = wsBalance.Range("B:B").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda

    'Impuesto Diferido Activo (IDA)

                      For Each celda In wsBaseID.Range("A196")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda


        'Impuesto Diferido Pasivo (IDP)


                      For Each celda In wsBaseID.Range("A201")
     Set celdaBuscar = wsBCInicial.Range("A:A").Find(What:=celda.Value, LookIn:=xlValues, LookAt:=xlWhole)
     If Not celdaBuscar Is Nothing Then
         celda.Offset(0, 3).Value = celdaBuscar.Offset(0, 2).Value ' Copiar el valor de la columna C en BC inicial a la columna D en Anexo n13
     Else
         celda.Offset(0, 3).Value = 0 ' Si no se encuentra la cuenta, establecer el valor a cero
        End If
    Next celda


    '--------------------------------------- Tools-----------------------------------------------------------

        ' LOCTI

    wsMenu.Range("M6").Value = Abs(wsRCF.Range("D12").Value * 0.005)

    With wsBalance
        ' Encontrar la última fila con datos en la columna B
        ultimaFila = .Cells(.Rows.Count, "B").End(xlUp).Row

        ' Buscar la cuenta en el rango especificado
        Set celdaBuscar = .Range("B2:B" & ultimaFila).Find(What:=713571001, LookIn:=xlValues, LookAt:=xlWhole)

        ' Si se encuentra la cuenta, copiar el valor a la hoja wsMenu
        If Not celdaBuscar Is Nothing Then
            wsMenu.Range("N6").Value = celdaBuscar.Offset(0, 2).Value ' Offset(0,2) para ir 2 columnas a la derecha
        Else
        End If
    End With

            ' Impuesto Diferido

    wsMenu.Range("M7").Value = Abs(wsBaseID.Range("E204").Value * 1)

    With wsBalance
        ' Encontrar la última fila con datos en la columna B
        ultimaFila = .Cells(.Rows.Count, "B").End(xlUp).Row

        ' Buscar la cuenta en el rango especificado
        Set celdaBuscar = .Range("B2:B" & ultimaFila).Find(What:=714013006, LookIn:=xlValues, LookAt:=xlWhole)

        ' Si se encuentra la cuenta, copiar el valor a la hoja wsMenu
        If Not celdaBuscar Is Nothing Then
            wsMenu.Range("N7").Value = celdaBuscar.Offset(0, 2).Value ' Offset(0,2) para ir 2 columnas a la derecha
        Else
        End If
    End With


        ' ISLR

    wsMenu.Range("M8").Value = Abs(wsRCF.Range("H59").Value * 0.34)

    With wsBalance
        ' Encontrar la última fila con datos en la columna B
        ultimaFila = .Cells(.Rows.Count, "B").End(xlUp).Row

        ' Buscar la cuenta en el rango especificado
        Set celdaBuscar = .Range("B2:B" & ultimaFila).Find(What:=213011010, LookIn:=xlValues, LookAt:=xlWhole)

        ' Si se encuentra la cuenta, copiar el valor a la hoja wsMenu
        If Not celdaBuscar Is Nothing Then
            wsMenu.Range("N8").Value = celdaBuscar.Offset(0, 2).Value ' Offset(0,2) para ir 2 columnas a la derecha
        Else
        End If
    End With

        ' Ley de deporte

    If wsRCF.Range("J51").Value > 0 Then
        wsMenu.Range("M9").Value = 0
    Else
        wsMenu.Range("M9").Value = Abs(wsRCF.Range("J51").Value * 0.01)
    End If

    With wsBalance
        ' Encontrar la última fila con datos en la columna B
        ultimaFila = .Cells(.Rows.Count, "B").End(xlUp).Row

        ' Buscar la cuenta en el rango especificado
        Set celdaBuscar = .Range("B2:B" & ultimaFila).Find(What:=714011046, LookIn:=xlValues, LookAt:=xlWhole)

        ' Si se encuentra la cuenta, copiar el valor a la hoja wsMenu
        If Not celdaBuscar Is Nothing Then
            wsMenu.Range("N9").Value = celdaBuscar.Offset(0, 2).Value ' Offset(0,2) para ir 2 columnas a la derecha
        Else
        End If

    End With


    ' 711 - Gastos de personal

    ' Configurar hojas
    Set wsBalance = ThisWorkbook.Worksheets("BC")
    Set wsMenu = ThisWorkbook.Worksheets("Menu")

    ' Inicializar total
    total = 0

    ' Definir rango de datos en columna B (ajusta el rango según tus datos)
    Set rangoBC = wsBalance.Range("B2:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)

    ' Recorrer cada celda en columna B
    For Each cell In rangoBC
        If Len(cell.Value) >= 3 Then
            primerosDigitos = Left(cell.Value, 3)

            ' Verificar si los dígitos son 11 o 12
            If primerosDigitos = "711" Then
                ' Sumar el valor de columna D en la misma fila
                total = total + cell.Offset(0, 2).Value ' D es 2 columnas a la derecha de B
            End If
        End If
    Next cell

    ' Colocar resultado en M16 de la hoja Menu
    wsMenu.Range("M17").Value = total


    ' 712 - Gastos financiero

    ' Configurar hojas
    Set wsBalance = ThisWorkbook.Worksheets("BC")
    Set wsMenu = ThisWorkbook.Worksheets("Menu")

    ' Inicializar total
    total = 0

    ' Definir rango de datos en columna B (ajusta el rango según tus datos)
    Set rangoBC = wsBalance.Range("B2:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)

    ' Recorrer cada celda en columna B
    For Each cell In rangoBC
        If Len(cell.Value) >= 3 Then
            primerosDigitos = Left(cell.Value, 3)

            ' Verificar si los dígitos son 11 o 12
            If primerosDigitos = "712" Then
                ' Sumar el valor de columna D en la misma fila
                total = total + cell.Offset(0, 2).Value ' D es 2 columnas a la derecha de B
            End If
        End If
    Next cell

    ' Colocar resultado en M16 de la hoja Menu
    wsMenu.Range("M18").Value = total



    ' 713 -  Otros gastos financieros



    ' Configurar hojas
    Set wsBalance = ThisWorkbook.Worksheets("BC")
    Set wsMenu = ThisWorkbook.Worksheets("Menu")

    ' Inicializar total
    total = 0

    ' Definir rango de datos en columna B (ajusta el rango según tus datos)
    Set rangoBC = wsBalance.Range("B2:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)

    ' Recorrer cada celda en columna B
    For Each cell In rangoBC
        If Len(cell.Value) >= 3 Then
            primerosDigitos = Left(cell.Value, 3)

            ' Verificar si los dígitos son 11 o 12
            If primerosDigitos = "713" Then
                ' Sumar el valor de columna D en la misma fila
                total = total + cell.Offset(0, 2).Value ' D es 2 columnas a la derecha de B
            End If
        End If
    Next cell

    ' Colocar resultado en M16 de la hoja Menu
    wsMenu.Range("M19").Value = total


    ' 714 - Apartados y reservas


    ' Configurar hojas
    Set wsBalance = ThisWorkbook.Worksheets("BC")
    Set wsMenu = ThisWorkbook.Worksheets("Menu")

    ' Inicializar total
    total = 0

    ' Definir rango de datos en columna B (ajusta el rango según tus datos)
    Set rangoBC = wsBalance.Range("B2:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)

    ' Recorrer cada celda en columna B
    For Each cell In rangoBC
        If Len(cell.Value) >= 3 Then
            primerosDigitos = Left(cell.Value, 3)

            ' Verificar si los dígitos son 11 o 12
            If primerosDigitos = "714" Then
                ' Sumar el valor de columna D en la misma fila
                total = total + cell.Offset(0, 2).Value ' D es 2 columnas a la derecha de B
            End If
        End If
    Next cell

    ' Colocar resultado en M16 de la hoja Menu
    wsMenu.Range("M20").Value = total


    ' 611 - Otros ingresos Mercantiles


    ' Configurar hojas
    Set wsBalance = ThisWorkbook.Worksheets("BC")
    Set wsMenu = ThisWorkbook.Worksheets("Menu")

    ' Inicializar total
    total = 0

    ' Definir rango de datos en columna B (ajusta el rango según tus datos)
    Set rangoBC = wsBalance.Range("B2:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)

    ' Recorrer cada celda en columna B
    For Each cell In rangoBC
        If Len(cell.Value) >= 3 Then
            primerosDigitos = Left(cell.Value, 3)

            ' Verificar si los dígitos son 11 o 12
            If primerosDigitos = "611" Then
                ' Sumar el valor de columna D en la misma fila
                total = total + cell.Offset(0, 2).Value ' D es 2 columnas a la derecha de B
            End If
        End If
    Next cell

    ' Colocar resultado en M16 de la hoja Menu
    wsMenu.Range("M24").Value = total


    ' 612 - Liberaciones de Revervas


    ' Configurar hojas
    Set wsBalance = ThisWorkbook.Worksheets("BC")
    Set wsMenu = ThisWorkbook.Worksheets("Menu")

    ' Inicializar total
    total = 0

    ' Definir rango de datos en columna B (ajusta el rango según tus datos)
    Set rangoBC = wsBalance.Range("B2:B" & wsBalance.Cells(wsBalance.Rows.Count, "B").End(xlUp).Row)

    ' Recorrer cada celda en columna B
    For Each cell In rangoBC
        If Len(cell.Value) >= 3 Then
            primerosDigitos = Left(cell.Value, 3)

            ' Verificar si los dígitos son 11 o 12
            If primerosDigitos = "612" Then
                ' Sumar el valor de columna D en la misma fila
                total = total + cell.Offset(0, 2).Value ' D es 2 columnas a la derecha de B
            End If
        End If
    Next cell

    ' Colocar resultado en M16 de la hoja Menu
    wsMenu.Range("M25").Value = total

    ' Mensaje de confirmación
    MsgBox "Los montos han sido transferidos correctamente."


End Sub
