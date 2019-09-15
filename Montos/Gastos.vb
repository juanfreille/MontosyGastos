
Imports System.Drawing.Printing

Public Class Gastos
    Dim MatrizGastos(25, 4) As Decimal
    Dim TOTAL_SERVICIO, TOTAL_CHIP, TOTAL_CORREO, TOTAL_FORMULARIO, TOTAL_AUTONOMO, TOTAL_ALQUILER, TOTAL_OBLEAS, TOTAL_LUZ, TOTAL_INTERNET, TOTAL_TELECOM As Decimal
    Dim TOTAL_AERPC, TOTAL_ACARA, TOTAL_SUELDOS, TOTAL_MINISTERIO, TOTAL_MANTENIMIENTO, TOTAL_INSUMOS, TOTAL_CONTADOR, TOTAL_VARIOS1, TOTAL_VARIOS2 As Decimal
    Dim TOTAL_VARIOS3, TOTAL_VARIOS4, TOTAL_GASTOS1, TOTAL_GASTOS2, TOTAL_GASTOS3, TOTAL_GASTOS4, TOTAL_GASTOS5 As Decimal
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim f As Integer
        Dim c As Integer

        If MessageBox.Show("¿Borrar todos los campos y empezar de nuevo?", "Mensaje de sistema", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Limpiarcampos(Me)
            REINICIAR_VARIABLES()
            For f = 0 To 25
                For c = 0 To 4
                    MatrizGastos(f, c) = 0
                Next
            Next
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim EXCEL As New Microsoft.Office.Interop.Excel.Application
        EXCEL.Visible = True
        EXCEL.Workbooks.Add()

        'COLOR DE FILAS
        With EXCEL.Range("A3:A28")
            .Interior.ColorIndex = 15
            .WrapText = True
            .Font.FontStyle = "Verdana"
            .Font.Bold = True
        End With
        With EXCEL.Range("A1")
            .Interior.ColorIndex = 17
            .Font.FontStyle = "Verdana"
            .Font.Underline = True
            .Font.Bold = True
        End With
        ' COLOR DE TOTALES
        With EXCEL.Range("H2")
            .Interior.ColorIndex = 10
            .WrapText = True
            .Font.FontStyle = "Verdana"
            .Font.Bold = True
        End With
        With EXCEL.Range("H30")
            .Interior.ColorIndex = 10
            .WrapText = True
            .Font.FontStyle = "Verdana"
            .Font.Bold = True
        End With
        With EXCEL.Range("A30")
            .Interior.ColorIndex = 10
            .WrapText = True
            .Font.FontStyle = "Verdana"
            .Font.Bold = True
        End With

        'COLOR DE COLUMNAS

        'NEGRITA PARA TOTALES
        With EXCEL.Range("H3:H30")
            .Font.Bold = True
        End With

        'FILAS
        EXCEL.Range("A1").Value = "PLANILLA DE GASTOS"
        EXCEL.Range("A1").ColumnWidth = 25
        EXCEL.Range("H1").ColumnWidth = 20
        EXCEL.Range("J1").ColumnWidth = 20
        EXCEL.Range("K1").ColumnWidth = 20
        EXCEL.Range("L1").ColumnWidth = 20
        EXCEL.Range("H2").Value = "TOTAL"
        EXCEL.Range("A30").Value = "SUMA TOTAL"
        EXCEL.Range("A3").Value = "Servicio Técnico"
        EXCEL.Range("A4").Value = "Chip Celular"
        EXCEL.Range("A5").Value = "Correo"
        EXCEL.Range("A6").Value = "Formulario"
        EXCEL.Range("A7").Value = "Autónomo"
        EXCEL.Range("A8").Value = "Alquiler"
        EXCEL.Range("A9").Value = "Obleas"
        EXCEL.Range("A10").Value = "Luz"
        EXCEL.Range("A11").Value = "Internet"
        EXCEL.Range("A12").Value = "Telecom"
        EXCEL.Range("A13").Value = "Aerpc Socios"
        EXCEL.Range("A14").Value = "Acara"
        EXCEL.Range("A15").Value = "Sueldos"
        EXCEL.Range("A16").Value = "Ministerio"
        EXCEL.Range("A17").Value = "Mantenimiento de cuenta"
        EXCEL.Range("A18").Value = "Insumos"
        EXCEL.Range("A19").Value = "Contador"
        EXCEL.Range("A20").Value = "Varios 1"
        EXCEL.Range("A21").Value = "Varios 2"
        EXCEL.Range("A22").Value = "Varios 3"
        EXCEL.Range("A23").Value = "Varios 4"
        EXCEL.Range("A24").Value = "Gastos Extraordinarios 1"
        EXCEL.Range("A25").Value = "Gastos Extraordinarios 2"
        EXCEL.Range("A26").Value = "Gastos Extraordinarios 3"
        EXCEL.Range("A27").Value = "Gastos Extraordinarios 4"
        EXCEL.Range("A28").Value = "Gastos Extraordinarios 5"
        'COLUMNAS


        'DATOS LUNES
        EXCEL.Range("B3").Value = MatrizGastos(0, 0)
        EXCEL.Range("B4").Value = MatrizGastos(1, 0)
        EXCEL.Range("B5").Value = MatrizGastos(2, 0)
        EXCEL.Range("B6").Value = MatrizGastos(3, 0)
        EXCEL.Range("B7").Value = MatrizGastos(4, 0)
        EXCEL.Range("B8").Value = MatrizGastos(5, 0)
        EXCEL.Range("B9").Value = MatrizGastos(6, 0)
        EXCEL.Range("B10").Value = MatrizGastos(7, 0)
        EXCEL.Range("B11").Value = MatrizGastos(8, 0)
        EXCEL.Range("B12").Value = MatrizGastos(9, 0)
        EXCEL.Range("B13").Value = MatrizGastos(10, 0)
        EXCEL.Range("B14").Value = MatrizGastos(11, 0)
        EXCEL.Range("B15").Value = MatrizGastos(12, 0)
        EXCEL.Range("B16").Value = MatrizGastos(13, 0)
        EXCEL.Range("B17").Value = MatrizGastos(14, 0)
        EXCEL.Range("B18").Value = MatrizGastos(15, 0)
        EXCEL.Range("B19").Value = MatrizGastos(16, 0)
        EXCEL.Range("B20").Value = MatrizGastos(17, 0)
        EXCEL.Range("B21").Value = MatrizGastos(18, 0)
        EXCEL.Range("B22").Value = MatrizGastos(19, 0)
        EXCEL.Range("B23").Value = MatrizGastos(20, 0)
        EXCEL.Range("B24").Value = MatrizGastos(21, 0)
        EXCEL.Range("B25").Value = MatrizGastos(22, 0)
        EXCEL.Range("B26").Value = MatrizGastos(23, 0)
        EXCEL.Range("B27").Value = MatrizGastos(24, 0)
        EXCEL.Range("B28").Value = MatrizGastos(25, 0)
        'DATOS MARTES
        EXCEL.Range("C3").Value = MatrizGastos(0, 1)
        EXCEL.Range("C4").Value = MatrizGastos(1, 1)
        EXCEL.Range("C5").Value = MatrizGastos(2, 1)
        EXCEL.Range("C6").Value = MatrizGastos(3, 1)
        EXCEL.Range("C7").Value = MatrizGastos(4, 1)
        EXCEL.Range("C8").Value = MatrizGastos(5, 1)
        EXCEL.Range("C9").Value = MatrizGastos(6, 1)
        EXCEL.Range("C10").Value = MatrizGastos(7, 1)
        EXCEL.Range("C11").Value = MatrizGastos(8, 1)
        EXCEL.Range("C12").Value = MatrizGastos(9, 1)
        EXCEL.Range("C13").Value = MatrizGastos(10, 1)
        EXCEL.Range("C14").Value = MatrizGastos(11, 1)
        EXCEL.Range("C15").Value = MatrizGastos(12, 1)
        EXCEL.Range("C16").Value = MatrizGastos(13, 1)
        EXCEL.Range("C17").Value = MatrizGastos(14, 1)
        EXCEL.Range("C18").Value = MatrizGastos(15, 1)
        EXCEL.Range("C19").Value = MatrizGastos(16, 1)
        EXCEL.Range("C20").Value = MatrizGastos(17, 1)
        EXCEL.Range("C21").Value = MatrizGastos(18, 1)
        EXCEL.Range("C22").Value = MatrizGastos(19, 1)
        EXCEL.Range("C23").Value = MatrizGastos(20, 1)
        EXCEL.Range("C24").Value = MatrizGastos(21, 1)
        EXCEL.Range("C25").Value = MatrizGastos(22, 1)
        EXCEL.Range("C26").Value = MatrizGastos(23, 1)
        EXCEL.Range("C27").Value = MatrizGastos(24, 1)
        EXCEL.Range("C28").Value = MatrizGastos(25, 1)

        'DATOS MIERCOLES
        EXCEL.Range("D3").Value = MatrizGastos(0, 2)
        EXCEL.Range("D4").Value = MatrizGastos(1, 2)
        EXCEL.Range("D5").Value = MatrizGastos(2, 2)
        EXCEL.Range("D6").Value = MatrizGastos(3, 2)
        EXCEL.Range("D7").Value = MatrizGastos(4, 2)
        EXCEL.Range("D8").Value = MatrizGastos(5, 2)
        EXCEL.Range("D9").Value = MatrizGastos(6, 2)
        EXCEL.Range("D10").Value = MatrizGastos(7, 2)
        EXCEL.Range("D11").Value = MatrizGastos(8, 2)
        EXCEL.Range("D12").Value = MatrizGastos(9, 2)
        EXCEL.Range("D13").Value = MatrizGastos(10, 2)
        EXCEL.Range("D14").Value = MatrizGastos(11, 2)
        EXCEL.Range("D15").Value = MatrizGastos(12, 2)
        EXCEL.Range("D16").Value = MatrizGastos(13, 2)
        EXCEL.Range("D17").Value = MatrizGastos(14, 2)
        EXCEL.Range("D18").Value = MatrizGastos(15, 2)
        EXCEL.Range("D19").Value = MatrizGastos(16, 2)
        EXCEL.Range("D20").Value = MatrizGastos(17, 2)
        EXCEL.Range("D21").Value = MatrizGastos(18, 2)
        EXCEL.Range("D22").Value = MatrizGastos(19, 2)
        EXCEL.Range("D23").Value = MatrizGastos(20, 2)
        EXCEL.Range("D24").Value = MatrizGastos(21, 2)
        EXCEL.Range("D25").Value = MatrizGastos(22, 2)
        EXCEL.Range("D26").Value = MatrizGastos(23, 2)
        EXCEL.Range("D27").Value = MatrizGastos(24, 2)
        EXCEL.Range("D28").Value = MatrizGastos(25, 2)

        'DATOS JUEVES
        EXCEL.Range("E3").Value = MatrizGastos(0, 3)
        EXCEL.Range("E4").Value = MatrizGastos(1, 3)
        EXCEL.Range("E5").Value = MatrizGastos(2, 3)
        EXCEL.Range("E6").Value = MatrizGastos(3, 3)
        EXCEL.Range("E7").Value = MatrizGastos(4, 3)
        EXCEL.Range("E8").Value = MatrizGastos(5, 3)
        EXCEL.Range("E9").Value = MatrizGastos(6, 3)
        EXCEL.Range("E10").Value = MatrizGastos(7, 3)
        EXCEL.Range("E11").Value = MatrizGastos(8, 3)
        EXCEL.Range("E12").Value = MatrizGastos(9, 3)
        EXCEL.Range("E13").Value = MatrizGastos(10, 3)
        EXCEL.Range("E14").Value = MatrizGastos(11, 3)
        EXCEL.Range("E15").Value = MatrizGastos(12, 3)
        EXCEL.Range("E16").Value = MatrizGastos(13, 3)
        EXCEL.Range("E17").Value = MatrizGastos(14, 3)
        EXCEL.Range("E18").Value = MatrizGastos(15, 3)
        EXCEL.Range("E19").Value = MatrizGastos(16, 3)
        EXCEL.Range("E20").Value = MatrizGastos(17, 3)
        EXCEL.Range("E21").Value = MatrizGastos(18, 3)
        EXCEL.Range("E22").Value = MatrizGastos(19, 3)
        EXCEL.Range("E23").Value = MatrizGastos(20, 3)
        EXCEL.Range("E24").Value = MatrizGastos(21, 3)
        EXCEL.Range("E25").Value = MatrizGastos(22, 3)
        EXCEL.Range("E26").Value = MatrizGastos(23, 3)
        EXCEL.Range("E27").Value = MatrizGastos(24, 3)
        EXCEL.Range("E28").Value = MatrizGastos(25, 3)

        'DATOS VIERNES
        EXCEL.Range("F3").Value = MatrizGastos(0, 4)
        EXCEL.Range("F4").Value = MatrizGastos(1, 4)
        EXCEL.Range("F5").Value = MatrizGastos(2, 4)
        EXCEL.Range("F6").Value = MatrizGastos(3, 4)
        EXCEL.Range("F7").Value = MatrizGastos(4, 4)
        EXCEL.Range("F8").Value = MatrizGastos(5, 4)
        EXCEL.Range("F9").Value = MatrizGastos(6, 4)
        EXCEL.Range("F10").Value = MatrizGastos(7, 4)
        EXCEL.Range("F11").Value = MatrizGastos(8, 4)
        EXCEL.Range("F12").Value = MatrizGastos(9, 4)
        EXCEL.Range("F13").Value = MatrizGastos(10, 4)
        EXCEL.Range("F14").Value = MatrizGastos(11, 4)
        EXCEL.Range("F15").Value = MatrizGastos(12, 4)
        EXCEL.Range("F16").Value = MatrizGastos(13, 4)
        EXCEL.Range("F17").Value = MatrizGastos(14, 4)
        EXCEL.Range("F18").Value = MatrizGastos(15, 4)
        EXCEL.Range("F19").Value = MatrizGastos(16, 4)
        EXCEL.Range("F20").Value = MatrizGastos(17, 4)
        EXCEL.Range("F21").Value = MatrizGastos(18, 4)
        EXCEL.Range("F22").Value = MatrizGastos(19, 4)
        EXCEL.Range("F23").Value = MatrizGastos(20, 4)
        EXCEL.Range("F24").Value = MatrizGastos(21, 4)
        EXCEL.Range("F25").Value = MatrizGastos(22, 4)
        EXCEL.Range("F26").Value = MatrizGastos(23, 4)
        EXCEL.Range("F27").Value = MatrizGastos(24, 4)
        EXCEL.Range("F28").Value = MatrizGastos(25, 4)


        'DATOS TOTALES POR FILA
        EXCEL.Range("H3").Formula = "=SUM(B3:F3)"
        EXCEL.Range("H4").Formula = "=SUM(B4:F4)"
        EXCEL.Range("H5").Formula = "=SUM(B5:F5)"
        EXCEL.Range("H6").Formula = "=SUM(B6:F6)"
        EXCEL.Range("H7").Formula = "=SUM(B7:F7)"
        EXCEL.Range("H8").Formula = "=SUM(B8:F8)"
        EXCEL.Range("H9").Formula = "=SUM(B9:F9)"
        EXCEL.Range("H10").Formula = "=SUM(B10:F10)"
        EXCEL.Range("H11").Formula = "=SUM(B11:F11)"
        EXCEL.Range("H12").Formula = "=SUM(B12:F12)"
        EXCEL.Range("H13").Formula = "=SUM(B13:F13)"
        EXCEL.Range("H14").Formula = "=SUM(B14:F14)"
        EXCEL.Range("H15").Formula = "=SUM(B15:F15)"
        EXCEL.Range("H16").Formula = "=SUM(B16:F16)"
        EXCEL.Range("H17").Formula = "=SUM(B17:F17)"
        EXCEL.Range("H18").Formula = "=SUM(B18:F18)"
        EXCEL.Range("H19").Formula = "=SUM(B19:F19)"
        EXCEL.Range("H20").Formula = "=SUM(B20:F20)"
        EXCEL.Range("H21").Formula = "=SUM(B21:F21)"
        EXCEL.Range("H22").Formula = "=SUM(B22:F22)"
        EXCEL.Range("H23").Formula = "=SUM(B23:F23)"
        EXCEL.Range("H24").Formula = "=SUM(B24:F24)"
        EXCEL.Range("H25").Formula = "=SUM(B25:F25)"
        EXCEL.Range("H26").Formula = "=SUM(B26:F26)"
        EXCEL.Range("H27").Formula = "=SUM(B27:F27)"
        EXCEL.Range("H28").Formula = "=SUM(B28:F28)"
        EXCEL.Range("H30").Formula = "=SUM(H3:H28)"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show("¿Volver a la pantalla principal?", "Mensaje de sistema", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress, TextBox9.KeyPress, TextBox8.KeyPress, TextBox7.KeyPress, TextBox6.KeyPress, TextBox5.KeyPress, TextBox4.KeyPress, TextBox3.KeyPress, TextBox2.KeyPress, TextBox11.KeyPress, TextBox10.KeyPress, TextBox66.KeyPress, TextBox65.KeyPress, TextBox64.KeyPress, TextBox63.KeyPress, TextBox62.KeyPress, TextBox61.KeyPress, TextBox60.KeyPress, TextBox59.KeyPress, TextBox58.KeyPress, TextBox57.KeyPress, TextBox56.KeyPress, TextBox55.KeyPress, TextBox54.KeyPress, TextBox53.KeyPress, TextBox52.KeyPress, TextBox51.KeyPress, TextBox50.KeyPress, TextBox49.KeyPress, TextBox48.KeyPress, TextBox47.KeyPress, TextBox46.KeyPress, TextBox45.KeyPress, TextBox44.KeyPress, TextBox43.KeyPress, TextBox42.KeyPress, TextBox41.KeyPress, TextBox40.KeyPress, TextBox39.KeyPress, TextBox38.KeyPress, TextBox37.KeyPress, TextBox36.KeyPress, TextBox35.KeyPress, TextBox34.KeyPress, TextBox33.KeyPress, TextBox32.KeyPress, TextBox31.KeyPress, TextBox30.KeyPress, TextBox29.KeyPress, TextBox28.KeyPress, TextBox27.KeyPress, TextBox26.KeyPress, TextBox25.KeyPress, TextBox24.KeyPress, TextBox23.KeyPress, TextBox22.KeyPress, TextBox21.KeyPress, TextBox20.KeyPress, TextBox19.KeyPress, TextBox18.KeyPress, TextBox17.KeyPress, TextBox16.KeyPress, TextBox15.KeyPress, TextBox14.KeyPress, TextBox13.KeyPress, TextBox12.KeyPress, TextBox99.KeyPress, TextBox98.KeyPress, TextBox97.KeyPress, TextBox96.KeyPress, TextBox95.KeyPress, TextBox94.KeyPress, TextBox93.KeyPress, TextBox92.KeyPress, TextBox91.KeyPress, TextBox90.KeyPress, TextBox89.KeyPress, TextBox88.KeyPress, TextBox87.KeyPress, TextBox86.KeyPress, TextBox85.KeyPress, TextBox84.KeyPress, TextBox83.KeyPress, TextBox82.KeyPress, TextBox81.KeyPress, TextBox80.KeyPress, TextBox79.KeyPress, TextBox78.KeyPress, TextBox77.KeyPress, TextBox76.KeyPress, TextBox75.KeyPress, TextBox74.KeyPress, TextBox73.KeyPress, TextBox72.KeyPress, TextBox162.KeyPress, TextBox161.KeyPress, TextBox160.KeyPress, TextBox159.KeyPress, TextBox158.KeyPress, TextBox157.KeyPress, TextBox156.KeyPress, TextBox155.KeyPress, TextBox154.KeyPress, TextBox153.KeyPress, TextBox152.KeyPress, TextBox151.KeyPress, TextBox150.KeyPress, TextBox149.KeyPress, TextBox148.KeyPress, TextBox147.KeyPress, TextBox146.KeyPress, TextBox145.KeyPress, TextBox144.KeyPress, TextBox143.KeyPress, TextBox142.KeyPress, TextBox141.KeyPress, TextBox140.KeyPress, TextBox139.KeyPress, TextBox138.KeyPress, TextBox137.KeyPress, TextBox136.KeyPress, TextBox135.KeyPress, TextBox134.KeyPress, TextBox133.KeyPress, TextBox132.KeyPress, TextBox131.KeyPress, TextBox130.KeyPress, TextBox129.KeyPress, TextBox128.KeyPress, TextBox127.KeyPress, TextBox126.KeyPress, TextBox125.KeyPress, TextBox124.KeyPress, TextBox123.KeyPress, TextBox122.KeyPress, TextBox121.KeyPress, TextBox120.KeyPress, TextBox119.KeyPress, TextBox118.KeyPress, TextBox117.KeyPress, TextBox116.KeyPress, TextBox115.KeyPress, TextBox114.KeyPress, TextBox113.KeyPress, TextBox112.KeyPress, TextBox111.KeyPress, TextBox110.KeyPress, TextBox109.KeyPress, TextBox108.KeyPress, TextBox107.KeyPress, TextBox106.KeyPress, TextBox105.KeyPress, TextBox104.KeyPress, TextBox103.KeyPress, TextBox102.KeyPress, TextBox101.KeyPress, TextBox100.KeyPress
        If e.KeyChar = "," Then
            e.KeyChar = "."
        End If
        If (IsNumeric(e.KeyChar) Or e.KeyChar = vbBack) Or e.KeyChar = "," Or e.KeyChar = "." Then
            e.Handled = False
        Else
            e.Handled = True
        End If
    End Sub
    Sub REINICIAR_VARIABLES()
        TOTAL_SERVICIO = 0
        TOTAL_CHIP = 0
        TOTAL_CORREO = 0
        TOTAL_FORMULARIO = 0
        TOTAL_AUTONOMO = 0
        TOTAL_ALQUILER = 0
        TOTAL_OBLEAS = 0
        TOTAL_LUZ = 0
        TOTAL_INTERNET = 0
        TOTAL_TELECOM = 0
        TOTAL_AERPC = 0
        TOTAL_ACARA = 0
        TOTAL_SUELDOS = 0
        TOTAL_MINISTERIO = 0
        TOTAL_MANTENIMIENTO = 0
        TOTAL_INSUMOS = 0
        TOTAL_CONTADOR = 0
        TOTAL_VARIOS1 = 0
        TOTAL_VARIOS2 = 0
        TOTAL_VARIOS3 = 0
        TOTAL_VARIOS4 = 0
        TOTAL_GASTOS1 = 0
        TOTAL_GASTOS2 = 0
        TOTAL_GASTOS3 = 0
        TOTAL_GASTOS4 = 0
        TOTAL_GASTOS5 = 0
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim F As Integer = 0

        Carga()
        REINICIAR_VARIABLES()

        'FILAS
        F = 0
        For C = 0 To 4
            TOTAL_SERVICIO = TOTAL_SERVICIO + MatrizGastos(F, C)
        Next
        F = 1
        For C = 0 To 4
            TOTAL_CHIP = TOTAL_CHIP + MatrizGastos(F, C)
        Next
        F = 2
        For C = 0 To 4
            TOTAL_CORREO = TOTAL_CORREO + MatrizGastos(F, C)
        Next
        F = 3
        For C = 0 To 4
            TOTAL_FORMULARIO = TOTAL_FORMULARIO + MatrizGastos(F, C)
        Next
        F = 4
        For C = 0 To 4
            TOTAL_AUTONOMO = TOTAL_AUTONOMO + MatrizGastos(F, C)
        Next
        F = 5
        For C = 0 To 4
            TOTAL_ALQUILER = TOTAL_ALQUILER + MatrizGastos(F, C)
        Next
        F = 6
        For C = 0 To 4
            TOTAL_OBLEAS = TOTAL_OBLEAS + MatrizGastos(F, C)
        Next
        F = 7
        For C = 0 To 4
            TOTAL_LUZ = TOTAL_LUZ + MatrizGastos(F, C)
        Next
        F = 8
        For C = 0 To 4
            TOTAL_INTERNET = TOTAL_INTERNET + MatrizGastos(F, C)
        Next
        F = 9
        For C = 0 To 4
            TOTAL_TELECOM = TOTAL_TELECOM + MatrizGastos(F, C)
        Next
        F = 10
        For C = 0 To 4
            TOTAL_AERPC = TOTAL_AERPC + MatrizGastos(F, C)
        Next
        F = 11
        For C = 0 To 4
            TOTAL_ACARA = TOTAL_ACARA + MatrizGastos(F, C)
        Next
        F = 12
        For C = 0 To 4
            TOTAL_SUELDOS = TOTAL_SUELDOS + MatrizGastos(F, C)
        Next
        F = 13
        For C = 0 To 4
            TOTAL_MINISTERIO = TOTAL_MINISTERIO + MatrizGastos(F, C)
        Next
        F = 14
        For C = 0 To 4
            TOTAL_MANTENIMIENTO = TOTAL_MANTENIMIENTO + MatrizGastos(F, C)
        Next
        F = 15
        For C = 0 To 4
            TOTAL_INSUMOS = TOTAL_INSUMOS + MatrizGastos(F, C)
        Next
        F = 16
        For C = 0 To 4
            TOTAL_CONTADOR = TOTAL_CONTADOR + MatrizGastos(F, C)
        Next
        F = 17
        For C = 0 To 4
            TOTAL_VARIOS1 = TOTAL_VARIOS1 + MatrizGastos(F, C)
        Next
        F = 18
        For C = 0 To 4
            TOTAL_VARIOS2 = TOTAL_VARIOS2 + MatrizGastos(F, C)
        Next
        F = 19
        For C = 0 To 4
            TOTAL_VARIOS3 = TOTAL_VARIOS3 + MatrizGastos(F, C)
        Next
        F = 20
        For C = 0 To 4
            TOTAL_VARIOS4 = TOTAL_VARIOS4 + MatrizGastos(F, C)
        Next
        F = 21
        For C = 0 To 4
            TOTAL_GASTOS1 = TOTAL_GASTOS1 + MatrizGastos(F, C)
        Next
        F = 22
        For C = 0 To 4
            TOTAL_GASTOS2 = TOTAL_GASTOS2 + MatrizGastos(F, C)
        Next
        F = 23
        For C = 0 To 4
            TOTAL_GASTOS3 = TOTAL_GASTOS3 + MatrizGastos(F, C)
        Next
        F = 24
        For C = 0 To 4
            TOTAL_GASTOS4 = TOTAL_GASTOS4 + MatrizGastos(F, C)
        Next
        F = 25
        For C = 0 To 4
            TOTAL_GASTOS5 = TOTAL_GASTOS5 + MatrizGastos(F, C)
        Next

        'FILAS
        TextBox56.Text = "$ " & TOTAL_SERVICIO
        TextBox57.Text = "$ " & TOTAL_CHIP
        TextBox58.Text = "$ " & TOTAL_CORREO
        TextBox59.Text = "$ " & TOTAL_FORMULARIO
        TextBox60.Text = "$ " & TOTAL_AUTONOMO
        TextBox61.Text = "$ " & TOTAL_ALQUILER
        TextBox62.Text = "$ " & TOTAL_OBLEAS
        TextBox63.Text = "$ " & TOTAL_LUZ
        TextBox64.Text = "$ " & TOTAL_INTERNET
        TextBox65.Text = "$ " & TOTAL_TELECOM
        TextBox66.Text = "$ " & TOTAL_AERPC
        TextBox73.Text = "$ " & TOTAL_ACARA
        TextBox148.Text = "$ " & TOTAL_SUELDOS
        TextBox149.Text = "$ " & TOTAL_MINISTERIO
        TextBox150.Text = "$ " & TOTAL_MANTENIMIENTO
        TextBox151.Text = "$ " & TOTAL_INSUMOS
        TextBox152.Text = "$ " & TOTAL_CONTADOR
        TextBox153.Text = "$ " & TOTAL_VARIOS1
        TextBox154.Text = "$ " & TOTAL_VARIOS2
        TextBox155.Text = "$ " & TOTAL_VARIOS3
        TextBox156.Text = "$ " & TOTAL_VARIOS4
        TextBox157.Text = "$ " & TOTAL_GASTOS1
        TextBox158.Text = "$ " & TOTAL_GASTOS2
        TextBox159.Text = "$ " & TOTAL_GASTOS3
        TextBox160.Text = "$ " & TOTAL_GASTOS4
        TextBox161.Text = "$ " & TOTAL_GASTOS5

        TextBox162.Text = "$ " & TOTAL_SERVICIO + TOTAL_CHIP + TOTAL_CORREO + TOTAL_FORMULARIO + TOTAL_AUTONOMO + TOTAL_ALQUILER + TOTAL_OBLEAS + TOTAL_LUZ + TOTAL_INTERNET + TOTAL_TELECOM + TOTAL_AERPC + TOTAL_ACARA + TOTAL_SUELDOS + TOTAL_MINISTERIO + TOTAL_MANTENIMIENTO + TOTAL_INSUMOS + TOTAL_CONTADOR + TOTAL_VARIOS1 + TOTAL_VARIOS2 + TOTAL_VARIOS3 + TOTAL_VARIOS4 + TOTAL_GASTOS1 + TOTAL_GASTOS2 + TOTAL_GASTOS3 + TOTAL_GASTOS4 + TOTAL_GASTOS5
    End Sub
    Private Sub Gastos_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub
    Sub Carga()
        MatrizGastos(0, 0) = Val(TextBox1.Text)
        MatrizGastos(1, 0) = Val(TextBox2.Text)
        MatrizGastos(2, 0) = Val(TextBox3.Text)
        MatrizGastos(3, 0) = Val(TextBox4.Text)
        MatrizGastos(4, 0) = Val(TextBox5.Text)
        MatrizGastos(5, 0) = Val(TextBox6.Text)
        MatrizGastos(6, 0) = Val(TextBox7.Text)
        MatrizGastos(7, 0) = Val(TextBox8.Text)
        MatrizGastos(8, 0) = Val(TextBox9.Text)
        MatrizGastos(9, 0) = Val(TextBox10.Text)
        MatrizGastos(10, 0) = Val(TextBox11.Text)
        MatrizGastos(11, 0) = Val(TextBox77.Text)
        MatrizGastos(12, 0) = Val(TextBox82.Text)
        MatrizGastos(13, 0) = Val(TextBox87.Text)
        MatrizGastos(14, 0) = Val(TextBox92.Text)
        MatrizGastos(15, 0) = Val(TextBox97.Text)
        MatrizGastos(16, 0) = Val(TextBox102.Text)
        MatrizGastos(17, 0) = Val(TextBox107.Text)
        MatrizGastos(18, 0) = Val(TextBox112.Text)
        MatrizGastos(19, 0) = Val(TextBox117.Text)
        MatrizGastos(20, 0) = Val(TextBox122.Text)
        MatrizGastos(21, 0) = Val(TextBox127.Text)
        MatrizGastos(22, 0) = Val(TextBox132.Text)
        MatrizGastos(23, 0) = Val(TextBox137.Text)
        MatrizGastos(24, 0) = Val(TextBox142.Text)
        MatrizGastos(25, 0) = Val(TextBox147.Text)
        MatrizGastos(0, 1) = Val(TextBox12.Text)
        MatrizGastos(1, 1) = Val(TextBox13.Text)
        MatrizGastos(2, 1) = Val(TextBox14.Text)
        MatrizGastos(3, 1) = Val(TextBox15.Text)
        MatrizGastos(4, 1) = Val(TextBox16.Text)
        MatrizGastos(5, 1) = Val(TextBox17.Text)
        MatrizGastos(6, 1) = Val(TextBox18.Text)
        MatrizGastos(7, 1) = Val(TextBox19.Text)
        MatrizGastos(8, 1) = Val(TextBox20.Text)
        MatrizGastos(9, 1) = Val(TextBox21.Text)
        MatrizGastos(10, 1) = Val(TextBox22.Text)
        MatrizGastos(11, 1) = Val(TextBox76.Text)
        MatrizGastos(12, 1) = Val(TextBox81.Text)
        MatrizGastos(13, 1) = Val(TextBox86.Text)
        MatrizGastos(14, 1) = Val(TextBox91.Text)
        MatrizGastos(15, 1) = Val(TextBox96.Text)
        MatrizGastos(16, 1) = Val(TextBox101.Text)
        MatrizGastos(17, 1) = Val(TextBox106.Text)
        MatrizGastos(18, 1) = Val(TextBox111.Text)
        MatrizGastos(19, 1) = Val(TextBox116.Text)
        MatrizGastos(20, 1) = Val(TextBox121.Text)
        MatrizGastos(21, 1) = Val(TextBox126.Text)
        MatrizGastos(22, 1) = Val(TextBox131.Text)
        MatrizGastos(23, 1) = Val(TextBox136.Text)
        MatrizGastos(24, 1) = Val(TextBox141.Text)
        MatrizGastos(25, 1) = Val(TextBox146.Text)
        MatrizGastos(0, 2) = Val(TextBox23.Text)
        MatrizGastos(1, 2) = Val(TextBox24.Text)
        MatrizGastos(2, 2) = Val(TextBox25.Text)
        MatrizGastos(3, 2) = Val(TextBox26.Text)
        MatrizGastos(4, 2) = Val(TextBox27.Text)
        MatrizGastos(5, 2) = Val(TextBox28.Text)
        MatrizGastos(6, 2) = Val(TextBox29.Text)
        MatrizGastos(7, 2) = Val(TextBox30.Text)
        MatrizGastos(8, 2) = Val(TextBox31.Text)
        MatrizGastos(9, 2) = Val(TextBox32.Text)
        MatrizGastos(10, 2) = Val(TextBox33.Text)
        MatrizGastos(11, 2) = Val(TextBox75.Text)
        MatrizGastos(12, 2) = Val(TextBox80.Text)
        MatrizGastos(13, 2) = Val(TextBox85.Text)
        MatrizGastos(14, 2) = Val(TextBox90.Text)
        MatrizGastos(15, 2) = Val(TextBox95.Text)
        MatrizGastos(16, 2) = Val(TextBox100.Text)
        MatrizGastos(17, 2) = Val(TextBox105.Text)
        MatrizGastos(18, 2) = Val(TextBox110.Text)
        MatrizGastos(19, 2) = Val(TextBox115.Text)
        MatrizGastos(20, 2) = Val(TextBox120.Text)
        MatrizGastos(21, 2) = Val(TextBox125.Text)
        MatrizGastos(22, 2) = Val(TextBox130.Text)
        MatrizGastos(23, 2) = Val(TextBox135.Text)
        MatrizGastos(24, 2) = Val(TextBox140.Text)
        MatrizGastos(25, 2) = Val(TextBox145.Text)
        MatrizGastos(0, 3) = Val(TextBox34.Text)
        MatrizGastos(1, 3) = Val(TextBox35.Text)
        MatrizGastos(2, 3) = Val(TextBox36.Text)
        MatrizGastos(3, 3) = Val(TextBox37.Text)
        MatrizGastos(4, 3) = Val(TextBox38.Text)
        MatrizGastos(5, 3) = Val(TextBox39.Text)
        MatrizGastos(6, 3) = Val(TextBox40.Text)
        MatrizGastos(7, 3) = Val(TextBox41.Text)
        MatrizGastos(8, 3) = Val(TextBox42.Text)
        MatrizGastos(9, 3) = Val(TextBox43.Text)
        MatrizGastos(10, 3) = Val(TextBox44.Text)
        MatrizGastos(11, 3) = Val(TextBox74.Text)
        MatrizGastos(12, 3) = Val(TextBox79.Text)
        MatrizGastos(13, 3) = Val(TextBox84.Text)
        MatrizGastos(14, 3) = Val(TextBox89.Text)
        MatrizGastos(15, 3) = Val(TextBox94.Text)
        MatrizGastos(16, 3) = Val(TextBox99.Text)
        MatrizGastos(17, 3) = Val(TextBox104.Text)
        MatrizGastos(18, 3) = Val(TextBox109.Text)
        MatrizGastos(19, 3) = Val(TextBox114.Text)
        MatrizGastos(20, 3) = Val(TextBox119.Text)
        MatrizGastos(21, 3) = Val(TextBox124.Text)
        MatrizGastos(22, 3) = Val(TextBox129.Text)
        MatrizGastos(23, 3) = Val(TextBox134.Text)
        MatrizGastos(24, 3) = Val(TextBox139.Text)
        MatrizGastos(25, 3) = Val(TextBox144.Text)
        MatrizGastos(0, 4) = Val(TextBox45.Text)
        MatrizGastos(1, 4) = Val(TextBox46.Text)
        MatrizGastos(2, 4) = Val(TextBox47.Text)
        MatrizGastos(3, 4) = Val(TextBox48.Text)
        MatrizGastos(4, 4) = Val(TextBox49.Text)
        MatrizGastos(5, 4) = Val(TextBox50.Text)
        MatrizGastos(6, 4) = Val(TextBox51.Text)
        MatrizGastos(7, 4) = Val(TextBox52.Text)
        MatrizGastos(8, 4) = Val(TextBox53.Text)
        MatrizGastos(9, 4) = Val(TextBox54.Text)
        MatrizGastos(10, 4) = Val(TextBox55.Text)
        MatrizGastos(11, 4) = Val(TextBox72.Text)
        MatrizGastos(12, 4) = Val(TextBox78.Text)
        MatrizGastos(13, 4) = Val(TextBox83.Text)
        MatrizGastos(14, 4) = Val(TextBox88.Text)
        MatrizGastos(15, 4) = Val(TextBox93.Text)
        MatrizGastos(16, 4) = Val(TextBox98.Text)
        MatrizGastos(17, 4) = Val(TextBox103.Text)
        MatrizGastos(18, 4) = Val(TextBox108.Text)
        MatrizGastos(19, 4) = Val(TextBox113.Text)
        MatrizGastos(20, 4) = Val(TextBox118.Text)
        MatrizGastos(21, 4) = Val(TextBox123.Text)
        MatrizGastos(22, 4) = Val(TextBox128.Text)
        MatrizGastos(23, 4) = Val(TextBox133.Text)
        MatrizGastos(24, 4) = Val(TextBox138.Text)
        MatrizGastos(25, 4) = Val(TextBox143.Text)
    End Sub


End Class