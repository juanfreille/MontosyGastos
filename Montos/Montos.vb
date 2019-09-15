Public Class Montos
    Dim Matriz(10, 4) As Decimal
    Dim TOTALLUNES, TOTALMARTES, TOTALMIERCOLES, TOTALJUEVES, TOTALVIERNES As Decimal
    Dim TOTALARACELAUTO, TOTALARACELMOTO, TOTALFORMULARIOAUTO, TOTALFORMULARIOMOTO, TOTALSELLADOAUTO, TOTALSELLADOMOTO As Decimal
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If MessageBox.Show("¿Esta seguro de borrar todos los campos y empezar de nuevo?", "Mensaje de sistema", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            Limpiarcampos(Me)
            REINICIAR_VARIABLES()
            For f = 0 To 10
                For c = 0 To 4
                    Matriz(f, c) = 0
                Next
            Next
        End If
    End Sub
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        If MessageBox.Show("¿Volver a la pantalla principal?", "Mensaje de sistema", MessageBoxButtons.YesNo) = DialogResult.Yes Then
            Me.Close()
        End If
    End Sub

    Dim TOTALSUCERPAUTO, TOTALSUCERPMOTO, TOTALSUGITAUTO, TOTALSUGITMOTO, TOTALCORREO As Decimal
    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress, TextBox9.KeyPress, TextBox8.KeyPress, TextBox7.KeyPress, TextBox6.KeyPress, TextBox5.KeyPress, TextBox4.KeyPress, TextBox3.KeyPress, TextBox2.KeyPress, TextBox11.KeyPress, TextBox10.KeyPress, TextBox66.KeyPress, TextBox65.KeyPress, TextBox64.KeyPress, TextBox63.KeyPress, TextBox62.KeyPress, TextBox61.KeyPress, TextBox60.KeyPress, TextBox59.KeyPress, TextBox58.KeyPress, TextBox57.KeyPress, TextBox56.KeyPress, TextBox55.KeyPress, TextBox54.KeyPress, TextBox53.KeyPress, TextBox52.KeyPress, TextBox51.KeyPress, TextBox50.KeyPress, TextBox49.KeyPress, TextBox48.KeyPress, TextBox47.KeyPress, TextBox46.KeyPress, TextBox45.KeyPress, TextBox44.KeyPress, TextBox43.KeyPress, TextBox42.KeyPress, TextBox41.KeyPress, TextBox40.KeyPress, TextBox39.KeyPress, TextBox38.KeyPress, TextBox37.KeyPress, TextBox36.KeyPress, TextBox35.KeyPress, TextBox34.KeyPress, TextBox33.KeyPress, TextBox32.KeyPress, TextBox31.KeyPress, TextBox30.KeyPress, TextBox29.KeyPress, TextBox28.KeyPress, TextBox27.KeyPress, TextBox26.KeyPress, TextBox25.KeyPress, TextBox24.KeyPress, TextBox23.KeyPress, TextBox22.KeyPress, TextBox21.KeyPress, TextBox20.KeyPress, TextBox19.KeyPress, TextBox18.KeyPress, TextBox17.KeyPress, TextBox16.KeyPress, TextBox15.KeyPress, TextBox14.KeyPress, TextBox13.KeyPress, TextBox12.KeyPress
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
        TOTALLUNES = 0
        TOTALMARTES = 0
        TOTALMIERCOLES = 0
        TOTALJUEVES = 0
        TOTALVIERNES = 0
        TOTALARACELAUTO = 0
        TOTALARACELMOTO = 0
        TOTALFORMULARIOAUTO = 0
        TOTALFORMULARIOMOTO = 0
        TOTALSELLADOAUTO = 0
        TOTALSELLADOMOTO = 0
        TOTALSUCERPAUTO = 0
        TOTALSUCERPMOTO = 0
        TOTALSUGITAUTO = 0
        TOTALSUGITMOTO = 0
        TOTALCORREO = 0
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs)
        Dim r As New Globalization.CultureInfo("es-ES")
        r.NumberFormat.CurrencyDecimalSeparator = ","
        r.NumberFormat.NumberDecimalSeparator = ","
        System.Threading.Thread.CurrentThread.CurrentCulture = r
    End Sub
    Sub Carga()
        Matriz(0, 0) = Val(TextBox1.Text)
        Matriz(1, 0) = Val(TextBox2.Text)
        Matriz(2, 0) = Val(TextBox3.Text)
        Matriz(3, 0) = Val(TextBox4.Text)
        Matriz(4, 0) = Val(TextBox5.Text)
        Matriz(5, 0) = Val(TextBox6.Text)
        Matriz(6, 0) = Val(TextBox7.Text)
        Matriz(7, 0) = Val(TextBox8.Text)
        Matriz(8, 0) = Val(TextBox9.Text)
        Matriz(9, 0) = Val(TextBox10.Text)
        Matriz(10, 0) = Val(TextBox11.Text)
        Matriz(0, 1) = Val(TextBox12.Text)
        Matriz(1, 1) = Val(TextBox13.Text)
        Matriz(2, 1) = Val(TextBox14.Text)
        Matriz(3, 1) = Val(TextBox15.Text)
        Matriz(4, 1) = Val(TextBox16.Text)
        Matriz(5, 1) = Val(TextBox17.Text)
        Matriz(6, 1) = Val(TextBox18.Text)
        Matriz(7, 1) = Val(TextBox19.Text)
        Matriz(8, 1) = Val(TextBox20.Text)
        Matriz(9, 1) = Val(TextBox21.Text)
        Matriz(10, 1) = Val(TextBox22.Text)
        Matriz(0, 2) = Val(TextBox23.Text)
        Matriz(1, 2) = Val(TextBox24.Text)
        Matriz(2, 2) = Val(TextBox25.Text)
        Matriz(3, 2) = Val(TextBox26.Text)
        Matriz(4, 2) = Val(TextBox27.Text)
        Matriz(5, 2) = Val(TextBox28.Text)
        Matriz(6, 2) = Val(TextBox29.Text)
        Matriz(7, 2) = Val(TextBox30.Text)
        Matriz(8, 2) = Val(TextBox31.Text)
        Matriz(9, 2) = Val(TextBox32.Text)
        Matriz(10, 2) = Val(TextBox33.Text)
        Matriz(0, 3) = Val(TextBox34.Text)
        Matriz(1, 3) = Val(TextBox35.Text)
        Matriz(2, 3) = Val(TextBox36.Text)
        Matriz(3, 3) = Val(TextBox37.Text)
        Matriz(4, 3) = Val(TextBox38.Text)
        Matriz(5, 3) = Val(TextBox39.Text)
        Matriz(6, 3) = Val(TextBox40.Text)
        Matriz(7, 3) = Val(TextBox41.Text)
        Matriz(8, 3) = Val(TextBox42.Text)
        Matriz(9, 3) = Val(TextBox43.Text)
        Matriz(10, 3) = Val(TextBox44.Text)
        Matriz(0, 4) = Val(TextBox45.Text)
        Matriz(1, 4) = Val(TextBox46.Text)
        Matriz(2, 4) = Val(TextBox47.Text)
        Matriz(3, 4) = Val(TextBox48.Text)
        Matriz(4, 4) = Val(TextBox49.Text)
        Matriz(5, 4) = Val(TextBox50.Text)
        Matriz(6, 4) = Val(TextBox51.Text)
        Matriz(7, 4) = Val(TextBox52.Text)
        Matriz(8, 4) = Val(TextBox53.Text)
        Matriz(9, 4) = Val(TextBox54.Text)
        Matriz(10, 4) = Val(TextBox55.Text)
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim I As Integer = 0
        Dim F, C As Integer

        Carga()
        REINICIAR_VARIABLES()

        C = 0
        For F = 0 To 10
            TOTALLUNES = TOTALLUNES + Matriz(F, C)
        Next
        C = 1
        For F = 0 To 10
            TOTALMARTES = TOTALMARTES + Matriz(F, C)
        Next
        C = 2
        For F = 0 To 10
            TOTALMIERCOLES = TOTALMIERCOLES + Matriz(F, C)
        Next
        C = 3
        For F = 0 To 10
            TOTALJUEVES = TOTALJUEVES + Matriz(F, C)
        Next
        C = 4
        For F = 0 To 10
            TOTALVIERNES = TOTALVIERNES + Matriz(F, C)
        Next

        'FILAS
        F = 0
        For C = 0 To 4
            TOTALARACELAUTO = TOTALARACELAUTO + Matriz(F, C)
        Next
        F = 1
        For C = 0 To 4
            TOTALARACELMOTO = TOTALARACELMOTO + Matriz(F, C)
        Next
        F = 2
        For C = 0 To 4
            TOTALFORMULARIOAUTO = TOTALFORMULARIOAUTO + Matriz(F, C)
        Next
        F = 3
        For C = 0 To 4
            TOTALFORMULARIOMOTO = TOTALFORMULARIOMOTO + Matriz(F, C)
        Next
        F = 4
        For C = 0 To 4
            TOTALSELLADOAUTO = TOTALSELLADOAUTO + Matriz(F, C)
        Next
        F = 5
        For C = 0 To 4
            TOTALSELLADOMOTO = TOTALSELLADOMOTO + Matriz(F, C)
        Next
        F = 6
        For C = 0 To 4
            TOTALSUCERPAUTO = TOTALSUCERPAUTO + Matriz(F, C)
        Next
        F = 7
        For C = 0 To 4
            TOTALSUCERPMOTO = TOTALSUCERPMOTO + Matriz(F, C)
        Next
        F = 8
        For C = 0 To 4
            TOTALSUGITAUTO = TOTALSUGITAUTO + Matriz(F, C)
        Next
        F = 9
        For C = 0 To 4
            TOTALSUGITMOTO = TOTALSUGITMOTO + Matriz(F, C)
        Next
        F = 10
        For C = 0 To 4
            TOTALCORREO = TOTALCORREO + Matriz(F, C)
        Next

        TextBox67.Text = "$ " & TOTALLUNES
        TextBox68.Text = "$ " & TOTALMARTES
        TextBox69.Text = "$ " & TOTALMIERCOLES
        TextBox70.Text = "$ " & TOTALJUEVES
        TextBox71.Text = "$ " & TOTALVIERNES
        'FILAS
        TextBox56.Text = "$ " & TOTALARACELAUTO
        TextBox57.Text = "$ " & TOTALARACELMOTO
        TextBox58.Text = "$ " & TOTALFORMULARIOAUTO
        TextBox59.Text = "$ " & TOTALFORMULARIOMOTO
        TextBox60.Text = "$ " & TOTALSELLADOAUTO
        TextBox61.Text = "$ " & TOTALSELLADOMOTO
        TextBox62.Text = "$ " & TOTALSUCERPAUTO
        TextBox63.Text = "$ " & TOTALSUCERPMOTO
        TextBox64.Text = "$ " & TOTALSUGITAUTO
        TextBox65.Text = "$ " & TOTALSUGITMOTO
        TextBox66.Text = "$ " & TOTALCORREO
        TextBox72.Text = "$ " & TOTALARACELAUTO + TOTALARACELMOTO
        '        TextBox72.Text = Val(TextBox56.Text) + Val(TextBox57.Text)
        TextBox73.Text = "$ " & TOTALLUNES + TOTALMARTES + TOTALMIERCOLES + TOTALJUEVES + TOTALVIERNES
        TextBox74.Text = "$ " & TOTALFORMULARIOAUTO + TOTALFORMULARIOMOTO
        TextBox75.Text = "$ " & TOTALSELLADOAUTO + TOTALSELLADOMOTO + TOTALSUCERPAUTO + TOTALSUCERPMOTO + TOTALSUGITAUTO

    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim EXCEL As New Microsoft.Office.Interop.Excel.Application
        EXCEL.Visible = True
        EXCEL.Workbooks.Add()

        'COLOR DE FILAS
        With EXCEL.Range("A3:A13")
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
        With EXCEL.Range("A15")
            .Interior.ColorIndex = 10
            .WrapText = True
            .Font.Bold = True
            .Font.FontStyle = "Verdana"
        End With
        With EXCEL.Range("H2")
            .Interior.ColorIndex = 10
            .WrapText = True
            .Font.FontStyle = "Verdana"
            .Font.Bold = True
        End With

        With EXCEL.Range("J2:L2")
            .Interior.ColorIndex = 10
            .WrapText = True
            .Font.FontStyle = "Verdana"
            .Font.Bold = True
        End With
        'COLOR DE COLUMNAS
        With EXCEL.Range("B2:F2")
            .Interior.ColorIndex = 24
            .WrapText = True
            .Font.FontStyle = "Verdana"
            .Font.Bold = True
        End With

        'NEGRITA PARA TOTALES
        With EXCEL.Range("H3:H15")
            .Font.Bold = True
        End With
        With EXCEL.Range("B15:F15")
            .Font.Bold = True
        End With
        With EXCEL.Range("J3:L3")
            .Font.Bold = True
        End With
        'aclaraciones
        With EXCEL.Range("J1:L1")
            .Interior.ColorIndex = 24
            .Font.Bold = True
        End With

        'FILAS
        EXCEL.Range("A1").ColumnWidth = 25
        EXCEL.Range("H1").ColumnWidth = 20
        EXCEL.Range("J1").ColumnWidth = 20
        EXCEL.Range("K1").ColumnWidth = 20
        EXCEL.Range("L1").ColumnWidth = 20
        EXCEL.Range("A1").Value = "PLANILLA DE MONTOS"
        EXCEL.Range("A3").Value = "Arancel AUTO"
        EXCEL.Range("A4").Value = "Arancel MOTO"
        EXCEL.Range("A5").Value = "Formulario AUTO"
        EXCEL.Range("A6").Value = "Formulario MOTO"
        EXCEL.Range("A7").Value = "Sellado AUTO"
        EXCEL.Range("A8").Value = "Sellado MOTO"
        EXCEL.Range("A9").Value = "SUCERP AUTO"
        EXCEL.Range("A10").Value = "SUCERP MOTO"
        EXCEL.Range("A11").Value = "SUGIT AUTO"
        EXCEL.Range("A12").Value = "SUGIT MOTO"
        EXCEL.Range("A13").Value = "Correo"
        EXCEL.Range("A15").Value = "TOTAL DIARIO"
        'COLUMNAS
        EXCEL.Range("B2").Value = "LUNES"
        EXCEL.Range("C2").Value = "MARTES"
        EXCEL.Range("D2").Value = "MIERCOLES"
        EXCEL.Range("E2").Value = "JUEVES"
        EXCEL.Range("F2").Value = "VIERNES"
        EXCEL.Range("H2").Value = "TOTAL SEMANAL"
        EXCEL.Range("J2").Value = "TOTAL ARANCELES"
        EXCEL.Range("J3").Formula = "=SUM(H3:H4)"
        EXCEL.Range("K2").Value = "TOTAL FORMULARIOS"
        EXCEL.Range("K3").Formula = "=SUM(H5:H6)"
        EXCEL.Range("L2").Value = "TOTAL OTROS"
        EXCEL.Range("L3").Formula = "=SUM(H7:H11)"
        EXCEL.Range("K1").Value = "Calculos semanales"

        'DATOS LUNES
        EXCEL.Range("B3").Value = Matriz(0, 0)
        EXCEL.Range("B4").Value = Matriz(1, 0)
        EXCEL.Range("B5").Value = Matriz(2, 0)
        EXCEL.Range("B6").Value = Matriz(3, 0)
        EXCEL.Range("B7").Value = Matriz(4, 0)
        EXCEL.Range("B8").Value = Matriz(5, 0)
        EXCEL.Range("B9").Value = Matriz(6, 0)
        EXCEL.Range("B10").Value = Matriz(7, 0)
        EXCEL.Range("B11").Value = Matriz(8, 0)
        EXCEL.Range("B12").Value = Matriz(9, 0)
        EXCEL.Range("B13").Value = Matriz(10, 0)
        EXCEL.Range("B15").Formula = "=SUM(B3:B13)"
        'DATOS MARTES
        EXCEL.Range("C3").Value = Matriz(0, 1)
        EXCEL.Range("C4").Value = Matriz(1, 1)
        EXCEL.Range("C5").Value = Matriz(2, 1)
        EXCEL.Range("C6").Value = Matriz(3, 1)
        EXCEL.Range("C7").Value = Matriz(4, 1)
        EXCEL.Range("C8").Value = Matriz(5, 1)
        EXCEL.Range("C9").Value = Matriz(6, 1)
        EXCEL.Range("C10").Value = Matriz(7, 1)
        EXCEL.Range("C11").Value = Matriz(8, 1)
        EXCEL.Range("C12").Value = Matriz(9, 1)
        EXCEL.Range("C13").Value = Matriz(10, 1)
        EXCEL.Range("C15").Formula = "=SUM(C3:C13)"

        'DATOS MIERCOLES
        EXCEL.Range("D3").Value = Matriz(0, 2)
        EXCEL.Range("D4").Value = Matriz(1, 2)
        EXCEL.Range("D5").Value = Matriz(2, 2)
        EXCEL.Range("D6").Value = Matriz(3, 2)
        EXCEL.Range("D7").Value = Matriz(4, 2)
        EXCEL.Range("D8").Value = Matriz(5, 2)
        EXCEL.Range("D9").Value = Matriz(6, 2)
        EXCEL.Range("D10").Value = Matriz(7, 2)
        EXCEL.Range("D11").Value = Matriz(8, 2)
        EXCEL.Range("D12").Value = Matriz(9, 2)
        EXCEL.Range("D13").Value = Matriz(10, 2)
        EXCEL.Range("D15").Formula = "=SUM(D3:D13)"

        'DATOS JUEVES
        EXCEL.Range("E3").Value = Matriz(0, 3)
        EXCEL.Range("E4").Value = Matriz(1, 3)
        EXCEL.Range("E5").Value = Matriz(2, 3)
        EXCEL.Range("E6").Value = Matriz(3, 3)
        EXCEL.Range("E7").Value = Matriz(4, 3)
        EXCEL.Range("E8").Value = Matriz(5, 3)
        EXCEL.Range("E9").Value = Matriz(6, 3)
        EXCEL.Range("E10").Value = Matriz(7, 3)
        EXCEL.Range("E11").Value = Matriz(8, 3)
        EXCEL.Range("E12").Value = Matriz(9, 3)
        EXCEL.Range("E13").Value = Matriz(10, 3)
        EXCEL.Range("E15").Formula = "=SUM(E3:E13)"

        'DATOS VIERNES
        EXCEL.Range("F3").Value = Matriz(0, 4)
        EXCEL.Range("F4").Value = Matriz(1, 4)
        EXCEL.Range("F5").Value = Matriz(2, 4)
        EXCEL.Range("F6").Value = Matriz(3, 4)
        EXCEL.Range("F7").Value = Matriz(4, 4)
        EXCEL.Range("F8").Value = Matriz(5, 4)
        EXCEL.Range("F9").Value = Matriz(6, 4)
        EXCEL.Range("F10").Value = Matriz(7, 4)
        EXCEL.Range("F11").Value = Matriz(8, 4)
        EXCEL.Range("F12").Value = Matriz(9, 4)
        EXCEL.Range("F13").Value = Matriz(10, 4)
        EXCEL.Range("F15").Formula = "=SUM(F3:F13)"

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
        EXCEL.Range("H15").Formula = "=SUM(B15:F15)"

    End Sub
End Class
