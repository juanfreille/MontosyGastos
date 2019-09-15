Module Limpiador
    Sub Limpiarcampos(ByVal Formulario As Form)
        Dim Text As Object
        For Each Text In Formulario.Controls
            If TypeOf Text Is TextBox Then
                Dim txtTemp As TextBox = CType(Text, TextBox)
                txtTemp.Text = ""
            End If
        Next
        For Each gb As GroupBox In Formulario.Controls.OfType(Of GroupBox)()
            For Each tb As TextBox In gb.Controls.OfType(Of TextBox)()
                tb.Clear()
            Next
        Next

    End Sub
End Module
