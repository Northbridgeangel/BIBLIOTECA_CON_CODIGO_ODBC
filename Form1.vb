Imports System.Data.Odbc

Public Class Form1


    '-------------------------------------------------

    '--- CREO LOS OBJETOS BD A LOS QUE DARE VALOR:

    '--- (LUEGO EN EL FORM_LOAD)

    '--------------------------------------------------

    Private MyCn As New Odbc.OdbcConnection  'MY CONEXION PARA TODOS

    Private MyDatAdp As New OdbcDataAdapter  'ESTE PARA TEMAS

    Private MyCmdBld As New OdbcCommandBuilder 'ESTE PARA TEMAS

    Private MyDataTbl As New DataTable  'ESTE PARA TEMAS

    Private MyRowPosition As Integer = 0   'ESTE PARA TEMAS
    '----------------------------------------------------------------
    Private MyDatAdp2 As New OdbcDataAdapter  'ESTE PARA AUTORES

    Private MyCmdBld2 As New OdbcCommandBuilder 'ESTE PARA AUTORES

    Private MyDataTbl2 As New DataTable  'ESTE PARA AUTORES

    Private MyRowPosition2 As Integer = 0   'ESTE PARA AUTORES
    '----------------------------------------------------------------

    Private MyDatAdp3 As New OdbcDataAdapter  'ESTE PARA LIBROS

    Private MyCmdBld3 As New OdbcCommandBuilder 'ESTE PARA LIBROS

    Private MyDataTbl3 As New DataTable  'ESTE PARA LIBROS

    Private MyRowPosition3 As Integer = 0   'ESTE PARA LIBROS
    '----------------------------------------------------------------
    Private MyDatAdp4 As New OdbcDataAdapter  'ESTE PARA EDITORIALES

    Private MyCmdBld4 As New OdbcCommandBuilder 'ESTE PARA EDITORIALES

    Private MyDataTbl4 As New DataTable  'ESTE PARA EDITORIALES

    Private MyRowPosition4 As Integer = 0   'ESTE PARA EDITORIALES
    '----------------------------------------------------------------
    Private MyDatAdp5 As New OdbcDataAdapter  'ESTE PARA SOCIO

    Private MyCmdBld5 As New OdbcCommandBuilder 'ESTE PARA SOCIO

    Private MyDataTbl5 As New DataTable  'ESTE PARA SOCIO

    Private MyRowPosition5 As Integer = 0   'ESTE PARA SOCIO
    '----------------------------------------------------------------


    Private Sub showrecords()
        If MyDataTbl.Rows.Count = 0 Then
            TextBox1.Text = ""
            TextBox2.Text = ""
            Exit Sub
        End If
        TextBox1.Text = MyDataTbl.Rows(MyRowPosition)("codtema").ToString()
        TextBox2.Text = MyDataTbl.Rows(MyRowPosition)("nombretema").ToString()
    End Sub

    Private Sub showrecords2()
        If MyDataTbl2.Rows.Count = 0 Then
            TextBox3.Text = ""
            TextBox4.Text = ""
            TextBox5.Text = ""
            Exit Sub
        End If
        TextBox3.Text = MyDataTbl2.Rows(MyRowPosition2)("codautor").ToString()
        TextBox4.Text = MyDataTbl2.Rows(MyRowPosition2)("nombre").ToString()
        TextBox5.Text = MyDataTbl2.Rows(MyRowPosition2)("nacionalidad").ToString()
    End Sub

    Private Sub showrecords3()
        If MyDataTbl3.Rows.Count = 0 Then
            TextBox6.Text = ""
            TextBox7.Text = ""
            TextBox8.Text = ""
            TextBox9.Text = ""
            TextBox10.Text = ""
            TextBox11.Text = ""
            TextBox12.Text = ""
            TextBox13.Text = ""
            TextBox14.Text = ""
            TextBox15.Text = ""
            TextBox16.Text = ""
            TextBox17.Text = ""
            TextBox18.Text = ""
            TextBox19.Text = ""
            Exit Sub
        End If
        TextBox6.Text = MyDataTbl3.Rows(MyRowPosition3)("codlibro").ToString()
        TextBox7.Text = MyDataTbl3.Rows(MyRowPosition3)("titulo").ToString()
        TextBox8.Text = MyDataTbl3.Rows(MyRowPosition3)("codautor").ToString()
        TextBox9.Text = MyDataTbl3.Rows(MyRowPosition3)("codeditorial").ToString()
        TextBox10.Text = MyDataTbl3.Rows(MyRowPosition3)("codtema").ToString()
        TextBox11.Text = MyDataTbl3.Rows(MyRowPosition3)("isbn").ToString()
        TextBox12.Text = MyDataTbl3.Rows(MyRowPosition3)("deplegal").ToString()
        TextBox13.Text = MyDataTbl3.Rows(MyRowPosition3)("idioma").ToString()
        TextBox14.Text = MyDataTbl3.Rows(MyRowPosition3)("anoedicion").ToString()
        TextBox15.Text = MyDataTbl3.Rows(MyRowPosition3)("numpaginas").ToString()
        TextBox16.Text = MyDataTbl3.Rows(MyRowPosition3)("numejemplares").ToString()
        TextBox17.Text = MyDataTbl3.Rows(MyRowPosition3)("anoprimeraedicion").ToString()
        TextBox18.Text = MyDataTbl3.Rows(MyRowPosition3)("cantidad").ToString()
        TextBox19.Text = MyDataTbl3.Rows(MyRowPosition3)("portada").ToString()
    End Sub
    Private Sub showrecords4()
        If MyDataTbl4.Rows.Count = 0 Then
            TextBox20.Text = ""
            TextBox21.Text = ""
            TextBox22.Text = ""
            TextBox23.Text = ""
            TextBox24.Text = ""
            TextBox25.Text = ""
            TextBox26.Text = ""
            TextBox27.Text = ""
            TextBox28.Text = ""
            TextBox29.Text = ""

            Exit Sub
        End If
        TextBox20.Text = MyDataTbl4.Rows(MyRowPosition4)("codeditorial").ToString()
        TextBox21.Text = MyDataTbl4.Rows(MyRowPosition4)("nombreeditorial").ToString()
        TextBox22.Text = MyDataTbl4.Rows(MyRowPosition4)("direccion").ToString()
        TextBox23.Text = MyDataTbl4.Rows(MyRowPosition4)("poblacion").ToString()
        TextBox24.Text = MyDataTbl4.Rows(MyRowPosition4)("provincia").ToString()
        TextBox25.Text = MyDataTbl4.Rows(MyRowPosition4)("codpostal").ToString()
        TextBox26.Text = MyDataTbl4.Rows(MyRowPosition4)("pais").ToString()
        TextBox27.Text = MyDataTbl4.Rows(MyRowPosition4)("telefono").ToString()
        TextBox28.Text = MyDataTbl4.Rows(MyRowPosition4)("email").ToString()
        TextBox29.Text = MyDataTbl4.Rows(MyRowPosition4)("web").ToString()

    End Sub
    Private Sub showrecords5()
        If MyDataTbl5.Rows.Count = 0 Then
            TextBox30.Text = ""
            TextBox31.Text = ""
            TextBox32.Text = ""
            TextBox33.Text = ""
            TextBox34.Text = ""
            TextBox35.Text = ""
            TextBox36.Text = ""
            TextBox37.Text = ""
            TextBox38.Text = ""
            TextBox39.Text = ""

            Exit Sub
        End If
        TextBox30.Text = MyDataTbl5.Rows(MyRowPosition5)("codsocio").ToString()
        TextBox31.Text = MyDataTbl5.Rows(MyRowPosition5)("nombre_y_apellidos").ToString()
        TextBox32.Text = MyDataTbl5.Rows(MyRowPosition5)("dni").ToString()
        TextBox33.Text = MyDataTbl5.Rows(MyRowPosition5)("direccion").ToString()
        TextBox34.Text = MyDataTbl5.Rows(MyRowPosition5)("poblacion").ToString()
        TextBox35.Text = MyDataTbl5.Rows(MyRowPosition5)("provincia").ToString()
        TextBox36.Text = MyDataTbl5.Rows(MyRowPosition5)("cp").ToString()
        TextBox37.Text = MyDataTbl5.Rows(MyRowPosition5)("telefono").ToString()
        TextBox38.Text = MyDataTbl5.Rows(MyRowPosition5)("fecha").ToString()
        TextBox39.Text = MyDataTbl5.Rows(MyRowPosition5)("email").ToString()
        TextBox40.Text = MyDataTbl5.Rows(MyRowPosition5)("foto").ToString()

    End Sub






    'Inicializo los objetos en la carga del formulario

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        '-----------------------
        'ESTÉTICA DE LOS BOTONES, QUE APAREZCA UN BOCADILLO DE TEXTO = ARRAY DE TOOLTIP
        Dim TL(42) As ToolTip 'DIMENSIONO SEGUN EN QUE CONTROLES QUIERO MI AYUDA
        For I = 0 To 42
            TL(I) = New ToolTip
            TL(I).IsBalloon = True
        Next
        TL(0).SetToolTip(Me.BTN_IRALPRIMERO, "Ir al primero")
        TL(1).SetToolTip(Me.BTN_RETROCEDER, "Ir al anterior")
        TL(2).SetToolTip(Me.BTN_AVANZAR, "Ir al siguiente")
        TL(3).SetToolTip(Me.BTN_IRALULTIMO, "Ir al último")
        TL(4).SetToolTip(Me.BTN_GRABAR, "Guardar Registro")
        TL(5).SetToolTip(Me.BTN_NUEVO, "Crear Registro")
        TL(6).SetToolTip(Me.BTN_BUSCAR, "Buscar")
        TL(7).SetToolTip(Me.BTN_BORRAR, "Borrar")
        TL(8).SetToolTip(Me.BTN_IRALPRIMERO2, "Ir al primero")
        TL(9).SetToolTip(Me.BTN_RETROCEDER2, "Ir al anterior")
        TL(10).SetToolTip(Me.BTN_AVANZAR2, "Ir al siguiente")
        TL(11).SetToolTip(Me.BTN_IRALULTIMO2, "Ir al último")
        TL(12).SetToolTip(Me.BTN_GRABAR2, "Guardar Registro")
        TL(13).SetToolTip(Me.BTN_NUEVO2, "Crear Registro")
        TL(14).SetToolTip(Me.BTN_BUSCAR2, "Buscar")
        TL(15).SetToolTip(Me.BTN_BORRAR2, "Borrar")
        TL(16).SetToolTip(Me.BTN_IRALPRIMERO3, "Ir al primero")
        TL(17).SetToolTip(Me.BTN_RETROCEDER3, "Ir al anterior")
        TL(18).SetToolTip(Me.BTN_AVANZAR3, "Ir al siguiente")
        TL(19).SetToolTip(Me.BTN_IRALULTIMO3, "Ir al último")
        TL(20).SetToolTip(Me.BTN_GRABAR3, "Guardar Registro")
        TL(21).SetToolTip(Me.BTN_NUEVO3, "Crear Registro")
        TL(22).SetToolTip(Me.BTN_BUSCAR3, "Buscar")
        TL(23).SetToolTip(Me.BTN_BORRAR3, "Borrar")
        TL(24).SetToolTip(Me.BTN_PORTADA, "Agregar portada")
        TL(25).SetToolTip(Me.BTN_IRALPRIMERO4, "Ir al primero")
        TL(26).SetToolTip(Me.BTN_RETROCEDER4, "Ir al anterior")
        TL(27).SetToolTip(Me.BTN_AVANZAR4, "Ir al siguiente")
        TL(28).SetToolTip(Me.BTN_IRALULTIMO4, "Ir al último")
        TL(29).SetToolTip(Me.BTN_GRABAR4, "Guardar Registro")
        TL(30).SetToolTip(Me.BTN_NUEVO4, "Crear Registro")
        TL(31).SetToolTip(Me.BTN_BUSCAR4, "Buscar")
        TL(32).SetToolTip(Me.BTN_BORRAR4, "Borrar")
        TL(33).SetToolTip(Me.BTN_IRALPRIMERO5, "Ir al primero")
        TL(34).SetToolTip(Me.BTN_RETROCEDER5, "Ir al anterior")
        TL(35).SetToolTip(Me.BTN_AVANZAR5, "Ir al siguiente")
        TL(36).SetToolTip(Me.BTN_IRALULTIMO5, "Ir al último")
        TL(37).SetToolTip(Me.BTN_GRABAR5, "Guardar Registro")
        TL(38).SetToolTip(Me.BTN_NUEVO5, "Crear Registro")
        TL(39).SetToolTip(Me.BTN_BUSCAR5, "Buscar")
        TL(40).SetToolTip(Me.BTN_BORRAR5, "Borrar")
        TL(41).SetToolTip(Me.BTN_FOTO, "Agregar foto")
        '-----------------------




        MyCn.ConnectionString = "Dsn=BIBLIOTECA CON CODIGO"
        MyCn.Open()
        '-------------------------------------------------------------- CARGA PARA TEMA

        MyDatAdp = New OdbcDataAdapter(“SELECT * FROM TEMA ORDER BY CODTEMA”, MyCn)

        MyCmdBld = New OdbcCommandBuilder(MyDatAdp)

        MyDatAdp.Fill(MyDataTbl)

        '-------------------------------------------------------------- CARGA PARA AUTORES

        MyDatAdp2 = New OdbcDataAdapter(“SELECT * FROM AUTORES ORDER BY CODAUTOR”, MyCn)

        MyCmdBld2 = New OdbcCommandBuilder(MyDatAdp2)

        MyDatAdp2.Fill(MyDataTbl2)

        '-------------------------------------------------------------- CARGA PARA LIBROS

        MyDatAdp3 = New OdbcDataAdapter(“SELECT * FROM LIBROS”, MyCn)

        MyCmdBld3 = New OdbcCommandBuilder(MyDatAdp3)

        MyDatAdp3.Fill(MyDataTbl3)

        '-------------------------------------------------------------- CARGA PARA EDITORIALES

        MyDatAdp4 = New OdbcDataAdapter(“SELECT * FROM EDITORIALES ORDER BY CODEDITORIAL”, MyCn)

        MyCmdBld4 = New OdbcCommandBuilder(MyDatAdp4)

        MyDatAdp4.Fill(MyDataTbl4)


        '-------------------------------------------------------------- CARGA PARA SOCIO
        MyDatAdp5 = New OdbcDataAdapter(“SELECT * FROM SOCIO”, MyCn)

        MyCmdBld5 = New OdbcCommandBuilder(MyDatAdp5)

        MyDatAdp5.Fill(MyDataTbl5)



        '___________________________________________________________________________
        'CARGA EN EL COMBO LOS VALORES (Antes de ir a los últimos - Importa el orden)
        '___________________________________________________________________________
        Call cargatemas()
        Call cargaeditoriales()
        Call cargaautores()

        '___________________________________________________________________
        'PARA QUE MUESTRE EL ÚLTIMO REGISTRO DE CADA TABLA (en el form load)
        '___________________________________________________________________
        BTN_IRALULTIMO.PerformClick()
        MyRowPosition = MyDataTbl.Rows.Count - 1
        Call showrecords()
        MyRowPosition2 = MyDataTbl2.Rows.Count - 1
        Call showrecords2()
        MyRowPosition3 = MyDataTbl3.Rows.Count - 1
        Call showrecords3()
        Call miraportada()
        MyRowPosition4 = MyDataTbl4.Rows.Count - 1
        Call showrecords4()
        MyRowPosition5 = MyDataTbl5.Rows.Count - 1
        Call showrecords5()
        Call mirafoto()


        '________________________________________________________________
        'CARGA EN EL COMBOBOX1 EL TÍTULO DE LOS LIBROS (En el Form1 Load)
        '________________________________________________________________
        Cursor.Current = Cursors.WaitCursor
        MyRowPosition3 = 0

        'MyDataTbl3 porque es la tabla libros
        For I = 1 To MyDataTbl3.Rows.Count
            Call showrecords3()
            'Añade al combo textbox7.text porque es el registro "TÍTULO" (Carga en el combo el título de los libros)
            ComboBox1.Items.Add(TextBox7.Text)
            MyRowPosition3 = I
        Next
        MyRowPosition3 = 0

        '________________________________________________________________
        'CARGA EN EL COMBOBOX2 EL NOMBRE DE LOS SOCIOS (En el Form1 Load)
        '________________________________________________________________
        Cursor.Current = Cursors.WaitCursor
        MyRowPosition5 = 0

        'MyDataTbl5 porque es la tabla socios
        For I = 1 To MyDataTbl5.Rows.Count
            Call showrecords5()
            'Añade al combo, textbox31.text (Carga en el combo el nombre de los socios)
            ComboBox2.Items.Add(TextBox31.Text)
            MyRowPosition5 = I
        Next
        MyRowPosition5 = 0


    End Sub





    Sub REFRESCA()
        MyCn.Close()
        MyCn.ConnectionString = "Dsn=BIBLIOTECA CON CODIGO"
        MyDatAdp = New OdbcDataAdapter("SELECT * FROM TEMA", MyCn)
        MyCmdBld = New OdbcCommandBuilder(MyDatAdp)
        MyDatAdp.Fill(MyDataTbl)
    End Sub

    Sub REFRESCA2()
        MyCn.Close()
        MyCn.ConnectionString = "Dsn=BIBLIOTECA CON CODIGO"
        MyDatAdp2 = New OdbcDataAdapter("SELECT * FROM AUTORES", MyCn)
        MyCmdBld2 = New OdbcCommandBuilder(MyDatAdp2)
        MyDatAdp2.Fill(MyDataTbl2)
    End Sub

    Sub REFRESCA3()
        MyCn.Close()
        MyCn.ConnectionString = "Dsn=BIBLIOTECA CON CODIGO"
        MyDatAdp3 = New OdbcDataAdapter("SELECT * FROM LIBROS", MyCn)
        MyCmdBld3 = New OdbcCommandBuilder(MyDatAdp3)
        MyDatAdp3.Fill(MyDataTbl3)
    End Sub
    Sub REFRESCA4()
        MyCn.Close()
        MyCn.ConnectionString = "Dsn=BIBLIOTECA CON CODIGO"
        MyDatAdp4 = New OdbcDataAdapter("SELECT * FROM EDITORIALES", MyCn)
        MyCmdBld4 = New OdbcCommandBuilder(MyDatAdp4)
        MyDatAdp4.Fill(MyDataTbl4)
    End Sub
    Sub REFRESCA5()
        MyCn.Close()
        MyCn.ConnectionString = "Dsn=BIBLIOTECA CON CODIGO"
        MyDatAdp5 = New OdbcDataAdapter("SELECT * FROM SOCIO", MyCn)
        MyCmdBld5 = New OdbcCommandBuilder(MyDatAdp5)
        MyDatAdp5.Fill(MyDataTbl5)
    End Sub



    '_______________________________________________________________________________________________
    'TAB 1 - TEMAS
    '_______________________________________________________________________________________________


    'CREO LOS BOTONES 

    Private Sub BTN_IRALPRIMERO_Click(sender As Object, e As EventArgs) Handles BTN_IRALPRIMERO.Click
        MyRowPosition = 0
        Call showrecords()
    End Sub

    Private Sub BTN_RETROCEDER_Click(sender As Object, e As EventArgs) Handles BTN_RETROCEDER.Click
        If MyRowPosition = 0 Then Exit Sub
        MyRowPosition = MyRowPosition - 1
        Call showrecords()
    End Sub

    Private Sub BTN_AVANZAR_Click(sender As Object, e As EventArgs) Handles BTN_AVANZAR.Click
        If MyDataTbl.Rows.Count - 1 = MyRowPosition Then Exit Sub
        MyRowPosition = MyRowPosition + 1
        Call showrecords()
    End Sub

    Private Sub BTN_IRALULTIMO_Click(sender As Object, e As EventArgs) Handles BTN_IRALULTIMO.Click
        MyRowPosition = MyDataTbl.Rows.Count - 1
        Call showrecords()
    End Sub


    Private Sub BTN_NUEVO_Click(sender As Object, e As EventArgs) Handles BTN_NUEVO.Click
        Dim MyNewRow As DataRow = MyDataTbl.NewRow()
        MyDataTbl.Rows.Add(MyNewRow)
        MyRowPosition = MyDataTbl.Rows.Count - 1
        Me.showrecords()
        TextBox1.Select()
        BTN_GRABAR.Enabled = True
        BTN_NUEVO.Enabled = False
    End Sub


    Private Sub BTN_GRABAR_Click(sender As Object, e As EventArgs) Handles BTN_GRABAR.Click
        'DATA ADAPTER
        Try
            If MyDataTbl.Rows.Count <> 0 Then
                MyDataTbl.Rows(MyRowPosition)("codtema") = TextBox1.Text
                MyDataTbl.Rows(MyRowPosition)("nombretema") = UCase(TextBox2.Text)
                MyDatAdp.Update(MyDataTbl)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call REFRESCA()
        BTN_GRABAR.Enabled = False
        BTN_NUEVO.Enabled = True
        Call cargatemas()
    End Sub

    Private Sub BTN_BUSCAR_Click(sender As Object, e As EventArgs) Handles BTN_BUSCAR.Click

        GroupBox1.Visible = True


    End Sub

    Private Sub BTN_BORRAR_Click(sender As Object, e As EventArgs) Handles BTN_BORRAR.Click
        Dim R As Integer 'VARIABLE R PARA EL SI = 6 O EL NO =7 DEL MSGBOX
        R = MsgBox("¿Estas seguro de querer borrar el registro?", vbYesNo, "BORRAR REGISTRO")
        If R = 7 Then Exit Sub
        If MyDataTbl.Rows.Count <> 0 Then  'SI HAY FILAS ENTONCES
            MyDataTbl.Rows(MyRowPosition).Delete()
            MyDatAdp.Update(MyDataTbl)
            MyRowPosition = 0
            Me.showrecords()
            Call REFRESCA()
        End If
    End Sub


    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        'checked cambia para ponerse activado o para ponerse desactivado
        If RadioButton1.Checked = True Then
            Dim sebusca As String
            Cursor.Current = Cursors.WaitCursor
            MyRowPosition = 0
            Call showrecords()

            sebusca = InputBox("Mete el código del tema", "SE BUSCA")
            For I = 0 To MyDataTbl.Rows.Count
                Call showrecords()
                'se busca por textbox1.text porque es el registro "CODTEMA" (La clave principal de tabla tema en Access)
                If InStr(TextBox1.Text, sebusca) <> 0 Then
                    GoTo sal
                End If
                MyRowPosition = I
            Next
            MsgBox("No esta")


            'El cursor del mouse le da forma de actualizar
sal:

            Cursor.Current = Cursors.Default
            GroupBox1.Visible = False
        End If
        RadioButton1.Checked = False

    End Sub

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        'checked cambia para ponerse activado o para ponerse desactivado
        If RadioButton2.Checked = True Then
            Dim sebusca As String
            Cursor.Current = Cursors.WaitCursor
            MyRowPosition = 0
            Call showrecords()

            'Ucase pone en mayusculas -  si metes en minusculas no lo coge
            sebusca = UCase(InputBox("Mete el nombre del tema", "SE BUSCA"))

            For I = 0 To MyDataTbl.Rows.Count
                Call showrecords()
                'se busca por textbox2.text porque es el registro "NOMBRE" (del tema)
                If InStr(TextBox2.Text, sebusca) <> 0 Then
                    GoTo sal
                End If
                MyRowPosition = I
            Next
            MsgBox("No esta")

sal:
            'El cursor del mouse le da forma de actualizar
            Cursor.Current = Cursors.Default

            GroupBox1.Visible = False
        End If
        RadioButton2.Checked = False
    End Sub



    '_______________________________________________________________________________________________
    'TAB 2 - AUTORES
    '_______________________________________________________________________________________________

    '-----------------------
    'ESTÉTICA DE LOS BOTONES, QUE APAREZCA UN BOCADILLO DE TEXTO = ARRAY DE TOOLTIP
    'Hacer después



    'CREO LOS BOTONES 

    Private Sub BTN_IRALPRIMERO2_Click(sender As Object, e As EventArgs) Handles BTN_IRALPRIMERO2.Click
        MyRowPosition2 = 0
        Call showrecords2()
    End Sub

    Private Sub BTN_RETROCEDER2_Click(sender As Object, e As EventArgs) Handles BTN_RETROCEDER2.Click
        If MyRowPosition2 = 0 Then Exit Sub
        MyRowPosition2 = MyRowPosition2 - 1
        Call showrecords2()
    End Sub

    Private Sub BTN_AVANZAR2_Click(sender As Object, e As EventArgs) Handles BTN_AVANZAR2.Click
        If MyDataTbl2.Rows.Count - 1 = MyRowPosition2 Then Exit Sub
        MyRowPosition2 = MyRowPosition2 + 1
        Call showrecords2()
    End Sub

    Private Sub BTN_IRALULTIMO2_Click(sender As Object, e As EventArgs) Handles BTN_IRALULTIMO2.Click
        MyRowPosition2 = MyDataTbl2.Rows.Count - 1
        Call showrecords2()
    End Sub


    Private Sub BTN_NUEVO2_Click(sender As Object, e As EventArgs) Handles BTN_NUEVO2.Click
        Dim MyNewRow As DataRow = MyDataTbl2.NewRow()
        MyDataTbl2.Rows.Add(MyNewRow)
        MyRowPosition2 = MyDataTbl2.Rows.Count - 1
        Me.showrecords2()
        TextBox3.Select()
        BTN_GRABAR2.Enabled = True
        BTN_NUEVO2.Enabled = False
    End Sub


    Private Sub BTN_GRABAR2_Click(sender As Object, e As EventArgs) Handles BTN_GRABAR2.Click
        'DATA ADAPTER
        Try
            If MyDataTbl2.Rows.Count <> 0 Then
                MyDataTbl2.Rows(MyRowPosition2)("codautor") = TextBox3.Text
                MyDataTbl2.Rows(MyRowPosition2)("nombre") = UCase(TextBox4.Text)
                MyDataTbl2.Rows(MyRowPosition2)("nacionalidad") = UCase(TextBox5.Text)
                MyDatAdp2.Update(MyDataTbl2)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call REFRESCA2()
        BTN_GRABAR2.Enabled = False
        BTN_NUEVO2.Enabled = True
        Call cargaautores()
    End Sub

    Private Sub BTN_BUSCAR2_Click(sender As Object, e As EventArgs) Handles BTN_BUSCAR2.Click

        GroupBox2.Visible = True


    End Sub

    Private Sub BTN_BORRAR2_Click(sender As Object, e As EventArgs) Handles BTN_BORRAR2.Click
        Dim R As Integer 'VARIABLE R PARA EL SI = 6 O EL NO =7 DEL MSGBOX
        R = MsgBox("¿Estas seguro de querer borrar el registro?", vbYesNo, "BORRAR REGISTRO")
        If R = 7 Then Exit Sub
        If MyDataTbl2.Rows.Count <> 0 Then  'SI HAY FILAS ENTONCES
            MyDataTbl2.Rows(MyRowPosition2).Delete()
            MyDatAdp2.Update(MyDataTbl2)
            MyRowPosition2 = 0
            Me.showrecords2()
            Call REFRESCA2()
        End If
    End Sub


    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        'checked cambia para ponerse activado o para ponerse desactivado
        If RadioButton3.Checked = True Then
            Dim sebusca As String
            Cursor.Current = Cursors.WaitCursor
            MyRowPosition2 = 0
            Call showrecords2()

            sebusca = InputBox("Mete el código del autor", "SE BUSCA")
            For I = 0 To MyDataTbl2.Rows.Count
                Call showrecords2()
                'se busca por textbox3.text porque es el registro "CODAUTOR" (La clave principal de autores en Access)
                If InStr(TextBox3.Text, sebusca) <> 0 Then
                    GoTo sal
                End If
                MyRowPosition2 = I
            Next
            MsgBox("No esta")


            'El cursor del mouse le da forma de actualizar
sal:

            Cursor.Current = Cursors.Default
            GroupBox2.Visible = False
        End If
        RadioButton3.Checked = False

    End Sub

    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        'checked cambia para ponerse activado o para ponerse desactivado
        If RadioButton4.Checked = True Then
            Dim sebusca As String
            Cursor.Current = Cursors.WaitCursor
            MyRowPosition2 = 0
            Call showrecords2()

            'Ucase pone en mayusculas -  si metes en minusculas no lo coge
            sebusca = UCase(InputBox("Mete el nombre del autor", "SE BUSCA"))

            For I = 0 To MyDataTbl2.Rows.Count
                Call showrecords2()
                'se busca por textbox4.text porque es el registro "NOMBREAUTOR" (de la tabla autores)
                If InStr(TextBox4.Text, sebusca) <> 0 Then
                    GoTo sal
                End If
                MyRowPosition2 = I
            Next
            MsgBox("No esta")


            'El cursor del mouse le da forma de actualizar
sal:

            Cursor.Current = Cursors.Default

            GroupBox2.Visible = False
        End If
        RadioButton4.Checked = False
    End Sub
    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        'checked cambia para ponerse activado o para ponerse desactivado
        If RadioButton5.Checked = True Then
            Dim sebusca As String
            Cursor.Current = Cursors.WaitCursor
            MyRowPosition2 = 0
            Call showrecords2()

            'Ucase pone en mayusculas -  si metes en minusculas no lo coge
            sebusca = UCase(InputBox("Mete la nacionalidad", "SE BUSCA"))

            For I = 0 To MyDataTbl2.Rows.Count
                Call showrecords2()
                'se busca por textbox5.text porque es el registro "NACIONALIDAD" 
                If InStr(TextBox5.Text, sebusca) <> 0 Then
                    GoTo sal
                End If
                MyRowPosition2 = I
            Next
            MsgBox("No esta")


            'El cursor del mouse le da forma de actualizar
sal:

            Cursor.Current = Cursors.Default

            GroupBox2.Visible = False
        End If
        RadioButton5.Checked = False
    End Sub



    '_______________________________________________________________________________________________
    'TAB 3 - LIBROS
    '_______________________________________________________________________________________________

    '-----------------------
    'ESTÉTICA DE LOS BOTONES, QUE APAREZCA UN BOCADILLO DE TEXTO = ARRAY DE TOOLTIP
    'Hacer después



    'CREO LOS BOTONES 

    Private Sub BTN_IRALPRIMERO3_Click(sender As Object, e As EventArgs) Handles BTN_IRALPRIMERO3.Click
        MyRowPosition3 = 0
        Call showrecords3()
        Call miraportada()
        ComboBox1.Text = "" 'Borra el texto del combo
    End Sub

    Private Sub BTN_RETROCEDER3_Click(sender As Object, e As EventArgs) Handles BTN_RETROCEDER3.Click
        If MyRowPosition3 = 0 Then Exit Sub
        MyRowPosition3 = MyRowPosition3 - 1
        Call showrecords3()
        Call miraportada()
        ComboBox1.Text = "" 'Borra el texto del combo
    End Sub

    Private Sub BTN_AVANZAR3_Click(sender As Object, e As EventArgs) Handles BTN_AVANZAR3.Click
        If MyDataTbl3.Rows.Count - 1 = MyRowPosition3 Then Exit Sub
        MyRowPosition3 = MyRowPosition3 + 1
        Call showrecords3()
        Call miraportada()
        ComboBox1.Text = "" 'Borra el texto del combo
    End Sub

    Private Sub BTN_IRALULTIMO3_Click(sender As Object, e As EventArgs) Handles BTN_IRALULTIMO3.Click
        MyRowPosition3 = MyDataTbl3.Rows.Count - 1
        Call showrecords3()
        Call miraportada()
        ComboBox1.Text = "" 'Borra el texto del combo
    End Sub

    Private Sub BTN_NUEVO3_Click(sender As Object, e As EventArgs) Handles BTN_NUEVO3.Click
        Dim MyNewRow As DataRow = MyDataTbl3.NewRow()
        MyDataTbl3.Rows.Add(MyNewRow)
        MyRowPosition3 = MyDataTbl3.Rows.Count - 1
        Me.showrecords3()
        TextBox6.Select()
        BTN_GRABAR3.Enabled = True
        BTN_NUEVO3.Enabled = False
        Label31.Visible = False 'Hace invisible en Label31 "LIBROS"(para botón nuevo)
        ComboBox1.Visible = False 'Hace invisible en combobox1(para botón nuevo)
        ComboBox1.Enabled = False 'Desabilita el Combobox1 (VAYA SER QUE TOQUEN AHÍ Y LE MANDE A OTRO LIBRO SIN QUERER)
        BTN_BUSCAR3.Enabled = False
        'IMAGEN POR DEFECTO
        PictureBox1.Image = Image.FromFile("C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\Imagenes_por_defecto\Imagen_base_portada.png")
        'LIMPIA EL TEXTO DE LOS COMBOS AL CREAR UN NUEVO LIBRO
        CMB_EDITORIAL.Text = ""
        CMB_AUTOR.Text = ""
        CMB_TEMA.Text = ""
    End Sub


    Private Sub BTN_GRABAR3_Click(sender As Object, e As EventArgs) Handles BTN_GRABAR3.Click
        'IMAGEN POR DEFECTO
        PictureBox1.Image = Image.FromFile("C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\Imagenes_por_defecto\Imagen_base_portada.png")
        'DATA ADAPTER
        Try
            If MyDataTbl3.Rows.Count <> 0 Then
                MyDataTbl3.Rows(MyRowPosition3)("codlibro") = TextBox6.Text
                MyDataTbl3.Rows(MyRowPosition3)("titulo") = UCase(TextBox7.Text)
                MyDataTbl3.Rows(MyRowPosition3)("codautor") = TextBox8.Text
                MyDataTbl3.Rows(MyRowPosition3)("codeditorial") = TextBox9.Text
                MyDataTbl3.Rows(MyRowPosition3)("codtema") = TextBox10.Text
                MyDataTbl3.Rows(MyRowPosition3)("isbn") = TextBox11.Text
                MyDataTbl3.Rows(MyRowPosition3)("deplegal") = TextBox12.Text
                MyDataTbl3.Rows(MyRowPosition3)("idioma") = UCase(TextBox13.Text)
                MyDataTbl3.Rows(MyRowPosition3)("anoedicion") = TextBox14.Text
                MyDataTbl3.Rows(MyRowPosition3)("numpaginas") = TextBox15.Text
                MyDataTbl3.Rows(MyRowPosition3)("numejemplares") = TextBox16.Text
                MyDataTbl3.Rows(MyRowPosition3)("anoprimeraedicion") = TextBox17.Text
                MyDataTbl3.Rows(MyRowPosition3)("cantidad") = TextBox18.Text
                MyDataTbl3.Rows(MyRowPosition3)("portada") = "C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\PortadaLibro\" & TextBox6.Text & ".jpg"
                MyDatAdp3.Update(MyDataTbl3)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call REFRESCA3()
        BTN_GRABAR3.Enabled = False
        BTN_NUEVO3.Enabled = True
        ComboBox1.Text = "" 'Borra el texto del combo
        BTN_BUSCAR3.Enabled = False
        BTN_PORTADA.Enabled = True
    End Sub

    Private Sub BTN_BUSCAR3_Click(sender As Object, e As EventArgs) Handles BTN_BUSCAR3.Click
        Dim sebusca As String
        Cursor.Current = Cursors.WaitCursor
        MyRowPosition3 = 0
        Call showrecords3()

        sebusca = UCase(InputBox("Mete el título del libro", "SE BUSCA"))
        For I = 0 To MyDataTbl3.Rows.Count
            Call showrecords3()
            'se busca por textbox7.text porque es el registro "TÍTULO" (de libros)
            If InStr(TextBox7.Text, sebusca) <> 0 Then
                GoTo sal
            End If
            MyRowPosition3 = I
        Next
        MsgBox("No esta")


        'El cursor del mouse le da forma de actualizar
sal:

        Cursor.Current = Cursors.Default



    End Sub

    Private Sub BTN_BORRAR3_Click(sender As Object, e As EventArgs) Handles BTN_BORRAR3.Click
        Dim R As Integer 'VARIABLE R PARA EL SI = 6 O EL NO =7 DEL MSGBOX
        R = MsgBox("¿Estas seguro de querer borrar el registro?", vbYesNo, "BORRAR REGISTRO")
        If R = 7 Then Exit Sub
        If MyDataTbl3.Rows.Count <> 0 Then  'SI HAY FILAS ENTONCES
            MyDataTbl3.Rows(MyRowPosition3).Delete()
            MyDatAdp3.Update(MyDataTbl3)
            MyRowPosition3 = 0
            Me.showrecords3()
            Call REFRESCA3()
            ComboBox1.Text = "" 'Borra el texto del combo
        End If
    End Sub





    '______________________________________________________________________________________________
    'CARGA EN EL TEXTBOX 7 (TÍTULO DEL LIBRO) AL SELECCIONAR DESDE EL TEXTO DEL ÍNDICE DEL COMBOBOX
    '______________________________________________________________________________________________
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        MyRowPosition3 = 0

        For I = 0 To MyDataTbl3.Rows.Count
            Call showrecords3()
            If InStr(TextBox7.Text, ComboBox1.Text) <> 0 Then
                GoTo IR
            End If
            '(Carga en el textbox 7, título, el libro señalado en el combo)

            MyRowPosition3 = I

        Next
IR:

    End Sub

    '_______________________________________________________________________________________________
    'CARGA EN EL TEXTBOX 31 (NOMBRE DEL SOCIO) AL SELECCIONAR DESDE EL TEXTO DEL ÍNDICE DEL COMBOBOX
    '_______________________________________________________________________________________________
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        MyRowPosition5 = 0

        For I = 0 To MyDataTbl5.Rows.Count
            Call showrecords5()
            If InStr(TextBox31.Text, ComboBox2.Text) <> 0 Then
                GoTo IR
            End If
            '(Carga en el textbox 31, nombre socio, el socio señalado en el combo)

            MyRowPosition5 = I

        Next
IR:
    End Sub








    '_______________________________________________________________________________________________
    'TAB 4 - EDITORIALES
    '_______________________________________________________________________________________________


    '-----------------------
    'ESTÉTICA DE LOS BOTONES, QUE APAREZCA UN BOCADILLO DE TEXTO = ARRAY DE TOOLTIP
    'Hacer después



    'CREO LOS BOTONES 

    Private Sub BTN_IRALPRIMERO4_Click(sender As Object, e As EventArgs) Handles BTN_IRALPRIMERO4.Click
        MyRowPosition4 = 0
        Call showrecords4()
    End Sub

    Private Sub BTN_RETROCEDER4_Click(sender As Object, e As EventArgs) Handles BTN_RETROCEDER4.Click
        If MyRowPosition4 = 0 Then Exit Sub
        MyRowPosition4 = MyRowPosition4 - 1
        Call showrecords4()
    End Sub

    Private Sub BTN_AVANZAR4_Click(sender As Object, e As EventArgs) Handles BTN_AVANZAR4.Click
        If MyDataTbl4.Rows.Count - 1 = MyRowPosition4 Then Exit Sub
        MyRowPosition4 = MyRowPosition4 + 1
        Call showrecords4()
    End Sub

    Private Sub BTN_IRALULTIMO4_Click(sender As Object, e As EventArgs) Handles BTN_IRALULTIMO4.Click
        MyRowPosition4 = MyDataTbl4.Rows.Count - 1
        Call showrecords4()
    End Sub

    Private Sub BTN_NUEVO4_Click(sender As Object, e As EventArgs) Handles BTN_NUEVO4.Click
        Dim MyNewRow As DataRow = MyDataTbl4.NewRow()
        MyDataTbl4.Rows.Add(MyNewRow)
        MyRowPosition4 = MyDataTbl4.Rows.Count - 1
        Me.showrecords4()
        TextBox20.Select()
        BTN_GRABAR4.Enabled = True
        BTN_NUEVO4.Enabled = False
    End Sub


    Private Sub BTN_GRABAR4_Click(sender As Object, e As EventArgs) Handles BTN_GRABAR4.Click
        'DATA ADAPTER
        Try
            If MyDataTbl4.Rows.Count <> 0 Then
                MyDataTbl4.Rows(MyRowPosition4)("codeditorial") = TextBox20.Text
                MyDataTbl4.Rows(MyRowPosition4)("nombreeditorial") = UCase(TextBox21.Text)
                MyDataTbl4.Rows(MyRowPosition4)("direccion") = UCase(TextBox22.Text)
                MyDataTbl4.Rows(MyRowPosition4)("poblacion") = UCase(TextBox23.Text)
                MyDataTbl4.Rows(MyRowPosition4)("provincia") = UCase(TextBox24.Text)
                MyDataTbl4.Rows(MyRowPosition4)("codpostal") = TextBox25.Text
                MyDataTbl4.Rows(MyRowPosition4)("pais") = UCase(TextBox26.Text)
                MyDataTbl4.Rows(MyRowPosition4)("telefono") = TextBox27.Text
                MyDataTbl4.Rows(MyRowPosition4)("email") = TextBox28.Text
                MyDataTbl4.Rows(MyRowPosition4)("web") = UCase(TextBox29.Text)

                MyDatAdp4.Update(MyDataTbl4)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call REFRESCA4()
        BTN_GRABAR4.Enabled = False
        BTN_NUEVO4.Enabled = True
        Call cargaeditoriales()
    End Sub

    Private Sub BTN_BUSCAR4_Click(sender As Object, e As EventArgs) Handles BTN_BUSCAR4.Click
        Dim sebusca As String
        Cursor.Current = Cursors.WaitCursor
        MyRowPosition4 = 0
        Call showrecords4()

        sebusca = UCase(InputBox("Mete el nombre de la editorial", "SE BUSCA"))
        For I = 0 To MyDataTbl4.Rows.Count
            Call showrecords4()
            'se busca por textbox21.text porque es el registro "NOMBRE EDITORIAL" 
            If InStr(TextBox21.Text, sebusca) <> 0 Then
                GoTo sal
            End If
            MyRowPosition4 = I
        Next
        MsgBox("No esta")

sal:
        'El cursor del mouse le da forma de actualizar
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub BTN_BORRAR4_Click(sender As Object, e As EventArgs) Handles BTN_BORRAR4.Click
        Dim R As Integer 'VARIABLE R PARA EL SI = 6 O EL NO =7 DEL MSGBOX
        R = MsgBox("¿Estas seguro de querer borrar el registro?", vbYesNo, "BORRAR REGISTRO")
        If R = 7 Then Exit Sub
        If MyDataTbl4.Rows.Count <> 0 Then  'SI HAY FILAS ENTONCES
            MyDataTbl4.Rows(MyRowPosition4).Delete()
            MyDatAdp4.Update(MyDataTbl4)
            MyRowPosition4 = 0
            Me.showrecords4()
            Call REFRESCA4()
        End If
    End Sub


    '_______________________________________________________________________________________________
    'TAB 5 - SOCIO
    '_______________________________________________________________________________________________


    '-----------------------
    'ESTÉTICA DE LOS BOTONES, QUE APAREZCA UN BOCADILLO DE TEXTO = ARRAY DE TOOLTIP
    'Hacer después



    'CREO LOS BOTONES 

    Private Sub BTN_IRALPRIMERO5_Click(sender As Object, e As EventArgs) Handles BTN_IRALPRIMERO5.Click
        MyRowPosition5 = 0
        Call showrecords5()
        Call mirafoto()
        ComboBox2.Text = "" 'Borra el texto del combo
    End Sub

    Private Sub BTN_RETROCEDER5_Click(sender As Object, e As EventArgs) Handles BTN_RETROCEDER5.Click
        If MyRowPosition5 = 0 Then Exit Sub
        MyRowPosition5 = MyRowPosition5 - 1
        Call showrecords5()
        Call mirafoto()
        ComboBox2.Text = "" 'Borra el texto del combo
    End Sub

    Private Sub BTN_AVANZAR5_Click(sender As Object, e As EventArgs) Handles BTN_AVANZAR5.Click
        If MyDataTbl5.Rows.Count - 1 = MyRowPosition5 Then Exit Sub
        MyRowPosition5 = MyRowPosition5 + 1
        Call showrecords5()
        Call mirafoto()
        ComboBox2.Text = "" 'Borra el texto del combo
    End Sub

    Private Sub BTN_IRALULTIMO5_Click(sender As Object, e As EventArgs) Handles BTN_IRALULTIMO5.Click
        MyRowPosition5 = MyDataTbl5.Rows.Count - 1
        Call showrecords5()
        Call mirafoto()
        ComboBox2.Text = "" 'Borra el texto del combo
    End Sub

    Private Sub BTN_NUEVO5_Click(sender As Object, e As EventArgs) Handles BTN_NUEVO5.Click
        Dim MyNewRow As DataRow = MyDataTbl5.NewRow()
        MyDataTbl5.Rows.Add(MyNewRow)
        MyRowPosition5 = MyDataTbl5.Rows.Count - 1
        Me.showrecords5()
        TextBox30.Select()
        BTN_GRABAR5.Enabled = True
        BTN_NUEVO5.Enabled = False
        Label43.Visible = False 'Hace invisible en Label43 "SOCIOS"(para botón nuevo)
        ComboBox2.Visible = False 'Hace invisible en combobox2 (para botón nuevo)
        ComboBox2.Enabled = False 'Desabilita el Combobox2 (VAYA SER QUE TOQUEN AHÍ Y LE MANDE A OTRO SOCIO SIN QUERER)
        BTN_BUSCAR5.Enabled = False
        'IMAGEN POR DEFECTO
        PictureBox2.Image = Image.FromFile("C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\Imagenes_por_defecto\Imagen_base_socio.png")
    End Sub


    Private Sub BTN_GRABAR5_Click(sender As Object, e As EventArgs) Handles BTN_GRABAR5.Click
        'IMAGEN POR DEFECTO
        PictureBox2.Image = Image.FromFile("C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\Imagenes_por_defecto\Imagen_base_socio.png")
        'DATA ADAPTER
        Try
            If MyDataTbl5.Rows.Count <> 0 Then
                MyDataTbl5.Rows(MyRowPosition5)("codsocio") = TextBox30.Text
                MyDataTbl5.Rows(MyRowPosition5)("nombre_y_apellidos") = UCase(TextBox31.Text)
                MyDataTbl5.Rows(MyRowPosition5)("dni") = TextBox32.Text
                MyDataTbl5.Rows(MyRowPosition5)("direccion") = UCase(TextBox33.Text)
                MyDataTbl5.Rows(MyRowPosition5)("poblacion") = UCase(TextBox34.Text)
                MyDataTbl5.Rows(MyRowPosition5)("provincia") = TextBox35.Text
                MyDataTbl5.Rows(MyRowPosition5)("cp") = TextBox36.Text
                MyDataTbl5.Rows(MyRowPosition5)("telefono") = TextBox37.Text
                MyDataTbl5.Rows(MyRowPosition5)("fecha") = TextBox38.Text
                MyDataTbl5.Rows(MyRowPosition5)("email") = TextBox39.Text
                MyDataTbl5.Rows(MyRowPosition5)("foto") = "C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\Fotosocios\" & TextBox30.Text & ".jpg"

                MyDatAdp5.Update(MyDataTbl5)
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Call REFRESCA5()
        BTN_GRABAR5.Enabled = False
        BTN_NUEVO5.Enabled = True
        BTN_BUSCAR5.Enabled = False
        BTN_FOTO.Enabled = True
        ComboBox2.Text = ""
    End Sub

    Private Sub BTN_BUSCAR5_Click(sender As Object, e As EventArgs) Handles BTN_BUSCAR5.Click
        Dim sebusca As String
        Cursor.Current = Cursors.WaitCursor
        MyRowPosition5 = 0
        Call showrecords5()

        sebusca = UCase(InputBox("Mete el nombre del socio", "SE BUSCA"))
        For I = 0 To MyDataTbl5.Rows.Count
            Call showrecords5()
            'se busca por textbox31.text porque es el registro "NOMBRE Y APELLIDOS SOCIO" 
            If InStr(TextBox31.Text, sebusca) <> 0 Then
                GoTo sal
            End If
            MyRowPosition5 = I
        Next
        MsgBox("No esta")

sal:
        'El cursor del mouse le da forma de actualizar
        Cursor.Current = Cursors.Default
    End Sub

    Private Sub BTN_BORRAR5_Click(sender As Object, e As EventArgs) Handles BTN_BORRAR5.Click
        Dim R As Integer 'VARIABLE R PARA EL SI = 6 O EL NO =7 DEL MSGBOX
        R = MsgBox("¿Estas seguro de querer borrar el registro?", vbYesNo, "BORRAR REGISTRO")
        If R = 7 Then Exit Sub
        If MyDataTbl5.Rows.Count <> 0 Then  'SI HAY FILAS ENTONCES
            MyDataTbl5.Rows(MyRowPosition5).Delete()
            MyDatAdp5.Update(MyDataTbl5)
            MyRowPosition5 = 0
            Me.showrecords5()
            Call REFRESCA5()
        End If
        ComboBox2.Text = "" 'Borra el texto del combo
    End Sub





    'FUNCIÓN PARA VISUALIZAR LAS PORTADAS DE LOS LIBROS
    Sub miraportada()
        PictureBox1.Image = Image.FromFile(TextBox19.Text)
    End Sub


    'FUNCIÓN PARA VISUALIZAR LAS FOTOS DE LOS SOCIOS
    Sub mirafoto()
        PictureBox2.Image = Image.FromFile(TextBox40.Text)
    End Sub




    Private Sub BTN_PORTADA_Click(sender As Object, e As EventArgs) Handles BTN_PORTADA.Click
        'CARGAR Y GUARDAR PORTADA EN EL DIRECTORIO
        PictureBox1.Image = Image.FromFile("C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\Imagenes_por_defecto\Imagen_base_portada.png")
        Dim origen As String
        Dim destino As String = "C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\PortadaLibro\" & TextBox6.Text & ".jpg"

        Dim openFileDialog1 As New OpenFileDialog

        openFileDialog1.InitialDirectory = "c:/"
        openFileDialog1.Filter = "JPG | *.jpg| PNG | *.png"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = False
        If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            origen = openFileDialog1.FileName
            My.Computer.FileSystem.CopyFile(origen, destino,
    Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
    Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
        End If
        BTN_PORTADA.Enabled = False

    End Sub

    Private Sub BTN_FOTO_Click(sender As Object, e As EventArgs) Handles BTN_FOTO.Click
        'CARGAR Y GUARDAR IMAGEN DEL SOCIO EN EL DIRECTORIO
        PictureBox2.Image = Image.FromFile("C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\Imagenes_por_defecto\Imagen_base_socio.png")
        Dim origen As String
        Dim destino As String = "C:\Users\Cristina\source\repos\BIBLIOTECA_CON_CODIGO_ODBC\Fotosocios\" & TextBox30.Text & ".jpg"

        Dim openFileDialog1 As New OpenFileDialog

        openFileDialog1.InitialDirectory = "c:/"
        openFileDialog1.Filter = "JPG | *.jpg| PNG | *.png"
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = False
        If openFileDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            origen = openFileDialog1.FileName
            My.Computer.FileSystem.CopyFile(origen, destino,
    Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
    Microsoft.VisualBasic.FileIO.UICancelOption.DoNothing)
        End If
        BTN_FOTO.Enabled = False
    End Sub



    'DEPURANDO EL PROGRAMA

    Private Sub TextBox19_TextChanged(sender As Object, e As EventArgs) Handles TextBox19.TextChanged
        '________________________________________________________________________________________________
        'CARGA LA IMAGEN EN EL PICTUREBOX1 (PORTADA DEL LIBRO) AL SELECCIONAR DESDE EL TEXTO DEL COMBOBOX
        If TextBox19.Text <> Nothing Then
            Call miraportada()
        End If
    End Sub

    Private Sub TextBox40_TextChanged(sender As Object, e As EventArgs) Handles TextBox40.TextChanged
        '________________________________________________________________________________________________
        'CARGA LA IMAGEN EN EL PICTUREBOX2 (IMAGEN DEL SOCIO) AL SELECCIONAR DESDE EL TEXTO DEL COMBOBOX
        If TextBox40.Text <> Nothing Then
            Call mirafoto()
        End If
    End Sub




    '_________________________________
    'FUNCIONES QUE RELLENAN LOS COMBOS
    '_________________________________

    Sub cargatemas()
        Dim a = MyRowPosition
        CMB_TEMA.Items.Clear()
        'MyDataTbl porque es la tabla temas
        For I = 0 To MyDataTbl.Rows.Count - 1
            MyRowPosition = I
            'Añade al combo las líneas de showrecords porque son los registros "cod tema & nombre tema" 
            CMB_TEMA.Items.Add(MyDataTbl.Rows(MyRowPosition)("codtema").ToString() & " - " & MyDataTbl.Rows(MyRowPosition)("nombretema").ToString())
        Next
        MyRowPosition = a
    End Sub

    Sub cargaeditoriales()
        Dim a = MyRowPosition4
        CMB_EDITORIAL.Items.Clear()
        'MyDataTbl4 porque es la tabla editoriales
        For I = 0 To MyDataTbl4.Rows.Count - 1
            MyRowPosition4 = I
            'Añade al combo las líneas de showrecords porque son los registros "codeditorial & nombreeditorial" 
            CMB_EDITORIAL.Items.Add(MyDataTbl4.Rows(MyRowPosition4)("codeditorial").ToString() & " - " & MyDataTbl4.Rows(MyRowPosition4)("nombreeditorial").ToString())
        Next
        MyRowPosition4 = a
    End Sub

    Sub cargaautores()
        Dim a = MyRowPosition2
        CMB_AUTOR.Items.Clear()
        'MyDataTbl2 porque es la tabla temas
        For I = 0 To MyDataTbl2.Rows.Count - 1
            MyRowPosition2 = I
            'Añade al combo las líneas de showrecords porque son los registros "cod tema & nombre tema" 
            CMB_AUTOR.Items.Add(MyDataTbl2.Rows(MyRowPosition2)("codautor").ToString() & " - " & MyDataTbl2.Rows(MyRowPosition2)("nombre").ToString())
        Next
        MyRowPosition2 = a
    End Sub




    '______________________________________________________________________________________________________
    'AL SELECCIONAR UN ELEMENTO DE LA IZQUIERDA DEL ÍNDICE DEL COMBO, ENLAZALO A SU TEXTBOX CORRESPONDIENTE
    '______________________________________________________________________________________________________

    Private Sub CMB_AUTOR_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_AUTOR.SelectedIndexChanged
        BTN_GRABAR2.Enabled = True
        Dim CUANTOS = InStr(CMB_AUTOR.SelectedItem, " ") - 1
        'Textbox8.text (coautor) Microsoft Visual Basic coge de la izquierda
        TextBox8.Text = Microsoft.VisualBasic.Left(CMB_AUTOR.SelectedItem, CUANTOS)
    End Sub

    Private Sub CMB_EDITORIAL_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_EDITORIAL.SelectedIndexChanged
        BTN_GRABAR4.Enabled = True
        Dim CUANTOS = InStr(CMB_EDITORIAL.SelectedItem, " ") - 1
        'Textbox9.text (codeditorial) Microsoft Visual Basic coge de la izquierda
        TextBox9.Text = Microsoft.VisualBasic.Left(CMB_EDITORIAL.SelectedItem, CUANTOS)
    End Sub
    Private Sub CMB_TEMA_SelectedIndexChanged(sender As Object, e As EventArgs) Handles CMB_TEMA.SelectedIndexChanged
        BTN_GRABAR.Enabled = True
        Dim CUANTOS = InStr(CMB_TEMA.SelectedItem, " ") - 1
        'Textbox10.text (codtema) Microsoft Visual Basic coge de la izquierda
        TextBox10.Text = Microsoft.VisualBasic.Left(CMB_TEMA.SelectedItem, CUANTOS)
    End Sub


    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        CMB_AUTOR.SelectedIndex = CMB_AUTOR.FindString(TextBox8.Text)
    End Sub
    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        CMB_EDITORIAL.SelectedIndex = CMB_EDITORIAL.FindString(TextBox9.Text)
    End Sub
    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        CMB_TEMA.SelectedIndex = CMB_TEMA.FindString(TextBox10.Text)
    End Sub

End Class
