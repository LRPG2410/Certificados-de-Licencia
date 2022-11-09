#Region "IMPORTS  */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"

':::ADO.NET es un conjunto de componentes que pueden ser usado para acceder a datos y a servicios de datos. Es parte de la biblioteca de clases base que están incluidas en el Microsoft .NET Framework.
Imports System.Data ':::Ofrece acceso a clases que representan la arquitectura de ADO.NET. ADO.NET permite crear componentes que administran datos de varios orígenes de datos con eficacia.

':::OLE DB es la sigla de Object Linking and Embedding for Databases, es una tecnología usada para tener acceso a fuentes de información, o bases de datos.
Imports System.Data.OleDb ':::Es el proveedor de datos .NET Framework para OLE DB.

Imports Microsoft.Office.Interop.Word ':::Para activar esta referencia se va a proyecto luego en agregar referencia com y se agrega Microsoft Word 9.0 Object Library dependiendo de la version de office
Imports System.IO ':::Contiene tipos que permiten leer y escribir en los archivos y secuencias de datos, así como tipos que proporcionan compatibilidad básica con los archivos y directorios.
Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Imports System.Windows.Controls
Imports System.Windows.Forms.VisualStyles.VisualStyleElement
Imports System.Windows.Media.Media3D
Imports stdole
Imports System.Runtime.InteropServices.ComTypes
Imports System.Security.Authentication.ExtendedProtection


#End Region

Public Class Form1
    ':::String -> Cadena de caracteres
    ':::Integer -> Valores de tipo entero

#Region "VARIABLES GLOBALES */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"

    ':::Se le asigna a la variable conexion la clase de OleDbConnection que representa una conexión única a un origen de datos.
    Dim conexion As New OleDbConnection

    ':::Se le asigna a la variable comandos la clase OleDbCommand que representa una instrucción SQL o un procedimiento almacenado que se va a ejecutar en un origen de datos.
    Dim comandos As New OleDbCommand

    ':::Se le asigna a la variable reader la clase OleDbDataReader que proporciona el modo de lectura de una secuencia de filas de datos de tipo sólo avance de un origen de datos.
    Dim reader As OleDbDataReader

    ':::Se le asigna a la variable MSWord el enlace entre vs y la aplicación de word.
    Dim MSWord As New Word.Application

    ':::Se le asigna a la variable documento el enlace entre vs y el documento de word.
    Dim documento As Word.Document

    ':::Se utiliza junto con la clase OleDbCommand. Sirve para enlazar la base de datos.
    Dim ConnectionString As String

#End Region

#Region "PROCEDIMIENTOS */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"

    ':::PROCEDIMIENTO para la consulta y le indicamos que debe pedir 2 parametros para ejecutarse correctamente (tabla, access)
    Sub consulta(ByVal tabla As DataGridView, ByVal access As String)
        ':::Instruccion Try para capturar errores
        Try

            ':::Creamos el objeto DataAdapter y le pasamos los dos parametros (Instruccion, conexión)
            '::OleDbDaraAdapter funciona como un puente mediante Fill para cargar datos del origen de datos en DataSet y usar Update para enviar los cambios realizados en el DataSet origen de datos.
            Dim DA As New OleDbDataAdapter(access, conexion)

            ':::Creamos el objeto DataTable que recibe la informacion del DataAdapter
            '::DataTable se utilizan para representar las tablas de un DataSet.
            '::Los datos son locales de la aplicación basada en .NET en la que residen, pero se pueden llenar desde un origen de datos como SQL Server mediante un DataAdapter, como en este caso.
            Dim DT As New Data.DataTable

            ':::Pasamos la informacion del DataAdapter al DataTable mediante la propiedad Fill, antes mencionada.
            DA.Fill(DT)

            ':::Ahora mostramos los datos en el DataGridView
            tabla.DataSource = DT
        Catch ex As Exception
            MsgBox("No se logro realizar la consulta por: " & ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    ':::PROCEDIMIENTO para Agregar, Actualizar y Eliminar ademas le indicamos que debe pedir 2 parametros para ejecutarse correctamente (tabla, access)
    Sub operaciones(ByVal tabla As DataGridView, ByVal access As String)
        ':::Instruccion Try para capturar errores
        Try

            ':::Creamos nuestro objeto de tipo Command que almacenara nuestras instrucciones, necesita 2 parametros (access, conexion)
            Dim cmd As New OleDbCommand(access, conexion)

            ':::Ejecutamos la instruccion mediante la propiedad ExecuteNonQuery del command
            '::Realiza operaciones de catálogo (por ejemplo, consultar la estructura de una base de datos o crear objetos de base de datos como tablas)
            '::o cambiar los datos de una base de datos sin usar, mediante DataSet, la ejecución de instrucciones UPDATE, INSERT o DELETE.
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("No se logro realizar la operación por: " & ex.Message, MsgBoxStyle.Critical, "ERROR")
        End Try
    End Sub

    ':::PROCEDIMIENTO para actualizar el DataGridView al momento de accionar cualquier boton
    Sub actualizar()
        ':Instrucción "Select * from [tabla] where [nombrecelda1]='" & [nombreherramienta1] & "'"
        ':Dim access As String = "Select * from Certificados where NombrePaciente='" & txtNPaciente.Text & "'"
        ':Instrucción "Select * from [tabla]"

        ':::Creamos la variable access que guarda la instruccion de tipo SQL
        Dim access As String = "Select * from Datos_Paciente"

        ':::Accedemos a nuestro procedimiento "consulta" y le pasamos los dos 2 parametros (dgvTabla, access)
        Me.consulta(dgvTabla, access)
        num()
    End Sub

    ':::PROCEDIMIENTO que lleva el contador de la cantidad de pacientes.
    Sub num()

        Dim Sql As String = "Select max(Contador) from Datos_Paciente"
        Dim cmd As New OleDbCommand(Sql, conexion)

        ':::ExecuteScalar es una operación que se utiliza para devolver un valor único
        Dim codigo = cmd.ExecuteScalar

        ':::IsDBNull es una operación que se utiliza para verificar si hay registros en la base de datos, lanza un valor booleano
        If IsDBNull(codigo) Then
            Me.lcontador.Text = "1"
        Else

            ':::CStr se encarga de convertir un valor numérico en un tipo String
            Me.lcontador.Text = CStr(codigo) + 1
        End If

    End Sub

    ':::PROCEDIMIENTO para generar los reportes en word
    Sub reporte()

        Try
            ':::FileCopy se encarga de copiar una plantilla ya creada en word y crean un nuevo documento igual a la plantilla, pero con los datos que toma del formulario de vb.
            FileCopy("C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\Reportes\Plantilla.docx",
            "C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\Reportes\" & txtNPaciente.Text & ".docx")

            ':::A documento se le asigna el documento de word que este especificado en la ruta.
            documento = MSWord.Documents.Open("C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\Reportes\" & txtNPaciente.Text & ".docx")

            'MsgBox("EL INFORME FUE GUARDADO EN C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\Reportes CON EL NOMBRE DE '" & txtNPaciente.Text & " '",
            'MsgBoxStyle.Information, "INFORMACIÓN")

            ':::Aqui hacemos referencia al documento de word, al marcador al que queres que se exporte la info y por ultimo la herramienta de la cual se tomara la información
            documento.Bookmarks.Item("lcontador").Range.Text = lcontador.Text
            documento.Bookmarks.Item("lFReporte").Range.Text = lFReporte.Text
            documento.Bookmarks.Item("cbProfesional").Range.Text = cbProfesional.SelectedItem
            documento.Bookmarks.Item("txtTransito").Range.Text = txtTransito.Text
            documento.Bookmarks.Item("txtSalud").Range.Text = txtSalud.Text
            documento.Bookmarks.Item("txtOftal").Range.Text = txtOftal.Text

            documento.Bookmarks.Item("txtNPaciente").Range.Text = txtNPaciente.Text & " " & txtAPaciente.Text
            documento = MSWord.Documents.Open("C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\Reportes\" & txtNPaciente.Text & ".docx")
            documento.Bookmarks.Item("txtDpi").Range.Text = txtDpi.Text
            documento.Bookmarks.Item("txtDate1").Range.Text = txtDate1.Text

            ':::Se copia al portapapeles la imagen que este en el PictureBox
            Clipboard.SetImage(Me.pbFoto.Image)

            ':::Se busca en el documento de word el marcador y se pega la imagen anteriormente copiada
            documento.Range.Bookmarks.Item("prueba").Range.Paste()

            ':::Se crea una condición para que, dependiendo de la elección, se marque en el documento de word.
            If cbGenero.SelectedItem = "Femenino" Then
                ':::Se marca una "X" en el lugar del marcador indicado.
                documento.Bookmarks.Item("cbGenero1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("cbGenero2").Range.Text = "X"
            End If

            documento.Bookmarks.Item("cbDepartamento").Range.Text = cbDepartamento.SelectedItem
            documento.Bookmarks.Item("cbMunicipio").Range.Text = cbMunicipio.SelectedItem
            documento.Bookmarks.Item("txtResidencia").Range.Text = txtResidencia.Text

            documento.Bookmarks.Item("cbAgudeza1").Range.Text = cbAgudeza1.Text
            documento.Bookmarks.Item("cbAgudeza2").Range.Text = cbAgudeza2.Text
            documento.Bookmarks.Item("cbAgudeza3").Range.Text = cbAgudeza3.Text

            If rbVision1.Checked = True Then
                documento.Bookmarks.Item("rbVision1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("rbVision2").Range.Text = "X"
            End If

            If rbSensibilidad1.Checked = True Then
                documento.Bookmarks.Item("rbSensibilidad1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("rbSensibilidad2").Range.Text = "X"
            End If

            If rbPrueba1.Checked = True Then
                documento.Bookmarks.Item("rbPrueba1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("rbPrueba2").Range.Text = "X"
            End If

            If rbSeg1.Checked = True Then
                documento.Bookmarks.Item("rbSeg1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("rbSeg2").Range.Text = "X"
            End If

            If rbAnteojos1.Checked = True Then
                documento.Bookmarks.Item("rbAnteojos1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("rbAnteojos2").Range.Text = "X"
            End If

            If rbLentes1.Checked = True Then
                documento.Bookmarks.Item("rbLentes1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("rbLentes2").Range.Text = "X"
            End If

            documento.Bookmarks.Item("nudCentral1").Range.Text = nudCentral1.Value
            documento.Bookmarks.Item("nudCentral2").Range.Text = nudCentral2.Value
            If rbCentral1.Checked = True Then
                documento.Bookmarks.Item("rbCentral1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("rbCentral2").Range.Text = "X"
            End If

            documento.Bookmarks.Item("nudPeriferico1").Range.Text = nudPeriferico1.Value
            documento.Bookmarks.Item("nudPeriferico2").Range.Text = nudPeriferico2.Value
            If rbCentral1.Checked = True Then
                documento.Bookmarks.Item("rbPeriferico1").Range.Text = "X"
            Else
                documento.Bookmarks.Item("rbPeriferico2").Range.Text = "X"
            End If

            If (cbA.Checked = True) Then
                documento.Bookmarks.Item("cbA").Range.Text = "X"
            End If

            If (cbB.Checked = True) Then
                documento.Bookmarks.Item("cbB").Range.Text = "X"
            End If

            If (cbE.Checked = True) Then
                documento.Bookmarks.Item("cbE").Range.Text = "X"
            End If

            If (cbC.Checked = True) Then
                documento.Bookmarks.Item("cbC").Range.Text = "X"
            End If

            If (cbM.Checked = True) Then
                documento.Bookmarks.Item("cbM").Range.Text = "X"
            End If

            If (cbNinguna.Checked = True) Then
                documento.Bookmarks.Item("cbNinguna").Range.Text = "X"
            End If

            documento.Bookmarks.Item("rtb1").Range.Text = rtb1.Text

        Catch ex As Exception

        Finally
            documento.Save()
            'MSWord.Quit()
        End Try

    End Sub

    ':::PROCEDIMIENTO para limpiar los campos
    Sub limpiar()
        ':::Limpiar los TextBox
        'txtTransito.Text = ""
        'txtSalud.Text = ""
        'txtOftal.Text = ""
        txtAPaciente.Text = ""
        txtNPaciente.Text = ""
        txtDpi.Text = ""
        txtResidencia.Text = ""
        txtEdad.Visible = False
        txtEdad.Text = ""

        ':::Limpiar los ComboBox
        'cbProfesional.SelectedValue = Nothing
        'cbProfesional.Text = Nothing
        cbDepartamento.SelectedValue = Nothing
        cbDepartamento.Text = Nothing
        cbMunicipio.SelectedValue = Nothing
        cbMunicipio.Text = Nothing
        cbGenero.SelectedValue = Nothing
        cbGenero.Text = Nothing
        cbAgudeza1.Text = "20/25"
        cbAgudeza2.Text = "20/25"
        cbAgudeza3.Text = "20/25"

        ':::Limpiar la fecha
        txtDate1.Text = Nothing

        ':::Limpiar los RadioButton
        rbVision1.Checked = False
        rbVision2.Checked = False
        rbCentral1.Checked = False
        rbCentral2.Checked = False
        rbPeriferico1.Checked = False
        rbPeriferico2.Checked = False
        rbSensibilidad1.Checked = False
        rbSensibilidad2.Checked = False
        rbPrueba1.Checked = False
        rbPrueba2.Checked = False
        rbSeg1.Checked = False
        rbSeg2.Checked = False
        rbAnteojos1.Checked = False
        rbAnteojos2.Checked = False
        rbLentes1.Checked = False
        rbLentes2.Checked = False

        ':::Limpiar el RichTextBox
        Me.rtb1.ForeColor = Color.Gray
        rtb1.Text = "Observaciones: "

        ':::Limpiar los NumericUpDown
        nudCentral1.Value = "20"
        nudCentral2.Value = "20"
        nudPeriferico1.Value = "85"
        nudPeriferico2.Value = "85"

        ':::Limpiar los CheckBox
        cbA.Checked = False
        cbB.Checked = False
        cbC.Checked = False
        cbE.Checked = False
        cbM.Checked = False

        ':::Limpiar la foto
        pbFoto.ImageLocation = "C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\usuario.png"

        ':::Instrucción para evitar que el # de paciente y correlativo del paciente actual se mezclen
        lpaciente.Visible = False
        lpaciente.Text = Nothing
        num()
        lcontador.Visible = True
    End Sub

    ':::PROCEDIMIENTO para generar el reporte en pdf
    Sub generarpdf()
        Dim wordApplication As New Microsoft.Office.Interop.Word.Application
        Dim wordDocument As Microsoft.Office.Interop.Word.Document = Nothing
        Dim outputFilename As String
        Try
            wordDocument = wordApplication.Documents.Open("C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\Reportes\" & txtNPaciente.Text & ".docx")
            outputFilename = System.IO.Path.ChangeExtension("C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\Reportes\" & txtNPaciente.Text & ".docx", "pdf")

            If Not wordDocument Is Nothing Then
                wordDocument.ExportAsFixedFormat(outputFilename, Microsoft.Office.Interop.Word.WdExportFormat.wdExportFormatPDF, True, Microsoft.Office.Interop.Word.WdExportOptimizeFor.wdExportOptimizeForOnScreen, Microsoft.Office.Interop.Word.WdExportRange.wdExportAllDocument, 0, 0, Microsoft.Office.Interop.Word.WdExportItem.wdExportDocumentContent, True, True, Microsoft.Office.Interop.Word.WdExportCreateBookmarks.wdExportCreateNoBookmarks, True, True, False)
            End If
        Catch ex As Exception
            'TODO: handle exception
        Finally
            If Not wordDocument Is Nothing Then
                wordDocument.Close(False)
                wordDocument = Nothing
            End If

            If Not wordApplication Is Nothing Then
                wordApplication.Quit()
                wordApplication = Nothing
            End If
        End Try
    End Sub

#End Region

#Region "INSTRUCCIONES  */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"

    ':::Switch de los DEPARTAMENTOS y municipios
    Private Sub cbDepartamento_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cbDepartamento.SelectedIndexChanged

        ':::Nueva variable local departamento
        Dim departamento As String

        ':::Asignamos la variabel departamento al comboBox para que capture la información
        departamento = cbDepartamento.Text

        ':::Limpiamos el contenido del comboBox de municipios para que se "actualice" la información dependiendo de la opción elegida en departamentos
        cbMunicipio.Items.Clear()

        ':::Instrucción Select Case para detectar lo que se selecciona en el comboBox de departamentos
        Select Case departamento

            ':::En cada Case va un departamento, el cual, dependiendo la elección de departamento se cambiaran los items en el comboBox de municipio
            Case "Alta Verapaz"
                cbMunicipio.Items.Add("Cobán")
                cbMunicipio.Items.Add("San Pedro Carchá")
                cbMunicipio.Items.Add("San Juan Chamelco")
                cbMunicipio.Items.Add("San Cristóbal Verapaz")
                cbMunicipio.Items.Add("Tactic")
                cbMunicipio.Items.Add("Tucurú")
                cbMunicipio.Items.Add("Tamahú")
                cbMunicipio.Items.Add("Panzós")
                cbMunicipio.Items.Add("Senahú")
                cbMunicipio.Items.Add("Cahabón")
                cbMunicipio.Items.Add("Lanquín")
                cbMunicipio.Items.Add("Chahal")
                cbMunicipio.Items.Add("Fray Bartolomé de las Casas")
                cbMunicipio.Items.Add("Chisec")
                cbMunicipio.Items.Add("Santa Cruz Verapaz")
                cbMunicipio.Items.Add("Santa Catalina La Tinta")
                cbMunicipio.Items.Add("Raxruhá")
            Case "Baja Verapaz"
                cbMunicipio.Items.Add("Salamá")
                cbMunicipio.Items.Add("Cubulco")
                cbMunicipio.Items.Add("Granados")
                cbMunicipio.Items.Add("Purulhá")
                cbMunicipio.Items.Add("Rabinal")
                cbMunicipio.Items.Add("San Jerónimo")
                cbMunicipio.Items.Add("San Miguel Chicaj")
                cbMunicipio.Items.Add("Santa Cruz El Chol")
            Case "Chimaltenango"
                cbMunicipio.Items.Add("Acatenango")
                cbMunicipio.Items.Add("Chimaltenango")
                cbMunicipio.Items.Add("El Tejar")
                cbMunicipio.Items.Add("Parramos")
                cbMunicipio.Items.Add("Patzicía")
                cbMunicipio.Items.Add("Patzún")
                cbMunicipio.Items.Add("San Andrés Itzapa")
                cbMunicipio.Items.Add("San José Poaquil")
                cbMunicipio.Items.Add("San Juan Comalapa")
                cbMunicipio.Items.Add("San Martín Jilotepeque")
                cbMunicipio.Items.Add("San Miguel Pochuta")
                cbMunicipio.Items.Add("San Pedro Yepocapa")
                cbMunicipio.Items.Add("Santa Apolonia")
                cbMunicipio.Items.Add("Santa Cruz Balanyá")
                cbMunicipio.Items.Add("Tecpán")
                cbMunicipio.Items.Add("Zaragoza")
            Case "Chiquimula"
                cbMunicipio.Items.Add("Camotán")
                cbMunicipio.Items.Add("Chiquimula")
                cbMunicipio.Items.Add("Concepción Las Minas")
                cbMunicipio.Items.Add("Esquipulas")
                cbMunicipio.Items.Add("Ipala")
                cbMunicipio.Items.Add("Jocotán")
                cbMunicipio.Items.Add("Olopa")
                cbMunicipio.Items.Add("Quezaltepeque")
                cbMunicipio.Items.Add("San Jacinto")
                cbMunicipio.Items.Add("San José La Arada")
                cbMunicipio.Items.Add("San Juan Ermita")
            Case "Guatemala"
                cbMunicipio.Items.Add("Amatitlán")
                cbMunicipio.Items.Add("Chinautla")
                cbMunicipio.Items.Add("Chuarrancho")
                cbMunicipio.Items.Add("Ciudad de Guatemala")
                cbMunicipio.Items.Add("Fraijanes")
                cbMunicipio.Items.Add("Mixco")
                cbMunicipio.Items.Add("Palencia")
                cbMunicipio.Items.Add("San José del Golfo")
                cbMunicipio.Items.Add("San José Pinula")
                cbMunicipio.Items.Add("San Juan Sacatepéquez")
                cbMunicipio.Items.Add("San Miguel Petapa")
                cbMunicipio.Items.Add("San Pedro Ayampuc")
                cbMunicipio.Items.Add("San Pedro Sacatepéquez")
                cbMunicipio.Items.Add("San Raymundo")
                cbMunicipio.Items.Add("Santa Catarina Pinula")
                cbMunicipio.Items.Add("Villa Canales")
                cbMunicipio.Items.Add("Villa Nueva")
            Case "El Progreso"
                cbMunicipio.Items.Add("El Jícaro")
                cbMunicipio.Items.Add("Guastatoya")
                cbMunicipio.Items.Add("Morazán")
                cbMunicipio.Items.Add("San Agustín Acasaguastlán")
                cbMunicipio.Items.Add("San Antonio La Paz")
                cbMunicipio.Items.Add("San Cristóbal Acasaguastlán")
                cbMunicipio.Items.Add("Sanarate")
                cbMunicipio.Items.Add("Sansare")
            Case "Escuintla"
                cbMunicipio.Items.Add("Escuintla")
                cbMunicipio.Items.Add("Guanagazapa")
                cbMunicipio.Items.Add("Iztapa")
                cbMunicipio.Items.Add("La Democracia")
                cbMunicipio.Items.Add("La Gomera")
                cbMunicipio.Items.Add("Masagua")
                cbMunicipio.Items.Add("Nueva Concepción")
                cbMunicipio.Items.Add("Palín")
                cbMunicipio.Items.Add("San José")
                cbMunicipio.Items.Add("San Vicente Pacaya")
                cbMunicipio.Items.Add("Santa Lucía Cotzumalguapa")
                cbMunicipio.Items.Add("Sipacate")
                cbMunicipio.Items.Add("Siquinalá")
                cbMunicipio.Items.Add("Tiquisate")
            Case "Huehuetenango"
                cbMunicipio.Items.Add("Aguacatán")
                cbMunicipio.Items.Add("Chiantla")
                cbMunicipio.Items.Add("Colotenango")
                cbMunicipio.Items.Add("Concepción Huista")
                cbMunicipio.Items.Add("Cuilco")
                cbMunicipio.Items.Add("Huehuetenango")
                cbMunicipio.Items.Add("Jacaltenango")
                cbMunicipio.Items.Add("La Democracia")
                cbMunicipio.Items.Add("La Libertad")
                cbMunicipio.Items.Add("Malacatancito")
                cbMunicipio.Items.Add("Nentón")
                cbMunicipio.Items.Add("Petatán")
                cbMunicipio.Items.Add("San Antonio Huista")
                cbMunicipio.Items.Add("San Gaspar Ixchil")
                cbMunicipio.Items.Add("San Ildefonso Ixtahuacán")
                cbMunicipio.Items.Add("San Juan Atitán")
                cbMunicipio.Items.Add("San Juan Ixcoy")
                cbMunicipio.Items.Add("San Mateo Ixtatán")
                cbMunicipio.Items.Add("San Miguel Acatán")
                cbMunicipio.Items.Add("San Pedro Nécta")
                cbMunicipio.Items.Add("San Pedro Soloma")
                cbMunicipio.Items.Add("San Rafael La Independencia")
                cbMunicipio.Items.Add("San Rafael Pétzal")
                cbMunicipio.Items.Add("San Sebastián Coatán")
                cbMunicipio.Items.Add("San Sebastián Huehuetenango")
                cbMunicipio.Items.Add("Santa Ana Huista")
                cbMunicipio.Items.Add("Santa Bárbara")
                cbMunicipio.Items.Add("Santa Cruz Barillas")
                cbMunicipio.Items.Add("Santa Eulalia")
                cbMunicipio.Items.Add("Santiago Chimaltenango")
                cbMunicipio.Items.Add("Tectitán")
                cbMunicipio.Items.Add("Todos Santos Cuchumatán")
                cbMunicipio.Items.Add("Unión Cantinil")
            Case "Izabal"
                cbMunicipio.Items.Add("El Estor")
                cbMunicipio.Items.Add("Livingston")
                cbMunicipio.Items.Add("Los Amates")
                cbMunicipio.Items.Add("Morales")
                cbMunicipio.Items.Add("Puerto Barrios")
            Case "Jalapa"
                cbMunicipio.Items.Add("Jalapa")
                cbMunicipio.Items.Add("Mataquescuintla")
                cbMunicipio.Items.Add("Monjas")
                cbMunicipio.Items.Add("San Carlos Alzatate")
                cbMunicipio.Items.Add("San Luis Jilotepeque")
                cbMunicipio.Items.Add("San Manuel Chaparrón")
                cbMunicipio.Items.Add("San Pedro Pinula")
            Case "Jutiapa"
                cbMunicipio.Items.Add("Agua Blanca")
                cbMunicipio.Items.Add("Asunción Mita")
                cbMunicipio.Items.Add("Atescatempa")
                cbMunicipio.Items.Add("Comapa")
                cbMunicipio.Items.Add("Conguaco")
                cbMunicipio.Items.Add("El Adelanto")
                cbMunicipio.Items.Add("El Progreso")
                cbMunicipio.Items.Add("Jalpatagua")
                cbMunicipio.Items.Add("Jerez")
                cbMunicipio.Items.Add("Jutiapa")
                cbMunicipio.Items.Add("Moyuta")
                cbMunicipio.Items.Add("Pasaco")
                cbMunicipio.Items.Add("Quesada")
                cbMunicipio.Items.Add("San José Acatempa")
                cbMunicipio.Items.Add("Santa Catarina Mita")
                cbMunicipio.Items.Add("Yupiltepeque")
                cbMunicipio.Items.Add("Zapotitlán")
            Case "Petén"
                cbMunicipio.Items.Add("Dolores")
                cbMunicipio.Items.Add("El Chal")
                cbMunicipio.Items.Add("Flores")
                cbMunicipio.Items.Add("La Libertad")
                cbMunicipio.Items.Add("Las Cruces")
                cbMunicipio.Items.Add("Melchor de Mencos")
                cbMunicipio.Items.Add("Poptún")
                cbMunicipio.Items.Add("San Andrés")
                cbMunicipio.Items.Add("San Benito")
                cbMunicipio.Items.Add("San Francisco")
                cbMunicipio.Items.Add("San José")
                cbMunicipio.Items.Add("San Luis")
                cbMunicipio.Items.Add("Santa Ana")
            Case "Quetzaltenango"
                cbMunicipio.Items.Add("Almolonga")
                cbMunicipio.Items.Add("Cabricán")
                cbMunicipio.Items.Add("Cajolá")
                cbMunicipio.Items.Add("Cantel")
                cbMunicipio.Items.Add("Coatepeque")
                cbMunicipio.Items.Add("Colomba Costa Cuca")
                cbMunicipio.Items.Add("Concepción Chiquirichapa")
                cbMunicipio.Items.Add("El Palmar")
                cbMunicipio.Items.Add("Flores Costa Cuca")
                cbMunicipio.Items.Add("Génova")
                cbMunicipio.Items.Add("Huitán")
                cbMunicipio.Items.Add("La Esperanza")
                cbMunicipio.Items.Add("Olintepeque")
                cbMunicipio.Items.Add("Palestina de Los Altos")
                cbMunicipio.Items.Add("Quetzaltenango")
                cbMunicipio.Items.Add("Salcajá")
                cbMunicipio.Items.Add("San Carlos Sija")
                cbMunicipio.Items.Add("San Francisco La Unión")
                cbMunicipio.Items.Add("San Juan Ostuncalco")
                cbMunicipio.Items.Add("San Martín Sacatepéquez")
                cbMunicipio.Items.Add("San Mateo")
                cbMunicipio.Items.Add("San Miguel Sigüilá")
                cbMunicipio.Items.Add("Sibilia")
                cbMunicipio.Items.Add("Zunil")
            Case "Quiché"
                cbMunicipio.Items.Add("Canillá")
                cbMunicipio.Items.Add("Chajul")
                cbMunicipio.Items.Add("Chicamán")
                cbMunicipio.Items.Add("Chiché")
                cbMunicipio.Items.Add("Chichicastenango (Santo Tomás Chichicastenango)")
                cbMunicipio.Items.Add("Chinique")
                cbMunicipio.Items.Add("Cunén")
                cbMunicipio.Items.Add("Ixcán")
                cbMunicipio.Items.Add("Joyabaj")
                cbMunicipio.Items.Add("Nebaj")
                cbMunicipio.Items.Add("Pachalum")
                cbMunicipio.Items.Add("Patzité")
                cbMunicipio.Items.Add("Sacapulas")
                cbMunicipio.Items.Add("San Andrés Sajcabajá")
                cbMunicipio.Items.Add("San Antonio Ilotenango")
                cbMunicipio.Items.Add("San Bartolomé Jocotenango")
                cbMunicipio.Items.Add("San Juan Cotzal")
                cbMunicipio.Items.Add("San Pedro Jocopilas")
                cbMunicipio.Items.Add("Santa Cruz del Quiché")
                cbMunicipio.Items.Add("Uspantán")
                cbMunicipio.Items.Add("Zacualpa")
            Case "Retalhuleu"
                cbMunicipio.Items.Add("Champerico")
                cbMunicipio.Items.Add("El Asintal")
                cbMunicipio.Items.Add("Nuevo San Carlos")
                cbMunicipio.Items.Add("Retalhuleu")
                cbMunicipio.Items.Add("San Andrés Villa Seca")
                cbMunicipio.Items.Add("San Felipe")
                cbMunicipio.Items.Add("San Martín Zapotitlán")
                cbMunicipio.Items.Add("San Sebastián")
                cbMunicipio.Items.Add("Santa Cruz Muluá")
            Case "Sacatepequez"
                cbMunicipio.Items.Add("Alotenango")
                cbMunicipio.Items.Add("Ciudad Vieja")
                cbMunicipio.Items.Add("Jocotenango")
                cbMunicipio.Items.Add("Antigua Guatemala")
                cbMunicipio.Items.Add("Magdalena Milpas Altas")
                cbMunicipio.Items.Add("Pastores")
                cbMunicipio.Items.Add("San Antonio Aguas Calientes")
                cbMunicipio.Items.Add("San Bartolomé Milpas Altas")
                cbMunicipio.Items.Add("San Lucas Sacatepéquez")
                cbMunicipio.Items.Add("San Miguel Dueñas")
                cbMunicipio.Items.Add("Santa Catarina Barahona")
                cbMunicipio.Items.Add("Santa Lucía Milpas Altas")
                cbMunicipio.Items.Add("Santa María de Jesús")
                cbMunicipio.Items.Add("Santiago Sacatepéquez")
                cbMunicipio.Items.Add("Santo Domingo Xenacoj")
                cbMunicipio.Items.Add("Sumpango")
            Case "San Marcos"
                cbMunicipio.Items.Add("Ayutla")
                cbMunicipio.Items.Add("Catarina")
                cbMunicipio.Items.Add("Comitancillo")
                cbMunicipio.Items.Add("Concepción Tutuapa")
                cbMunicipio.Items.Add("El Quetzal")
                cbMunicipio.Items.Add("El Tumbador")
                cbMunicipio.Items.Add("Esquipulas Palo Gordo")
                cbMunicipio.Items.Add("Ixchiguán")
                cbMunicipio.Items.Add("La Blanca")
                cbMunicipio.Items.Add("La Reforma")
                cbMunicipio.Items.Add("Malacatán")
                cbMunicipio.Items.Add("Nuevo Progreso")
                cbMunicipio.Items.Add("Ocós")
                cbMunicipio.Items.Add("Pajapita")
                cbMunicipio.Items.Add("Río Blanco")
                cbMunicipio.Items.Add("San Antonio Sacatepéquez")
                cbMunicipio.Items.Add("San Cristóbal Cucho")
                cbMunicipio.Items.Add("San José El Rodeo")
                cbMunicipio.Items.Add("San José Ojetenam")
                cbMunicipio.Items.Add("San Lorenzo")
                cbMunicipio.Items.Add("San Marcos")
                cbMunicipio.Items.Add("San Miguel Ixtahuacán")
                cbMunicipio.Items.Add("San Pablo")
                cbMunicipio.Items.Add("San Pedro Sacatepéquez")
                cbMunicipio.Items.Add("San Rafael Pie de la Cuesta")
                cbMunicipio.Items.Add("Sibinal")
                cbMunicipio.Items.Add("Sipacapa")
                cbMunicipio.Items.Add("Tacaná")
                cbMunicipio.Items.Add("Tajumulco")
                cbMunicipio.Items.Add("Tejutla")
            Case "Santa Rosa"
                cbMunicipio.Items.Add("Barberena")
                cbMunicipio.Items.Add("Casillas")
                cbMunicipio.Items.Add("Chiquimulilla")
                cbMunicipio.Items.Add("Cuilapa")
                cbMunicipio.Items.Add("Guazacapán")
                cbMunicipio.Items.Add("Nueva Santa Rosa")
                cbMunicipio.Items.Add("Oratorio")
                cbMunicipio.Items.Add("Pueblo Nuevo Viñas")
                cbMunicipio.Items.Add("San Juan Tecuaco")
                cbMunicipio.Items.Add("San Rafael las Flores")
                cbMunicipio.Items.Add("Santa Cruz Naranjo")
                cbMunicipio.Items.Add("Santa María Ixhuatán")
                cbMunicipio.Items.Add("Santa Rosa de Lima")
                cbMunicipio.Items.Add("Taxisco")
            Case "Sololá"
                cbMunicipio.Items.Add("Concepción")
                cbMunicipio.Items.Add("Nahualá")
                cbMunicipio.Items.Add("Panajachel")
                cbMunicipio.Items.Add("San Andrés Semetabaj")
                cbMunicipio.Items.Add("San Antonio Palopó")
                cbMunicipio.Items.Add("San José Chacayá")
                cbMunicipio.Items.Add("San Juan La Laguna")
                cbMunicipio.Items.Add("San Lucas Tolimán")
                cbMunicipio.Items.Add("San Marcos La Laguna")
                cbMunicipio.Items.Add("San Pablo La Laguna")
                cbMunicipio.Items.Add("San Pedro La Laguna")
                cbMunicipio.Items.Add("Santa Catarina Ixtahuacán")
                cbMunicipio.Items.Add("Santa Catarina Palopó")
                cbMunicipio.Items.Add("Santa Clara La Laguna")
                cbMunicipio.Items.Add("Santa Cruz La Laguna")
                cbMunicipio.Items.Add("Santa Lucía Utatlán")
                cbMunicipio.Items.Add("Santa María Visitación")
                cbMunicipio.Items.Add("Santiago Atitlán")
                cbMunicipio.Items.Add("Sololá")
            Case "Suchitepequez"
                cbMunicipio.Items.Add("Chicacao")
                cbMunicipio.Items.Add("Cuyotenango")
                cbMunicipio.Items.Add("Mazatenango")
                cbMunicipio.Items.Add("Patulul")
                cbMunicipio.Items.Add("Pueblo Nuevo")
                cbMunicipio.Items.Add("Río Bravo")
                cbMunicipio.Items.Add("Samayac")
                cbMunicipio.Items.Add("San Antonio Suchitepéquez")
                cbMunicipio.Items.Add("San Bernardino")
                cbMunicipio.Items.Add("San Francisco Zapotitlán")
                cbMunicipio.Items.Add("San Gabriel")
                cbMunicipio.Items.Add("San José El Idolo")
                cbMunicipio.Items.Add("San José La Máquina")
                cbMunicipio.Items.Add("San Juan Bautista")
                cbMunicipio.Items.Add("San Lorenzo")
                cbMunicipio.Items.Add("San Miguel Panán")
                cbMunicipio.Items.Add("San Pablo Jocopilas")
                cbMunicipio.Items.Add("Santa Bárbara")
                cbMunicipio.Items.Add("Santo Domingo Suchitepéquez")
                cbMunicipio.Items.Add("Santo Tomás La Unión")
                cbMunicipio.Items.Add("Zunilito")
            Case "Totonicapán"
                cbMunicipio.Items.Add("Momostenango")
                cbMunicipio.Items.Add("San Andrés Xecul")
                cbMunicipio.Items.Add("San Bartolo")
                cbMunicipio.Items.Add("San Cristóbal Totonicapán")
                cbMunicipio.Items.Add("San Francisco El Alto")
                cbMunicipio.Items.Add("Santa Lucía La Reforma")
                cbMunicipio.Items.Add("Santa María Chiquimula")
                cbMunicipio.Items.Add("Totonicapán")
            Case "Zacapa"
                cbMunicipio.Items.Add("Cabañas")
                cbMunicipio.Items.Add("Estanzuela")
                cbMunicipio.Items.Add("Gualán")
                cbMunicipio.Items.Add("Huité")
                cbMunicipio.Items.Add("La Unión")
                cbMunicipio.Items.Add("Río Hondo")
                cbMunicipio.Items.Add("San Diego")
                cbMunicipio.Items.Add("San Jorge")
                cbMunicipio.Items.Add("Teculután")
                cbMunicipio.Items.Add("Usumatlán")
                cbMunicipio.Items.Add("Zacapa")
        End Select
    End Sub

    ':::CERRAR el programa
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        ':::Instrucción para detener la ejecución del programa
        Me.Close()
    End Sub

    ':::LIMPIAR el programa
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        limpiar()
    End Sub

    ':::Instrucción para mostrar la información en el formulario desde el DataGridView
    Private Sub dgvTabla_CellDoubleClick(sender As Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles dgvTabla.CellDoubleClick
        Dim I As Integer = e.RowIndex
        Me.cbProfesional.Text = dgvTabla.Rows(I).Cells(0).Value.ToString
        Me.txtTransito.Text = dgvTabla.Rows(I).Cells(1).Value.ToString
        Me.txtSalud.Text = dgvTabla.Rows(I).Cells(2).Value.ToString
        Me.txtOftal.Text = dgvTabla.Rows(I).Cells(3).Value.ToString
        Me.txtAPaciente.Text = dgvTabla.Rows(I).Cells(4).Value.ToString
        Me.txtNPaciente.Text = dgvTabla.Rows(I).Cells(5).Value.ToString
        Me.txtDpi.Text = dgvTabla.Rows(I).Cells(6).Value.ToString
        Me.cbDepartamento.Text = dgvTabla.Rows(I).Cells(7).Value.ToString
        Me.cbMunicipio.Text = dgvTabla.Rows(I).Cells(8).Value.ToString
        Me.txtDate1.Value = dgvTabla.Rows(I).Cells(9).Value.ToString
        Me.cbGenero.Text = dgvTabla.Rows(I).Cells(10).Value.ToString
        Me.txtResidencia.Text = dgvTabla.Rows(I).Cells(11).Value.ToString
        Me.cbAgudeza1.Text = dgvTabla.Rows(I).Cells(12).Value.ToString
        Me.cbAgudeza2.Text = dgvTabla.Rows(I).Cells(13).Value.ToString
        Me.cbAgudeza3.Text = dgvTabla.Rows(I).Cells(14).Value.ToString

        ':::Instrucción para que los RadioButton se marquen correctamente al momento de enviarlos al formulario OPCION [1]
        If CStr(dgvTabla.Rows(I).Cells(15).Value) = "NORMAL" Then
            rbVision1.Checked = True
        Else
            rbVision2.Checked = True
        End If

        ':::OPCION [2]
        'rbVision1.Checked = CStr(dgvTabla.Rows(I).Cells(15).Value) = "NORMAL"
        'rbVision2.Checked = CStr(dgvTabla.Rows(I).Cells(15).Value) = "DEFICIENTE"

        Me.nudCentral1.Value = dgvTabla.Rows(I).Cells(16).Value.ToString
        Me.nudCentral2.Value = dgvTabla.Rows(I).Cells(17).Value.ToString
        If CStr(dgvTabla.Rows(I).Cells(18).Value) = "NORMAL" Then
            rbCentral1.Checked = True
        Else
            rbCentral2.Checked = True
        End If

        Me.nudPeriferico1.Value = dgvTabla.Rows(I).Cells(19).Value.ToString
        Me.nudPeriferico2.Value = dgvTabla.Rows(I).Cells(20).Value.ToString
        If CStr(dgvTabla.Rows(I).Cells(21).Value) = "NORMAL" Then
            rbPeriferico1.Checked = True
        Else
            rbPeriferico2.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(22).Value) = "NORMAL" Then
            rbSensibilidad1.Checked = True
        Else
            rbSensibilidad2.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(23).Value) = "SI HAY ESTEREOPSIS" Then
            rbPrueba1.Checked = True
        Else
            rbPrueba2.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(24).Value) = "SI" Then
            rbSeg1.Checked = True
        Else
            rbSeg2.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(25).Value) = "SI" Then
            rbAnteojos1.Checked = True
        Else
            rbAnteojos2.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(26).Value) = "SI" Then
            rbLentes1.Checked = True
        Else
            rbLentes2.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(27).Value) = "A" Then
            cbA.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(28).Value) = "B" Then
            cbB.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(29).Value) = "E" Then
            cbE.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(30).Value) = "C" Then
            cbC.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(31).Value) = "M" Then
            cbM.Checked = True
        End If

        If CStr(dgvTabla.Rows(I).Cells(32).Value) = "NINGUNA" Then
            cbNinguna.Checked = True
        End If

        Me.rtb1.Text = dgvTabla.Rows(I).Cells(33).Value.ToString

        Me.lpaciente.Text = dgvTabla.Rows(I).Cells(35).Value.ToString
        Me.lcontador.Visible = False
        Me.lpaciente.Location = New Drawing.Point(840, 13)
        Me.lpaciente.Visible = True

    End Sub

    ':::Instrucción que verifica el estado del CheckBox NINGUNA
    Private Sub cbNinguna_Click(sender As Object, e As System.EventArgs) Handles cbNinguna.Click
        If (cbNinguna.Checked = True) Then
            cbA.Checked = False
            cbB.Checked = False
            cbE.Checked = False
            cbC.Checked = False
            cbM.Checked = False
        End If
    End Sub

    ':::Instrucción que verifica el estado del CheckBox A
    Private Sub cbA_Click(sender As Object, e As System.EventArgs) Handles cbA.Click
        If (cbA.Checked = True) Then
            cbNinguna.Checked = False
        End If
    End Sub

    ':::Instrucción que verifica el estado del CheckBox B
    Private Sub cbB_Click(sender As Object, e As System.EventArgs) Handles cbB.Click
        If (cbB.Checked = True) Then
            cbNinguna.Checked = False
        End If
    End Sub

    ':::Instrucción que verifica el estado del CheckBox E
    Private Sub cbE_Click(sender As Object, e As System.EventArgs) Handles cbE.Click
        If (cbE.Checked = True) Then
            cbNinguna.Checked = False
        End If
    End Sub

    ':::Instrucción que verifica el estado del CheckBox C
    Private Sub cbC_Click(sender As Object, e As System.EventArgs) Handles cbC.Click
        If (cbC.Checked = True) Then
            cbNinguna.Checked = False
        End If
    End Sub

    ':::Instrucción que verifica el estado del CheckBox M
    Private Sub cbM_Click(sender As Object, e As System.EventArgs) Handles cbM.Click
        If (cbM.Checked = True) Then
            cbNinguna.Checked = False
        End If
    End Sub

    ':::Boton para LLENAR todos los campos
    Private Sub btnAgregar_Click(sender As System.Object, e As System.EventArgs)
        'cbProfesional.Text = "Luis Porras"
        'txtTransito.Text = "123"
        'txtSalud.Text = "456"
        'txtOftal.Text = "789"
        'txtNPaciente.Text = "Alejandro Polares"
        'txtDpi.Text = "457"
        'cbDepartamento.Text = "Guatemala"
        'cbMunicipio.Text = "Ciudad de Guatemala"
        'cbGenero.Text = "Masculino"
        'txtResidencia.Text = "Vivo por ahi"
        'cbAgudeza1.Text = "67"
        'cbAgudeza2.Text = "34"
        'cbAgudeza3.Text = "90"
        'rbVision1.Checked = True
        'nudCentral1.Value = "45"
        'nudCentral2.Value = "32"
        'rbCentral1.Checked = True
        'nudPeriferico1.Value = "5"
        'nudPeriferico2.Value = "57"
        'rbPeriferico1.Checked = True
        'rbSensibilidad1.Checked = True
        'rbPrueba1.Checked = True
        'rbSeg1.Checked = True
        'rbAnteojos1.Checked = True
        'rbLentes1.Checked = True
        'rtb1.Text = "Esto es una prueba"
    End Sub

    ':::Instrucción para abrir el explorador de archivos y buscar la FOTO
    Private Sub PictureBox1_Click(sender As System.Object, e As System.EventArgs) Handles pbFoto.Click

        ':::Variable que utlizamos para guardar la cadena que contiene la localización de la imagen.
        Dim curFileName As String = ""

        ':::Buscamos la imagen a grabar
        Dim SaveImage As Boolean = False
        Dim openDlg As OpenFileDialog = New OpenFileDialog()
        openDlg.Filter = "Todos los archivos JPEG|*.jpg"
        Dim filter As String = openDlg.Filter
        openDlg.Title = "Abrir archivos JPEG"
        If (openDlg.ShowDialog() = DialogResult.OK) Then
            curFileName = openDlg.FileName
            SaveImage = True
            'Mostrando la foto en el picture
            Me.pbFoto.ImageLocation = curFileName.ToString
            Label14.Text = curFileName
        Else
            Exit Sub
        End If
    End Sub

    ':::Instrucción para calcular la edad del paciente
    Private Sub txtDate1_ValueChanged(sender As Object, e As EventArgs) Handles txtDate1.ValueChanged
        ':::Variables para obtener el valor del día, mes y año actuales
        Dim DiaHoy As String = Date.Now.Date.Day
        Dim MesHoy As String = Date.Now.Date.Month
        Dim AnioHoy As String = Date.Now.Date.Year

        ':::Variables para obtener el valor del día, mes y año de nacimiento
        Dim DiaNacer As String = txtDate1.Value.Day
        Dim MesNacer As String = txtDate1.Value.Month
        Dim AnioNacer As String = txtDate1.Value.Year

        ':::La variable edad se encarga de restar el año actual con el año de nacimiento para obtener la "edad"
        Dim edad = AnioHoy - AnioNacer

        ':::La instrucción If se encarga de verificar que el día y mes actuales sean iguales a los de nacimiento para confirmar si ya cumplio años
        If (edad < 0) Then
            MsgBox("Ingrese una fecha de nacimiento valida", vbExclamation, "AVISO")
            txtDate1.Value = Date.Now.Date
        ElseIf (edad = 0) Then
            txtEdad.Text = "Tiene " & edad & " años."
        ElseIf (edad > 0) Then
            If (MesNacer < MesHoy) Then
                txtEdad.Text = "Tiene " & edad & " años."
            ElseIf (MesNacer > MesHoy) Then
                txtEdad.Text = "Tiene " & edad - 1 & " años."
            ElseIf (MesNacer = MesHoy) Then
                If (DiaNacer < DiaHoy) Then
                    txtEdad.Text = "Tiene " & edad & " años."
                ElseIf (DiaNacer = DiaHoy) Then
                    txtEdad.Text = "¡Cumpleaños! Tiene " & edad & " años."
                ElseIf (DiaNacer > DiaHoy) Then
                    txtEdad.Text = "Tiene " & edad - 1 & " años."
                End If
            End If
        End If

        '::::Hacemos visible el Label
        txtEdad.Visible = True
    End Sub

    ':::Instrucción para que solo acepte número en el txtTransito
    Private Sub txtTransito_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtTransito.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else : e.Handled = True
            MsgBox("INTRODUZCA SÓLO VALORES NÚMERICOS", vbCritical, "ERROR")
        End If
    End Sub

    ':::Instrucción para que solo acepte número en el txtSalud
    Private Sub txtSalud_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtSalud.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else : e.Handled = True
            MsgBox("INTRODUZCA SÓLO VALORES NÚMERICOS", vbCritical, "ERROR")
        End If
    End Sub

    ':::Instrucción para que solo acepte número en el txtOftal
    Private Sub txtOftal_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtOftal.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else : e.Handled = True
            MsgBox("INTRODUZCA SÓLO VALORES NÚMERICOS", vbCritical, "ERROR")
        End If
    End Sub

    ':::Instrucción para que solo acepte número en el txtDpi
    Private Sub txtDpi_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtDpi.KeyPress
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else : e.Handled = True
            MsgBox("INTRODUZCA SÓLO VALORES NÚMERICOS", vbCritical, "ERROR")
        End If
    End Sub

    ':::Instrucción para que solo acepte letras en el txtAPaciente
    Private Sub txtAPaciente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAPaciente.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else : e.Handled = True
            MsgBox("INTRODUZCA SÓLO LETRAS", vbCritical, "ERROR")
        End If
    End Sub

    ':::Instrucción para que solo acepte letras en el txtNPaciente
    Private Sub txtNPaciente_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtNPaciente.KeyPress
        If Char.IsLetter(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else : e.Handled = True
            MsgBox("INTRODUZCA SÓLO LETRAS", vbCritical, "ERROR")
        End If
    End Sub

    ':::Instrucción para que solo acepte números en el txtAPaciente
    Private Sub txtAgudeza1_KeyPress(sender As Object, e As KeyPressEventArgs)
        If Char.IsNumber(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsControl(e.KeyChar) Then
            e.Handled = False
        ElseIf Char.IsSeparator(e.KeyChar) Then
            e.Handled = False
        Else : e.Handled = True
            MsgBox("INTRODUZCA SÓLO NÚMEROS", vbCritical, "ERROR")
        End If
    End Sub

    ':::Instrucción para "placeholder" en el RichTextBox
    Private Sub rtb1_Click(sender As Object, e As EventArgs) Handles rtb1.Click
        If rtb1.Text = "Observaciones: " Then
            Me.rtb1.Text = Nothing
            Me.rtb1.ForeColor = Color.Black
        End If
    End Sub

    ':::Boton para imprimir los REPORTES en word
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        gbBotones.Visible = True
        rbGenerarpdf.Visible = True
        rbGenerarpdf.Location = New Drawing.Point(138, 25)
        rbGenerarword.Visible = True
        rbGenerarword.Location = New Drawing.Point(276, 25)
        rbAbrir.Visible = True
        rbAbrir.Location = New Drawing.Point(430, 25)
        Button6.Image = Nothing
        Button6.Image = pbCargando.Image
        GroupBox3.Enabled = False
        GroupBox2.Enabled = False
        txtAPaciente.Enabled = False
        txtDpi.Enabled = False
        cbDepartamento.Enabled = False
        cbMunicipio.Enabled = False
        txtDate1.Enabled = False
        cbGenero.Enabled = False
        txtResidencia.Enabled = False
        pbFoto.Enabled = False
        BtnGuardar.Enabled = False
        Button5.Enabled = False
        Button4.Enabled = False

        If (txtNPaciente.Text = Nothing) Then
            MsgBox("No ha ingresado el nombre del paciente", vbExclamation, "AVISO")
            pbSalir_Click(sender, e)
        Else
            If rbGenerarpdf.Checked = True Then
                generarpdf()
            ElseIf rbGenerarword.Checked = True Then
                reporte()
            ElseIf rbAbrir.Checked = True Then
                Call Shell("explorer.exe " & "C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\Reportes", vbNormalFocus)
            End If
        End If

    End Sub

    ':::"BOTON" para sali del submenu de la generacion de reportes y la busqueda de paciente.
    Private Sub pbSalir_Click(sender As Object, e As EventArgs) Handles pbSalir.Click
        Button6.Image = pbImpresora.Image
        Button4.Image = pbLupa.Image
        gbBotones.Visible = False

        rbGenerarpdf.Visible = False
        rbGenerarpdf.Checked = False
        rbGenerarword.Visible = False
        rbGenerarword.Checked = False
        rbAbrir.Visible = False
        rbAbrir.Checked = False

        rbNombre.Visible = False
        rbNombre.Checked = False
        rbApellido.Visible = False
        rbApellido.Checked = False
        rbDpi.Visible = False
        rbDpi.Checked = False
        rbTodos.Visible = False
        rbTodos.Checked = False

        GroupBox3.Enabled = True
        GroupBox2.Enabled = True
        txtAPaciente.Enabled = True
        txtNPaciente.Enabled = True
        txtDpi.Enabled = True
        cbDepartamento.Enabled = True
        cbMunicipio.Enabled = True
        txtDate1.Enabled = True
        cbGenero.Enabled = True
        txtResidencia.Enabled = True
        pbFoto.Enabled = True
        BtnGuardar.Enabled = True
        Button6.Enabled = True
        Button5.Enabled = True
        Button4.Enabled = True
    End Sub

    ':::Instrucción para colocar los #'s de registro cuando se selecciona al profesional en cbProfesional
    Private Sub cbProfesional_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cbProfesional.SelectedIndexChanged
        If cbProfesional.Text = "Ramiro Faillace Poggio" Then
            txtTransito.Text = "070213"
            txtSalud.Text = "404"
        End If

    End Sub

    ':::Instrucción para habilitar txtNPaciente si se selecciona en el menu de [BUSCAR]
    Private Sub rbNombre_Click(sender As Object, e As EventArgs) Handles rbNombre.Click
        txtNPaciente.Enabled = True
        txtAPaciente.Enabled = False
        txtDpi.Enabled = False
    End Sub

    ':::Instrucción para habilitar txtAPaciente si se selecciona en el menu de [BUSCAR]
    Private Sub rbApellido_Click(sender As Object, e As EventArgs) Handles rbApellido.Click
        txtNPaciente.Enabled = False
        txtAPaciente.Enabled = True
        txtDpi.Enabled = False
    End Sub

    ':::Instrucción para habilitar txtDpi si se selecciona en el menu de [BUSCAR]
    Private Sub rbDpi_Click(sender As Object, e As EventArgs) Handles rbDpi.Click
        txtNPaciente.Enabled = False
        txtAPaciente.Enabled = False
        txtDpi.Enabled = True
    End Sub

#End Region

#Region "BASE DE DATOS  */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"
    ':::CONEXIÓN a la base de datos Access
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ':::Instrucción para mostrar la fecha de hoy
        Me.lFecha.Text = Date.Now.Date
        Dim fechaReporte = Format(Date.Now.Date, "Long Date")
        lFReporte.Text = fechaReporte
        cbAgudeza1.Text = "20/25"
        cbAgudeza2.Text = "20/25"
        cbAgudeza3.Text = "20/25"

        Me.rtb1.ForeColor = Color.Gray
        Me.rtb1.Text = "Observaciones: "

        cbProfesional.Text = "Ramiro Faillace Poggio"

        ':::Instrucción Try para capturar errores
        Try

            ':::Usamos la variable conexion para el enlace a la base de datos
            conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\CERTIFICADO_DE_LICENCIA_BD.accdb"
            conexion.Open()
            num()
            MsgBox("SE CONECTO EXITOSAMENTE A LA BASE DE DATOS", vbInformation, "CORRECTO")
        Catch ex As Exception
            MsgBox("ERROR AL CONECTAR A LA BASE DE DATOS NO PODRA GUARDAR NIGUN DATO", vbCritical, "ERROR")
            MsgBox(ex.Message)
        End Try
    End Sub

    ':::Boton que ENVIA [GUARDAR] la información a la base de datos
    Private Sub BtnGuardar_Click(sender As System.Object, e As System.EventArgs) Handles BtnGuardar.Click
        ':::Instrucción Try para capturar errores
        Try

            ':::Usamos el comando y le indicamos con la instrucción "INSERT INTO [tabla] ([nombrecelda1],[nombrecelda2]) VALUES([@nombreherramienta1],[@nombreherramienta2]", conexion)
            comandos = New OleDbCommand("INSERT INTO Datos_Paciente (Pro_Nombre, Pro_regTransito, Pro_regSalud, Pro_regOft, Pac_Apellido, Pac_Nombre, Pac_Dpi, Pac_Departamento, Pac_Municipio, Pac_Nacimiento, Pac_Genero, Pac_Residencia, Res_Agudeza1, Res_Agudeza2, Res_Agudeza3, Res_Vision, Res_CampoCentralOD, Res_CampoCentralOI, Res_CampoCentral, Res_CampoPerifericoOD, Res_CampoPerifericoOI, Res_CampoPeriferico, Res_Sensibilidad, Res_Prueba, Res_Seg, Res_Anteojos, Res_Lentes, Res_Licencia1, Res_Licencia2, Res_Licencia3, Res_Licencia4, Res_Licencia5, Res_Licencia6, Res_Obs, Contador) VALUES (@cbProfesional, @txtTransito, @txtSalud, @txtOftal, @txtAPaciente, @txtNPaciente, @txtDpi, @cbDepartamento, @cbMunicipio, @txtDate1, @cbGenero, @txtResidencia, @cbAgudeza1, @cbAgudeza2, @cbAgudeza3, @rbVision1, @nudCentral1, @nudCentral2, @rbCentral1, @nudPeriferico1, @nudPeriferico2, @rbPeriferico1, @rbSensibilidad1, @rbPrueba1, @rbSeg1, @rbAnteojos1, @rbLentes1, @cbA, @cbB, @cbE, @cbC, @cbM, @cbNinguno, @rtb1, @)", conexion)

            ':::Instrucción comandos.Parameters.AddWithValue ([@nombrecelda1],[nombreherramienta1]) para agregar los datos a la tabla de la base de datos
            If cbProfesional.SelectedItem = Nothing Then
                MsgBox("Aún no ha selecciona a un médico", vbExclamation, "ERROR")
            Else
                comandos.Parameters.AddWithValue("@Pro_Nombre", cbProfesional.SelectedItem)
            End If

            comandos.Parameters.AddWithValue("@Pro_regTransito", txtTransito.Text)
            comandos.Parameters.AddWithValue("@Pro_regSalud", txtSalud.Text)
            comandos.Parameters.AddWithValue("@Pro_regOft", txtOftal.Text)

            If txtAPaciente.Text = Nothing Then
                MsgBox("Ingrese apellido del paciente", vbExclamation, "AVISO")
            Else
                comandos.Parameters.AddWithValue("@Pac_Apellido", txtAPaciente.Text)
            End If

            If txtNPaciente.Text = Nothing Then
                MsgBox("Ingrese nombre del paciente", vbExclamation, "AVISO")
            Else
                comandos.Parameters.AddWithValue("@Pac_Nombre", txtNPaciente.Text)
            End If

            If txtDpi.Text = Nothing Then
                MsgBox("Ingrese DPI del paciente", vbExclamation, "AVISO")
            Else
                comandos.Parameters.AddWithValue("@Pac_Dpi", txtDpi.Text)
            End If

            If cbDepartamento.SelectedItem = Nothing Then
                MsgBox("Debe ingresar el departamento en el que reside el paciente", vbExclamation, "AVISO")
            Else
                comandos.Parameters.AddWithValue("@Pac_Departamento", cbDepartamento.SelectedItem)
            End If

            If cbMunicipio.SelectedItem = Nothing Then
                MsgBox("Debe ingresar el municipio en el que reside el paciente", vbExclamation, "AVISO")
            Else
                comandos.Parameters.AddWithValue("@Pac_Municipio", cbMunicipio.SelectedItem)
            End If

            If txtDate1.Text = Nothing Then
                MsgBox("Debe ingresar la fecha de nacimiento del paciente", vbExclamation, "AVISO")
            Else
                comandos.Parameters.AddWithValue("@Pac_Nacimiento", txtDate1.Text)
            End If

            If cbGenero.SelectedItem = Nothing Then
                MsgBox("Debe ingresar el género del paciente", vbExclamation, "AVISO")
            Else
                comandos.Parameters.AddWithValue("@Pac_Genero", cbGenero.SelectedItem)
            End If

            If txtResidencia.Text = Nothing Then
                MsgBox("Debe ingresar la dirección en la que reside del paciente", vbExclamation, "AVISO")
            Else
                comandos.Parameters.AddWithValue("@Pac_Residencia", txtResidencia.Text)
            End If

            comandos.Parameters.AddWithValue("@Res_Agudeza1", cbAgudeza1.SelectedItem)
            comandos.Parameters.AddWithValue("@Res_Agudeza2", cbAgudeza2.SelectedItem)
            comandos.Parameters.AddWithValue("@Res_Agudeza3", cbAgudeza3.SelectedItem)

            If rbVision1.Checked = Nothing And rbVision2.Checked = Nothing Then
                MsgBox("Debe ingresar la visión de colores del paciente", vbExclamation, "AVISO")
            Else
                ':::Instrucción If para los RadioButton. Hace la condición de guardar el texto del RadioButton que esté marcado.
                If (rbVision1.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Vision", rbVision1.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Vision", rbVision2.Text)
                End If
            End If

            comandos.Parameters.AddWithValue("@Res_CampoCentralOD", nudCentral1.Value)
            comandos.Parameters.AddWithValue("@Res_CampoCentralOI", nudCentral2.Value)
            If rbCentral1.Checked = Nothing And rbCentral2.Checked = Nothing Then
                MsgBox("Debe ingresar el campo visual central del paciente", vbExclamation, "AVISO")
            Else
                If (rbCentral1.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_CampoCentral", rbCentral1.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_CampoCentral", rbCentral2.Text)
                End If
            End If

            comandos.Parameters.AddWithValue("@Res_CampoPerifericoOD", nudPeriferico1.Value)
            comandos.Parameters.AddWithValue("@Res_CampoPerifericoOI", nudPeriferico2.Value)
            If rbPeriferico1.Checked = Nothing And rbPeriferico2.Checked = Nothing Then
                MsgBox("Debe ingresar el campo visual periférico del paciente", vbExclamation, "AVISO")
            Else
                If (rbPeriferico1.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_CampoPeriferico", rbPeriferico1.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_CampoPeriferico", rbPeriferico2.Text)
                End If
            End If

            If rbSensibilidad1.Checked = Nothing And rbSensibilidad2.Checked = Nothing Then
                MsgBox("Debe seleccionar la sensibilidad al contraste del paciente", vbExclamation, "AVISO")
            Else
                If (rbSensibilidad1.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Sensibilidad", rbSensibilidad1.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Sensibilidad", rbSensibilidad2.Text)
                End If
            End If

            If rbPrueba1.Checked = Nothing And rbPrueba2.Checked = Nothing Then
                MsgBox("Debe seleccionar la sensibilidad al contraste del paciente", vbExclamation, "AVISO")
            Else
                If (rbPrueba1.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Prueba", rbPrueba1.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Prueba", rbPrueba2.Text)
                End If
            End If

            If rbSeg1.Checked = Nothing And rbSeg2.Checked = Nothing Then
                MsgBox("Debe seleccionar a 600 segundos del paciente", vbExclamation, "AVISO")
            Else
                If (rbSeg1.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Seg", rbSeg1.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Seg", rbSeg2.Text)
                End If
            End If

            If rbAnteojos1.Checked = Nothing And rbAnteojos2.Checked = Nothing Then
                MsgBox("Debe seleccionar si el paciente usa lentes o no.", vbExclamation, "AVISO")
            Else
                If (rbAnteojos1.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Anteojos", rbAnteojos1.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Anteojos", rbAnteojos2.Text)
                End If
            End If

            If rbLentes1.Checked = Nothing And rbLentes2.Checked = Nothing Then
                MsgBox("Debe seleccionar si el paciente usa lentes de contacto o no.", vbExclamation, "AVISO")
            Else
                If (rbLentes1.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Lentes", rbLentes1.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Lentes", rbLentes2.Text)
                End If
            End If

            If cbA.Checked = Nothing And cbB.Checked = Nothing And cbE.Checked = Nothing And cbC.Checked = Nothing And cbM.Checked = Nothing And cbNinguna.Checked = Nothing Then
                MsgBox("Debe seleccionar para que tipo de licencia se encuentra apto el paciente.", vbExclamation, "AVISO")
            Else
                If (cbA.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Licencia1", cbA.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Licencia1", cbA.Text = " ")
                End If

                If (cbB.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Licencia2", cbB.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Licencia2", cbB.Text = " ")
                End If
                If (cbE.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Licencia3", cbE.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Licencia3", cbE.Text = " ")
                End If
                If (cbC.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Licencia4", cbC.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Licencia4", cbC.Text = " ")
                End If
                If (cbM.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Licencia5", cbM.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Licencia5", cbM.Text = " ")
                End If

                If (cbNinguna.Checked = True) Then
                    comandos.Parameters.AddWithValue("@Res_Licencia6", cbNinguna.Text)
                Else
                    comandos.Parameters.AddWithValue("@Res_Licencia6", cbNinguna.Text = " ")
                End If
            End If

            comandos.Parameters.AddWithValue("@Res_Obs", rtb1.Text)
            comandos.Parameters.AddWithValue("@Contador", lcontador.Text)

            ':::Ejecutamos la instruccion mediante la propiedad ExecuteNonQuery del command
            comandos.ExecuteNonQuery()
            MsgBox("DATOS GUARDADOS EXITOSAMENTE", vbInformation, "CORRECTO")
            'reporte()
            actualizar()
            limpiar()
        Catch ex As Exception
            MsgBox("ERROR AL GUARDAR EL FORMULARIO", vbCritical, "ERROR")
            MsgBox(ex.Message)
        End Try
    End Sub

    ':::Boton que MUESTRA [VER] la información almacenada en la base de datos
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        gbBotones.Visible = True
        rbNombre.Visible = True
        rbNombre.Location = New Drawing.Point(228, 25)
        rbApellido.Visible = True
        rbApellido.Location = New Drawing.Point(318, 25)
        rbDpi.Visible = True
        rbDpi.Location = New Drawing.Point(175, 25)
        rbTodos.Visible = True
        rbTodos.Location = New Drawing.Point(413, 25)
        Button4.Image = Nothing
        Button4.Image = pbCargando.Image
        GroupBox3.Enabled = False
        GroupBox2.Enabled = False
        txtAPaciente.Enabled = False
        txtNPaciente.Enabled = False
        txtDpi.Enabled = False
        cbDepartamento.Enabled = False
        cbMunicipio.Enabled = False
        txtDate1.Enabled = False
        cbGenero.Enabled = False
        txtResidencia.Enabled = False
        pbFoto.Enabled = False
        BtnGuardar.Enabled = False
        Button6.Enabled = False
        Button5.Enabled = False
        Dim nombre As String = "Select * from Datos_Paciente where Pac_Nombre = '" & txtNPaciente.Text & "'"
        Dim apellido As String = "Select * from Datos_Paciente where Pac_Apellido = '" & txtAPaciente.Text & "'"
        Dim dpi As String = "Select * from Datos_Paciente where Pac_Dpi = '" & txtDpi.Text & "'"
        Dim access As String = "Select * from Datos_Paciente"

        If (rbNombre.Checked = True) Then
            If (txtNPaciente.Text = Nothing) Then
                MsgBox("No ha ingresado el nombre del paciente", vbExclamation, "AVISO")
                pbSalir_Click(sender, e)
            Else
                Me.consulta(dgvTabla, nombre)
            End If
        ElseIf (rbApellido.Checked = True) Then
            If (txtNPaciente.Text = Nothing) Then
                MsgBox("No ha ingresado el apellido del paciente", vbExclamation, "AVISO")
                pbSalir_Click(sender, e)
            Else
                Me.consulta(dgvTabla, apellido)
            End If
        ElseIf (rbDpi.Checked = True) Then
            If (txtNPaciente.Text = Nothing) Then
                MsgBox("No ha ingresado el DPI del paciente", vbExclamation, "AVISO")
                pbSalir_Click(sender, e)
            Else
                Me.consulta(dgvTabla, dpi)
            End If
        ElseIf (rbTodos.Checked = True) Then
            Me.consulta(dgvTabla, access)
        End If



        ':::Creamos la variable access que guarda la instruccion de tipo SQL

        ':::Instrucción "Select * from [tabla] where [nombrecelda1]='" & [nombreherramienta1] & "'"

        'Dim access As String = "Select * from Certificados where NombrePaciente='" & txtNPaciente.Text & "'"

        ':::Accedemos a nuestro procedimiento "consulta" y le pasamos los dos (2) parametros (dgvTabla, access)

    End Sub

    ':::Boton que ACTUALIZA [ACTUALIZAR] la informacion almacenada en la base de datos
    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        ':::Creamos la variable access que guardar la instruccion de tipo SQL

        ':::Instrucción "Update [tabla] set [nombrecelda1]='" & [nombreherramienta1] & "'" where [nombrecelda2]= " & [nombreherramienta1] & ""
        'Dim access As String = "Update Datos_Paciente Set Pac_Apellido = '" & txtAPaciente.Text & " where Contador = " & lpaciente.Text & ""
        Dim access As String = "Update Datos_Paciente Set Pac_Apellido='" & txtAPaciente.Text & "' where Contador=" & lcontador.Text & ""

        Me.operaciones(dgvTabla, access)

        actualizar()
        limpiar()
    End Sub

    ':::Boton que ELIMINA [ELIMINAR] la información de la base de datos
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ':::Creamos la variable access que guardar la instruccion de tipo SQL
        If MessageBox.Show("¿Seguro que quiere elimar este paciente?", "ELIMINACIÓN DE PACIENTE", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = vbYes Then
            Dim access As String = "Delete * From Datos_Paciente Where Pac_Nombre='" & txtNPaciente.Text & "'"
            Me.operaciones(dgvTabla, access)
            actualizar()
            limpiar()
        End If

        ':::Instrucción "Delete * From [tabla] Where [nombrecelda1]=" & [nombreherramienta1] & ""

    End Sub

#End Region

End Class