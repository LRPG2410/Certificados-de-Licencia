#Region "IMPORTS  */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"
':::ADO.NET es un conjunto de componentes que pueden ser usado para acceder a datos y a servicios de datos. Es parte de la biblioteca de clases base que están incluidas en el Microsoft .NET Framework.
Imports System.Data ':::Ofrece acceso a clases que representan la arquitectura de ADO.NET. ADO.NET permite crear componentes que administran datos de varios orígenes de datos con eficacia.

':::OLE DB es la sigla de Object Linking and Embedding for Databases, es una tecnología usada para tener acceso a fuentes de información, o bases de datos.
Imports System.Data.OleDb ':::Es el proveedor de datos .NET Framework para OLE DB.

':::Imports Microsoft.Office.Interop.Word ' para activar esta referencia se va a proyecto luego en agregar referencia com y se agrega Microsoft Word 9.0 Object Library dependiendo de la version de office

Imports System.IO ':::Contiene tipos que permiten leer y escribir en los archivos y secuencias de datos, así como tipos que proporcionan compatibilidad básica con los archivos y directorios.

':::Imports Microsoft.Office.interop

Imports System.Runtime.InteropServices
#End Region

Public Class Form1

#Region "VARIABLES GLOBALES */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"
    Public sRuta As String
    Dim conexion As New OleDbConnection
    Dim comandos As New OleDbCommand
    Dim registro As New DataSet
    'Dim MSWord As New Word.Application
    'Dim documento As Word.Document
    Dim connectionstring As String
    Private SelectedDevice As WIA.Device
    Private SavePath As String
    Private SavedFilePath As String
    Dim curFileName As String = ""
    Dim SaveImage As Boolean = False

#End Region

#Region "PROCEDIMIENTOS */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"

    ':::PROCEDIMIENTO para la consulta y le indicamos que debe pedir 2 parametros para ejecutarse correctamente (tabla, access)
    Sub consulta(ByVal tabla As DataGridView, ByVal access As String)
        ':::Instruccion Try para capturar errores
        Try

            ':::Creamos el objeto DataAdapter y le pasamos los dos parametros (Instruccion, conexión)
            Dim DA As New OleDbDataAdapter(access, conexion)

            ':::Creamos el objeto DataTable que recibe la informacion del DataAdapter
            Dim DT As New DataTable

            ':::Pasamos la informacion del DataAdapter al DataTable mediante la propiedad Fill
            DA.Fill(DT)

            ':::Ahora mostramos los datos en el DataGridView
            tabla.DataSource = DT
        Catch ex As Exception
            MsgBox("No se logro realizar la consulta por: " & ex.Message, MsgBoxStyle.Critical, "Tutorial CRUD")
        End Try
    End Sub

    ':::PROCEDIMIENTO para Agregar, Actualizar y Eliminar ademas le indicamos que debe pedir 2 parametros para ejecutarse correctamente (tabla, access)
    Sub operaciones(ByVal tabla As DataGridView, ByVal access As String)
        ':::Instruccion Try para capturar errores
        Try

            ':::Creamos nuestro objeto de tipo Command que almacenara nuestras instrucciones, necesita 2 parametros (access, conexion)
            Dim cmd As New OleDbCommand(access, conexion)

            ':::Ejecutamos la instruccion mediante la propiedad ExecuteNonQuery del command
            cmd.ExecuteNonQuery()
        Catch ex As Exception
            MsgBox("No se logro realizar la operación por: " & ex.Message, MsgBoxStyle.Critical, "Tutorial CRUD")
        End Try
    End Sub

    ':::PROCEDIMIENTO para actualizar el DataGridView al momento de accionar cualquier boton
    Sub actualizar()
        ':::Creamos la variable access que guarda la instruccion de tipo SQL

        ':::Instrucción "Select * from [tabla] where [nombrecelda1]='" & [nombreherramienta1] & "'"
        'Dim access As String = "Select * from Certificados where NombrePaciente='" & txtNPaciente.Text & "'"

        ':::Instrucción "Select * from [tabla]"
        Dim access As String = "Select * from Datos_Paciente"

        ':::Accedemos a nuestro procedimiento "consulta" y le pasamos los dos (2) parametros (dgvTabla, access)
        Me.consulta(dgvTabla, access)
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
        ':::Limpiar los TextBox
        txtTransito.Text = ""
        txtSalud.Text = ""
        txtOftal.Text = ""
        txtAPaciente.Text = ""
        txtNPaciente.Text = ""
        txtDpi.Text = ""
        txtResidencia.Text = ""
        txtAgudeza1.Text = ""
        txtAgudeza2.Text = ""
        txtAgudeza3.Text = ""
        txtEdad.Visible = False
        txtEdad.Text = ""

        ':::Limpiar los ComboBox
        cbProfesional.SelectedValue = Nothing
        cbProfesional.Text = Nothing
        cbDepartamento.SelectedValue = Nothing
        cbDepartamento.Text = Nothing
        cbMunicipio.SelectedValue = Nothing
        cbMunicipio.Text = Nothing
        cbGenero.SelectedValue = Nothing
        cbGenero.Text = Nothing

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
        rtb1.Text = ""

        ':::Limpiar los NumericUpDown
        nudCentral1.Value = Nothing
        nudCentral2.Value = Nothing
        nudPeriferico1.Value = Nothing
        nudPeriferico2.Value = Nothing

        ':::Limpiar los CheckBox
        cbA.Checked = False
        cbB.Checked = False
        cbC.Checked = False
        cbE.Checked = False
        cbM.Checked = False

        ':::Limpiar la foto
        pbFoto.Image = Nothing
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

        Me.txtAgudeza1.Text = dgvTabla.Rows(I).Cells(12).Value.ToString
        Me.txtAgudeza2.Text = dgvTabla.Rows(I).Cells(13).Value.ToString
        Me.txtAgudeza3.Text = dgvTabla.Rows(I).Cells(14).Value.ToString

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
    Private Sub btnAgregar_Click(sender As System.Object, e As System.EventArgs) Handles btnAgregar.Click
        cbProfesional.Text = "Luis Porras"
        txtTransito.Text = "123"
        txtSalud.Text = "456"
        txtOftal.Text = "789"
        txtNPaciente.Text = "Alejandro Polares"
        txtDpi.Text = "457"
        cbDepartamento.Text = "Guatemala"
        cbMunicipio.Text = "Ciudad de Guatemala"
        cbGenero.Text = "Masculino"
        txtResidencia.Text = "Vivo por ahi"
        txtAgudeza1.Text = "67"
        txtAgudeza2.Text = "34"
        txtAgudeza3.Text = "90"
        rbVision1.Checked = True
        nudCentral1.Value = "45"
        nudCentral2.Value = "32"
        rbCentral1.Checked = True
        nudPeriferico1.Value = "5"
        nudPeriferico2.Value = "57"
        rbPeriferico1.Checked = True
        rbSensibilidad1.Checked = True
        rbPrueba1.Checked = True
        rbSeg1.Checked = True
        rbAnteojos1.Checked = True
        rbLentes1.Checked = True
        rtb1.Text = "Esto es una prueba"
    End Sub

    ':::Instrucción para abrir el explorador de archivos y buscar la FOTO
    Private Sub PictureBox1_Click(sender As System.Object, e As System.EventArgs) Handles pbFoto.Click
        ':::Buscamos la imagen a grabar
        Dim openDlg As OpenFileDialog = New OpenFileDialog()
        openDlg.Filter = "Todos los archivos JPEG|*.jpg"
        Dim filter As String = openDlg.Filter
        openDlg.Title = "Abrir archivos JPEG"
        If (openDlg.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            curFileName = openDlg.FileName
            SaveImage = True

            ':::Mostrando la foto en el picture
            Me.pbFoto.ImageLocation = curFileName.ToString
        Else
            Exit Sub
        End If
    End Sub

    ':::Instrucción para calcular la edad del paciente
    Private Sub txtDate1_CloseUp(sender As Object, e As System.EventArgs) Handles txtDate1.CloseUp
        ':::Variable para obtener el valor del día, mes y año actuales
        Dim vHoyDia As String = Date.Now.Date.Day
        Dim vHoyMes As String = Date.Now.Date.Month
        Dim vHoyAnio As String = Date.Now.Date.Year

        ':::Variable para obtener el valor del día, mes y año de nacimiento
        Dim vDia As String = txtDate1.Value.Day
        Dim vMes As String = txtDate1.Value.Month
        Dim vAnio As String = txtDate1.Value.Year

        ':::La variable edad se encarga de restar el año actual con el año de nacimiento para obtener la "edad"
        Dim edad = vHoyAnio - vAnio

        ':::La instrucción If se encarga de verificar que el día y mes actuales sean iguales a los de nacimiento para confirmar si ya cumplio años
        If (vHoyDia = vDia And vHoyMes = vMes) Then
            txtEdad.Text = "Tiene " & edad & " años."

            ':::La instrucción Else se encarga de restarle uno a edad, ya que el día y mes actuales no son igual a los de nacimiento
        Else
            Dim nuevaEdad = edad - 1
            txtEdad.Text = "Tiene " & nuevaEdad & " años."
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
    Private Sub txtAgudeza1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtAgudeza1.KeyPress
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

#End Region

#Region "BASE DE DATOS  */*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/*/"
    ':::CONEXIÓN a la base de datos Access
    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ':::Instrucción para mostrar la fecha de hoy
        Me.lFecha.Text = Date.Now.Date

        ':::Instrucción Try para capturar errores
        Try

            ':::Usamos la variable conexion para el enlace a la base de datos
            conexion.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\sistemas.INTEVISA\Desktop\Proyectos\CERTIFICADO_DE_LICENCIA\CERTIFICADO_DE_LICENCIA\Recursos\CERTIFICADO_DE_LICENCIA_BD.accdb"
            conexion.Open()
            MsgBox("SE CONECTO EXITOSAMENTE A LA BASE DE DATOS", vbInformation, "CORRECTO")
        Catch ex As Exception
            MsgBox("ERROR AL CONECTAR A LA BASE DE DATOS NO PODRA GUARDAR NIGUN DATO", vbCritical, "ERROR")
            MsgBox(ex.Message)
        End Try
    End Sub

    ':::Boton que ENVIA la información a la base de datos
    Private Sub BtnGuardar_Click(sender As System.Object, e As System.EventArgs) Handles BtnGuardar.Click

        ':::Instrucción Try para capturar errores
        Try

            ':::Usamos el comando y le indicamos con la instrucción "INSERT INTO [tabla] ([nombrecelda1],[nombrecelda2]) VALUES([@nombreherramienta1],[@nombreherramienta2]", conexion)
            comandos = New OleDbCommand("INSERT INTO Datos_Paciente (Pro_Nombre, Pro_regTransito, Pro_regSalud, Pro_regOft, Pac_Apellido, Pac_Nombre, Pac_Dpi, Pac_Departamento, Pac_Municipio, Pac_Nacimiento, Pac_Genero, Pac_Residencia, Res_Agudeza1, Res_Agudeza2, Res_Agudeza3, Res_Vision, Res_CampoCentralOD, Res_CampoCentralOI, Res_CampoCentral, Res_CampoPerifericoOD, Res_CampoPerifericoOI, Res_CampoPeriferico, Res_Sensibilidad, Res_Prueba, Res_Seg, Res_Anteojos, Res_Lentes, Res_Licencia1, Res_Licencia2, Res_Licencia3, Res_Licencia4, Res_Licencia5, Res_Licencia6, Res_Obs) VALUES (@cbProfesional, @txtTransito, @txtSalud, @txtOftal, @txtAPaciente, @txtNPaciente, @txtDpi, @cbDepartamento, @cbMunicipio, @txtDate1, @cbGenero, @txtResidencia, @txtAgudeza1, @txtAgudeza2, @txtAgudeza3, @rbVision1, @nudCentral1, @nudCentral2, @rbCentral1, @nudPeriferico1, @nudPeriferico2, @rbPeriferico1, @rbSensibilidad1, @rbPrueba1, @rbSeg1, @rbAnteojos1, @rbLentes1, @cbA, @cbB, @cbE, @cbC, @cbM, @cbNinguno, @rtb1)", conexion)

            ':::Instrucción comandos.Parameters.AddWithValue ([@nombrecelda1],[nombreherramienta1]) para agregar los datos a la tabla de la base de datos
            comandos.Parameters.AddWithValue("@Pro_Nombre", cbProfesional.SelectedItem)
            comandos.Parameters.AddWithValue("@Pro_regTransito", txtTransito.Text)
            comandos.Parameters.AddWithValue("@Pro_regSalud", txtSalud.Text)
            comandos.Parameters.AddWithValue("@Pro_regOft", txtOftal.Text)

            comandos.Parameters.AddWithValue("@Pac_Apellido", txtAPaciente.Text)
            comandos.Parameters.AddWithValue("@Pac_Nombre", txtNPaciente.Text)
            comandos.Parameters.AddWithValue("@Pac_Dpi", txtDpi.Text)
            comandos.Parameters.AddWithValue("@Pac_Departamento", cbDepartamento.SelectedItem)
            comandos.Parameters.AddWithValue("@Pac_Municipio", cbMunicipio.SelectedItem)
            comandos.Parameters.AddWithValue("@Pac_Nacimiento", txtDate1.Text)
            comandos.Parameters.AddWithValue("@Pac_Genero", cbGenero.SelectedItem)
            comandos.Parameters.AddWithValue("@Pac_Residencia", txtResidencia.Text)

            comandos.Parameters.AddWithValue("@Res_Agudeza1", txtAgudeza1.Text)
            comandos.Parameters.AddWithValue("@Res_Agudeza2", txtAgudeza2.Text)
            comandos.Parameters.AddWithValue("@Res_Agudeza3", txtAgudeza3.Text)

            ':::Instrucción If para los RadioButton. Hace la condición de guardar el texto del RadioButton que esté marcado.
            If (rbVision1.Checked = True) Then
                comandos.Parameters.AddWithValue("@Res_Vision", rbVision1.Text)
            Else
                comandos.Parameters.AddWithValue("@Res_Vision", rbVision2.Text)
            End If

            comandos.Parameters.AddWithValue("@Res_CampoCentralOD", nudCentral1.Value)
            comandos.Parameters.AddWithValue("@Res_CampoCentralOI", nudCentral2.Value)
            If (rbCentral1.Checked = True) Then
                comandos.Parameters.AddWithValue("@Res_CampoCentral", rbCentral1.Text)
            Else
                comandos.Parameters.AddWithValue("@Res_CampoCentral", rbCentral2.Text)
            End If

            comandos.Parameters.AddWithValue("@Res_CampoPerifericoOD", nudPeriferico1.Value)
            comandos.Parameters.AddWithValue("@Res_CampoPerifericoOI", nudPeriferico2.Value)
            If (rbPeriferico1.Checked = True) Then
                comandos.Parameters.AddWithValue("@Res_CampoPeriferico", rbPeriferico1.Text)
            Else
                comandos.Parameters.AddWithValue("@Res_CampoPeriferico", rbPeriferico2.Text)
            End If

            If (rbSensibilidad1.Checked = True) Then
                comandos.Parameters.AddWithValue("@Res_Sensibilidad", rbSensibilidad1.Text)
            Else
                comandos.Parameters.AddWithValue("@Res_Sensibilidad", rbSensibilidad2.Text)
            End If

            If (rbPrueba1.Checked = True) Then
                comandos.Parameters.AddWithValue("@Res_Prueba", rbPrueba1.Text)
            Else
                comandos.Parameters.AddWithValue("@Res_Prueba", rbPrueba2.Text)
            End If

            If (rbSeg1.Checked = True) Then
                comandos.Parameters.AddWithValue("@Res_Seg", rbSeg1.Text)
            Else
                comandos.Parameters.AddWithValue("@Res_Seg", rbSeg2.Text)
            End If

            If (rbAnteojos1.Checked = True) Then
                comandos.Parameters.AddWithValue("@Res_Anteojos", rbAnteojos1.Text)
            Else
                comandos.Parameters.AddWithValue("@Res_Anteojos", rbAnteojos2.Text)
            End If

            If (rbLentes1.Checked = True) Then
                comandos.Parameters.AddWithValue("@Res_Lentes", rbLentes1.Text)
            Else
                comandos.Parameters.AddWithValue("@Res_Lentes", rbLentes2.Text)
            End If

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

            comandos.Parameters.AddWithValue("@Res_Obs", rtb1.Text)

            ':::Ejecutamos la instruccion mediante la propiedad ExecuteNonQuery del command
            comandos.ExecuteNonQuery()
            MsgBox("DATOS GUARDADOS EXITOSAMENTE", vbInformation, "CORRECTO")
            actualizar()
        Catch ex As Exception
            MsgBox("ERROR AL GUARDAR EL FORMULARIO", vbCritical, "ERROR")
            MsgBox(ex.Message)
        End Try
    End Sub

    ':::Boton que MUESTRA la información almacenada en la base de datos
    Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click
        ':::Creamos la variable access que guarda la instruccion de tipo SQL

        ':::Instrucción "Select * from [tabla] where [nombrecelda1]='" & [nombreherramienta1] & "'"

        'Dim access As String = "Select * from Certificados where NombrePaciente='" & txtNPaciente.Text & "'"
        Dim access As String = "Select * from Datos_Paciente"

        ':::Accedemos a nuestro procedimiento "consulta" y le pasamos los dos (2) parametros (dgvTabla, access)
        Me.consulta(dgvTabla, access)
    End Sub

    ':::Boton que ACTUALIZA la informacion almacenada en la base de datos
    Private Sub Button5_Click(sender As System.Object, e As System.EventArgs) Handles Button5.Click
        ':::Creamos la variable access que guardar la instruccion de tipo SQL

        ':::Instrucción "Update [tabla] set [nombrecelda1]='" & [nombreherramienta1] & "'" where [nombrecelda2]= " & [nombreherramienta1] & "" 
        Dim access As String = "Update Datos_Paciente Set NombrePaciente='" & txtNPaciente.Text & "', Genero='" & cbGenero.Text & "' where Id=" & txtOftal.Text & ""
        Me.operaciones(dgvTabla, access)
    End Sub

    ':::Boton que ELIMINA la información de la base de datos
    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        ':::Creamos la variable access que guardar la instruccion de tipo SQL

        ':::Instrucción "Delete * From [tabla] Where [nombrecelda1]=" & [nombreherramienta1] & ""
        Dim access As String = "Delete * From Datos_Paciente Where Pac_Nombre='" & txtNPaciente.Text & "'"
        Me.operaciones(dgvTabla, access)
        actualizar()
    End Sub

#End Region

#Region "FOTOGRAFÍA"
    Dim DATOS As IDataObject
    Dim IMAGEN As Image
    Dim CARPETA As String
    Dim FECHA As String = DateTime.Now.ToShortDateString().Replace("/", "_") + "_" + DateTime.Now.ToLongTimeString().Replace(":", "_")
    Dim DIRECTORIO As String = "C:\Users\sistemas.INTEVISA\Desktop\" ' AQUI COLOCA LA RUTA A TU ESCRITORIO
    Dim DESTINO As String
    Dim CONTADOR As Integer = 1
    Dim CARPETAS_DIARIAS As String
    Public Const WM_CAP As Short = &H400S
    Public Const WM_CAP_DLG_VIDEOFORMAT As Integer = WM_CAP + 41
    Public Const WM_CAP_DRIVER_CONNECT As Integer = WM_CAP + 10
    Public Const WM_CAP_DRIVER_DISCONNECT As Integer = WM_CAP + 11
    Public Const WM_CAP_EDIT_COPY As Integer = WM_CAP + 30
    Public Const WM_CAP_SEQUENCE As Integer = WM_CAP + 62
    Public Const WM_CAP_FILE_SAVEAS As Integer = WM_CAP + 23
    Public Const WM_CAP_SET_PREVIEW As Integer = WM_CAP + 50
    Public Const WM_CAP_SET_PREVIEWRATE As Integer = WM_CAP + 52
    Public Const WM_CAP_SET_SCALE As Integer = WM_CAP + 53
    Public Const WS_CHILD As Integer = &H40000000
    Public Const WS_VISIBLE As Integer = &H10000000
    Public Const SWP_NOMOVE As Short = &H2S
    Public Const SWP_NOSIZE As Short = 1
    Public Const SWP_NOZORDER As Short = &H4S
    Public Const HWND_BOTTOM As Short = 1
    Public Const WM_CAP_STOP As Integer = WM_CAP + 68

    Public iDevice As Integer = 0 ' Current device ID
    Public hHwnd As Integer ' Handle to preview window

    Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
        (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer,
        <MarshalAs(UnmanagedType.AsAny)> ByVal lParam As Object) As Integer

    Public Declare Function SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hwnd As Integer,
        ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer,
        ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer) As Integer

    Public Declare Function DestroyWindow Lib "user32" (ByVal hndw As Integer) As Boolean

    Public Declare Function capCreateCaptureWindowA Lib "avicap32.dll" _
        (ByVal lpszWindowName As String, ByVal dwStyle As Integer,
        ByVal x As Integer, ByVal y As Integer, ByVal nWidth As Integer,
        ByVal nHeight As Short, ByVal hWndParent As Integer,
        ByVal nID As Integer) As Integer

    Public Declare Function capGetDriverDescriptionA Lib "avicap32.dll" (ByVal wDriver As Short,
        ByVal lpszName As String, ByVal cbName As Integer, ByVal lpszVer As String,
        ByVal cbVer As Integer) As Boolean

    ':::Abrir la vista
    Public Sub OpenPreviewWindow()

        ' Open Preview window in picturebox
        '
        hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 640,
        480, VISOR.Handle.ToInt32, 0)

        ' Connect to device
        '
        SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0)
        If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then
            '
            'Set the preview scale

            SendMessage(hHwnd, WM_CAP_SET_SCALE, True, 0)

            'Set the preview rate in milliseconds
            '
            SendMessage(hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0)

            'Start previewing the image from the camera
            '
            SendMessage(hHwnd, WM_CAP_SET_PREVIEW, True, 0)

            ' Resize window to fit in picturebox
            '
            SetWindowPos(hHwnd, HWND_BOTTOM, 0, 0, VISOR.Width, VISOR.Height,
            SWP_NOMOVE Or SWP_NOZORDER)

        Else
            ' Error connecting to device close window
            ' 
            DestroyWindow(hHwnd)

        End If
    End Sub

    ':::Abrir la prevista 
    Public Sub OpenPreviewWindowCliente()

        ' Open Preview window in picturebox
        '
        hHwnd = capCreateCaptureWindowA(iDevice, WS_VISIBLE Or WS_CHILD, 0, 0, 600,
           480, Me.pbFoto.Handle.ToInt32, 0)

        ' Connect to device
        '
        SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0)
        If SendMessage(hHwnd, WM_CAP_DRIVER_CONNECT, iDevice, 0) Then
            '
            'Set the preview scale

            SendMessage(hHwnd, WM_CAP_SET_SCALE, True, 0)

            'Set the preview rate in milliseconds
            '
            SendMessage(hHwnd, WM_CAP_SET_PREVIEWRATE, 66, 0)

            'Start previewing the image from the camera
            '
            SendMessage(hHwnd, WM_CAP_SET_PREVIEW, True, 0)

            ' Resize window to fit in picturebox
            '
            SetWindowPos(hHwnd, HWND_BOTTOM, 0, 0, Me.pbFoto.Width, Me.pbFoto.Height,
                    SWP_NOMOVE Or SWP_NOZORDER)

        Else
            ' Error connecting to device close window
            ' 
            DestroyWindow(hHwnd)

        End If
    End Sub

    ':::Capturar la fotografia
    Public Sub CapturarCliente()
        ' Copy image to clipboard
        '
        SendMessage(hHwnd, WM_CAP_EDIT_COPY, 0, 0)

        ' Get image from clipboard and convert it to a bitmap
        '
        DATOS = Clipboard.GetDataObject()

        IMAGEN = CType(DATOS.GetData(GetType(System.Drawing.Bitmap)), Image)
        Me.pbFoto.Image = IMAGEN
        'GUARDAR.Visible = True
    End Sub

    ':::Cerrar ventana anterior
    Public Sub ClosePreviewWindow()
        '
        ' Disconnect from device
        '
        SendMessage(hHwnd, WM_CAP_DRIVER_DISCONNECT, 0, 0)
        '
        ' close window
        '
        DestroyWindow(hHwnd)
    End Sub

    ':::Abrir camara
    Private Sub cmdCamara_Click(sender As Object, e As EventArgs) Handles cmdCamara.Click
        Me.OpenPreviewWindowCliente()
        cmdCamara.Enabled = False
    End Sub

#End Region
End Class