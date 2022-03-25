using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Drawing.Printing;
using System.Net.Mail;
using System.Windows.Forms;
using System.Resources;


namespace InvFun
{
    class Funciones
    {

		#region Definicion de Variable
		// Objeto conexion sql server
		private SqlConnection LO_SQLServerConexion = new SqlConnection();
		// Variable que me permite acceder a las funciones
		public static InvFun.Funciones LF_Funciones = new InvFun.Funciones();
		// Objeto de tipo DataTable, aqui se guardan todos los datos obtenidos desde Sql Server
		private DataSet LO_ResultadoSQLSERVER = new DataSet();
		private DateTime LD_FechaSistema = DateTime.Today;
		private string LS_NombreUsuario = "";
		#endregion
		#region Definicion de propiedades
		public string NombreUsuario
		{
			get { return LF_Funciones.LS_NombreUsuario; }
			set { LF_Funciones.LS_NombreUsuario = value; }
		}
		public SqlConnection SQLSERVERConexion
		{
			get { return LF_Funciones.LO_SQLServerConexion; }
			set { LF_Funciones.LO_SQLServerConexion = value; }
		}
		public DataSet ResultadoSQLSERVER
		{
			get { return LF_Funciones.LO_ResultadoSQLSERVER; }
			set { LF_Funciones.LO_ResultadoSQLSERVER = value; }
		}
		public DateTime FechaSistema
		{
			get { return LF_Funciones.LD_FechaSistema; }
			set { LF_Funciones.LD_FechaSistema = value; }
		}
		#endregion
		#region LimpiarResultadoSQL: Limpia todo el DataTable
		public void LimpiarResultadoSQLSERVER()
		{
			if (LF_Funciones.ResultadoSQLSERVER.Tables.Count > 0)
			{
				LF_Funciones.ResultadoSQLSERVER.Tables.Clear();
			}
		}
		#endregion
		#region EjecutarComandoSQLSERVER: funcion que ejecuta cualquier comando SQL SERVER como si estuvieramos en el Query Analyze, devuelve falso si no hubo un error y verdadero si hubo error
		public bool EjecutarComandoSQLSERVER(string LS_ComandoSQLSERVER, bool LB_LimpiarResultadoSQLSERVER, bool LB_MostrarError)
		{
			// Bandera de Error
			bool LB_Error = false;
			// Limpiamos el DataSet ResultadoSQLServer
			if (LB_LimpiarResultadoSQLSERVER == true)
				LF_Funciones.LimpiarResultadoSQLSERVER();
			// Se Crea la Instancia del Comando
			SqlCommand LO_SqlCommand = new SqlCommand();
			// Le Decimos que es de tipo texto es decir que se le enviara una cadea
			LO_SqlCommand.CommandType = CommandType.Text;
			// Se le envia la cadena con el comando
			LO_SqlCommand.CommandText = LS_ComandoSQLSERVER;
			// Se le decie la conexion cual es
			LO_SqlCommand.Connection = LF_Funciones.SQLSERVERConexion;
			// Y se le dice que no hay tiempo de espera
			LO_SqlCommand.CommandTimeout = 0;
			try
			{
				// Se Crea la Instancia SqlDataAdapter para Crear un Nuevo Comando
				SqlDataAdapter LO_SqlDataAdapter = new SqlDataAdapter(LO_SqlCommand);
				// Se LLena el DataSet
				LO_SqlDataAdapter.Fill(LF_Funciones.ResultadoSQLSERVER);
				// Se Limpia el Comando
				LO_SqlDataAdapter = null/* TODO Change to default(_) if this is not a reference type */;
				LB_Error = false;
			}
			catch (Exception ex)
			{
				LB_Error = true;
				if (LB_MostrarError == true)
				{
					MessageBox.Show("Comando a ejecutar: " + (Char)13 + LS_ComandoSQLSERVER.Trim() + (Char)13, "Error al Ejecutar Comando SQL", MessageBoxButtons.OK, MessageBoxIcon.Error);
					MessageBox.Show("Excepcion:" + (Char)13 + ex.ToString().Trim(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					Form LO_Formulario = new Form();

					// *********************************************
					// Desmarcar para ver el Query en SQL Server
					// *********************************************
					TextBox LO_Textbox = new TextBox();
					LO_Formulario.Controls.Add(LO_Textbox);
					LO_Textbox.Multiline = true;
					LO_Textbox.Top = 0;
					LO_Textbox.Left = 0;
					LO_Textbox.Width = LO_Formulario.Width;
					LO_Textbox.Height = LO_Formulario.Height;
					LO_Textbox.Text = LS_ComandoSQLSERVER.Trim();
					LO_Textbox.Anchor = (((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
					| System.Windows.Forms.AnchorStyles.Left)
					| System.Windows.Forms.AnchorStyles.Right);

					LO_Formulario.Text = "Error Comando SQL Server";
					LO_Formulario.StartPosition = FormStartPosition.CenterScreen;
					LO_Formulario.WindowState = FormWindowState.Maximized;
					LO_Formulario.MaximizeBox = false;
					LO_Formulario.MinimizeBox = false;
					LO_Formulario.Show();
					LO_Formulario.WindowState = FormWindowState.Minimized;
				}
			}
			return LB_Error;
		}

		#endregion

		#region ConsultasSQL: funcion que devuelve el ComandoSQL para las diferentes consultas a la base de datos        
		public string ConsultaSQLVentasProductoPrecio(string LS_FechaDesdeReporte, string LS_FechaHastaReporte)
		{
			string sqlVentasVentasProductoPrecio =
				@"

				SET LANGUAGE US_ENGLISH
				DECLARE @Reporte TABLE(Fecha DATETIME, Factura VARCHAR(20), Vendedor VARCHAR(5), NombreVendedor VARCHAR(30), Producto VARCHAR(16), DescripcionProducto VARCHAR(40), Cliente VARCHAR(10),NombreCliente VARCHAR(120), Etico NUMERIC(18,4), GenP1_Dimedics NUMERIC(18,4), GenP1 NUMERIC(18,4), GenP4 NUMERIC(18,4), MmP1 NUMERIC(18,4), MmP4 NUMERIC(18,4), Impuesto NUMERIC(18,4), Descuento NUMERIC(18,4), DescProPago NUMERIC(18,4), TotalFactura NUMERIC(18,4))
				DECLARE @LD_FechaDesde DATETIME, @LD_FechaHasta DATETIME

				SET @LD_FechaDesde = '" + LS_FechaDesdeReporte + @"'
				SET @LD_FechaHasta = '" + LS_FechaHastaReporte + @"'


				INSERT INTO @Reporte

				SELECT A.FECHA_DOC, A.CVE_DOC, V.CVE_VEND, V.NOMBRE, I.CVE_ART, I.DESCR, C.CLAVE, IIF(A.CVE_CLPV='MOSTR',CC.NOMBRE, C.NOMBRE), 
				IIF(SUBSTRING(I.LIN_PROD,1,1)='E' AND I.LIN_PROD NOT LIKE'%11',B.CANT*B.PREC,0), IIF(I.LIN_PROD LIKE'G%01' AND B.PREC>=P.PRECIO AND I.LIN_PROD NOT LIKE'%11',B.CANT*B.PREC,0),IIF(SUBSTRING(I.LIN_PROD,1,1)='G' AND I.LIN_PROD NOT LIKE'G%01' AND B.PREC>=P.PRECIO AND I.LIN_PROD NOT LIKE'%11',B.CANT*B.PREC,0), IIF(SUBSTRING(I.LIN_PROD,1,1)='G' AND B.PREC<P.PRECIO AND I.LIN_PROD NOT LIKE'%11',B.CANT*B.PREC,0), 
				IIF(I.LIN_PROD LIKE'%11' AND B.PREC>=P.PRECIO,B.CANT*B.PREC,0), IIF(I.LIN_PROD LIKE'%11' AND B.PREC<P.PRECIO,B.CANT*B.PREC,0),
				B.TOTIMP4, (B.CANT*B.PREC*B.DESC1/100), 0, A.IMPORTE
				FROM FACTF01 A
				INNER JOIN PAR_FACTF01 B ON A.CVE_DOC=B.CVE_DOC AND B.TIPO_ELEM='N'
				INNER JOIN INVE01 I ON B.CVE_ART=I.CVE_ART
				INNER JOIN CLIE01 C ON A.CVE_CLPV=C.CLAVE
				LEFT JOIN INFCLI01 CC ON A.DAT_MOSTR=CC.CVE_INFO
				INNER JOIN PRECIO_X_PROD01 P ON B.CVE_ART=P.CVE_ART AND P.CVE_PRECIO=1
				INNER JOIN VEND01 V ON A.CVE_VEND=V.CVE_VEND
				WHERE A.STATUS<>'C' AND A.FECHA_DOC>=@LD_FechaDesde AND A.FECHA_DOC<=@LD_FechaHasta

				INSERT INTO @Reporte

				SELECT A.FECHA_DOC, A.CVE_DOC, V.CVE_VEND, V.NOMBRE, I.CVE_ART, I.DESCR, C.CLAVE, IIF(A.CVE_CLPV='MOSTR',CC.NOMBRE, C.NOMBRE), 
				IIF(SUBSTRING(I.LIN_PROD,1,1)='E' AND I.LIN_PROD NOT LIKE'%11',(B.CANT*B.PREC)*-1,0), IIF(I.LIN_PROD LIKE'G%01' AND B.PREC>=P.PRECIO AND I.LIN_PROD NOT LIKE'%11',(B.CANT*B.PREC)*-1,0),IIF(SUBSTRING(I.LIN_PROD,1,1)='G' AND I.LIN_PROD NOT LIKE'G%01' AND B.PREC>=P.PRECIO AND I.LIN_PROD NOT LIKE'%11',(B.CANT*B.PREC)*-1,0), IIF(SUBSTRING(I.LIN_PROD,1,1)='G' AND B.PREC<P.PRECIO AND I.LIN_PROD NOT LIKE'%11',(B.CANT*B.PREC)*-1,0), 
				IIF(I.LIN_PROD LIKE'%11' AND B.PREC>=P.PRECIO,(B.CANT*B.PREC)*-1,0), IIF(I.LIN_PROD LIKE'%11' AND B.PREC<P.PRECIO,(B.CANT*B.PREC)*-1,0),
				B.TOTIMP4*-1, (B.CANT*B.PREC*B.DESC1/100)*-1, IIF(I.LIN_PROD ='98',(B.CANT*B.PREC),0), A.IMPORTE*-1
				FROM FACTD01 A
				INNER JOIN PAR_FACTD01 B ON A.CVE_DOC=B.CVE_DOC AND B.TIPO_ELEM='N'
				INNER JOIN INVE01 I ON B.CVE_ART=I.CVE_ART
				INNER JOIN CLIE01 C ON A.CVE_CLPV=C.CLAVE
				LEFT JOIN INFCLI01 CC ON A.DAT_MOSTR=CC.CVE_INFO
				INNER JOIN PRECIO_X_PROD01 P ON B.CVE_ART=P.CVE_ART AND P.CVE_PRECIO=1
				INNER JOIN VEND01 V ON A.CVE_VEND=V.CVE_VEND
				WHERE A.STATUS<>'C' AND A.FECHA_DOC>=@LD_FechaDesde AND A.FECHA_DOC<=@LD_FechaHasta


				SELECT Fecha, Factura AS 'Documento', Vendedor, NombreVendedor, Cliente, NombreCliente, Etico, GenP1_Dimedics as'Generico Precio1 Dimedics', GenP1 as'Generico Precio1',GenP4 as'Generico Precio4',MmP1 as'Material Medico Precio1',MmP4 as'Material Medico Precio4',Impuesto,Descuento,DescProPago AS'Descuento Pronto Pago',TotalFactura AS'Total Documento',Producto,DescripcionProducto FROM @Reporte
				ORDER BY Fecha, Factura, Producto


	            ";

			return sqlVentasVentasProductoPrecio;
		}

		public string ConsultaSQLVentasProductoBonificacion(string LS_FechaDesdeReporte, string LS_FechaHastaReporte)
		{
			string sqlVentasVentasProductoBonificacion =
				@"

				SET LANGUAGE US_ENGLISH
				DECLARE @Reporte TABLE(Fecha DATETIME, Factura VARCHAR(20), Producto VARCHAR(16), DescripcionProducto VARCHAR(40), Cliente VARCHAR(10),NombreCliente VARCHAR(120), CantidadVenta NUMERIC(18,4), Bonificacion NUMERIC(18,4), ValorVenta NUMERIC(18,4), Descuento NUMERIC(18,4), Impuesto NUMERIC(18,4), Costo NUMERIC(18,4), TipoPrecio VARCHAR(50))
				DECLARE @LD_FechaDesde DATETIME, @LD_FechaHasta DATETIME

				SET @LD_FechaDesde = '" + LS_FechaDesdeReporte + @"'
				SET @LD_FechaHasta = '" + LS_FechaHastaReporte + @"'


				INSERT INTO @Reporte

				SELECT A.FECHA_DOC, A.CVE_DOC, I.CVE_ART, I.DESCR, C.CLAVE, IIF(A.CVE_CLPV='MOSTR',CC.NOMBRE, C.NOMBRE), IIF(B.PREC>0,B.CANT,0), IIF(B.PREC=0,B.CANT,0), (B.CANT*B.PREC-(B.CANT*B.PREC*B.DESC1/100)), (B.CANT*B.PREC*B.DESC1/100), B.TOTIMP4, (B.CANT*B.COST), IIF((B.CANT*B.PREC-(B.CANT*B.PREC*B.DESC1/100))<P.PRECIO,'Menor que publico','Publico') 
				FROM FACTF01 A
				INNER JOIN PAR_FACTF01 B ON A.CVE_DOC=B.CVE_DOC AND B.TIPO_PROD='P'
				INNER JOIN INVE01 I ON B.CVE_ART=I.CVE_ART
				INNER JOIN CLIE01 C ON A.CVE_CLPV=C.CLAVE
				LEFT JOIN INFCLI01 CC ON A.DAT_MOSTR=CC.CVE_INFO
				INNER JOIN PRECIO_X_PROD01 P ON B.CVE_ART=P.CVE_ART AND P.CVE_PRECIO=1
				WHERE A.FECHA_DOC>=@LD_FechaDesde AND A.FECHA_DOC<=@LD_FechaHasta


				INSERT INTO @Reporte

				SELECT A.FECHA_DOC, A.CVE_DOC, I.CVE_ART, I.DESCR, C.CLAVE, IIF(A.CVE_CLPV='MOSTR',CC.NOMBRE, C.NOMBRE), IIF(B.PREC>0,B.CANT*-1,0), IIF(B.PREC=0,B.CANT*-1,0), (B.CANT*B.PREC-(B.CANT*B.PREC*B.DESC1/100))*-1, (B.CANT*B.PREC*B.DESC1/100)*-1, (B.TOTIMP4*-1), (B.CANT*B.COST)*-1, IIF((B.CANT*B.PREC-(B.CANT*B.PREC*B.DESC1/100))<P.PRECIO,'Menor que publico','Publico')
				FROM FACTD01 A
				INNER JOIN PAR_FACTD01 B ON A.CVE_DOC=B.CVE_DOC AND B.TIPO_PROD='P'
				INNER JOIN INVE01 I ON B.CVE_ART=I.CVE_ART
				INNER JOIN CLIE01 C ON A.CVE_CLPV=C.CLAVE
				LEFT JOIN INFCLI01 CC ON A.DAT_MOSTR=CC.CVE_INFO
				INNER JOIN PRECIO_X_PROD01 P ON B.CVE_ART=P.CVE_ART AND P.CVE_PRECIO=1
				WHERE A.FECHA_DOC>=@LD_FechaDesde AND A.FECHA_DOC<=@LD_FechaHasta


				SELECT Fecha, Factura, Producto, DescripcionProducto, Cliente, NombreCliente, SUM(CantidadVenta) AS CantidadVenta, SUM(Bonificacion) AS Bonificacion, SUM(ValorVenta) AS ValorVenta, SUM(Descuento) AS Descuento, SUM(Impuesto) AS Impuesto, SUM(Costo) AS Costo, TipoPrecio  FROM @Reporte
				GROUP BY Fecha, Factura, Producto, DescripcionProducto, Cliente, NombreCliente,TipoPrecio


	            ";

			return sqlVentasVentasProductoBonificacion;
		}

		public string ConsultaSQLAntiguedadClienteZona(string LS_FechaHastaReporte)
		{
			string sqlAntiguedadClienteZona =
				@"

                SET LANGUAGE US_ENGLISH
                DECLARE @LV_NombreUsuario VARCHAR(MAX), @LD_FechaAl DATETIME,@LV_CuentaEnviarCorreo NVARCHAR(MAX), @LV_AsuntoCorreo NVARCHAR(MAX), @LV_CuerpoCorreo NVARCHAR(MAX)
                --Inicializamos Variable
                --SET @LV_NombreUsuario='ALEJANDRO ZUÑIGA'
                --SET @LV_CuentaEnviarCorreo='azunigapaz@gmail.com'
                SET @LD_FechaAl='" + LS_FechaHastaReporte + @"'
                --Creamos el Cursor
                --Movimientos
                DECLARE @ReporteMovimiento TABLE (Documento VARCHAR(20), Monto NUMERIC(18,4))
                --Reporte Final
                DECLARE @Reporte TABLE (Vendedor VARCHAR(5), NombreVendedor VARCHAR(30), Cliente VARCHAR(10), NombreCliente VARCHAR(120), Documento VARCHAR(20), Fecha DATETIME, FechaVencimiento DATETIME, Saldo NUMERIC(18,4), DiaAntiguedad SMALLINT, Zona VARCHAR(10), NombreZona VARCHAR(100))
                --Insertamos los Movimientos  CUENT_M
                INSERT INTO @ReporteMovimiento (Documento,Monto)
                SELECT REFER,IMPORTE*SIGNO
                FROM CUEN_M01 WHERE FECHA_APLI<=@LD_FechaAl
                --Insertamos la Informacion  CUENT_DET
                INSERT INTO @ReporteMovimiento (Documento,Monto)
                SELECT REFER,IMPORTE*SIGNO
                FROM CUEN_DET01 WHERE FECHA_APLI<=@LD_FechaAl
                --Insertamos en el Reporte la Informacion
                INSERT INTO @Reporte (Documento,Saldo)
                SELECT Documento,SUM(Monto)
                FROM @ReporteMovimiento
                GROUP BY Documento
                HAVING SUM(Monto)<>0 
                --Actualizamos Fecha y la Fecha de  Vencimiento
                UPDATE A SET A.Fecha=(SELECT TOP 1 AA.FECHA_APLI FROM  CUEN_M01 AS AA WHERE  AA.REFER=A.Documento),A.FechaVencimiento=(SELECT TOP 1 AA.FECHA_VENC FROM  CUEN_M01 AS AA WHERE  AA.REFER=A.Documento)
                FROM @Reporte AS A
                WHERE A.Fecha IS NULL
                --Actualizamos Fecha y la Fecha de  Vencimiento
                UPDATE A SET A.Fecha=(SELECT TOP 1 AA.FECHA_APLI FROM  CUEN_DET01 AS AA WHERE  AA.REFER=A.Documento),A.FechaVencimiento=(SELECT TOP 1 AA.FECHA_VENC FROM  CUEN_DET01 AS AA WHERE  AA.REFER=A.Documento)
                FROM @Reporte AS A
                WHERE A.Fecha IS NULL
                --Actualizamos Cliente
                UPDATE A SET A.Cliente=(SELECT TOP 1 AA.CVE_CLIE FROM  CUEN_M01 AS AA WHERE  AA.REFER=A.Documento AND LEN(LTRIM(AA.CVE_CLIE))>0)
                FROM @Reporte AS A
                WHERE A.Cliente IS NULL
                --Actualizamos Cliente
                UPDATE A SET A.Cliente=(SELECT TOP 1 AA.CVE_CLIE FROM  CUEN_DET01 AS AA WHERE  AA.REFER=A.Documento AND LEN(LTRIM(AA.CVE_CLIE))>0)
                FROM @Reporte AS A
                WHERE A.Cliente IS NULL
                --Actualizamos Vendedor
                UPDATE A SET A.Vendedor=(SELECT TOP 1 AA.STRCVEVEND FROM  CUEN_M01 AS AA WHERE  AA.REFER=A.Documento AND LEN(LTRIM(AA.STRCVEVEND))>0)
                FROM @Reporte AS A
                WHERE A.Vendedor IS NULL
                --Actualizamos Vendedor
                UPDATE A SET A.Vendedor=(SELECT TOP 1 AA.STRCVEVEND FROM  CUEN_DET01 AS AA WHERE  AA.REFER=A.Documento AND LEN(LTRIM(AA.STRCVEVEND))>0)
                FROM @Reporte AS A
                WHERE A.Vendedor IS NULL
                --Actualizamos el Cliente
                UPDATE A SET A.NombreCliente=B.NOMBRE
                FROM @Reporte AS A
                INNER JOIN CLIE01 B ON B.CLAVE=A.Cliente
                --Actualizamos la Zona
                UPDATE A SET A.Zona=B.CVE_ZONA
                FROM @Reporte AS A
                INNER JOIN CLIE01 B ON B.CLAVE=A.Cliente
                --Actualizamos el nombre de la Zona
                UPDATE A SET A.NombreZona=B.TEXTO
                FROM @Reporte AS A
                INNER JOIN ZONA01 B ON B.CVE_ZONA=A.Zona
                --Actualizamos el Vendedor
                UPDATE A SET A.NombreVendedor=B.NOMBRE
                FROM @Reporte AS A
                INNER JOIN VEND01 B ON B.CVE_VEND=A.Vendedor
                --Actualizamos Cliente que no Existan
                UPDATE @Reporte SET Cliente='N/D',NombreCliente='(NO DEFINIDO)' WHERE Cliente IS NULL OR LEN(LTRIM(Cliente))=0
                --Actualizamos el  NombreCliente
                UPDATE @Reporte SET NombreCliente='(NO DEFINIDO)' WHERE NombreCliente IS NULL OR LEN(LTRIM(NombreCliente))=0
                --Actualizamos Vendedor que no Existan
                UPDATE @Reporte SET Vendedor='N/D',NombreVendedor='(NO DEFINIDO)' WHERE Vendedor IS NULL OR LEN(LTRIM(Vendedor))=0
                --Actualizamos el  NombreVendedor
                UPDATE @Reporte SET NombreVendedor='(NO DEFINIDO)' WHERE NombreVendedor IS NULL OR LEN(LTRIM(NombreVendedor))=0
                --Actualizamos los Dias de Antiguedad
                UPDATE @Reporte SET DiaAntiguedad=DATEDIFF(DD,FechaVencimiento,@LD_FechaAl)
                --Obtenemos  los Vendedores
                DECLARE @ReporteVendedor TABLE (Vendedor VARCHAR(5), NombreVendedor VARCHAR(30))
                --Insertamos los Vendedores
                INSERT INTO @ReporteVendedor (Vendedor,NombreVendedor)
                SELECT DISTINCT Vendedor,NombreVendedor FROM @Reporte
                --Obtenemos  los Clientes
                DECLARE @ReporteCliente TABLE (Cliente VARCHAR(10), NombreCliente VARCHAR(120))
                --Insertamos los Vendedores
                INSERT INTO @ReporteCliente (Cliente,NombreCliente)
                SELECT DISTINCT Cliente,NombreCliente FROM @Reporte
								
                SELECT 
                Documento,C.FECHA_APLI AS Fecha,Cliente,NombreCliente, NombreZona,Z2.TEXTO AS Departamento, NombreVendedor,
                ISNULL((SELECT SUM(Saldo) FROM @Reporte AA WHERE AA.Cliente=A.Cliente AND AA.Documento=A.Documento AND DiaAntiguedad <=0),0) AS 'AlCorriente',
                ISNULL((SELECT SUM(Saldo) FROM @Reporte AA WHERE AA.Cliente=A.Cliente AND AA.Documento=A.Documento AND DiaAntiguedad >0),0) AS 'Vencido',
                SUM(SALDO) AS Total,
                ISNULL((SELECT SUM(Saldo) FROM @Reporte AA WHERE AA.Cliente=A.Cliente AND AA.Documento=A.Documento AND DiaAntiguedad >=1 AND DiaAntiguedad <=30),0) AS 'TreintaDias',
                ISNULL((SELECT SUM(Saldo) FROM @Reporte AA WHERE AA.Cliente=A.Cliente AND AA.Documento=A.Documento AND DiaAntiguedad >=31 AND DiaAntiguedad <=60),0) AS 'SesentaDias',
                ISNULL((SELECT SUM(Saldo) FROM @Reporte AA WHERE AA.Cliente=A.Cliente AND AA.Documento=A.Documento AND DiaAntiguedad >=61 AND DiaAntiguedad <=90),0) AS 'NoventaDias',
                ISNULL((SELECT SUM(Saldo) FROM @Reporte AA WHERE AA.Cliente=A.Cliente AND AA.Documento=A.Documento AND DiaAntiguedad >=91 AND DiaAntiguedad <=120),0) AS 'CientoVeinteDias',
                ISNULL((SELECT SUM(Saldo) FROM @Reporte AA WHERE AA.Cliente=A.Cliente AND AA.Documento=A.Documento AND DiaAntiguedad >=121),0) AS 'MasCientoVeintiuno'

                FROM @Reporte A
				INNER JOIN ZONA01 Z1 ON Z1.CVE_ZONA=Zona
				INNER JOIN ZONA01 Z2 ON Z2.CVE_ZONA=Z1.CVE_PADRE
				INNER JOIN CUEN_M01 C ON C.REFER=A.Documento
--				WHERE Cliente='       174'
                GROUP BY Documento, C.FECHA_APLI,Cliente,NombreCliente,NombreZona,Z2.TEXTO,NombreVendedor
                ORDER BY SUM(SALDO) DESC
				
                    
                ";

			return sqlAntiguedadClienteZona;
		}

		public string ConsultaSQLComisiones(string LS_FechaDesdeReporte, string LS_FechaHastaReporte)
		{
			string sqlComisiones =
				@"

				SET LANGUAGE US_ENGLISH
				DECLARE @ReporteVentas TABLE(Factura VARCHAR(20),Fecha DATETIME, Producto VARCHAR(20),Cantidad NUMERIC(18,4),Categoria VARCHAR(5),NombreCategoria VARCHAR(60),SubCategoria VARCHAR(5), NombreSubCategoria VARCHAR(60), Casa VARCHAR(5), NombreCasa VARCHAR(60), Precio NUMERIC(18,4), PrecioLista1 NUMERIC(18,4) DEFAULT 0, SubTotal NUMERIC(18,4), Impuesto NUMERIC(18,4), Total NUMERIC(18,4))
				DECLARE @ReporteCobros TABLE(Factura VARCHAR(20),Recibo VARCHAR(20), FechaRecibo DATETIME, ConceptoPago INT, Cliente VARCHAR(20), NombreCliente VARCHAR(200), Vendedor VARCHAR(10), NombreVendedor VARCHAR(60),Importe NUMERIC(18,4), DescuentoProntoPago NUMERIC(18,4))
				DECLARE @ReporteBaseComision TABLE(Factura VARCHAR(20),FechaFactura DATETIME,Recibo VARCHAR(20),FechaRecibo DATETIME,Concepto INT, Cliente VARCHAR(20),NombreCliente VARCHAR(200),Vendedor VARCHAR(5),NombreVendedor VARCHAR(100),MontoRecibo NUMERIC(18,4),DescuentoProntoPago NUMERIC(18,4),Impuesto NUMERIC(18,4),TotalGenericoPrecioPublico NUMERIC(18,4),TotalGenericoMenorPrecioPublico NUMERIC(18,4),TotalEtico NUMERIC(18,4),TotalEquipoMedicoPrecioPublico NUMERIC(18,4),TotalEquipoMedicoMenorPrecioPublico NUMERIC(18,4),Otros NUMERIC(18,4),TotalFactura NUMERIC(18,4),PorcImpuesto NUMERIC(18,4),PorcTotalGenericoPrecioPublico NUMERIC(18,4),PorcTotalGenericoMenorPrecioPublico NUMERIC(18,4),PorcTotalEtico NUMERIC(18,4),PorcTotalEquipoMedicoPrecioPublico NUMERIC(18,4),PorcTotalEquipoMedicoMenorPrecioPublico NUMERIC(18,4),PorcOtros NUMERIC(18,4),MontoCobroImpuesto NUMERIC(18,4),MontoCobroGenericoPrecioPublico NUMERIC(18,4),MontoCobroGenericoMenorPrecioPublico NUMERIC(18,4),MontoCobroEtico NUMERIC(18,4),MontoCobroEquipoMedicoPrecioPublico NUMERIC(18,4),MontoCobroEquipoMedicoMenorPrecioPublico NUMERIC(18,4),MontoCobroOtros NUMERIC(18,4))

				DECLARE @LD_FechaDesde DATETIME, @LD_FechaHasta DATETIME

				SET @LD_FechaDesde = '" + LS_FechaDesdeReporte + @"'
				SET @LD_FechaHasta = '" + LS_FechaHastaReporte + @"'

				INSERT INTO @ReporteCobros(Factura,Recibo,FechaRecibo,ConceptoPago, Cliente, NombreCliente,Vendedor,NombreVendedor,Importe, DescuentoProntoPago)

				SELECT IIF(CD.ID_MOV = 3, 
				(SELECT TOP 1 CM.NO_FACTURA FROM CUEN_M01 AS CM WHERE CM.DOCTO = CD.REFER)
				,CD.REFER),
				(SELECT TOP 1 DOCTO FROM CUEN_DET01 CDA WHERE CDA.REFER = CD.REFER AND CDA.FECHA_APLI >= @LD_FechaDesde AND CDA.FECHA_APLI <= @LD_FechaHasta AND CDA.NUM_CPTO <> 24 AND CDA.NUM_CPTO <> 13 ORDER BY CDA.FECHA_APLI DESC) AS Recibo,
				(SELECT TOP 1 FECHA_APLI FROM CUEN_DET01 CDA WHERE CDA.REFER = CD.REFER AND CDA.FECHA_APLI >= @LD_FechaDesde AND CDA.FECHA_APLI <= @LD_FechaHasta AND CDA.NUM_CPTO <> 24 AND CDA.NUM_CPTO <> 13 AND CDA.NUM_CPTO <> 17 AND CDA.NUM_CPTO <> 25 ORDER BY CDA.FECHA_APLI DESC) AS FechaRecibo,
				(SELECT TOP 1 NUM_CPTO FROM CUEN_DET01 CDA WHERE CDA.REFER = CD.REFER AND CDA.FECHA_APLI >= @LD_FechaDesde AND CDA.FECHA_APLI <= @LD_FechaHasta AND CDA.NUM_CPTO <> 24 AND CDA.NUM_CPTO <> 13 AND CDA.NUM_CPTO <> 17 AND CDA.NUM_CPTO <> 25 ORDER BY CDA.FECHA_APLI DESC) AS ConceptoPago,
				CD.CVE_CLIE,CL.NOMBRE,CD.STRCVEVEND,VE.NOMBRE, SUM(CD.IMPORTE),
				ISNULL((SELECT SUM(CDA.IMPORTE) FROM CUEN_DET01 CDA WHERE CDA.REFER = CD.REFER AND CDA.FECHA_APLI >= @LD_FechaDesde AND CDA.FECHA_APLI <= @LD_FechaHasta AND CDA.NUM_CPTO = 24),0) AS DescuentoProntoPago
				FROM CUEN_DET01 AS CD 
				INNER JOIN CONC01 CO ON CD.NUM_CPTO = CO.NUM_CPTO
				INNER JOIN CLIE01 CL ON CD.CVE_CLIE = CL.CLAVE
				LEFT JOIN VEND01 VE ON CD.STRCVEVEND = VE.CVE_VEND 
				WHERE FECHA_APLI >= @LD_FechaDesde AND FECHA_APLI <= @LD_FechaHasta AND CO.ES_FMA_PAG = 'S' AND CO.NUM_CPTO <> 24 AND CO.NUM_CPTO <> 13 AND CO.NUM_CPTO <> 17 AND CO.NUM_CPTO <> 25
				GROUP BY CD.ID_MOV, CD.REFER,CD.CVE_CLIE,CL.NOMBRE,CD.STRCVEVEND,VE.NOMBRE

				-- Insertamos las facturas pagadas
				INSERT INTO @ReporteVentas(Factura,Fecha,Producto,Cantidad,Categoria,NombreCategoria,SubCategoria,NombreSubCategoria,Casa,NombreCasa,Precio,PrecioLista1, SubTotal,Impuesto,Total)
				SELECT 
				FE.CVE_DOC,FE.FECHA_DOC, FD.CVE_ART,FD.CANT,
				SUBSTRING( IV.LIN_PROD,1,1) AS Categoria,
				(SELECT DESC_LIN FROM CLIN01 AA WHERE SUBSTRING( AA.CVE_LIN,1,1) = SUBSTRING( IV.LIN_PROD,1,1) AND LEN( AA.CVE_LIN ) = 1 ) AS NombreCategoria,
				SUBSTRING( IV.LIN_PROD,2,1) AS SubCategoria,
				(SELECT DISTINCT DESC_LIN FROM CLIN01 AA WHERE SUBSTRING( AA.CVE_LIN,2,1) = SUBSTRING( IV.LIN_PROD,2,1) AND LEN( AA.CVE_LIN ) = 2) AS NombreSubCategoria,
				SUBSTRING( IV.LIN_PROD,3,2) AS Casa,
				(SELECT DISTINCT DESC_LIN FROM CLIN01 AA WHERE SUBSTRING( AA.CVE_LIN,3,2) = SUBSTRING( IV.LIN_PROD,3,2) AND LEN( AA.CVE_LIN ) = 4) AS NombreCasa,				
				FD.PREC-(FD.PREC * FD.DESC1)/100 AS PREC, (SELECT CAMPLIB1 FROM PAR_FACTF_CLIB01 AS PXP WHERE PXP.CLAVE_DOC = FD.CVE_DOC AND PXP.NUM_PART = FD.NUM_PAR) AS PrecioLista1,
				FD.CANT * (FD.PREC-(FD.PREC * FD.DESC1)/100) AS SubTotal, FD.TOTIMP4,FD.CANT * (FD.PREC-(FD.PREC * FD.DESC1)/100) + FD.TOTIMP4 AS TOTAL
				FROM FACTF01 FE
				INNER JOIN PAR_FACTF01 FD ON FE.CVE_DOC = FD.CVE_DOC
				INNER JOIN INVE01 IV ON FD.CVE_ART = IV.CVE_ART
				INNER JOIN @ReporteCobros RC ON FE.CVE_DOC = RC.Factura

				-- Insertamos las facturas pagadas del saldo inicial
				INSERT INTO @ReporteVentas(Factura,Fecha,Producto,Cantidad,Categoria,NombreCategoria,SubCategoria,NombreSubCategoria,Casa,NombreCasa,Precio,SubTotal,Impuesto,Total)
				SELECT RC.Factura,CE.FECHA_APLI,'N/A' AS Producto, 0 AS Cantidad,'N/A' AS Categoria, 'N/A' AS NombreCategoria, 'N/A' AS SubCategoria, 'N/A' AS NombreSubCategoria,
				'N/A' AS Casa, 'N/A' AS NombreCasa,0 AS Precio,CE.IMPORTE AS SubTotal, 0 AS Impuesto, CE.IMPORTE AS Total
					FROM @ReporteCobros AS RC
				LEFT JOIN FACTF01 AS FE ON RC.Factura = FE.CVE_DOC
				INNER JOIN CUEN_M01 AS CE ON RC.Factura = CE.REFER
				WHERE FE.CVE_DOC IS NULL AND CE.NUM_CPTO <> 3

				INSERT INTO @ReporteBaseComision

				SELECT RC.Factura,RV.Fecha AS FechaFactura,RC.Recibo,RC.FechaRecibo, RC.ConceptoPago, RC.Cliente,RC.NombreCliente,RC.Vendedor,RC.NombreVendedor,RC.Importe AS MontoRecibo,RC.DescuentoProntoPago,
				RV.Impuesto,RV.TotalGenericoPrecioPublico,RV.TotalGenericoMenorPrecioPublico, RV.TotalEtico,RV.TotalEquipoMedicoPrecioPublico, RV.TotalEquipoMedicoMenorPrecioPublico, RV.Otros,RV.Total AS TotalFactura,RV.PorcImpuesto,RV.PorcTotalGenericoPrecioPublico, RV.PorcTotalGenericoMenorPrecioPublico, RV.PorcTotalEtico,PorcTotalEquipoMedicoPrecioPublico,PorcTotalEquipoMedicoMenorPrecioPublico,PorcOtros,
				RC.Importe * RV.PorcImpuesto AS MontoCobroImpuesto, RC.Importe * RV.PorcTotalGenericoPrecioPublico AS MontoCobroGenericoPrecioPublico, RC.Importe * RV.PorcTotalGenericoMenorPrecioPublico AS MontoCobroGenericoMenorPrecioPublico, RC.Importe * RV.PorcTotalEtico AS MontoCobroEtico, RC.Importe * RV.PorcTotalEquipoMedicoPrecioPublico AS MontoCobroEquipoMedicoPrecioPublico,RC.Importe * RV.PorcTotalEquipoMedicoMenorPrecioPublico AS MontoCobroEquipoMedicoMenorPrecioPublico,
				RC.Importe * RV.PorcOtros AS MontoCobroOtros
					FROM
				(SELECT 
				Factura,Recibo,FechaRecibo,ConceptoPago, Cliente,NombreCliente,Vendedor,NombreVendedor, Importe,DescuentoProntoPago
				FROM @ReporteCobros AS RC) RC

				INNER JOIN

				(SELECT 
				Factura,Fecha,Impuesto,TotalGenericoPrecioPublico,TotalGenericoMenorPrecioPublico,TotalEtico,TotalEquipoMedicoPrecioPublico,TotalEquipoMedicoMenorPrecioPublico,Otros,Total,
				IIF(Impuesto>0,Impuesto/Total,0) AS PorcImpuesto,  
				IIF(TotalGenericoPrecioPublico>0,TotalGenericoPrecioPublico/Total,0) AS PorcTotalGenericoPrecioPublico,
				IIF(TotalGenericoMenorPrecioPublico>0,TotalGenericoMenorPrecioPublico/Total,0) AS PorcTotalGenericoMenorPrecioPublico,
				IIF(TotalEtico>0,TotalEtico/Total,0) AS PorcTotalEtico,
				IIF(TotalEquipoMedicoPrecioPublico>0,TotalEquipoMedicoPrecioPublico/Total,0) AS PorcTotalEquipoMedicoPrecioPublico,
				IIF(TotalEquipoMedicoMenorPrecioPublico>0,TotalEquipoMedicoMenorPrecioPublico/Total,0) AS PorcTotalEquipoMedicoMenorPrecioPublico,
				IIF(Otros>0,Otros/Total,0) AS PorcOtros
				FROM
					(SELECT 
					Factura,Fecha,SUM(Impuesto) AS Impuesto,
					SUM(IIF(Categoria='G',IIF(SubCategoria <> 'E', IIF(Precio >= PrecioLista1,SubTotal,0), 0),0)) AS TotalGenericoPrecioPublico,
					SUM(IIF(Categoria='G',IIF(SubCategoria <> 'E', IIF(Precio < PrecioLista1,SubTotal,0), 0),0)) AS TotalGenericoMenorPrecioPublico,
					SUM(IIF(Categoria='E',IIF(SubCategoria <> 'E', SubTotal, 0),0)) AS TotalEtico,
					SUM(IIF(SubCategoria = 'E', IIF(Precio >= PrecioLista1,SubTotal,0), 0)) AS TotalEquipoMedicoPrecioPublico,
					SUM(IIF(SubCategoria = 'E', IIF(Precio < PrecioLista1,SubTotal,0), 0)) AS TotalEquipoMedicoMenorPrecioPublico,
					SUM(IIF(Categoria = 'N/A', SubTotal, 0)) AS Otros,
					SUM(Total) AS Total
					FROM @ReporteVentas AS RV
					GROUP BY Factura,Fecha) RV) RV
				ON RC.Factura = RV.Factura
				ORDER BY RC.Factura

				SELECT 
				Vendedor,NombreVendedor,MontoCobroGenericoPrecioPublico,MontoCobroGenericoMenorPrecioPublico,MontoCobroEtico,MontoCobroEquipoMedicoPrecioPublico,MontoCobroEquipoMedicoMenorPrecioPublico,
				MontoCobroOtros,ComisionGenericoPrecioPublico,ComisionGenericoMenorPrecioPublico,ComisionEtico,ComisionEquipoMedicoPrecioPublico,ComisionEquipoMedicoMenorPrecioPublico,ComisionOtros,
				(ComisionGenericoPrecioPublico+ComisionGenericoMenorPrecioPublico+ComisionEtico+ComisionEquipoMedicoPrecioPublico+ComisionEquipoMedicoMenorPrecioPublico+ComisionOtros) AS TotalComision
				FROM 
					(SELECT 
					Vendedor,NombreVendedor,MontoCobroGenericoPrecioPublico,MontoCobroGenericoMenorPrecioPublico,MontoCobroEtico,MontoCobroEquipoMedicoPrecioPublico,MontoCobroEquipoMedicoMenorPrecioPublico,MontoCobroOtros,
					IIF(MontoCobroGenericoPrecioPublico <= 100000,MontoCobroGenericoPrecioPublico*0.05,IIF(MontoCobroGenericoPrecioPublico > 100000 AND MontoCobroGenericoPrecioPublico <= 200000,MontoCobroGenericoPrecioPublico*0.06,IIF(MontoCobroGenericoPrecioPublico>200000,MontoCobroGenericoPrecioPublico*0.07,0))) AS ComisionGenericoPrecioPublico,
					MontoCobroGenericoMenorPrecioPublico * 0.03 AS ComisionGenericoMenorPrecioPublico,
					MontoCobroEtico * 0.03 AS ComisionEtico,
					MontoCobroEquipoMedicoPrecioPublico * 0.05 AS ComisionEquipoMedicoPrecioPublico,
					MontoCobroEquipoMedicoMenorPrecioPublico * 0.03 AS ComisionEquipoMedicoMenorPrecioPublico,
					MontoCobroOtros * 0.04 AS ComisionOtros
					FROM
						(SELECT 
						Vendedor,NombreVendedor, 
						SUM(MontoCobroGenericoPrecioPublico) AS MontoCobroGenericoPrecioPublico, SUM(MontoCobroGenericoMenorPrecioPublico) AS MontoCobroGenericoMenorPrecioPublico,
						SUM(MontoCobroEtico) AS MontoCobroEtico, SUM(MontoCobroEquipoMedicoPrecioPublico) AS MontoCobroEquipoMedicoPrecioPublico, SUM(MontoCobroEquipoMedicoMenorPrecioPublico) AS MontoCobroEquipoMedicoMenorPrecioPublico,
						SUM(MontoCobroOtros) AS MontoCobroOtros
						FROM @ReporteBaseComision
						WHERE NombreVendedor IS NOT NULL
						GROUP BY Vendedor,NombreVendedor) RV) RV
				ORDER BY Vendedor

				SELECT * FROM @ReporteBaseComision ORDER BY Factura
                    
                ";

			return sqlComisiones;
		}


		#endregion

	}
}
