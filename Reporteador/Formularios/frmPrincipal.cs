using DevExpress.XtraBars;
using DevExpress.XtraBars.Docking2010;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Reporteador
{
    public partial class FrmPrincipal : DevExpress.XtraBars.FluentDesignSystem.FluentDesignForm
    {
        public FrmPrincipal()
        {
            InitializeComponent();
        }

        // Declaramos variables y metodos globales
        private InvFun.Funciones LF_Funciones = new InvFun.Funciones();
        private string LS_Conexion = "";

        // Definimos los dataSets
        DataSet dsConfiguracionEmpresa = new DataSets.dsConfiguracionEmpresa();
        DataSet dsTempConfiguracionEmpresa = new DataSets.dsConfiguracionEmpresa();
        DataSet dsConfiguracionPerfiles = new DataSets.dsConfiguracionPerfiles();
        DataSet dsTempConfiguracionPerfiles = new DataSets.dsConfiguracionPerfiles();

        // Definimos los DataTable
        private DataTable LO_TablaReporteComisionPorVendedor = new DataTable();
        private DataTable LO_TablaReporteComisionDetalle = new DataTable();
        private DataTable LO_TablaReporteVentasProductoPrecio = new DataTable();
        private DataTable LO_TablaReporteVentasProductoBonificacion = new DataTable();
        private DataTable LO_TablaReporteAntiguedadClienteZona = new DataTable();

        private void FrmPrincipal_Load(object sender, EventArgs e)
        {
            // Leemos los datos de conexion, almacenados en el archivo ConfiguracionEmpresa.cfg
            try
            {
                dsConfiguracionEmpresa.ReadXml("ConfiguracionEmpresa.cfg");
                //dgvConfiguracionEmpresa.DataSource = dsConfiguracionEmpresa.Tables[0];
                // DateTime.Now.ToString("dd/MM/yyyy")
                dateEditComisionesFchaDesde.DateTime = DateTime.Now;
                dateEditComisionesFchaHasta.DateTime = DateTime.Now;
                dateEditDesdeVentasProductoPrecio.DateTime = DateTime.Now;
                dateEditHastaVentasProductoPrecio.DateTime = DateTime.Now;
                dateEditDesdeVentasProductoBonificacion.DateTime = DateTime.Now;
                dateEditHastaVentasProductoBonificacion.DateTime = DateTime.Now;
                dateEditHastaAntiguedadClienteZona.DateTime = DateTime.Now;
            }
            catch (Exception exConfiguracionEmpresa)
            {
                MessageBox.Show("No se encontro el archivo de configuración de empresa, favor revisar los parametros\n" + exConfiguracionEmpresa.ToString());
                /*
                tabControlPrincipal.SelectedIndex = 10;
                tabControlConfiguracion.SelectedIndex = 0;
                textBoxParametrosEmpresa.Focus();
                */
            }
            try
            {
                dsConfiguracionPerfiles.ReadXml("ConfiguracionPerfiles.cfg");
                //DgvConfiguracionPerfiles.DataSource = dsConfiguracionPerfiles.Tables[0];
            }
            catch (Exception exConfiguracionPerfiles)
            {
                MessageBox.Show("No se encontro el archivo de configuración de perfiles, favor revisar los parametros\n" + exConfiguracionPerfiles.ToString());
                /*
                tabControlPrincipal.SelectedIndex = 10;
                tabControlConfiguracion.SelectedIndex = 1;
                textBoxPerfilesNombreDeUsuario.Focus();
                */
            }
        }

        private void accordionControlElementComisiones_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("Click comisiones");
            xtraTabPageComisiones.PageVisible = true;
            xtraTabPageVentasProductoTipoPrecio.PageVisible = false;
            xtraTabPageVentasProductoBonificacion.PageVisible = false;
            xtraTabPageAntiguedadSaldosClienteZona.PageVisible = false;
            xtraTabPageDefinirBonificaciones.PageVisible = false;
        }
        // al dar click en el boton ventas por producto y tipo de precios
        private void accordionControlElement2_Click(object sender, EventArgs e)
        {
            xtraTabPageVentasProductoTipoPrecio.PageVisible = true;
            xtraTabPageComisiones.PageVisible = false;
            xtraTabPageVentasProductoBonificacion.PageVisible = false;
            xtraTabPageAntiguedadSaldosClienteZona.PageVisible = false;
            xtraTabPageDefinirBonificaciones.PageVisible = false;
        }
        // al dar click en ventas por producto bonificacion
        private void accordionControlElement3_Click(object sender, EventArgs e)
        {
            xtraTabPageVentasProductoBonificacion.PageVisible = true;
            xtraTabPageVentasProductoTipoPrecio.PageVisible = false;
            xtraTabPageComisiones.PageVisible = false;
            xtraTabPageAntiguedadSaldosClienteZona.PageVisible = false;
            xtraTabPageDefinirBonificaciones.PageVisible = false;
        }
        // al dar click en antiguedad por cliente y zona
        private void accordionControlElement5_Click(object sender, EventArgs e)
        {
            xtraTabPageAntiguedadSaldosClienteZona.PageVisible = true;
            xtraTabPageVentasProductoBonificacion.PageVisible = false;
            xtraTabPageVentasProductoTipoPrecio.PageVisible = false;
            xtraTabPageComisiones.PageVisible = false;
            xtraTabPageDefinirBonificaciones.PageVisible = false;
        }
        // al dar click en definir configuraciones
        private void accordionControlElement6_Click(object sender, EventArgs e)
        {
            xtraTabPageDefinirBonificaciones.PageVisible = true;
            xtraTabPageAntiguedadSaldosClienteZona.PageVisible = false;
            xtraTabPageVentasProductoBonificacion.PageVisible = false;
            xtraTabPageVentasProductoTipoPrecio.PageVisible = false;
            xtraTabPageComisiones.PageVisible = false;
        }

        private void windowsUIButtonPanel1_ButtonClick(object sender, DevExpress.XtraBars.Docking2010.ButtonEventArgs e)
        {
            WindowsUIButton btn = e.Button as WindowsUIButton;
            if (btn.Tag != null && btn.Tag.Equals("Filtrar"))
            {
                //MessageBox.Show("Click filtrar");

                // vaciamos los objetos de ventas por ruta
                LO_TablaReporteComisionDetalle.Clear();
                gridControlDetalleComision.DataSource = null;
                //this.LF_Funciones.ResultadoSQLSERVER.Tables.Clear();
                this.LF_Funciones.LimpiarResultadoSQLSERVER();

                // Llenamos la cadena de conexion con el archivo de configuracion xml
                LS_Conexion =
                    "server=" + dsConfiguracionEmpresa.Tables[0].Rows[0][5].ToString() +
                    ";database=" + dsConfiguracionEmpresa.Tables[0].Rows[0][8].ToString() +
                    ";Persist Security Info=True;User ID=" + dsConfiguracionEmpresa.Tables[0].Rows[0][6].ToString() +
                    "; Password=" + dsConfiguracionEmpresa.Tables[0].Rows[0][7].ToString() + "";

                // Enlazamos la cadena de conexion a la funcion SQL
                LF_Funciones.SQLSERVERConexion.ConnectionString = LS_Conexion.Trim();

                // Llenamos el string comando SQL con el Query de ventas alojado en la funcion ConsultaSQLVentas y le enviamos la fecha como parametro
                string LS_ComandoSQL = LF_Funciones.ConsultaSQLComisiones(dateEditComisionesFchaDesde.DateTime.ToString("yyyy/MM/dd"), dateEditComisionesFchaHasta.DateTime.ToString("yyyy/MM/dd"));
                //string LS_ComandoSQL = LF_Funciones.ConsultaSQLComisiones("2021/09/01","2021/09/30");

                // Ejecutamos el comando SQL
                if (LF_Funciones.EjecutarComandoSQLSERVER(LS_ComandoSQL, true, true) == false)
                {
                    // Validamos que la consulta haya devuelto registros
                    if (this.LF_Funciones.ResultadoSQLSERVER.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show("No hay filas para mostrar", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    // verificamos que la tabla 0 del dataTable no este vacia
                    if (this.LF_Funciones.ResultadoSQLSERVER.Tables[0].Rows.Count > 0)
                    {

                        // llenamos los dataTablet de ventas por ruta
                        LO_TablaReporteComisionPorVendedor = this.LF_Funciones.ResultadoSQLSERVER.Tables[0];
                        LO_TablaReporteComisionDetalle = this.LF_Funciones.ResultadoSQLSERVER.Tables[1];

                        // Enlazamos el dataTable al dataGridView                        
                        this.gridControlDetalleComision.DataSource = LO_TablaReporteComisionDetalle;
                        this.gridControlComisionPorVendedor.DataSource = LO_TablaReporteComisionPorVendedor;


                        gridView1.OptionsView.ColumnAutoWidth = false;
                        gridView1.OptionsView.BestFitMaxRowCount = -1;
                        gridView1.BestFitColumns();

                        gridView2.OptionsView.ColumnAutoWidth = false;
                        gridView2.OptionsView.BestFitMaxRowCount = -1;
                        gridView2.BestFitColumns();

                    }
                }

            }
            if (btn.Tag != null && btn.Tag.Equals("Exportar"))
            {
                //MessageBox.Show("Click Exportar");

                if(comboBoxEditSeleccionReporte.Text == "Comisión Detalle")
                {

                    if (gridView1.RowCount > 0)
                    {
                        string path = "ComisionDetalle.xlsx";
                        gridControlDetalleComision.ExportToXlsx(path);
                        // Open the created XLSX file with the default application.
                        Process.Start(path);
                    }
                    else
                    {
                        MessageBox.Show("Debe realizar un filtro para poder exportar a Excel");
                    }

                }
                else if(comboBoxEditSeleccionReporte.Text == "Comisión Resumen")
                {
                    if (gridView1.RowCount > 0)
                    {
                        string path = "ComisionResumen.xlsx";
                        gridControlComisionPorVendedor.ExportToXlsx(path);
                        // Open the created XLSX file with the default application.
                        Process.Start(path);
                    }
                    else
                    {
                        MessageBox.Show("Debe realizar un filtro para poder exportar a Excel");
                    }
                }
                else
                {
                    MessageBox.Show("Debe seleccionar un reporte para Exportar");
                }

            }
            if (btn.Tag != null && btn.Tag.Equals("Imprimir"))
            {
                MessageBox.Show("Click Imprimir");
            }
        }

        private void windowsUIButtonPanelVentasProductoPrecio_ButtonClick(object sender, ButtonEventArgs e)
        {
            WindowsUIButton btn = e.Button as WindowsUIButton;
            if (btn.Tag != null && btn.Tag.Equals("Filtrar"))
            {
                //MessageBox.Show("Click filtrar");

                // vaciamos los objetos de ventas por ruta
                LO_TablaReporteVentasProductoPrecio.Clear();
                gridControlVentasProductoPrecio.DataSource = null;
                this.LF_Funciones.LimpiarResultadoSQLSERVER();

                // Llenamos la cadena de conexion con el archivo de configuracion xml
                LS_Conexion =
                    "server=" + dsConfiguracionEmpresa.Tables[0].Rows[0][5].ToString() +
                    ";database=" + dsConfiguracionEmpresa.Tables[0].Rows[0][8].ToString() +
                    ";Persist Security Info=True;User ID=" + dsConfiguracionEmpresa.Tables[0].Rows[0][6].ToString() +
                    "; Password=" + dsConfiguracionEmpresa.Tables[0].Rows[0][7].ToString() + "";

                // Enlazamos la cadena de conexion a la funcion SQL
                LF_Funciones.SQLSERVERConexion.ConnectionString = LS_Conexion.Trim();

                // Llenamos el string comando SQL con el Query de ventas alojado en la funcion ConsultaSQLVentas y le enviamos la fecha como parametro
                string LS_ComandoSQL = LF_Funciones.ConsultaSQLVentasProductoPrecio(dateEditDesdeVentasProductoPrecio.DateTime.ToString("yyyy/MM/dd"), dateEditHastaVentasProductoPrecio.DateTime.ToString("yyyy/MM/dd"));
                //string LS_ComandoSQL = LF_Funciones.ConsultaSQLComisiones("2021/09/01","2021/09/30");

                // Ejecutamos el comando SQL
                if (LF_Funciones.EjecutarComandoSQLSERVER(LS_ComandoSQL, true, true) == false)
                {
                    // Validamos que la consulta haya devuelto registros
                    if (this.LF_Funciones.ResultadoSQLSERVER.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show("No hay filas para mostrar", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    // verificamos que la tabla 0 del dataTable no este vacia
                    if (this.LF_Funciones.ResultadoSQLSERVER.Tables[0].Rows.Count > 0)
                    {

                        // llenamos los dataTablet de ventas por ruta
                        LO_TablaReporteVentasProductoPrecio = this.LF_Funciones.ResultadoSQLSERVER.Tables[0];                        

                        // Enlazamos el dataTable al dataGridView                        
                        this.gridControlVentasProductoPrecio.DataSource = LO_TablaReporteVentasProductoPrecio;
                        

                        gridView3.OptionsView.ColumnAutoWidth = false;
                        gridView3.OptionsView.BestFitMaxRowCount = -1;
                        gridView3.BestFitColumns();

                    }
                }

            }
            if (btn.Tag != null && btn.Tag.Equals("Exportar"))
            {
                //MessageBox.Show("Click Exportar");

                if (gridView3.RowCount > 0)
                {
                    string path = "VentasProductoPrecio.xlsx";
                    gridControlVentasProductoPrecio.ExportToXlsx(path);
                    // Open the created XLSX file with the default application.
                    Process.Start(path);
                }
                else
                {
                    MessageBox.Show("Debe realizar un filtro para poder exportar a Excel");
                }

            }
            if (btn.Tag != null && btn.Tag.Equals("Imprimir"))
            {
                MessageBox.Show("Click Imprimir");
            }

        }

        private void windowsUIButtonPanelVentasProductoBonificacion_ButtonClick(object sender, ButtonEventArgs e)
        {

            WindowsUIButton btn = e.Button as WindowsUIButton;
            if (btn.Tag != null && btn.Tag.Equals("Filtrar"))
            {
                //MessageBox.Show("Click filtrar");

                // vaciamos los objetos de ventas por ruta
                LO_TablaReporteVentasProductoBonificacion.Clear();
                gridControlVentasProductoBonificacion.DataSource = null;
                this.LF_Funciones.LimpiarResultadoSQLSERVER();

                // Llenamos la cadena de conexion con el archivo de configuracion xml
                LS_Conexion =
                    "server=" + dsConfiguracionEmpresa.Tables[0].Rows[0][5].ToString() +
                    ";database=" + dsConfiguracionEmpresa.Tables[0].Rows[0][8].ToString() +
                    ";Persist Security Info=True;User ID=" + dsConfiguracionEmpresa.Tables[0].Rows[0][6].ToString() +
                    "; Password=" + dsConfiguracionEmpresa.Tables[0].Rows[0][7].ToString() + "";

                // Enlazamos la cadena de conexion a la funcion SQL
                LF_Funciones.SQLSERVERConexion.ConnectionString = LS_Conexion.Trim();

                // Llenamos el string comando SQL con el Query de ventas alojado en la funcion ConsultaSQLVentas y le enviamos la fecha como parametro
                string LS_ComandoSQL = LF_Funciones.ConsultaSQLVentasProductoBonificacion(dateEditDesdeVentasProductoBonificacion.DateTime.ToString("yyyy/MM/dd"), dateEditHastaVentasProductoBonificacion.DateTime.ToString("yyyy/MM/dd"));
                //string LS_ComandoSQL = LF_Funciones.ConsultaSQLComisiones("2021/09/01","2021/09/30");

                // Ejecutamos el comando SQL
                if (LF_Funciones.EjecutarComandoSQLSERVER(LS_ComandoSQL, true, true) == false)
                {
                    // Validamos que la consulta haya devuelto registros
                    if (this.LF_Funciones.ResultadoSQLSERVER.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show("No hay filas para mostrar", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    // verificamos que la tabla 0 del dataTable no este vacia
                    if (this.LF_Funciones.ResultadoSQLSERVER.Tables[0].Rows.Count > 0)
                    {

                        // llenamos los dataTablet de ventas por ruta
                        LO_TablaReporteVentasProductoBonificacion = this.LF_Funciones.ResultadoSQLSERVER.Tables[0];

                        // Enlazamos el dataTable al dataGridView                        
                        this.gridControlVentasProductoBonificacion.DataSource = LO_TablaReporteVentasProductoBonificacion;


                        gridView4.OptionsView.ColumnAutoWidth = false;
                        gridView4.OptionsView.BestFitMaxRowCount = -1;
                        gridView4.BestFitColumns();

                    }
                }

            }
            if (btn.Tag != null && btn.Tag.Equals("Exportar"))
            {
                //MessageBox.Show("Click Exportar");

                if (gridView4.RowCount > 0)
                {
                    string path = "VentasProductoBonificacion.xlsx";
                    gridControlVentasProductoBonificacion.ExportToXlsx(path);
                    // Open the created XLSX file with the default application.
                    Process.Start(path);
                }
                else
                {
                    MessageBox.Show("Debe realizar un filtro para poder exportar a Excel");
                }

            }
            if (btn.Tag != null && btn.Tag.Equals("Imprimir"))
            {
                MessageBox.Show("Click Imprimir");
            }

        }

        private void windowsUIButtonPanelAntiguedadClienteZona_ButtonClick(object sender, ButtonEventArgs e)
        {

            WindowsUIButton btn = e.Button as WindowsUIButton;
            if (btn.Tag != null && btn.Tag.Equals("Filtrar"))
            {
                //MessageBox.Show("Click filtrar");

                // vaciamos los objetos de ventas por ruta
                LO_TablaReporteAntiguedadClienteZona.Clear();
                gridControlAntiguedadClienteZona.DataSource = null;
                this.LF_Funciones.LimpiarResultadoSQLSERVER();

                // Llenamos la cadena de conexion con el archivo de configuracion xml
                LS_Conexion =
                    "server=" + dsConfiguracionEmpresa.Tables[0].Rows[0][5].ToString() +
                    ";database=" + dsConfiguracionEmpresa.Tables[0].Rows[0][8].ToString() +
                    ";Persist Security Info=True;User ID=" + dsConfiguracionEmpresa.Tables[0].Rows[0][6].ToString() +
                    "; Password=" + dsConfiguracionEmpresa.Tables[0].Rows[0][7].ToString() + "";

                // Enlazamos la cadena de conexion a la funcion SQL
                LF_Funciones.SQLSERVERConexion.ConnectionString = LS_Conexion.Trim();

                // Llenamos el string comando SQL con el Query de ventas alojado en la funcion ConsultaSQLVentas y le enviamos la fecha como parametro
                string LS_ComandoSQL = LF_Funciones.ConsultaSQLAntiguedadClienteZona(dateEditHastaAntiguedadClienteZona.DateTime.ToString("yyyy/MM/dd"));
                //string LS_ComandoSQL = LF_Funciones.ConsultaSQLComisiones("2021/09/01","2021/09/30");

                // Ejecutamos el comando SQL
                if (LF_Funciones.EjecutarComandoSQLSERVER(LS_ComandoSQL, true, true) == false)
                {
                    // Validamos que la consulta haya devuelto registros
                    if (this.LF_Funciones.ResultadoSQLSERVER.Tables[0].Rows.Count == 0)
                    {
                        MessageBox.Show("No hay filas para mostrar", "Informacion", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        return;
                    }
                    // verificamos que la tabla 0 del dataTable no este vacia
                    if (this.LF_Funciones.ResultadoSQLSERVER.Tables[0].Rows.Count > 0)
                    {

                        // llenamos los dataTablet de ventas por ruta
                        LO_TablaReporteAntiguedadClienteZona = this.LF_Funciones.ResultadoSQLSERVER.Tables[0];

                        // Enlazamos el dataTable al dataGridView                        
                        this.gridControlAntiguedadClienteZona.DataSource = LO_TablaReporteAntiguedadClienteZona;


                        gridView5.OptionsView.ColumnAutoWidth = false;
                        gridView5.OptionsView.BestFitMaxRowCount = -1;
                        gridView5.BestFitColumns();

                    }
                }

            }
            if (btn.Tag != null && btn.Tag.Equals("Exportar"))
            {
                //MessageBox.Show("Click Exportar");

                if (gridView5.RowCount > 0)
                {
                    string path = "AntiguedadClienteZona.xlsx";
                    gridControlAntiguedadClienteZona.ExportToXlsx(path);
                    // Open the created XLSX file with the default application.
                    Process.Start(path);
                }
                else
                {
                    MessageBox.Show("Debe realizar un filtro para poder exportar a Excel");
                }

            }
            if (btn.Tag != null && btn.Tag.Equals("Imprimir"))
            {
                MessageBox.Show("Click Imprimir");
            }

        }
    }
}
