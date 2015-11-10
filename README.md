# factura-electronica
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using System.Collections;
using System.Collections.Specialized;
using CustomListView;
using System.Data.SqlClient;
using System.Xml;
using System.Xml.Schema;
using System.Xml.Linq;
using System.Security.Cryptography.X509Certificates;
using MessagingToolkit.QRCode.Codec;
using System.Threading.Tasks;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.Threading;

// 6 de Mayo de 2014
// se aplico a la ultima actualizacion un catalogo de permisos de usuarios, modificacion del xml para que salgan
// los nodos de retencion de impuestos asi como asignar el valor del subtotal al total por la razon de retencion

// 8 de Mayo de 2014
// se modifico la generacion del pdf acorde a el iva retenido

// 12 de Mayo 2014
// se agrego un permiso solicitado en base 21 el cual es: 
// impresion de documentos, nota credito y nota credito electronica

//24 de Julio 2015
// Se modifico para que el tipo de documento lo hiciera de manera automatica si es factura es ingreso y si es Nota de credito egreso

namespace FacturaElectronica
{
    public partial class frmFactura : Form
    {
        
        string sMoneda;
        string sMoneda_ini;
        string sDetalle="";
        int iTipoFactura;
        int i = 0;
        int f = 0;
        int vMov, nFila; 
        string[,] arrayObtenerArticulos;
        string[,] arrayArticulos;
        int incremento;
        string[] sDatoLeido;
        ArrayList aUnidadMedida;
        StringCollection sUnidadMedida;
        string sTipoDocumento;
        List<double> lPrecioUnitario;
        List<double> lImporte;
        bool bXmlOk;
        int iXmlArticulos;
        int iTiempoTimbrado;
        string sConsultaSql;
        bool bDiferenciaIva;
        bool bGuardarXml;

        bool bFacturaElectronica = false; 
        bool bImpresionDocumentos = false; 
        bool bNotaCredito = false; 
        bool bNotaCreditoElectronica = false;
        bool bPrueba = false;

        cCliente cliente;
        cEmpresa empresa;
        cArticulos articulo;
        cLeerXml leerXml;
        
        //declaracion de componenetes OLEDB para establecer la conexion y obtener datos
        OleDbConnection conexion;
        SqlConnection sqlConexion;

        //lectura de las diferentes tablas de Megapaq
        OleDbDataReader leer, leer2, leer3, leer4, leer5, leerProveedor;
        OleDbDataReader leer6, leer7, leer8, leer9, leer10, leer11;
        SqlDataReader sqlLeer, sqlLeerBuscarFacturaBD;

        OleDbCommand comando, comando2, comando3, comando4, comando5, comando6;
        OleDbCommand comando7, comando8, comando9, comando10, comando11, comandoProveedor;
        SqlCommand sqlComando, sqlComando2;

        RadioButton[] rbMetodoPago;
        RadioButton[] rbMetodoPagoNE;
        string sMetodoPago;
        string sTipoIva;

        //coleccion que guarda los elementos leidos de la consulta realizada en el load
        AutoCompleteStringCollection coleccion;
        AutoCompleteStringCollection coleccionNotas;

        //variables para almacenar las consultas para las tablas de megapaq
        string consulta, consulta2, consulta3, consulta4, consulta5, consulta6, consulta7;
        string consulta8, consulta9, consulta10, consulta11;

        string cadenaConexion;

        private FacturaElectronica.leerArchivoConfig config;

        public frmFactura(FacturaElectronica.leerArchivoConfig config)
        {
            this.config = config;
            InitializeComponent();
        }

        #region notaCreditoElectronica

            private void limpiarPantallaNE()
            {
                //foreach (TabPage tab in tabControl1.TabPages)
                //{
                //    IEnumerable<TextBox> texts = tab.Controls.OfType<TextBox>();

                //    if (tab.Name == "tabPage4")
                //    {
                //        foreach (TextBox text in texts)
                //        {
                //            text.Text = "";
                //        }
                //    }
                //}
                txtNEcliente.Text = "";
                txtNEcodigo.Text = "";
                txtNEcolonia.Text = "";
                txtNEcp.Text = "";
                txtNEdireccion.Text = "";
                txtNEestado.Text = "";
                txtNEFactura.Text = "";
                txtNEfecha.Text = "";
                txtNEformapago.Text = "";
                txtNEimporte.Text = "";
                txtNEiva.Text = "";
                txtNEmoneda.Text = "";
                txtNEordencompra.Text = "";
                txtNEpoblacion.Text = "";
                txtNEreferencia.Text = "";
                txtNErfc.Text = "";
                txtNEserie.Text = "";
                txtNEsubtotal.Text = "";
                txtNEtelefono.Text = "";
                txtNEtotal.Text = "";
                txtBancoNE.Text = "";
                txtDetallesNotaCredito.Text = "";
                rbMetodoPagoNE[1].Checked = true;
                cmbNEiva.SelectedItem = "16";
                txtNEMensajes.Text = "";
                txtNEMensajes.BackColor = Color.Silver;
                listNEarticulos.Items.Clear();
                grbEstadoSistemaNE.BackColor = Color.Transparent;
                rdbIvaNERet.Checked = false;
                rdbIvaNETrans.Checked = true;
                tabControl1.Update();
            }

            private void mostrarDatosPantallaNE()
            {
                if (cliente.sCodigo == "C0000" || cliente.sCodigo == "D0000")
                {
                    txtNEserie.ReadOnly = false;
                    txtNErfc.ReadOnly = false;
                    txtNEordencompra.ReadOnly = false;
                    txtNEpoblacion.ReadOnly = false;
                    txtNEreferencia.ReadOnly = false;
                    txtNEcp.ReadOnly = false;
                    txtNEestado.ReadOnly = false;
                    txtNEcolonia.ReadOnly = false;
                    txtNEtelefono.ReadOnly = false;
                    txtNEpoblacion.ReadOnly = false;
                    txtNEdireccion.ReadOnly = false;
                    txtNEcliente.ReadOnly = false;
                }

                else
                {
                    txtNEserie.ReadOnly = true;
                    txtNErfc.ReadOnly = true;
                    txtNEordencompra.ReadOnly = true;
                    txtNEpoblacion.ReadOnly = true;
                    txtNEreferencia.ReadOnly = true;
                    txtNEcp.ReadOnly = true;
                    txtNEestado.ReadOnly = true;
                    txtNEcolonia.ReadOnly = true;
                    txtNEtelefono.ReadOnly = true;
                    txtNEpoblacion.ReadOnly = true;
                    txtNEdireccion.ReadOnly = true;
                    txtNEcliente.ReadOnly = true;
                }

                
                string sTipoMoneda = "";

                if (cliente.sMoneda == "PESOS")
                { sTipoMoneda = "MXN"; }
                if (cliente.sMoneda == "DOLARES")
                { sTipoMoneda = "USD"; }
                if (cliente.sMoneda == "EUROS")
                { sTipoMoneda = "EUR"; }

                //se muestra el dato de la columna especificada en la etiqueta 
                txtNEFactura.Text = cliente.sDocumento;
                txtNEcodigo.Text = cliente.sCodigo;
                txtNEserie.Text = cliente.sSerie;
                txtNErfc.Text = cliente.sRfc;
                txtNEordencompra.Text = cliente.sOrdenCompra;
                txtNEfecha.Text = cliente.sFecha;
                txtNEpoblacion.Text = cliente.sPoblacion;
                txtNEreferencia.Text = cliente.sReferencia;
                txtNEcliente.Text = cliente.sCliente;
                txtNEdireccion.Text = cliente.sColonia;
                txtNEcp.Text = cliente.sCP;
                txtNEtelefono.Text = cliente.sTelefono;
                txtNEcolonia.Text = cliente.sColonia;
                txtNEestado.Text = cliente.sEstado;
                txtNEdireccion.Text = cliente.sDireccion;
                txtNEmoneda.Text = cliente.sMoneda;
                txtNEtotal.Text = Convert.ToString(cliente.dTotal) + " " + sTipoMoneda;
                txtNEsubtotal.Text = Convert.ToString(cliente.sSubtotal) + " " + sTipoMoneda;
                txtNEiva.Text = Convert.ToString(cliente.dIva) + " " + sTipoMoneda;
                txtNEformapago.Text = cliente.sFormaPago;

                if (bDiferenciaIva)
                {
                    cliente.dIva = cliente.sSubtotal;
                    cliente.sSubtotal = 0.00;
                    txtNEiva.Text = txtNEsubtotal.Text;
                    txtNEsubtotal.Text = 0.00 + " " + sTipoMoneda;
                }

                Numeros_letras convertir = new Numeros_letras();

                txtNEimporte.Text = convertir.enletras(Convert.ToString(cliente.dTotal), sTipoMoneda, cliente.sMoneda);
                cliente.sImporteLetra = txtImporte.Text;
                cliente.sTipoMoneda = sTipoMoneda;

                mensajesOk("NOTA DE CRÉDITO ENCONTRADA");

                btnGenerarNE.Enabled = true;

            }

            private void btnNEbuscar_Click(object sender, EventArgs e)
            {
                bNotaCreditoElectronica = true;
                estadoSistema(" Buscando...");
                limpiarPantallaNE();
                mensajesOk("BUSCANDO...");
                lPrecioUnitario = new List<double>();
                lImporte = new List<double>();
                consultasNotaCredito();
                estadoSistema("");
                bNotaCreditoElectronica = false;
            }

            private void btnNEsalir_Click(object sender, EventArgs e)
            {
                if (conexion.State == ConnectionState.Open) { conexion.Close(); }
                Close();
            }

            private void btnGenerarNE_Click(object sender, EventArgs e)
            {
                TimeSpan ts = DateTime.Today.Date - cliente.dtFecha.Date;

                int iDiferenciaDias = ts.Days;

                if (iDiferenciaDias <= config.iDiasTimbrado)
                {
                bNotaCreditoElectronica = true;
                estadoSistema("En espera...");
                mensajesOk("EN ESPERA...");
                sDetalle = txtDetallesNotaCredito.Text;
                facturar();
                estadoSistema("");
                bNotaCreditoElectronica = false;
                btnNEbuscar.Enabled = true;
                ActiveControl = txtNEbuscar;
                }
                else
                {
                    MessageBox.Show("No se puede Timbrar, debido a que la fecha esta fuera del limite permitido ", "Advertencia ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            private void txtBancoNE_KeyPress(object sender, KeyPressEventArgs e)
            {
                if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
                {
                    e.Handled = true;
                }
            }

            private void txtNEbuscar_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                { btnNEbuscar.PerformClick(); }
            }

        #endregion

        #region notaCredito

            private void btnNbuscar_Click(object sender, EventArgs e)
            {
                bNotaCredito = true;
                estadoSistema(" Buscando....");
                limpiarPantallaNotas();
                mensajesOk("BUSCANDO...");
                lPrecioUnitario = new List<double>();
                lImporte = new List<double>();
                consultasNotaCredito();
                estadoSistema("");
                bNotaCredito = false;
            }

            private void consultasNotaCredito()
            {
                try
                {
                    cliente = new cCliente();
                    articulo = new cArticulos();

                    iTipoFactura = 0;
                    vMov = 1;
                    nFila = 1;
                    string sNota = "";
                    string sNotaAño = "";

                    if (bNotaCredito)
                    {
                        sNota = txtNbuscar.Text.Trim();
                        sNotaAño = cmbAñoNC.Text.Trim();
                    }
                    if (bNotaCreditoElectronica)
                    {
                        sNota = txtNEbuscar.Text.Trim();
                        sNotaAño = cmbAñoNE.Text.Trim();
                    }

                    //cadenaConexion = "Provider=vfpoledb.1;Data Source='" + cGlobal.sRutaBase + "';Collating Sequence=machine;";

                    conexion = new OleDbConnection(cadenaConexion);

                    conexion.Open();

                    consulta = @"select Numdocto, Numtipodoc, Codcteprov, Fecdocto, Seriedocto, Referdocto, Ordencompr,
                        Imptotaldo, Monedadoct, Diaesp, Aniodocto, Impnetodoc, Impivadoct, Perdocto, Fecpedido 
                        from MGP10008.DBF where LTRIM(RTRIM(Numtipodoc))='7' 
                        and LTRIM(RTRIM(Numdocto))='" + sNota + "' and LTRIM(RTRIM(Aniodocto))='" + sNotaAño + "';";

                    comando = new OleDbCommand(consulta, conexion);

                    leer = comando.ExecuteReader();

                    if (leer.HasRows)
                    { //if (1)
                        while (leer.Read())
                        { //while (1)

                            cliente.sDocumento = leer.GetString(0).Trim(); //numdocto
                            cliente.sNumtipodocto = leer.GetString(1).Trim(); //Numtipodoc
                            cliente.sCodigo = leer.GetString(2).Trim(); //codcteprov
                            cliente.dtFecha = Convert.ToDateTime(leer.GetValue(3)); //Fecdocto
                            cliente.sFecha = Convert.ToString(cliente.dtFecha.Day + "/" + cliente.dtFecha.Month + "/" + cliente.dtFecha.Year);
                            cliente.sSerie = leer.GetString(4).Trim(); //seriedocto
                            cliente.sFechaSQL = Convert.ToString(cliente.dtFecha.Year + "-" + cliente.dtFecha.Month + "-" + cliente.dtFecha.Day);
                            cliente.sReferencia = leer.GetString(5).Trim(); //referdocto
                            cliente.sOrdenCompra = leer.GetString(6).Trim(); //ordencompr
                            cliente.dTotal = Convert.ToDouble(leer.GetValue(7)); //imptotaldo
                            cliente.sAniodocto = leer.GetString(10).Trim(); //aniodocto
                            cliente.sPerdocto = leer.GetString(13).Trim(); //perdocto
                            articulo.dtFechaPedimento = Convert.ToDateTime(leer.GetValue(14)); //Fecpedido
                            articulo.sFecha = Convert.ToString(articulo.dtFechaPedimento.Day + "/" +
                                            articulo.dtFechaPedimento.Month + "/" + articulo.dtFechaPedimento.Year);
                            articulo.sFechaSql = Convert.ToString(articulo.dtFechaPedimento.Year + "-" +
                                            articulo.dtFechaPedimento.Month + "-" + articulo.dtFechaPedimento.Day);

                            if (leer.GetString(8).Trim() == "1")
                            {
                                sMoneda = "PESOS";
                                sMoneda_ini = "MXN";
                            }
                            else
                            {
                                sMoneda = "DOLARES";
                                sMoneda_ini = "USD";
                            }

                            cliente.sMoneda = sMoneda;
                            cliente.sMonedaIni = sMoneda_ini;

                            if (leer.GetString(9).Trim() == "1")
                            {
                                cliente.sFormaPago = "CONTADO";
                            }
                            else
                            {
                                cliente.sFormaPago = "CRÉDITO";
                            }

                            if (leer.GetString(2).Trim() == "C0088") //MGP10008.Codcteprov
                            { //if (2)
                                iTipoFactura = 2;

                                consulta2 = @"select Aniodocto, Perdocto, Numtipodoc, Seriedocto, Numdocto, Tipo, Rfc, 
                                    Razsocial, Domicilio, Codpostal, Coloniacte, Estadocte, Poblacte from MGP10011.DBF
                                    where LTRIM(RTRIM(Aniodocto))='" + cliente.sAniodocto +
                                "' and LTRIM(RTRIM(Perdocto))='" + cliente.sPerdocto +
                                "' and LTRIM(RTRIM(Numtipodoc))='" + cliente.sNumtipodocto +
                                "' and LTRIM(RTRIM(Seriedocto))='" + cliente.sSerie +
                                "' and LTRIM(RTRIM(Numdocto))='" + cliente.sDocumento +
                                "' and LTRIM(RTRIM(Tipo))='F';";

                                comando2 = new OleDbCommand(consulta2, conexion);

                                leer2 = comando2.ExecuteReader();

                                if (leer2.HasRows)
                                {
                                    while (leer2.Read())
                                    {
                                        iTipoFactura = 0;

                                        cliente.sRfc = leer2.GetString(6).Trim();

                                        if (cliente.sRfc == "")
                                        {
                                            iTipoFactura = 1;
                                        }
                                        cliente.sCliente = leer2.GetString(7).Trim(); //MGP10011.Razsocial

                                        if (leer2.GetString(7).Trim() == "CLIENTE CONTADO")
                                        {
                                            iTipoFactura = 2;
                                        }

                                        cliente.sDireccion = leer2.GetString(8).Trim(); //MGP10011.direccion
                                        cliente.sCP = leer2.GetString(9).Trim(); //MGP10011.codpostal
                                        cliente.sColonia = leer2.GetString(10).Trim(); // MGP10011.coloniacte
                                        cliente.sEstaDocte = leer2.GetString(11).Trim(); //MGP10011.estadocte
                                        cliente.sPoblacte = leer2.GetString(12).Trim(); //MGP10011.poblacte

                                        consulta3 = @"select Tipotabla, Estadocte, Descripcio, Numtabla from MGP10021.DBF
                                            where LTRIM(RTRIM(Tipotabla))='4' and 
                                            LTRIM(RTRIM(Numtabla))'" + cliente.sEstaDocte + "';";

                                        comando3 = new OleDbCommand(consulta3, conexion);

                                        leer3 = comando3.ExecuteReader();

                                        while (leer3.Read())
                                        {
                                            cliente.sEstado = leer3.GetString(2).Trim(); //MGP10021.Descripcio
                                        }

                                        consulta4 = @"select Tipotabla, Poblacte, Descripcio, Numtabla from MGP10021.DBF
                                            where LTRIM(RTRIM(Tipotabla))='7' and 
                                            LTRIM(RTRIM(Numtabla))'" + cliente.sPoblacte + "';";

                                        comando4 = new OleDbCommand(consulta4, conexion);

                                        leer4 = comando4.ExecuteReader();

                                        while (leer4.Read())
                                        {
                                            cliente.sPoblacion = leer4.GetString(2).Trim();
                                        }
                                    }
                                }
                            } //fin del if (2)

                            else
                            { // else del if (2)
                                consulta5 = @"select Codcteprov, Rfc, Razsocial, Domicilio, Codpostal, 
                                        Telefono1, Coloniacte, Estadocte, Poblacte from MGP10002.DBF 
                                        where LTRIM(RTRIM(Codcteprov))='" + cliente.sCodigo + "';";

                                comando5 = new OleDbCommand(consulta5, conexion);

                                leer5 = comando5.ExecuteReader();

                                while (leer5.Read())
                                {
                                    if (leer5.GetString(1).Trim() == "") //MGP10002.Rfc
                                    { iTipoFactura = 1; }

                                    cliente.sRfc = leer5.GetString(1).Trim(); //MGP10002.Rfc
                                    cliente.sCliente = leer5.GetString(2).Trim(); //MGP10002.Razsocial
                                    cliente.sDireccion = leer5.GetString(3).Trim(); // MGP10002.domicilio
                                    cliente.sCP = leer5.GetString(4).Trim(); // MGP10002.codpostal
                                    cliente.sTelefono = leer5.GetString(5).Trim(); //MGP10002.Telefono1
                                    cliente.sColonia = leer5.GetString(6).Trim(); //MGP10002.coloniacte
                                    cliente.sEstaDocte = leer5.GetString(7).Trim(); //MGP10002.estadocte
                                    cliente.sPoblacte = leer5.GetString(8).Trim(); //MGP10002.poblacte

                                    consulta6 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.DBF
                                            where LTRIM(RTRIM(Tipotabla))='4' and LTRIM(RTRIM(Numtabla))=
                                            '" + cliente.sEstaDocte + "';";

                                    comando6 = new OleDbCommand(consulta6, conexion);

                                    leer6 = comando6.ExecuteReader();

                                    while (leer6.Read())
                                    {
                                        cliente.sEstado = leer6.GetString(2).Trim(); //MGP10021.Descripcio
                                    }

                                    consulta7 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.DBF
                                            where LTRIM(RTRIM(Tipotabla))='7' and LTRIM(RTRIM(Numtabla))=
                                            '" + cliente.sPoblacte + "';";

                                    comando7 = new OleDbCommand(consulta7, conexion);

                                    leer7 = comando7.ExecuteReader();

                                    while (leer7.Read())
                                    {
                                        cliente.sPoblacion = leer7.GetString(2).Trim(); //MGP10021.Descripcio
                                    }

                                }

                            } //fin de else

                            //if (iTipoFactura == 1 || iTipoFactura == 2)
                            //{
                            //    cliente.sSubtotal = leer.GetDouble(7); //imptotaldo
                            //    cliente.dIva = 0.00;
                            //}
                            //else
                            //{
                                cliente.sSubtotal = Convert.ToDouble(leer.GetValue(11)); //impnetodoc
                                cliente.dIva = Convert.ToDouble(leer.GetValue(12)); //impivadoct
                            //}

                                cliente.dTotal = Convert.ToDouble(leer.GetValue(7)); //Imptotaldo

                            consulta8 = @"select Aniodocto, Perdocto, Numtipodoc, Seriedocto, Numdocto, Tipomov, Unidades,
                                        Nummovto, Porcdescau, Porcdesces, Porcivamov, Impnetomov, Impivamovt, Codprodser, 
                                        Refermovto, Precio from MGP10010.DBF 
                                        where LTRIM(RTRIM(Aniodocto))='" + cliente.sAniodocto +
                            "' and LTRIM(RTRIM(Perdocto))='" + cliente.sPerdocto +
                            "' and LTRIM(RTRIM(Numtipodoc))='" + cliente.sNumtipodocto +
                            "' and LTRIM(RTRIM(Seriedocto))='" + cliente.sSerie +
                            "' and LTRIM(RTRIM(Numdocto))='" + cliente.sDocumento +
                            "' and LTRIM(RTRIM(Tipomov))='N';";

                            comando8 = new OleDbCommand(consulta8, conexion);

                            leer8 = comando8.ExecuteReader();

                            while (leer8.Read())
                            {
                                articulo.dCantidad = Convert.ToDouble(leer8.GetValue(6)); //MGP10010.Unidades
                                articulo.sRefermovto = leer8.GetString(14).Trim(); //MGP10010.Refermovto
                                cliente.sCodprodser = leer8.GetString(13).Trim(); //MGP10010.Codprodser

                                if (iTipoFactura == 1 || iTipoFactura == 2)
                                {
                                    //articulo.dPrecioUnitario = ((Convert.ToDouble(leer8.GetValue(15))) //MGP10010.precio
                                    //            * (1 - ((Convert.ToDouble(leer8.GetValue(8)) / 100))) //MGP10010.porcdescau
                                    //            * (1 - ((Convert.ToDouble(leer8.GetValue(9)) / 100))) //MGP10010.porcdesces
                                    //            * (1 + ((Convert.ToDouble(leer8.GetValue(10)) / 100)))); //MGP10010.Porcivamov
                                    //articulo.dImporte = (Convert.ToDouble(leer8.GetValue(11)) //MGP10010.Impnetomov
                                    //                   + Convert.ToDouble(leer8.GetValue(12))); //MGP10010.Impivamovt

                                    articulo.dPrecioUnitario = ((Convert.ToDouble(leer8.GetValue(15))) //MGP10010.precio
                                                            * (1 - ((Convert.ToDouble(leer8.GetValue(8)) / 100))) //MGP10010.porcdescau
                                                            * (1 - ((Convert.ToDouble(leer8.GetValue(9)) / 100)))); //MGP10010.porcdesces

                                    articulo.dImporte = Convert.ToDouble(leer8.GetValue(11)); //MGP10010.impnetomov
                                }
                                else
                                {
                                    articulo.dPrecioUnitario = ((Convert.ToDouble(leer8.GetValue(15))) //MGP10010.precio
                                                            * (1 - ((Convert.ToDouble(leer8.GetValue(8)) / 100))) //MGP10010.porcdescau
                                                            * (1 - ((Convert.ToDouble(leer8.GetValue(9)) / 100)))); //MGP10010.porcdesces

                                    articulo.dImporte = Convert.ToDouble(leer8.GetValue(11)); //MGP10010.impnetomov

                                }

                                //+++++++++++++++++++++++++++ agregado para redondear decimales 23/04/2014
                                //articulo.dPrecioUnitario = Math.Round(articulo.dPrecioUnitario, 2);
                                //articulo.dImporte = Math.Round(articulo.dImporte, 2);
                                //+++++++++++++++++++++++++++

                                if (cGlobal.sNombreBase == cGlobal.sBaseMostrarExtra2)
                                {
                                    //consulta modificada para base 28
                                    consulta9 = @"select Codprodser, Descrippro, Extra2 from MGP10004.DBF
                                                where LTRIM(RTRIM(Codprodser))='" + cliente.sCodprodser + "';";
                                }
                                else
                                {
                                    consulta9 = @"select Codprodser, Descrippro from MGP10004.DBF
                                        where LTRIM(RTRIM(Codprodser))='" + cliente.sCodprodser + "';";
                                }

                                comando9 = new OleDbCommand(consulta9, conexion);

                                leer9 = comando9.ExecuteReader();

                                if (leer9.HasRows)
                                {
                                    while (leer9.Read())
                                    {
                                        if (cliente.sCodprodser == "DIFERENCIA")
                                        {
                                            articulo.sRefermovto = leer8.GetString(14).Trim(); //MGP10010.Refermovto

                                            if (cGlobal.sNombreBase == cGlobal.sBaseMostrarExtra2)
                                            {
                                                //mgp10004.Extra2, modificada para 28
                                                articulo.sDescripcion = leer9.GetString(2).Trim(); //MGP10004.Extra2
                                            }
                                            else
                                            {
                                                articulo.sDescripcion = leer9.GetString(1).Trim(); //MGP10004.Descripro
                                            }

                                            articulo.sCodigo = leer9.GetString(0).Trim(); //MGP10004.Codprodser
                                        }
                                        else
                                        {
                                            if (cGlobal.sNombreBase == cGlobal.sBaseMostrarExtra2)
                                            {
                                                //mgp10004.Extra2, modificada para 28
                                                articulo.sDescripcion = leer9.GetString(2).Trim(); //MGP10004.Extra2
                                            }
                                            else
                                            {
                                                articulo.sDescripcion = leer9.GetString(1).Trim(); //MGP10004.Descripro
                                            }
                                            //articulo.sDescripcion = leer9.GetString(1).Trim(); //MGP10004.Descripro
                                            articulo.sCodigo = leer9.GetString(0).Trim(); //MGP10004.Codprodser
                                        }
                                    }
                                }

                                else
                                {
                                    articulo.sDescripcion = "";
                                }

                                articulo.sPedimento = "";
                                articulo.sFecha = Convert.ToString(articulo.dtFechaPedimento.Day + "/" +
                                                  articulo.dtFechaPedimento.Month + "/" +
                                                  articulo.dtFechaPedimento.Year);

                                consulta10 = @"select Origen, Anumdocto, Aaniodocto, Aperdocto, Anumtipodo, Aseriedoct, 
                                        Anummovto from MGP10028.DBF
                                        where LTRIM(RTRIM(Anumdocto))='" + leer.GetString(0).Trim() + //MGP10008.Numdocto
                                "' and LTRIM(RTRIM(Origen))='L' and LTRIM(RTRIM(Aseriedoct))='" + leer.GetString(4).Trim() + //MGP10008.seriedocto
                                "' and LTRIM(RTRIM(Aperdocto))='" + leer.GetString(13).Trim() + //MGP10008.perdocto
                                "' and LTRIM(RTRIM(Anumtipodo))='" + leer.GetString(1).Trim() + //MGP10008.numtipodoc
                                "' and LTRIM(RTRIM(Aaniodocto))='" + leer.GetString(10).Trim() + //MGP10008.aniodocto
                                "';";

                                comando10 = new OleDbCommand(consulta10, conexion);

                                leer10 = comando10.ExecuteReader();

                                while (leer10.Read())
                                {
                                    cliente.sAAniodocto = leer10.GetString(2).Trim(); //MGP10028.Aaniodocto
                                    cliente.sAPerdocto = leer10.GetString(3).Trim(); //MGP10028.Aperdocto
                                    cliente.sANumtipodoc = leer10.GetString(4).Trim(); //MGP10028.Anumtipodo
                                    cliente.sASeriedocto = leer10.GetString(5).Trim(); //MGP10028.Aseriedoct
                                    cliente.sANumdocto = leer10.GetString(1).Trim(); //MGP10028.Anumdocto
                                    cliente.sANummovto = leer10.GetString(6).Trim(); //MGP10028.Anummovto

                                    consulta11 = @"select Entanio, Entper, Enttipodoc, Entserie, Entnumdoc, 
                                            Entnummov, Fecha, Numpedim from MGP10025.DBF
                                            where LTRIM(RTRIM(Entanio))='" + cliente.sAAniodocto +
                                    "' and LTRIM(RTIRM(Entper))='" + cliente.sAPerdocto +
                                    "' and LTRIM(RTRIM(Enttipodoc))='" + cliente.sANumtipodoc +
                                    "' and LTRIM(RTRIM(Entserie))='" + cliente.sASeriedocto +
                                    "' and LTRIM(RTRIM(Entnumdoc))='" + cliente.sANumdocto +
                                    "' and LTRIM(RTRIM(Entnummov))='" + cliente.sANummovto + "';";

                                    comando11 = new OleDbCommand(consulta11, conexion);

                                    leer11 = comando11.ExecuteReader();

                                    while (leer11.Read())
                                    {
                                        articulo.dtFechaPedimento = Convert.ToDateTime(leer11.GetValue(6)); //MGP10025.Fecha
                                        articulo.sFecha = Convert.ToString(articulo.dtFechaPedimento.Day + "/" +
                                                            articulo.dtFechaPedimento.Month + "/" + articulo.dtFechaPedimento.Year);
                                        articulo.sFechaSql = Convert.ToString(articulo.dtFechaPedimento.Year + "-" +
                                                            articulo.dtFechaPedimento.Month + "-" + articulo.dtFechaPedimento.Day);
                                        articulo.sPedimento = leer11.GetString(7).Trim(); //MGP10025.Numpedim
                                    }
                                }

                                llenarListArticulos();
                                nFila = nFila + 1;
                            }

                            vMov = vMov + 1;

                            //} //fin del while (0)
                        }//fin del while (1)
                        if (bNotaCredito)
                        {
                            mostrarDatosPantallaNotas();
                        }
                        else if (bNotaCreditoElectronica) { mostrarDatosPantallaNE(); }

                    }//fin de if (1)

                    else //else del if (1)
                    {
                        if (bNotaCredito) //pestaña nota de crédito
                        {
                            if (txtNbuscar.Text.Trim() == "")
                            {
                                limpiarPantallaNotas();
                                mensajesAdvertencia("FALTAN DATOS EN LA BUSQUEDA");
                                txtNbuscar.Focus();
                            }
                            else
                            {
                                limpiarPantallaNotas();
                                mensajesError("NOTA DE CREDITO INEXISTENTE");
                                txtNbuscar.Focus();
                            }
                        }
                        if (bNotaCreditoElectronica) // pestaña nota de crédito electrónica
                        {
                            if (txtNEbuscar.Text.Trim() == "")
                            {
                                limpiarPantallaNE();
                                mensajesAdvertencia("FALTAN DATOS EN LA BUSQUEDA");
                                txtNEbuscar.Focus();
                            }
                            else
                            {
                                limpiarPantallaNE();
                                mensajesError("NOTA DE CREDITO INEXISTENTE");
                                txtNEbuscar.Focus();
                            }
                        }
                    }
                    //}

                } //fin del try

                catch (Exception ex)
                {
                    MessageBox.Show("Error en la consulta, motivo: " + ex.Message);
                }

                //se cierra el datareader y se cierra la conexion con la base de datos
                if (!leer.IsClosed) { leer.Close(); }

                if (conexion.State == ConnectionState.Open) { conexion.Close(); }

            }

            private void guardarNotaCredito()
            {
                SqlCommand comandoNota;
                string sInsertarDatosNota;
                string sConsultaDatosNota;
                SqlDataReader sLeerNota;

                if (cliente.sCodigo == "C0000" || cliente.sCodigo == "D0000")
                    {
                        cliente.sSerie = txtNserie.Text;
                        cliente.sRfc = txtNrfc.Text.Trim();
                        cliente.sOrdenCompra = txtNorden.Text;
                        cliente.sPoblacion = txtNpoblacion.Text;
                        cliente.sReferencia = txtNreferencia.Text;
                        cliente.sCP = txtNcp.Text;
                        cliente.sEstado = txtNestado.Text;
                        cliente.sColonia = txtNcolonia.Text;
                        cliente.sTelefono = txtNtelefono.Text;
                        cliente.sPoblacion = txtNpoblacion.Text;
                        cliente.sDireccion = txtNdireccion.Text;
                        cliente.sCliente = txtNcliente.Text;
                    }

                SqlConnection conexionNotaCredito = new SqlConnection(cGlobal.sCadenaSql);
                

                if (cliente.sFormaPago == "CRÉDITO")
                {
                    cliente.sLeyenda = "VENTA A CRÉDITO";
                }
                else
                {
                    cliente.sLeyenda = "PAGO EN UNA SOLA EXHIBICIÓN";
                }

                string sConsultaSql = "truncate table tblArticulosNotaCredito";

                conexionNotaCredito.Open();

                sqlComando2 = new SqlCommand(sConsultaSql, conexionNotaCredito);

                sqlComando2.ExecuteNonQuery();

                conexionNotaCredito.Close();

                sInsertarDatosNota = @"Insert into tblNotaCredito (nombre, direccion, rfc, poblacion, fecha, nota, ncFactura, 
                                    formaPago, cp, codigo, colonia, telefono, leyenda, referencia, importeLetra, subtotal,
                                    total, iva, usuario, moneda)
                                    Values (@nombre, @direccion, @rfc, @poblacion, @fecha, @nota, @ncFactura, 
                                    @formaPago, @cp, @codigo, @colonia, @telefono, @leyenda, @referencia, @importeLetra, @subtotal,
                                    @total, @iva, @usuario, @moneda);";

                conexionNotaCredito.Open();

                comandoNota = new SqlCommand(sInsertarDatosNota, conexionNotaCredito);

                comandoNota.Parameters.AddWithValue("nombre", ((object)cliente.sCliente) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("direccion", ((object)cliente.sDireccion) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("rfc", ((object)cliente.sRfc) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("poblacion", ((object)cliente.sPoblacion) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("fecha", ((object)cliente.sFechaSQL) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("nota", ((object)cliente.sDocumento) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("ncFactura", ((object)cliente.sNcFactura) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("formaPago", ((object)cliente.sFormaPago) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("cp", ((object)cliente.sCP) ?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("codigo", ((object)cliente.sCodigo)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("colonia", ((object)cliente.sColonia)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("telefono", ((object)cliente.sTelefono)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("leyenda", ((object)cliente.sLeyenda)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("referencia", ((object)cliente.sReferencia)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("importeLetra", ((object)cliente.sImporteLetra)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("subtotal", ((object)cliente.sSubtotal)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("total", ((object)cliente.dTotal)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("iva", ((object)cliente.dIva)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("usuario", ((object)cGlobal.sUserOk)?? DBNull.Value);
                comandoNota.Parameters.AddWithValue("moneda", ((object)cliente.sMonedaIni) ?? DBNull.Value);

                comandoNota.ExecuteNonQuery();

                conexionNotaCredito.Close();

                sConsultaDatosNota = @"select id, nota from tblNotaCredito where 
                                    nota=@nota;";

                conexionNotaCredito.Open();

                comandoNota = new SqlCommand(sConsultaDatosNota, conexionNotaCredito);

                comandoNota.Parameters.AddWithValue("nota", ((object)cliente.sDocumento) ?? DBNull.Value);

                sLeerNota = comandoNota.ExecuteReader();

                while (sLeerNota.Read())
                {
                    cGlobal.iDatoLeido = Convert.ToInt32(sLeerNota.GetValue(0));
                }

                conexionNotaCredito.Close();

                int iObtenerRegistros = 0;
                int iObtenerArticulosFila = 0;
                int iObtenerArticulosColumna = 0;

                iObtenerRegistros = listNArticulos.Items.Count;
                iObtenerArticulosFila = listNArticulos.Items.Count;
                iObtenerArticulosColumna = listNArticulos.Columns.Count;

                arrayObtenerArticulos = new string[iObtenerArticulosFila, iObtenerArticulosColumna];

                for (i = 0; i < iObtenerRegistros; i++)
                {
                    listNArticulos.Items[i].Selected = true;

                    for (int c = 0; c < iObtenerArticulosColumna; c++)
                    {
                        arrayObtenerArticulos[i, c] = listNArticulos.SelectedItems[i].SubItems[c].Text;
                    }

                }

                for (int f = 0; f < iObtenerRegistros; f++)
                {
                    sInsertarDatosNota = @"Insert into tblArticulosNotaCredito (cantidad, idNotaCredito,
                                        codigo, descripcion, precioUnitario, importe) 
                                        Values (@cantidad, @idNotaCredito,
                                        @codigo, @descripcion, @precioUnitario, @importe);";                                           //importe


                    conexionNotaCredito.Open();

                    comandoNota = new SqlCommand(sInsertarDatosNota, conexionNotaCredito);

                    comandoNota.Parameters.AddWithValue("cantidad", ((object)arrayObtenerArticulos[f, 1]) ?? DBNull.Value);
                    comandoNota.Parameters.AddWithValue("idNotaCredito",((object)cGlobal.iDatoLeido)?? DBNull.Value);
                    comandoNota.Parameters.AddWithValue("codigo",((object)arrayObtenerArticulos[f, 3])?? DBNull.Value);
                    comandoNota.Parameters.AddWithValue("descripcion",((object)arrayObtenerArticulos[f, 4])?? DBNull.Value);
                    comandoNota.Parameters.AddWithValue("precioUnitario",((object)lPrecioUnitario[f])?? DBNull.Value);
                    comandoNota.Parameters.AddWithValue("importe", ((object)lImporte[f]) ?? DBNull.Value);

                    comandoNota.ExecuteNonQuery();

                    conexionNotaCredito.Close();
                }

            }

            private void btnGenerarNota_Click(object sender, EventArgs e)
            {
                bNotaCredito = true;
                estadoSistema(" En espera...");
                mensajesOk("GENERANDO NOTA, ESPERE POR FAVOR");
                guardarNotaCredito();
                cGlobal.sDocumento = cliente.sDocumento;
                frmReporteNotaCredito reporte = new frmReporteNotaCredito();
                reporte.Show();
                estadoSistema("");
                limpiarMensaje();
                grbEstadoSistemaNota.Update();
                bNotaCredito = false;
            }

            private void txtNbuscar_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                { btnNbuscar.PerformClick(); }
            }

            private void limpiarPantallaNotas()
            {
                //foreach (TabPage tab in tabControl1.TabPages)
                //{
                //    IEnumerable<TextBox> texts = tab.Controls.OfType<TextBox>();

                //    if (tab.Name == "tabPage3")
                //    {
                //        foreach (TextBox text in texts)
                //        {
                //            text.Text = "";
                //        }
                //    }
                //}

                txtNcliente.Text = "";
                txtNcodigo.Text = "";
                txtNcolonia.Text = "";
                txtNcp.Text = "";
                txtNdireccion.Text = "";
                txtNestado.Text = "";
                txtNfactura.Text = "";
                txtNfecha.Text = "";
                txtNformapago.Text = "";
                txtNimporte.Text = "";
                txtNiva.Text = "";
                txtNmoneda.Text = "";
                txtNorden.Text = "";
                txtNpoblacion.Text = "";
                txtNreferencia.Text = "";
                txtNrfc.Text = "";
                txtNserie.Text = "";
                txtNsubtotal.Text = "";
                txtNtelefono.Text = "";
                txtNtotal.Text = "";
                txtNmensajes.Text = "";
                txtNmensajes.BackColor = Color.Silver;
                listNArticulos.Items.Clear();
                grbEstadoSistemaNota.BackColor = Color.Transparent;
                tabControl1.Update();
            }

            private void mostrarDatosPantallaNotas()
            {
                if (cliente.sCodigo == "C0000" || cliente.sCodigo == "D0000")
                {
                    txtNserie.ReadOnly = false;
                    txtNrfc.ReadOnly = false;
                    txtNorden.ReadOnly = false;
                    txtNpoblacion.ReadOnly = false;
                    txtNreferencia.ReadOnly = false;
                    txtNcp.ReadOnly = false;
                    txtNestado.ReadOnly = false;
                    txtNcolonia.ReadOnly = false;
                    txtNtelefono.ReadOnly = false;
                    txtNpoblacion.ReadOnly = false;
                    txtNdireccion.ReadOnly = false;
                    txtNcliente.ReadOnly = false;
                }

                else
                {
                    txtNserie.ReadOnly = true;
                    txtNrfc.ReadOnly = true;
                    txtNorden.ReadOnly = true;
                    txtNpoblacion.ReadOnly = true;
                    txtNreferencia.ReadOnly = true;
                    txtNcp.ReadOnly = true;
                    txtNestado.ReadOnly = true;
                    txtNcolonia.ReadOnly = true;
                    txtNtelefono.ReadOnly = true;
                    txtNpoblacion.ReadOnly = true;
                    txtNdireccion.ReadOnly = true;
                    txtNcliente.ReadOnly = true;
                }

                string sTipoMoneda = "";

                if (cliente.sMoneda == "PESOS")
                { sTipoMoneda = "MXN"; }
                if (cliente.sMoneda == "DOLARES")
                { sTipoMoneda = "USD"; }
                if (cliente.sMoneda == "EUROS")
                { sTipoMoneda = "EUR"; }

                //se muestra el dato de la columna especificada en la etiqueta 
                txtNfactura.Text = cliente.sDocumento;
                txtNcodigo.Text = cliente.sCodigo;
                txtNserie.Text = cliente.sSerie;
                txtNrfc.Text = cliente.sRfc;
                txtNorden.Text = cliente.sOrdenCompra;
                txtNfecha.Text = cliente.sFecha;
                txtNpoblacion.Text = cliente.sPoblacion;
                txtNreferencia.Text = cliente.sReferencia;
                txtNcliente.Text = cliente.sCliente;
                txtNdireccion.Text = cliente.sColonia;
                txtNcp.Text = cliente.sCP;
                txtNtelefono.Text = cliente.sTelefono;
                txtNcolonia.Text = cliente.sColonia;
                txtNestado.Text = cliente.sEstado;
                txtNdireccion.Text = cliente.sDireccion;
                txtNmoneda.Text = cliente.sMoneda;
                txtNtotal.Text = Convert.ToString(cliente.dTotal) + " " + sTipoMoneda;
                txtNsubtotal.Text = Convert.ToString(cliente.sSubtotal) + " " + sTipoMoneda;
                txtNiva.Text = Convert.ToString(cliente.dIva) + " " + sTipoMoneda;
                txtNformapago.Text = cliente.sFormaPago;

                if (bDiferenciaIva)
                {
                    cliente.dIva = cliente.sSubtotal;
                    cliente.sSubtotal = 0.00;
                    txtNiva.Text = txtNsubtotal.Text;
                    txtNsubtotal.Text = 0.00 + " " + sTipoMoneda;
                }

                Numeros_letras convertir = new Numeros_letras();

                txtNimporte.Text = convertir.enletras(Convert.ToString(cliente.dTotal), sTipoMoneda, cliente.sMoneda);

                cliente.sImporteLetra = txtNimporte.Text;

                mensajesOk("NOTA DE CRÉDITO ENCONTRADA");

                btnGenerarNota.Enabled = true;
            }

        #endregion

        #region factura

            private void limpiarPantalla()
            {
                //foreach (TabPage tab in tabControl1.TabPages)
                //{
                //    IEnumerable<TextBox> texts = tab.Controls.OfType<TextBox>();

                //    if (tab.Text == "FACTURACIÓN")
                //    {
                //        foreach (TextBox text in texts)
                //        {
                //            text.Text = "";
                //        }
                //    }
                //}
                txtCliente.Text = "";
                txtCodigo.Text = "";
                txtColonia.Text = "";
                txtCP.Text = "";
                txtDireccion.Text = "";
                txtEstado.Text = "";
                txtFactura.Text = ""; 
                txtFecha.Text = "";
                txtFormaPago.Text = "";
                txtImporte.Text = "";
                txtIVA.Text = "";
                txtMoneda.Text = "";
                txtOrdenCompra.Text = "";
                txtPoblacion.Text = "";
                txtReferencia.Text = "";
                txtRFC.Text = "";
                txtSerie.Text = "";
                txtSubtotal.Text = "";
                txtTelefono.Text = "";
                txtTotal.Text = "";
                txtMensajes.Text = "";
                txtBancoF.Text = "";
                txtMensajes.Text = "";
                txtDetallesFactura.Text = "";
                rbMetodoPago[0].Checked = true;
                cmbIva.SelectedItem = "16";
                txtMensajes.BackColor = Color.Silver;
                listArticulos.Items.Clear();
                grbEstadoSistema.BackColor = Color.Transparent;
                rbdIvaRet.Checked = false;
                rdbIvaTrans.Checked = true;
                tabControl1.Update();

            }

            private void mostrarArticulos()
            {
                vMov = 1;
                //vMov2 = 1;
                nFila = 1;
                incremento = 0;

                try
                {

                    //consultaArticulos();
                    //mostrarMetodoPagoIva();

                    
                    consulta8 = @"select Aniodocto, Perdocto, Numtipodoc, Seriedocto, Numdocto, Tipomov, Precio,
                            Codprodser, Impnetomov, Refermovto, Unidades, Porcdescau, Porcdesces,
                            Porcivamov, Impivamovt, Nummovto from MGP10010.DBF 
                            where LTRIM(RTRIM(Aniodocto))='" + cliente.sAniodocto +
                            "' and LTRIM(RTRIM(Perdocto))='" + cliente.sPerdocto +
                            "' and LTRIM(RTRIM(Numtipodoc))='" + cliente.sNumtipodocto +
                            "' and LTRIM(RTRIM(Seriedocto))='" + cliente.sSerie +
                            "' and LTRIM(RTRIM(Numdocto))='" + cliente.sDocumento +
                            "' and LTRIM(RTRIM(Tipomov))='N';";

                    comando8 = new OleDbCommand(consulta8, conexion);

                    try
                    { //try (1)
                        leer8 = comando8.ExecuteReader();

                        if (leer8.HasRows) //si hay lectura de filas el metodo HasRows se convierte en True
                        { //if (5)
                            while (leer8.Read())
                            { //while (1)
                                if ((Convert.ToDouble(leer8.GetValue(6))) != 0)
                                { //if precio diferente de cero (2)

                                    articulo.sCodigo = leer8.GetString(7).Trim();//obtener el codigo para buscar los articulos relacionados

                                    if (articulo.sCodigo == "DESC")
                                    {
                                        articulo.dImporte = Convert.ToDouble(leer8.GetValue(8)); //Impnetomov
                                        articulo.sDescripcion = leer8.GetString(9).Trim(); //Refermovto
                                    }
                                    else
                                    { // else (4) del if anterior
                                        articulo.dCantidad = Convert.ToDouble(leer8.GetValue(10)); //unidades equivale a cantidad

                                        if (iTipoFactura == 1 || iTipoFactura == 2)
                                        {
                                            //articulo.dPrecioUnitario = (Convert.ToDouble(leer8.GetValue(6))) *
                                            //                       (1 - ((Convert.ToDouble(leer8.GetValue(11))) / 100)) * //Porcdescau
                                            //                       (1 - ((Convert.ToDouble(leer8.GetValue(12))) / 100)) * //Porcdesces
                                            //                       (1 + ((Convert.ToDouble(leer8.GetValue(13))) / 100));  //Porcivamov
                                            //articulo.dImporte = (Convert.ToDouble(leer8.GetValue(8))) + (Convert.ToDouble(leer8.GetValue(14))); //Impivamovt
                                            articulo.dPrecioUnitario = (Convert.ToDouble(leer8.GetValue(6))) *
                                                                   (1 - ((Convert.ToDouble(leer8.GetValue(11))) / 100)) * //Porcdescau
                                                                   (1 - ((Convert.ToDouble(leer8.GetValue(12))) / 100));  //Porcdesces
                                            articulo.dImporte = Convert.ToDouble(leer8.GetValue(8)); //Impnetomov;
                                        }
                                        else
                                        {
                                            articulo.dPrecioUnitario = (Convert.ToDouble(leer8.GetValue(6))) *
                                                                   (1 - ((Convert.ToDouble(leer8.GetValue(11))) / 100)) * //Porcdescau
                                                                   (1 - ((Convert.ToDouble(leer8.GetValue(12))) / 100));  //Porcdesces
                                            articulo.dImporte = Convert.ToDouble(leer8.GetValue(8)); //Impnetomov;
                                        }

                                        //+++++++++++++++++++++++++++ agregado para redondear decimales 23/04/2014
                                        //articulo.dPrecioUnitario = Math.Round(articulo.dPrecioUnitario, 2);
                                        //articulo.dImporte = Math.Round(articulo.dImporte, 2);
                                        //+++++++++++++++++++++++++++
                                        articulo.sDescripcion = leer8.GetString(9).Trim();

                                        if (articulo.sCodigo.StartsWith("ARM") && articulo.sDescripcion != String.Empty)
                                        {
                                            articulo.sDescripcion = leer8.GetString(9).Trim(); //Refermovto;    
                                        }
                                        else
                                        { //inicio del else(3) del if anterior

                                            if (cGlobal.sNombreBase == cGlobal.sBaseMostrarExtra2)
                                            {
                                                //consulta modificada para base 28, campo añadido Extra2
                                                consulta9 = @"select Codprodser, Descrippro, Extra2 from MGP10004.DBF 
                                                             where LTRIM(RTRIM(Codprodser))='" + articulo.sCodigo + "';";
                                            }
                                            else
                                            {
                                                //consulta para obtener el nombre del articulo, el cual esta en el campo "Descrippro"
                                                consulta9 = @"select Codprodser, Descrippro from MGP10004.DBF 
                                                    where LTRIM(RTRIM(Codprodser))='" + articulo.sCodigo + "';";
                                            }


                                            comando9 = new OleDbCommand(consulta9, conexion);

                                            try
                                            {
                                                leer9 = comando9.ExecuteReader();

                                                while (leer9.Read())
                                                {
                                                    if (articulo.sCodigo == leer9.GetString(0).Trim())
                                                    { 
                                                        
                                                        if (cGlobal.sNombreBase == cGlobal.sBaseMostrarExtra2)
                                                        {
                                                            //Extra2, nombre del articulo, modificada para 28
                                                         articulo.sDescripcion = leer9.GetString(2).Trim();  
                                                        }
                                                        else
                                                        {
                                                            articulo.sDescripcion = leer9.GetString(1).Trim();
                                                        } //Descripro, nombre del articulo
                                                    }
                                                    
                                                    else
                                                    {
                                                        articulo.sDescripcion = "";
                                                    }
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                MessageBox.Show("Error en la consulta Mega 0: " + ex.Message);
                                            }
                                        } // fin del else (3)

                                        consulta10 = @"select Origen, Aaniodocto, Aperdocto, Anumtipodo, Aseriedoct,
                                                Anumdocto, Anummovto from MGP10028.DBF 
                                                where LTRIM(RTRIM(Aaniodocto))='" + cliente.sAniodocto +
                                        "' and LTRIM(RTRIM(Aperdocto))='" + cliente.sPerdocto +
                                        "' and LTRIM(RTRIM(Anumtipodo))='" + cliente.sNumtipodocto +
                                        "' and LTRIM(RTRIM(Aseriedoct))='" + cliente.sSerie +
                                        "' and LTRIM(RTRIM(Anumdocto))='" + cliente.sDocumento + "';";

                                        comando10 = new OleDbCommand(consulta10, conexion);

                                        try
                                        {// try (2)
                                            leer10 = comando10.ExecuteReader();

                                            if (leer10.HasRows)
                                            {
                                                while (leer10.Read())
                                                {

                                                    cliente.sAAniodocto = leer10.GetString(1).Trim();
                                                    cliente.sAPerdocto = leer10.GetString(2).Trim();
                                                    cliente.sANumtipodoc = leer10.GetString(3).Trim();
                                                    cliente.sASeriedocto = leer10.GetString(4).Trim();
                                                    cliente.sANumdocto = leer10.GetString(5).Trim();
                                                    cliente.sANummovto = leer10.GetString(6).Trim();

                                                    consulta11 = @"select Entserie, Entanio, Entnumdoc, Entnummov,
                                                            Entper, Fecha, Numpedim, Enttipodoc from MGP10025.DBF 
                                                            where LTRIM(RTRIM(Entserie))='" + cliente.sASeriedocto +
                                                    "' and LTRIM(RTRIM(Entanio))='" + cliente.sAAniodocto +
                                                    "' and LTRIM(RTRIM(Entnumdoc))= '" + cliente.sANumdocto +
                                                    "' and LTRIM(RTRIM(Entnummov))= '" + cliente.sANummovto +
                                                    "' and LTRIM(RTRIM(Entper))= '" + cliente.sAPerdocto +
                                                    "' and LTRIM(RTRIM(Enttipodoc))='" + cliente.sANumtipodoc + "';";

                                                    comando11 = new OleDbCommand(consulta11, conexion);

                                                    leer11 = comando11.ExecuteReader();

                                                    if (leer11.HasRows)
                                                    {
                                                        while (leer11.Read())
                                                        {
                                                            articulo.sFecha = Convert.ToString(leer11.GetValue(5)); //Fecha
                                                            articulo.sNumPedido = leer11.GetString(6); //Numpedim
                                                        }
                                                    }
                                                }
                                            }
                                        } //fin del try (2)

                                        catch (Exception ex)
                                        {
                                            MessageBox.Show("Error en la consulta Mega 1: " + ex.Message);
                                        }

                                        llenarListArticulos();

                                    } //fin del else (4)

                                    //vMov = vMov + 1;
                                    nFila = nFila + 1;

                                } //end if precio diferente de cero (2)

                            } //fin del While (1)

                        } //fin del if (5)

                        else //else del if (5), el cual indica que la factura no fue encontrada
                        {
                            mensajesError("FACTURA INEXISTENTE");
                            ActiveControl = txtBuscar;
                        }

                    } //fin de try (1)
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error en la consulta de Mega 2: " + ex.Message);
                    }

                } //fin del try auxiliar
                catch (Exception ex)
                {
                    MessageBox.Show("Error en la consulta de Mega 3: " + ex.Message);
                }
            }

            private void mostrarDatosPantalla()
            {

                if (cliente.sCodigo == "C0000" || cliente.sCodigo == "D0000")
                {
                    txtSerie.ReadOnly = false;
                    txtRFC.ReadOnly = false;
                    txtOrdenCompra.ReadOnly = false;
                    txtPoblacion.ReadOnly = false;
                    txtReferencia.ReadOnly = false;
                    txtCP.ReadOnly = false;
                    txtEstado.ReadOnly = false;
                    txtColonia.ReadOnly = false;
                    txtTelefono.ReadOnly = false;
                    txtPoblacion.ReadOnly = false;
                    txtDireccion.ReadOnly = false;
                    txtCliente.ReadOnly = false;
                }

                else
                {
                    txtSerie.ReadOnly = true;
                    txtRFC.ReadOnly = true;
                    txtOrdenCompra.ReadOnly = true;
                    txtPoblacion.ReadOnly = true;
                    txtReferencia.ReadOnly = true;
                    txtCP.ReadOnly = true;
                    txtEstado.ReadOnly = true;
                    txtColonia.ReadOnly = true;
                    txtTelefono.ReadOnly = true;
                    txtPoblacion.ReadOnly = true;
                    txtDireccion.ReadOnly = true;
                    txtCliente.ReadOnly = true;
                }

                string sTipoMoneda = "";

                if (cliente.sMoneda == "PESOS")
                { sTipoMoneda = "MXN"; }
                if (cliente.sMoneda == "DOLARES")
                { sTipoMoneda = "USD"; }
                if (cliente.sMoneda == "EUROS")
                { sTipoMoneda = "EUR"; }

                //se muestra el dato de la columna especificada en la etiqueta 
                txtFactura.Text = cliente.sDocumento;
                txtCodigo.Text = cliente.sCodigo;
                txtSerie.Text = cliente.sSerie;
                txtRFC.Text = cliente.sRfc;
                txtOrdenCompra.Text = cliente.sOrdenCompra;
                txtFecha.Text = cliente.sFecha;
                txtPoblacion.Text = cliente.sPoblacion;
                txtReferencia.Text = cliente.sReferencia;
                txtCliente.Text = cliente.sCliente;
                txtDireccion.Text = cliente.sColonia;
                txtCP.Text = cliente.sCP;
                txtTelefono.Text = cliente.sTelefono;
                txtColonia.Text = cliente.sColonia;
                txtEstado.Text = cliente.sEstado;
                txtDireccion.Text = cliente.sDireccion;
                txtMoneda.Text = cliente.sMoneda;
                txtTotal.Text = Convert.ToString(cliente.dTotal) + " " + sTipoMoneda;
                txtSubtotal.Text = Convert.ToString(cliente.sSubtotal) + " " + sTipoMoneda;
                txtIVA.Text = Convert.ToString(cliente.dIva) + " " + sTipoMoneda;
                txtFormaPago.Text = cliente.sFormaPago;

                if (bDiferenciaIva)
                {
                    cliente.dIva = cliente.sSubtotal;
                    cliente.sSubtotal = 0.00;
                    txtIVA.Text = txtSubtotal.Text;
                    txtSubtotal.Text = 0.00 + " " + sTipoMoneda;
                }

                Numeros_letras convertir = new Numeros_letras();

                txtImporte.Text = convertir.enletras(Convert.ToString(cliente.dTotal), sTipoMoneda, cliente.sMoneda);
                cliente.sImporteLetra = txtImporte.Text;
                cliente.sTipoMoneda = sTipoMoneda;

                mensajesOk("FACTURA ENCONTRADA");

                btnFacturar.Enabled = true;

            }

            private void consultaArticulos()
            {
                string sConsultaSql;

                int a = 0;

                sConsultaSql = @"select idAuxiliar, unidadMedida from tblUnidadMedida where idAuxiliar in 
                                (select id from tblAuxiliar where numFactura='" + cliente.sDocumento + "');";

                sqlConexion = new SqlConnection(cGlobal.sCadenaSql);

                sqlConexion.Open();

                sqlComando = new SqlCommand(sConsultaSql, sqlConexion);

                SqlDataAdapter adaptador = new SqlDataAdapter(sConsultaSql, sqlConexion);

                DataTable tabla = new DataTable();

                adaptador.Fill(tabla);

                sqlLeer = sqlComando.ExecuteReader();

                sDatoLeido = new string[tabla.Rows.Count];

                while (sqlLeer.Read())
                {
                    sDatoLeido[a] = sqlLeer.GetString(1).Trim();
                    a = a + 1;
                }

            }

            private void btnBuscar_Click(object sender, EventArgs e)
            {
                bFacturaElectronica = true;
                estadoSistema(" Buscando....");
                limpiarPantalla(); 
                mensajesOk("BUSCANDO...");
                sTipoDocumento = "1";
                lPrecioUnitario = new List<double>();
                lImporte = new List<double>();
                buscar(txtBuscar.Text.Trim());
                estadoSistema("");
                bFacturaElectronica = false;
                
            }

            private void buscar(string sDocumento)
            {
                try
                {
                    cliente = new cCliente();
                    articulo = new cArticulos();

                    //cadenaConexion = "Provider=vfpoledb;Data Source='" + cGlobal.sRutaBase + "';Collating Sequence=machine;";

                    consulta = @"select Numdocto, Numtipodoc, Codcteprov, Fecdocto, Seriedocto, Referdocto, 
                    Ordencompr, Imptotaldo, Monedadoct, Diaesp, Impnetodoc, Impivadoct, Aniodocto, Perdocto from MGP10008.DBF
                    where LTRIM(RTRIM(Numdocto))='" + sDocumento + "' and LTRIM(RTRIM(Numtipodoc))='" + sTipoDocumento +
                    "' and LTRIM(RTRIM(Aniodocto))='" + cmbAño.Text + "';";

                    conexion = new OleDbConnection(cadenaConexion);

                    comando = new OleDbCommand(consulta, conexion);

                    conexion.Open();

                    leer = comando.ExecuteReader();

                    if (leer.HasRows)
                    { //if (3)
                        while (leer.Read())
                        { //While (3)

                            cliente.sCodigo = leer.GetString(2).Trim(); //Codtecprov

                            cliente.sSerie = leer.GetString(4).Trim(); //Seriedocto
                            cliente.sDocumento = leer.GetString(0).Trim(); //Numdocto
                            cliente.sOrdenCompra = leer.GetString(6).Trim(); //Ordencompr
                            cliente.sAniodocto = leer.GetString(12).Trim(); //Aniodocto
                            cliente.sPerdocto = leer.GetString(13).Trim(); //Perdocto
                            cliente.sNumtipodocto = leer.GetString(1).Trim(); //Numtipodoct
                            cliente.sOrdenCompra = leer.GetString(6).Trim(); //OrdenCompra
                            cliente.sReferencia = leer.GetString(5).Trim(); //Referdocto
                            cliente.dtFecha = Convert.ToDateTime(leer.GetValue(3)); //Fecdocto
                            cliente.sFecha = Convert.ToString(cliente.dtFecha.Day + "/" + cliente.dtFecha.Month + "/" + cliente.dtFecha.Year);
                            cliente.sFechaSQL = Convert.ToString(cliente.dtFecha.Year + "-" + cliente.dtFecha.Month + "-" + cliente.dtFecha.Day);
                            //dCantidad = leer.GetDouble(7); //Imptotaldo

                            if (leer.GetString(8).Trim() == "1")
                            {
                                sMoneda = "PESOS";
                                sMoneda_ini = "MXN";
                            }
                            else
                            {
                                sMoneda = "DOLARES";
                                sMoneda_ini = "USD";
                            }

                            cliente.sMoneda = sMoneda;

                            if (leer.GetString(9).Trim() == "1") //Diaesp
                            {
                                cliente.sFormaPago = "CONTADO";
                            }
                            else
                            {
                                cliente.sFormaPago = "CRÉDITO";
                            }

                            if (cliente.sCodigo == "C0000")
                            {

                                iTipoFactura = 2;

                                consulta2 = @"select Aniodocto, Perdocto, Numtipodoc, Tipo, Rfc, Razsocial, Domicilio, Codpostal, 
                                    Coloniacte, Estadocte, Poblacte, Numdocto from MGP10011.dbf 
                                where LTRIM(RTRIM(Perdocto))='" + cliente.sPerdocto +
                                "' and LTRIM(RTRIM(Aniodocto))='" + cliente.sAniodocto +
                                "' and LTRIM(RTRIM(Numtipodoc))='" + cliente.sNumtipodocto +
                                "' and LTRIM(RTRIM(Numdocto))='" + cliente.sDocumento +
                                "' and LTRIM(RTRIM(Tipo))='F';";

                                comando2 = new OleDbCommand(consulta2, conexion);

                                leer2 = comando2.ExecuteReader();

                                while (leer2.Read())
                                {
                                    iTipoFactura = 0;
                                    //string datos;
                                    //datos = leer2.GetString(1).Trim(); //Perdocto

                                    cliente.sRfc = leer2.GetString(4).Trim(); //Rfc

                                    if (cliente.sRfc == "")
                                    {
                                        iTipoFactura = 1;
                                    }

                                    cliente.sCliente = leer2.GetString(5).Trim(); //Razsocial

                                    if (cliente.sCliente == "Cliente contado")
                                    {
                                        iTipoFactura = 2;
                                    }

                                    cliente.sDireccion = leer2.GetString(6).Trim(); //Domicilio
                                    cliente.sCP = leer2.GetString(7).Trim(); //Codpostal
                                    cliente.sColonia = leer2.GetString(8).Trim(); //Coloniacte
                                    cliente.sEstaDocte = leer2.GetString(9).Trim(); //Estadocte
                                    cliente.sPoblacte = leer2.GetString(10).Trim(); //Poblacte
                                    cliente.sRfc = leer2.GetString(4).Trim(); //RFC

                                }

                                consulta3 = @"select Descripcio, Tipotabla, Numtabla from MGP10021.DBF
                                    where LTRIM(RTRIM(Tipotabla))='4' and LTRIM(RTRIM(Numtabla))='" + cliente.sEstaDocte + "'";

                                comando3 = new OleDbCommand(consulta3, conexion);

                                leer3 = comando3.ExecuteReader();

                                while (leer3.Read())
                                {
                                    cliente.sEstado = leer3.GetString(0).Trim(); //Descripcio
                                    
                                }

                                consulta4 = @"select Descripcio, Numtabla, Tipotabla from MGP10021.dbf 
                                    where LTRIM(RTRIM(Tipotabla))='7'
                                    and LTRIM(RTRIM(Numtabla))='" + cliente.sPoblacte + "';";

                                comando4 = new OleDbCommand(consulta4, conexion);

                                leer4 = comando4.ExecuteReader();

                                if (leer4.HasRows) //si hay lectura de filas el metodo HasRows se convierte en True
                                {
                                    while (leer4.Read())
                                    {
                                        cliente.sPoblacion = leer4.GetString(0).Trim(); //Descripcio

                                    }
                                }

                            }

                            else
                            { //else (03)

                                consulta5 = @"select Codcteprov, Rfc, Razsocial, Domicilio, Codpostal, 
                                    Telefono1, Coloniacte, Estadocte, Poblacte from MGP10002.dbf 
                                    where LTRIM(RTRIM(Codcteprov))='" + cliente.sCodigo + "'";

                                comando5 = new OleDbCommand(consulta5, conexion);

                                leer5 = comando5.ExecuteReader();

                                if (leer5.HasRows) //si hay lectura de filas el metodo HasRows se convierte en True
                                { //if (4)
                                    while (leer5.Read())
                                    { //while (4)
                                        if (leer5.GetString(1).Trim() == "") //Rfc
                                        {
                                            iTipoFactura = 1;
                                        }

                                        cliente.sRfc = leer5.GetString(1).Trim(); //Rfc
                                        cliente.sCliente = leer5.GetString(2).Trim(); //Razon social
                                        cliente.sDireccion = leer5.GetString(3).Trim(); //Domicilio
                                        cliente.sCP = leer5.GetString(4).Trim(); //Codpostal
                                        cliente.sTelefono = leer5.GetString(5).Trim(); //Telefono1
                                        cliente.sColonia = leer5.GetString(6).Trim(); //Coloniacte
                                        cliente.sEstaDocte = leer5.GetString(7).Trim(); //Estadocte
                                        cliente.sPoblacte = leer5.GetString(8).Trim(); //Poblacte

                                        consulta6 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.dbf 
                                                where LTRIM(RTRIM(Tipotabla))='4' and 
                                                LTRIM(RTRIM(Numtabla))='" + cliente.sEstaDocte + "'";

                                        comando6 = new OleDbCommand(consulta6, conexion);

                                        leer6 = comando6.ExecuteReader();

                                        if (leer6.HasRows) //si hay lectura de filas el metodo HasRows se convierte en True
                                        {
                                            while (leer6.Read())
                                            {
                                                cliente.sEstado = leer6.GetString(2).Trim(); //Descripcio
                                            }
                                        }

                                        consulta7 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.dbf 
                                                where LTRIM(RTRIM(Tipotabla))='7' and 
                                                LTRIM(RTRIM(Numtabla))='" + cliente.sPoblacte + "';";

                                        comando7 = new OleDbCommand(consulta7, conexion);

                                        leer7 = comando7.ExecuteReader();

                                        if (leer7.HasRows) //si hay lectura de filas el metodo HasRows se convierte en True
                                        {
                                            while (leer7.Read())
                                            {
                                                cliente.sPoblacion = leer7.GetString(2).Trim(); //Descripcio
                                            }
                                        }
                                    } //while (4)
                                } //fin del if (4)

                            } //fin del else (03)

                            //if (iTipoFactura == 1 || iTipoFactura == 2)
                            //{
                            //    cliente.sSubtotal = Convert.ToDouble(leer.GetValue(7)); //Imptotaldo
                            //    cliente.dIva = 0.00;
                            //}

                            //else
                            //{
                                
                                cliente.sSubtotal = Convert.ToDouble(leer.GetValue(10)); //Impnetodoc
                                cliente.dIva = Convert.ToDouble(leer.GetValue(11)); //Impivadoct
                            //}

                            cliente.dTotal = Convert.ToDouble(leer.GetValue(7)); //Imptotaldo

                            //se almacena en la variable el ultimo registro leido segun la consulta
                            cliente.sDocumento = leer.GetString(0).Trim(); //Numdocto

                            if (cliente.sCodigo == "TICKET")
                            { }
                            else { mostrarArticulos(); }

                            if (bFacturaElectronica) { mostrarDatosPantalla(); }
                            if (bImpresionDocumentos) { mostrarDatosPantallaImpresion(); }

                        } // While (3)
                    } //if (3)

                    else // else del if (3)
                    {
                        if (bFacturaElectronica)
                        {
                            limpiarPantalla();
                            mensajesError("NO EXISTE EL DOCUMENTO");
                            txtBuscar.Focus();
                        }
                        else if (bImpresionDocumentos)
                        {
                            limpiarPantallaImpresion();
                            mensajesError("NO EXISTE EL DOCUMENTO");
                            txtIbuscar.Focus();
                        }
                    }

                    if (bImpresionDocumentos)
                    {
                        if (txtIbuscar.Text.Trim() == "")
                        {
                            limpiarPantallaImpresion();
                            mensajesAdvertencia("FALTAN DATOS EN LA BUSQUEDA");
                        }
                    }

                    if (bFacturaElectronica)
                    {
                        if (txtBuscar.Text.Trim() == "")
                        {
                            limpiarPantalla();
                            mensajesAdvertencia("FALTAN DATOS EN LA BUSQUEDA");
                        }
                    }

                    //se cierra el datareader y se cierra la conexion con la base de datos
                    if (!leer.IsClosed) { leer.Close(); }

                    if (conexion.State == ConnectionState.Open) { conexion.Close(); }

                }// fin del try

                catch (Exception ex)
                {
                    MessageBox.Show("Error en la Consulta, motivo: " + ex.Message, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private void btnSalir_Click_1(object sender, EventArgs e)
            {
                if (conexion.State == ConnectionState.Open) { conexion.Close(); }
                Close();
            }

            private void btnFacturar_Click(object sender, EventArgs e)
            {
                TimeSpan ts = DateTime.Today.Date - cliente.dtFecha.Date;

                int iDiferenciaDias = ts.Days;

                if (iDiferenciaDias <= config.iDiasTimbrado)
                {
                    bFacturaElectronica = true;
                    estadoSistema("En espera...");
                    mensajesOk("EN ESPERA...");
                    facturar();
                    btnSalir.Enabled = true;
                    btnBuscar.Enabled = true;
                    estadoSistema("");
                    bFacturaElectronica = false;
                    ActiveControl = txtBuscar;
                }
                else
                {
                    MessageBox.Show("No se puede Timbrar la factura, debido a que la fecha esta fuera del limite permitido ", "Advertencia ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }

            private void obtenerUnidadMedida(string[] uMedida)
            {
                int iObtenerRegistros;

                listArticulos.FullRowSelect = true;

                iObtenerRegistros = listArticulos.Items.Count;

                //unidadMedida = new string[iObtenerRegistros];

                for (i = 0; i < iObtenerRegistros; i++)
                {
                    listArticulos.Items[i].Selected = true;
                }

                for (int j = 0; j < iObtenerRegistros; j++)
                {
                    uMedida[j] = listArticulos.SelectedItems[j].SubItems[3].Text;
                }
            }

            private void txtBuscar_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                { btnBuscar.PerformClick(); }
            }

            private void txtBancoF_KeyPress(object sender, KeyPressEventArgs e)
            {
                if (!Char.IsNumber(e.KeyChar) && e.KeyChar != (char)8)
                {
                    e.Handled = true;
                }

            }

            private void obtenerArticulos(string[,] arrayArticulos)
            {
                int iObtenerArticulos;

                listArticulos.FullRowSelect = true;

                iObtenerArticulos = listArticulos.Items.Count;

                for (i = 0; i < iObtenerArticulos; i++)
                {
                    listArticulos.Items[i].Selected = true;
                }

                for (int f = 0; f < iObtenerArticulos; f++)
                {
                    for (int c = 0; c < 6; c++)
                    {
                        arrayArticulos[f, c] = listArticulos.SelectedItems[f].SubItems[c].Text;
                    }
                }
            }

        #endregion

        #region impresionDocumentos

            private void cmbDocumento_SelectedValueChanged(object sender, EventArgs e)
            {
                if (cmbDocumento.SelectedItem == "Compra")
                {
                    cmbProveedor.Enabled = true;
                }
                else
                {
                    cmbProveedor.Enabled = false;
                }
            }

            private void limpiarPantallaImpresion()
            {

                //foreach (TabPage tab in tabControl1.TabPages)
                //{
                //    IEnumerable<TextBox> texts = tab.Controls.OfType<TextBox>();

                //    if (tab.Name == "tabPage2")
                //    {
                //        foreach (TextBox text in texts)
                //        {
                //            text.Text = "";
                //        }
                //    }
                //}

                txtIcliente.Text = "";
                txtIcodigo.Text = "";
                txtIcolonia.Text = "";
                txtIcp.Text = "";
                txtIdireccion.Text = "";
                txtIestado.Text = "";
                txtIfactura.Text = "";
                txtIfecha.Text = "";
                txtIformapago.Text = "";
                txtIimporte.Text = "";
                txtIiva.Text = "";
                txtImoneda.Text = "";
                txtIorden.Text = "";
                txtIpoblacion.Text = "";
                txtIreferencia.Text = "";
                txtIrfc.Text = "";
                txtIserie.Text = "";
                txtIsubtotal.Text = "";
                txtItelefono.Text = "";
                txtItotal.Text = "";
                txtImensajes.Text = "";
                txtImensajes.BackColor = Color.Silver;
                listArticulosImpresion.Items.Clear();
                grbEstadoSistemaImpresion.BackColor = Color.Transparent;
                tabControl1.Update();

            }

            private void mostrarDatosPantallaImpresion()
            {

                if (cliente.sCodigo == "C0000" || cliente.sCodigo == "D0000")
                {
                    txtIserie.ReadOnly = false;
                    txtIrfc.ReadOnly = false;
                    txtIorden.ReadOnly = false;
                    txtIpoblacion.ReadOnly = false;
                    txtIreferencia.ReadOnly = false;
                    txtIcp.ReadOnly = false;
                    txtIestado.ReadOnly = false;
                    txtIcolonia.ReadOnly = false;
                    txtItelefono.ReadOnly = false;
                    txtIpoblacion.ReadOnly = false;
                    txtIdireccion.ReadOnly = false;
                    txtIcliente.ReadOnly = false;
                }

                else
                {
                    txtIserie.ReadOnly = true;
                    txtIrfc.ReadOnly = true;
                    txtIorden.ReadOnly = true;
                    txtIpoblacion.ReadOnly = true;
                    txtIreferencia.ReadOnly = true;
                    txtIcp.ReadOnly = true;
                    txtIestado.ReadOnly = true;
                    txtIcolonia.ReadOnly = true;
                    txtItelefono.ReadOnly = true;
                    txtIpoblacion.ReadOnly = true;
                    txtIdireccion.ReadOnly = true;
                    txtIcliente.ReadOnly = true;
                }

                string sTipoMoneda = "";

                if (cliente.sMoneda == "PESOS")
                { sTipoMoneda = "MXN"; }
                else if (cliente.sMoneda == "DOLARES")
                { sTipoMoneda = "USD"; }
                else if (cliente.sMoneda == "EUROS")
                { sTipoMoneda = "EUR"; }

                //se muestra el dato de la columna especificada en la etiqueta 
                txtIfactura.Text = cliente.sDocumento;
                txtIcodigo.Text = cliente.sCodigo;
                txtIserie.Text = cliente.sSerie;
                txtIrfc.Text = cliente.sRfc;
                txtIorden.Text = cliente.sOrdenCompra;
                txtIfecha.Text = cliente.sFecha;
                txtIpoblacion.Text = cliente.sPoblacion;
                txtIreferencia.Text = cliente.sReferencia;
                txtIcliente.Text = cliente.sCliente;
                txtIdireccion.Text = cliente.sColonia;
                txtIcp.Text = cliente.sCP;
                txtItelefono.Text = cliente.sTelefono;
                txtIcolonia.Text = cliente.sColonia;
                txtIestado.Text = cliente.sEstado;
                txtIdireccion.Text = cliente.sDireccion;
                txtImoneda.Text = cliente.sMoneda;
                txtItotal.Text = Convert.ToString(cliente.dTotal) + " " + sTipoMoneda;
                txtIsubtotal.Text = Convert.ToString(cliente.sSubtotal) + " " + sTipoMoneda;
                txtIiva.Text = Convert.ToString(cliente.dIva) + " " + sTipoMoneda;
                txtIformapago.Text = cliente.sFormaPago;
                //txtImetodoPago.Text = cliente.sMetodoPago;
                //txtItipoIva.Text = cliente.sTipoIva;

                if (bDiferenciaIva)
                {
                    cliente.dIva = cliente.sSubtotal;
                    cliente.sSubtotal = 0.00;
                    txtIiva.Text = txtIsubtotal.Text;
                    txtIsubtotal.Text = 0.00 + " " + sTipoMoneda;
                }

                Numeros_letras convertir = new Numeros_letras();

                txtIimporte.Text = convertir.enletras(Convert.ToString(cliente.dTotal), sTipoMoneda, cliente.sMoneda);

                mensajesOk("DOCUMENTO ENCONTRADO");

                btnImprimir.Enabled = true;
            }

            private void btnIbuscar_Click(object sender, EventArgs e)
            {
                bImpresionDocumentos = true;

                if (cmbDocumento.Text != "")
                {
                    if (txtIbuscar.Text != "")
                    {
                        if (cmbDocumento.Text == "Factura")
                            sTipoDocumento = "1";
                        else if (cmbDocumento.Text == "Nota de cargo")
                            sTipoDocumento = "3";
                        else if (cmbDocumento.Text == "Pedido")
                            sTipoDocumento = "23";
                        else if (cmbDocumento.Text == "Remision")
                            sTipoDocumento = "28";
                        else if (cmbDocumento.Text == "Compra")
                            sTipoDocumento = "31";
                        else if (cmbDocumento.Text == "Devolucion proveedor")
                            sTipoDocumento = "37";
                        else { sTipoDocumento = ""; }

                        if (sTipoDocumento == "1") //tipo de documento: Factura
                        {
                            bImpresionDocumentos = true;
                            estadoSistema(" Buscando...");
                            limpiarPantallaImpresion();
                            mensajesOk("BUSCANDO...");
                            lPrecioUnitario = new List<double>();
                            lImporte = new List<double>();
                            buscarDocumentos(txtIbuscar.Text.Trim(), sTipoDocumento);
                            estadoSistema("");
                            bImpresionDocumentos = false;

                        }
                        else
                        {
                            if (txtIbuscar.Text != "" && sTipoDocumento != "")
                            {
                                bImpresionDocumentos = true;
                                estadoSistema(" Buscando...");
                                limpiarPantallaImpresion();
                                mensajesOk("BUSCANDO...");
                                lPrecioUnitario = new List<double>();
                                lImporte = new List<double>();
                                buscarDocumentos(txtIbuscar.Text.Trim(), sTipoDocumento);
                                estadoSistema("");
                                bImpresionDocumentos = false;
                            }
                            else
                            {
                                limpiarPantallaImpresion();
                                mensajesAdvertencia("FALTAN DATOS");
                            }
                        }
                    }
                    else
                    {
                        mensajesAdvertencia("FALTAN DATOS EN LA BUSQUEDA");
                        txtIbuscar.Focus();
                    }
                }
                else
                {
                    mensajesAdvertencia("SELECCIONE TIPO DE DOCUMENTO");
                    cmbDocumento.Focus();
                }

                bImpresionDocumentos = false;
            }

            private void guardarDatosDocumentos()
            {

                if (cliente.sCodigo == "C0000" || cliente.sCodigo == "D0000")
                {
                    cliente.sSerie = txtIserie.Text;
                    cliente.sRfc = txtIrfc.Text.Trim();
                    cliente.sOrdenCompra = txtIorden.Text;
                    cliente.sPoblacion = txtIpoblacion.Text;
                    cliente.sReferencia = txtIreferencia.Text;
                    cliente.sCP = txtIcp.Text;
                    cliente.sEstado = txtIestado.Text;
                    cliente.sColonia = txtIcolonia.Text;
                    cliente.sTelefono = txtItelefono.Text;
                    cliente.sPoblacion = txtIpoblacion.Text;
                    cliente.sDireccion = txtIdireccion.Text;
                    cliente.sCliente = txtIcliente.Text;
                }



                int iObtenerRegistros = 0;
                int iObtenerArticulosFila = 0;
                int iObtenerArticulosColumna = 0;
                string sConsultaSql;
                
                try
                {
                sqlConexion = new SqlConnection(cGlobal.sCadenaSql);

                if (listArticulosImpresion.Items.Count > 0)
                {
                    iObtenerRegistros = listArticulosImpresion.Items.Count;

                    iObtenerArticulosFila = listArticulosImpresion.Items.Count;
                    iObtenerArticulosColumna = listArticulosImpresion.Columns.Count;

                    arrayObtenerArticulos = new string[iObtenerArticulosFila, iObtenerArticulosColumna];

                    for (i = 0; i < iObtenerRegistros; i++)
                    {
                        listArticulosImpresion.Items[i].Selected = true;

                        for (int c = 0; c < iObtenerArticulosColumna; c++)
                        {
                            arrayObtenerArticulos[i, c] = listArticulosImpresion.SelectedItems[i].SubItems[c].Text;
                        }

                    }

                    sConsultaSql = "truncate table tblArticulosDocumentos";

                    sqlConexion.Open();

                    sqlComando2 = new SqlCommand(sConsultaSql, sqlConexion);

                    sqlComando2.ExecuteNonQuery();

                    sqlConexion.Close();

                    
                        string sInsertarDatos;

                        sMetodoPago = cmbMetodoPagoI.Text;
                        cliente.sCondicionPago = cmbCondicionPagoI.Text;
                        cliente.sTipoIva = cmbTipoIvaI.Text;
                        cliente.sMonedaIni = sMoneda_ini;

                        sInsertarDatos = @"Insert into tblAuxiliarDocumentos (metodoPago, documento, fecha, serie, referencia, 
                                ordenCompra, nombre, rfc, codigo, poblacion, cp, direccion, colonia, estado, telefono, moneda, formaPago, 
                                subtotal, iva, total, importeLetra, porcentajeIva, tipoDocumento, usuario, condicionPago, tipoIva, 
                                monedaIniciales, sTipoDocumento, fechaImpresion)
                                Values (@metodoPago, @documento, @fecha, @serie, @referencia, 
                                @ordenCompra, @nombre, @rfc, @codigo, @poblacion, @cp, @direccion, @colonia, @estado, @telefono, @moneda, @formaPago, 
                                @subtotal, @iva, @total, @importeLetra, @porcentajeIva, @tipoDocumento, @usuario, @condicionPago, @tipoIva, 
                                @monedaIniciales, @sTipoDocumento, @fechaImpresion);";

                        sqlConexion.Open();

                        sqlComando = new SqlCommand(sInsertarDatos, sqlConexion);


                        sqlComando.Parameters.AddWithValue("metodoPago", ((object)sMetodoPago) ?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("documento", ((object)cliente.sDocumento)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("fecha", ((object)cliente.sFechaSQL)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("serie", ((object)cliente.sSerie)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("referencia", ((object)cliente.sReferencia)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("ordenCompra", ((object)cliente.sOrdenCompra)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("nombre", ((object)cliente.sCliente)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("rfc", ((object)cliente.sRfc)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("codigo", ((object)cliente.sCodigo)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("poblacion", ((object)cliente.sPoblacion)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("cp", ((object)cliente.sCP)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("direccion", ((object)cliente.sDireccion)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("colonia", ((object)cliente.sColonia)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("estado", ((object)cliente.sEstado)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("telefono", ((object)cliente.sTelefono)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("moneda", ((object)cliente.sMoneda)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("formaPago", ((object)cliente.sFormaPago)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("subtotal", ((object)cliente.sSubtotal)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("iva", ((object)cliente.dIva)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("total", ((object)cliente.dTotal)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("importeLetra", ((object)txtIimporte.Text)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("porcentajeIva", ((object)cmbIva.Text)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("tipoDocumento", ((object)Convert.ToInt32(cliente.sNumtipodocto))?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("usuario", ((object)cGlobal.sUserOk)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("condicionPago", ((object)cliente.sCondicionPago)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("tipoIva", ((object)cliente.sTipoIva)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("monedaIniciales", ((object)cliente.sMonedaIni)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("sTipoDocumento", ((object)cmbDocumento.Text)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("fechaImpresion", DateTime.Now);

                        sqlComando.ExecuteNonQuery();

                        sqlConexion.Close();

                        sConsultaSql = @"select id from tblAuxiliarDocumentos where documento=@documento;";

                        sqlConexion.Open();

                        sqlComando2 = new SqlCommand(sConsultaSql, sqlConexion);

                        sqlComando2.Parameters.AddWithValue("documento", cliente.sDocumento);

                        sqlLeer = sqlComando2.ExecuteReader();

                        while (sqlLeer.Read())
                        {
                            cGlobal.iDatoLeido = Convert.ToInt32(sqlLeer.GetValue(0));
                        }

                        sqlConexion.Close();

                        for (int f = 0; f < iObtenerRegistros; f++)
                        {
                            sInsertarDatos = @"Insert into tblArticulosDocumentos (fechaPedimento, idDocumento, cantidad, 
                                    codigo, descripcion, precioUnitario, importe) 
                                    Values (@fechaPedimento, @idDocumento, @cantidad, 
                                    @codigo, @descripcion, @precioUnitario, @importe);";

                            //arrayObtenerArticulos[f,0] numero de linea
                            //arrayObtenerArticulos[f,1] cantidad
                            //arrayObtenerArticulos[f,2] codigo
                            //arrayObtenerArticulos[f,3] descripcion
                            //arrayObtenerArticulos[f,4] fecha de pedimento
                            //arrayObtenerArticulos[f,5] precio unitario
                            //arrayObtenerArticulos[f,6] importe

                            sqlConexion.Open();

                            sqlComando2 = new SqlCommand(sInsertarDatos, sqlConexion);

                            sqlComando2.Parameters.AddWithValue("fechaPedimento", ((object)articulo.sFechaSql)?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("idDocumento", ((object)cGlobal.iDatoLeido)?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("cantidad", ((object)Convert.ToDouble(arrayObtenerArticulos[f, 1]))?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("codigo", ((object)arrayObtenerArticulos[f, 2])?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("descripcion", ((object)arrayObtenerArticulos[f, 3])?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("precioUnitario", ((object)lPrecioUnitario[f])?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("importe", ((object)lImporte[f]) ?? DBNull.Value);

                            sqlComando2.ExecuteNonQuery();

                            sqlConexion.Close();
                        }

                        limpiarPantallaImpresion();
                        mensajesOk("SE GENERO EXITOSAMENTE EL DOCUMENTO");
                        btnFacturar.Enabled = false;
                    //}

                    if (cmbDocumento.Text == "Factura")
                    {
                        frmImpresionFactura reporte = new frmImpresionFactura();
                        reporte.Show();
                    }
                    else
                    {
                        frmCompras reporte = new frmCompras();
                        reporte.Show();
                    }
                 }
                }
                    catch(Exception ex)
                    {
                        mensajesError("ERROR EN LA CREACION DEL DOCUMENTO");
                    }
            }

            private void buscarDocumentos(string sDocumento, string sTipoDoc)
            {
                cliente = new cCliente();
                articulo = new cArticulos();

                string sConsultaDoc, sConsultaDoc2, sConsultaDoc3, sConsultaDoc4;
                OleDbCommand cmdDoc, cmdDoc2, cmdDoc3, cmdDoc4;
                OleDbDataReader leerDoc, leerDoc2, leerDoc3, leerDoc4;

                //string sTipoDoc2 = string.Format("{0,3}", sTipoDoc);
                //string sDocumento2 = string.Format("{0,12}",sDocumento);

                if (sTipoDoc != "31")
                {
                    sConsultaDoc = @"select Numdocto, Numtipodoc, Codcteprov, Fecdocto, Seriedocto, Referdocto, Ordencompr, 
                imptotaldo, Monedadoct, Diaesp, Aniodocto, Perdocto, Impnetodoc, Impivadoct, Fecpedido
                from MGP10008.DBF where LTRIM(RTRIM(Numdocto))='" + sDocumento +
                    "' and LTRIM(RTRIM(Numtipodoc))='" + sTipoDoc + "' and LTRIM(RTRIM(Aniodocto))='" + cmbAñoDoc.Text + "';";
                }
                else
                {
                    sConsultaDoc = @"select Numdocto, Numtipodoc, Codcteprov, Fecdocto, Seriedocto, Referdocto, Ordencompr, 
                    imptotaldo, Monedadoct, Diaesp, Aniodocto, Perdocto, Impnetodoc, Impivadoct, Fecpedido
                    from MGP10008.DBF where LTRIM(RTRIM(Numdocto))='" + sDocumento +
                    "' and LTRIM(RTRIM(Numtipodoc))='" + sTipoDoc + "' and LTRIM(RTRIM(Aniodocto))='" + cmbAñoDoc.Text +
                    "' and LTRIM(RTRIM(Codcteprov))like'%" + cmbProveedor.Text + "%';";
                }


//                sConsultaDoc = @"select Numdocto, Numtipodoc, Codcteprov, Fecdocto, Seriedocto, Referdocto, Ordencompr, 
//                imptotaldo, Monedadoct, Diaesp, Aniodocto, Perdocto, Impnetodoc, Impivadoct, Fecpedido
//                from MGP10008.DBF where Aniodocto='" + cmbAñoDoc.Text +
//                                              "' and Numtipodoc='" + sTipoDoc2 +
//                                              "' and Numdocto='" + sDocumento2 + "';";

                //cadenaConexion = "Provider=vfpoledb;Data Source='" + cGlobal.sRutaBase + "';Collating Sequence=machine;";

                conexion = new OleDbConnection(cadenaConexion);

                cmdDoc = new OleDbCommand(sConsultaDoc, conexion);

                conexion.Open();

                leerDoc = cmdDoc.ExecuteReader();

                if (leerDoc.HasRows)
                {
                    while (leerDoc.Read())
                    { //While (3)
                        cliente.sCodigo = leerDoc.GetString(2).Trim(); //Codteprov
                        cliente.dtFecha = Convert.ToDateTime(leerDoc.GetValue(3)); //Fecdocto
                        cliente.sFecha = Convert.ToString(cliente.dtFecha.Day + "/" + cliente.dtFecha.Month +
                                        "/" + cliente.dtFecha.Year);
                        cliente.sFechaSQL = Convert.ToString(cliente.dtFecha.Year + "-" + cliente.dtFecha.Month +
                                        "-" + cliente.dtFecha.Day);
                        cliente.sSerie = leerDoc.GetString(4).Trim(); //Seriedocto
                        cliente.sDocumento = leerDoc.GetString(0).Trim(); //Numdocto
                        cliente.sReferencia = leerDoc.GetString(5).Trim(); //Referdocto
                        cliente.sOrdenCompra = leerDoc.GetString(6).Trim(); //Ordencompr
                        cliente.dTotal = Convert.ToDouble(leerDoc.GetValue(7)); //imptotaldo
                        cliente.sAniodocto = leerDoc.GetString(10).Trim(); //Aniodocto
                        cliente.sPerdocto = leerDoc.GetString(11).Trim(); //Perdocto

                        articulo.dtFechaPedimento = Convert.ToDateTime(leerDoc.GetValue(14)); //fechaPedimento
                        articulo.sFecha = Convert.ToString(articulo.dtFechaPedimento.Day +
                            "/" + articulo.dtFechaPedimento.Month + "/" + articulo.dtFechaPedimento.Year);
                        articulo.sFechaSql = Convert.ToString(articulo.dtFechaPedimento.Year +
                            "-" + articulo.dtFechaPedimento.Month + "-" + articulo.dtFechaPedimento.Day);

                        if (leerDoc.GetString(8).Trim() == "1") //monedadoc
                        {
                            sMoneda = "PESOS";
                            sMoneda_ini = "MXN";
                        }
                        else
                        {
                            sMoneda = "DOLARES";
                            sMoneda_ini = "USD";
                        }
                        cliente.sMoneda = sMoneda;

                        if (leerDoc.GetString(9).Trim() == "1")
                        {
                            cliente.sFormaPago = "CONTADO";
                        }
                        else
                        {
                            cliente.sFormaPago = "CRÉDITO";
                        }

                        if (sTipoDoc == "31" || sTipoDoc == "33") // if (1)
                        {
                            iTipoFactura = 0;

                            //string sCodigo = string.Format("{0,20}", cliente.sCodigo);

                            sConsultaDoc2 = @"select Codcteprov, Rfcproved, Nombprovee, Direcprove, Codpostpro,
                            Telprov1, Coloniapro, Estadoprov, Poblaprove from MGP10013.DBF 
                            where LTRIM(RTRIM(Codcteprov))='" + cliente.sCodigo + "';";

                            cmdDoc2 = new OleDbCommand(sConsultaDoc2, conexion);

                            leerDoc2 = cmdDoc2.ExecuteReader();

                            while (leerDoc2.Read())
                            { //While (3)
                                cliente.sRfc = leerDoc2.GetString(1).Trim(); //Rfcproved
                                cliente.sCliente = leerDoc2.GetString(2).Trim(); //Nomprovee
                                cliente.sDireccion = leerDoc2.GetString(3).Trim(); //Direcprove
                                cliente.sCP = leerDoc2.GetString(4).Trim(); //Codpostro
                                cliente.sTelefono = leerDoc2.GetString(5).Trim(); //Telprov1
                                cliente.sColonia = leerDoc2.GetString(6).Trim(); //Coloniapro
                                cliente.sEstaDocte = leerDoc2.GetString(7).Trim(); //Estadoprov
                                cliente.sPoblacte = leerDoc2.GetString(8).Trim(); //Poblaprove

                                sConsultaDoc3 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.DBF where LTRIM(RTRIM(Tipotabla))='4'
                                and LTRIM(RTRIM(Numtabla))= '" + cliente.sEstaDocte + "';";

                                cmdDoc3 = new OleDbCommand(sConsultaDoc3, conexion);

                                leerDoc3 = cmdDoc3.ExecuteReader();

                                while (leerDoc3.Read())
                                {
                                    cliente.sEstado = leerDoc3.GetString(2).Trim(); //Descripcio
                                }


                                sConsultaDoc4 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.DBF where LTRIM(RTRIM(Tipotabla))='7'
                                and LTRIM(RTRIM(Numtabla))= '" + cliente.sPoblacte + "';";

                                cmdDoc4 = new OleDbCommand(sConsultaDoc4, conexion);

                                leerDoc4 = cmdDoc4.ExecuteReader();

                                while (leerDoc4.Read())
                                {
                                    cliente.sPoblacion = leerDoc4.GetString(2).Trim(); //Descripcio
                                }
                            }
                        }
                        else //else de if (1)
                        {
                            if (cliente.sCodigo == "C0000")
                            { // if (2)
                                iTipoFactura = 1;

                                //string serie = string.Format("{0,4}", cliente.sSerie);

                                sConsultaDoc2 = @"select Aniodocto, Perdocto, Numtipodoc, Seriedocto, Numdocto, Tipo, Rfc, Razsocial,
                                 Domicilio, Codpostal, Coloniacte, Estadocte, Poblacte
                                 from MGP10011.DBF where LTRIM(RTRIM(Aniodocto))= '" + cliente.sAniodocto +
                                "' and LTRIM(RTRIM(Perdocto))='" + cliente.sPerdocto + "' and LTRIM(RTRIM(Numtipodoc))='" + sTipoDoc +
                                "' and LTRIM(RTRIM(Seriedocto))='" + cliente.sSerie + "' and LTRIM(RTRIM(Numdocto))='" + cliente.sDocumento +
                                "' and LTRIM(RTRIM(Tipo))='F';";

                                cmdDoc2 = new OleDbCommand(sConsultaDoc2, conexion);

                                leerDoc2 = cmdDoc2.ExecuteReader();

                                while (leerDoc2.Read())
                                {
                                    iTipoFactura = 0;
                                    cliente.sRfc = leerDoc2.GetString(6).Trim();
                                    if (cliente.sRfc == "")
                                    {
                                        iTipoFactura = 1;
                                    }

                                    cliente.sCliente = leerDoc2.GetString(7).Trim();
                                    cliente.sDireccion = leerDoc2.GetString(8).Trim();
                                    cliente.sCP = leerDoc2.GetString(9).Trim();
                                    cliente.sColonia = leerDoc2.GetString(10).Trim();
                                    cliente.sEstaDocte = leerDoc2.GetString(11).Trim();
                                    cliente.sPoblacte = leerDoc2.GetString(12).Trim();

                                    sConsultaDoc3 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.DBF where LTRIM(RTRIM(Tipotabla))='4'
                                    and LTRIM(RTRIM(Numtabla))= '" + cliente.sEstaDocte + "';";

                                    cmdDoc3 = new OleDbCommand(sConsultaDoc3, conexion);

                                    leerDoc3 = cmdDoc3.ExecuteReader();

                                    while (leerDoc3.Read())
                                    {
                                        cliente.sEstado = leerDoc3.GetString(2).Trim(); //Descripcio
                                    }


                                    sConsultaDoc4 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.DBF where LTRIM(RTRIM(Tipotabla))='7'
                                    and LTRIM(RTRIM(Numtabla))= '" + cliente.sPoblacte + "';";

                                    cmdDoc4 = new OleDbCommand(sConsultaDoc4, conexion);

                                    leerDoc4 = cmdDoc4.ExecuteReader();

                                    while (leerDoc4.Read())
                                    {
                                        cliente.sPoblacion = leerDoc4.GetString(2).Trim(); //Descripcio
                                    }

                                }

                            } // if (2)
                            else
                            { // else (2)

                                //string sCodigo2 = string.Format("{0,20}", cliente.sCodigo);

                                sConsultaDoc2 = @"select Codcteprov, Rfc, Razsocial, Domicilio, Codpostal, Telefono1, Coloniacte,
                                Estadocte, Poblacte
                                from MGP10002.DBF where LTRIM(RTRIM(Codcteprov))='" + cliente.sCodigo + "'";

                                cmdDoc2 = new OleDbCommand(sConsultaDoc2, conexion);

                                leerDoc2 = cmdDoc2.ExecuteReader();

                                while (leerDoc2.Read())
                                {
                                    if (leerDoc2.GetString(1).Trim() == "")
                                        iTipoFactura = 1;

                                    cliente.sRfc = leerDoc2.GetString(1).Trim(); //Rfc
                                    cliente.sCliente = leerDoc2.GetString(2).Trim(); //Razsocial
                                    cliente.sDireccion = leerDoc2.GetString(3).Trim(); //Domicilio
                                    cliente.sCP = leerDoc2.GetString(4).Trim(); //Codpostal
                                    cliente.sTelefono = leerDoc2.GetString(5).Trim(); //Telefono1
                                    cliente.sColonia = leerDoc2.GetString(6).Trim(); //Coloniacte
                                    cliente.sEstaDocte = leerDoc2.GetString(7).Trim(); //Estadocte
                                    cliente.sPoblacte = leerDoc2.GetString(8).Trim(); //Poblacte

                                    sConsultaDoc3 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.DBF where LTRIM(RTRIM(Tipotabla))='4'
                                     and LTRIM(RTRIM(Numtabla))= '" + cliente.sEstaDocte + "';";

                                    cmdDoc3 = new OleDbCommand(sConsultaDoc3, conexion);

                                    leerDoc3 = cmdDoc3.ExecuteReader();

                                    while (leerDoc3.Read())
                                    {
                                        cliente.sEstado = leerDoc3.GetString(2).Trim();
                                    }


                                    sConsultaDoc4 = @"select Tipotabla, Numtabla, Descripcio from MGP10021.DBF where LTRIM(RTRIM(Tipotabla))='7'
                                    and LTRIM(RTRIM(Numtabla))= '" + cliente.sPoblacte + "';";

                                    cmdDoc4 = new OleDbCommand(sConsultaDoc4, conexion);

                                    leerDoc4 = cmdDoc4.ExecuteReader();

                                    while (leerDoc4.Read())
                                    {
                                        cliente.sPoblacion = leerDoc4.GetString(2).Trim();
                                    }

                                }

                            } //else (2)
                        }// else (1)
                        //if (iTipoFactura == 1 || iTipoFactura == 2)
                        //{
                        //    cliente.sSubtotal = Convert.ToDouble(leerDoc.GetValue(7)); //imptotaldo
                        //    cliente.dIva = 0.00;
                        //}
                        //else
                        //{
                            cliente.sSubtotal = Convert.ToDouble(leerDoc.GetValue(12)); //impnetodoc
                            cliente.dIva = Convert.ToDouble(leerDoc.GetValue(13)); //impivadoct
                        //}

                        cliente.dTotal = Convert.ToDouble(leerDoc.GetValue(7)); //imptotaldo

                        mostrarArticulosDocumentos();

                    }//fin del while (3)
                }
                else 
                {
                    mensajesOk("NO SE ENCONTRO EL DOCUMENTO");
                }
            }

            private void mostrarArticulosDocumentos()
            {
                nFila = 1;
                incremento = 0;

                try
                {
                    //string sDoc2 = string.Format("{0,12}", cliente.sDocumento);
                    //string sSerie2 = string.Format("{0,12}", cliente.sSerie);
                    //string sTipoDoc2 = string.Format("{0,3}", sTipoDocumento);
                    //string sPerdocto2 = string.Format("{0,2}", cliente.sPerdocto);

                    consulta8 = @"select Aniodocto, Perdocto, Numtipodoc, Seriedocto, Numdocto, Tipomov, Precio,
                        Codprodser, Impnetomov, Refermovto, Unidades, Porcdescau, Porcdesces, Porcivamov, Impivamovt, 
                        Nummovto, Fecdocto from MGP10010.DBF 
                        where LTRIM(RTRIM(Aniodocto))='" + cliente.sAniodocto +
                    "' and LTRIM(RTRIM(Perdocto))='" + cliente.sPerdocto +
                    "' and LTRIM(RTRIM(Numtipodoc))='" + sTipoDocumento +
                    "' and LTRIM(RTRIM(Seriedocto))='" + cliente.sSerie +
                    "' and LTRIM(RTRIM(Numdocto))='" + cliente.sDocumento +
                    "' and LTRIM(RTRIM(Tipomov))='N';";

                    comando8 = new OleDbCommand(consulta8, conexion);

                    leer8 = comando8.ExecuteReader();

                    if (leer8.HasRows) //si hay lectura de filas el metodo HasRows se convierte en True
                    { //if (5)
                        while (leer8.Read())
                        { //while (1)

                            articulo.dCantidad = Convert.ToDouble(leer8.GetValue(10)); //unidades equivale a cantidad
                            articulo.sCodigo = leer8.GetString(7).Trim();
                            //string sCodigo2 = leer8.GetString(7);

                            if (iTipoFactura == 1 || iTipoFactura == 2)
                            {
                                //articulo.dPrecioUnitario = (Convert.ToDouble(leer8.GetValue(6))) *
                                //                       (1 - ((Convert.ToDouble(leer8.GetValue(11))) / 100)) * //Porcdescau
                                //                       (1 - ((Convert.ToDouble(leer8.GetValue(12))) / 100)) * //Porcdesces
                                //                       (1 + ((Convert.ToDouble(leer8.GetValue(13))) / 100));  //Porcivamov
                                //articulo.dImporte = (Convert.ToDouble(leer8.GetValue(8))) + (Convert.ToDouble(leer8.GetValue(14))); //Impivamovt
                                articulo.dPrecioUnitario = (Convert.ToDouble(leer8.GetValue(6))) *
                                                       (1 - ((Convert.ToDouble(leer8.GetValue(11))) / 100)) * //Porcdescau
                                                       (1 - ((Convert.ToDouble(leer8.GetValue(12))) / 100));  //Porcdesces
                                articulo.dImporte = Convert.ToDouble(leer8.GetValue(8)); //Impnetomov
                            }
                            else
                            {
                                articulo.dPrecioUnitario = (Convert.ToDouble(leer8.GetValue(6))) *
                                                       (1 - ((Convert.ToDouble(leer8.GetValue(11))) / 100)) * //Porcdescau
                                                       (1 - ((Convert.ToDouble(leer8.GetValue(12))) / 100));  //Porcdesces
                                articulo.dImporte = Convert.ToDouble(leer8.GetValue(8)); //Impnetomov
                            }

                            //+++++++++++++++++++++++++++ agregado para redondear decimales 23/04/2014
                            //articulo.dPrecioUnitario = Math.Round(articulo.dPrecioUnitario, 2);
                            //articulo.dImporte = Math.Round(articulo.dImporte, 2);
                            //+++++++++++++++++++++++++++

                            articulo.dtFechaPedimento = Convert.ToDateTime(leer8.GetValue(16));
                            articulo.sFecha = Convert.ToString(articulo.dtFechaPedimento.Day + "/" +
                                articulo.dtFechaPedimento.Month + "/" + articulo.dtFechaPedimento.Year);
                            articulo.sFechaSql = Convert.ToString(articulo.dtFechaPedimento.Year +
                                "-" + articulo.dtFechaPedimento.Month + "-" + articulo.dtFechaPedimento.Day);

                             consulta9 = @"select Codprodser, Descrippro, Extra2 from MGP10004.DBF 
                                                where LTRIM(RTRIM(Codprodser))='" + articulo.sCodigo + "';";

                            comando9 = new OleDbCommand(consulta9, conexion);

                            leer9 = comando9.ExecuteReader();

                            while (leer9.Read())
                            {
                                if (articulo.sCodigo == leer9.GetString(0).Trim()) //Codprodser
                                { 
                                    if(cGlobal.sNombreBase == cGlobal.sBaseMostrarExtra2)
                                    {
                                        articulo.sDescripcion = leer9.GetString(2).Trim();
                                    }
                                    else
                                    {
                                        articulo.sDescripcion = leer9.GetString(1).Trim(); 
                                    }
                                    
                                } //Descripro, nombre del articulo
                                else
                                {
                                    articulo.sDescripcion = "";
                                }
                            }

                            if (articulo.sCodigo == "ARM")
                            {
                                articulo.sDescripcion = leer8.GetString(9).Trim(); //Refermovto;
                            }

                            consulta10 = @"select Origen, Aaniodocto, Aperdocto, Anumtipodo, Aseriedoct,
                                            Anumdocto, Anummovto from MGP10028.DBF 
                                            where LTRIM(RTRIM(Aaniodocto))='" + cliente.sAniodocto +
                            "' and LTRIM(RTRIM(Aperdocto))='" + cliente.sPerdocto +
                            "' and LTRIM(RTRIM(Anumtipodo))='" + sTipoDocumento +
                            "' and LTRIM(RTRIM(Aseriedoct))='" + cliente.sSerie +
                            "' and LTRIM(RTRIM(Anumdocto))='" + cliente.sDocumento + "';";
                            //"' and LTRIM(RTRIM(Anummovto))='" + Convert.ToString(vMov) + "';";

//                            consulta10 = @"select Origen, Aaniodocto, Aperdocto, Anumtipodo, Aseriedoct,
//                                            Anumdocto, Anummovto from MGP10028.DBF 
//                                            where Aaniodocto='" + cliente.sAniodocto +
//                                                            "' and Aperdocto='" + sPerdocto2 +
//                                                            "' and Anumtipodo='" + sTipoDoc2 +
//                                                            "' and Aseriedoct='" + sSerie2 +
//                                                            "' and Anumdocto='" + sDoc2 + "';";

                            comando10 = new OleDbCommand(consulta10, conexion);

                            leer10 = comando10.ExecuteReader();

                            if (leer10.HasRows)
                            {
                                while (leer10.Read())
                                {

                                    cliente.sAAniodocto = leer10.GetString(1).Trim();
                                    cliente.sAPerdocto = leer10.GetString(2).Trim();
                                    cliente.sANumtipodoc = leer10.GetString(3).Trim();
                                    cliente.sASeriedocto = leer10.GetString(4).Trim();
                                    cliente.sANumdocto = leer10.GetString(5).Trim();
                                    cliente.sANummovto = leer10.GetString(6).Trim();

                                    consulta11 = @"select Entserie, Entanio, Entnumdoc, Entnummov, 
                                                        Entper, Fecha, Numpedim, Enttipodoc from MGP10025.DBF
                                                        where LTRIM(RTRIM(Entserie))='" + cliente.sASeriedocto +
                                    "' and LTRIM(RTRIM(Entanio))='" + cliente.sAAniodocto +
                                    "' and LTRIM(RTRIM(Entnumdoc))= '" + cliente.sANumdocto +
                                    "' and LTRIM(RTRIM(Entnummov))= '" + cliente.sANummovto +
                                    "' and LTRIM(RTRIM(Entper))= '" + cliente.sAPerdocto +
                                    "' and LTRIM(RTRIM(Enttipodoc))='" + cliente.sANumtipodoc + "';";

//                                    consulta11 = @"select Entserie, Entanio, Entnumdoc, Entnummov, 
//                                                        Entper, Fecha, Numpedim, Enttipodoc from MGP10025.DBF
//                                                        where Entanio='" + cliente.sAAniodocto +
//                                    "' and Entper='" + cliente.sAPerdocto +
//                                    "' and Enttipodoc= '" + cliente.sANumtipodoc +
//                                    "' and Entserie= '" + cliente.sASeriedocto +
//                                    "' and Entnumdoc= '" + cliente.sANumdocto + "';";

                                    comando11 = new OleDbCommand(consulta11, conexion);

                                    leer11 = comando11.ExecuteReader();

                                    if (leer11.HasRows)
                                    {
                                        while (leer11.Read())
                                        {
                                            articulo.dtFechaPedimento = Convert.ToDateTime(leer11.GetValue(5));
                                            articulo.sFecha = Convert.ToString(articulo.dtFechaPedimento.Day + "/"
                                                + articulo.dtFechaPedimento.Month + "/" + articulo.dtFechaPedimento.Year);
                                            articulo.sFechaSql = Convert.ToString(articulo.dtFechaPedimento.Year + "-"
                                                + articulo.dtFechaPedimento.Month + "-" + articulo.dtFechaPedimento.Day);
                                            articulo.sNumPedido = leer11.GetString(6); //Numpedim
                                        }
                                    }
                                }
                            }

                            llenarListaArticulosDocumentos();

                            //vMov = vMov + 1;
                            nFila = nFila + 1;

                        } //fin del While (1)
                        mostrarDatosPantallaImpresion();
                    } //fin del if (5)

                    else //else del if (5), el cual indica que la factura no fue encontrada
                    {
                        txtMensajes.Visible = true;
                        txtMensajes.Text = "FACTURA INEXISTENTE";
                        txtMensajes.ForeColor = Color.White;
                        txtMensajes.BackColor = Color.Red;
                        txtBuscar.Focus();
                    }

                } //fin del try auxiliar
                catch (Exception ex)
                {
                    MessageBox.Show("Error en la consulta Documentos, motivo: " + ex.Message);
                }
            }

            private void llenarListaArticulosDocumentos()
            {
                bDiferenciaIva = false;

                ListViewItem articulo1 = new ListViewItem(Convert.ToString(nFila));

                listArticulosImpresion.AddSubItem = true;

                articulo1.SubItems.Add(Convert.ToString(articulo.dCantidad));
                articulo1.SubItems.Add(articulo.sCodigo);
                articulo1.SubItems.Add(articulo.sDescripcion);

                    if (articulo.sCodigo == cGlobal.sCodigoDiferenciaIva)
                    {
                        bDiferenciaIva = true;
                    }

                articulo1.SubItems.Add(Convert.ToString(articulo.sFecha));

                if (bDiferenciaIva)
                {
                    articulo1.SubItems.Add(0.00 + " " + sMoneda_ini);
                    articulo1.SubItems.Add(0.00 + " " + sMoneda_ini);

                    lPrecioUnitario.Add(0.00);
                    lImporte.Add(0.00);
                }
                else
                {
                    articulo1.SubItems.Add(Convert.ToString(articulo.dPrecioUnitario) + " " + sMoneda_ini);
                    articulo1.SubItems.Add(Convert.ToString(articulo.dImporte) + " " + sMoneda_ini);

                    lPrecioUnitario.Add(articulo.dPrecioUnitario);
                    lImporte.Add(articulo.dImporte);
                }

              listArticulosImpresion.Items.Add(articulo1);

            }

            private void txtIbuscar_KeyDown(object sender, KeyEventArgs e)
            {
                if (e.KeyCode == Keys.Enter)
                { btnIbuscar.PerformClick(); }
            }

            private void btnImprimir_Click(object sender, EventArgs e)
            {
                estadoSistema("En espera...");
                mensajesOk("IMPRIMIENDO...");
                cGlobal.sDocumento = cliente.sDocumento;
                guardarDatosDocumentos();
                estadoSistema("");
                limpiarMensaje();
            }

        #endregion

        #region cargaInicial

            private void btnBusca_Load(object sender, EventArgs e)
            {
                //la aplicacion utiliza la cultura indicada

                CultureInfo ci = new CultureInfo("es-MX");

                Application.CurrentCulture = ci;

                //ci.NumberFormat.NumberDecimalSeparator = ".";
                //ci.NumberFormat.NumberGroupSeparator = ".";
                //Application.CurrentCulture = ci;

                lblBase.Text = "BASE " + config.sBase;

                this.ActiveControl = txtBuscar;

                this.Text = "Grupo Hergosa Facturación e Impresión de documentos,  " + "Usuario: " + cGlobal.sUserOk;

                cadenaConexion = "Provider=vfpoledb;Data Source='" + config.sRutaBD + "'; Collating Sequence=machine; Mode=Read";

                //error la tabla no tiene el formato esperado
                //cadenaConexion = "Provider=Microsoft.Jet.OLEDB.4.0;Extended Properties='dBASE 5.0';Data Source='" + sBaseDatos + "';"; 

                //selecciona solo las facturas, las identifica mediante el Numtipodoc = 1
                //consulta = @"select LTRIM(RTRIM(Numdocto)), LTRIM(RTRIM(Numtipodoc)) from MGP10008.DBF where LTRIM(RTRIM(Numtipodoc))='1' and LTRIM(RTRIM(Aniodocto))='" + DateTime.Now.Year + "';";
                consulta = @"select Numdocto, Numtipodoc from MGP10008.DBF where LTRIM(RTRIM(Aniodocto))='" + DateTime.Now.Year + "' and LTRIM(RTRIM(Numtipodoc))='1';";

                //selecciona solo las Notas de crédito, las identifica mediante el Numtipodoc = 7
                //consulta2 = @"select LTRIM(RTRIM(Numdocto)), LTRIM(RTRIM(Numtipodoc)) from MGP10008.DBF where LTRIM(RTRIM(Numtipodoc))='7' and LTRIM(RTRIM(Aniodocto))='" + DateTime.Now.Year + "';";
                consulta2 = @"select Numdocto, Numtipodoc from MGP10008.DBF where LTRIM(RTRIM(Aniodocto))='" + DateTime.Now.Year + "' and LTRIM(RTRIM(Numtipodoc))='7';";

                conexion = new OleDbConnection(cadenaConexion);

                string consultaProveedor = @"select Distinct(Codcteprov) from  MGP10008.DBF where LTRIM(RTRIM(Numtipodoc))='31'";

                string sFactura;
                string sNota;

                comando = new OleDbCommand(consulta, conexion);
                comando2 = new OleDbCommand(consulta2, conexion);
                comandoProveedor = new OleDbCommand(consultaProveedor, conexion);

                coleccion = new AutoCompleteStringCollection();
                coleccionNotas = new AutoCompleteStringCollection();

                try
                {
                    conexion.Open();

                    leer = comando.ExecuteReader();
                    leer2 = comando2.ExecuteReader();
                    leerProveedor = comandoProveedor.ExecuteReader();
                    cmbProveedor.Items.Add(' ');

                    while (leer.Read())
                    {
                        sFactura = leer.GetString(0).Trim();
                        coleccion.Add(Convert.ToString(sFactura));
                    }

                    while (leer2.Read())
                    {
                        sNota = leer2.GetString(0).Trim();
                        coleccionNotas.Add(Convert.ToString(sNota));
                    }
                    while (leerProveedor.Read())
                    {
                        cmbProveedor.Items.Add(leerProveedor.GetString(0).Trim());
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Error en la consutla, motivo: \n" + ex.Message, "Error " ,MessageBoxButtons.OK,MessageBoxIcon.Error);
                }

                leer.Close();
                leer2.Close();
                leerProveedor.Close();
                conexion.Close();

                txtBuscar.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                txtBuscar.AutoCompleteSource = AutoCompleteSource.CustomSource;

                txtNEbuscar.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                txtNEbuscar.AutoCompleteSource = AutoCompleteSource.CustomSource;

                txtNbuscar.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
                txtNbuscar.AutoCompleteSource = AutoCompleteSource.CustomSource;

                txtBuscar.AutoCompleteCustomSource = coleccion;
                txtNEbuscar.AutoCompleteCustomSource = coleccionNotas;
                txtNbuscar.AutoCompleteCustomSource = coleccionNotas;

                txtBuscar.Focus();

                rdbIvaTrans.Checked = true;
                rdbIvaNETrans.Checked = true;

                leerArchivoConfig();

                //permisos
                //+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

                if (cGlobal.sPermisoOk == "2") //facturacion, impresion, generar pdf
                {
                    tabControl1.TabPages.Remove(tabPage3);
                    tabControl1.TabPages.Remove(tabPage4);
                    tabControl1.TabPages.Remove(tabPage5);
                }
                else if (cGlobal.sPermisoOk == "3") //nota credito, nota credito electronica, generar pdf
                {
                    tabControl1.TabPages.Remove(tabPage1);
                    tabControl1.TabPages.Remove(tabPage2);
                    tabControl1.TabPages.Remove(tabPage5);
                }
                else if (cGlobal.sPermisoOk == "4") //impresion, nota de credito
                {
                    tabControl1.TabPages.Remove(tabPage1);
                    tabControl1.TabPages.Remove(tabPage4);
                    tabControl1.TabPages.Remove(tabPage5);
                    tabControl1.TabPages.Remove(tabPage6);
                }
                else if (cGlobal.sPermisoOk == "5") //nota credito, nota credito electronica
                {
                    tabControl1.TabPages.Remove(tabPage1);
                    tabControl1.TabPages.Remove(tabPage2);
                    tabControl1.TabPages.Remove(tabPage5);
                    tabControl1.TabPages.Remove(tabPage6);
                }
                else if (cGlobal.sPermisoOk == "6") //reporte
                {
                    tabControl1.TabPages.Remove(tabPage1);
                    tabControl1.TabPages.Remove(tabPage2);
                    tabControl1.TabPages.Remove(tabPage3);
                    tabControl1.TabPages.Remove(tabPage4);
                    tabControl1.TabPages.Remove(tabPage6);
                }
                else if (cGlobal.sPermisoOk == "7") //facturacion, impresion
                {
                    tabControl1.TabPages.Remove(tabPage3);
                    tabControl1.TabPages.Remove(tabPage4);
                    tabControl1.TabPages.Remove(tabPage5);
                    tabControl1.TabPages.Remove(tabPage6);
                }
                else if (cGlobal.sPermisoOk == "8") //generar pdf
                {
                    tabControl1.TabPages.Remove(tabPage1);
                    tabControl1.TabPages.Remove(tabPage2);
                    tabControl1.TabPages.Remove(tabPage3);
                    tabControl1.TabPages.Remove(tabPage4);
                    tabControl1.TabPages.Remove(tabPage5);
                }
                else if (cGlobal.sPermisoOk == "9") //reporte, generar pdf
                {
                    tabControl1.TabPages.Remove(tabPage1);
                    tabControl1.TabPages.Remove(tabPage2);
                    tabControl1.TabPages.Remove(tabPage3);
                    tabControl1.TabPages.Remove(tabPage4);
                }
                else if (cGlobal.sPermisoOk == "10") //impresion de documentos, nota de credito, nota de credito electronica
                {
                    tabControl1.TabPages.Remove(tabPage1);
                    tabControl1.TabPages.Remove(tabPage5);
                    tabControl1.TabPages.Remove(tabPage6);
                }
            }

            private void leerArchivoConfig()
            {
                string path;
                //elementos = new ArrayList();
                path = Application.StartupPath;
                aUnidadMedida = new ArrayList();
                cmbTipoIvaI.Text = "IVA Transferido";

                try
                {
                    foreach (string dato in cGlobal.condicionesPago)
                    {
                        cmbCondicionesPago.Items.Add(dato);
                        cmbCondicionesPago.Text = dato;
                        cmbCondicionesPagoNE.Items.Add(dato);
                        cmbCondicionesPagoNE.Text = dato;
                        cmbCondicionPagoI.Items.Add(dato);
                        cmbCondicionPagoI.Text = dato;
                    }

                    foreach (string dato in cGlobal.años)
                    {
                        cmbAño.Items.Add(dato);
                        cmbAño.Text = dato;
                        cmbAñoNC.Items.Add(dato);
                        cmbAñoNC.Text = dato;
                        cmbAñoNE.Items.Add(dato);
                        cmbAñoNE.Text = dato;
                        cmbAñoDoc.Items.Add(dato);
                        cmbAñoDoc.Text = dato;
                    }

                    foreach (string dato in cGlobal.unidadMedida)
                    {
                        aUnidadMedida.Add(dato);
                    }
                   
                    int iPosicion, iPosicion2;
                    iPosicion = 15;
                    iPosicion2 = 15;

                    int i = 0;

                    rbMetodoPago = new RadioButton[cGlobal.metodosPago.Length];
                    rbMetodoPagoNE = new RadioButton[cGlobal.metodosPago.Length];

                    //rutina para generar metodos de pago adicionales agregados al archivo de configuracion inicial
                    
                    foreach (string smetodoPago in cGlobal.metodosPago)
                    {
                        if (iPosicion <= 115 || iPosicion2 <= 115)//|| iPosicion3 <= 115)
                        {
                            if (iPosicion <= 115)
                            {
                                rbMetodoPago[i] = new RadioButton();
                                rbMetodoPago[i].AutoSize = true;
                                rbMetodoPago[i].Text = smetodoPago;
                                rbMetodoPago[i].Name = "rb" + smetodoPago.Trim();
                                rbMetodoPago[i].Location = new Point(10, iPosicion);
                                gpbMetodoPago.Controls.Add(rbMetodoPago[i]);
                                rbMetodoPagoNE[i] = new RadioButton();
                                rbMetodoPagoNE[i].AutoSize = true;
                                rbMetodoPagoNE[i].Text = smetodoPago;
                                cmbMetodoPagoI.Items.Add(smetodoPago);
                                rbMetodoPagoNE[i].Name = "rb" + smetodoPago.Trim();
                                rbMetodoPagoNE[i].Location = new Point(10, iPosicion);
                                gpbNEMetodoPago.Controls.Add(rbMetodoPagoNE[i]);
                                iPosicion = iPosicion + 20;
                                i = i + 1;
                            }
                            else
                            {
                                rbMetodoPago[i] = new RadioButton();
                                rbMetodoPago[i].AutoSize = true;
                                rbMetodoPago[i].Text = smetodoPago;
                                rbMetodoPago[i].Name = "rb" + smetodoPago.Trim();
                                rbMetodoPago[i].Location = new Point(250, iPosicion2);
                                gpbMetodoPago.Controls.Add(rbMetodoPago[i]);
                                rbMetodoPagoNE[i] = new RadioButton();
                                rbMetodoPagoNE[i].AutoSize = true;
                                rbMetodoPagoNE[i].Text = smetodoPago;
                                cmbMetodoPagoI.Items.Add(smetodoPago);
                                rbMetodoPagoNE[i].Name = "rb" + smetodoPago.Trim();
                                rbMetodoPagoNE[i].Location = new Point(250, iPosicion2);
                                gpbNEMetodoPago.Controls.Add(rbMetodoPagoNE[i]);
                                iPosicion2 = iPosicion2 + 20;
                                i = i + 1;
                            }

                        }
                    }

                    rbMetodoPago[0].Checked = true;
                    rbMetodoPagoNE[1].Checked = true;

                    //foreach (string sLeerElemento in porcentajeIva)
                    foreach (string sPorcentajeIva in cGlobal.porcentajeIva)
                    {
                        cmbIva.Items.Add(sPorcentajeIva);
                        cmbIva.Text = sPorcentajeIva;
                        cmbNEiva.Items.Add(sPorcentajeIva);
                        cmbNEiva.Text = sPorcentajeIva;
                    }

                    foreach (string sBancos in cGlobal.bancos)
                    {
                        cmbBancos.Items.Add(sBancos);
                        cmbBancosNE.Items.Add(sBancos);
                    }

                    cmbMetodoPagoI.Text = "Efectivo";

                }//fin del try
                catch (Exception ex)
                {
                    MessageBox.Show("Error en la lectura del archivo de configuracion, motivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        #endregion

        #region compartida

            private void llenarListArticulos()
            {
                bDiferenciaIva = false;

                ListViewItem articulo1 = new ListViewItem(Convert.ToString(nFila));

                listArticulos.AddSubItem = true;
                listNEarticulos.AddSubItem = true;

                sUnidadMedida = new StringCollection();

                foreach (string dato in aUnidadMedida)
                {
                    sUnidadMedida.Add(dato);
                }
                articulo1.SubItems.Add(Convert.ToString(articulo.dCantidad));
                articulo1.SubItems.Add(articulo.sCodigo);
                articulo1.SubItems.Add(articulo.sDescripcion);

                if (articulo.sCodigo == cGlobal.sCodigoDiferenciaIva)
                {
                    bDiferenciaIva = true;
                }

                if (bFacturaElectronica) //pestaña facturacion
                {
                    articulo1.SubItems.Add(Convert.ToString(articulo.sUnidadMedida));
                    listArticulos.AddEditableCell(-1, 7);
                    listArticulos.AddEditableCell(-1, 8);
                    listArticulos.AddEditableCell(-1, 9);
                }
                else if (bImpresionDocumentos) //pestaña impresion
                {
                    //articulo1.SubItems.Add(Convert.ToString(sDatoLeido[incremento]));
                    articulo1.SubItems.Add(Convert.ToString(articulo.sUnidadMedida));
                    incremento = incremento + 1;
                }
                else if (bNotaCreditoElectronica) //pestaña nota de crédito electrónica
                {
                    articulo1.SubItems.Add(Convert.ToString(articulo.sUnidadMedida));
                }

                if (bDiferenciaIva)
                {
                    articulo1.SubItems.Add(0.00 + " " + sMoneda_ini);
                    articulo1.SubItems.Add(0.00 + " " + sMoneda_ini);

                    lPrecioUnitario.Add(0.00);
                    lImporte.Add(0.00);
                }
                else if(!bNotaCredito)
                {
                    articulo1.SubItems.Add(Convert.ToString(articulo.dPrecioUnitario) + " " + sMoneda_ini);
                    articulo1.SubItems.Add(Convert.ToString(articulo.dImporte) + " " + sMoneda_ini);

                    lPrecioUnitario.Add(articulo.dPrecioUnitario);
                    lImporte.Add(articulo.dImporte);
                }

                if (bFacturaElectronica) //pestaña de facturacion
                {
                    listArticulos.AddComboBoxCell(-1, 4, sUnidadMedida);

                    if (bDiferenciaIva || articulo.sCodigo == config.sCodigoConsolidado)
                    {
                        listArticulos.AddEditableCell(-1, 3);
                    }

                    listArticulos.Items.Add(articulo1);
                }

                else if (bImpresionDocumentos) //pestaña de impresion
                {
                    listArticulosImpresion.AddComboBoxCell(-1, 3, sUnidadMedida);
                    listArticulosImpresion.Items.Add(articulo1);
                }

                else if (bNotaCreditoElectronica) //pestaña nota de crédito electrónica
                {
                    listNEarticulos.AddComboBoxCell(-1, 4, sUnidadMedida);

                    if (bDiferenciaIva || articulo.sCodigo == config.sCodigoConsolidado)
                    {
                        listNEarticulos.AddEditableCell(-1, 3);
                    }

                    listNEarticulos.Items.Add(articulo1); 
                }

                else if (bNotaCredito)//pestaña de nota de credito
                {
                    articulo1 = new ListViewItem(Convert.ToString(nFila));
                    articulo1.SubItems.Add(Convert.ToString(articulo.dCantidad));
                    articulo1.SubItems.Add(articulo.sCodigo);
                    articulo1.SubItems.Add(articulo.sCodigo);
                    articulo1.SubItems.Add(articulo.sDescripcion);
                    articulo1.SubItems.Add(Convert.ToString(articulo.sFecha));
                    articulo1.SubItems.Add(articulo.sPedimento);
                    articulo1.SubItems.Add(Convert.ToString(articulo.dPrecioUnitario) + " " + sMoneda_ini);
                    articulo1.SubItems.Add(Convert.ToString(articulo.dImporte) + " " + sMoneda_ini);
                    
                    lPrecioUnitario.Add(articulo.dPrecioUnitario);
                    lImporte.Add(articulo.dImporte);

                    listNArticulos.Items.Add(articulo1);
                }
            }

            private void tabControl1_Selected(object sender, TabControlEventArgs e)
            {
                if (tabControl1.SelectedTab.Text == "IMPRESIÓN DE DOCUMENTOS")
                {
                    tabControl1.Size = new Size(787, 510);
                    frmFactura.ActiveForm.Size = new Size(802, 550);
                    btnImprimir.Enabled = false;
                    this.ActiveControl = txtIbuscar;
                }
                if (tabControl1.SelectedTab.Text == "FACTURACIÓN")
                {
                    tabControl1.Size = new Size(787, 525);
                    frmFactura.ActiveForm.Size = new Size(802, 565);
                    btnFacturar.Enabled = false;
                    this.ActiveControl = txtBuscar;
                }
                if (tabControl1.SelectedTab.Text == "NOTA DE CRÉDITO")
                {
                    tabControl1.Size = new Size(787, 505);
                    frmFactura.ActiveForm.Size = new Size(802, 545);
                    btnGenerarNota.Enabled = false;
                    this.ActiveControl = txtNbuscar;
                }
                if (tabControl1.SelectedTab.Text == "NOTA DE CRÉDITO ELECTRÓNICA")
                {
                    tabControl1.Size = new Size(787, 525);
                    frmFactura.ActiveForm.Size = new Size(802, 565);
                    btnGenerarNE.Enabled = false;
                    this.ActiveControl = txtNEbuscar;
                }
                if (tabControl1.SelectedTab.Text == "REPORTE")
                {
                    tabControl1.Size = new Size(787, 340);
                    frmFactura.ActiveForm.Size = new Size(802, 380);
                    btnGenerarNE.Enabled = false;
                    this.ActiveControl = txtPassReporte;
                }
                if (tabControl1.SelectedTab.Text == "GENERAR PDF")
                {
                    tabControl1.Size = new Size(787, 290);
                    frmFactura.ActiveForm.Size = new Size(802, 330);
                    btnGenerarNE.Enabled = false;
                    this.ActiveControl = txtCodigoPdf;
                }
            }

        #endregion

        #region facturaNotaCreditoElectronica

            private void obtenerMetodoPago()
            {
                int i;

                if (bFacturaElectronica)//pesataña facturación
                {
                    for (i = 0; i < config.sMetodosPago.Count; i++)
                    {
                        if (rbMetodoPago[i].Checked)
                        {
                            sMetodoPago = rbMetodoPago[i].Text;
                            i = i + 1;
                        }
                    }

                    if (rbdIvaRet.Checked)
                    {
                        sTipoIva = rbdIvaRet.Text;
                    }
                    else
                    {
                        sTipoIva = rdbIvaTrans.Text;
                    }
                }

                else if (bNotaCreditoElectronica) //pestaña nota de crédito electrónica
                {
                    for (i = 0; i < config.sMetodosPago.Count; i++)
                    {
                        if (rbMetodoPagoNE[i].Checked)
                        {
                            sMetodoPago = rbMetodoPagoNE[i].Text;
                            i = i + 1;
                        }
                    }

                    if (rdbIvaNERet.Checked)
                    {
                        sTipoIva = rdbIvaNERet.Text;
                    }
                    else
                    {
                        sTipoIva = rdbIvaNETrans.Text;
                    }
                }

                empresa = new cEmpresa();

                empresa.sColonia = config.sColonia;
                empresa.sCp = config.sCp;
                empresa.sDireccion = config.sDomicilio;
                empresa.sEmpresa = config.sNombre;
                empresa.sEstado = config.sEstado;
                empresa.sMunicipio = config.sMunicipio;
                empresa.sPais = config.spais;
                empresa.sRfc = config.sRfc;
                empresa.sTelefono = config.sTelefono;
                empresa.sRegimen = config.sRegimen;

                cliente.sPais = config.spais;

                //cliente.sMunicipio = "";
                //cliente.sPais = "";
                //cliente.sRegimen = "";
                //cliente.sLugarExpedicion = "";

                //articulo.dImporte = 1;
                //articulo.dCantidad = 1;
                //articulo.dPrecioUnitario = 1;
                //articulo.sCodigo = "";
                //articulo.sDescripcion = "";
                //articulo.sNumPedido = "";
                //articulo.sUnidadMedida = "";
            }

            private void facturar()
            {
                int iObtenerRegistros = 0;
                int iObtenerArticulosFila = 0;
                int iObtenerArticulosColumna = 0;
                string[] unidadMedida;

                bool bDatosCompletos = false;

                //comentado por la modificacion de iva retenido
                //if (bFacturaElectronica)
                //{
                //    listArticulos.FullRowSelect = true;
                //}
                //else if (bNotaCreditoElectronica)
                //{
                //    listNEarticulos.FullRowSelect = true;
                //}

                if (bFacturaElectronica) 
                { 
                    listArticulos.FullRowSelect = true;

                    //retencion de iva 21/04/2014
                    if (rbdIvaRet.Checked)
                    {
                        cliente.dTotal = cliente.sSubtotal;
                    }
                }
                else if (bNotaCreditoElectronica) 
                { 
                    listNEarticulos.FullRowSelect = true;

                    //retencion de iva 21/04/2014
                    if (rdbIvaNERet.Checked)
                    {
                        cliente.dTotal = cliente.sSubtotal;
                    }
                }

                if (cliente.sCodigo == "C0000" || cliente.sCodigo == "D0000")
                {
                    if (bFacturaElectronica)
                    {
                        cliente.sSerie = txtSerie.Text;
                        cliente.sRfc = txtRFC.Text.Trim();
                        cliente.sOrdenCompra = txtOrdenCompra.Text;
                        cliente.sPoblacion = txtPoblacion.Text;
                        cliente.sReferencia = txtReferencia.Text;
                        cliente.sCP = txtCP.Text;
                        cliente.sEstado = txtEstado.Text;
                        cliente.sColonia = txtColonia.Text;
                        cliente.sTelefono = txtTelefono.Text;
                        cliente.sPoblacion = txtPoblacion.Text;
                        cliente.sDireccion = txtDireccion.Text;
                        cliente.sCliente = txtCliente.Text;      
                    }
                    else if (bNotaCreditoElectronica)
                    {
                        cliente.sSerie = txtNEserie.Text;
                        cliente.sRfc = txtNErfc.Text.Trim();
                        cliente.sOrdenCompra = txtNEordencompra.Text;
                        cliente.sPoblacion = txtNEpoblacion.Text;
                        cliente.sReferencia = txtNEreferencia.Text;
                        cliente.sCP = txtNEcp.Text;
                        cliente.sEstado = txtNEestado.Text;
                        cliente.sColonia = txtNEcolonia.Text;
                        cliente.sTelefono = txtNEtelefono.Text;
                        cliente.sPoblacion = txtNEpoblacion.Text;
                        cliente.sDireccion = txtNEdireccion.Text;
                        cliente.sCliente = txtNEcliente.Text;
                    }
                }


                try
                {
                    //************************************************
                    // validar los combobox en la listview en la facturacion
                    //***********************************************

                    if (bFacturaElectronica) //pestaña de facturación
                    {
                        iObtenerRegistros = listArticulos.Items.Count;

                        iObtenerArticulosFila = listArticulos.Items.Count;
                        iObtenerArticulosColumna = listArticulos.Columns.Count;
                    }
                    else if (bNotaCreditoElectronica) //pestaña de nota de crédito electrónica
                    {
                        iObtenerRegistros = listNEarticulos.Items.Count;

                        iObtenerArticulosFila = listNEarticulos.Items.Count;
                        iObtenerArticulosColumna = listNEarticulos.Columns.Count;
                    }

                    unidadMedida = new string[iObtenerRegistros];

                    arrayObtenerArticulos = new string[iObtenerArticulosFila, iObtenerArticulosColumna];

                    for (i = 0; i < iObtenerRegistros; i++)
                    {
                        if (bFacturaElectronica)
                        {
                            listArticulos.Items[i].Selected = true;

                            if (listArticulos.SelectedItems[i].SubItems[4].Text == "") //columna 4 unidad de medida en el listArticulos
                            {
                                mensajesAdvertencia("INGRESE UNIDAD DE MEDIDA");
                                bDatosCompletos = false;
                                break;
                            }
                            else if (cmbCondicionesPago.Text == "")
                            {
                                mensajesAdvertencia("INGRESE CONDICIÓN DE PAGO");
                                bDatosCompletos = false;
                                break;
                            }
                            else
                            {
                                unidadMedida[i] = listArticulos.SelectedItems[i].SubItems[4].Text;

                                for (int c = 0; c < iObtenerArticulosColumna; c++)
                                {
                                    arrayObtenerArticulos[i, c] = listArticulos.SelectedItems[i].SubItems[c].Text;
                                }

                                bDatosCompletos = true;
                            }
                        }
                        else if (bNotaCreditoElectronica)
                        {
                            listNEarticulos.Items[i].Selected = true;

                            if (listNEarticulos.SelectedItems[i].SubItems[4].Text == "") //columna 4 unidad de medida en el listArticulos
                            {
                                mensajesAdvertencia("INGRESE UNIDAD DE MEDIDA");
                                bDatosCompletos = false;
                                break;
                            }
                            else if (cmbCondicionesPago.Text == "")
                            {
                                mensajesAdvertencia("INGRESE CONDICIÓN DE PAGO");
                                bDatosCompletos = false;
                                break;
                            }
                            else
                            {
                                unidadMedida[i] = listNEarticulos.SelectedItems[i].SubItems[4].Text;

                                for (int c = 0; c < iObtenerArticulosColumna; c++)
                                {
                                    arrayObtenerArticulos[i, c] = listNEarticulos.SelectedItems[i].SubItems[c].Text;
                                }

                                bDatosCompletos = true;
                            }
                        }
                    }
                    if (bDatosCompletos)
                    { obtenerMetodoPago(); }

                    cliente.sNumCtaPago = "NO IDENTIFICADO";
                    cliente.sBanco = "";
                    cliente.sDigitosBanco = "";

                    if (bDatosCompletos && sMetodoPago == "Transferencia Electronica" || sMetodoPago == "Cheque Nominativo" || sMetodoPago == "Tarjeta de cred. Deb o Serv.")
                    {
                        if (bNotaCreditoElectronica) //pestaña nota de crédito electrónica
                        {
                            if (txtBancoNE.TextLength == 4 && cmbBancosNE.Text != "")
                            {
                                cliente.sDigitosBanco = txtBancoNE.Text;
                                cliente.sBanco = cmbBancosNE.Text;
                                cliente.sNumCtaPago = cliente.sDigitosBanco + " " + cliente.sBanco;
                               
                            }
                            else
                            {
                                mensajesAdvertencia("FALTAN DATOS DEL BANCO");
                                //MessageBox.Show("Faltan datos del Banco", "Banco", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                bDatosCompletos = false;
                            }
                        }
                        else if (bFacturaElectronica) //pestaña facturación
                        {
                            if (txtBancoF.TextLength == 4 && cmbBancos.Text != "")
                            {
                                cliente.sDigitosBanco = txtBancoF.Text;
                                cliente.sBanco = cmbBancos.Text;
                                cliente.sNumCtaPago = cliente.sDigitosBanco + " " + cliente.sBanco;
                            }
                            else
                            {
                                //MessageBox.Show("Faltan datos del Banco", "Banco", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                                mensajesAdvertencia("FALTAN DATOS DEL BANCO");
                                bDatosCompletos = false;
                            }
                        }
                    }

                    if (bDatosCompletos)
                    {
                        guardarDatosSql(iObtenerArticulosFila, unidadMedida.Count());
                    }
                }
                catch (Exception ex)
                {
                    limpiarMensaje();
                    MessageBox.Show("Error en metodo factura, motivo: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }

            private void cfdiWs()
            {
                ProcessStartInfo cfdiInfo = new ProcessStartInfo(config.sRutaGetCfdi + config.sNombreGetCfdi);
                cfdiInfo.WindowStyle = ProcessWindowStyle.Minimized;
                
                Process.Start(cfdiInfo);
            }

            private void guardarDatosSql(int iObtenerArticulosFila, int iUnidadMedida)
            {
                Thread tCfdiws = new Thread(cfdiWs);

                try
                {
                    if (bFacturaElectronica) //factura
                    {
                        cliente.sCondicionPago = cmbCondicionesPago.Text;
                        cliente.dPorcentajeIva = Convert.ToDouble(cmbIva.Text);
                        cliente.sImporteLetra = txtImporte.Text;
                        cliente.sTipoComprobante = ComprobanteTipoDeComprobante.ingreso;
                    }
                    else if (bNotaCreditoElectronica) //nota electrónica
                    {
                        cliente.sCondicionPago = cmbCondicionesPagoNE.Text;
                        cliente.dPorcentajeIva = Convert.ToDouble(cmbNEiva.Text);
                        cliente.sImporteLetra = txtNEimporte.Text;
                        cliente.sTipoComprobante = ComprobanteTipoDeComprobante.egreso;
                    }

                    if (File.Exists(config.sRutaSalida + cGlobal.sNombreXml))
                    {
                        try
                        {
                            File.Delete(config.sRutaSalida + cGlobal.sNombreXml);
                        }
                        catch
                        { }
                    }

                    sConsultaSql = @"select id, numFactura from tblAuxiliar where numFactura=@numFactura;";

                    sqlConexion = new SqlConnection(cGlobal.sCadenaSql);

                    sqlConexion.Open();

                    sqlComando2 = new SqlCommand(sConsultaSql, sqlConexion);

                    sqlComando2.Parameters.AddWithValue("numFactura",((object)cliente.sDocumento) ?? DBNull.Value);

                    sqlLeerBuscarFacturaBD = sqlComando2.ExecuteReader();

                    if (sqlLeerBuscarFacturaBD.HasRows == false)
                    { // if (1) existe en la base de datos

                        sqlConexion.Close();

                            var confirmacion = MessageBox.Show("Debe confirmar para continuar", "Factura", MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation);

                            if (confirmacion == DialogResult.OK)
                            {
                                estadoSistema("Facturando...");
                                mensajesOk("GENERANDO XML CFD, ESPERE POR FAVOR...");

                                string sInsertarDatos;

                                crearXML(iObtenerArticulosFila);

                                if (bXmlOk)
                                {
                                    cGlobal.bTimbrado = false;
                                    iTiempoTimbrado = 0;

                                    if (bPrueba == false)
                                    {
                                        tCfdiws.Start(); //lanza el GetCfdi.exe
                                    }
                                    else if (bPrueba)
                                    {
                                        System.IO.File.Move(config.sRutaSalida + cGlobal.sNombreXml, config.sRutaEntrada + config.sIniciales + cGlobal.sNombreXml);
                                    }
                                    
                                    while (esperarTimbrado() == false)
                                    {
                                        mensajesAdvertencia("ESPERANDO TIMBRADO...");

                                        if (bFacturaElectronica)
                                        {
                                            btnBuscar.Enabled = false;
                                            btnFacturar.Enabled = false;
                                            btnSalir.Enabled = false;
                                        }
                                        else if (bNotaCreditoElectronica)
                                        {
                                            btnNEbuscar.Enabled = false;
                                            btnFacturar.Enabled = false;
                                            btnSalir.Enabled = false;
                                        }

                                        if (iTiempoTimbrado >= config.iTiempoTimbrado)
                                        {
                                            mensajesError("TIEMPO DE TIMBRADO EXCEDIDO, INTENTE DE NUEVO POR FAVOR");

                                            if (tCfdiws.ThreadState == System.Threading.ThreadState.Running)
                                            {
                                                tCfdiws.Abort();
                                            }

                                            System.Diagnostics.Process[] procesosGetCfdi;

                                            try
                                            {
                                                procesosGetCfdi = System.Diagnostics.Process.GetProcessesByName(config.sProcesoGetCfdi);
                                            
                                                foreach (System.Diagnostics.Process proceso in procesosGetCfdi)
                                                {
                                                    proceso.CloseMainWindow();
                                                }
                                            }
                                            catch
                                            { }

                                            if (bFacturaElectronica)
                                            {
                                                btnBuscar.Enabled = true;
                                                btnSalir.Enabled = true;
                                            }
                                            else if (bNotaCreditoElectronica)
                                            {
                                                btnNEbuscar.Enabled = true;
                                                btnSalir.Enabled = true;
                                            }

                                            break;
                                        }
                                    }


                                    if (File.Exists(config.sRutaSalida + cGlobal.sNombreXml))
                                    {
                                        try
                                        {
                                            File.Delete(config.sRutaSalida + cGlobal.sNombreXml);
                                        }
                                        catch
                                        { }
                                    }

                                    btnBuscar.Enabled = true;

                                    if (cGlobal.bTimbrado)
                                    {
                                        leerXML(iXmlArticulos, 6, config.sRutaEntrada + config.sIniciales + cGlobal.sNombreXml);

                                        if (cGlobal.bLecturaXmlFallo)
                                        {
                                            string sValor = cliente.dTotal.ToString("0000000000.000000", CultureInfo.InvariantCulture);

                                            QRCodeEncoder encoder = new QRCodeEncoder();
                                            cGlobal.cbb = encoder.Encode("?re=" + empresa.sRfc + "&rr=" + cliente.sRfc + "&tt=" + sValor + "&id=" + leerXml.sFolioFiscal,System.Text.Encoding.UTF8);

                                            sqlConexion = new SqlConnection(cGlobal.sCadenaSql);

                                            XDocument xDocXml;

                                            xDocXml = XDocument.Load(config.sRutaEntrada + config.sIniciales + cGlobal.sNombreXml);

                                            try
                                            {
                                                sInsertarDatos = @"Insert into tblAuxiliar (numFactura, usuario, xml,  
                                                tipoDocumento, fecha) Values (@numFactura,@usuario,'" + xDocXml + "',@tipoDocumento,@fecha);";
                                                
                                                sqlConexion.Open();

                                                sqlComando = new SqlCommand(sInsertarDatos, sqlConexion);

                                                sqlComando.Parameters.AddWithValue("numFactura", ((object)cliente.sDocumento) ?? DBNull.Value);
                                                sqlComando.Parameters.AddWithValue("usuario", ((object)cGlobal.sUserOk) ?? DBNull.Value);
                                                sqlComando.Parameters.AddWithValue("tipoDocumento", ((object)Convert.ToInt32(cliente.sNumtipodocto)) ?? DBNull.Value);
                                                sqlComando.Parameters.AddWithValue("fecha", ((object)DateTime.Now) ?? DBNull.Value);

                                                sqlComando.ExecuteNonQuery();

                                                sqlConexion.Close();
                                            }
                                            catch
                                            {
                                                if (sqlConexion.State == ConnectionState.Open) { sqlConexion.Close(); }

                                                sInsertarDatos = @"Insert into tblAuxiliar (numFactura, usuario,  
                                                tipoDocumento, fecha) Values (@numFactura,@usuario,@tipoDocumento,@fecha);";
                                                
                                                sqlConexion.Open();

                                                sqlComando = new SqlCommand(sInsertarDatos, sqlConexion);

                                                sqlComando.Parameters.AddWithValue("numFactura", ((object)cliente.sDocumento) ?? DBNull.Value);
                                                sqlComando.Parameters.AddWithValue("usuario", ((object)cGlobal.sUserOk) ?? DBNull.Value);
                                                sqlComando.Parameters.AddWithValue("tipoDocumento", ((object)Convert.ToInt32(cliente.sNumtipodocto)) ?? DBNull.Value);
                                                sqlComando.Parameters.AddWithValue("fecha", ((object)DateTime.Now) ?? DBNull.Value);

                                                sqlComando.ExecuteNonQuery();

                                                sqlConexion.Close();
                                            }

                                            sConsultaSql = @"select id from tblAuxiliar where numFactura=@numFactura
                                            and tipoDocumento=@tipoDocumento;";

                                            sqlConexion.Open();

                                            sqlComando2 = new SqlCommand(sConsultaSql, sqlConexion);

                                            sqlComando2.Parameters.AddWithValue("numFactura", ((object) cliente.sDocumento)?? DBNull.Value);
                                            sqlComando2.Parameters.AddWithValue("tipoDocumento", ((object) Convert.ToInt32(cliente.sNumtipodocto)) ?? DBNull.Value);

                                            sqlLeer = sqlComando2.ExecuteReader();

                                            while (sqlLeer.Read())
                                            {
                                                cGlobal.iDatoLeido = Convert.ToInt32(sqlLeer.GetValue(0));
                                            }

                                            sqlConexion.Close();

                                            guardarXml(iXmlArticulos);

                                            limpiarPantalla();
                                            btnFacturar.Enabled = false;

                                            cGlobal.sDocumento = cliente.sDocumento;
                                            cGlobal.sCodigoCliente = cliente.sCodigo;

                                            cGlobal.iTipoDocumento = Convert.ToInt32(cliente.sNumtipodocto);

                                            if (bFacturaElectronica && bGuardarXml)
                                            {
                                                frmReporteFactura reporte = new frmReporteFactura();
                                                reporte.Show();

                                                limpiarPantalla();

                                                mensajesOk("SE GENERO EXITOSAMENTE EL DOCUMENTO");

                                                estadoSistema("");

                                                btnFacturar.Enabled = false;

                                            }
                                            else if (bNotaCreditoElectronica && bGuardarXml)
                                            {
                                                frmReporteNotaElectronica reporte = new frmReporteNotaElectronica();
                                                reporte.Show();

                                                limpiarPantallaNE();
                                                
                                                mensajesOk("SE GENERO EXITOSAMENTE EL DOCUMENTO");

                                                estadoSistema("");

                                                btnGenerarNE.Enabled = false;

                                            }
                                            else 
                                            {
                                                mensajesOk("ERROR AL GUARDAR ARCHIVO XML EN BASE DE DATOS");

                                                estadoSistema("");
                                            }

                                        }
                                        else
                                        {
                                            mensajesError("ERROR EN LA LECTURA DEL ARCHIVO XML");
                                        }
                                    }
                                    else
                                    {
                                        mensajesError("ERROR EN EL TIMBRADO, INTENTE DE NUEVO POR FAVOR");
                                    }
                                }
                                else
                                {
                                    limpiarMensaje();
                                }
                            }
                            else
                            {
                                limpiarMensaje();
                            }

                        if (conexion.State == ConnectionState.Open) { conexion.Close(); }

                    }// if (1)
                    else
                    {
                        

                        if (bFacturaElectronica)
                        {  
                            limpiarPantalla();
                            btnFacturar.Enabled = false;
                        }
                        else if (bNotaCreditoElectronica)
                        { 
                            limpiarPantallaNE();
                            btnGenerarNE.Enabled = false;
                        }

                        limpiarMensaje();
                        mensajesAdvertencia("YA EXISTE EL DOCUMENTO EN LA BASE DE DATOS, VERIFIQUE POR FAVOR");   
                    }
                }
                catch (Exception ex)
                {
                    if (tCfdiws.ThreadState == System.Threading.ThreadState.Running)
                    {
                        tCfdiws.Abort();
                    }

                    if (File.Exists(config.sRutaSalida + cGlobal.sNombreXml))
                    {
                        File.Delete(config.sRutaSalida + cGlobal.sNombreXml);
                    }
                    limpiarMensaje();
                    MessageBox.Show("Error en el proceso de guardar datos de factura, motivo: " + ex.Message, "Error ",MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            private void mostrarMetodoPagoIva()
            {
                string sConsultaSql;

                sConsultaSql = @"select metodoPago, tipoIva from tblAuxiliar where numFactura = '" + cliente.sDocumento + "';";

                sqlConexion = new SqlConnection(cGlobal.sCadenaSql);

                sqlConexion.Open();

                sqlComando = new SqlCommand(sConsultaSql, sqlConexion);

                sqlLeer = sqlComando.ExecuteReader();

                while (sqlLeer.Read())
                {
                    cliente.sMetodoPago = sqlLeer.GetString(0).Trim();
                    cliente.sTipoIva = sqlLeer.GetString(1).Trim();
                }

                sqlLeer.Close();

                sqlConexion.Close();
            }

        #endregion

        #region xml     

            private bool esperarTimbrado()
            {
                if (File.Exists(config.sRutaEntrada + config.sIniciales + cGlobal.sNombreXml))
                {
                    System.Threading.Thread.Sleep(5000);
                    File.OpenRead(config.sRutaEntrada + config.sIniciales + cGlobal.sNombreXml);
                    cGlobal.bTimbrado = true;
                    return true;
                }
                else
                {
                    cGlobal.bTimbrado = false;
                    System.Threading.Thread.Sleep(2000);
                    iTiempoTimbrado = iTiempoTimbrado + 2000;
                    return false;
                }
            }    

            protected void crearXML(int articulos)
            {

                cCFDxml.VersionCFD version;
                version = cCFDxml.VersionCFD.CFDv3_2;

                cCFDxml cfds = new cCFDxml();
                DateTime fecha = DateTime.Now;

                string mes = fecha.Month.ToString("00", CultureInfo.InvariantCulture);
                string año = Convert.ToString(fecha.Year);
                string dia = fecha.Day.ToString("00", CultureInfo.InvariantCulture);
                string hora = fecha.Hour.ToString("00", CultureInfo.InvariantCulture);
                string minuto = fecha.Minute.ToString("00", CultureInfo.InvariantCulture);
                string segundo = fecha.Minute.ToString("00", CultureInfo.InvariantCulture);

                string fecha2 = año + "-" + mes + "-" + dia + "T" + hora + ":" + minuto + ":" + segundo;

                cfds.comprobante(version, cliente.sDocumento, fecha2, cliente.sCondicionPago, cliente.sSubtotal,
                    cliente.dTotal, cliente.sTipoComprobante, sMetodoPago, empresa.sMunicipio + ", " + empresa.sEstado,
                    cliente.sSerie, "", "", 0, "", cliente.sTipoMoneda, cliente.sNumCtaPago);
                cfds.AgregarEmisor(empresa.sRfc, empresa.sDireccion, empresa.sMunicipio, empresa.sEstado, empresa.sPais,
                    empresa.sCp, empresa.sEmpresa, "", "", empresa.sColonia);
                cfds.AgregarReceptor(cliente.sRfc, cliente.sCliente, cliente.sDireccion, cliente.sPoblacion, cliente.sEstado,
                    cliente.sPais, cliente.sCP, "", "", cliente.sColonia, cliente.sPoblacion, cliente.sReferencia);
                cfds.AgregaRegimenFiscal(empresa.sRegimen);
                cfds.AgregarEmisorExpedidoEn();

                for (f = 0; f < articulos; f++)
                {

                    cfds.AgregaConcepto(Convert.ToDouble(arrayObtenerArticulos[f, 1]), //cantidad 
                        arrayObtenerArticulos[f, 4],                                   //unidad medida
                        Convert.ToString(arrayObtenerArticulos[f, 3]),                 //descripcion
                        lPrecioUnitario[f],                                            //valor unitario
                        lImporte[f],                                                   //importe
                        arrayObtenerArticulos[f, 2]);                                  //codigo o noIdentificacion
                }

                //comentado para probar el codigo del iva retenido
                //cfds.AgregaImpuesto(ComprobanteImpuestosTrasladoImpuesto.IVA, cliente.dIva, cliente.dPorcentajeIva);

                //retencion de impuestos 21/04/2014
                if (bFacturaElectronica & rbdIvaRet.Checked)
                {
                    cfds.AgregaImpuesto(ComprobanteImpuestosTrasladoImpuesto.IVA, 0, cliente.dPorcentajeIva, ComprobanteImpuestosRetencionImpuesto.IVA, cliente.dIva);
                }
                else if (bNotaCreditoElectronica & rdbIvaNERet.Checked)
                {
                    cfds.AgregaImpuesto(ComprobanteImpuestosTrasladoImpuesto.IVA, 0, cliente.dPorcentajeIva, ComprobanteImpuestosRetencionImpuesto.IVA, cliente.dIva);
                }
                else
                {
                    cfds.AgregaImpuesto(ComprobanteImpuestosTrasladoImpuesto.IVA, cliente.dIva, cliente.dPorcentajeIva, ComprobanteImpuestosRetencionImpuesto.IVA, 0);
                }
                //retencion de impuestos


                string sPath = config.sRutaCertificado;
                string sCerFile = System.IO.Path.Combine(sPath, config.sNomCertificado);
                string sKeyFile = System.IO.Path.Combine(sPath, config.sNomArchivoKey);
                string sKeyPass = config.sPassCertificado;
                string sErrores = "";

                bXmlOk = false;
                iXmlArticulos = 0;

                if ((cfds.crearFacturaXML(sKeyFile, sKeyPass, sCerFile, sErrores, cliente.sDocumento)) == false)
                {
                    MessageBox.Show("Se encontraron los siguientes errores: " + sErrores, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    iXmlArticulos = articulos;
                    bXmlOk = true;
                    mensajesOk("ARCHIVO XML CFD CREADO EXITOSAMENTE, ESPERANDO TIMBRADO...");
                }

            }

            protected void leerXML(int iFila, int iColumna, string sNombreXml)
            {
                try
                {
                    leerXml = new cLeerXml();

                    XmlReader reader = XmlReader.Create(sNombreXml);
     
                    int fila = 0;
                    int col = 0;

                    arrayArticulos = new string[iFila, iColumna];

                    //Array.Clear(arrayObtenerArticulos, iFila, iColumna);

                    while (reader.Read())
                    {
                        if (reader.IsStartElement()) // checa si hay algo que leer.
                        {
                            // obtiene el nombre de cada elemento
                            switch (reader.Name)
                            {
                                case "cfdi:Comprobante":
                                    leerXml.sFolio = reader["folio"];
                                    leerXml.sLugarExpedicion = reader["LugarExpedicion"];
                                    leerXml.sMoneda = reader["Moneda"];
                                    leerXml.sNumCtaPago = reader["NumCtaPago"];
                                    leerXml.sCertificadoEmisor = reader["certificado"];
                                    leerXml.sNoCertificado = reader["noCertificado"];
                                    leerXml.sMetodoPago = reader["metodoDePago"];
                                    leerXml.sFormaPago = reader["formaDePago"];
                                    leerXml.sSelloDigitalEmisor = reader["sello"];
                                    leerXml.sFecha = reader["fecha"];
                                    leerXml.dSubtotal = Convert.ToDouble(reader["subTotal"]);
                                    leerXml.dTotal = Convert.ToDouble(reader["total"]);
                                    leerXml.sSerie = reader["serie"];
                                    leerXml.sImporteLetra = cliente.sImporteLetra;
                                    break;

                                case "cfdi:Emisor":
                                    leerXml.sEmpresaRfc = reader["rfc"];
                                    leerXml.sEmpresa = reader["nombre"];
                                    leerXml.sEmpresaTelefono = empresa.sTelefono;
                                    break;

                                case "cfdi:DomicilioFiscal":
                                    leerXml.sEmpresaDireccion = reader["calle"];
                                    leerXml.sEmpresaCp = reader["codigoPostal"];
                                    leerXml.sEmpresaEstado = reader["estado"];
                                    leerXml.sEmpresaMunicipio = reader["municipio"];
                                    leerXml.sEmpresaPais = reader["pais"];
                                    leerXml.sEmpresaColonia = reader["colonia"];
                                    break;

                                case "cfdi:ExpedidoEn":
                                    leerXml.sEmpresaCp = reader["codigoPostal"];
                                    leerXml.sEmpresaEstado = reader["estado"];
                                    leerXml.sEmpresaMunicipio = reader["municipio"];
                                    leerXml.sEmpresaPais = reader["pais"];
                                    break;

                                case "cfdi:RegimenFiscal":
                                    leerXml.sEmpresaRegimenFiscal = reader["Regimen"];
                                    break;

                                case "cfdi:Receptor":
                                    leerXml.sNombre = reader["nombre"];
                                    leerXml.sRfc = reader["rfc"];
                                    leerXml.sNoCte = cliente.sCodigo;
                                    leerXml.sOrdenCompra = cliente.sOrdenCompra;
                                    break;

                                case "cfdi:Domicilio":
                                    leerXml.scalle = reader["calle"];
                                    leerXml.sCp = reader["codigoPostal"];
                                    leerXml.sColonia = reader["colonia"];
                                    leerXml.sEstado = reader["estado"];
                                    leerXml.sCiudad = reader["municipio"];
                                    leerXml.sPais = reader["pais"];
                                    leerXml.sReferencia = reader["referencia"];
                                    leerXml.sTelefono = cliente.sTelefono;
                                    break;

                                case "cfdi:Concepto":
                                    col = 0;
                                    arrayArticulos[fila, col] = reader["cantidad"];
                                    col = col + 1;
                                    arrayArticulos[fila, col] = reader["descripcion"];
                                    col = col + 1;
                                    arrayArticulos[fila, col] = reader["importe"];
                                    col = col + 1;
                                    arrayArticulos[fila, col] = reader["noIdentificacion"];
                                    col = col + 1;
                                    arrayArticulos[fila, col] = reader["unidad"];
                                    col = col + 1;
                                    arrayArticulos[fila, col] = reader["valorUnitario"];
                                    fila = fila + 1;
                                    break;

                                case "cfdi:Impuestos":
                                    leerXml.sTotalImpuestosTrasladados = reader["totalImpuestosTrasladados"];
                                    leerXml.sTotalImpuestosRetenidos = reader["totalImpuestosRetenidos"];
                                    break;

                                case "cfdi:Traslado":
                                    leerXml.sIva = reader["importe"];
                                    leerXml.sPorcentajeIva = reader["tasa"];
                                    leerXml.sTipoImpuesto = reader["impuesto"];
                                    break;

                                case "cfdi:Retencion":
                                    leerXml.sIvaRetenido = reader["importe"];
                                    leerXml.sPorcentajeIvaRetenido = reader["tasa"];
                                    leerXml.sTipoImpuestoRetenido = reader["impuesto"];
                                    break;

                                case "tfd:TimbreFiscalDigital":
                                    leerXml.sFolioFiscal = reader["UUID"];
                                    leerXml.sFechaCertificadoCfdi = reader["FechaTimbrado"];
                                    leerXml.dtFechaCertificadoCfdi = Convert.ToDateTime(reader["FechaTimbrado"]);
                                    leerXml.sCertificadoSat = reader["noCertificadoSAT"];
                                    leerXml.sSelloDigitalSAT = reader["selloSAT"];
                                    leerXml.sSelloDigitalEmisor = reader["selloCFD"];

                                    leerXml.sCadenaOriginal = "||1.0|" + leerXml.sFolioFiscal + "|" +
                                        leerXml.sFechaCertificadoCfdi + "|" + leerXml.sSelloDigitalEmisor + "|" +
                                        leerXml.sCertificadoSat + "||";

                                    break;

                            }
                        }

                    }

                    cGlobal.bLecturaXmlFallo = true;
                }
                catch (Exception ex)
                {
                    limpiarMensaje();
                    MessageBox.Show("Error en la lectura del archivo XML, motivo: " + ex.Message, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    cGlobal.bLecturaXmlFallo = false;
                }

            }

            private void guardarXml(int iArticulos)
            {
                try
                {
                    string sInsertarDatos;
                    string sConsulta;
                    string sInsertar;
                    string sInsertarDatos2;
                    string sConsultaSql;
                    SqlCommand sqlComandoEmpresa;

                    sqlConexion = new SqlConnection(cGlobal.sCadenaSql);

                    DateTime fechaHora;
                    fechaHora = Convert.ToDateTime(leerXml.sFecha);

                    PictureBox codigo = new PictureBox();
                    codigo.Image = cGlobal.cbb as Image;

                    MemoryStream ms = new MemoryStream();

                    codigo.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                    byte[] imgByte = ms.GetBuffer();

                    bGuardarXml = false;

                    if (bFacturaElectronica)
                        { sDetalle = txtDetallesFactura.Text; }
                    else if (bNotaCreditoElectronica)
                        { sDetalle = txtDetallesNotaCredito.Text; }
                    else
                        { sDetalle = ""; }

                    leerXml.dIva = leerXml.dTotal - leerXml.dSubtotal;

                    sConsulta = @"select id from tblEmpresa where empresa = @empresa
                                    and telefono = @telefono
                                    and direccion = @direccion
                                    and rfc = @rfc
                                    and regimenFiscal = @regimenFiscal
                                    and cp = @cp
                                    and colonia = @colonia
                                    and estado = @estado
                                    and municipio = @municipio
                                    and pais = @pais
                                    and certificado = @certificado
                                    and noCertificado= @noCertificado;";

                    sqlConexion.Open();

                    sqlComando2 = new SqlCommand(sConsulta, sqlConexion);

                    sqlComando2.Parameters.AddWithValue("empresa", ((object)leerXml.sEmpresa) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("telefono", ((object)leerXml.sEmpresaTelefono) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("direccion", ((object)leerXml.sEmpresaDireccion) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("rfc", ((object)leerXml.sEmpresaRfc) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("regimenFiscal", ((object)leerXml.sEmpresaRegimenFiscal) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("cp", ((object)leerXml.sEmpresaCp) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("colonia", ((object)leerXml.sEmpresaColonia) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("estado", ((object)leerXml.sEmpresaEstado) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("municipio", ((object)leerXml.sEmpresaMunicipio) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("pais", ((object)leerXml.sEmpresaPais) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("certificado", ((object)leerXml.sCertificadoEmisor) ?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("noCertificado", ((object)leerXml.sNoCertificado) ?? DBNull.Value);

                    sqlLeer = sqlComando2.ExecuteReader();

                    if (sqlLeer.HasRows)
                    {
                        while (sqlLeer.Read())
                        {
                            cGlobal.iDatoLeidoXml = Convert.ToInt32(sqlLeer.GetValue(0));
                        }
                    }
                    else
                    {
                        sInsertar = @"Insert into tblEmpresa (empresa, telefono, direccion, rfc,
                                    regimenFiscal, cp, colonia, estado, municipio, pais, certificado, noCertificado)
                                    Values (@empresa, @telefono, @direccion, @rfc, @regimenFiscal, 
                                    @cp, @colonia, @estado, @municipio, @pais, @certificado, @noCertificado);";

                        sqlLeer.Close();

                        sqlComandoEmpresa = new SqlCommand(sInsertar, sqlConexion);

                        sqlComandoEmpresa.Parameters.AddWithValue("empresa", ((object)leerXml.sEmpresa) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("telefono", ((object)leerXml.sEmpresaTelefono) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("direccion", ((object)leerXml.sEmpresaDireccion) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("rfc", ((object)leerXml.sEmpresaRfc) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("regimenFiscal", ((object)leerXml.sEmpresaRegimenFiscal) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("cp", ((object)leerXml.sEmpresaCp) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("colonia", ((object)leerXml.sEmpresaColonia) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("estado", ((object)leerXml.sEmpresaEstado) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("municipio", ((object)leerXml.sEmpresaMunicipio) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("pais", ((object)leerXml.sEmpresaPais) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("certificado", ((object)leerXml.sCertificadoEmisor) ?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("noCertificado", ((object)leerXml.sNoCertificado) ?? DBNull.Value);

                        sqlComandoEmpresa.ExecuteNonQuery();

                        sqlConexion.Close();

                        sConsulta = @"select id from tblEmpresa where empresa = @empresa
                                    and telefono = @telefono
                                    and direccion = @direccion
                                    and rfc = @rfc
                                    and regimenFiscal = @regimenFiscal
                                    and cp = @cp
                                    and colonia = @colonia
                                    and estado = @estado
                                    and municipio = @municipio
                                    and pais = @pais
                                    and certificado = @certificado
                                    and noCertificado= @noCertificado;";

                        sqlConexion.Open();

                        sqlComando2 = new SqlCommand(sConsulta, sqlConexion);

                        sqlComando2.Parameters.AddWithValue("empresa", ((object)leerXml.sEmpresa)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("telefono", ((object)leerXml.sEmpresaTelefono)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("direccion", ((object)leerXml.sEmpresaDireccion)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("rfc", ((object)leerXml.sEmpresaRfc)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("regimenFiscal", ((object)leerXml.sEmpresaRegimenFiscal)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("cp", ((object)leerXml.sEmpresaCp)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("colonia", ((object)leerXml.sEmpresaColonia)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("estado", ((object)leerXml.sEmpresaEstado)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("municipio", ((object)leerXml.sEmpresaMunicipio)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("pais", ((object)leerXml.sEmpresaPais)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("certificado", ((object)leerXml.sCertificadoEmisor)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("noCertificado", ((object)leerXml.sNoCertificado) ?? DBNull.Value);

                        sqlLeer = sqlComando2.ExecuteReader();

                        while (sqlLeer.Read())
                        {
                            cGlobal.iDatoLeidoXml = Convert.ToInt32(sqlLeer.GetValue(0));
                        }
                    }

                    sqlLeer.Close();

                    sqlConexion.Close();

                    sInsertarDatos = @"Insert into tblXmlPdf (idEmpresa, idAuxiliar, folio, folioFiscalSat, certificadoSat, fechaCertificadoCfdi, 
                                    fecha, condiciones, metodoPago, lugarExpedicion, tasaCambio, moneda, subtotal, total, 
                                    porcentajeIva, retenciones, importeLetra, cadenaOriginal, selloDigitalEmisor, 
                                    selloDigitalSat, descuento, nombre, calle, municipio, estado, telefono, rfc, noCte,
                                    numCtaPago, colonia, cp, formaPago, imagen, serie, ordenCompra, iva, detalle)
                                    Values (@idEmpresa, @idAuxiliar, @folio, @folioFiscalSat, @certificadoSat, @fechaCertificadoCfdi, 
                                    @smallFecha, @condiciones, @metodoPago, @lugarExpedicion, @tasaCambio, @moneda, @subtotal, @total, 
                                    @porcentajeIva, @retenciones, @importeLetra, @cadenaOriginal, @selloDigitalEmisor,
                                    @selloDigitalSat, @descuento, @nombre, @calle, @municipio, @estado, @telefono, @rfc, @noCte,
                                    @numCtaPago, @colonia, @cp, @formaPago, @imagen, @serie, @ordenCompra, @iva, @detalle);";

                    sqlConexion.Open();
                    sqlComando = new SqlCommand(sInsertarDatos, sqlConexion);

                    if (leerXml.dtFechaCertificadoCfdi.Year < 1900)
                    {
                        leerXml.dtFechaCertificadoCfdi = DateTime.Now;
                    }

                    sqlComando.Parameters.AddWithValue("idEmpresa", ((object)cGlobal.iDatoLeidoXml)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("idAuxiliar", ((object)cGlobal.iDatoLeido)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("folio", ((object)leerXml.sFolio)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("folioFiscalSat", ((object)leerXml.sFolioFiscal)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("certificadoSat", ((object)leerXml.sCertificadoSat)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("condiciones", ((object)leerXml.sCondiciones)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("metodoPago", ((object)leerXml.sMetodoPago)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("lugarExpedicion", ((object)leerXml.sLugarExpedicion)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("tasaCambio", ((object)cGlobal.sTasaCambio)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("moneda", ((object)leerXml.sMoneda)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("subtotal", ((object)leerXml.dSubtotal)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("total", ((object)leerXml.dTotal)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("porcentajeIva", ((object)Convert.ToString(cliente.dPorcentajeIva))?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("retenciones", ((object)Convert.ToDouble(leerXml.sTotalImpuestosRetenidos))?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("importeLetra", ((object)leerXml.sImporteLetra)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("cadenaOriginal", ((object)leerXml.sCadenaOriginal)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("selloDigitalEmisor", ((object)leerXml.sSelloDigitalEmisor)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("selloDigitalSat", ((object)leerXml.sSelloDigitalSAT)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("descuento", ((object)Convert.ToDouble(leerXml.sDescuento))?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("calle", ((object)leerXml.scalle)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("municipio", ((object)leerXml.sCiudad)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("estado", ((object)leerXml.sEstado)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("telefono", ((object)leerXml.sTelefono)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("rfc", ((object)leerXml.sRfc)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("noCte", ((object)leerXml.sNoCte)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("numCtaPago", ((object)leerXml.sNumCtaPago)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("colonia", ((object)leerXml.sColonia)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("cp", ((object)leerXml.sCp)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("formaPago", ((object)leerXml.sFormaPago)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("serie", ((object)leerXml.sSerie)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("ordenCompra", ((object)leerXml.sOrdenCompra)?? DBNull.Value);
                    
                    //reemplazado para probar que se guarde el iva en el pdf aunque sea iva retenido 08/05/2014
                    //sqlComando.Parameters.AddWithValue("iva", ((object)leerXml.dIva) ?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("iva", ((object)cliente.dIva) ?? DBNull.Value);

                    sqlComando.Parameters.AddWithValue("detalle", ((object)sDetalle) ?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("nombre", ((object)leerXml.sNombre)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("imagen", ((object)imgByte)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("fechaCertificadoCfdi", ((object)leerXml.dtFechaCertificadoCfdi)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("smallFecha", ((object)fechaHora) ?? DBNull.Value);

                    sqlComando.ExecuteNonQuery();

                    sqlConexion.Close();

                    sConsultaSql = @"select id, folio from tblXmlPdf where folio=@folio;";

                    sqlConexion.Open();

                    sqlComando2 = new SqlCommand(sConsultaSql, sqlConexion);

                    sqlComando2.Parameters.AddWithValue("folio", ((object)leerXml.sFolio) ?? DBNull.Value);

                    sqlLeer = sqlComando2.ExecuteReader();

                    while (sqlLeer.Read())
                    {
                        cGlobal.iDatoLeidoXmlArticulos = Convert.ToInt32(sqlLeer.GetValue(0));
                    }

                    sqlConexion.Close();

                    for (int f = 0; f < iArticulos; f++)
                    {
                        sInsertarDatos2 = @"Insert into tblArticulosXmlPdf (idXml, herco, codigo, 
                                        cantidad, unidadMedida, descripcion, precioUnitario, importe, oc, atencion, planta) 
                                        Values(@idXml, @herco, @codigo, 
                                        @cantidad, @unidadMedida, @descripcion, @precioUnitario, @importe, @oc, @atencion, @planta);";

                        sqlConexion.Open();

                        sqlComando2 = new SqlCommand(sInsertarDatos2, sqlConexion);

                        sqlComando2.Parameters.AddWithValue("idXml", ((object)cGlobal.iDatoLeidoXmlArticulos)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("herco", ((object)arrayArticulos[f, 3])?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("codigo", ((object)arrayArticulos[f, 3])?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("cantidad", ((object)Convert.ToString(arrayObtenerArticulos[f, 1]))?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("unidadMedida", ((object)arrayArticulos[f, 4])?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("descripcion", ((object)arrayArticulos[f, 1])?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("precioUnitario", ((object)Convert.ToDouble(arrayArticulos[f, 5]))?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("importe", ((object)Convert.ToDouble(arrayArticulos[f, 2])) ?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("oc", ((object)Convert.ToString(arrayObtenerArticulos[f, 7])) ?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("atencion", ((object)Convert.ToString(arrayObtenerArticulos[f, 8])) ?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("planta", ((object)Convert.ToString(arrayObtenerArticulos[f, 9])) ?? DBNull.Value);

                        sqlComando2.ExecuteNonQuery();

                        sqlConexion.Close();
                    }


                    if (conexion.State == ConnectionState.Open) { conexion.Close(); }

                    bGuardarXml = true;

                }
                catch (Exception ex)
                {
                    bGuardarXml = false;
                    limpiarMensaje();
                    MessageBox.Show("Error en el proceso de guardar, motivo: " + ex.Message, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

        #endregion

        #region mensajesSistema

                private void mensajesAdvertencia(string mensaje)
                {
                    if (bFacturaElectronica)
                    {
                        txtMensajes.Visible = true;
                        txtMensajes.Text = mensaje;
                        txtMensajes.ForeColor = Color.Black;
                        txtMensajes.BackColor = Color.Yellow;
                    }
                    else if (bImpresionDocumentos)
                    {
                        txtImensajes.Visible = true;
                        txtImensajes.Text = mensaje;
                        txtImensajes.ForeColor = Color.Black;
                        txtImensajes.BackColor = Color.Yellow;
                    }
                    else if (bNotaCredito)
                    {
                        txtNmensajes.Visible = true;
                        txtNmensajes.Text = mensaje;
                        txtNmensajes.ForeColor = Color.Black;
                        txtNmensajes.BackColor = Color.Yellow;
                    }
                    else if (bNotaCreditoElectronica)
                    {
                        txtNEMensajes.Visible = true;
                        txtNEMensajes.Text = mensaje;
                        txtNEMensajes.ForeColor = Color.Black;
                        txtNEMensajes.BackColor = Color.Yellow;
                    }

                }

                private void mensajesOk(string mensaje)
                {
                    if (bFacturaElectronica)
                    {
                        txtMensajes.Visible = true;
                        txtMensajes.Text = mensaje;
                        txtMensajes.ForeColor = Color.White;
                        txtMensajes.BackColor = Color.Green;
                        txtMensajes.Update();
                    }
                    else if (bImpresionDocumentos)
                    {
                        txtImensajes.Visible = true;
                        txtImensajes.Text = mensaje;
                        txtImensajes.ForeColor = Color.White;
                        txtImensajes.BackColor = Color.Green;
                        txtImensajes.Update();
                    }
                    else if (bNotaCredito)
                    {
                        txtNmensajes.Visible = true;
                        txtNmensajes.Text = mensaje;
                        txtNmensajes.ForeColor = Color.White;
                        txtNmensajes.BackColor = Color.Green;
                        txtNmensajes.Update();
                    }
                    else if (bNotaCreditoElectronica)
                    {
                        txtNEMensajes.Visible = true;
                        txtNEMensajes.Text = mensaje;
                        txtNEMensajes.ForeColor = Color.White;
                        txtNEMensajes.BackColor = Color.Green;
                        txtNEMensajes.Update();
                    }
                }

                private void mensajesError(string mensaje)
                {
                    if (bFacturaElectronica)
                    {
                        txtMensajes.Visible = true;
                        txtMensajes.Text = mensaje;
                        txtMensajes.ForeColor = Color.White;
                        txtMensajes.BackColor = Color.Green;
                    }
                    else if (bImpresionDocumentos)
                    {
                        txtImensajes.Visible = true;
                        txtImensajes.Text = mensaje;
                        txtImensajes.ForeColor = Color.White;
                        txtImensajes.BackColor = Color.Green;
                    }
                    else if (bNotaCredito)
                    {
                        txtNmensajes.Visible = true;
                        txtNmensajes.Text = mensaje;
                        txtNmensajes.ForeColor = Color.White;
                        txtNmensajes.BackColor = Color.Green;
                    }
                    else if (bNotaCreditoElectronica)
                    {
                        txtNEMensajes.Visible = true;
                        txtNEMensajes.Text = mensaje;
                        txtNEMensajes.ForeColor = Color.White;
                        txtNEMensajes.BackColor = Color.Green;
                    }
                }

                private void limpiarMensaje()
                {
                    if (bFacturaElectronica)
                    {
                        txtMensajes.Text = "";
                        txtMensajes.BackColor = Color.Silver;
                        txtMensajes.Update();
                    }
                    else if (bImpresionDocumentos)
                    {
                        txtImensajes.Text = "";
                        txtImensajes.BackColor = Color.Silver;
                        txtImensajes.Update();
                    }
                    else if (bNotaCredito)
                    {
                        txtNmensajes.Text = "";
                        txtNmensajes.BackColor = Color.Silver;
                        txtNmensajes.Update();
                    }
                    else if (bNotaCreditoElectronica)
                    {
                        txtNEMensajes.Text = "";
                        txtNEMensajes.BackColor = Color.Silver;
                        txtNEMensajes.Update();
                    }
                }

                private void estadoSistema(string sMensaje)
                {
                    if (bFacturaElectronica)
                    {
                        if (sMensaje != "")
                        {
                            txtEstadoSistema.Text = sMensaje;
                            txtEstadoSistema.BackColor = Color.Gold;
                            txtEstadoSistema.Update();
                        }
                        else
                        {
                            txtEstadoSistema.Text = sMensaje;
                            txtEstadoSistema.BackColor = Color.Silver;
                            txtEstadoSistema.Update();
                        }
                    }
                    else if (bImpresionDocumentos)
                    {
                        if (sMensaje != "")
                        {
                            txtEstadoSistemaImpresion.Text = sMensaje;
                            txtEstadoSistemaImpresion.BackColor = Color.Gold;
                            txtEstadoSistemaImpresion.Update();
                        }
                        else
                        {
                            txtEstadoSistemaImpresion.Text = sMensaje;
                            txtEstadoSistemaImpresion.BackColor = Color.Silver;
                            txtEstadoSistemaImpresion.Update();
                        }
                    }
                    else if (bNotaCredito)
                    {
                        if (sMensaje != "")
                        {
                            txtEstadoSistemaNota.Text = sMensaje;
                            txtEstadoSistemaNota.BackColor = Color.Gold;
                            txtEstadoSistemaNota.Update();
                        }
                        else
                        {
                            txtEstadoSistemaNota.Text = sMensaje;
                            txtEstadoSistemaNota.BackColor = Color.Silver;
                            txtEstadoSistemaNota.Update();
                        }
                    }
                    else if (bNotaCreditoElectronica)
                    {
                        if (sMensaje != "")
                        {
                            txtEstadoSistemaNE.Text = sMensaje;
                            txtEstadoSistemaNE.BackColor = Color.Gold;
                            txtEstadoSistemaNE.Update();
                        }
                        else
                        {
                            txtEstadoSistemaNE.Text = sMensaje;
                            txtEstadoSistemaNE.BackColor = Color.Silver;
                            txtEstadoSistemaNE.Update();
                        }
                    }
                }

        #endregion

        private void frmFactura_FormClosed(object sender, FormClosedEventArgs e)
        {
            this.Dispose();
            Application.Exit();
        }

        private void btnGenerarReporte_Click(object sender, EventArgs e)
        {
            if (config.sPassReporte == txtPassReporte.Text)
            {
                if (mcFechaFinal.SelectionEnd.Date >= mcFechaInicial.SelectionEnd.Date)
                {
                    txtMensajesReportes.Text = "GENERANDO REPORTE, ESPERE POR FAVOR...";
                    txtMensajesReportes.BackColor = Color.Yellow;
                    txtMensajesReportes.ForeColor = Color.Black;
                    txtMensajesReportes.Update();

                    cGlobal.dtFechaInicial = mcFechaInicial.SelectionEnd.Date.AddHours(00).AddMinutes(00).AddSeconds(00);
                    cGlobal.dtFechaFinal = mcFechaFinal.SelectionEnd.Date.AddHours(23).AddMinutes(59).AddSeconds(59);

                    frmReporteFacturasTimbradas reporte = new frmReporteFacturasTimbradas();
                    reporte.Show();

                    txtMensajesReportes.Text = "REPORTE GENERADO EXITOSAMENTE";
                    txtPassReporte.Text = "";
                    txtMensajesReportes.BackColor = Color.Green;
                    txtMensajesReportes.ForeColor = Color.White;
                    txtMensajesReportes.Update();
                }
                else
                {
                    txtMensajesReportes.Text = "";
                    txtPassReporte.Text = "";
                    txtMensajesReportes.BackColor = Color.Silver;
                    txtMensajesReportes.Update();
                    MessageBox.Show("Rango de fechas incorrecto ", "Verificar ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            else
            { 
                txtMensajesReportes.Text = "";
                txtPassReporte.Text = "";
                txtMensajesReportes.BackColor = Color.Silver;
                txtMensajesReportes.Update();
                MessageBox.Show("No tiene autorización ", "Verificar ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        
        }

        #region generarPdf

        List<string> lcantidad;
        List<string> lDescripcion;
        List<string> limporte;
        List<string> lnoIdentificacion;
        List<string> lunidad;
        List<string> lvalorUnitario;
        bool bLecturaXmlFallo;
        double dIvaGenerarPdf;

        private void btnGenerarPDF_Click(object sender, EventArgs e)
        {
            if (txtCodigoPdf.Text == "" | cmbTipoDocPdf.Text == "")
            {
                MessageBox.Show("Faltan datos ", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }

            else
            {
                txtMensajesPdf.Text="GENERANDO PDF, ESPERE POR FAVOR...";
                txtMensajesPdf.BackColor = Color.Yellow;
                txtMensajesPdf.ForeColor = Color.Black;
                txtMensajesPdf.Update();

                bGuardarXml = false;
                bLecturaXmlFallo = false;
                lcantidad = new List<string>();
                lDescripcion = new List<string>();
                limporte = new List<string>();
                lnoIdentificacion = new List<string>();
                lunidad = new List<string>();
                lvalorUnitario = new List<string>();
                

                openFileDialog1.Filter = "Archivos XML(*.xml)|*.xml";
                openFileDialog1.Title = "Archivos XML";

                try
                {
                    leerXml = new cLeerXml();

                    if (openFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        cGlobal.sRutaXmlPdf = openFileDialog1.FileName;
                    }

                    openFileDialog1.Dispose();

                    XmlReader reader = XmlReader.Create(cGlobal.sRutaXmlPdf);

                    while (reader.Read())
                    {
                        if (reader.IsStartElement()) // checa si hay algo que leer.
                        {
                            // obtiene el nombre de cada elemento
                            switch (reader.Name)
                            {
                                case "cfdi:Comprobante":
                                    leerXml.sFolio = reader["folio"];
                                    leerXml.sLugarExpedicion = reader["LugarExpedicion"];
                                    leerXml.sMoneda = reader["Moneda"];
                                    leerXml.sNumCtaPago = reader["NumCtaPago"];
                                    leerXml.sCertificadoEmisor = reader["certificado"];
                                    leerXml.sNoCertificado = reader["noCertificado"];
                                    leerXml.sMetodoPago = reader["metodoDePago"];
                                    leerXml.sFormaPago = reader["formaDePago"];
                                    leerXml.sSelloDigitalEmisor = reader["sello"];
                                    leerXml.sFecha = reader["fecha"];
                                    leerXml.dSubtotal = Convert.ToDouble(reader["subTotal"]);
                                    leerXml.dTotal = Convert.ToDouble(reader["total"]);
                                    leerXml.sSerie = reader["serie"];
                                    break;

                                case "cfdi:Emisor":
                                    leerXml.sEmpresaRfc = reader["rfc"];
                                    leerXml.sEmpresa = reader["nombre"];
                                    leerXml.sEmpresaTelefono = config.sTelefono;
                                    break;

                                case "cfdi:DomicilioFiscal":
                                    leerXml.sEmpresaDireccion = reader["calle"];
                                    leerXml.sEmpresaCp = reader["codigoPostal"];
                                    leerXml.sEmpresaEstado = reader["estado"];
                                    leerXml.sEmpresaMunicipio = reader["municipio"];
                                    leerXml.sEmpresaPais = reader["pais"];
                                    leerXml.sEmpresaColonia = reader["colonia"];
                                    break;

                                case "cfdi:ExpedidoEn":
                                    leerXml.sEmpresaCp = reader["codigoPostal"];
                                    leerXml.sEmpresaEstado = reader["estado"];
                                    leerXml.sEmpresaMunicipio = reader["municipio"];
                                    leerXml.sEmpresaPais = reader["pais"];
                                    break;

                                case "cfdi:RegimenFiscal":
                                    leerXml.sEmpresaRegimenFiscal = reader["Regimen"];
                                    break;

                                case "cfdi:Receptor":
                                    leerXml.sNombre = reader["nombre"];
                                    leerXml.sRfc = reader["rfc"];
                                    leerXml.sNoCte = txtCodigoPdf.Text.Trim();
                                    leerXml.sOrdenCompra = txtOrdenPdf.Text.Trim();
                                    break;

                                case "cfdi:Domicilio":
                                    leerXml.scalle = reader["calle"];
                                    leerXml.sCp = reader["codigoPostal"];
                                    leerXml.sColonia = reader["colonia"];
                                    leerXml.sEstado = reader["estado"];
                                    leerXml.sCiudad = reader["municipio"];
                                    leerXml.sPais = reader["pais"];
                                    leerXml.sReferencia = reader["referencia"];
                                    leerXml.sTelefono = txtTelefonoReceptor.Text;
                                    break;

                                case "cfdi:Concepto":
                                    lcantidad.Add(reader["cantidad"]);
                                    lDescripcion.Add(reader["descripcion"]);
                                    limporte.Add(reader["importe"]);
                                    lnoIdentificacion.Add(reader["noIdentificacion"]);
                                    lunidad.Add(reader["unidad"]);
                                    lvalorUnitario.Add(reader["valorUnitario"]);
                                    break;

                                case "cfdi:Impuestos":
                                    leerXml.sTotalImpuestosTrasladados = reader["totalImpuestosTrasladados"];
                                    leerXml.sTotalImpuestosRetenidos = reader["totalImpuestosRetenidos"];
                                    break;

                                case "cfdi:Traslado":
                                    leerXml.sIva = reader["importe"];
                                    leerXml.sPorcentajeIva = reader["tasa"];
                                    leerXml.sTipoImpuesto = reader["impuesto"];
                                    break;

                                case "cfdi:Retencion":
                                    leerXml.sIvaRetenido = reader["importe"];
                                    leerXml.sPorcentajeIvaRetenido = reader["tasa"];
                                    leerXml.sTipoImpuestoRetenido = reader["impuesto"];
                                    break;

                                case "tfd:TimbreFiscalDigital":
                                    leerXml.sFolioFiscal = reader["UUID"];
                                    leerXml.sFechaCertificadoCfdi = reader["FechaTimbrado"];
                                    leerXml.dtFechaCertificadoCfdi = Convert.ToDateTime(reader["FechaTimbrado"]);
                                    leerXml.sCertificadoSat = reader["noCertificadoSAT"];
                                    leerXml.sSelloDigitalSAT = reader["selloSAT"];
                                    leerXml.sSelloDigitalEmisor = reader["selloCFD"];

                                    leerXml.sCadenaOriginal = "||1.0|" + leerXml.sFolioFiscal + "|" +
                                        leerXml.sFechaCertificadoCfdi + "|" + leerXml.sSelloDigitalEmisor + "|" +
                                        leerXml.sCertificadoSat + "||";

                                    break;

                            }
                        }

                    }

                    //Agregado para adecuar la generacion del pdf acorde al iva retenido 08/05/2014

                    if (leerXml.sTotalImpuestosTrasladados != null & Convert.ToDouble(leerXml.sTotalImpuestosTrasladados) == 0 & leerXml.sTotalImpuestosRetenidos != null)
                    {
                        dIvaGenerarPdf = Convert.ToDouble(leerXml.sTotalImpuestosRetenidos);
                    }
                    else if (leerXml.sTotalImpuestosTrasladados != null & leerXml.sTotalImpuestosTrasladados != "0")
                    {
                        dIvaGenerarPdf = Convert.ToDouble(leerXml.sTotalImpuestosTrasladados);
                    }
                    else
                    { 
                        dIvaGenerarPdf = Convert.ToDouble(leerXml.sIva);
                    }

                    //////////////////////////////////////////////////////////////////////////////

                    string sTipoMoneda = "";

                    if (leerXml.sMoneda == "MXN")
                    { sTipoMoneda = "PESOS"; }
                    else if (leerXml.sMoneda == "USD")
                    { sTipoMoneda = "DOLARES"; }
                    else if (leerXml.sMoneda == "EUR")
                    { sTipoMoneda = "EUROS"; }

                    Numeros_letras convertirXmlLeido = new Numeros_letras();
                    leerXml.sImporteLetra = convertirXmlLeido.enletras(Convert.ToString(leerXml.dTotal), leerXml.sMoneda, sTipoMoneda);

                    bLecturaXmlFallo = true;
                    guardarXmlPdf(lnoIdentificacion.Count());
                }
                catch (Exception ex)
                {
                    txtMensajesPdf.Text = "";
                    txtMensajesPdf.BackColor = Color.Silver;
                    txtMensajesPdf.Update();

                    MessageBox.Show("Error en la lectura del archivo XML, motivo: " + ex.Message, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    bLecturaXmlFallo = false;
                }

            }
        }

        private void guardarXmlPdf(int iArticulos)
        {
            try
            {
                string sInsertarDatos;
                string sConsulta;
                string sInsertar;
                string sInsertarDatos2;
                bool bActualizar;
                SqlCommand sqlComandoEmpresa;

                bActualizar = false;

                sqlConexion = new SqlConnection(cGlobal.sCadenaSql);

                config.iTipoDocumento = Convert.ToInt32(cmbTipoDocPdf.Text);

                string sValor = leerXml.dTotal.ToString("0000000000.000000", CultureInfo.InvariantCulture);

                QRCodeEncoder encoder = new QRCodeEncoder();

                //opciones para resolver el problema con le código qr 07/07/2014, se utilizo el System.Text.Encoding.UTF8
                //encoder.QRCodeErrorCorrect = QRCodeEncoder.ERROR_CORRECTION.Q;
                //encoder.QRCodeVersion = 0;

                cGlobal.cbb = encoder.Encode("?re=" + leerXml.sEmpresaRfc + "&rr=" + leerXml.sRfc + "&tt=" + sValor + "&id=" + leerXml.sFolioFiscal, System.Text.Encoding.UTF8);

                DateTime fechaHora;
                fechaHora = Convert.ToDateTime(leerXml.sFecha);

                PictureBox codigo = new PictureBox();
                codigo.Image = cGlobal.cbb as Image;

                MemoryStream ms = new MemoryStream();

                codigo.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg);
                byte[] imgByte = ms.GetBuffer();

                XDocument xDocXml;

                xDocXml = XDocument.Load(cGlobal.sRutaXmlPdf);
                

                sConsultaSql = @"select id from tblAuxiliar where numFactura=@numFactura
                                            and tipoDocumento=@tipoDocumento;";

                sqlConexion.Open();

                sqlComando2 = new SqlCommand(sConsultaSql, sqlConexion);

                sqlComando2.Parameters.AddWithValue("numFactura", ((object)leerXml.sFolio) ?? DBNull.Value);
                sqlComando2.Parameters.AddWithValue("tipoDocumento", ((object)Convert.ToInt32(cmbTipoDocPdf.Text)) ?? DBNull.Value);
                
                sqlLeer = sqlComando2.ExecuteReader();

                if (sqlLeer.HasRows)
                {
                    bActualizar = true;
                    //txtMensajesPdf.Text = "";
                    //txtMensajesPdf.BackColor = Color.Silver;
                    //txtMensajesPdf.Update();
                    //sqlConexion.Close();
                    //MessageBox.Show("Ya existe el documento en la base de datos ", "Advertencia ", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                    sqlLeer.Close();

                    sqlConexion.Close();

                    try
                    {
                        if (bActualizar)
                        {
                            //actualizar
                            sInsertarDatos = @"update tblAuxiliar set numFactura=@numFactura, usuario=@usuario, 
                            xml='" + xDocXml + "', tipoDocumento=@tipoDocumento, fecha=@fecha where numFactura=@numFactura and tipoDocumento=@tipoDocumento;";
                        }
                        else
                        {
                            //insertar
                            sInsertarDatos = @"Insert into tblAuxiliar (numFactura, usuario, xml, tipoDocumento, fecha) 
                            Values (@numFactura , @usuario , '" + xDocXml + "' , @tipoDocumento , @fecha);";
                        }

                        sqlConexion.Open();

                        sqlComando = new SqlCommand(sInsertarDatos, sqlConexion);

                        sqlComando.Parameters.AddWithValue("fecha", ((object)Convert.ToDateTime(leerXml.sFecha)) ?? DBNull.Value);
                        //sqlComando.Parameters.AddWithValue("xml",  ((object)xDocXml)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("numFactura", ((object)leerXml.sFolio) ?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("usuario", ((object)cGlobal.sUserOk) ?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("tipoDocumento", ((object)Convert.ToInt32(cmbTipoDocPdf.Text)) ?? DBNull.Value);

                        sqlComando.ExecuteNonQuery();

                        sqlConexion.Close();
                    }
                    catch
                    {
                        if (sqlConexion.State == ConnectionState.Open) { sqlConexion.Close(); }

                        if (bActualizar)
                        {
                            sInsertarDatos = @"update tblAuxiliar set numFactura=@numFactura, usuario=@usuario, tipoDocumento=@tipoDocumento, fecha=@fecha 
                            where numFactura=@numFactura and tipoDocumento=@tipoDocumento;";
                        }
                        else
                        {
                            sInsertarDatos = @"Insert into tblAuxiliar (numFactura, usuario, tipoDocumento, fecha) 
                            Values (@numFactura, @usuario, @tipoDocumento, @fecha);";
                        }

                        sqlConexion.Open();

                        sqlComando = new SqlCommand(sInsertarDatos, sqlConexion);

                        sqlComando.Parameters.AddWithValue("fecha", ((object)Convert.ToDateTime(leerXml.sFecha)) ?? DBNull.Value);
                        //sqlComando.Parameters.AddWithValue("xml",  ((object)xDocXml)?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("numFactura", ((object)leerXml.sFolio) ?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("usuario", ((object)cGlobal.sUserOk) ?? DBNull.Value);
                        sqlComando.Parameters.AddWithValue("tipoDocumento", ((object)Convert.ToInt32(cmbTipoDocPdf.Text)) ?? DBNull.Value);

                        sqlComando.ExecuteNonQuery();

                        sqlConexion.Close();
                    }

                    sConsultaSql = @"select id from tblAuxiliar where numFactura=@numFactura
                                                and tipoDocumento=@tipoDocumento;";

                    sqlConexion.Open();

                    sqlComando2 = new SqlCommand(sConsultaSql, sqlConexion);

                    sqlComando2.Parameters.AddWithValue("numFactura", ((object)leerXml.sFolio)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("tipoDocumento", ((object)Convert.ToInt32(cmbTipoDocPdf.Text)) ?? DBNull.Value);

                    sqlLeer = sqlComando2.ExecuteReader();

                    while (sqlLeer.Read())
                    {
                        cGlobal.iDatoLeido = Convert.ToInt32(sqlLeer.GetValue(0));

                    }
                    cGlobal.sDocumento = leerXml.sFolio;
                    cGlobal.sCodigoCliente = leerXml.sNoCte;

                    sqlLeer.Close();

                    sqlConexion.Close();

                    bGuardarXml = false;

                    leerXml.dIva = leerXml.dTotal - leerXml.dSubtotal;

                    sConsulta = @"select id from tblEmpresa where empresa = @empresa
                                    and telefono = @telefono
                                    and direccion = @direccion
                                    and rfc = @rfc
                                    and regimenFiscal = @regimenFiscal
                                    and cp = @cp
                                    and colonia = @colonia
                                    and estado = @estado
                                    and municipio = @municipio
                                    and pais = @pais
                                    and certificado = @certificado
                                    and noCertificado= @noCertificado;";

                    sqlConexion.Open();

                    sqlComando2 = new SqlCommand(sConsulta, sqlConexion);

                    sqlComando2.Parameters.AddWithValue("empresa",  ((object)leerXml.sEmpresa)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("telefono",  ((object)leerXml.sEmpresaTelefono)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("direccion",  ((object)leerXml.sEmpresaDireccion)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("rfc",  ((object)leerXml.sEmpresaRfc)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("regimenFiscal",  ((object)leerXml.sEmpresaRegimenFiscal)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("cp",  ((object)leerXml.sEmpresaCp)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("colonia",  ((object)leerXml.sEmpresaColonia)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("estado",  ((object)leerXml.sEmpresaEstado)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("municipio",  ((object)leerXml.sEmpresaMunicipio)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("pais",  ((object)leerXml.sEmpresaPais)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("certificado",  ((object)leerXml.sCertificadoEmisor)?? DBNull.Value);
                    sqlComando2.Parameters.AddWithValue("noCertificado", ((object)leerXml.sNoCertificado) ?? DBNull.Value);

                    sqlLeer = sqlComando2.ExecuteReader();

                    if (sqlLeer.HasRows)
                    {
                        while (sqlLeer.Read())
                        {
                            cGlobal.iDatoLeidoXml = Convert.ToInt32(sqlLeer.GetValue(0));
                        }
                    }
                    else
                    {
                        sInsertar = @"Insert into tblEmpresa (empresa, telefono, direccion, rfc,
                                    regimenFiscal, cp, colonia, estado, municipio, pais, certificado, noCertificado)
                                    Values (@empresa, @telefono, @direccion, @rfc, @regimenFiscal, 
                                    @cp, @colonia, @estado, @municipio, @pais, @certificado, @noCertificado);";

                        sqlLeer.Close();

                        sqlComandoEmpresa = new SqlCommand(sInsertar, sqlConexion);

                        sqlComandoEmpresa.Parameters.AddWithValue("empresa", ((object)leerXml.sEmpresa)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("telefono", ((object)leerXml.sEmpresaTelefono)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("direccion", ((object)leerXml.sEmpresaDireccion)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("rfc", ((object)leerXml.sEmpresaRfc)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("regimenFiscal", ((object)leerXml.sEmpresaRegimenFiscal)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("cp", ((object)leerXml.sEmpresaCp)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("colonia", ((object)leerXml.sEmpresaColonia)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("estado", ((object)leerXml.sEmpresaEstado)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("municipio", ((object)leerXml.sEmpresaMunicipio)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("pais", ((object)leerXml.sEmpresaPais)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("certificado", ((object)leerXml.sCertificadoEmisor)?? DBNull.Value);
                        sqlComandoEmpresa.Parameters.AddWithValue("noCertificado", ((object)leerXml.sNoCertificado) ?? DBNull.Value);
                        
                        sqlComandoEmpresa.ExecuteNonQuery();

                        sqlConexion.Close();

                        sConsulta = @"select id from tblEmpresa where empresa = @empresa
                                    and telefono = @telefono
                                    and direccion = @direccion
                                    and rfc = @rfc
                                    and regimenFiscal = @regimenFiscal
                                    and cp = @cp
                                    and colonia = @colonia
                                    and estado = @estado
                                    and municipio = @municipio
                                    and pais = @pais
                                    and certificado = @certificado
                                    and noCertificado= @noCertificado;";

                        sqlConexion.Open();

                        sqlComando2 = new SqlCommand(sConsulta, sqlConexion);

                        sqlComando2.Parameters.AddWithValue("empresa", ((object)leerXml.sEmpresa)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("telefono", ((object)leerXml.sEmpresaTelefono)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("direccion", ((object)leerXml.sEmpresaDireccion)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("rfc", ((object)leerXml.sEmpresaRfc)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("regimenFiscal", ((object)leerXml.sEmpresaRegimenFiscal)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("cp", ((object)leerXml.sEmpresaCp)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("colonia", ((object)leerXml.sEmpresaColonia)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("estado", ((object)leerXml.sEmpresaEstado)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("municipio", ((object)leerXml.sEmpresaMunicipio)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("pais", ((object)leerXml.sEmpresaPais)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("certificado", ((object)leerXml.sCertificadoEmisor)?? DBNull.Value);
                        sqlComando2.Parameters.AddWithValue("noCertificado", ((object)leerXml.sNoCertificado) ?? DBNull.Value);

                        sqlLeer = sqlComando2.ExecuteReader();

                        while (sqlLeer.Read())
                        {
                            cGlobal.iDatoLeidoXml = Convert.ToInt32(sqlLeer.GetValue(0));
                        }
                    }

                    sqlLeer.Close();

                    sqlConexion.Close();

                    if (bActualizar)
                    {
                        sInsertarDatos = @"update tblXmlPdf  set idEmpresa=@idEmpresa, idAuxiliar=@idAuxiliar, 
                                    folio=@folio, folioFiscalSat=@folioFiscalSat, certificadoSat=@certificadoSat, 
                                    fechaCertificadoCfdi=@fechaCertificadoCfdi, fecha=@smallFecha, 
                                    condiciones=@condiciones, metodoPago=@metodoPago, lugarExpedicion=@lugarExpedicion, 
                                    tasaCambio=@tasaCambio, moneda=@moneda, subtotal=@subtotal, total=@total, 
                                    porcentajeIva=@porcentajeIva, retenciones=@retenciones, importeLetra=@importeLetra, 
                                    cadenaOriginal=@cadenaOriginal, selloDigitalEmisor=@selloDigitalEmisor, 
                                    selloDigitalSat=@selloDigitalSat, descuento=@descuento, nombre=@nombre, 
                                    calle=@calle, municipio=@municipio, estado=@estado, telefono=@telefono, rfc=@rfc, noCte=@noCte,
                                    numCtaPago=@numCtaPago, colonia=@colonia, cp=@cp, formaPago=@formaPago, 
                                    imagen=@imagen, serie=@serie, ordenCompra=@ordenCompra, iva=@iva, detalle=@detalle
                                    where idAuxiliar=@idAuxiliar;";
                    }
                    else
                    {
                        sInsertarDatos = @"Insert into tblXmlPdf (idEmpresa, idAuxiliar, folio, folioFiscalSat, certificadoSat, fechaCertificadoCfdi, 
                                    fecha, condiciones, metodoPago, lugarExpedicion, tasaCambio, moneda, subtotal, total, 
                                    porcentajeIva, retenciones, importeLetra, cadenaOriginal, selloDigitalEmisor, 
                                    selloDigitalSat, descuento, nombre, calle, municipio, estado, telefono, rfc, noCte,
                                    numCtaPago, colonia, cp, formaPago, imagen, serie, ordenCompra, iva, detalle)
                                    Values (@idEmpresa, @idAuxiliar, @folio, @folioFiscalSat, @certificadoSat, @fechaCertificadoCfdi, 
                                    @smallFecha, @condiciones, @metodoPago, @lugarExpedicion, @tasaCambio, @moneda, @subtotal, @total, 
                                    @porcentajeIva, @retenciones, @importeLetra, @cadenaOriginal, @selloDigitalEmisor,
                                    @selloDigitalSat, @descuento, @nombre, @calle, @municipio, @estado, @telefono, @rfc, @noCte,
                                    @numCtaPago, @colonia, @cp, @formaPago, @imagen, @serie, @ordenCompra, @iva, @detalle);";
                    }

                    sqlConexion.Open();

                    sqlComando = new SqlCommand(sInsertarDatos, sqlConexion);

                    if (leerXml.dtFechaCertificadoCfdi.Year < 1900)
                    {
                        leerXml.dtFechaCertificadoCfdi = DateTime.Now;
                    }

                    sqlComando.Parameters.AddWithValue("idEmpresa", ((object)cGlobal.iDatoLeidoXml)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("idAuxiliar", ((object)cGlobal.iDatoLeido)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("folio", ((object)leerXml.sFolio)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("folioFiscalSat", ((object)leerXml.sFolioFiscal)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("certificadoSat", ((object)leerXml.sCertificadoSat)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("condiciones", ((object)leerXml.sCondiciones)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("metodoPago", ((object)leerXml.sMetodoPago)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("lugarExpedicion", ((object)leerXml.sLugarExpedicion)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("tasaCambio", ((object)cGlobal.sTasaCambio)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("moneda", ((object)leerXml.sMoneda)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("subtotal", ((object)leerXml.dSubtotal)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("total", ((object)leerXml.dTotal)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("porcentajeIva", ((object)leerXml.sPorcentajeIva)?? DBNull.Value);
                    
                    //modificado para probar la retencion de impuestos en la pestaña generar pdf
                    //sqlComando.Parameters.AddWithValue("retenciones", ((object)Convert.ToDouble(leerXml.sRetenciones))?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("retenciones", ((object)Convert.ToDouble(leerXml.sTotalImpuestosRetenidos)) ?? DBNull.Value);

                    sqlComando.Parameters.AddWithValue("importeLetra", ((object)leerXml.sImporteLetra)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("cadenaOriginal", ((object)leerXml.sCadenaOriginal)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("selloDigitalEmisor", ((object)leerXml.sSelloDigitalEmisor)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("selloDigitalSat", ((object)leerXml.sSelloDigitalSAT)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("descuento", ((object)Convert.ToDouble(leerXml.sDescuento))?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("calle", ((object)leerXml.scalle)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("municipio", ((object)leerXml.sCiudad)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("estado", ((object)leerXml.sEstado)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("telefono", ((object)leerXml.sTelefono)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("rfc", ((object)leerXml.sRfc)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("noCte", ((object)leerXml.sNoCte)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("numCtaPago", ((object)leerXml.sNumCtaPago)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("colonia", ((object)leerXml.sColonia)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("cp", ((object)leerXml.sCp)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("formaPago", ((object)leerXml.sFormaPago)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("serie", ((object)leerXml.sSerie)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("ordenCompra", ((object)leerXml.sOrdenCompra)?? DBNull.Value);

                    //modificado para probar el iva retenido 08/05/2014
                    //sqlComando.Parameters.AddWithValue("iva", ((object)leerXml.dIva) ?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("iva", ((object)dIvaGenerarPdf) ?? DBNull.Value);

                    sqlComando.Parameters.AddWithValue("detalle", ((object)txtDetallePDF.Text)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("nombre", ((object)leerXml.sNombre)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("imagen", ((object)imgByte)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("fechaCertificadoCfdi", ((object)leerXml.dtFechaCertificadoCfdi)?? DBNull.Value);
                    sqlComando.Parameters.AddWithValue("smallFecha", ((object)fechaHora) ?? DBNull.Value);

                    sqlComando.ExecuteNonQuery();

                    sqlConexion.Close();

                    sConsultaSql = @"select id, folio from tblXmlPdf where folio=@folio;";

                    sqlConexion.Open();

                    sqlComando2 = new SqlCommand(sConsultaSql, sqlConexion);

                    sqlComando2.Parameters.AddWithValue("folio", leerXml.sFolio);

                    sqlLeer = sqlComando2.ExecuteReader();

                    while (sqlLeer.Read())
                    {
                        cGlobal.iDatoLeidoXmlArticulos = Convert.ToInt32(sqlLeer.GetValue(0));
                    }

                    sqlLeer.Close();

                    sqlConexion.Close();

                    if (bActualizar)
                    {
                        string sArticulos;
                        int iArticulosActualizar;
                        SqlDataReader sqlActualizar;
                        SqlCommand sqlComandoBuscar;
                        SqlCommand sqlComandoActualizar;
                        SqlConnection sqlConexionActualizar = new SqlConnection(cGlobal.sCadenaSql);

                            sArticulos = @"select id from tblArticulosXmlPdf where idXml=@idXml";

                            sqlConexionActualizar.Open();

                            sqlComandoBuscar = new SqlCommand(sArticulos, sqlConexionActualizar);

                            sqlComandoBuscar.Parameters.AddWithValue("idXml", cGlobal.iDatoLeidoXmlArticulos);

                            sqlActualizar = sqlComandoBuscar.ExecuteReader();

                            sqlConexion.Open();

                            f = 0;

                            while (sqlActualizar.Read() & f < lnoIdentificacion.Count())
                            {
                                iArticulosActualizar = Convert.ToInt32(sqlActualizar.GetValue(0));

                                sInsertarDatos2 = @"update tblArticulosXmlPdf set idXml=@idXml, herco=@herco, codigo=@codigo, 
                                        cantidad=@cantidad, unidadMedida=@unidadMedida, descripcion=@descripcion, 
                                        precioUnitario=@precioUnitario, importe=@importe 
                                        where id=@iArticulosActualizar;";

                                sqlComandoActualizar = new SqlCommand(sInsertarDatos2, sqlConexion);

                                sqlComandoActualizar.Parameters.AddWithValue("iArticulosActualizar", ((object)iArticulosActualizar) ?? DBNull.Value);
                                sqlComandoActualizar.Parameters.AddWithValue("idXml", ((object)cGlobal.iDatoLeidoXmlArticulos) ?? DBNull.Value);
                                sqlComandoActualizar.Parameters.AddWithValue("herco", ((object)lnoIdentificacion[f]) ?? DBNull.Value);
                                sqlComandoActualizar.Parameters.AddWithValue("codigo", ((object)lnoIdentificacion[f]) ?? DBNull.Value);
                                sqlComandoActualizar.Parameters.AddWithValue("cantidad", ((object)lcantidad[f]) ?? DBNull.Value);
                                sqlComandoActualizar.Parameters.AddWithValue("unidadMedida", ((object)lunidad[f]) ?? DBNull.Value);
                                sqlComandoActualizar.Parameters.AddWithValue("descripcion", ((object)lDescripcion[f]) ?? DBNull.Value);
                                sqlComandoActualizar.Parameters.AddWithValue("precioUnitario", ((object)Convert.ToDouble(lvalorUnitario[f])) ?? DBNull.Value);
                                sqlComandoActualizar.Parameters.AddWithValue("importe", ((object)Convert.ToDouble(limporte[f])) ?? DBNull.Value);

                                sqlComandoActualizar.ExecuteNonQuery();

                                f=f+1;
                            }

                            sqlActualizar.Close();

                            sqlConexionActualizar.Close();

                            sqlConexion.Close();
                    }
                    else
                    {
                        for (int f = 0; f < lnoIdentificacion.Count(); f++)
                        {
                            sInsertarDatos2 = @"Insert into tblArticulosXmlPdf (idXml, herco, codigo, 
                                        cantidad, unidadMedida, descripcion, precioUnitario, importe) 
                                        Values(@idXml, @herco, @codigo, 
                                        @cantidad, @unidadMedida, @descripcion, @precioUnitario, @importe);";


                            sqlConexion.Open();

                            sqlComando2 = new SqlCommand(sInsertarDatos2, sqlConexion);

                            sqlComando2.Parameters.AddWithValue("idXml", ((object)cGlobal.iDatoLeidoXmlArticulos) ?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("herco", ((object)lnoIdentificacion[f]) ?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("codigo", ((object)lnoIdentificacion[f]) ?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("cantidad", ((object)lcantidad[f]) ?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("unidadMedida", ((object)lunidad[f]) ?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("descripcion", ((object)lDescripcion[f]) ?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("precioUnitario", ((object)Convert.ToDouble(lvalorUnitario[f])) ?? DBNull.Value);
                            sqlComando2.Parameters.AddWithValue("importe", ((object)Convert.ToDouble(limporte[f])) ?? DBNull.Value);

                            sqlComando2.ExecuteNonQuery();

                            sqlConexion.Close();
                        }
                    }

                    bGuardarXml = true;

                    cGlobal.iTipoDocumento = config.iTipoDocumento;

                    if (config.iTipoDocumento == 7)
                    {
                        frmReporteNotaElectronica reporte = new frmReporteNotaElectronica();
                        reporte.Show();
                    }
                    else
                    {

                        frmReporteFactura reporte = new frmReporteFactura();
                        reporte.Show();
                    }

                    txtMensajesPdf.Text = "DOCUMENTO PDF GENERADO EXITOSAMENTE";
                    txtMensajesPdf.BackColor = Color.Green;
                    txtMensajesPdf.ForeColor = Color.White;
                    txtMensajesPdf.Update();

                    txtOrdenPdf.Text = "";
                    txtCodigoPdf.Text = "";
                    txtTelefonoReceptor.Text = "";
                    txtDetallePDF.Text = "";
                
            }
            catch (Exception ex)
            {
                txtMensajesPdf.Text = "";
                txtMensajesPdf.BackColor = Color.Silver;
                txtMensajesPdf.Update();

                bGuardarXml = false;
                //limpiarMensaje();
                MessageBox.Show("Error en el proceso de guardar, motivo: " + ex.Message, "Error ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        
        }

#endregion


              
 

    }

    
}

factura electronica
