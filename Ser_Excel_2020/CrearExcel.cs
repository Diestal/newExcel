using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;
using System.Windows;
using OfficeOpenXml;//Librería EPPlus
using OfficeOpenXml.Table;
using OfficeOpenXml.Drawing;
using System.Drawing;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;//Librería de Excel

namespace Ser_Excel_2020
{
    class CrearExcel
    {
        #region VARIABLES LIBRERÍA EPPLUS
        private ExcelPackage AppExcel;//Librería EPPlus - Instancia para crear el Excel
        private ExcelWorksheet hojaXls;
        #endregion

        #region VARIABLES EXCEL
        private Excel.Application objExcel;
        private Excel.Workbook objLibro;
        #endregion

        #region VARIABLES CLASES
        private LOG log;
        private Datos CargaDatos;
        #endregion

        #region VARIABLES LISTAS-ARRAYS
        private List<string> vechojas, vecActual;
        private List<String> vecDatos;
        private ArrayList variable;
        #endregion

        #region VARIABLES FICHEROS
        private FileInfo existeArchivo;
        #endregion

        #region VARIABLES GENERALES
        private string RutaAplicacion = System.AppDomain.CurrentDomain.BaseDirectory.ToString();
        private string nomarchivo, numsesion, rutaReporte, rutaReporte1, var;
        private string rutaPlantilla, hojaActual, rep, tipoReporte, hojaActual1, hojaActualindicador;
        private string[] param_req;
        private char sep;
        private string rutaPlantilla1, indicador, nomHojaPrincipal, nomplantilla, pass, Archivoxls;
        //private string[] vecParam = new string[3];
        private int i, j, k, l, m;
        private int conterr = 0, zoom = 0, orie, tama;
        private int numCol;
        private long posIndi, f1, f2;
        private bool existe, cambio_numInd, existeindicador;
        private string variable1, myStr;
        private string utilita;
        private int posicion;
        private string linea_convierte;
        #endregion

        public CrearExcel()
        {
            string rtaLog = RutaAplicacion.Replace(@"SIIFNET\", "Documentos\\LOGS\\");
            CargaDatos = new Datos();
            log = new LOG(rtaLog, "Ser_Excel");
        }

        //--------------Metodo principal-------------------
        //public void crear_reporte(List<String> param, string ruta_Timelog)
        public void crear_reporte(string[] param, string ruta_Timelog)
        {
            try { 
            DateTime Hora;
            Random numale;

            indicador = "??FIN??";
            sep = (char)1;
            posIndi = 1;
            //vecParam = param.Split(sep);

            //Vector para almacenar las hojas que se encuentran en la Plantilla a excepción de la PRINCIPAL
            vechojas = new List<string>();

            //param[0]= @"C:\Users\wilder.lopez\Documents\Requerimientos\2021\Septiembre\75813\Ser excel Nuevo\Ser_Excel_2020\Ser_Excel_2020\bin\Debug\Ser_Excel.exe"; //Ruta del ser_excel -> C:\Users\xxxx.xxxx\source\repos\Ejecuta_Ser_Excel\Ejecuta_Ser_Excel\bin\Debug\Ser_Excel.exe
            //param[0]= "P000017188" + sep+"57632";

            param_req = param[1].Split(sep);

            //if (param.Count > 1)
            if (param_req.Length > 1)
            {
                nomarchivo = param_req[0].Substring(1);
                utilita = param_req[0].Substring(0, 1);
                numsesion = param_req[1];//Ej 44123

                /*---------DEFINE LA RUTA DEL REPORTE PARA EXTRAER LA INFORMACIÓN-------------*/
                //ruta 23 Ej -> SIIF_WWB\AYUDAS\PAGINAS\reportes\
                rutaReporte = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[23] + nomarchivo + ".txt";

                //VERIFICAR SI EXISTE EL REPORTE *.txt
                if (File.Exists(rutaReporte))
                {
                    //Obtiene las líneas del Reporte(*.txt)
                    vecDatos = LeerReporte(rutaReporte);
                    if (vecDatos.Count > 1)
                    {
                        //DIVIDE LA PRIMERA LÍNEA - REFERENTE A LA PLANTILLA A UTILIZAR
                        //---Ej: Primera linea -> WWB033,H,011
                        vecActual = vecDatos[0].Split(sep).ToList();

                        //TOMA EL NOMBRE DE PLANTILLA EJ: WWB033
                        nomplantilla = vecActual[0];

                        //RUTA DONDE ESTA LA PLANTILLA Y TOMA LA PLANTILLA EN EXCEL
                        //---RUTAS[3] -> Ej: SIIF_NOMINA\PLANTILLAS\
                        rutaPlantilla1 = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[3].ToString() + nomplantilla + ".xlsx";

                        numale = new Random();

                        //UBICACIÓN DE ARCHIVO TEMPORAL EN RUTAS 1 -> DOCUMENTOS\ARCHIVOS
                        rutaPlantilla = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "Temp" + numale.Next(1000, 9999) + ".xlsx";

                        //NOMBRE DEL EXCEL CONSOLIDADO
                        Archivoxls = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "E" + numsesion + ".xls";

                        //DETERMINA TIPO DE REPORTE, QUE VA EN LA PRIMERA LÍNEA DEL REPORTE TXT
                        tipoReporte = vecActual[1]; 

                        //TOMA EL NÚMERO QUE EQUIVALE A LA ULTIMA COLUMNA DONDE SE ENCUENTRAN DATOS REFERENTE A LA LÍNEA DE LA PLANTILLA
                        ////Ej: Primera linea -> WWB033,H,011<----
                        numCol = int.Parse(vecActual[2]);//SE PRETENDE QUITAR

                        try
                        {
                            File.Delete(rutaPlantilla);
                            //PREGUNTA SI EL REPORTE A GENERAR TIENE COMO NOMBRE LA SECUENCIA
                            if (param_req[1] == "PRRE")
                                File.Delete(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1] + "E" + nomarchivo + ".xlsx");
                            else
                                File.Delete(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1] + "E" + param_req[1] + ".xlsx");
                        }
                        catch { }

                        
                        /*COPIA CONTENIDO DEL ARCHIVO-PLANTILLA EN EL ARCHIVO TEMPORAL,
                         * ESTO CON EL FIN DE MANEJAR DIFERENTES INSTANCIAS DE USUARIOS
                         */
                        try
                        {
                            File.Copy(rutaPlantilla1, rutaPlantilla);
                        }
                        catch (Exception ex)
                        {
                            if (File.Exists(rutaPlantilla))
                            {
                                File.Delete(rutaPlantilla);
                            }
                            log.EscribeLog("Error al copiar: " + rutaPlantilla1 + " como: " + rutaPlantilla, " -- " + ex.ToString(), true);
                        }

                        //Si existe archivo temporal TempXXXX.xlsx
                        if (File.Exists(rutaPlantilla))
                        {
                            existeArchivo = Utilidades.ObtieneInfoArchivo(rutaPlantilla, false);//false -> si no se desea eliminar el archivo

                            //INSTANCIA PRINCIPAL PARA LECTURA DE ARCHIVO XLSX POR MEDIO DE LA LIBRERIA
                            //y SE LIBERE EL OBJETO APENAS CULMINE EL MERCHADO DE LA INFORMACIÓN
                            using (AppExcel = new ExcelPackage(existeArchivo))
                            {
                                try
                                {
                                    
                                    if (tipoReporte == "I")
                                    {
                                        //COPIA EL FORMATO DE LA HOJA PRINCIPAL A UNA NUEVA HOJA LLAMADA PRINCIPALAUX
                                        hojaXls = AppExcel.Workbook.Worksheets.Copy(AppExcel.Workbook.Worksheets.First().ToString(), "PrincipalAux");
                                        hojaActual = "PrincipalAux";
                                        nomHojaPrincipal = AppExcel.Workbook.Worksheets[1].Name;
                                        m += 1;
                                        vechojas.Add(hojaActual.ToString());
                                    }
                                }
                                catch (Exception ex)
                                {
                                    if (File.Exists(rutaPlantilla))
                                    {
                                        File.Delete(rutaPlantilla);
                                    }
                                    log.EscribeLog("Error copiando tipo de reporte I de hoja Principal a hoja nueva PrincipalAux", ex.ToString(), true);
                                }
                                try
                                {

                                    //RECORRE EL REPORTE LÍNEA POR LÍNEA
                                    //No cuenta la primera posición, porque no hay hoja a leer
                                    //Ej: FINE088-H-099
                                    for (int i = 1; i < vecDatos.Count - 1; i++)//Para no tomar la última línea que viene vacía
                                    {
                                        int posicion;
                                        posicion = vecDatos[i].IndexOf(sep);

                                        if (posicion == -1)
                                        {
                                            //Se Agregó un Separador para Proteger la Integridad del Programa
                                            vecDatos[i] = vecDatos[i] + sep;
                                        }
                                        
                                        //TOMA FILA DEL REPORTE REFERENTE A LA CABECERA DEL DOCUMENTO
                                        vecActual = vecDatos[i].Split(sep).ToList(); //Línea actual

                                        //NOMBRE DE HOJA A POSICIONAR EN LA HOJA PRINCIPAL
                                        hojaActual = vecActual[0];
                                        
                                        existe = false;

                                        //********************************CONFIGURAR ENCABEZADO DE PÁGINA**********************************
                                        if(hojaActual == "998" && i == vecDatos.Count - 2)
                                        {
                                            //ACTIVA HOJA PRINCIPAL
                                            AppExcel.Workbook.Worksheets.First();
                                            
                                            string cab2 = hojaXls.HeaderFooter.OddHeader.CenteredText;
                                            for(int j=1;j<vecActual.Count;j++)//Ciclo Merge
                                            {
                                                var = "<VAR" + j.ToString("000") + ">";//VAR001, VAR002 .... VAR005
                                                cab2 = cab2.Replace(var, vecActual[j]);
                                            }
                                            try
                                            {
                                                hojaXls.HeaderFooter.OddHeader.CenteredText = cab2;
                                            }
                                            catch(Exception ex)
                                            {
                                                if (File.Exists(rutaPlantilla))
                                                {
                                                    File.Delete(rutaPlantilla);
                                                }
                                                log.EscribeLog("Error al asignar Encabezado de la Página " , ex.ToString(), true);
                                            }

                                        }
                                        //********************************CONFIGURAR PIE DE PÁGINA**********************************
                                        if(hojaActual == "999" && i == vecDatos.Count - 1)
                                        {
                                            //ACTIVA HOJA PRINCIPAL
                                            AppExcel.Workbook.Worksheets.First();
                                            
                                            string cab2 = hojaXls.HeaderFooter.OddFooter.LeftAlignedText;
                                            for(int j=1;j<vecActual.Count;j++)//Ciclo Merge
                                            {
                                                var = "<VAR" + j.ToString("000") + ">";//VAR001, VAR002 .... VAR005
                                                cab2 = cab2.Replace(var, vecActual[j]);
                                            }
                                            try
                                            {
                                                hojaXls.HeaderFooter.OddFooter.LeftAlignedText = cab2;
                                            }
                                            catch(Exception ex)
                                            {
                                                if (File.Exists(rutaPlantilla))
                                                {
                                                    File.Delete(rutaPlantilla);
                                                }
                                                log.EscribeLog("Error al asignar el Pie de Página " , ex.ToString(), true);
                                            }
                                            break;//Para salir del bucle de vecdatos, ya que termina el Reporte y no tiene hoja en la Plantilla

                                        }
                                        
                                        //****************************************************************************************************
                                        //------------------->OBTENIENDO NÚMERO DE HOJAS DEL ACTUAL LIBRO EXCEL PARA NO REPETIR HOJA----------
                                        //****************************************************************************************************
                                        for (int l = 0; l < m; l++)
                                        {
                                            if (vechojas[l] == hojaActual)
                                            {
                                                existe = true;
                                                break;
                                            }
                                        }
                                        if (!(existe))
                                        {
                                            //vechojas[m] = hojaActual1;
                                            m += 1;
                                            vechojas.Add(hojaActual.ToString());
                                        }

                                        try
                                        {
                                            string IndicaCelda = "";
                                            int IndicaNumFila = 0;
                                            hojaXls = AppExcel.Workbook.Worksheets[hojaActual];
                                            
                                            //BUSCA EL INDICADOR ??FIN?? EN LA COLUMNA A
                                            var celdaIndicadorHoja = (from celda in hojaXls.Cells["A:A"] where celda.Value is "??FIN??" select celda);
                                            foreach (var celda in celdaIndicadorHoja)
                                            {
                                                IndicaCelda = celda.Address;//Celda dónde se encontró el indicador Ej: A7
                                                IndicaNumFila = int.Parse(celda.Address.Replace("A", "")); //Se quita el A(columna) para dejar solo la fila
                                                break;//SE PODRÍA QUITAR
                                            }

                                            if (celdaIndicadorHoja.Count() > 0)
                                            {
                                                    /*SE VERIFICA SI LA HOJA ACTUAL TIENE IMAGEN*/
                                                    string Indimg = "";
                                                    hojaXls = AppExcel.Workbook.Worksheets[hojaActual];
                                                    //BUSCA EL INDICADOR ??FIN?? EN LA COLUMNA A
                                                    var celdaIndimg = (from celda in hojaXls.Cells["A:A"] where celda.Value is "??IMG??" select celda);
                                                    if (celdaIndimg.Count() > 0 )
                                                    {
                                                    
                                                       
                                                        
                                                            string IndicaCelda_HojaPrincipal = "";
                                                            int IndicaNumFilaImagen = 0;
                                                            int fincolumna = 0;
                                                            hojaXls = AppExcel.Workbook.Worksheets[1];//Se selecciona hoja PRINCIPAL
                                                    
                                                            //Busca el indicador ??FIN?? en la columna A
                                                            var celdaIndicadorHojaA = (from celda in hojaXls.Cells["A:A"] where celda.Value is "??FIN??" select celda);
                                                            foreach (var celda in celdaIndicadorHojaA)
                                                            {
                                                                IndicaCelda_HojaPrincipal = celda.Address;//Celda dónde se encontró el indicador
                                                                IndicaNumFilaImagen = int.Parse(celda.Address.Replace("A", "")); //Se quita el A(columna) para dejar solo la fila
                                                                fincolumna = hojaXls.Dimension.Columns;//Da el número de la columna
                                                                break;
                                                            }

                                                        if (celdaIndicadorHojaA.Count() > 0)
                                                        {

                                                            StreamReader archParam = new StreamReader(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[3].ToString() + nomplantilla + ".txt");
                                                            ArrayList datosParamtxt = new ArrayList();
                                                            while (!archParam.EndOfStream)
                                                            {
                                                                datosParamtxt.Add(archParam.ReadLine());
                                                            }
                                                            archParam.Close();
                                                            releaseObject(archParam);


                                                            hojaXls = AppExcel.Workbook.Worksheets[hojaActual];


                                                            //COPIA LAS CELDAS QUE CONTIENEN LAS VARIABLES DE LA HOJA ACTUAL A LA HOJA PRINCIPAL, SIN UTILIZAR EL PORTAPAPELES
                                                            hojaXls.Cells[1, 1, IndicaNumFila, fincolumna].Copy(AppExcel.Workbook.Worksheets.First().Cells[IndicaCelda_HojaPrincipal]);

                                                            //SE SELECCIONA HOJA PRINCIPAL PARA REALIZAR EL MERCHADO
                                                            hojaXls = AppExcel.Workbook.Worksheets.First();
                                                            AppExcel.Workbook.Worksheets.First().Select();

                                                            //j=1 -> Para no tomar el indicador de la hoja 001, 002, 003, 998 en la línea Actual
                                                           
                                                            
                                                            for (int y = 0; y < datosParamtxt.Count; y++) {
                                                                string[] datosParam;
                                                                datosParam = datosParamtxt[y].ToString().Split('|');
                                                                ////SE ASIGNA EL NOMBRE DE LA PLANTILLA A LA IMAGEN

                                                                /*------------------DETERMINA PARAMETROS PARA EL TAMAÑO Y UBICACION DE LA IMAGEN-------------------------------------------------*/
                                                                if (datosParam.Length <= 0)
                                                                {
                                                                    if (File.Exists(rutaPlantilla))
                                                                    {
                                                                        File.Delete(rutaPlantilla);
                                                                    }
                                                                    log.EscribeLog("El archivo de Parametros para la IMAGEN esta vacio: " + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[3].ToString() + nomplantilla + ".txt", "", true);
                                                                }
                                                                else if (hojaActual==datosParam[0])
                                                                    {
                                                                        

                                                                        if (datosParam.Length == 7)
                                                                        {
                                                                            Image img = Image.FromFile(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[3].ToString() + datosParam[1] + ".png");

                                                                            //Imagen + i -> Nombre del objeto dentro del Excel en el CUADRO DE NOMBRES + i -> por si hay mas imagenes

                                                                            var excelImage = hojaXls.Drawings.AddPicture("Imagen" + y, img);
                                                                            excelImage.SetPosition(IndicaNumFilaImagen + Int32.Parse(datosParam[6].ToString()),                   //Fila dónde va ubicada la IMAGEN
                                                                                                    int.Parse(datosParam[2].ToString()),  //Altura de la IMAGEN en píxeles
                                                                                                    int.Parse(datosParam[3].ToString()),  //Columna dónde va ubicada la IMAGEN
                                                                                                    int.Parse(datosParam[4].ToString())); //Ancho de la IMAGEN en píxeles
                                                                            excelImage.SetSize(int.Parse(datosParam[5].ToString()));//Tamaño porcentual sin el porcentaje, se va a trabajar en un archivo plano en PLANTILLAS
                                                                        }
                                                                        else
                                                                        {
                                                                            if (File.Exists(rutaPlantilla))
                                                                            {
                                                                                File.Delete(rutaPlantilla);
                                                                            }
                                                                            log.EscribeLog("No existen todos los parametros para usar la IMAGEN, en el archivo: " + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[3].ToString() + nomplantilla + ".txt", "", true);
                                                                        }
                                                                    
                                                                    }
                                                                    else
                                                                    {
                                                                        if (File.Exists(rutaPlantilla))
                                                                        {
                                                                            File.Delete(rutaPlantilla);
                                                                        }
                                                                        log.EscribeLog("No existe IMAGEN: " + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[3].ToString() + nomplantilla + ".png", "", true);
                                                                }

                                                                    
                                                            
                                                            }
                                                            

                                                        }
                                                        
                                                    }//vecActual -> IMAGEN

                                                else
                                                {
                                                    string IndicaCelda_HojaPrincipal = "";
                                                    int IndicaNumFila_HojaPrincipal = 0;
                                                    //Se selecciona Hoja Principal

                                                    hojaXls = AppExcel.Workbook.Worksheets.First();//o se puede tipo vector [1] sin el First()
                                                    //Busca el indicador ??FIN?? en la columna A
                                                    var celdaIndicadorHojaA = (from celda in hojaXls.Cells["A:A"] where celda.Value is "??FIN??" select celda);
                                                    foreach (var celda in celdaIndicadorHojaA)
                                                    {
                                                        IndicaCelda_HojaPrincipal = celda.Address;//Celda dónde se encontró el indicador
                                                        IndicaNumFila_HojaPrincipal = int.Parse(celda.Address.Replace("A", "")); //Se quita el A(columna) para dejar solo la fila
                                                    }

                                                    if(celdaIndicadorHojaA.Count() > 0)
                                                    {
                                                        hojaXls = AppExcel.Workbook.Worksheets[hojaActual];
                                                        var finCol = hojaXls.Dimension.Columns;//Fin Celda

                                                        //COPIA LAS CELDAS QUE CONTIENEN LAS VARIABLES DE LA HOJA ACTUAL A LA HOJA PRINCIPAL, SIN UTILIZAR EL PORTAPAPELES
                                                        hojaXls.Cells[1, 1, IndicaNumFila, finCol].Copy(AppExcel.Workbook.Worksheets.First().Cells[IndicaCelda_HojaPrincipal]);

                                                        //SE SELECCIONA HOJA PRINCIPAL PARA REALIZAR EL MERCHADO
                                                        hojaXls = AppExcel.Workbook.Worksheets.First();
                                                        AppExcel.Workbook.Worksheets.First().Select();

                                                        //j=1 -> Para no tomar el indicador de la hoja 001, 002, 003, 998 en la línea Actual
                                                        for (int j = 1; j < vecActual.Count; j++)
                                                        {
                                                            var = "<VAR" + j.ToString("000") + ">";//VAR001, VAR002 .... VAR005

                                                                bool tempVarRemp = false;       //BUSCA VARIABLE PARA REALIZAR MERCHADO
                                                                do { 
                                                                    var reemplazaVAR = from celdaVAR in hojaXls.Cells["A:XFD"]
                                                                                       where celdaVAR.Value?.ToString().Contains(var) == true
                                                                                       select celdaVAR;
                                                                
                                                                    if (reemplazaVAR.Count()>0)
                                                                      {
                                                                            tempVarRemp = true;
                                                                    }
                                                                    else
                                                                    {
                                                                        tempVarRemp = false;
                                                                    }
                                                                    foreach (var celdaVAR in reemplazaVAR)
                                                                    {
                                                                        //VERIFICA SI LA CELDA TIENE FORMATO DE TEXTO ENRIQUECIDO (RTF)
                                                                        if (celdaVAR.IsRichText)
                                                                        {
                                                                            celdaVAR.IsRichText = false;//Se desactiva el formato enriquecido, dejandolo como XML
                                                                            celdaVAR.Value = celdaVAR.Value.ToString().Replace(var.Replace("<", "&lt;").Replace(">", "&gt;"), vecActual[j]);//Como se desactiva el formato RTF, los símbolos de < y > toman el código en formato HTML
                                                                            celdaVAR.IsRichText = true;//Se vuelve activar el formato RTF para dejar la celda como estaba antes
                                                                        }
                                                                        else
                                                                        {
                                                                                //SE MERCHA LA INFORMACIÓN
                                                                                celdaVAR.Value = celdaVAR.Value.ToString().Replace(var, vecActual[j]);
                                                                                //celdaVAR.Value = vecActual[j];

                                                                                //VALIDA SI LA CELDA TIENE CONTENIDO NUMÉRICO
                                                                                float num;
                                                                            bool validasiesnum = false;
                                                                            validasiesnum = float.TryParse(celdaVAR.Value.ToString(), out num);
                                                                            if (validasiesnum)
                                                                            {
                                                                                celdaVAR.Value = decimal.Parse(celdaVAR.Value.ToString());
                                                                            }

                                                                        }

                                                                        break;
                                                                    }
                                                                } while (tempVarRemp);
                                                            }//Bucle Merchar Variables

                                                            
                                                        }
                                                    else
                                                    {
                                                        if (File.Exists(rutaPlantilla))
                                                        {
                                                            File.Delete(rutaPlantilla);
                                                        }
                                                        log.EscribeLog("Indicador vacío en la Hoja [Principal] --> en la línea: (" + (i + 1) + ") que contiene " + vecDatos[i], "", true);//true -> Para finalizar el programa
                                                    }

                                                }
                                            }
                                            else
                                            {
                                                if (File.Exists(rutaPlantilla))
                                                {
                                                    File.Delete(rutaPlantilla);
                                                }
                                                log.EscribeLog("Error: No se encuentra el indicador en la Hoja Actual - [" + hojaActual + "] --> en la línea: (" + (i + 1) + ") que contiene " + vecDatos[i], "", true);//true -> Para finalizar el programa
                                            }

                                        }
                                        catch(Exception ex)
                                        {
                                            if (File.Exists(rutaPlantilla))
                                            {
                                                //File.Delete(rutaPlantilla);
                                            }
                                            log.EscribeLog("Error en la Hoja Actual - [" + hojaActual + "] --> en la línea: (" + (i + 1) + ") que contiene " + vecDatos[i] + "--", ex.ToString(), true);//true -> Para finalizar el programa
                                        }

                                    }//Bucle VecDatos
                                    
                                }
                                catch (Exception ex)
                                {
                                    if (File.Exists(rutaPlantilla))
                                    {
                                        File.Delete(rutaPlantilla);
                                    }
                                    log.EscribeLog("Error al recorrer vector VecDatos, información contenida en el reporte. Línea Reporte: (" + (i + 1) + ")", ex.ToString(), true);
                                }

                                hojaXls = AppExcel.Workbook.Worksheets.First();
                                hojaXls.Select("A1");
                                
                                //ELIMINA INDICADOR DE LA HOJA PRINCIPAL
                                var consultaindicador = (from celda in hojaXls.Cells["A:A"] where celda.Value is "??FIN??" select celda);
                                foreach(var eliminaindi in consultaindicador)
                                {
                                    hojaXls.Cells[eliminaindi.Address].Value = null;
                                }


                                //SE ELIMINAN HOJAS EXCEPTO LA PRINCIPAL
                                foreach (string h in vechojas)
                                {
                                    AppExcel.Workbook.Worksheets.Delete(h);
                                    //hojaXls = AppExcel.Workbook.Worksheets[h];
                                    //hojaXls.Hidden = eWorkSheetHidden.VeryHidden;
                                }

                                AppExcel.Compression = CompressionLevel.BestSpeed;
                               
                                //SE GUARDA EL ARCHIVO FINAL
                                var ArchivoExcel = Utilidades.ObtieneInfoArchivo(rutaPlantilla, false);
                                AppExcel.SaveAs(ArchivoExcel);

                                //------------------------CONVERTIR DE XLSX A XLS POR MICROSOFT EXCEL--------------------------------
                                try
                                {
                                    //SE VERIFICA SI EXCEL ESTA INSTALADO
                                    bool isExcelInstalled = Type.GetTypeFromProgID("Excel.Application") != null ? true : false;
                                    if (isExcelInstalled)
                                    {
                                        objExcel = new Excel.Application();
                                        objLibro = objExcel.Workbooks.Open(rutaPlantilla);
                                        if(File.Exists(Archivoxls))
                                        {
                                            File.Delete(Archivoxls);
                                        }
                                        objExcel.Application.DisplayAlerts = false;//Para que no muestre anuncios del programa
                                        objLibro.SaveAs(Archivoxls,Excel.XlFileFormat.xlWorkbookNormal);//Formato 97-2003 (.xls)
                                        File.Delete(rutaPlantilla);//Se elimina archivo temporal
                                        objLibro.Close();
                                        releaseObject(objLibro);
                                        releaseObject(objExcel);
                                        log.EscribeLog("Archivo generado exitosamente !!! --> [ " + Archivoxls + " ]");
                                    }
                                    else
                                    {
                                        if (File.Exists(rutaPlantilla))
                                        {
                                            File.Delete(rutaPlantilla);
                                        }
                                        log.EscribeLog("El programa Excel no esta instalado --> [ " + Archivoxls + " ]");
                                    }

                                }
                                catch(Exception ex)
                                {
                                    if (File.Exists(rutaPlantilla))
                                    {
                                        File.Delete(rutaPlantilla);
                                    }
                                    log.EscribeLog(" Error al pasar de xlsx a xls ", ex.ToString(), true);

                                }

                            }//using AppExcel

                        }//Si existe TempXXXX.xlsx
                        else
                        {
                            if (File.Exists(rutaPlantilla))
                            {
                                File.Delete(rutaPlantilla);
                            }
                            log.EscribeLog("La Plantilla: " + rutaPlantilla + " No Existe", "", true);
                        }
                    }//VecDatos, si no esta vacío
                    else
                    {
                        if (File.Exists(rutaPlantilla))
                        {
                            File.Delete(rutaPlantilla);
                        }
                        log.EscribeLog("Reporte: [" + rutaReporte + ".txt] vacío", "", true);//true -> Termina el programa
                    }

                }//Reporte txt si existe
                else
                {
                    if (File.Exists(rutaPlantilla))
                    {
                        File.Delete(rutaPlantilla);
                    }
                    log.EscribeLog("No existe archivo Reporte: [" + rutaReporte + ".txt]", "", true);//true -> Termina el programa
                }

                //***********************Prin_Pdf_OpenOffice***********************
                try
                {

                    if(utilita == "P")
                    {
                            if (File.Exists(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "Temp_E" + numsesion + ".xls")) {
                                File.Delete(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "Temp_E" + numsesion + ".xls");
                            }
                        File.Copy(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "E" + numsesion + ".xls", CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "Temp_E" + numsesion + ".xls");
                            if (File.Exists(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "E" + numsesion + ".xls")) {
                                File.Delete(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "E" + numsesion + ".xls");
                            }
                            StreamWriter wr = new StreamWriter(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "PDF" + numsesion + ".txt", true);
                            linea_convierte = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "Temp_E" + numsesion + ".xls" + "-*-" + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "E" + numsesion + ".pdf";
                            wr.WriteLine(linea_convierte);
                            wr.Flush();
                            wr.Close();
                            wr.Dispose();
                            wr.Dispose();
                            //wr.Close();
                            //releaseObject(wr);
                            //Process.Start(CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[2] + "prin_pdf_OpenOffice.exe", "1");
                            ProcessStartInfo start = new ProcessStartInfo();
                            start.FileName = "prin_pdf_OpenOffice.exe";
                            start.Arguments = numsesion;
                            start.UseShellExecute = false;// Do not use OS shell
                            start.CreateNoWindow = true; // We don't need new window
                            start.RedirectStandardOutput = true;// Any output, generated by application will be redirected back
                            start.RedirectStandardError = true; // Any error in standard output will be redirected back (for example exceptions)
                            start.WorkingDirectory = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[2];
                            Process.Start(start);

                            log.EscribeLog("Generando pdf: LLamando a Prin_Pdf_OpenOffice.exe con parametro de archivo PDFXXXX.txt : " + numsesion,"",true);
                    }
                }
                catch(Exception ex)
                {
                    if (File.Exists(rutaPlantilla))
                    {
                        File.Delete(rutaPlantilla);
                    }
                    log.EscribeLog("Error al ejecutar : [" + CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[2] + "prin_pdf_OpenOffice.exe ]", ex.ToString(), true);//true -> Termina el programa
                }
                //******************************************************************
            
            }//VecParam si viene vacío
            else
            {
                if (File.Exists(rutaPlantilla))
                {
                    File.Delete(rutaPlantilla);
                }
                log.EscribeLog("Faltan parametros", "", true);//true -> Termina el programa
            }
            //Controlador de tiempo en generar el reporte
            Hora = DateTime.Now;
            StreamWriter timelog = new StreamWriter(ruta_Timelog, true, Encoding.UTF8);
            timelog.WriteLine("***************************************************************************************");
            timelog.WriteLine("Hora de Finalización:    " + Hora.ToString("F"));//F - Friday, February 27, 2009 12:12:22 PM;
            timelog.Flush();
            timelog.Close();
            System.Environment.Exit(0);
            }
            catch (Exception e)
            {
                log.EscribeLog("Error: ", " -- " + e.ToString(), true);

            }

        }//Cierre Metodo CrearReporte

        //SE LEE REPORTE TXT Y SE CUENTA POR SALTO DE LÍNEA
        private List<String> LeerReporte(string rutaReporte)
        {
            string ContenidoRepor = "";
            List<String> vecReporte;
            StreamReader leeReporte = new StreamReader(rutaReporte);
            ContenidoRepor = leeReporte.ReadToEnd();
            leeReporte.Close();
            vecReporte = ContenidoRepor.Split(new string[] { Environment.NewLine }, StringSplitOptions.None).ToList();
            return vecReporte;
        }

        //LIBERAR OBJETOS
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                log.EscribeLog("No es posible liberar el objeto --> " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        private static Boolean EsNumero(string valor)
        {
            int result;
            return int.TryParse(valor, out result);
        }
    }
}
