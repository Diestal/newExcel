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
using System.Threading;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Ser_Excel_2020
{
    public class CrearExcel
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
        private string nomarchivo, numsesion, rutaReporte, rutaReporte1, var, archivoXls;
        private string rutaPlantilla, hojaActual, rep, tipoReporte, hojaActual1, hojaActualindicador;
        private string[] param_req;
        private string auxDebug;
        private char sep;
        private string rutaPlantilla1, indicador, nomHojaPrincipal, nomplantilla, pass, Archivoxls;
        //private string[] vecParam = new string[3];
        private int i, j, k, l, m;
        //private int conterr = 0, zoom = 0, orie, tama;
        //private int numCol;
        //private long posIndi, f1, f2;
        private bool existe;//, cambio_numInd, existeindicador;
        //private string variable1, myStr;
        private string utilita;
        //private int posicion;
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
        public Task<string> crear_reporte(string[] param, string ruta_Timelog)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DateTime Hora;
            Random numale;
            indicador = "??FIN??";
            sep = (char)1;
            //posIndi = 1;
            //vecParam = param.Split(sep);
            //Vector para almacenar las hojas que se encuentran en la Plantilla a excepción de la PRINCIPAL
            vechojas = new List<string>();
            //param[0]: Ruta del ser_excel -> C:\Users\xxxx.xxxx\source\repos\Ejecuta_Ser_Excel\Ejecuta_Ser_Excel\bin\Debug\Ser_Excel.exe
            //param[1]: Parametros -> Y000116139#57632
            param_req = param[1].Split(sep);

            //if (param.Count > 1)
            if (param_req.Length > 1)
            {
                nomarchivo = param_req[0].Substring(1);
                utilita = param_req[0].Substring(0, 1);
                numsesion = param_req[1];//Ej 44123
                /*---------DEFINE LA RUTA DEL REPORTE PARA EXTRAER LA INFORMACIÓN-------------*/
                //ruta 23 Ej -> SIIF_WWB\AYUDAS\PAGINAS\reportes\
                rutaReporte = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[23].ToString().Trim() + nomarchivo + ".txt";
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
                        rutaPlantilla = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "Temp" + numale.Next(10000, 99999) + ".xlsx";
                        //NOMBRE DEL EXCEL CONSOLIDADO
                        Archivoxls = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "E" + numsesion + ".xls";
                        string NombreArchivoXls = "E" + numsesion + ".xls";
                        //DETERMINA TIPO DE REPORTE, QUE VA EN LA PRIMERA LÍNEA DEL REPORTE TXT
                        tipoReporte = vecActual[1];
                        //TOMA EL NÚMERO QUE EQUIVALE A LA ULTIMA COLUMNA DONDE SE ENCUENTRAN DATOS REFERENTE A LA LÍNEA DE LA PLANTILLA
                        ////Ej: Primera linea -> WWB033,H,011<----
                        //numCol = int.Parse(vecActual[2]);//SE PRETENDE QUITAR

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
                        catch (System.Exception ex)
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
                                catch (System.Exception ex)
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
                                    // Bottleneck -> Lugar del programa donde más se ocupa tiempo para hacer una operación
                                    // Main Bottleneck -> donde más memoria se ocupa para generar cada reporte
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
                                            catch(System.Exception ex)
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
                                            catch(System.Exception ex)
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
                                        // Bottleneck 3°
                                        if (!vechojas.Any(h => h == hojaActual))
                                        {
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
                                                // SE RETIRA LA FUNCIONALIDAD DONDE SIEMPRE REVISA QUE HAYA IMÁGENES, DEBIDO A QUE NO LAS MERCHABA DE TODAS MANERAS
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
                                                        auxDebug = IndicaCelda_HojaPrincipal.ToString() + " fila: " + IndicaNumFila.ToString() + " columna: " + finCol.ToString();
                                                        hojaXls.Cells[1, 1, IndicaNumFila, finCol].Copy(AppExcel.Workbook.Worksheets.First().Cells[IndicaCelda_HojaPrincipal]);

                                                        //SE SELECCIONA HOJA PRINCIPAL PARA REALIZAR EL MERCHADO
                                                        hojaXls = AppExcel.Workbook.Worksheets.First();
                                                        AppExcel.Workbook.Worksheets.First().Select();

                                                        //***********************||   GENERACIÓN DE REPORTE EN PDF   ||***********************
                                                        //j=1 -> Para no tomar el indicador de la hoja 001, 002, 003, 998 en la línea Actual
                                                        // Bottleneck 2° -> RECORRÍA UN CICLO POR CADA VARIABLE, LAS ITERACIONES SE AMUENTABAN -> NUM_REGISTROS ** NUM_VARIABLES
                                                        var var = "VAR";
                                                        // AHORA REVISA SI LOS CAMPOS CONTIENEN VAR PARA MERCHAR
                                                        var reemplazaVAR = from celdaVAR in hojaXls.Cells["A:XFD"]
                                                                           where celdaVAR.Value?.ToString().Contains(var) == true
                                                                           select celdaVAR;
                                                        // SE ALMACENAN LOS CAMPOS POR MERCHAR
                                                        var matchingCells = reemplazaVAR.ToList();
                                                        // SE DEJA UN INDICE PARA NO TOMAR EL NÚMERO DE LA HOJA
                                                        int indexOfCombined = 1;
                                                        if (matchingCells.Count > 0)
                                                        {
                                                        // TOMA TODO EL REGISTRO, EN LUGAR DE ITERAR POR ELEMENTOS DEL REGISTRO
                                                            var combinedValue = vecActual;
                                                            foreach (var cell in matchingCells)
                                                            {
                                                            // SI LA CELULA TIENE ESTILOS O ALGO MÁS, SE ELIMINA
                                                                if (cell.IsRichText)
                                                                {
                                                                // SE TIENE PRESENTE CADA VEZ QUE SE ELIMINA
                                                                    log.EscribeLog("Vars rtf: " + cell.Value);
                                                                    cell.IsRichText = false;
                                                                    cell.Value = cell.Value.ToString().Replace(var.Replace("<", "&lt;").Replace(">", "&gt;"), combinedValue[indexOfCombined]);
                                                                    cell.IsRichText = true;
                                                                    indexOfCombined++;
                                                                }
                                                                else
                                                                {
                                                                // SE MERCHA LA CELDA CON EL DATO
                                                                    cell.Value = combinedValue[indexOfCombined];
                                                                    indexOfCombined++;
                                                                    bool validasiesnum = false;
                                                                    if (cell.Value.ToString().IndexOf(",") != -1)
                                                                    {
                                                                    // SI TIENE UN INDICE NUMERICO, LO CONVIERTE PARA NO GENERAR EXCEPCION DE TIPO DE EPPLUS
                                                                        float num;
                                                                        validasiesnum = float.TryParse(cell.Value.ToString(), out num);
                                                                        if (validasiesnum)
                                                                        {
                                                                            cell.Value = float.Parse(cell.Value.ToString());
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        long num;
                                                                        validasiesnum = long.TryParse(cell.Value.ToString(), out num);
                                                                        if (validasiesnum)
                                                                        {
                                                                            cell.Value = long.Parse(cell.Value.ToString());
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
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
                                                File.Delete(rutaPlantilla);
                                            }
                                            log.EscribeLog("Error en la Hoja Actual - [" + hojaActual + "] --> en la línea: (" + (i + 1) + ") que contiene " + vecDatos[i].Length + ". La celda límite es: " + auxDebug + "--", ex.ToString(), true);//true -> Para finalizar el programa
                                        }
                                    }//Bucle VecDatos
                                }
                                catch (System.Exception ex)
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

                                AppExcel.Compression = CompressionLevel.BestSpeed;
                                //SE GUARDA EL ARCHIVO FINAL
                                var ArchivoExcel = Utilidades.ObtieneInfoArchivo(rutaPlantilla, false);
                                AppExcel.SaveAs(ArchivoExcel);
                                //------------------------CONVERTIR DE XLSX A XLS POR EPPLUS--------------------------------
                                try
                                {
                                    //SE VERIFICA SI EXCEL ESTA INSTALADO
                                    bool isExcelInstalled = Type.GetTypeFromProgID("Excel.Application") != null ? true : false;
                                    if (isExcelInstalled)
                                    {
                                        FileInfo plantilla = new FileInfo(rutaPlantilla);
                                        string archivoXls = Path.Combine(Path.GetDirectoryName(rutaPlantilla), NombreArchivoXls);
                                        // VERIFICA QUE NO HAYA NINGÚN ARCHIVO NOMBRADO IGUAL
                                        // -> CONSECUTIVO -> EJ. -> 1234
                                        // -> EXTENSION -> EJ. -> .xls
                                        if (File.Exists(archivoXls))
                                        {
                                            File.Delete(archivoXls);
                                        }
                                        // SE USA EPPLUS PARA CONVERTIR MÁS RÁPIDO EL DOCUMENTO DE XLSX A XLS
                                        ExcelPackage package = new ExcelPackage(plantilla);
                                            try
                                            {
                                                // ELIMINA TODAS LAS HOJAS EXCEPTO LA PRINCIPAL
                                                for (int i = package.Workbook.Worksheets.Count - 1; i > 0; i--)
                                                {
                                                    ExcelWorksheet worksheet = package.Workbook.Worksheets[i];
                                                    package.Workbook.Worksheets.Delete(worksheet);
                                                }
                                            }
                                            catch(Exception ex)
                                            {
                                                package.Dispose();
                                                log.EscribeLog("Error limpiando el reporte: " + ex.Message);
                                            }
                                        try
                                        {
                                            // 1) GUARDA EL ARCHIVO EN RUTA -> AMBIENTE/SIIF_AMBIENTE/DOCUMENTOS/ARCHIVOS/ARCHIVO.xls
                                            // 2) SE LIBERA LA MEMORIA DEL ARCHIVO CON Dispose
                                            // 3) SE ELIMINA EL ARCHIVO XLSS
                                            package.SaveAs(new FileInfo(archivoXls));
                                            package.Dispose();
                                            File.Delete(rutaPlantilla);
                                            log.EscribeLog("Archivo generado exitosamente !!! --> [ " + archivoXls + " ]");
                                        }
                                        catch (Exception ex)
                                        {
                                            // SOLO OCURRE SI HAY PROBLEMAS CON EL FORMATO, SI EL EXCEL ESTÁ CORRUPTO O SI SE ELIMINA ANTES EL TEMPORAL
                                            package.Dispose();
                                            log.EscribeLog("Error al guardar archivo: " + ex.Message);
                                        }
                                    }
                                }

                                catch (System.Exception ex)
                                {
                                    log.EscribeLog(" Error al pasar de xlsx a xls ", ex.ToString(), true);
                                    if (File.Exists(rutaPlantilla))
                                    {
                                        File.Delete(rutaPlantilla);
                                    }
                                    log.EscribeLog(" Error al pasar de xlsx a xls ", ex.ToString(), true);
                                }

                            }

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
                //***********************||   GENERACIÓN DE REPORTE EN PDF   ||***********************
                try
                {
                    if (utilita == "P")
                    {
                        // SE TOMA EL ARCHIVO
                        string excelFilePath = CargaDatos.CargaCarpetaRaiz + CargaDatos.RUTAS[1].ToString() + "E" + numsesion + ".xls";
                        // SE CREA UNA NUEVA INSTANCIA DE INTEROP
                        Application excelApp = new Excel.Application();
                        // SE ABRE EL EXCEL USANDO INTEROP
                        Workbook workbook = excelApp.Workbooks.Open(excelFilePath);
                        try
                        {
                            // USANDO INTEROP, SE REALIZA LA CONVERSIÓN DE XLS A PDF
                            // SE USAN PARÁMETROS DE INTEROP PARA LA CONVERSIÓN DE PDF
                            string pdfFilePath = excelFilePath.Replace(".xls", ".pdf");
                            workbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfFilePath);
                            log.EscribeLog("Archivo PDF guardado exitosamente.");
                        }
                        catch (Exception ex)
                        {
                            log.EscribeLog("Error al guardar el archivo PDF: " + ex.Message);
                        }
                        finally
                        {
                            // CERRAR Y LIBERAR RECURSOS
                            workbook.Close(false);
                            excelApp.Quit();
                            Marshal.ReleaseComObject(workbook);
                            Marshal.ReleaseComObject(excelApp);
                        }
                    }
                }
                catch (Exception ex)
                {
                    if (File.Exists(rutaPlantilla))
                    {
                        File.Delete(rutaPlantilla);
                    }
                    log.EscribeLog("Error al generar el reporte en PDF: ", ex.ToString(), true);//true -> Termina el programa
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
            return null;
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
        // ** OPCIONAL **
        // NO SE USA, PERO SI SE DESEA USAR COM PARA USAR LECTURAS SE CONSERVA
        private void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (System.Exception ex)
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
