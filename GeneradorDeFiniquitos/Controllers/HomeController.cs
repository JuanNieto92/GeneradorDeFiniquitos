using GeneradorDeFiniquitos.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Xceed.Words.NET;

namespace GeneradorDeFiniquitos.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }

        [HttpGet]
        public ActionResult GenerarFiniquitos()
        {
            ViewBag.Message = "Your contact page.";
            return View("../Forms/GenerarFiniquitos");
        }

        [HttpPost]
        public FileResult ReadExcel(ModeloArchivoFiniquito modeloFiniquito)
        {
            // se valida la extension del archivo
            if (Path.GetExtension(modeloFiniquito.archivo.FileName).Contains("xlsx") || Path.GetExtension(modeloFiniquito.archivo.FileName).Contains("xls"))
            {
                ExcelPackage package = new ExcelPackage(modeloFiniquito.archivo.InputStream);
                DataTable listaTrabajadores = ExcelPackageExtensions.ToDataTable(package);

                string[] rutTrabajadores = new string[listaTrabajadores.Rows.Count];
                string[] nombreTrabajadores = new string[listaTrabajadores.Rows.Count];
                string[] fechaContratoTrabajadores = new string[listaTrabajadores.Rows.Count];
                string[] fechaFiniquitoTrabajadores = new string[listaTrabajadores.Rows.Count];
                string[] vacacionesProporcionales = new string[listaTrabajadores.Rows.Count];
                string[] timporServicio = new string[listaTrabajadores.Rows.Count];
                string[] totalMontos = new string[listaTrabajadores.Rows.Count];
                string[] nombresParaArchivo = new string[listaTrabajadores.Rows.Count];

                for (int i = 0; i < listaTrabajadores.Rows.Count; i++)
                {
                    rutTrabajadores[i] = listaTrabajadores.Rows[i][0].ToString();
                    nombreTrabajadores[i] = listaTrabajadores.Rows[i][1].ToString();
                    fechaContratoTrabajadores[i] = listaTrabajadores.Rows[i][2].ToString();
                    fechaFiniquitoTrabajadores[i] = listaTrabajadores.Rows[i][3].ToString();
                    vacacionesProporcionales[i] = listaTrabajadores.Rows[i][4].ToString();
                    timporServicio[i] = listaTrabajadores.Rows[i][5].ToString();
                    totalMontos[i] = listaTrabajadores.Rows[i][6].ToString();
                    nombresParaArchivo[i] = listaTrabajadores.Rows[i][7].ToString();
                }

               return GenerarDocumentosWord(rutTrabajadores, nombreTrabajadores);
            }

            return null;
        }

        public FileResult GenerarDocumentosWord(string[] rutPacientesExcel, string[] nombrePacientesExcel /*, string rutProfesional, string nombreArchivo*/)
        {

            using (var document = DocX.Create("Prueba.docx"))
            {
                document.Save();
            }

            int tamanioArreglo = rutPacientesExcel.Length;
            string[] rutPcrGenerado = new string[tamanioArreglo];
            string[] nombrePcrGenerado = new string[tamanioArreglo];
            string[] correlativoMicroGenerado = new string[tamanioArreglo];
            string[] correlativoLabGenerado = new string[tamanioArreglo];
            string[] fechaPcrGenerado = new string[tamanioArreglo];
            string[] observacionesPcr = new string[tamanioArreglo];

            // validar si el paciente esta en la pac paciente 
            //BD_ENTI_CORPORATIVAEntities context = new BD_ENTI_CORPORATIVAEntities();

            for (int i = 0; i < tamanioArreglo; i++)
            {
                //var rutPacientesBusqueda = rutPacientesExcel[i];
                //PAC_Paciente paciente = context.PAC_Paciente.Where(p => p.PAC_PAC_Rut == rutPacientesBusqueda).FirstOrDefault();
                // si el paciente es null guarda info en los arreglos para posterior generacion del excel
                //if (paciente == null)
                //{
                //    rutPcrGenerado[i] = rutPacientesExcel[i];
                //    nombrePcrGenerado[i] = nombrePacientesExcel[i];
                //    correlativoMicroGenerado[i] = "S/C";
                //    correlativoLabGenerado[i] = "S/C";
                //    fechaPcrGenerado[i] = "S/F";
                //    observacionesPcr[i] = "Rut mal digitado o no existe en la base de datos";
                //    ViewBag.mensajeError = "El RUT no fue encontrado";
                //}
                //else
                //{
                // si existe se manda a generar el PCR automatico
                //SolicitudMicroController smc = new SolicitudMicroController();
                //ModeloResultadoSolicitud resultadoSolicitud = smc.SolicitarSolicitudPCR(paciente, rutProfesional);
                // si ocurre algun error muestra el mensaje correspondiente
                //if (resultadoSolicitud.error)
                //{
                //    ViewBag.mensajeError = resultadoSolicitud.mensajeError;
                //}
                //else
                //{
                //    // si se genero correctamente PCR guarda la info para posterior generacion de excel 
                //    rutPcrGenerado[i] = rutPacientesExcel[i];
                //    nombrePcrGenerado[i] = paciente.PAC_PAC_Nombre + " " + paciente.PAC_PAC_ApellPater + " " + paciente.PAC_PAC_ApellMater;
                //    correlativoMicroGenerado[i] = resultadoSolicitud.correlativoMicro;
                //    correlativoLabGenerado[i] = resultadoSolicitud.correlativoLab;
                //    fechaPcrGenerado[i] = resultadoSolicitud.fechaGeneracionPCR;
                //    observacionesPcr[i] = "PCR generado correctamente";
                //    ViewBag.exito = true;
                //}
                //}
            }
            // se envian los arreglos cargados con la información
            //return generarExcelPcrGenerados(rutPcrGenerado, nombrePcrGenerado, correlativoMicroGenerado, correlativoLabGenerado, fechaPcrGenerado, observacionesPcr, nombreArchivo);
            return null;
        }


        // se genera el excel con los datos
        public FileResult generarExcelPcrGenerados(string[] rutPcrGenerado, string[] nombrePcrGenerado, string[] correlativoMicroGenerado, string[] correlativoLabGenerado, string[] fechaPcrGenerado, string[] observacionesPcr, string nombreArchivo)
        {
            using (ExcelPackage excel = new ExcelPackage())
            {
                // se le da nombre a la hoja 
                var messageBook = excel.Workbook.Worksheets.Add("PCR_Generados_" + DateTime.Now.ToShortDateString());

                // se setea los titulos en la cabecera
                messageBook.Cells["A1"].Value = "Correlativo Microbiologia";
                messageBook.Cells["B1"].Value = "Correlativo Laboratorio";
                messageBook.Cells["C1"].Value = "Rut paciente (funcionario)";
                messageBook.Cells["D1"].Value = "Nombre paciente (funcionario)";
                messageBook.Cells["E1"].Value = "Fecha PCR generado";
                messageBook.Cells["F1"].Value = "Observación";

                //se recorren los arreglos para ir poniendo la informacion 
                int fila = 1;
                for (int i = 0; i < rutPcrGenerado.Length; i++)
                {
                    fila++;
                    // si al menos el correlativo micro viene con un "Sin comentario" se marca la fila completa en amarillo
                    if (correlativoMicroGenerado[i] == "S/C")
                    {
                        messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.Font.Bold = true;
                        messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
                        messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                        messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    }
                    // se agrega la informacion a las celdas
                    messageBook.Cells["A" + (fila).ToString()].Value = correlativoMicroGenerado[i];
                    messageBook.Cells["B" + (fila).ToString()].Value = correlativoLabGenerado[i];
                    messageBook.Cells["C" + (fila).ToString()].Value = rutPcrGenerado[i];
                    messageBook.Cells["D" + (fila).ToString()].Value = nombrePcrGenerado[i];
                    messageBook.Cells["E" + (fila).ToString()].Value = fechaPcrGenerado[i];
                    messageBook.Cells["F" + (fila).ToString()].Value = observacionesPcr[i];

                }

                // Formato
                messageBook.Cells.Style.Font.SetFromFont(new Font("Calibri", 11));
                // cabecera
                messageBook.Cells["A1:F1"].Style.Font.Bold = true;
                messageBook.Cells["A1:F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                messageBook.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
                messageBook.Cells["A1:F1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                messageBook.Cells["A1:F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                // messageBook.Cells["A1:F1"].Style.Font.Color.SetColor(Color.White);
                messageBook.Cells["A1:F1"].Style.Font.Color.SetColor(Color.FromArgb(255, 128, 0));
                messageBook.Cells.AutoFitColumns();

                // poner bordes a las celdas
                messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].AutoFilter = true;
                var stream = new MemoryStream(excel.GetAsByteArray());

                // nombre del reporte y descarga
                return File(stream, "application/excel", "AutoPCR_listado_" + nombreArchivo + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
            }
        }
    }
}