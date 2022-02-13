using GeneradorDeFiniquitos.Models;
using OfficeOpenXml;
using System;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Xceed.Document.NET;
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
                string[] tiempoServicio = new string[listaTrabajadores.Rows.Count];
                string[] totalMontos = new string[listaTrabajadores.Rows.Count];
                string[] nombresParaArchivo = new string[listaTrabajadores.Rows.Count];

                for (int i = 0; i < listaTrabajadores.Rows.Count; i++)
                {
                    rutTrabajadores[i] = listaTrabajadores.Rows[i][0].ToString();
                    nombreTrabajadores[i] = listaTrabajadores.Rows[i][1].ToString();
                    fechaContratoTrabajadores[i] = listaTrabajadores.Rows[i][2].ToString();
                    fechaFiniquitoTrabajadores[i] = listaTrabajadores.Rows[i][3].ToString();
                    vacacionesProporcionales[i] = listaTrabajadores.Rows[i][4].ToString();
                    tiempoServicio[i] = listaTrabajadores.Rows[i][5].ToString();
                    totalMontos[i] = listaTrabajadores.Rows[i][6].ToString();
                    nombresParaArchivo[i] = listaTrabajadores.Rows[i][7].ToString();
                }

                return GenerarDocumentosWord(rutTrabajadores, nombreTrabajadores, fechaContratoTrabajadores, fechaFiniquitoTrabajadores, vacacionesProporcionales, tiempoServicio, totalMontos, nombresParaArchivo);
            }

            return null;
        }

        public FileResult GenerarDocumentosWord(string[] rutTrabajadores, string[] nombreTrabajadores, string[] fechaContrato, string[] fechaFiniquito,
            string[] montoVacaciones, string[] montoTiempoServicio, string[] totalMontos, string[] nombreParaArchivo)
        {
            int tamArreglo = rutTrabajadores.Length;
            for (int i = 0; i < tamArreglo; i++)
            {
                using (var document = DocX.Create("Finiquito " + nombreParaArchivo[i] + ".docx"))
                {
                    //Cabecera documento
                    document.AddHeaders();

                    var titulo = document.InsertParagraph();
                    var ParrafoUno = document.InsertParagraph();
                    var ParrafoDos = document.InsertParagraph();
                    var ParrafoTres = document.InsertParagraph();
                    //var header = document.Headers.Odd.InsertParagraph();

                    titulo.Append("FINIQUITO DE CONTRATO DE TRABAJO").Font("Cambria").FontSize(19).SpacingAfter(10d).Alignment = Alignment.center;

                    //.Font("Cambria").FontSize(12).Alignment = Alignment.both;

                    ParrafoUno.Append("En Santiago Chile, " + DateTime.Parse(fechaFiniquito[i]).ToString("d 'de' MMMM 'de' yyyy") + " entre ")
                        .Font("Cambria").FontSize(12)
                        .Append("Framberry Agrícola SA.")
                        .Font("Cambria").FontSize(12).Bold()
                        .Append(" RUT Nº 76.470.708-7, representada por don Carlos Luis Brito Claissac, chileno, casado, cédula nacional de identidad Nº 8.975.812-2,")
                        .Font("Cambria").FontSize(12)
                        .Append("ambos domiciliados para estos afectos en Avda. Américo Vespucio Norte Nº 2500, Comuna de Vitacura, Santiago, en adelante e indistintamente ")
                        .Font("Cambria").FontSize(12)
                        .Append("el “Empleador”; por una parte, y por la otra, don(a) ")
                        .Font("Cambria").FontSize(12)
                        .Append(nombreTrabajadores[i])
                        .Font("Cambria").FontSize(12).Bold()
                        .Append(", nacionalidad chilena, Obrero Agrícola, cédula nacional de identidad Nº " + rutTrabajadores[i] + " en adelante indistintamente el “Trabajador”; ")
                        .Font("Cambria").FontSize(12)
                        .Append("quienes han acordado el siguiente finiquito: ")
                        .Font("Cambria").FontSize(12).SpacingAfter(15d);

                    ParrafoUno.Alignment = Alignment.both;

                    ParrafoDos.Append("PRIMERO: ").Font("Cambria").FontSize(12).Bold()
                        .Append("El trabajador declara haber prestado servicios para el Empleador, ejecutando las labores de Obrero Agrícola, desde el " +
                        DateTime.Parse(fechaContrato[i]).ToString("d MMMM yyyy") + " hasta el " + DateTime.Parse(fechaFiniquito[i]).ToString("d MMMM yyyy") + ", " +
                        "fecha en que ambas partes pusieron término al Contrato de Trabajo por la causal contemplada en el artículo 159 n° 5, del Código del Trabajo, ")
                        .Font("Cambria").FontSize(12)
                        .Append("esto es por “Conclusión del trabajo o servicio que dio origen al Contrato”")
                        .Font("Cambria").FontSize(12).Bold()
                        .Append(".").Font("Cambria").FontSize(12).SpacingAfter(15d);

                    ParrafoDos.Alignment = Alignment.both;

                    ParrafoTres.Append("SEGUNDO: ").Font("Cambria").FontSize(12).Bold()
                        .Append("El trabajador declara haber recibido con anterioridad a esta fecha de partes de")
                        .Font("Cambria").FontSize(12)
                        .Append(" Framberry Agrícola S.A.").Font("Cambria").FontSize(12).Bold()
                        .Append(", a su entera satisfacción, los valores que se indican por los conceptos siguientes:")
                        .Font("Cambria").FontSize(12).SpacingAfter(15d);

                    ParrafoTres.Alignment = Alignment.both;

                    //var table = document.InsertTable(1, 3);

                    //table.AutoFit = AutoFit.Contents;
                    //var border = new Border(BorderStyle.Tcbs_single, BorderSize.one, 0, Color.Black);
                    //table.SetBorder(TableBorderType.InsideH, border);

                    //var tableHeaders = table.Rows[0];
                    //tableHeaders.Cells[0].InsertParagraph().Append("Detalle").Bold();
                    //tableHeaders.Cells[1].InsertParagraph().Append("Haberes").Bold();
                    //tableHeaders.Cells[2].InsertParagraph().Append("Descuentos").Bold();

                    //var tableRow = table.InsertRow();
                    //tableRow.Cells[0].InsertParagraph().Append(lecture.Id.ToString());
                    //tableRow.Cells[1].InsertParagraph().Append(lecture.Name);
                    //tableRow.Cells[2].InsertParagraph().Append(lecture.Level);


                    // Add a Table of 5 rows and 2 columns into the document and sets its values.
                    var t = document.AddTable(4, 3);
                    t.Design = TableDesign.TableGrid;
                    t.Alignment = Alignment.center;
                    t.Rows[3].MergeCells(1, 2);
                    t.SetColumnWidth(0, 220.0,false);
                    t.SetColumnWidth(1, 60.0,false);
                    t.SetColumnWidth(2, 60.0,false);

                    var detalles = t.Rows[0].Cells[0];
                    var haberes = t.Rows[0].Cells[1];
                    var descuentos = t.Rows[0].Cells[2];

                    detalles.Paragraphs[0].Append("Detalle").Font("Arial").FontSize(10).Bold().Italic().Alignment = Alignment.left;
                    haberes.Paragraphs[0].Append("Haberes").Font("Arial").FontSize(10).Bold().Italic().Alignment = Alignment.right;
                    descuentos.Paragraphs[0].Append("Descuentos").Font("Arial").FontSize(10).Bold().Italic().Alignment = Alignment.right;

                    t.Rows[1].Cells[0].Paragraphs[0].Append("Vacaciones Proporcionales \n").Font("Arial").FontSize(10)
                        .Append("Indemnización por Tiempo Servido \n").Font("Arial").FontSize(10).Alignment = Alignment.left;
                    t.Rows[1].Cells[1].Paragraphs[0]
                        .Append(montoVacaciones[i]+ "\n").Font("Arial").FontSize(10)
                        .Append(montoTiempoServicio[i]).Font("Arial").FontSize(10).Alignment = Alignment.right;

                    t.Rows[2].Cells[0].Paragraphs[0].Append("Total").Font("Arial").Bold()
                        .FontSize(10).Alignment = Alignment.left;
                    t.Rows[2].Cells[1].Paragraphs[0].Append(totalMontos[i]).Font("Arial").Bold()
                        .FontSize(10).Alignment = Alignment.right;
                    t.Rows[2].Cells[2].Paragraphs[0].Append("0").Font("Arial").Bold()
                        .FontSize(10).Alignment = Alignment.right;

                    t.Rows[3].Cells[0].Paragraphs[0].Append("Líquido a Pagar").Font("Arial").Bold().Italic()
                       .FontSize(10).Alignment = Alignment.center;
                    t.Rows[3].Cells[1].Paragraphs[0].Append(totalMontos[i]).Font("Arial").Bold()
                        .FontSize(10).Alignment = Alignment.center;

                    //t.Rows[2].Cells[0].Paragraphs[0].Append("Carl");
                    //t.Rows[2].Cells[1].Paragraphs[0].Append("60");
                    //t.Rows[3].Cells[0].Paragraphs[0].Append("Michael");
                    //t.Rows[3].Cells[1].Paragraphs[0].Append("59");
                    //t.Rows[4].Cells[0].Paragraphs[0].Append("Shawn");
                    //t.Rows[4].Cells[1].Paragraphs[0].Append("57");
                    document.InsertTable(t);

                    //DocX.ConvertToPdf(document, "ConvertedDocument.pdf");

                    document.Save();
                }
            }



            //int tamArreglo = rutPacientesExcel.Length;
            //string[] rutPcrGenerado = new string[tamArreglo];
            //string[] nombrePcrGenerado = new string[tamArreglo];
            //string[] correlativoMicroGenerado = new string[tamArreglo];
            //string[] correlativoLabGenerado = new string[tamArreglo];
            //string[] fechaPcrGenerado = new string[tamArreglo];
            //string[] observacionesPcr = new string[tamArreglo];

            // validar si el paciente esta en la pac paciente 
            //BD_ENTI_CORPORATIVAEntities context = new BD_ENTI_CORPORATIVAEntities();

            //for (int i = 0; i < tamArreglo; i++)
            //{
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
            //}
            // se envian los arreglos cargados con la información
            //return generarExcelPcrGenerados(rutPcrGenerado, nombrePcrGenerado, correlativoMicroGenerado, correlativoLabGenerado, fechaPcrGenerado, observacionesPcr, nombreArchivo);
            return null;
        }


        // se genera el excel con los datos
        //public FileResult generarExcelPcrGenerados(string[] rutPcrGenerado, string[] nombrePcrGenerado, string[] correlativoMicroGenerado, string[] correlativoLabGenerado, string[] fechaPcrGenerado, string[] observacionesPcr, string nombreArchivo)
        //{
        //    using (ExcelPackage excel = new ExcelPackage())
        //    {
        //        // se le da nombre a la hoja 
        //        var messageBook = excel.Workbook.Worksheets.Add("PCR_Generados_" + DateTime.Now.ToShortDateString());

        //        // se setea los titulos en la cabecera
        //        messageBook.Cells["A1"].Value = "Correlativo Microbiologia";
        //        messageBook.Cells["B1"].Value = "Correlativo Laboratorio";
        //        messageBook.Cells["C1"].Value = "Rut paciente (funcionario)";
        //        messageBook.Cells["D1"].Value = "Nombre paciente (funcionario)";
        //        messageBook.Cells["E1"].Value = "Fecha PCR generado";
        //        messageBook.Cells["F1"].Value = "Observación";

        //        //se recorren los arreglos para ir poniendo la informacion 
        //        int fila = 1;
        //        for (int i = 0; i < rutPcrGenerado.Length; i++)
        //        {
        //            fila++;
        //            // si al menos el correlativo micro viene con un "Sin comentario" se marca la fila completa en amarillo
        //            if (correlativoMicroGenerado[i] == "S/C")
        //            {
        //                messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.Font.Bold = true;
        //                messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //                messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.Fill.BackgroundColor.SetColor(Color.Yellow);
        //                messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //                messageBook.Cells["A" + (fila) + ":F" + (fila)].Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
        //            }
        //            // se agrega la informacion a las celdas
        //            messageBook.Cells["A" + (fila).ToString()].Value = correlativoMicroGenerado[i];
        //            messageBook.Cells["B" + (fila).ToString()].Value = correlativoLabGenerado[i];
        //            messageBook.Cells["C" + (fila).ToString()].Value = rutPcrGenerado[i];
        //            messageBook.Cells["D" + (fila).ToString()].Value = nombrePcrGenerado[i];
        //            messageBook.Cells["E" + (fila).ToString()].Value = fechaPcrGenerado[i];
        //            messageBook.Cells["F" + (fila).ToString()].Value = observacionesPcr[i];

        //        }


        //        // Formato
        //        messageBook.Cells.Style.Font.SetFromFont(new System.Drawing.Font("Calibri", 11));
        //        // cabecera
        //        messageBook.Cells["A1:F1"].Style.Font.Bold = true;
        //        messageBook.Cells["A1:F1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
        //        messageBook.Cells["A1:F1"].Style.Fill.BackgroundColor.SetColor(Color.FromArgb(242, 242, 242));
        //        messageBook.Cells["A1:F1"].Style.VerticalAlignment = ExcelVerticalAlignment.Center;
        //        messageBook.Cells["A1:F1"].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        //        // messageBook.Cells["A1:F1"].Style.Font.Color.SetColor(Color.White);
        //        messageBook.Cells["A1:F1"].Style.Font.Color.SetColor(Color.FromArgb(255, 128, 0));
        //        messageBook.Cells.AutoFitColumns();

        //        // poner bordes a las celdas
        //        messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].Style.Border.Top.Style = ExcelBorderStyle.Thin;
        //        messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
        //        messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].Style.Border.Left.Style = ExcelBorderStyle.Thin;
        //        messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].Style.Border.Right.Style = ExcelBorderStyle.Thin;
        //        messageBook.Cells["A1:F" + (rutPcrGenerado.Length + 1)].AutoFilter = true;
        //        var stream = new MemoryStream(excel.GetAsByteArray());

        //        // nombre del reporte y descarga
        //        return File(stream, "application/excel", "AutoPCR_listado_" + nombreArchivo + "_" + DateTime.Now.ToShortDateString() + ".xlsx");
        //    }
        //}
    }
}