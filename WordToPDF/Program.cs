using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;

namespace WordToPDF
{
    class Program
    {
        public static Application AppWord { get; set; }

        static void Main(string[] args)
        {
            Console.WriteLine("*-------------------------------------------------------*");
            Console.WriteLine($"> Microsoft Word to PDF");
            Console.WriteLine($"> Versión: { FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).FileVersion }");
            Console.WriteLine($"> Inicio: { DateTime.Now.ToString($"dd/MM/yyyy HH:mm:ss") }");
            Console.WriteLine("*-------------------------------------------------------*");

            try
            {
                // used as auxiliar to write errors on log and console
                string message = string.Empty;

                string directoryToProcess = Directory.GetCurrentDirectory();

                // get files to process
                var files = GetFilePaths(directoryToProcess);

                // get the error when no files detected
                if (!files.Any())
                {
                    message = $"> El directorio \"{ directoryToProcess }\" no contiene archivos para procesar.";
                    Console.WriteLine(message);
                }
                else
                {
                    AppWord = new Application();
                    Console.WriteLine($"> Cantidad de archivos a procesar: { files.Count }");
                    Console.WriteLine("*-------------------------------------------------------*");
                    int correctlyExported = 0;

                    foreach (string file in files)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(file);
                        Console.WriteLine($"> Procesando archivo: {fileName}");

                        if (!File.Exists(file))
                        {
                            message = $">> El archivo: {fileName} no existe o no está disponible para ser procesado.";
                            Console.WriteLine(message);
                        }
                        else
                        {
                            if (ConvertToPDF(file, fileName))
                            {
                                message = $">> El archivo: {fileName} fue exportado correctamente.";
                                Console.WriteLine(message);
                                correctlyExported++;
                            }
                        }
                    }
                    Console.WriteLine("*-------------------------------------------------------*");
                    Console.WriteLine($"Archivos exportados: {correctlyExported} de {files.Count}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"** Ocurrió un error en el proceso. Error: { ex.Message }.{ ex.InnerException?.Message }**");
            }
            finally
            {
                // close word app
                AppWord?.Quit();
                // show end of the process
                Console.WriteLine("*-------------------------------------------------------*");
                Console.WriteLine($"> Término: { DateTime.Now.ToString($"dd/MM/yyyy HH:mm:ss") }");
                Console.WriteLine("*-------------------------------------------------------*");
                Console.Write($"> Proceso completo, presione la tecla Enter para salir...");
                Console.ReadLine();
            }
        }

        /// <summary>
        /// Convert a docx file to pdf with Microsoft Word Interop
        /// </summary>
        /// <param name="file"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        private static bool ConvertToPDF(string file, string fileName)
        {
            bool exported = true;

            Document wordDocument = null;
            try
            {
                // convert the file to pdf
                wordDocument = AppWord.Documents.Open(file);
                wordDocument.ExportAsFixedFormat($"{Path.GetDirectoryName(file)}\\{fileName}.pdf", WdExportFormat.wdExportFormatPDF);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ocurrió un error al intentar procesar el archivo { file }. Error: { ex.Message }.{ ex.InnerException?.Message }");
                exported = false;
            }
            finally
            {
                wordDocument?.Close();
            }

            return exported;
        }

        /// <summary>
        /// Get files from given basePath. If no base path, use the current directory.
        /// </summary>
        /// <param name="basePath"></param> 
        /// <returns></returns>
        private static IList<string> GetFilePaths(string basePath)
        {
            if (string.IsNullOrWhiteSpace(basePath))
                basePath = Directory.GetCurrentDirectory();

            return Directory.GetFiles(basePath, "*.docx").ToList();
        }
    }
}
