using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using WordToPDF.Exceptions;

namespace WordToPDF
{
    class Program
    {
        public static Microsoft.Office.Interop.Word.Application AppWord { get; set; }
        public static string ExportPath { get; set; }

        [STAThread]
        static void Main(string[] args)
        {
            Console.WriteLine("*-------------------------------------------------------*");
            Console.WriteLine($"> Microsoft Word to PDF");
            Console.WriteLine($"> Versión: { FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location).ProductVersion }");
            Console.WriteLine($"> Inicio: { DateTime.Now.ToString($"dd/MM/yyyy HH:mm:ss") }");
            Console.WriteLine("*-------------------------------------------------------*");

            try
            {
                // used as auxiliar to write errors on log and console
                string message = string.Empty;

                // used as default base path
                string directoryToProcess = Directory.GetCurrentDirectory();

                // get files to process
                var files = GetFilePaths(false, directoryToProcess);

                // get the error when no files detected
                if (!files.Any())
                {
                    message = $"> El directorio \"{ directoryToProcess }\" no contiene archivos para procesar.";
                    Console.WriteLine(message);
                }
                else
                {
                    AppWord = new Microsoft.Office.Interop.Word.Application();
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
            catch (CancellationException cEx)
            {
                Console.WriteLine($"** {cEx.Message}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"** Ocurrió un error en el proceso. Error: { ex.Message }.{ ex.InnerException?.Message }**");
            }
            finally
            {
                // close all open documents
                if (AppWord?.Documents?.Count > 0)
                    AppWord?.Documents?.Close();
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
                string exportPath = GetExportPath(false);

                // convert the file to pdf
                wordDocument = AppWord.Documents.Open(file);
                wordDocument.ExportAsFixedFormat($"{exportPath}\\{fileName}.pdf", WdExportFormat.wdExportFormatPDF);
            }
            catch (CancellationException)
            {
                throw;
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
        /// Get export path, If the user don't select, use default export path
        /// </summary>
        /// <param name="useDefaultExportPath"></param>
        /// <returns>Export Path or an CancellationException</returns>
        private static string GetExportPath(bool useDefaultExportPath)
        {
            if (!string.IsNullOrWhiteSpace(ExportPath))
                return ExportPath;

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog
            {
                Description = $"Seleccione la ruta desde donde se guardarán los archivos transformados a PDF.", 
                RootFolder = Environment.SpecialFolder.Desktop
            };
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                ExportPath = folderBrowserDialog.SelectedPath;
            }
            else if (useDefaultExportPath)
            {
                ExportPath = Path.Combine(Directory.GetCurrentDirectory(), "exports");
                Console.WriteLine($"> El usuario no seleccionó ninguna ruta, se utilizará la ruta por defecto: \"{ ExportPath }\"");
            }
            else
            {
                throw new CancellationException("GetExportPath", $"El usuario canceló su petición.");
            }

            return ExportPath;
        }

        /// <summary>
        /// Get files from given default basePath. If the user don't select, use default basePath
        /// </summary>
        /// <param name="basePath"></param> 
        /// <returns>IList<string> with files on selected folder.</returns>
        private static IList<string> GetFilePaths(bool useBasePath, string basePath)
        {
            IList<string> paths = new List<string>();

            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog()
            {
                Description = $"Seleccione la ruta desde donde se obtendrán los archivos de Microsoft Word.",
                RootFolder = Environment.SpecialFolder.Desktop
            };
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                paths = Directory.GetFiles(folderBrowserDialog.SelectedPath, "*.doc*").ToList();
            }
            else if (useBasePath && !string.IsNullOrWhiteSpace(basePath))
            {
                Console.WriteLine($"> El usuario no seleccionó ninguna ruta, se utilizará la ruta por defecto: \"{ basePath }\"");
                basePath = Directory.GetCurrentDirectory();
                paths = Directory.GetFiles(basePath, "*.doc*").ToList();
            }
            else
            {
                throw new CancellationException("GetFilePaths", $"El usuario canceló su petición.");
            }

            return paths;
        }
    }
}
