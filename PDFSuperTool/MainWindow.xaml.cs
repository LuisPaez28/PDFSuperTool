using System;
using System.Collections.Generic;
using System.Diagnostics; // Necesario para ejecutar LibreOffice
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32; // Aquí están los diálogos nativos de WPF

// Alias para PDF (Separar/Unir) - Mantenemos esto que sí funcionaba
using SharpDoc = PdfSharp.Pdf.PdfDocument;
using PdfSharp.Pdf.IO;

namespace PDFSuperTool
{
    public partial class MainWindow : Window
    {
        private List<string> archivosParaUnir = new List<string>();

        public MainWindow()
        {
            InitializeComponent();
        }

        // ==========================================
        // 1. SEPARAR PDF (Sin cambios, funciona bien)
        // ==========================================
        private async void BtnSeparar_Click(object sender, RoutedEventArgs e)
        {
            string pdfPath = txtPathSplit.Text;

            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
            {
                MessageBox.Show("Selecciona un PDF válido.");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "Guardar páginas...",
                FileName = "Pagina.pdf",
                Filter = "PDF|*.pdf"
            };

            if (sfd.ShowDialog() == true)
            {
                string carpetaDestino = Path.GetDirectoryName(sfd.FileName);
                string nombreBase = Path.GetFileNameWithoutExtension(pdfPath);

                try
                {
                    lblStatus.Text = "Procesando...";
                    await Task.Run(() =>
                    {
                        using (SharpDoc input = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
                        {
                            for (int i = 0; i < input.PageCount; i++)
                            {
                                using (SharpDoc output = new SharpDoc())
                                {
                                    output.AddPage(input.Pages[i]);
                                    output.Save(Path.Combine(carpetaDestino, $"{nombreBase}_pag_{i + 1}.pdf"));
                                }
                            }
                        }
                    });
                    MessageBox.Show("Separación completada.");
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
                finally { lblStatus.Text = "Listo."; }
            }
        }

        // ==========================================
        // 2. UNIR PDF (Sin cambios, funciona bien)
        // ==========================================
        private async void BtnUnir_Click(object sender, RoutedEventArgs e)
        {
            if (archivosParaUnir.Count < 2)
            {
                MessageBox.Show("Faltan archivos para unir.");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF|*.pdf", FileName = "Unido.pdf" };
            if (sfd.ShowDialog() == true)
            {
                try
                {
                    lblStatus.Text = "Uniendo...";
                    string salida = sfd.FileName;
                    await Task.Run(() =>
                    {
                        using (SharpDoc output = new SharpDoc())
                        {
                            foreach (string f in archivosParaUnir)
                            {
                                using (SharpDoc input = PdfReader.Open(f, PdfDocumentOpenMode.Import))
                                    foreach (var p in input.Pages) output.AddPage(p);
                            }
                            output.Save(salida);
                        }
                    });
                    MessageBox.Show("Unión lista.");
                }
                catch (Exception ex) { MessageBox.Show("Error: " + ex.Message); }
                finally { lblStatus.Text = "Listo."; }
            }
        }

        // ==========================================
        // 3. CONVERTIR (CORREGIDO: 100% WPF NATIVO)
        // ==========================================
        private async void BtnConvertir_Click(object sender, RoutedEventArgs e)
        {
            string pdfPath = txtPathConvert.Text;
            bool esExcel = chkEsExcel.IsChecked ?? false;

            // 1. Validaciones
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
            {
                MessageBox.Show("Por favor, selecciona el PDF a convertir.");
                return;
            }

            string rutaLibreOffice = BuscarLibreOffice();
            if (string.IsNullOrEmpty(rutaLibreOffice))
            {
                MessageBox.Show("No encontré LibreOffice.");
                return;
            }

            // 2. Elegir carpeta destino
            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "Selecciona la carpeta donde guardar",
                Filter = esExcel ? "Carpeta|*.xlsx" : "Carpeta|*.docx",
                FileName = "Guardar_Aqui"
            };

            if (sfd.ShowDialog() == true)
            {
                string carpetaDestino = Path.GetDirectoryName(sfd.FileName);

                // Creamos una carpeta temporal única para asegurar que capturamos el archivo correcto
                string carpetaTemporal = Path.Combine(carpetaDestino, "TEMP_CONVERSION_" + Guid.NewGuid().ToString().Substring(0, 8));

                lblStatus.Text = "Convirtiendo...";

                try
                {
                    // Creamos la carpeta temporal (El cuarto aislado)
                    Directory.CreateDirectory(carpetaTemporal);

                    await Task.Run(() =>
                    {
                        // A. Convertimos y guardamos en la carpeta temporal vacía
                        ConvertirConLibreOffice(rutaLibreOffice, pdfPath, carpetaTemporal, esExcel);

                        // Esperamos un segundo por seguridad de disco
                        System.Threading.Thread.Sleep(1000);
                    });

                    // B. Buscamos qué archivo se creó ahí dentro
                    string[] archivosGenerados = Directory.GetFiles(carpetaTemporal);

                    if (archivosGenerados.Length > 0)
                    {
                        // ¡Lo encontramos! Es el único archivo en esa carpeta
                        string archivoTemporal = archivosGenerados[0];
                        string nombreArchivo = Path.GetFileName(archivoTemporal);
                        string rutaFinal = Path.Combine(carpetaDestino, nombreArchivo);

                        // C. Lo movemos a la carpeta real (Sobrescribiendo si existe)
                        if (File.Exists(rutaFinal)) File.Delete(rutaFinal);
                        File.Move(archivoTemporal, rutaFinal);

                        // D. Limpieza: Borramos la carpeta temporal
                        try { Directory.Delete(carpetaTemporal, true); } catch { }

                        MessageBox.Show("¡Éxito! Archivo generado correctamente.");

                        // Abrir explorador
                        Process.Start("explorer.exe", $"/select,\"{rutaFinal}\"");
                    }
                    else
                    {
                        // Si la carpeta sigue vacía, LibreOffice falló silenciosamente
                        MessageBox.Show("Error extraño: LibreOffice terminó pero no dejó ningún archivo en la carpeta temporal.");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    // Aseguramos borrar la carpeta temporal si hubo error y quedó ahí
                    if (Directory.Exists(carpetaTemporal))
                    {
                        try { Directory.Delete(carpetaTemporal, true); } catch { }
                    }
                    lblStatus.Text = "Listo.";
                }
            }
        }

        // ==========================================
        // LÓGICA LIBREOFFICE (CMD)
        // ==========================================
        private void ConvertirConLibreOffice(string exePath, string inputFile, string outputDir, bool esExcel)
        {
            string formato = esExcel ? "xlsx" : "docx";

            // --outdir le dice a LibreOffice: "Guarda el resultado aquí, pero no cambies el nombre del archivo"
            string args = $"--headless --convert-to {formato} --outdir \"{outputDir}\" \"{inputFile}\"";

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = args,
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardError = true // Para capturar si LibreOffice se queja
            };

            using (Process process = Process.Start(startInfo))
            {
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    throw new Exception($"LibreOffice reportó un error (Código {process.ExitCode}):\n{error}");
                }
            }
        }
        private string BuscarLibreOffice()
        {
            // Rutas típicas en Windows 64 bits
            string[] rutas = {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };

            foreach (var r in rutas) if (File.Exists(r)) return r;
            return null;
        }

        // Helpers UI
        private void BtnSeleccionar_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "PDF|*.pdf" };
            if (ofd.ShowDialog() == true) txtPathSplit.Text = ofd.FileName;
        }

        private void BtnAgregarALista_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Multiselect = true, Filter = "PDF|*.pdf" };
            if (ofd.ShowDialog() == true)
            {
                foreach (string f in ofd.FileNames)
                {
                    archivosParaUnir.Add(f);
                    listArchivosUnir.Items.Add(Path.GetFileName(f));
                }
            }
        }

        private void BtnLimpiar_Click(object sender, RoutedEventArgs e)
        {
            archivosParaUnir.Clear();
            listArchivosUnir.Items.Clear();
            txtPathSplit.Clear();
            txtPathConvert.Clear();
            lblStatus.Text = "Limpiado.";
        }

        private void BtnSeleccionarConvert_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "PDF|*.pdf" };
            if (ofd.ShowDialog() == true) txtPathConvert.Text = ofd.FileName;
        }

        private string BuscarArchivoReciente(string carpeta, bool esExcel, DateTime horaMinima)
        {
            // Esperamos un momento para asegurar que el sistema de archivos se actualice
            System.Threading.Thread.Sleep(1500);

            string patron = esExcel ? "*.xlsx" : "*.docx";
            DirectoryInfo dirInfo = new DirectoryInfo(carpeta);

            // Obtenemos todos los Excel/Word de la carpeta
            var archivos = dirInfo.GetFiles(patron);

            // Ordenamos por fecha de modificación (el más nuevo primero)
            // y filtramos para que solo tome los que se modificaron DESPUÉS de que empezamos el proceso.
            var archivoReciente = archivos
                .Where(f => f.LastWriteTime >= horaMinima.AddSeconds(-5) || f.CreationTime >= horaMinima.AddSeconds(-5))
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault();

            if (archivoReciente != null)
            {
                return archivoReciente.FullName;
            }

            return null; // No encontramos nada nuevo
        }
    }
}