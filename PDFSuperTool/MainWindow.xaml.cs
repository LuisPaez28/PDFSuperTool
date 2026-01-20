using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

// Alias para PDF (Separar/Unir)
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
        // 1. SEPARAR PDF (Funciona perfecto)
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
        // 2. UNIR PDF (Funciona perfecto)
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
        // 3. CONVERTIR (SOLUCIÓN BLINDADA FINAL)
        // ==========================================
        private async void BtnConvertir_Click(object sender, RoutedEventArgs e)
        {
            string pdfPath = txtPathConvert.Text;
            bool esExcel = chkEsExcel.IsChecked ?? false;

            // Validaciones iniciales
            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
            {
                MessageBox.Show("Selecciona el PDF a convertir.");
                return;
            }

            string rutaLibreOffice = BuscarLibreOffice();
            if (string.IsNullOrEmpty(rutaLibreOffice))
            {
                MessageBox.Show("No encontré LibreOffice. Verifica la instalación.");
                return;
            }

            // 1. Matar procesos viejos para liberar memoria
            MatarProcesosLibreOffice();

            // 2. Preguntar dónde guardar el archivo final
            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "Selecciona dónde guardar el resultado",
                Filter = esExcel ? "Excel (*.xlsx)|*.xlsx" : "Word (*.docx)|*.docx",
                FileName = Path.GetFileNameWithoutExtension(pdfPath) // Sugerir nombre original
            };

            if (sfd.ShowDialog() == true)
            {
                string rutaDestinoFinal = sfd.FileName;

                // 3. Crear carpeta TEMP del sistema (Segura y sin espacios raros)
                string carpetaTempSistema = Path.Combine(Path.GetTempPath(), "PDFTool_" + DateTime.Now.Ticks);
                Directory.CreateDirectory(carpetaTempSistema);

                // 4. Copiar PDF original a la Temp llamándolo "Input.pdf" 
                // (Esto evita errores si tu archivo original tiene nombres raros)
                string pdfTemporal = Path.Combine(carpetaTempSistema, "Input.pdf");
                File.Copy(pdfPath, pdfTemporal, true);

                lblStatus.Text = "Convirtiendo...";

                try
                {
                    await Task.Run(() =>
                    {
                        // Convertimos el "Input.pdf"
                        ConvertirConLibreOffice(rutaLibreOffice, pdfTemporal, carpetaTempSistema, esExcel);

                        // Pequeña espera técnica
                        System.Threading.Thread.Sleep(1500);
                    });

                    // 5. Buscar el resultado (se llamará Input.xlsx o Input.docx)
                    string extension = esExcel ? ".xlsx" : ".docx";
                    string archivoGeneradoTemp = Path.Combine(carpetaTempSistema, "Input" + extension);

                    if (File.Exists(archivoGeneradoTemp))
                    {
                        // 6. Mover el archivo generado a donde pidió el usuario
                        if (File.Exists(rutaDestinoFinal)) File.Delete(rutaDestinoFinal);
                        File.Move(archivoGeneradoTemp, rutaDestinoFinal);

                        MessageBox.Show("¡Conversión Exitosa!");

                        // Abrir explorador
                        Process.Start("explorer.exe", $"/select,\"{rutaDestinoFinal}\"");
                    }
                    else
                    {
                        // Debugging: ver qué pasó si falló
                        string[] archivos = Directory.GetFiles(carpetaTempSistema);
                        string lista = string.Join("\n", archivos);
                        MessageBox.Show($"Error: LibreOffice terminó pero no generó el archivo esperado.\nContenido carpeta temp:\n{lista}");
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error crítico: " + ex.Message);
                }
                finally
                {
                    // Limpieza: Borrar carpeta temp
                    try { if (Directory.Exists(carpetaTempSistema)) Directory.Delete(carpetaTempSistema, true); } catch { }
                    lblStatus.Text = "Listo.";
                }
            }
        }

        // ==========================================
        // HELPERS (FUNCIONES DE APOYO)
        // ==========================================

        private void MatarProcesosLibreOffice()
        {
            try
            {
                foreach (var proc in Process.GetProcessesByName("soffice")) proc.Kill();
                foreach (var proc in Process.GetProcessesByName("soffice.bin")) proc.Kill();
            }
            catch { /* Ignorar errores de permisos */ }
        }

        private void ConvertirConLibreOffice(string exePath, string inputFile, string outputDir, bool esExcel)
        {
            string formato = esExcel ? "xlsx" : "docx";

            // Perfil temporal limpio para evitar errores de configuración
            string userProfileTemp = Path.Combine(outputDir, "user_profile");

            // Argumentos blindados
            string args = $"-env:UserInstallation=\"file:///{userProfileTemp.Replace(@"\", "/")}\" --headless --convert-to {formato} --outdir \"{outputDir}\" \"{inputFile}\"";

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                FileName = exePath,
                Arguments = args,
                WindowStyle = ProcessWindowStyle.Hidden,
                CreateNoWindow = true,
                UseShellExecute = false,
                RedirectStandardError = true
            };

            using (Process process = Process.Start(startInfo))
            {
                process.WaitForExit();
            }
        }

        private string BuscarLibreOffice()
        {
            string[] rutas = {
                @"C:\Program Files\LibreOffice\program\soffice.exe",
                @"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
            };
            foreach (var r in rutas) if (File.Exists(r)) return r;
            return null;
        }

        // Botones de UI (Seleccionar, Limpiar, etc.)
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
    }
}