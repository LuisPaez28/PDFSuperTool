using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

// ==========================================
// PARTE 1: Librerías Nativas de C# (Rápidas para Separar/Unir)
// ==========================================
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
        // 1. SEPARAR PDF (Usamos C# nativo porque es rapidísimo)
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
                string carpeta = Path.GetDirectoryName(sfd.FileName);
                string baseName = Path.GetFileNameWithoutExtension(pdfPath);

                try
                {
                    lblStatus.Text = "Separando (Motor C#)...";
                    await Task.Run(() =>
                    {
                        using (SharpDoc input = PdfReader.Open(pdfPath, PdfDocumentOpenMode.Import))
                        {
                            for (int i = 0; i < input.PageCount; i++)
                            {
                                using (SharpDoc output = new SharpDoc())
                                {
                                    output.AddPage(input.Pages[i]);
                                    output.Save(Path.Combine(carpeta, $"{baseName}_pag_{i + 1}.pdf"));
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
        // 2. UNIR PDF (Usamos C# nativo porque es rapidísimo)
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
                    lblStatus.Text = "Uniendo (Motor C#)...";
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
        // 3. CONVERTIR (Usamos PYTHON porque es más inteligente)
        // ==========================================
        private async void BtnConvertir_Click(object sender, RoutedEventArgs e)
        {
            string pdfPath = txtPathConvert.Text;
            bool esExcel = chkEsExcel.IsChecked ?? false;

            if (string.IsNullOrEmpty(pdfPath) || !File.Exists(pdfPath))
            {
                MessageBox.Show("Selecciona el PDF a convertir.");
                return;
            }

            // Validar si Python existe
            string pythonPath = ObtenerRutaPython();
            if (string.IsNullOrEmpty(pythonPath))
            {
                MessageBox.Show("No encontré Python instalado.\nInstálalo y ejecuta: pip install pdf2docx pandas pdfplumber openpyxl");
                return;
            }

            SaveFileDialog sfd = new SaveFileDialog
            {
                Title = "Guardar archivo convertido",
                Filter = esExcel ? "Excel (*.xlsx)|*.xlsx" : "Word (*.docx)|*.docx",
                FileName = Path.GetFileNameWithoutExtension(pdfPath)
            };

            if (sfd.ShowDialog() == true)
            {
                string rutaSalida = sfd.FileName;
                lblStatus.Text = "Procesando con Python (IA)...";

                try
                {
                    await Task.Run(() =>
                    {
                        EjecutarScriptPython(pythonPath, pdfPath, rutaSalida, esExcel);
                    });

                    MessageBox.Show("¡Conversión Exitosa!");
                    try { Process.Start("explorer.exe", $"/select,\"{rutaSalida}\""); } catch { }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error en Python:\n" + ex.Message);
                }
                finally
                {
                    lblStatus.Text = "Listo.";
                }
            }
        }

        // ==========================================
        // LÓGICA DE PUENTE C# <-> PYTHON
        // ==========================================
        private void EjecutarScriptPython(string pythonExe, string inputPdf, string output, bool esExcel)
        {
            string scriptContent = GenerarCodigoPython(esExcel);
            string tempScriptPath = Path.Combine(Path.GetTempPath(), "convertidor_temp.py");

            // Escribimos el script .py en el disco temporalmente
            File.WriteAllText(tempScriptPath, scriptContent);

            // Argumentos con comillas para proteger espacios
            string argumentos = $"\"{tempScriptPath}\" \"{inputPdf}\" \"{output}\"";

            ProcessStartInfo start = new ProcessStartInfo
            {
                FileName = pythonExe,
                Arguments = argumentos,
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            using (Process process = Process.Start(start))
            {
                string errors = process.StandardError.ReadToEnd();
                process.WaitForExit();

                // Borramos el script temporal
                try { File.Delete(tempScriptPath); } catch { }

                if (process.ExitCode != 0)
                {
                    throw new Exception($"El script de Python falló:\n{errors}");
                }
            }
        }

        private string GenerarCodigoPython(bool paraExcel)
        {
            if (!paraExcel)
            {
                // Script para WORD (pdf2docx)
                return @"
import sys
from pdf2docx import Converter

def main(pdf_file, docx_file):
    try:
        cv = Converter(pdf_file)
        cv.convert(docx_file, start=0, end=None)
        cv.close()
    except Exception as e:
        print(f'ERROR: {e}', file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
";
            }
            else
            {
                // Script para EXCEL (pdfplumber + pandas)
                return @"
import sys
import pdfplumber
import pandas as pd

def main(pdf_file, xlsx_file):
    try:
        all_tables = []
        with pdfplumber.open(pdf_file) as pdf:
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    # Ignorar tablas vacías o rotas
                    if table:
                        # Convertir a DataFrame. Usamos la fila 0 como headers
                        df = pd.DataFrame(table[1:], columns=table[0])
                        all_tables.append(df)
        
        if not all_tables:
            print('No se detectaron tablas claras en el PDF.', file=sys.stderr)
            sys.exit(1)

        final_df = pd.concat(all_tables, ignore_index=True)
        final_df.to_excel(xlsx_file, index=False)

    except Exception as e:
        print(f'ERROR: {e}', file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    main(sys.argv[1], sys.argv[2])
";
            }
        }

        private string ObtenerRutaPython()
        {
            // Intenta detectar el comando 'python' global
            try
            {
                ProcessStartInfo psi = new ProcessStartInfo("python", "--version") { UseShellExecute = false, CreateNoWindow = true };
                Process.Start(psi).WaitForExit();
                return "python";
            }
            catch
            {
                // Búsqueda manual si no está en el PATH
                string[] rutas = { @"C:\Python39\python.exe", @"C:\Python310\python.exe", @"C:\Python311\python.exe", @"C:\Python312\python.exe" };
                foreach (var r in rutas) if (File.Exists(r)) return r;
                return null;
            }
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
                foreach (string f in ofd.FileNames) { archivosParaUnir.Add(f); listArchivosUnir.Items.Add(Path.GetFileName(f)); }
            }
        }
        private void BtnLimpiar_Click(object sender, RoutedEventArgs e)
        {
            archivosParaUnir.Clear(); listArchivosUnir.Items.Clear(); txtPathSplit.Clear(); txtPathConvert.Clear(); lblStatus.Text = "Limpiado.";
        }
        private void BtnSeleccionarConvert_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog { Filter = "PDF|*.pdf" };
            if (ofd.ShowDialog() == true) txtPathConvert.Text = ofd.FileName;
        }
    }
}