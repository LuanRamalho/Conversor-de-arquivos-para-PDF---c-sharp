using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
// Aliases para evitar conflitos
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using iTextDocument = iText.Layout.Document;
using iTextParagraph = iText.Layout.Element.Paragraph;
using iTextPdfWriter = iText.Kernel.Pdf.PdfWriter;
using iTextPdfDocument = iText.Kernel.Pdf.PdfDocument;
using Spire.Presentation;

namespace ConversorPDF
{
    public partial class MainForm : Form
    {
        private TextBox? txtFilePath;
        private Button? btnSelect, btnConvert;
        private Label? lblTitle;

        public MainForm()
        {
            ConfigurarInterface();
        }

        private void ConfigurarInterface()
        {
            this.Text = "Conversor Profissional para PDF - .NET 10";
            this.Size = new Size(500, 350);
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.FromArgb(255, 128, 245); // Rosa vibrante

            lblTitle = new Label
            {
                Text = "Conversor de Arquivos para PDF",
                Font = new Font("Arial", 16, FontStyle.Bold),
                ForeColor = Color.White,
                BackColor = Color.FromArgb(76, 175, 80), // Verde
                TextAlign = ContentAlignment.MiddleCenter,
                Dock = DockStyle.Top,
                Height = 60
            };

            Label lblSelect = new Label
            {
                Text = "Selecione o arquivo (Docx, Xlsx, Pptx ou Txt):",
                Font = new Font("Arial", 11),
                Location = new Point(20, 80),
                AutoSize = true
            };

            txtFilePath = new TextBox
            {
                Location = new Point(20, 110),
                Width = 440,
                Font = new Font("Arial", 11),
                ReadOnly = true
            };

            btnSelect = new Button
            {
                Text = "Escolher Arquivo",
                Location = new Point(150, 150),
                Size = new Size(200, 40),
                Font = new Font("Arial", 11, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(2, 14, 128), // Azul
                ForeColor = Color.White
            };
            btnSelect.Click += (s, e) => SelectFile();

            btnConvert = new Button
            {
                Text = "CONVERTER COM FORMATAÇÃO",
                Location = new Point(100, 220),
                Size = new Size(300, 50),
                Font = new Font("Arial", 12, FontStyle.Bold),
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.FromArgb(3, 103, 1), // Verde Escuro
                ForeColor = Color.White
            };
            btnConvert.Click += (s, e) => ConvertFile();

            this.Controls.AddRange(new Control[] { lblTitle, lblSelect, txtFilePath, btnSelect, btnConvert });
        }

        private void SelectFile()
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Arquivos do Office e Texto|*.docx;*.xlsx;*.pptx;*.txt";
                if (ofd.ShowDialog() == DialogResult.OK) txtFilePath!.Text = ofd.FileName;
            }
        }

        private void ConvertFile()
        {
            string input = txtFilePath?.Text ?? "";
            if (string.IsNullOrEmpty(input)) return;

            using (SaveFileDialog sfd = new SaveFileDialog { Filter = "PDF Files|*.pdf" })
            {
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        Cursor = Cursors.WaitCursor;
                        ExecuteConversion(input, sfd.FileName);
                        MessageBox.Show("Conversão concluída com sucesso!", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro: {ex.Message}", "Falha na Conversão", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally { Cursor = Cursors.Default; }
                }
            }
        }

        private void ExecuteConversion(string input, string output)
        {
            string ext = Path.GetExtension(input).ToLower();

            switch (ext)
            {
                case ".docx": ConvertWord(input, output); break;
                case ".xlsx": ConvertExcel(input, output); break;
                case ".pptx": ConvertPowerPoint(input, output); break;
                case ".txt": ConvertText(input, output); break;
                default: throw new Exception("Formato não suportado.");
            }
        }

        private void ConvertWord(string input, string output)
        {
            dynamic wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!);
            try
            {
                var doc = wordApp.Documents.Open(input);
                doc.ExportAsFixedFormat(output, 17); // 17 é o valor para wdExportFormatPDF
                doc.Close(false);
            }
            finally { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); }
        }

        private void ConvertExcel(string input, string output)
        {
            dynamic excelApp = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
            try
            {
                var workbook = excelApp.Workbooks.Open(input);
                workbook.ExportAsFixedFormat(0, output); // 0 é o valor para xlTypePDF
                workbook.Close(false);
            }
            finally { excelApp.Quit(); Marshal.ReleaseComObject(excelApp); }
        }

        private void ConvertPowerPoint(string input, string output)
        {
            try
            {
                // Cria um objeto de apresentação independente do Microsoft Office
                Presentation presentation = new Presentation();
                
                // Carrega o seu arquivo .pptx
                presentation.LoadFromFile(input);
                
                // Salva diretamente como PDF
                // Isso ignora completamente o erro de "argumento" do Office
                presentation.SaveToFile(output, FileFormat.PDF);
                
                presentation.Dispose();
            }
            catch (Exception ex)
            {
                throw new Exception("Erro na conversão direta: " + ex.Message);
            }
        }

        private void ConvertText(string input, string output)
        {
            using (var writer = new iTextPdfWriter(output))
            using (var pdf = new iTextPdfDocument(writer))
            using (var document = new iTextDocument(pdf))
            {
                document.Add(new iTextParagraph(File.ReadAllText(input)));
            }
        }
    }
}