using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace GeradorDocumentos
{
    static class Program
    {
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new MainForm());
        }
    }

    public class MainForm : Form
    {
        private TextBox txtExcel = null!;
        private TextBox txtWord = null!;
        private Button btnSelectExcel = null!;
        private Button btnSelectWord = null!;
        private Button btnGerar = null!;
        private Label lblStatus = null!;
        private ProgressBar progressBar = null!;

        public MainForm()
        {
            InitializeComponents();
        }

        private void InitializeComponents()
        {
            this.Text = "Gerador de Documentos Word";
            this.Size = new Size(580, 320);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;

            Label lblExcel = new Label() { Text = "Planilha Excel (.xlsx):", Location = new Point(20, 20), AutoSize = true };
            txtExcel = new TextBox() { Location = new Point(20, 45), Width = 400, ReadOnly = true };
            btnSelectExcel = new Button() { Text = "Procurar...", Location = new Point(430, 43), Width = 110 };
            btnSelectExcel.Click += (s, e) => SelecionarArquivo(txtExcel, "Planilha Excel (*.xlsx)|*.xlsx");

            Label lblWord = new Label() { Text = "Modelo Word (.docx):", Location = new Point(20, 85), AutoSize = true };
            txtWord = new TextBox() { Location = new Point(20, 110), Width = 400, ReadOnly = true };
            btnSelectWord = new Button() { Text = "Procurar...", Location = new Point(430, 108), Width = 110 };
            btnSelectWord.Click += (s, e) => SelecionarArquivo(txtWord, "Modelo Word (*.docx)|*.docx");

            btnGerar = new Button()
            {
                Text = "Gerar Documentos",
                Location = new Point(20, 155),
                Width = 520,
                Height = 40,
                BackColor = System.Drawing.Color.FromArgb(0, 122, 204),
                ForeColor = System.Drawing.Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnGerar.Click += BtnGerar_Click;

            progressBar = new ProgressBar() { Location = new Point(20, 210), Width = 520, Height = 20 };
            lblStatus = new Label() { Text = "Selecione os arquivos para iniciar.", Location = new Point(20, 240), AutoSize = true };

            this.Controls.Add(lblExcel);
            this.Controls.Add(txtExcel);
            this.Controls.Add(btnSelectExcel);

            this.Controls.Add(lblWord);
            this.Controls.Add(txtWord);
            this.Controls.Add(btnSelectWord);

            this.Controls.Add(btnGerar);
            this.Controls.Add(progressBar);
            this.Controls.Add(lblStatus);
        }

        private void SelecionarArquivo(TextBox campoTexto, string filtro)
        {
            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                dialog.Filter = filtro;
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    campoTexto.Text = dialog.FileName;
                }
            }
        }

        private async void BtnGerar_Click(object? sender, EventArgs e)
        {
            string excelPath = txtExcel.Text;
            string templatePath = txtWord.Text;

            if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
            {
                MessageBox.Show("Selecione um arquivo Excel válido.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (string.IsNullOrWhiteSpace(templatePath) || !File.Exists(templatePath))
            {
                MessageBox.Show("Selecione um modelo Word válido.", "Atenção", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            btnGerar.Enabled = false;
            btnSelectExcel.Enabled = false;
            btnSelectWord.Enabled = false;
            lblStatus.Text = "Lendo planilha...";
            progressBar.Value = 0;

            string outputFolder = Path.Combine(Path.GetDirectoryName(excelPath) ?? AppDomain.CurrentDomain.BaseDirectory, "Documentos_Gerados");
            Directory.CreateDirectory(outputFolder);

            try
            {
                await Task.Run(() =>
                {
                    using (var workbook = new XLWorkbook(excelPath))
                    {
                        var worksheet = workbook.Worksheet(1);
                        var rangeUsed = worksheet.RangeUsed();

                        if (rangeUsed == null)
                        {
                            throw new Exception("A planilha selecionada está vazia.");
                        }

                        var firstRow = worksheet.Row(1);
                        int lastCol = rangeUsed.LastColumnUsed()?.ColumnNumber() ?? 0;

                        // Mapeia o número da coluna para o nome do cabeçalho
                        var colunas = new Dictionary<int, string>();
                        for (int col = 1; col <= lastCol; col++)
                        {
                            string nomeHeader = firstRow.Cell(col).GetFormattedString().Trim();
                            if (string.IsNullOrEmpty(nomeHeader))
                            {
                                nomeHeader = firstRow.Cell(col).Value.ToString().Trim();
                            }

                            if (!string.IsNullOrEmpty(nomeHeader))
                            {
                                colunas[col] = nomeHeader;
                            }
                        }

                        var rows = rangeUsed.RowsUsed().Skip(1).ToList();
                        int total = rows.Count;
                        int processados = 0;

                        foreach (var row in rows)
                        {
                            var dadosLinha = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                            foreach (var kvp in colunas)
                            {
                                int colIndex = kvp.Key;
                                string nomeColuna = kvp.Value;
                                string valorCelula = row.Cell(colIndex).GetFormattedString().Trim();
                                if (string.IsNullOrEmpty(valorCelula))
                                {
                                    valorCelula = row.Cell(colIndex).Value.ToString().Trim();
                                }
                                dadosLinha[nomeColuna] = valorCelula;
                            }

                            string identificador = dadosLinha.ContainsKey("NOME") && !string.IsNullOrWhiteSpace(dadosLinha["NOME"])
                                ? dadosLinha["NOME"]
                                : $"Documento_{processados + 1}";

                            string nomeArquivoLimpo = string.Join("_", identificador.Split(Path.GetInvalidFileNameChars()));
                            string outputPath = Path.Combine(outputFolder, $"{nomeArquivoLimpo}.docx");

                            File.Copy(templatePath, outputPath, overwrite: true);

                            // Aplica a substituição estruturada no Word
                            SubstituirVariaveisWord(outputPath, dadosLinha);

                            processados++;
                            int percentual = total > 0 ? (int)((processados / (double)total) * 100) : 100;

                            this.Invoke(new Action(() =>
                            {
                                progressBar.Value = percentual;
                                lblStatus.Text = $"Gerando {processados} de {total}...";
                            }));
                        }
                    }
                });

                lblStatus.Text = "Processo concluído com sucesso!";
                MessageBox.Show($"Arquivos gerados com sucesso na pasta:\n{outputFolder}", "Sucesso", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                lblStatus.Text = "Erro durante o processamento.";
                MessageBox.Show($"Erro ao processar: {ex.Message}", "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnGerar.Enabled = true;
                btnSelectExcel.Enabled = true;
                btnSelectWord.Enabled = true;
            }
        }

        private void SubstituirVariaveisWord(string filePath, Dictionary<string, string> dados)
        {
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, true))
            {
                var mainPart = wordDoc.MainDocumentPart;
                if (mainPart == null) return;

                var body = mainPart.Document?.Body;
                if (body != null)
                {
                    SubstituirNoElemento(body, dados);
                }

                foreach (var header in mainPart.HeaderParts)
                {
                    if (header.Header != null) SubstituirNoElemento(header.Header, dados);
                }
                foreach (var footer in mainPart.FooterParts)
                {
                    if (footer.Footer != null) SubstituirNoElemento(footer.Footer, dados);
                }

                mainPart.Document?.Save();
            }
        }

        private void SubstituirNoElemento(OpenXmlElement parentElement, Dictionary<string, string> dados)
        {
            var paragrafos = parentElement.Descendants<Paragraph>().ToList();

            foreach (var p in paragrafos)
            {
                string textoCompleto = p.InnerText;
                if (string.IsNullOrWhiteSpace(textoCompleto)) continue;

                bool alterado = false;
                string textoModificado = textoCompleto;

                foreach (var item in dados)
                {
                    if (string.IsNullOrEmpty(item.Key)) continue;

                    string chave = item.Key.Trim();
                    string valor = item.Value ?? "";

                    string tagChaves = $"{{{{{chave}}}}}"; // {{NOME}}
                    string tagColchetes = $"[{chave}]";    // [NOME]

                    if (textoModificado.Contains(tagChaves, StringComparison.OrdinalIgnoreCase))
                    {
                        textoModificado = Regex.Replace(textoModificado, Regex.Escape(tagChaves), valor, RegexOptions.IgnoreCase);
                        alterado = true;
                    }

                    if (textoModificado.Contains(tagColchetes, StringComparison.OrdinalIgnoreCase))
                    {
                        textoModificado = Regex.Replace(textoModificado, Regex.Escape(tagColchetes), valor, RegexOptions.IgnoreCase);
                        alterado = true;
                    }

                    // Caso a coluna no Excel seja 'CNPJ' e no Word a tag esteja como '{{CPF}}'
                    if (chave.Equals("CNPJ", StringComparison.OrdinalIgnoreCase) && textoModificado.Contains("{{CPF}}", StringComparison.OrdinalIgnoreCase))
                    {
                        textoModificado = Regex.Replace(textoModificado, Regex.Escape("{{CPF}}"), valor, RegexOptions.IgnoreCase);
                        alterado = true;
                    }
                }

                if (alterado)
                {
                    var runs = p.Elements<Run>().ToList();
                    if (runs.Count > 0)
                    {
                        var primeiroRun = runs.First();

                        var textElement = primeiroRun.Elements<Text>().FirstOrDefault();
                        if (textElement == null)
                        {
                            textElement = new Text();
                            primeiroRun.AppendChild(textElement);
                        }

                        textElement.Text = textoModificado;
                        textElement.Space = SpaceProcessingModeValues.Preserve;

                        for (int i = 1; i < runs.Count; i++)
                        {
                            runs[i].Remove();
                        }
                    }
                }
            }
        }
    }
}