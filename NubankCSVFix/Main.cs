using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace NubankCSVFix {
    public partial class Main : Form {
        public static readonly char[] CellChars = new char[] { 'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z' };

        public Main() {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e) {
            var openFileDialog = new OpenFileDialog {
                Filter = @"Nubank CSV|*.csv"
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK) {
                textBox1.Text = openFileDialog.FileName;
            }
        }

        private void Button2_Click(object sender, EventArgs e) {
            if (!File.Exists(textBox1.Text)) {
                MessageBox.Show(@"Falha ao abrir, o arquivo não existe!", "Erro no arquivo", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            try {
                using (var workbook = new XLWorkbook()) {
                    var worksheet = workbook.Worksheets.Add(@"Nubank CSV");
                    using (var fileStream = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                        using (var streamReader = new StreamReader(fileStream)) {
                            string lineContent;
                            var currentRow = 1;
                            while ((lineContent = streamReader.ReadLine()) != null) {
                                var lineData = lineContent.Split(',');
                                for (var i = 0; i < lineData.Length; i++) {
                                    worksheet.Cell(CellChars[i] + currentRow.ToString()).Value = lineData[i];
                                }
                                currentRow++;
                            }
                        }
                    }
                    var saveFileDialog = new SaveFileDialog {
                        Filter = @"Excel Worksheets|*.xlsx"
                    };
                    if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                        workbook.SaveAs(saveFileDialog.FileName);
                    }
                    if (MessageBox.Show(@"Arquivo convertido, deseja abri-lo ?", @"Conversão bem sucedida",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                        Process.Start(saveFileDialog.FileName);
                    }
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, @"Erro ao converter", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void Button3_Click(object sender, EventArgs e) {
            if (!File.Exists(textBox1.Text)) {
                MessageBox.Show(@"Falha ao abrir, o arquivo não existe!", @"Erro no arquivo", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return;
            }
            try {
                using (var workbook = new XLWorkbook()) {
                    var worksheet = workbook.Worksheets.Add(@"Nubank CSV");
                    using (var fileStream = new FileStream(textBox1.Text, FileMode.Open, FileAccess.Read, FileShare.Read)) {
                        using (var streamReader = new StreamReader(fileStream)) {
                            string lineContent;
                            var currentRow = 1;
                            var currentColumn = 0;
                            var sumList = new List<string>();
                            var nextColor = true;
                            while ((lineContent = streamReader.ReadLine()) != null) {
                                var lineData = lineContent.Split(',');
                                for (currentColumn = 0; currentColumn < lineData.Length; currentColumn++) {
                                    var cell = worksheet.Cell(CellChars[currentColumn] + currentRow.ToString());
                                    cell.Value = lineData[currentColumn];
                                    if (currentColumn == 3 && double.TryParse(lineData[currentColumn], out var result)) {
                                        if (result > 0) {
                                            sumList.Add(CellChars[currentColumn] + result.ToString("{0:0.00}"));
                                            cell.Value = result.ToString("{0:0.00}");
                                        }
                                    }
                                    cell.Style.Fill.BackgroundColor = nextColor ? XLColor.FromHtml("#D9E1F2") : XLColor.White;
                                }
                                nextColor = !nextColor;
                                currentRow++;
                            }
                            var range = worksheet.Range(1, 1, currentRow, currentColumn);
                            var table = range.CreateTable();
                            table.Theme = XLTableTheme.TableStyleLight12;
                            currentRow += 2;
                            currentColumn--;
                            var cellCalcule = "=SUM(";
                            foreach (var cell in sumList) {
                                cellCalcule += $"{cell},";
                            }
                            cellCalcule = cellCalcule.Remove(cellCalcule.Length - 1, 1);
                            cellCalcule += ")";
                            worksheet.Cell(CellChars[currentColumn] + currentRow.ToString()).FormulaA1 = cellCalcule;
                            worksheet.Cell(CellChars[currentColumn] + currentRow.ToString()).Style.Fill.BackgroundColor = XLColor.FromHtml("#FFFF00");
                            currentRow ++;
                            worksheet.Cell(CellChars[currentColumn] + currentRow.ToString()).Value = "Versão 1.3";
                            worksheet.Column("A").Width = 12;
                            worksheet.Column("B").Width = 18;
                            worksheet.Column("C").Width = 30;
                            worksheet.Column("D").Width = 10;
                        }
                    }
                    var saveFileDialog = new SaveFileDialog {
                        Filter = @"Excel Worksheets|*.xlsx"
                    };
                    if (saveFileDialog.ShowDialog() == DialogResult.OK) {
                        workbook.SaveAs(saveFileDialog.FileName);
                    }
                    if (MessageBox.Show(@"Arquivo convertido, deseja abri-lo ?", @"Conversão bem sucedida",
                            MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes) {
                        Process.Start(saveFileDialog.FileName);
                    }
                }
            }
            catch (Exception ex) {
                MessageBox.Show(ex.Message, @"Erro ao converter", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }
    }
}
