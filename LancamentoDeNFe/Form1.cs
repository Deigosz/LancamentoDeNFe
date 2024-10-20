using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace LancamentoDeNFe
{
    public partial class Form1 : Form
    {
        private string caminhoArquivo;

        public Form1()
        {
            InitializeComponent();
            dataGridView1.CellEndEdit += dataGridView1_CellEndEdit;
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                Filter = "Excel Files|*.xlsx",
                Title = "Selecione o arquivo Excel"
            };

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                caminhoArquivo = ofd.FileName;
                CarregarExcel(ofd.FileName);
                MessageBox.Show("Arquivo carregado com sucesso!");
            }
        }

        private void CarregarExcel(string Path) 
        {
            try
            {
                using (var package = new ExcelPackage(new FileInfo(Path)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    DataTable dt = new DataTable();

                    for (int i = worksheet.Dimension.Start.Column; i <= worksheet.Dimension.End.Column; i++)
                    {
                        dt.Columns.Add(worksheet.Cells[1, i].Text);
                    }

                    for (int i = worksheet.Dimension.Start.Row + 1; i <= worksheet.Dimension.End.Row; i++)
                    {
                        DataRow row = dt.NewRow();
                        for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                        {
                            row[j - 1] = worksheet.Cells[i, j].Text;
                        }
                        dt.Rows.Add(row);
                    }
                    dataGridView1.DataSource = dt;
                }
                    
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao carregar o arquivo: {ex.Message}");
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            SalvarCelulaEditada(e.RowIndex, e.ColumnIndex);
        }

        private void SalvarCelulaEditada(int rowIndex, int columnIndex)
        {
            if (string.IsNullOrEmpty(caminhoArquivo))
            {
                MessageBox.Show("Nenhum arquivo Excel foi carregado.");
                return;
            }
            try
            {
                using (var package = new ExcelPackage(new FileInfo(caminhoArquivo)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0];
                    var valorEditado = dataGridView1.Rows[rowIndex].Cells[columnIndex].Value?.ToString();
                    worksheet.Cells[rowIndex + 2, columnIndex + 1].Value = valorEditado;
                    package.Save();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Erro ao salvar o arquivo: {ex.Message}");
            }
        }
    }
}
