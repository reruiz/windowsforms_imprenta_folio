using System;
using System.IO;
using OfficeOpenXml;
using System.Windows.Forms;


namespace windowsforms_imprenta_folio
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            textFilas.KeyPress += new KeyPressEventHandler(SoloNumerosEnteros_KeyPress);
            textColumnas.KeyPress += new KeyPressEventHandler(SoloNumerosEnteros_KeyPress);
            textInicio.KeyPress += new KeyPressEventHandler(SoloNumerosEnteros_KeyPress);
            textSalto.KeyPress += new KeyPressEventHandler(SoloNumerosEnteros_KeyPress);
            textMaximo.KeyPress += new KeyPressEventHandler(SoloNumerosEnteros_KeyPress);
        }

        private void buttonCancelar_Click(object sender, EventArgs e)
        {
            this.Close();
        }


        

        private void SoloNumerosEnteros_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Permite solo dígitos y teclas de control (como backspace)
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void buttonCrear_Click(object sender, EventArgs e)
        {
            int inicial = Convert.ToInt32(textInicio.Text); // Asume que tienes un TextBox llamado txtInicial
            int salto = Convert.ToInt32(textSalto.Text);     // Asume que tienes un TextBox llamado txtSalto
            int maximo = Convert.ToInt32(textMaximo.Text);   // Asume que tienes un TextBox llamado txtMaximo
            int filas = Convert.ToInt32(textFilas.Text);     // Asume que tienes un TextBox llamado txtFilas
            int columnas = Convert.ToInt32(textColumnas.Text); // Asume que tienes un TextBox llamado txtColumnas

            using (var package = new ExcelPackage())
            {
                var worksheet = package.Workbook.Worksheets.Add("Matriz");

                int valorActual = inicial;
                int correlativo = 1;
                int filaExcel = 1;

                while (valorActual <= maximo)
                {
                    for (int columna = 0; columna < columnas; columna++)
                    {
                        if (valorActual > maximo) break;

                        for (int fila = 0; fila < filas; fila++)
                        {
                            if (valorActual > maximo) break;

                            // Escribe primero el correlativo y luego el valor actual
                            worksheet.Cells[filaExcel + fila, columna * 2 + 1].Value = correlativo;
                            worksheet.Cells[filaExcel + fila, columna * 2 + 2].Value = valorActual;

                            valorActual += salto;
                            correlativo++;
                        }
                    }

                    filaExcel += filas + 2; // Salta dos filas después de cada matriz
                }

                // Guarda el archivo
                var saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel files (*.xlsx)|*.xlsx",
                    Title = "Guardar como"
                };

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    var fi = new FileInfo(saveFileDialog.FileName);
                    package.SaveAs(fi);
                }
            }
        }




    }
}
