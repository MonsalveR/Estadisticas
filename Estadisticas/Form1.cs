using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.CSharp;
using SpreadsheetLight;

namespace Estadisticas
{
    public partial class frmEstadisticas : Form
    {
        //private string path = @"C:\Users\carlo\source\repos\Archivos.xlsx";
        public frmEstadisticas()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void btnImportar_Click(object sender, EventArgs e)
        {
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path;
                path = openFileDialog1.FileName;
                SLDocument sl = new SLDocument(path);
                int row = 2;
                List<Producto> listProducto = new List<Producto>();
                while (!string.IsNullOrEmpty(sl.GetCellValueAsString(row, 1)))
                {
                    Producto producto = new Producto();
                    producto.Codigo = sl.GetCellValueAsString(row, 1);
                    producto.Nombre = sl.GetCellValueAsString(row, 2);
                    producto.Precio = sl.GetCellValueAsInt32(row, 3);
                    producto.Cantidad = sl.GetCellValueAsInt32(row, 4);

                    listProducto.Add(producto);

                    row += 1;
                }

                dgvDatos.DataSource = listProducto;
            }
        }

        private void btnGuardar_Click(object sender, EventArgs e)
        {
            if(saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string path;
                SLDocument sl = new SLDocument();
                int iR = 2; int iC = 1;
                foreach (DataGridViewColumn clum in dgvDatos.Columns)
                {
                    sl.SetCellValue(1, iC, clum.HeaderText.ToString());
                    iC += 1;
                }
                foreach (DataGridViewRow row in dgvDatos.Rows)
                {
                    sl.SetCellValue(iR, 1, row.Cells[0].Value.ToString());
                    sl.SetCellValue(iR, 2, row.Cells[1].Value.ToString());
                    sl.SetCellValue(iR, 3, row.Cells[2].Value.ToString());
                    sl.SetCellValue(iR, 4, row.Cells[3].Value.ToString());
                    iR += 1;
                }
                path = saveFileDialog1.FileName;
                sl.SaveAs(path);
            }
           
        }

        private void btnEliminar_Click(object sender, EventArgs e)
        {
            SLDocument sl = new SLDocument(@"C:\Users\carlo\source\repos\Archivos.xlsx");
            int row = 2;
            List<Producto> listProducto = new List<Producto>();
            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(row, 1)))
            {
                Producto producto = new Producto();
                producto.Codigo = sl.GetCellValueAsString(row, 1);
                producto.Nombre = sl.GetCellValueAsString(row, 2);
                producto.Precio = sl.GetCellValueAsInt32(row, 3);
                producto.Cantidad = sl.GetCellValueAsInt32(row, 4);

                listProducto.Add(producto);

                row += 1;
            }
            int rows;
            int iR = 1;
            rows = dgvDatos.CurrentRow.Index;
            listProducto.Clear();
            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(iR, 1)))
            {
                Producto producto = new Producto();
                if (rows == iR)
                {
                    iR += 1;
                }else
                {
                    producto.Codigo = sl.GetCellValueAsString(iR, 1);
                    producto.Nombre = sl.GetCellValueAsString(iR, 2);
                    producto.Precio = sl.GetCellValueAsInt32(iR, 3);
                    producto.Cantidad = sl.GetCellValueAsInt32(iR, 4);

                    listProducto.Add(producto);

                    iR += 1;
                }
            }
            dgvDatos.DataSource = listProducto;
        }
        private void frmEstadisticas_Load(object sender, EventArgs e)
        {
            SLDocument sl = new SLDocument(@"C:\Users\carlo\source\repos\Archivos.xlsx");
            int row = 2;
            List<Producto> listProducto = new List<Producto>();
            while (!string.IsNullOrEmpty(sl.GetCellValueAsString(row, 1)))
            {
                Producto producto = new Producto();
                producto.Codigo = sl.GetCellValueAsString(row, 1);
                producto.Nombre = sl.GetCellValueAsString(row, 2);
                producto.Precio = sl.GetCellValueAsInt32(row, 3);
                producto.Cantidad = sl.GetCellValueAsInt32(row, 4);

                listProducto.Add(producto);

                row += 1;
            }

            dgvDatos.DataSource = listProducto;
        }
    }
}
