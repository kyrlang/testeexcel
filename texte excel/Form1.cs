using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace texte_excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            try
            {
                List<Teste> testes = new List<Teste>();
                testes.Add(new Teste
                {
                    Texto = "Um",
                    Texto1 = "Dois"

                });

                testes.Add(new Teste
                {
                    Texto = "A",
                    Texto1 = "C"

                });

                testes.Add(new Teste
                {
                    Texto = "D",
                    Texto1 = "B"

                });

                testes.Add(new Teste
                {
                    Texto = "E",
                    Texto1 = "F"

                });

                FileStream file = new FileStream(ConfigurationManager.AppSettings["planilha"], FileMode.Open, FileAccess.Read);

                HSSFWorkbook hSSFWorkbook = new HSSFWorkbook(file);
                Sheet sheet = hSSFWorkbook.GetSheet("Planilha1");
                int indice = 1;

                foreach (var item in testes)
                {
                    GetCell(indice, 0, sheet);
                    GetCell(indice, 1, sheet);

                    sheet.GetRow(indice).GetCell(0).SetCellValue(item.Texto);
                    sheet.GetRow(indice).GetCell(1).SetCellValue(item.Texto1);
                    indice += 1;
                }

                using (FileStream fileWrite = new FileStream(@"c:\temp\testepreenchido.xls", FileMode.Create))
                {
                    hSSFWorkbook.Write(fileWrite);
                    file.Close();
                }
            }
            catch (Exception ex)
            {

                throw;
            }




        }

        private Cell GetCell(int linha, int coluna, Sheet sheet)
        {
            Row row;
            row = sheet.GetRow(linha);
            if (row == null)
                row = sheet.CreateRow(linha);

            Cell cell;
            cell = row.GetCell(coluna);
            if (cell == null)
                cell = row.CreateCell(coluna);

            return cell;
        }
       
    }

    public class Teste{

        public string Texto { get; set; }
        public string Texto1 { get; set; }
    }
}
