using System;
using System.IO;
using NetOffice.ExcelApi;

namespace excel
{
    class Program
    {
        static void Main(string[] args)
        {
            if (!File.Exists("cliente.xls"))
            {
                CriarExcel();
            }
            LerExcel();
        }
        
        static void CriarExcel()
        {    
            Application ex = new Application();
            ex.Workbooks.Add();
            ex.Cells[1,1].Value = "Ford";
            ex.Cells[1,2].Value = "Fiesta";
            ex.Cells[1,3].Value = "1.8";

            ex.Cells[2,1].Value = "Fiat";
            ex.Cells[2,2].Value = "Pálio";
            ex.Cells[2,3].Value = "1.4";

            ex.ActiveWorkbook.SaveAs(@"C:\Users\39694603870\Desktop\Projetos\Projeto-6\cliente.xls");
            ex.Quit();

        }
    
        static void LerExcel()
        {
            Application ex = new Application();
            ex.Workbooks.Open(@"C:\Users\39694603870\Desktop\Projetos\Projeto-6\cliente.xls");
            string valor = ex.Cells[1,3].Value.ToString();
            Console.WriteLine(valor);
            ex.Quit();
        }
    
    
    }

}