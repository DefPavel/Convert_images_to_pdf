using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using ExcelApp = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace ConsoleApp2
{
    internal class Program
    {
        static DataTable ReadExcel()
        {

            ExcelApp.Application excelApp = null;
            ExcelApp.Workbooks workbooks = null;

            DataRow myNewRow;
            DataTable myTable;

            try
            {
                excelApp = new ExcelApp.Application();
                workbooks = excelApp.Workbooks;

                if (excelApp == null)
                {
                    Console.WriteLine("Упс,Кажется Excel не установлен");

                }
                ExcelApp.Workbook excelBook = excelApp.Workbooks.Open($"{Environment.CurrentDirectory}\\testdata.xlsx");
                ExcelApp._Worksheet excelSheet = excelBook.Sheets[1];
                ExcelApp.Range excelRange = excelSheet.UsedRange;

                int rows = excelRange.Rows.Count;
                int cols = excelRange.Columns.Count;

                //Set DataTable Name and Columns Name
                myTable = new DataTable("MyDataTable");
                myTable.Columns.Add("FullName", typeof(string));




                //Пропускаем первую строку для хедара
                for (int i = 1; i <= rows; i++)
                {
                    myNewRow = myTable.NewRow();
                    myNewRow["FullName"] = excelRange.Cells[i, 3].Value2.ToString(); 

                    myTable.Rows.Add(myNewRow);
                }

                return myTable;

            }
            finally
            {

                if (workbooks != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(workbooks);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
        }

        static void UpdateExcel(List<Persons> persons)
        {
            Microsoft.Office.Interop.Excel.Application application = null;
            Microsoft.Office.Interop.Excel.Workbooks workbooks = null;
            Microsoft.Office.Interop.Excel.Workbook workbook = null;
            Microsoft.Office.Interop.Excel.Sheets worksheets = null;
            Microsoft.Office.Interop.Excel.Worksheet worksheet = null;


            Microsoft.Office.Interop.Excel.Range cell = null;

            try
            {
                application = new Microsoft.Office.Interop.Excel.Application
                {
                    Visible = true
                };
                workbooks = application.Workbooks;
                workbook = workbooks.Add();
                worksheets = workbook.Worksheets; //получаем доступ к коллекции рабочих листов
                worksheet = worksheets.Item[1];//получаем доступ к первому листу

                for (int i = 0; i < persons.Count; i++)
                {
                    cell = worksheet.Cells[i + 1, 3];
                    cell.Value = persons[i].FullName; //

                    cell = worksheet.Cells[i + 1, 4];
                    cell.Value = persons[i].Birthday; //

                    cell = worksheet.Cells[i + 1, 5];
                    cell.Value = persons[i].DepartmentPosition; //

                    cell = worksheet.Cells[i + 1, 6];
                    cell.Value = persons[i].Education; //

                    cell = worksheet.Cells[i + 1, 7];
                    cell.Value = persons[i].Phone; //
                }

                Console.WriteLine("Done!");
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }

        static async Task Main(string[] args)
        {
            List<Persons> persons = new List<Persons>();
            ReadExcel();

            foreach (DataRow dr in ReadExcel().Rows)
            {
                
                var FullName = dr[0].ToString().Split(' ');
                var person = await FireService.getPersons(FullName[0] , FullName[1], FullName[2]);
                persons.Add(person);

                var doc = await FireService.getDocuments(person.Id);

                var path = $"{person.FullName}";
                if(path.Any())
                {
                    // Создаём папку с ФИО
                    Directory.CreateDirectory(path);
                    // Заносим в папку все файлы
                    foreach (var item in doc)
                    {
                        File.WriteAllBytes($"{path}\\{item.Name}.jpg", item.Data);
                    }
                }
              
            }
            UpdateExcel(persons);
            /*var doc = await FireService.getDocuments(7042);

            var path = $"Божок Дарья Николаевна";
            if (path.Any())
            {
                Directory.CreateDirectory(path);

                foreach (var item in doc)
                {
                    File.WriteAllBytes($"{path}\\{item.Name}.jpg", item.Data);
                }
            }
            */




            Console.WriteLine("Нажмите любую клавишу для завершения работы...");

            Console.ReadKey();

        }
    }
}
