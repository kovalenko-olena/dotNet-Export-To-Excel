//using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExportToExcel
{
    internal class Program
    {
        static void ExportToExcel(List<PayVo> payVo)
        {
            // Загрузить Excel и затем создать новую пустую рабочую книгу.
            Excel.Application excelApp = new Excel.Application();
            excelApp.Workbooks.Add();
            // В этом примере используется единственный рабочий лист.
            Excel._Worksheet worksheet = excelApp.ActiveSheet;
            // Установить заголовки столбцов в ячейках.
            worksheet.Cells[1, 1] = "VoId";
            worksheet.Cells[1, 2] = "VoCd";
            worksheet.Cells[1, "C"] = "VoName";
            // Отобразить все данные из List<PayVo> на ячейки электронной таблицы.
            int row = 1;
            foreach (PayVo c in payVo)
            {
                row++;
                worksheet.Cells[row, "A"] = c.VoId;
                worksheet.Cells[row, "B"] = c.VoCd;
                worksheet.Cells[row, "C"] = c.VoName;
            }
            // Придать симпатичный вид табличным данным.
            worksheet.Range["A1"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic1);
            // Сохранить файл, завершить работу Excel и отобразить сообщение пользователю.
            worksheet.SaveAs($@"{Environment.CurrentDirectory}\VoCd.xlsx");
            excelApp.Quit();
            // Файл VoCd.xslx сохранен в папке приложения.
            Console.WriteLine("The VoCd.xslx file has been saved to your app folder");
        }

        static void Main(string[] args)
        {
            List<PayVo> vo = new List<PayVo>
            {
                new PayVo{VoId=1,VoCd=101,VoName="Paid time off"},
                new PayVo{VoId=2,VoCd=102,VoName="Unpaid leave"},
                new PayVo{VoId=3,VoCd=103,VoName="Paid vacation"},
                new PayVo{VoId=4,VoCd=104,VoName="Sick leave"},
                new PayVo{VoId=5,VoCd=105,VoName="Medical leave"},
                new PayVo{VoId=6,VoCd=106,VoName="Family leave"},
                new PayVo{VoId=7,VoCd=107,VoName="Short-term disability"},
                new PayVo{VoId=8,VoCd=108,VoName="Bereavement leave"}
            };
            ExportToExcel(vo);
        }
    }
}