using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using SaveFileDialog = System.Windows.Forms.SaveFileDialog;

namespace RevitAPITrainingReadWrite_4
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            string roomInfo = string.Empty;

            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx");
            // сохранение файла на рабочем столе, название файла
            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))// создание файла
            {
                IWorkbook workbook = new HSSFWorkbook(); // создание книги
                ISheet sheet = workbook.CreateSheet("Лист1"); // создание листа

                int rowIndex = 0;                       //1 строка
                foreach (var room in rooms)             // проходимся по каждому помещению в модели
                {
                    sheet.SetCellValue(rowIndex, columnIndex: 0, room.Name); // записываю 1 значение в 1 столбец
                    sheet.SetCellValue(rowIndex, columnIndex: 1, room.Number); // 2 столбец
                    sheet.SetCellValue(rowIndex, columnIndex: 2, room.Area); // 3 столбец
                    rowIndex++;                         //переход на следующую строку
                }

                workbook.Write(stream);                 // закрываем запись в ехселе
                workbook.Close();
            }

            System.Diagnostics.Process.Start(excelPath); // запускаем файл ехселя автоматически

            return Result.Succeeded;
        }
    }
}

