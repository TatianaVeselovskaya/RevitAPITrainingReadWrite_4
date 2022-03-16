using NPOI.SS.UserModel;
using NPOI.SS.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPITrainingReadWrite_4
{
    public static class SheetExts
    {
        //метод для заполнения листа даннными
        public static void SetCellValue<T>(this ISheet sheet, int rowIndex, int columnIndex, T value)
        // адрес ячейки: номер строки, номер столбца, запись значения
        {
            var cellReference = new CellReference(rowIndex, columnIndex);// ссылка на ячейку
            var row = sheet.GetRow(cellReference.Row); // проверяем наличме строки
            if (row == null)                            // если строки нет
                row = sheet.CreateRow(cellReference.Row); // создаем 

            var cell = row.GetCell(cellReference.Col); // создаем ссылку на ячейку и в выбранной строке, конкренная ячейка
            if (cell == null)                        // если ячейка пустая
                cell = row.CreateCell(cellReference.Col); // создаем 

            if (value is string) //если передан текст то записываем данное значение как текст
            {
                cell.SetCellValue((string)(object)value); // делаем преобразование в object потом в string, на прямую нельзя
            }
            else if (value is double) // если значение double
            {
                cell.SetCellValue((double)(object)value);
            }
            else if (value is int)  //если значение целочисленное
            {
                cell.SetCellValue((int)(object)value);
            }
        }    
    }
}
