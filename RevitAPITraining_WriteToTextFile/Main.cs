using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Transactions;
using System.Windows.Forms;
using Transaction = Autodesk.Revit.DB.Transaction;

namespace RevitAPITraining_ReadWriteFile
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            #region Сбор помещений в коллектор
            var rooms = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .ToList();
            #endregion

            #region Запись параметров помещений в строку
            //string roomInfo = string.Empty;
            //foreach (var room in rooms)
            //{
            //    string roomName = room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
            //    roomInfo += $"{roomName}\t{room.Number}\t{room.Area}{Environment.NewLine}";
            //}
            #endregion

            #region Сохранение файла автоматически в папку по умолчанию
            //string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //string csvPath = Path.Combine(desktopPath, "roomInfo.csv");

            //File.WriteAllText(csvPath, roomInfo);
            #endregion

            #region Cохранение файла с выбором пути
            //var saveDialog = new SaveFileDialog
            //{
            //    OverwritePrompt = true,
            //    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            //    Filter = "All files (*.*)|*.*",
            //    FileName = "roomInfo.csv",
            //    DefaultExt = "*.csv",
            //};

            //string selectedFilePath = string.Empty;

            //if (saveDialog.ShowDialog() == DialogResult.OK)
            //{
            //    selectedFilePath = saveDialog.FileName;
            //}
            //if (string.IsNullOrEmpty(selectedFilePath))
            //{
            //    return Result.Cancelled;
            //}

            //File.WriteAllText(selectedFilePath, roomInfo);
            #endregion

            #region Чтение данных из текстового файла. Создание RoomData.cs
            ////Диалог открытия файла
            //OpenFileDialog openFileDialog = new OpenFileDialog();
            //openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            //openFileDialog.Filter = "All files (*.*)|*.*";

            ////Указание пути
            //string filePath = string.Empty;
            //if (openFileDialog.ShowDialog() == DialogResult.OK)
            //{
            //    filePath = openFileDialog.FileName;
            //}

            //if (string.IsNullOrEmpty(filePath))
            //    return Result.Cancelled;

            ////Чтение файла в список элементов типа RoomData
            //var lines = File.ReadAllLines(filePath).ToList();

            //List<RoomData> roomDataList = new List<RoomData>();
            //foreach (var line in lines)
            //{
            //    List<string> values = line.Split(';').ToList();
            //    roomDataList.Add(new RoomData
            //    {
            //        Name = values[0],
            //        Number = values[1],
            //    });
            //}

            ////Записываем значение "имени комнаты" из файла => в проект (по номеру комнаты)
            //using (var ts = new Transaction(doc, "Set parameters"))
            //{
            //    ts.Start();
            //    foreach (RoomData roomData in roomDataList)
            //    {
            //        Room room = rooms.FirstOrDefault(r => r.Number.Equals(roomData.Number));
            //        if (room == null)
            //            continue;
            //        room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(roomData.Name);
            //    }
            //    ts.Commit();
            //}
            #endregion

            #region Запись данных в эксель. Создание SheetExts.cs
            ////Указание пути, названия файла
            //string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "rooms.xlsx");

            ////Создание стрим(запись данных в файл)
            //using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            //{
            //    //Создание файла ексель
            //    IWorkbook workbook = new XSSFWorkbook();

            //    //Создание листа
            //    ISheet sheet = workbook.CreateSheet("Лист1");

            //    int rowIndex = 0;
            //    foreach (var room in rooms)
            //    {
            //        sheet.SetCellValue(rowIndex, columnIndex: 0, room.Name);
            //        sheet.SetCellValue(rowIndex, columnIndex: 1, room.Number);
            //        sheet.SetCellValue(rowIndex, columnIndex: 2, room.Area);
            //        rowIndex++;
            //    }
            //    //Запись данных в файл
            //    workbook.Write(stream);

            //    //Закрытие файла
            //    workbook.Close();
            //}

            ////Открытие файла в эксель
            //System.Diagnostics.Process.Start(excelPath);
            #endregion

            #region Чтение из файла эксель. Запись в проект
            //Диалог открытия файла
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "Excel files (*.xlsx)|*.xlsx"
            };

            //Указание пути
            string filePath = string.Empty;
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                filePath = openFileDialog.FileName;
            }
            if (string.IsNullOrEmpty(filePath))
                return Result.Cancelled;

            using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(filePath);
                ISheet sheet = workbook.GetSheetAt(index: 0);

                int rowIndex = 0;
                while (sheet.GetRow(rowIndex) != null)
                {
                    if (sheet.GetRow(rowIndex).GetCell(0) == null ||
                            sheet.GetRow(rowIndex).GetCell(1) == null)
                    {
                        rowIndex++;
                        continue;
                    }
                    
                    string name = sheet.GetRow(rowIndex).GetCell(0).StringCellValue;
                    string number = sheet.GetRow(rowIndex).GetCell(1).StringCellValue;

                    var room = rooms.FirstOrDefault(r => r.Number.Equals(number));

                    if (room != null)
                    {
                        rowIndex++;
                        continue;
                    }

                    using (var ts = new Transaction(doc, "Set parameter"))
                    {
                        ts.Start();
                        room.get_Parameter(BuiltInParameter.ROOM_NAME).Set(name);
                        ts.Commit();
                    }
                    rowIndex++;
                }
            }
            #endregion

            return Result.Succeeded;
        }
    }
}
