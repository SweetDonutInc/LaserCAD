using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using Microsoft.Win32;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Util;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using WW.Cad.Base;
using WW.Cad.IO;
using WW.Cad.Model;
using WW.Cad.Model.Entities;
using WW.Cad.Model.Tables;
using WW.Math;
using System.Threading;
using System.Windows.Media;
using System.ComponentModel;

namespace laserPj
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private readonly string fontPath = "Files/ISOCPEUR.ttf";
        private readonly double fontSize = 8d;

        List<ExcelCubeData> excelList = new List<ExcelCubeData>(); //В этот лист записываются строки из таблицы Ecxel, сгенерированной в Кубе
        public bool isMountAir { get; set; }

        public MainWindow()
        {
            InitializeComponent();
        }

        private void OpenCubeExcel_Click(object sender, RoutedEventArgs e)
        {
            string excelCube_path = "";
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel files (*.xls)|*.xls";

            if (openFileDialog.ShowDialog() == true)
            {
                excelCube_path = openFileDialog.FileName;
                addExcelCubeList(excelCube_path);
            }
        }

        private void addExcelCubeList(string path) //Метод для считывания Excel-файла и записи его в лист excelList (строка 18)
        {
            IWorkbook workbook;
            using (FileStream fileStream = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                workbook = new HSSFWorkbook(fileStream); // Считываем загруженный файл
            }

            ISheet sheet = workbook.GetSheetAt(0); // Открываем первый лист

            for (int i = 4; i < sheet.LastRowNum; i++) // Первые 4 строки не содержат данных, поэтому начинаем с 4 индекса (5 строка)
            {
                // Добавляем в лист всё, что содержит борты
                if (sheet.GetRow(i).GetCell(3).StringCellValue.Contains("борты"))
                {
                    excelList.Add(new ExcelCubeData
                    {
                        lineNum = i + 1,
                        mark = sheet.GetRow(i).GetCell(0).StringCellValue,
                        type = sheet.GetRow(i).GetCell(1).StringCellValue,
                        article = sheet.GetRow(i).GetCell(2).StringCellValue,
                        name = sheet.GetRow(i).GetCell(3).StringCellValue,
                        width = (int)sheet.GetRow(i).GetCell(4).NumericCellValue,
                        height = (int)sheet.GetRow(i).GetCell(5).NumericCellValue,
                        count = (int)sheet.GetRow(i).GetCell(6).NumericCellValue,
                        mass = (int)sheet.GetRow(i).GetCell(7).NumericCellValue,
                        holeWidth = sheet.GetRow(i).GetCell(8).NumericCellValue,
                        holeHeight = sheet.GetRow(i).GetCell(9).NumericCellValue
                    });
                }
            }
            if (isMountAir)
            {
                MountAir_createFilesForCad(excelList);
            }
            else
            {
                AirWay_createFilesForCad(excelList);
            }
        }

        private void MountAir_createFilesForCad(List<ExcelCubeData> excelList) // Метод для создания файлов dxf MA
        {
            List<ExcelCubeData> tempExcelList = new List<ExcelCubeData>(excelList); // Копия основного листа со всеми строками (для удобства работы)

            for (int i = 0; i < tempExcelList.Count; i++)
            {
                DxfModel model = new DxfModel(); // Создаём пустой файл
                int border = tempExcelList[i].name.Contains("35мм") ? 35 : 23; //Если в названии встречается 35мм, то борты 35мм, иначе - 23мм

                int newWidth = tempExcelList[i].width - (2 * border); // Ширина листа без бортов
                int newHeight = tempExcelList[i].height - (2 * border); // Высота листа без бортов

                // Рисуем линии основной фигуры
                model.Entities.Add(new DxfLine(new Point2D(0, border), new Point2D(0, newHeight + border)));
                model.Entities.Add(new DxfLine(new Point2D(0, newHeight + border), new Point2D(border, newHeight + border)));
                model.Entities.Add(new DxfLine(new Point2D(border, newHeight + border), new Point2D(border, newHeight + (2 * border))));
                model.Entities.Add(new DxfLine(new Point2D(border, newHeight + (2 * border)), new Point2D(newWidth + border, newHeight + (2 * border))));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + border, newHeight + (2 * border)), new Point2D(newWidth + border, newHeight + border)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + border, newHeight + border), new Point2D(newWidth + (2 * border), newHeight + border)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + (2 * border), newHeight + border), new Point2D(newWidth + (2 * border), border)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + (2 * border), border), new Point2D(newWidth + border, border)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + border, border), new Point2D(newWidth + border, 0)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + border, 0), new Point2D(border, 0)));
                model.Entities.Add(new DxfLine(new Point2D(border, 0), new Point2D(border, border)));
                model.Entities.Add(new DxfLine(new Point2D(border, border), new Point2D(0, border)));
                // Линии нарисованы

                //Добавляем вырез
                double tw = tempExcelList[i].width;
                double hw = tempExcelList[i].holeWidth;
                double n = (tw - hw) / 2;

                double th = tempExcelList[i].height;
                double hh = tempExcelList[i].holeHeight;
                double m = (th - hh) / 2;

                model.Entities.Add(new DxfLine(new Point2D(n,m), new Point2D(n, m + hh)));
                model.Entities.Add(new DxfLine(new Point2D(n, m + hh), new Point2D(n + hw, m + hh)));
                model.Entities.Add(new DxfLine(new Point2D(n + hw, m + hh), new Point2D(n + hw, m)));
                model.Entities.Add(new DxfLine(new Point2D(n + hw, m), new Point2D(n,m)));

                //Вырез добавлен

                string t = "";
                if (tempExcelList[i].type.Contains("Глухая")) t = "Г";
                if (tempExcelList[i].type.Contains("Сервис")) t = "С";
                if (tempExcelList[i].type.Contains("Створка")) t = "О";

                //Добавляем гравировку
                string grav = orderNum_int.Text + " "
                    + tempExcelList[i].mark + " "
                    + t + " "
                    + tempExcelList[i].width.ToString() + "x" + tempExcelList[i].height.ToString()
                    + "_" + tempExcelList[i].article[0];

                DxfText gravText;
                if (newWidth < 160)
                {
                    gravText = new DxfText(grav, new Point3D(border / 2, newHeight + border - 5, 0), fontSize)
                    {
                        Rotation = -(Math.PI / 2)
                    };
                }
                else
                {
                    gravText = new DxfText(grav, new Point3D(border + 5, border / 2, 0), fontSize);
                }

                gravText.Color = EntityColors.Green;
                DxfTextStyle textStyle = new DxfTextStyle("MYSTYLE", "Files/ISOCPEUR.ttf");
                model.TextStyles.Add(textStyle);
                gravText.Style = textStyle;

                DxfLayer layer = new DxfLayer("textLayer");
                model.Layers.Add(layer);
                gravText.Layer = layer;

                model.Entities.Add(gravText);

                //Собираем имя файла
                string tempRal = tempExcelList[i].name.Contains("RAL") ? "RAL" : "Нерж";

                string crop = hw > 0 ? $"(Вырез {hw} x {hh})" : "";

                string filename = tempExcelList[i].lineNum.ToString() + "_"
                    + tempExcelList[i].mark + " "
                    + t + " "
                    + tempExcelList[i].width.ToString() + "x" + tempExcelList[i].height.ToString() + "_"
                    + tempExcelList[i].article[0] + "_" + tempRal + "_" + tempExcelList[i].count.ToString() + "шт"
                    + crop;
                //Записываем файл
                if (tempExcelList[i].name.Contains("RAL"))
                {
                    DxfWriter.Write($@"{directoryRal.Text}\{filename}.dxf", model);
                }
                else
                {
                    DxfWriter.Write($@"{directory.Text}\{filename}.dxf", model);
                }
            }

            excel_path.Text = "Файлы по листу МА успешно созданы";
            excel_path.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Green);
        }

        private void AirWay_createFilesForCad(List<ExcelCubeData> excelList) // Метод для создания файлов dxf AW
        {
            List<ExcelCubeData> tempExcelList = new List<ExcelCubeData>(excelList); // Копия основного листа со всеми строками (для удобства работы)

            for (int i = 0; i < tempExcelList.Count; i++)
            {
                DxfModel model = new DxfModel(); // Создаём пустой файл

                int border = 0;
                if (tempExcelList[i].name.Contains("23мм")) border = 23;
                if (tempExcelList[i].name.Contains("24мм")) border = 24;
                if (tempExcelList[i].name.Contains("43мм")) border = 43;
                if (tempExcelList[i].name.Contains("44мм")) border = 44;

                int newWidth = tempExcelList[i].width - (2 * border); // Ширина листа без бортов
                int newHeight = tempExcelList[i].height - (2 * border); // Высота листа без бортов

                // Рисуем линии по точкам
                model.Entities.Add(new DxfLine(new Point2D(0, border), new Point2D(0, newHeight + border)));
                model.Entities.Add(new DxfLine(new Point2D(0, newHeight + border), new Point2D(border, newHeight + border)));
                model.Entities.Add(new DxfLine(new Point2D(border, newHeight + border), new Point2D(border, newHeight + (2 * border))));
                model.Entities.Add(new DxfLine(new Point2D(border, newHeight + (2 * border)), new Point2D(newWidth + border, newHeight + (2 * border))));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + border, newHeight + (2 * border)), new Point2D(newWidth + border, newHeight + border)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + border, newHeight + border), new Point2D(newWidth + (2 * border), newHeight + border)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + (2 * border), newHeight + border), new Point2D(newWidth + (2 * border), border)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + (2 * border), border), new Point2D(newWidth + border, border)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + border, border), new Point2D(newWidth + border, 0)));
                model.Entities.Add(new DxfLine(new Point2D(newWidth + border, 0), new Point2D(border, 0)));
                model.Entities.Add(new DxfLine(new Point2D(border, 0), new Point2D(border, border)));
                model.Entities.Add(new DxfLine(new Point2D(border, border), new Point2D(0, border)));
                // Линии нарисованы

                //Добавляем вырез
                double tw = tempExcelList[i].width;
                double hw = tempExcelList[i].holeWidth;
                double n = (tw - hw) / 2;

                double th = tempExcelList[i].height;
                double hh = tempExcelList[i].holeHeight;
                double m = (th - hh) / 2;

                model.Entities.Add(new DxfLine(new Point2D(n, m), new Point2D(n, m + hh)));
                model.Entities.Add(new DxfLine(new Point2D(n, m + hh), new Point2D(n + hw, m + hh)));
                model.Entities.Add(new DxfLine(new Point2D(n + hw, m + hh), new Point2D(n + hw, m)));
                model.Entities.Add(new DxfLine(new Point2D(n + hw, m), new Point2D(n, m)));

                //Вырез добавлен

                string t = "";
                if (tempExcelList[i].type.Contains("Глухая")) t = "Г";
                if (tempExcelList[i].type.Contains("Сервис")) t = "С";
                if (tempExcelList[i].type.Contains("Открывающаяся")) t = "О";

                //Добавляем гравировку
                string grav = orderNum_int.Text + " "
                    + tempExcelList[i].mark + " "
                    + t + " "
                    + tempExcelList[i].width.ToString() + "x" + tempExcelList[i].height.ToString()
                    + "_" + tempExcelList[i].article[0];

                DxfText gravText;
                if (newWidth < 160)
                {
                    gravText = new DxfText(grav, new Point3D(border / 2, newHeight + border - 5, 0), fontSize)
                    {
                        Rotation = -(Math.PI / 2)
                    };
                }
                else
                {
                    gravText = new DxfText(grav, new Point3D(border + 5, border / 2, 0), fontSize);
                }

                gravText.Color = EntityColors.Green;
                DxfTextStyle textStyle = new DxfTextStyle("MYSTYLE", fontPath);
                model.TextStyles.Add(textStyle);
                gravText.Style = textStyle;

                DxfLayer layer = new DxfLayer("textLayer");
                model.Layers.Add(layer);
                gravText.Layer = layer;

                model.Entities.Add(gravText);

                //Собираем имя файла
                string tempRal = tempExcelList[i].name.Contains("RAL") ? "Ral 0.5" : "оц 0.7";

                string crop = hw > 0 ? $"(Вырез {hw} x {hh})" : "";

                string filename = tempExcelList[i].lineNum.ToString() + "_"
                    + tempExcelList[i].mark + " "
                    + t + " "
                    + tempExcelList[i].width.ToString() + "x" + tempExcelList[i].height.ToString() + "_"
                    + tempExcelList[i].name[0] + "_" + tempRal + "_" + tempExcelList[i].count.ToString() + "шт" + crop;
                //Записываем файл
                if (tempExcelList[i].name.Contains("RAL"))
                {
                    DxfWriter.Write($@"{directoryRal.Text}\{filename}.dxf", model);
                }
                else
                {
                    DxfWriter.Write($@"{directory.Text}\{filename}.dxf", model);
                }
            }

            excel_path.Text = "Файлы по листу AW успешно созданы";
            excel_path.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Green);
        }

        private void SaveDxf_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Создать папку проекта";
            if(saveFileDialog1.ShowDialog() == true)
            {
                DirectoryInfo dirInfo = new DirectoryInfo(saveFileDialog1.FileName);
                dirInfo.Create();
                directory.Text = dirInfo.FullName;
                directoryRal.Visibility = Visibility.Visible;
                dirRelBtn.Visibility = Visibility.Visible;
            }
        }

        private void SaveDxfRal_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Создать папку проекта";
            if (saveFileDialog1.ShowDialog() == true)
            {
                DirectoryInfo dirInfo = new DirectoryInfo(saveFileDialog1.FileName);
                dirInfo.Create();
                directoryRal.Text = dirInfo.FullName;
                MA.Visibility = Visibility.Visible;
                AW.Visibility = Visibility.Visible;
            }
        }

        private void MountAir_Click(object sender, RoutedEventArgs e)
        {
            isMountAir = true;
            MA.Opacity = 1f;
            AW.Opacity = 0.5f;
            excbtn.Visibility = Visibility.Visible;
            excel_path.Visibility = Visibility.Visible;
        }

        private void AirWay_Click(object sender, RoutedEventArgs e)
        {
            isMountAir = false;
            AW.Opacity = 1f;
            MA.Opacity = 0.5f;
            excbtn.Visibility = Visibility.Visible;
            excel_path.Visibility = Visibility.Visible;
        }
    }

    public class ExcelCubeData
    {
        //Класс для считывания Excel файла, сгенерированного программой Куб

        public int lineNum; //Номер строки в таблице (Вероятно пригодится потом)
        public string mark; //Маркировка
        public string type; //Тип
        public string article; //Артикул
        public string name; //Название
        public int width; //Ширина
        public int height; //Высота
        public int count; //Количество
        public int mass; //Масса
        public double holeWidth; //Ширина выреза
        public double holeHeight; //Высота выреза
    }
}
