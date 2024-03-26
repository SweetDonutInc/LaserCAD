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
using System.Threading.Tasks;
using System.Windows.Media;
using System.Windows.Controls;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace laserPj
{
    public partial class MainWindow : Window
    {
        private readonly string fontPath = "Files/ISOCPEUR.ttf"; //Шрифт
        private readonly double fontSize = 8d; //Размер шрифта
        

        List<ExcelCubeData> excelList = new List<ExcelCubeData>(); //В этот лист записываются строки из таблицы Ecxel, сгенерированной в Кубе
        public bool isMountAir { get; set; }

        public MainWindow()
        {
            InitializeComponent();
            version.Text = "Версия 1.4";
        }

        private void OpenCubeExcel_Click(object sender, RoutedEventArgs e)
        {
            try
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
            catch
            {
                MessageBox.Show("Данный Excel файл уже открыт. Закройте его и повторите попытку снова!");
            }
        }

        private void addExcelCubeList(string path) //Метод для считывания Excel-файла и записи его в лист excelList
        {
            try
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
                            holeHeight = sheet.GetRow(i).GetCell(9).NumericCellValue,
                            isD = (int)sheet.GetRow(i).GetCell(10).NumericCellValue
                        });
                    }
                }
                progressBar.Minimum = 0;
                progressBar.Maximum = excelList.Count;

                pbText.Visibility = Visibility.Visible;
                progressBar.Visibility = Visibility.Visible;

                if (isMountAir)
                {
                    MountAir_createFilesForCad(excelList);
                }
                else
                {
                    AirWay_createFilesForCad(excelList);
                }
            }
            catch
            {
                MessageBox.Show("Не удалось считать файл Excel. Возможно он повреждён или заполнен неправильно!");
            }
        }

        private async void MountAir_createFilesForCad(List<ExcelCubeData> excelList) // Метод для создания файлов dxf MA
        {
            List<ExcelCubeData> tempExcelList = new List<ExcelCubeData>(excelList); // Копия основного листа со всеми строками (для удобства работы)
            try
            {
                for (int i = 0; i < excelList.Count; i++)
                {
                    DxfModel model = new DxfModel(); // Создаём пустой файл
                    int border = tempExcelList[i].name.Contains("35мм") ? 35 : 23; //Если в названии встречается 35мм, то борты 35мм, иначе - 23мм

                    int newWidth = tempExcelList[i].width - (2 * border); // Ширина листа без бортов
                    int newHeight = tempExcelList[i].height - (2 * border); // Высота листа без бортов

                    // Рисуем линии основной фигуры
                    model = PerimetrDrawing(model, border, newWidth, newHeight);
                    // Линии нарисованы

                    //Добавляем вырез
                    double tw = tempExcelList[i].width;
                    double hw = tempExcelList[i].holeWidth;
                    double n = (tw - hw) / 2;

                    double th = tempExcelList[i].height;
                    double hh = tempExcelList[i].holeHeight;
                    double m = (th - hh) / 2;

                    if (hw > 0)
                    {
                        model = AddHole(model, n, m, hh, hw);
                    }
                    //Вырез добавлен

                    //Добавляем круг если есть
                    if (tempExcelList[i].isD == 1) model = AddCircle(model, tw, th);
                    //Круг добавлен

                    string t = "";
                    if (tempExcelList[i].type.Contains("Глухая")) t = "Г";
                    if (tempExcelList[i].type.Contains("Сервис")) t = "С";
                    if (tempExcelList[i].type.Contains("Створка")) t = "О";

                    //Добавляем гравировку
                    //model = AddGrav(model, tempExcelList[i].mark, t, tempExcelList[i].width, tempExcelList[i].height, tempExcelList[i].article, newWidth, border, newHeight);

                    //Собираем имя файла

                    string filename = Filename(tempExcelList[i].name, hw, hh, tempExcelList[i].lineNum, tempExcelList[i].mark, t,
                        tempExcelList[i].width, tempExcelList[i].height, tempExcelList[i].article, tempExcelList[i].count);

                    ////Записываем файл
                    bool checkPanel = tempExcelList[i].name.Contains("П54");
                    SaveAllFiles(checkPanel, filename, model, tempExcelList[i].article.ToLower().StartsWith("наруж"));

                    progressBar.Value = i + 1;
                    pbText.Text = $"{progressBar.Value}/{excelList.Count}";
                    await Task.Delay(1);
                }


                excel_path.Text = "Файлы по листу МА успешно созданы";
                excel_path.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Green);
            }
            catch
            {
                MessageBox.Show("Не удалось создать DXF файлы! Проверьте правильность заполнение Excel файлов");
            }
        }

        private async void AirWay_createFilesForCad(List<ExcelCubeData> excelList) // Метод для создания файлов dxf AW
        {
            List<ExcelCubeData> tempExcelList = new List<ExcelCubeData>(excelList); // Копия основного листа со всеми строками (для удобства работы)

            try
            {
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
                    model = PerimetrDrawing(model, border, newWidth, newHeight);
                    // Линии нарисованы

                    //Добавляем вырез
                    double tw = tempExcelList[i].width;
                    double hw = tempExcelList[i].holeWidth;
                    double n = (tw - hw) / 2;

                    double th = tempExcelList[i].height;
                    double hh = tempExcelList[i].holeHeight;
                    double m = (th - hh) / 2;

                    if (hw > 0)
                    {
                        model = AddHole(model, n, m, hh, hw);
                    };

                    //Вырез добавлен

                    //Добавляем круг если есть
                    if (tempExcelList[i].isD == 1) model = AddCircle(model, tw, th);
                    //Круг добавлен

                    string t = "";
                    if (tempExcelList[i].type.Contains("Глухая")) t = "Г";
                    if (tempExcelList[i].type.Contains("Сервис")) t = "С";
                    if (tempExcelList[i].type.Contains("Открывающаяся")) t = "О";

                    //Добавляем гравировку
                    model = AddGrav(model, tempExcelList[i].mark, t, tempExcelList[i].width, tempExcelList[i].height, tempExcelList[i].article, newWidth, border, newHeight);

                    //Собираем имя файла
                    string filename = Filename(tempExcelList[i].name, hw, hh, tempExcelList[i].lineNum, tempExcelList[i].mark, t,
                        tempExcelList[i].width, tempExcelList[i].height, tempExcelList[i].article, tempExcelList[i].count);

                    //Записываем файл
                    bool checkPanel = tempExcelList[i].name.Contains("П54");
                    SaveAllFiles(checkPanel, filename, model, tempExcelList[i].article.ToLower().StartsWith("наруж"));

                    progressBar.Value = i + 1;
                    pbText.Text = $"{progressBar.Value}/{excelList.Count}";
                    await Task.Delay(1);
                }

                excel_path.Text = "Файлы по листу AW успешно созданы";
                excel_path.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Colors.Green);
            }
            catch
            {
                MessageBox.Show("Не удалось создать DXF файлы! Проверьте правильность заполнение Excel файлов");
            }
        }

        private void SaveDxf_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Title = "Создать папку проекта";
            if (saveFileDialog1.ShowDialog() == true)
            {
                DirectoryInfo dirInfo = new DirectoryInfo(saveFileDialog1.FileName);
                dirInfo.Create();

                DirectoryInfo dw1 = new DirectoryInfo(saveFileDialog1.FileName + "/П54 внешние");
                dw1.Create();
                DirectoryInfo dw2 = new DirectoryInfo(saveFileDialog1.FileName + "/П54 внутренние");
                dw2.Create();
                DirectoryInfo dw3 = new DirectoryInfo(saveFileDialog1.FileName + "/П42 внешние");
                dw3.Create();
                DirectoryInfo dw4 = new DirectoryInfo(saveFileDialog1.FileName + "/П42 внутренние");
                dw4.Create();
                directory.Text = dirInfo.FullName;
                AW.Visibility = Visibility.Visible;
                MA.Visibility = Visibility.Visible;
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

        //Рисуем периметр листа
        private DxfModel PerimetrDrawing(DxfModel model, int border, int newWidth, int newHeight)
        {
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
            return model;
        }

        //Добавляем вырез
        private DxfModel AddHole(DxfModel model, double n, double m, double hh, double hw)
        {
            model.Entities.Add(new DxfLine(new Point2D(n, m), new Point2D(n, m + hh)));
            model.Entities.Add(new DxfLine(new Point2D(n, m + hh), new Point2D(n + hw, m + hh)));
            model.Entities.Add(new DxfLine(new Point2D(n + hw, m + hh), new Point2D(n + hw, m)));
            model.Entities.Add(new DxfLine(new Point2D(n + hw, m), new Point2D(n, m)));
            return model;
        }

        //Добавляем круговой вырез
        private DxfModel AddCircle(DxfModel model, double tw, double th)
        {
            model.Entities.Add(new DxfCircle(new Point2D(tw / 2, th / 2), 100));
            return model;
        }

        //Добавляем гравировку
        private DxfModel AddGrav(DxfModel model, string mark, string t, int width, int height, string article, int newWidth, int border, int newHeight)
        {
            //Гравировка собирается так: Номер заказа Маркировка Первый символ типа Ширина х Высота_Первый символ артикула 
            string grav = orderNum_int.Text + " "
                    + mark + " " + t + " "
                    + width.ToString() + "x" + height.ToString()
                    + "_" + article[0];

            //Если панель узкая по ширине, то поворачиваем гравировку
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
            return model;
        }

        //Создадим имя файла
        private string Filename(string name, double hw, double hh, int lineNum, string mark, string t, int width, int height, string article, int count)
        {
            string tempRal = name.Contains("RAL") ? "RAL" : "Нерж";

            string crop = hw > 0 ? $" (Вырез {hw} x {hh})" : "";

            string filename = lineNum.ToString() + "_"
                + mark + " "
                + t + " "
                + width.ToString() + "x" + height.ToString() + "_"
                + article[0] + "_" + tempRal + "_" + count.ToString() + "шт"
                + crop;
            return filename;
        }

        //Сохраняем рисунки в файлы
        private void SaveAllFiles(bool BigPanel, string filename, DxfModel model, bool outside)
        {
            if (BigPanel && outside)
            {
                DxfWriter.Write($@"{directory.Text}\П54 внешние\{filename}.dxf", model);
            }
            else if (BigPanel && !outside)
            {
                DxfWriter.Write($@"{directory.Text}\П54 внутренние\{filename}.dxf", model);
            }
            else if (!BigPanel && outside)
            {
                DxfWriter.Write($@"{directory.Text}\П42 внешние\{filename}.dxf", model);
            }
            else
            {
                DxfWriter.Write($@"{directory.Text}\П42 внутренние\{filename}.dxf", model);
            }
        }

        private void AllClear_Click(object sender, RoutedEventArgs e)
        {
            orderNum_int.Clear();
            excelList.Clear();

            directory.Text = "Пусть к каталогу с DXF файлами внутренних листов";

            MA.Visibility = Visibility.Hidden;
            MA.Opacity = 1f;
            AW.Visibility = Visibility.Hidden;
            AW.Opacity = 1f;

            progressBar.Value = 0;
            pbText.Text = "0/0";
            pbText.Visibility = Visibility.Hidden;
            progressBar.Visibility = Visibility.Hidden;

            excbtn.Visibility = Visibility.Hidden;
            excel_path.Text = "";
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
        public int isD;
    }    
}
