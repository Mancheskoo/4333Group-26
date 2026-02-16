using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Group4333
{
    /// <summary>
    /// Логика взаимодействия для _4333_MonichArtem.xaml
    /// </summary>
    public partial class _4333_MonichArtem : Window
    {
        public _4333_MonichArtem()
        {
            InitializeComponent();
        }

        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            // 1. Выбор файла через диалоговое окно
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "Файлы Excel (*.xlsx)|*.xlsx",
                Title = "Выберите файл для импорта (3.xlsx)"
            };

            if (!(ofd.ShowDialog() == true))
                return;

            // 2. Запуск Excel и открытие книги
            Excel.Application app = new Excel.Application();
            Excel.Workbook workbook = app.Workbooks.Open(ofd.FileName);
            Excel.Worksheet worksheet = workbook.Sheets[1]; // Работаем с первым листом

            // Определяем диапазон заполненных данных
            Excel.Range lastCell = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int rowCount = lastCell.Row;

            try
            {
                using (MonichContext db = new MonichContext())
                {
                    // Цикл со 2-й строки (пропускаем заголовки таблицы)
                    for (int i = 2; i <= rowCount; i++)
                    {
                        // Считываем данные из ячеек (номера столбцов должны соответствовать вашему Excel)
                        // Предположим порядок: 1-ФИО, 2-Код, 3-ДатаРожд, 4-Индекс, 5-Город, 6-Улица, 7-Дом, 8-Кварт, 9-Email
                        string fio = worksheet.Cells[i, 1].Text;
                        string code = worksheet.Cells[i, 2].Text;
                        string dateStr = worksheet.Cells[i, 3].Text;
                        string index = worksheet.Cells[i, 4].Text;
                        string city = worksheet.Cells[i, 5].Text;
                        string street = worksheet.Cells[i, 6].Text;
                        string house = worksheet.Cells[i, 7].Text;
                        string apartment = worksheet.Cells[i, 8].Text;
                        string email = worksheet.Cells[i, 9].Text;

                        if (DateTime.TryParse(dateStr, out DateTime birthDate))
                        {
                            // --- ТРЕБОВАНИЕ ВАРИАНТА 7: Расчет возраста программно ---
                            int age = DateTime.Now.Year - birthDate.Year;
                            if (DateTime.Now.DayOfYear < birthDate.DayOfYear)
                                age--;

                            // Создаем объект сущности (согласно вашей структуре БД)
                            Clients newClient = new Clients
                            {
                                FullName = fio,
                                ClientCode = code,
                                BirthDate = birthDate,
                                IndexCode = index,
                                City = city,
                                Street = street,
                                House = house,
                                Apartment = apartment,
                                Email = email,
                                Age = age // Наш вычисленный атрибут
                            };

                            // Добавляем в контекст
                            db.Clients.Add(newClient);
                        }
                    }

                    // Сохраняем все изменения в базу данных одним махом
                    db.SaveChanges();
                    MessageBox.Show("Данные успешно импортированы в БД и возраст рассчитан!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при сохранении в БД: {ex.Message}");
            }
            finally
            {
                // Обязательно закрываем Excel, чтобы процесс не висел в памяти
                workbook.Close(false);
                app.Quit();

                // Освобождаем COM-объекты (рекомендуется в лекции)
                System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
        }

        private void BtnExport_Click(object sender, RoutedEventArgs e)
        {
            List<Clients> allClients;

            // 1. Извлекаем все данные из базы данных
            using (MonichContext db = new MonichContext())
            {
                allClients = db.Clients.ToList();
            }

            if (allClients.Count == 0)
            {
                MessageBox.Show("База данных пуста. Сначала выполните импорт.");
                return;
            }

            // 2. Группируем данные по критерию Варианта 7
            var category1 = allClients.Where(c => c.Age >= 20 && c.Age <= 29).ToList();
            var category2 = allClients.Where(c => c.Age >= 30 && c.Age <= 39).ToList();
            var category3 = allClients.Where(c => c.Age >= 40).ToList();

            // 3. Инициализация Excel
            Excel.Application app = new Excel.Application();
            app.SheetsInNewWorkbook = 3; // Сразу создаем книгу с 3 листами
            Excel.Workbook workbook = app.Workbooks.Add();

            // Подготавливаем данные для цикла: списки категорий и названия листов
            var categories = new List<List<Clients>> { category1, category2, category3 };
            string[] sheetNames = { "20-29 лет", "30-39 лет", "от 40 лет" };

            try
            {
                for (int i = 0; i < 3; i++)
                {
                    // Выбираем лист (индексация в Excel начинается с 1)
                    Excel.Worksheet worksheet = workbook.Worksheets[i + 1];
                    worksheet.Name = sheetNames[i];

                    // 4. Установка заголовков согласно формату Варианта 7
                    // Столбцы: Код, ФИО клиента, Возраст, E-mail
                    worksheet.Cells[1, 1] = "Код";
                    worksheet.Cells[1, 2] = "ФИО клиента";
                    worksheet.Cells[1, 3] = "Возраст";
                    worksheet.Cells[1, 4] = "E-mail";

                    // Оформление заголовков (жирный шрифт)
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[1, 4]];
                    headerRange.Font.Bold = true;

                    // 5. Заполнение данных из соответствующей категории
                    int currentRow = 2;
                    foreach (var client in categories[i])
                    {
                        worksheet.Cells[currentRow, 1] = client.ClientCode;
                        worksheet.Cells[currentRow, 2] = client.FullName;
                        worksheet.Cells[currentRow, 3] = client.Age;
                        worksheet.Cells[currentRow, 4] = client.Email;
                        currentRow++;
                    }

                    // 6. Форматирование таблицы (границы и автоподбор ширины)
                    if (currentRow > 2) // Если в категории есть хотя бы один человек
                    {
                        Excel.Range tableRange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[currentRow - 1, 4]];
                        tableRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous; // Сплошная линия границ
                        worksheet.Columns.AutoFit(); // Автоподбор ширины столбцов
                    }
                }

                // Делаем Excel видимым для пользователя
                app.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Произошла ошибка при экспорте: {ex.Message}");
            }
            finally
            {
                // Освобождаем ресурсы, чтобы процессы Excel не висели в диспетчере задач
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(app);
            }
        }
    }
}
