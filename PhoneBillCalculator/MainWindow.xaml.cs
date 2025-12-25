using System;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using Word = Microsoft.Office.Interop.Word;

namespace PhoneBillCalculator
{
    public partial class MainWindow : Window
    {
        // Храним данные последнего расчета для генерации квитанции
        private double lastCalculationAmount = 0;
        private int lastExtraMinutes = 0;
        private int lastMinutes = 0;
        private bool isTariff1 = true;

        public MainWindow()
        {
            InitializeComponent();
        }

        // Универсальный метод расчета для любого тарифа
        // Возвращает стоимость и количество минут сверх лимита
        public static (double cost, int extraMinutes) CalculateTariff(int minutes, int includedMinutes, double normalRate, double extraRate)
        {
            if (minutes < 0)
            {
                return (0, 0);
            }

            int extraMinutes = 0;
            double cost;

            // Расчет стоимости: если минут меньше или равно включенным, считаем по обычному тарифу
            if (minutes <= includedMinutes)
            {
                cost = minutes * normalRate;
            }
            else
            {
                // Иначе добавляем оплату сверхлимитных минут по повышенному тарифу
                extraMinutes = minutes - includedMinutes;
                cost = (includedMinutes * normalRate) + (extraMinutes * extraRate);
            }

            if (cost < 0)
            {
                cost = 0;
            }

            return (cost, extraMinutes);
        }

        private void btnCalculate_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateInput())
                return;

            int minutes = int.Parse(txtMinutes.Text);
            lastMinutes = minutes;

            if (rbTariff1.IsChecked == true)
            {
                isTariff1 = true;
                // Тариф 1: 200 минут по 0.7, сверх - 1.6
                var result = CalculateTariff(minutes, 200, 0.7, 1.6);
                lastCalculationAmount = result.cost;
                lastExtraMinutes = result.extraMinutes;
            }
            else
            {
                isTariff1 = false;
                // Тариф 2: 100 минут по 0.3, сверх - 1.6
                var result = CalculateTariff(minutes, 100, 0.3, 1.6);
                lastCalculationAmount = result.cost;
                lastExtraMinutes = result.extraMinutes;
            }

            txtResult.Text = $"Сумма к оплате: {lastCalculationAmount:F2} руб.";
            txtExtraMinutes.Text = $"Минут сверх нормы: {lastExtraMinutes}";

            // Активируем кнопку генерации только после успешного расчета
            btnGenerateReceipt.IsEnabled = true;
        }

        private void btnGenerateReceipt_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateReceiptData())
                return;

            Word.Application wordApp = null;
            Word.Document wordDoc = null;

            try
            {
                // Ищем шаблон в папке с программой
                string templatePath = Path.Combine(Directory.GetCurrentDirectory(), "ЧекШаблон.docx");

                if (!File.Exists(templatePath))
                {
                    MessageBox.Show($"Файл шаблона не найден:\n{templatePath}\n\n" +
                                  "Поместите файл 'Квитанция_шаблон.docx' в папку с программой.",
                                  "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                wordApp = new Word.Application();
                wordApp.Visible = false;
                wordDoc = wordApp.Documents.Open(templatePath);

                // Заменяем метки в шаблоне на реальные данные
                Repwo("{ФИО плательщика}", txtCustomerName.Text, wordDoc);
                Repwo("{Тариф}", isTariff1 ? "Тариф 1" : "Тариф 2", wordDoc);
                Repwo("{Адрес плательщика}", txtCustomerAddress.Text, wordDoc);

                string amountFormatted = lastCalculationAmount.ToString("F2").Replace(".", ",");
                Repwo("{Сумма платежа}", amountFormatted, wordDoc);
                Repwo("{Дата платежа}", DateTime.Now.ToString("dd.MM.yyyy"), wordDoc);

                // Генерируем уникальное имя файла с номером квитанции и датой
                string receiptNumber = new Random().Next(1, 1000000000).ToString();
                string fileName = $"Чек_{receiptNumber}_{DateTime.Now:dd.MM.yyyy}.docx";
                string appPath = AppDomain.CurrentDomain.BaseDirectory;
                string savePath = Path.Combine(appPath, fileName);

                wordDoc.SaveAs2(savePath);

                MessageBox.Show($"Квитанция успешно сформирована:\n{savePath}",
                    "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при формировании квитанции:\n{ex.Message}",
                    "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                // Освобождаем COM-объекты Word для избежания утечек памяти
                if (wordDoc != null)
                {
                    wordDoc.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                }
            }
        }

        // Метод для замены текста в документе Word (Find and Replace)
        private void Repwo(string textToFind, string replacementText, Word.Document wordDoc)
        {
            Word.Range range = wordDoc.Content;
            range.Find.ClearFormatting();
            range.Find.Text = textToFind;
            range.Find.Replacement.ClearFormatting();
            range.Find.Replacement.Text = replacementText;

            object replaceAll = Word.WdReplace.wdReplaceAll;
            range.Find.Execute(
                FindText: textToFind,
                ReplaceWith: replacementText,
                Replace: replaceAll
            );
        }

        private bool ValidateInput()
        {
            if (string.IsNullOrWhiteSpace(txtMinutes.Text))
            {
                MessageBox.Show("Введите количество минут", "Ошибка ввода",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (!int.TryParse(txtMinutes.Text, out int minutes))
            {
                MessageBox.Show("Введите целое число", "Ошибка ввода",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            // Предотвращаем ввод нереальных значений
            if (minutes > 1000000)
            {
                MessageBox.Show("Слишком большое количество минут (макс. 1,000,000)", "Ошибка ввода",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }

        private bool ValidateReceiptData()
        {
            if (string.IsNullOrWhiteSpace(txtCustomerName.Text))
            {
                MessageBox.Show("Введите Ф.И.О. плательщика", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            if (string.IsNullOrWhiteSpace(txtCustomerAddress.Text))
            {
                MessageBox.Show("Введите адрес плательщика", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            return true;
        }
    }
}