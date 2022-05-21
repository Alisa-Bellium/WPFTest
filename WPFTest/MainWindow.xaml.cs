using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;

namespace WPFTest
{

    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        static int countQ;
        StackPanel[] Panels;
        TextBlock resultBlock;
        Object[,] Question;
        Label TimerLabel;
        int secondsPassed;
        DispatcherTimer timer;

        // Индексы правильных ответов
        int[] answers;

        // Индексы пользовательских ответов
        int[] userAnswers;

        // Количество ответов на вопрос
        const int optionsPerQuestion = 5;

        // Смещение для работы со строками Excel
        const int excelRowOffset = 2;

        // Смещение для работы с колонками Excel
        const int excelColumnOffset = 1;

        // Время (в секундах) для прохождения теста
        const int secondsToAnwer = 100;

        private void Window_Initialized(object sender, EventArgs e)
        {
            InitialMethod();
        }

        private void InitialMethod()
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook excelBook = excelApp.Workbooks.Open(@"C:\Users\Xiaomi\OneDrive\Рабочий стол\myTest.xlsx", 0, true, optionsPerQuestion + 1, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            Microsoft.Office.Interop.Excel.Worksheet excelSheet = (Microsoft.Office.Interop.Excel.Worksheet)excelBook.Worksheets.get_Item(1); ;
            Microsoft.Office.Interop.Excel.Range excelRange = excelSheet.UsedRange;

            countQ = excelRange.Rows.Count - 1;

            // Делаем по числу вопросов + 1 для отображения результатов + 1 для таймера
            Panels = new StackPanel[countQ + 2];

            // Таймер
            Panels[0] = new StackPanel();
            TimerLabel = new Label();
            TimerLabel.Content = "";
            Panels[0].Children.Add(TimerLabel);
            MainPanel.Children.Add(Panels[0]);
            timer = new DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(1);
            timer.Tick += TimerTick;

            // Резервируем 0ую колонку под вопрос, и еще optionsPerQuestion на ответы
            Question = new object[countQ, optionsPerQuestion + 1];

            answers = new int[countQ];
            userAnswers = new int[countQ];

            for (int i = 0; i < countQ; i++)
            {
                Panels[i+1] = new StackPanel();
                Panels[i+1].Orientation = Orientation.Vertical;
                Panels[i+1].Background = new SolidColorBrush(Colors.CadetBlue);
            }
            for (int i = 0; i < countQ; i++)
            {
                Question[i, 0] = new Label();
                (Question[i, 0] as Label).Content = Convert.ToString((excelRange.Cells[i + excelRowOffset, excelColumnOffset] as Microsoft.Office.Interop.Excel.Range).Value2);
                (Question[i, 0] as Label).FontSize = 20;
                Panels[i+1].Children.Add(Question[i, 0] as Label);

                for (int j = 1; j <= optionsPerQuestion; j++)
                {
                    Question[i, j] = new RadioButton();
                    (Question[i, j] as RadioButton).Content = Convert.ToString((excelRange.Cells[i + excelRowOffset, j + excelColumnOffset] as Microsoft.Office.Interop.Excel.Range).Value2);
                    (Question[i, j] as RadioButton).FontSize = 20;
                    Panels[i + 1].Children.Add(Question[i, j] as RadioButton);
                }

                // В последней колонке лежат номера правильных ответов, забираем их по +1
                answers[i] = Convert.ToInt32((excelRange.Cells[i + excelRowOffset, optionsPerQuestion + excelColumnOffset + 1] as Microsoft.Office.Interop.Excel.Range).Value2);
                Panels[i + 1].Margin = new Thickness(3, 5, 0, 0);
                MainPanel.Children.Add(Panels[i+1]);
            }

            excelBook.Close(true, null, null);
            excelApp.Quit();

            Button btn = new Button();
            btn.Content = "Ok";
            btn.Width = 85;
            btn.Margin = new Thickness(10, 10, 10, 10);
            btn.HorizontalAlignment = HorizontalAlignment.Right;
            btn.Click += Btn_Click;
            MainPanel.Children.Add(btn);

            // Панель для показа результатов
            resultBlock = new TextBlock();
            resultBlock.Padding = new Thickness(5, 20, 5, 20);
            resultBlock.Background = new SolidColorBrush(Colors.LightGray);
            resultBlock.Visibility = Visibility.Collapsed;
            MainPanel.Children.Add(resultBlock);

            timer.Start();
        }

        private void TimerTick(object sender, EventArgs e)
        {
            secondsPassed++;
            int secondsRemaining = secondsToAnwer - secondsPassed;
            int minutesRemaining = secondsRemaining / 60;

            if (secondsRemaining < 0)
            {
                TimerLabel.Content = "Время на прохождение теста закончилось.";
                FinishTest();
            } else if (minutesRemaining < 1)
            {
                TimerLabel.Content = "На прохождение теста осталось секунд: " + secondsRemaining.ToString();
            } else
            {
                TimerLabel.Content = "На прохождение теста осталось минут: " + minutesRemaining.ToString();
            }
        }

        private void Btn_Click(object sender, RoutedEventArgs e)
        {
            FinishTest();
        }

        private void FinishTest()
        {
            timer.Stop();
            BeforeCheckedMethod();
            CheckedMethod();
        }

        private void BeforeCheckedMethod()
        {
            for (int i = 0; i < countQ; i++)
            {
                for (int j = 1; j <= optionsPerQuestion; j++)
                {
                    if ((Question[i, j] as RadioButton).IsChecked == true)
                    {
                        userAnswers[i] = j;
                    }
                }
            }


            for (int i = 0; i < countQ; i++)
            {
                for (int j = 1; j <= optionsPerQuestion; j++)
                {
                    (Question[i, j] as RadioButton).IsEnabled = false;
                }
            }
        }

        private void CheckedMethod()
        {
            int success = 0;
            int answered = 0;
            for (int i = 0; i < countQ; i++)
            {
                if (userAnswers[i] != 0)
                {
                    answered++;
                    if (userAnswers[i] == answers[i])
                    {
                        (Question[i, userAnswers[i]] as RadioButton).Foreground = new SolidColorBrush(Colors.Green);
                        success++;
                    }
                    else
                    {
                        (Question[i, userAnswers[i]] as RadioButton).Foreground = new SolidColorBrush(Colors.Red);
                    }
                }
                else
                {
                    (Question[i, answers[i]] as RadioButton).Foreground = new SolidColorBrush(Colors.Green);
                }
            }

            float percentSuccess = 0;

            if (answered != 0)
            {
                percentSuccess = (float) success / (float)answered * 100;
            }

            resultBlock.Text = "Вы завершили тест.\n" +
                "Количество вопросов: " + countQ.ToString() + "\n" +
                "Вы ответили на вопросов: " + answered.ToString() + "\n" +
                "Правильных ответов: " + success.ToString() + "\n" +
                "Неправильных ответов: " + (answered - success).ToString() + "\n" +
                "Процент правильных ответов: " + percentSuccess.ToString() + "%";
            resultBlock.Visibility = Visibility.Visible;

            TimerLabel.Content = "Тест завершен.";
        }
    }
}
