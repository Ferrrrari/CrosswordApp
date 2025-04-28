using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Data;

namespace CrosswordApp
{
    public partial class MainWindow : Window
    {

       
        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                DataTable table = new DataTable("Статистика");
                table.Columns.Add("Игрок", typeof(string));
                table.Columns.Add("Успешные попытки", typeof(int));
                table.Columns.Add("Ошибки", typeof(int));

                table.Rows.Add("Игрок 1", 10, 2);
                table.Rows.Add("Игрок 2", 8, 5);

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Статистика");

                    worksheet.Cell(1, 1).Value = "Игрок";
                    worksheet.Cell(1, 2).Value = "Успешные попытки";
                    worksheet.Cell(1, 3).Value = "Ошибки";

                    for (int i = 0; i < table.Rows.Count; i++)
                    {
                        worksheet.Cell(i + 2, 1).Value = table.Rows[i][0].ToString(); 
                        worksheet.Cell(i + 2, 2).Value = (int)table.Rows[i][1];      
                        worksheet.Cell(i + 2, 3).Value = (int)table.Rows[i][2];      
                    }

                    var dataRange = worksheet.RangeUsed();
                    var excelTable = dataRange.CreateTable();

                    excelTable.ShowAutoFilter = false;
                    excelTable.Theme = XLTableTheme.TableStyleMedium2;

                    workbook.SaveAs("Статистика.xlsx");
                    MessageBox.Show("Данные успешно экспортированы!");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        private bool _isDragging = false;
        private Point _startPosition;

        private string _questionsPath = "Data/questions.txt";
        private string _hintsPath = "Data/hints.txt";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void CheckAnswers_Click(object sender, RoutedEventArgs e)
        {
            foreach (TextBox textBox in FindVisualChildren<TextBox>(this))
            {
                if (textBox.Tag != null && textBox.Tag.ToString() == textBox.Text.ToUpper())
                {
                    textBox.Background = Brushes.LightGreen;
                }
                else
                {
                    textBox.Background = Brushes.LightCoral;
                }
            }
        }

        private void ClearFields_Click(object sender, RoutedEventArgs e)
        {
            foreach (TextBox textBox in FindVisualChildren<TextBox>(this))
            {
                textBox.Text = "";
                textBox.Background = Brushes.White;
            }
        }

        private void ShowAnswers_Click(object sender, RoutedEventArgs e)
        {
            foreach (TextBox textBox in FindVisualChildren<TextBox>(this))
            {
                if (textBox.Tag != null)
                {
                    textBox.Text = textBox.Tag.ToString();
                    textBox.Background = Brushes.LightGreen;
                }
            }
        }

        private void ShowHint(string text)
        {
            HintText.Text = text;
            var animation = new DoubleAnimation
            {
                From = 0,
                To = 1,
                Duration = TimeSpan.FromSeconds(0.3)
            };
            HintText.BeginAnimation(OpacityProperty, animation);
        }

        private void HideHint()
        {
            var animation = new DoubleAnimation
            {
                From = 1,
                To = 0,
                Duration = TimeSpan.FromSeconds(0.3)
            };
            HintText.BeginAnimation(OpacityProperty, animation);
        }

        private void TextBlock_MouseEnter(object sender, MouseEventArgs e)
        {
            var textBlock = sender as TextBlock;
            ShowHint(textBlock.Text);
        }

        private void TextBlock_MouseLeave(object sender, MouseEventArgs e)
        {
            HideHint();
        }

        private void TaskImage_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            _isDragging = true;
            _startPosition = e.GetPosition(ImageCanvas);
            TaskImage.CaptureMouse();
        }

        private void TaskImage_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            _isDragging = false;
            TaskImage.ReleaseMouseCapture();
        }

        private void TaskImage_MouseMove(object sender, MouseEventArgs e)
        {
            if (_isDragging)
            {
                Point currentPosition = e.GetPosition(ImageCanvas);
                double deltaX = currentPosition.X - _startPosition.X;
                double deltaY = currentPosition.Y - _startPosition.Y;
                Canvas.SetLeft(TaskImage, Canvas.GetLeft(TaskImage) + deltaX);
                Canvas.SetTop(TaskImage, Canvas.GetTop(TaskImage) + deltaY);
                _startPosition = currentPosition;
            }
        }

        private void BtnZoomIn_Click(object sender, RoutedEventArgs e)
        {
            ImageScale.ScaleX *= 1.2;
            ImageScale.ScaleY *= 1.2;
        }

        private void BtnZoomOut_Click(object sender, RoutedEventArgs e)
        {
            ImageScale.ScaleX /= 1.2;
            ImageScale.ScaleY /= 1.2;
        }

        private void BtnCloseImage_Click(object sender, RoutedEventArgs e)
        {
            TaskImage.Visibility = Visibility.Collapsed;
        }

        private void LoadQuestions_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var questions = File.ReadAllLines(_questionsPath);
                UpdateQuestions(questions);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void LoadHints_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var hints = File.ReadAllLines(_hintsPath);
                UpdateHints(hints);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        private void UpdateQuestions(string[] questions)
        {
            var stackPanel = (StackPanel)FindName("QuestionsPanel");
            stackPanel.Children.Clear();

            foreach (var question in questions)
            {
                var textBlock = new TextBlock
                {
                    Text = question,
                    Margin = new Thickness(0, 0, 0, 5),
                    TextWrapping = TextWrapping.Wrap
                };
                stackPanel.Children.Add(textBlock);
            }
        }

        private void UpdateHints(string[] hints)
        {
            var stackPanel = (StackPanel)FindName("HintsPanel");
            stackPanel.Children.Clear();

            foreach (var hint in hints)
            {
                var textBlock = new TextBlock
                {
                    Text = hint,
                    Margin = new Thickness(0, 0, 0, 5),
                    TextWrapping = TextWrapping.Wrap
                };
                stackPanel.Children.Add(textBlock);
            }
        }

        private void ThemeSelector_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var themeDict = new ResourceDictionary();
            switch (ThemeSelector.SelectedIndex)
            {
                case 0:
                    themeDict.Source = new Uri("LightTheme.xaml", UriKind.Relative);
                    break;
                case 1:
                    themeDict.Source = new Uri("DarkTheme.xaml", UriKind.Relative);
                    break;
                case 2:
                    themeDict.Source = new Uri("ColorfulTheme.xaml", UriKind.Relative);
                    break;
            }

            Application.Current.Resources.MergedDictionaries.Clear();
            Application.Current.Resources.MergedDictionaries.Add(themeDict);
        }

        private static IEnumerable<T> FindVisualChildren<T>(DependencyObject depObj) where T : DependencyObject
        {
            if (depObj != null)
            {
                for (int i = 0; i < VisualTreeHelper.GetChildrenCount(depObj); i++)
                {
                    DependencyObject child = VisualTreeHelper.GetChild(depObj, i);
                    if (child != null && child is T)
                    {
                        yield return (T)child;
                    }
                    foreach (T childOfChild in FindVisualChildren<T>(child))
                    {
                        yield return childOfChild;
                    }
                }
            }
        }
    }
}