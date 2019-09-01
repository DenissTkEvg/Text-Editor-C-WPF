using Microsoft.Win32;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Path = System.IO.Path;

namespace TextEd
{

    public partial class MainWindow : Window
    {
        bool saveCheck = false;
        public MainWindow()
        {
            InitializeComponent();
            cmbFontFamily.ItemsSource = Fonts.SystemFontFamilies.OrderBy(f => f.Source);
            cmbFontSize.ItemsSource = new List<double>() { 8, 9, 10, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72 };
            cmbFontFamily.SelectedValue = Fonts.SystemFontFamilies.First();
            cmbFontSize.SelectedValue = (double)12;

        }

        private void ButtonSave_Click(object sender, RoutedEventArgs e)
        {
            saveCheck = true;
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Filter = "doc files (*.doc)|*.doc|Text Files (*.txt)|*.txt|XAML Files (*.xaml)|*.xaml|All files (*.*)|*.*";
            if (sfd.ShowDialog() == true)
            {
                TextRange doc = new TextRange(rich.Document.ContentStart, rich.Document.ContentEnd);
                using (FileStream fs = File.Create(sfd.FileName))
                {

                    if (Path.GetExtension(sfd.FileName).ToLower() == ".txt")
                        doc.Save(fs, DataFormats.Text);
                    else if (Path.GetExtension(sfd.FileName).ToLower() == ".doc")
                    {
                        doc.Save(fs, DataFormats.Text);
                    }
                    else

                        doc.Save(fs, DataFormats.Xaml);
                }
            }

        }

        private void ButtonOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "doc files (*.doc)|*.doc|Text Files (*.txt)|*.txt|XAML Files (*.xaml)|*.xaml|All files (*.*)|*.*";

            if (ofd.ShowDialog() == true)
            {
                TextRange doc = new TextRange(rich.Document.ContentStart, rich.Document.ContentEnd);
                using (FileStream fs = new FileStream(ofd.FileName, FileMode.Open))
                {

                    if (Path.GetExtension(ofd.FileName).ToLower() == ".txt")
                        doc.Load(fs, DataFormats.Text);
                    else if (Path.GetExtension(ofd.FileName).ToLower() == ".doc")
                    {
                        doc.Load(fs, DataFormats.Text);
                    }


                    else
                        doc.Load(fs, DataFormats.Xaml);
                }
            }



        }
        private void rich_SelectionChanged(object sender, RoutedEventArgs e)
        {
            object temp = rich.Selection.GetPropertyValue(Inline.FontWeightProperty);
            btnBold.IsChecked = (temp != DependencyProperty.UnsetValue) && (temp.Equals(FontWeights.Bold));
            temp = rich.Selection.GetPropertyValue(Inline.FontStyleProperty);

            btnItalic.IsChecked = (temp != DependencyProperty.UnsetValue) && (temp.Equals(FontStyles.Italic));
            temp = rich.Selection.GetPropertyValue(Inline.TextDecorationsProperty);

            btnUnderline.IsChecked = (temp != DependencyProperty.UnsetValue) && (temp.Equals(TextDecorations.Underline));
            temp = rich.Selection.GetPropertyValue(Inline.FontFamilyProperty);
            cmbFontFamily.SelectedItem = temp;
            temp = rich.Selection.GetPropertyValue(Inline.FontSizeProperty);
            cmbFontSize.Text = temp.ToString();
        }


        private void cmbFontFamily_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cmbFontFamily.SelectedItem != null)
                rich.Selection.ApplyPropertyValue(Inline.FontFamilyProperty, cmbFontFamily.SelectedItem);
        }
        private void cmbFontSize_TextChanged(object sender, TextChangedEventArgs e)
        {
            rich.Selection.ApplyPropertyValue(Inline.FontSizeProperty, cmbFontSize.Text);
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if (saveCheck == false)
            {

                var result = MessageBox.Show("Сохранить файл?", "Сохранение файла", MessageBoxButton.YesNo);
                if (result == MessageBoxResult.Yes)
                {
                    saveCheck = true;
                    SaveFileDialog sfd = new SaveFileDialog();
                    sfd.Filter = "doc files (*.doc)|*.doc|Text Files (*.txt)|*.txt|XAML Files (*.xaml)|*.xaml|All files (*.*)|*.*";
                    if (sfd.ShowDialog() == true)
                    {
                        TextRange doc = new TextRange(rich.Document.ContentStart, rich.Document.ContentEnd);
                        using (FileStream fs = File.Create(sfd.FileName))
                        {

                            if (Path.GetExtension(sfd.FileName).ToLower() == ".txt")
                                doc.Save(fs, DataFormats.Text);
                            else if (Path.GetExtension(sfd.FileName).ToLower() == ".doc")
                            {
                                doc.Save(fs, DataFormats.Text);
                            }
                            else

                                doc.Save(fs, DataFormats.Xaml);
                        }

                    }


                }
            }
        }
    }
}
