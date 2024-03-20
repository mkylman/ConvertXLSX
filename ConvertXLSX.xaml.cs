using System;
using Microsoft.Win32;  // for OpenFileDialog and SaveFileDialog
using Spire.Xls;        // for handling Excel files
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Linq.Expressions;

namespace ConvertXLSX
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        String filename;
        String newfilename;

        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnGetFile_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel|*.xlsx|Excel|*.xls";
            if (openFileDialog.ShowDialog() == true)
            {
                filename = openFileDialog.FileName;
            }
        }

        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(filename))
            {
                MessageBox.Show("Please pick an Excel file to convert first");
            }
            else
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(filename);
                Worksheet sheet = workbook.Worksheets[0];

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "CSV|*.csv";
                if (saveFileDialog.ShowDialog() == true)
                {
                    newfilename = saveFileDialog.FileName;
                    sheet.SaveToFile(newfilename, ",", Encoding.UTF8);
                }
            }            
        }

        private void btnInsert_Click(object sender, RoutedEventArgs e)
        {

            if (String.IsNullOrEmpty(newfilename))
            {
                MessageBox.Show("No converted CSV file found");
            }
            else if (String.IsNullOrEmpty(txtStatus.Text) || String.IsNullOrEmpty(txtAccount.Text))
            {
                MessageBox.Show("Please fill in both fields", "Missing Data");
            }
            else
            {
                List<String> lines = new List<String>();
                String[] prepend = { txtStatus.Text, txtAccount.Text };

                using (StreamReader reader = new StreamReader(newfilename))
                {
                    String line;
                    bool first_line = true;

                    while ((line = reader.ReadLine()) != null)
                    {
                        if (first_line == false) // First line in file will be headers, which we don't want
                        {
                            String[] split = line.Split(',');
                            if (split[1].Contains('/')) // check if the field is mM/dD/YYYY
                            {
                                String[] date = split[1].Split('/');
                                for (int i = 0; i < 2; i++)
                                {
                                    if (date[i].Length < 2) // add leading zeroes for MMDDYY format
                                    {
                                        date[i] = "0" + date[i];
                                    }
                                }
                                date[2] = date[2].Substring(2); // trim year to YY
                                split[1] = String.Join("", date); // reassemble
                            }
                            Array.Resize(ref split, split.Length - 1); // drop the last element (vendor id)
                            split = prepend.ToList().Concat(split.ToList()).ToArray();
                            line = String.Join(",", split);
                            lines.Add(line);
                        } else { first_line = false; }
                    }
                }

                using (StreamWriter writer = new StreamWriter(newfilename, false))
                {
                    foreach (String line in lines)
                        writer.WriteLine(line);
                }
                MessageBox.Show("Done - please review your file to verify");
            }
        }
    }
}
