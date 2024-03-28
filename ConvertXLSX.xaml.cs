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
using System.Security.Principal;

namespace ConvertXLSX
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        String xlsx_filename;
        String csv_filename;
        public enum Header : int
        {
            CHECK_NUM,
            CHECK_DATE,
            VENDOR,
            NAME,
            CHECK_AMT
        }

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
                xlsx_filename = openFileDialog.FileName;
            }
        }

        private void btnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(xlsx_filename))
            {
                MessageBox.Show("Please pick an Excel file to convert first");
            }
            else
            {
                Workbook workbook = new Workbook();
                workbook.LoadFromFile(xlsx_filename);
                Worksheet sheet = workbook.Worksheets[0];

                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "CSV|*.csv";
                if (saveFileDialog.ShowDialog() == true)
                {
                    csv_filename = saveFileDialog.FileName;
                    sheet.SaveToFile(csv_filename, "|", Encoding.UTF8); // switch to | instead of , for error checking
                }
            }            
        }

        private void btnInsert_Click(object sender, RoutedEventArgs e)
        {
            if (String.IsNullOrEmpty(csv_filename))
            {
                MessageBox.Show("No converted CSV file found");
            }
            else if (String.IsNullOrEmpty(txtStatus.Text) || String.IsNullOrEmpty(txtAccount.Text))
            {
                MessageBox.Show("Please fill in both fields", "Missing Data");
            }
            else
            {
                List<string> lines = new List<string>();
                string[] prepend = { txtStatus.Text, txtAccount.Text };

                // Read all lines from the file
                string[] allLines = File.ReadAllLines(csv_filename);
                // Skip the first line (headers) and process the rest
                // Headers should match enum Headers
                var processedLines = allLines.Skip(1).Select(line =>
                {
                    string[] split = line.Split('|');

                    // Check if the field is in MM/DD/YYYY format and convert it to MMDDYY
                    if (split[(int)Header.CHECK_DATE].Contains('/'))
                    {
                        string[] date = split[(int)Header.CHECK_DATE].Split('/');
                        date[0] = date[0].PadLeft(2, '0'); // Add leading zero for month
                        date[1] = date[1].PadLeft(2, '0'); // Add leading zero for day
                        date[2] = date[2].Substring(2);    // Trim year to YY
                        split[(int)Header.CHECK_DATE] = string.Concat(date);    // Reassemble
                    }

                    // Check that the name field doesn't have commas
                    if (split[(int)Header.NAME].Contains(','))
                    {
                        string[] name = split[(int)Header.NAME].Split(',');
                        name[1] = string.Concat(name[1].Split()); // remove whitespace (thank you stackoverflow)
                        split[(int)Header.NAME] = name[1] + " " + name[0];
                    }

                    // Overwrite the vendor id with check amount
                    split[^3] = split[^1];
                    // Drop the last element
                    return prepend.Concat(split.Take(split.Length - 1));
                });

                // Write the processed lines back to the file
                File.WriteAllLines(csv_filename, processedLines.Select(split => string.Join(",", split)));

                MessageBox.Show("Done - please review your file to verify");

                // Code to open the file
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = csv_filename,
                    UseShellExecute = true,
                    Verb = "open"
                });
            }
        }
    }
}
