using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace OpenCustomizedApps
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
        [DllImport("user32.dll")]
        private static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, uint dwExtraInfo);


        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            // Create a list to store the data
            var dataList = new List<DataItem>();

            // Collect data from the TextBoxes and ComboBoxes
            dataList.Add(new DataItem
            {
                TextBoxValue = TextBox1.Text,
                ComboBoxValue = (Option1.SelectedItem as ComboBoxItem)?.Content.ToString()
            });
            dataList.Add(new DataItem
            {
                TextBoxValue = TextBox2.Text,
                ComboBoxValue = (Option2.SelectedItem as ComboBoxItem)?.Content.ToString()
            });
            dataList.Add(new DataItem
            {
                TextBoxValue = TextBox3.Text,
                ComboBoxValue = (Option3.SelectedItem as ComboBoxItem)?.Content.ToString()
            });

            dataList.Add(new DataItem
            {
                TextBoxValue = TextBox4.Text,
                ComboBoxValue = GetRichTextBoxText().ToString()
            });

            // Convert the list to an array if needed
            var dataArray = dataList.ToArray();

            // Do something with the data (e.g., display it, save it, etc.)
            // For demonstration, we will just show a message with the collected data
            string message = "Saved";
            
            MessageBox.Show(message);
        }
        private void saysomething_TextChanged(object sender, TextChangedEventArgs e)
        {
            
            TextRange textRange = new TextRange(saysomething.Document.ContentStart, saysomething.Document.ContentEnd);
            string inputText = textRange.Text.Trim();

            // Check each TextBox for a match (case-insensitive)
            if (string.Equals(TextBox1.Text, inputText, StringComparison.OrdinalIgnoreCase))
            {
                //  MessageBox.Show($"Match found: {GetComboBoxValue(Option1)}");
                RunApplication(GetComboBoxValue(Option1));
            }
            else if (string.Equals(TextBox2.Text, inputText, StringComparison.OrdinalIgnoreCase))
            {
                //MessageBox.Show($"Match found: {GetComboBoxValue(Option2)}");
                RunApplication(GetComboBoxValue(Option2));
            }
            else if (string.Equals(TextBox3.Text, inputText, StringComparison.OrdinalIgnoreCase))
            {
                //  MessageBox.Show($"Match found: {GetComboBoxValue(Option3)}");
                RunApplication(GetComboBoxValue(Option3));
            }
            else if (string.Equals(TextBox4.Text, inputText, StringComparison.OrdinalIgnoreCase))
            {
                saysomething.Document.Blocks.Clear();
                string richTextBoxText = GetRichTextBoxText();
                saysomething.Document.Blocks.Add(new Paragraph(new Run(richTextBoxText)));
               
            }
        }

        private string GetComboBoxValue(ComboBox comboBox)
        {
            return (comboBox.SelectedItem as ComboBoxItem)?.Content.ToString() ?? "No selection";
        }

        private void RunApplication(string appName)
        {
            string appPath = string.Empty;
            string arguments = string.Empty;

            switch (appName)
            {
                case "Word":
                    //Process.Start("WINWORD.EXE");
                    appPath = "C:\\Program Files\\Microsoft Office\\root\\Office16\\WINWORD.EXE"; 
                    break;
                case "Google Chrome":
                    appPath = "chrome.exe";
                    break;
                case "File Manager":
                    appPath = "explorer.exe";
                    break;
                case "ChatGpt":
                    // Assume ChatGpt is a local application. Adjust path as needed
                    //  Process.Start("C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe", "https://chatgpt.com/");
                    appPath = "chrome.exe";
                    arguments = "https://chatgpt.com/";
                    break;
                case "Mobile Link":
                    // Example: Assume Mobile Link is a local application. Adjust path as needed
                    appPath = "MobileLink.exe";
                    break;
                default:
                   // MessageBox.Show("Application not recognized.");
                    return;
            }

            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo
                {
                    FileName = appPath,
                    Arguments = arguments,
                    UseShellExecute = true
                };

                Process.Start(startInfo);
                saysomething.Document.Blocks.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to start application: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        private string GetRichTextBoxText()
        {
            
            TextRange textRange = new TextRange(Project.Document.ContentStart, Project.Document.ContentEnd);
            return textRange.Text.Trim(); 
        }


        private void SendWinHKeys()
        {
            
            const int VK_H = 0x48; 
            const int VK_LWIN = 0x5B; 
            const int KEYEVENTF_KEYUP = 0x0002;

            keybd_event((byte)VK_LWIN, 0, 0, 0); 
            keybd_event((byte)VK_H, 0, 0, 0);    
            keybd_event((byte)VK_H, 0, KEYEVENTF_KEYUP, 0);   
            keybd_event((byte)VK_LWIN, 0, KEYEVENTF_KEYUP, 0); 
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            saysomething.Document.Blocks.Clear();
            saysomething.Focus();
            SendWinHKeys();
        }
    }


    public class DataItem
    {
        public string TextBoxValue { get; set; }
        public string ComboBoxValue { get; set; }
    }
}