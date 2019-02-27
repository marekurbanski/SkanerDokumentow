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
using System.Windows.Navigation;
using System.Windows.Shapes;
using WIA;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Threading;


namespace Skaner
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Skanuj()
        {
            var deviceManager = new DeviceManager();

            DeviceInfo firstScannerAvailable = null;

            for (int i = 1; i <= deviceManager.DeviceInfos.Count; i++)
            {
                if (deviceManager.DeviceInfos[i].Type != WiaDeviceType.ScannerDeviceType)
                {
                    continue;
                }

                firstScannerAvailable = deviceManager.DeviceInfos[i];

                break;
            }

            var device = firstScannerAvailable.Connect();

            var scannerItem = device.Items[1];

            int resolution = 150;
            int width_pixel = 1250;
            int height_pixel = 1700;
            int color_mode = 1;
            AdjustScannerSettings(scannerItem, resolution, 0, 0, width_pixel, height_pixel, 0, 0, color_mode);

            var imageFile = (ImageFile)scannerItem.Transfer(FormatID.wiaFormatJPEG);

            numer_zlecenia.Text = numer_zlecenia.Text.Replace("\\", "_");
            numer_zlecenia.Text = numer_zlecenia.Text.Replace("/", "_");
            //numer_zlecenia.Text = numer_zlecenia.Text.ToUpper();
            DoEvents();

            if (!Directory.Exists("C:\\Users\\USER\\Dysk Google\\Dokumenty\\" + numer_zlecenia.Text))
            {
                Directory.CreateDirectory("C:\\Users\\USER\\Dysk Google\\Dokumenty\\" + numer_zlecenia.Text);
            }

            int numer = 1;
            String co = "inne";

            if (protokol.IsChecked == true)
                co = "protokol";
            if (oplaty_drogowe.IsChecked == true)
                co = "oplaty";
            if (notatki.IsChecked == true)
                co = "notatki";

            String sciezka = "Dokumenty";
            if (wyjasnienie.IsChecked == true)
            {
                sciezka = "Wyjasnienia";
                co = "dokument";
            }

            Sprawdzanie:

            var path = @"C:\\Users\\USER\\Dysk Google\\"+ sciezka + "\\" + numer_zlecenia.Text + "\\" + co.ToString() +"_"+numer.ToString() + ".jpeg";



                if (File.Exists(path))
            {
                numer++;
                goto Sprawdzanie;
            }

            imageFile.SaveFile(path);
            MessageBox.Show("Zrobione","Info");
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (numer_zlecenia.Text == "")
            {
                MessageBox.Show("Wpisz numer zlecenia","Uwaga !!!");
            }
            else
            {
                try
                {
                    wait_on();
                    Skanuj();
                    wait_off();
                }
                catch(COMException err)
                {
                    wait_off();
                    uint errorCode = (uint)err.ErrorCode;

                    if (errorCode == 0x80210006)
                    {
                        MessageBox.Show("Skaner nie jest dostepny albo jest zajety", "Błąd");
                    }
                    else if (errorCode == 0x80210064)
                    {
                        MessageBox.Show("Skanowanie zostało anulowane","Błąd");
                    }
                    MessageBox.Show(err.ToString(), "Błąd !!!");
                }
            }
        }

        private static void AdjustScannerSettings(IItem scannnerItem, int scanResolutionDPI, int scanStartLeftPixel, int scanStartTopPixel, int scanWidthPixels, int scanHeightPixels, int brightnessPercents, int contrastPercents, int colorMode)
        {
            const string WIA_SCAN_COLOR_MODE = "6146";
            const string WIA_HORIZONTAL_SCAN_RESOLUTION_DPI = "6147";
            const string WIA_VERTICAL_SCAN_RESOLUTION_DPI = "6148";
            const string WIA_HORIZONTAL_SCAN_START_PIXEL = "6149";
            const string WIA_VERTICAL_SCAN_START_PIXEL = "6150";
            const string WIA_HORIZONTAL_SCAN_SIZE_PIXELS = "6151";
            const string WIA_VERTICAL_SCAN_SIZE_PIXELS = "6152";
            const string WIA_SCAN_BRIGHTNESS_PERCENTS = "6154";
            const string WIA_SCAN_CONTRAST_PERCENTS = "6155";
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_RESOLUTION_DPI, scanResolutionDPI);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_START_PIXEL, scanStartLeftPixel);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_START_PIXEL, scanStartTopPixel);
            SetWIAProperty(scannnerItem.Properties, WIA_HORIZONTAL_SCAN_SIZE_PIXELS, scanWidthPixels);
            SetWIAProperty(scannnerItem.Properties, WIA_VERTICAL_SCAN_SIZE_PIXELS, scanHeightPixels);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_BRIGHTNESS_PERCENTS, brightnessPercents);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_CONTRAST_PERCENTS, contrastPercents);
            SetWIAProperty(scannnerItem.Properties, WIA_SCAN_COLOR_MODE, colorMode);
        }

        private static void SetWIAProperty(IProperties properties, object propName, object propValue)
        {
            Property prop = properties.get_Item(ref propName);
            prop.set_Value(ref propValue);
        }

        private void wait_on()
        {
            skanuj.IsEnabled = false;
            skanuj.Content = "Czekaj...";
            DoEvents();
        }

        private void wait_off()
        {
            skanuj.IsEnabled = true;
            skanuj.Content = "Skanuj";
            DoEvents();
        }


        public static void DoEvents()
        {
            Application.Current.Dispatcher.Invoke(DispatcherPriority.Render,
                                                      new Action(delegate { }));
        }

    }
}
