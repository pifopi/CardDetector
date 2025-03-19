using CsvHelper;
using OpenCvSharp.WpfExtensions;
using System.IO;

namespace CardDetector.CardDetector
{
    enum CardLocation
    {
        TopLeft,
        TopMiddle,
        TopRight,
        BottomLeft,
        BottomRight
    }

    enum Pack
    {
        Mewtwo,
        Dracaufeu,
        Pikachu,
        Mew,
        Dialga,
        Palkia,
        Arceus
    }

    class CardTemplate
    {
        public Pack pack { get; private set; }

        public string ID { get; private set; } = "";

        public OpenCvSharp.Mat template { get; private set; } = new();

        public string informationToDisplay = "";

        public CardTemplate(Pack pack, string ID, OpenCvSharp.Mat template)
        {
            this.pack = pack;
            this.ID = ID;
            this.template = template;
        }
    };


    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private List<CardTemplate> templates = new();

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool AddClipboardFormatListener(IntPtr hwnd);

        private void MainWindow_SourceInitialized(object sender, EventArgs e)
        {
            IntPtr handle = (new System.Windows.Interop.WindowInteropHelper(this)).Handle;
            System.Windows.Interop.HwndSource.FromHwnd(handle)?.AddHook(WndProc);
            AddClipboardFormatListener(handle);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            int WM_CLIPBOARDUPDATE = 0x031D;

            if (msg == WM_CLIPBOARDUPDATE)
            {
                if (System.Windows.Clipboard.ContainsImage())
                {
                    ReadGP();
                }
            }
            return IntPtr.Zero;
        }


        private static OpenCvSharp.Rect GetRect(CardLocation cardLocation)
        {
            int width = 367;
            int height = 512;
            int padding = 20;
            int shift = 193;

            switch (cardLocation)
            {
                case CardLocation.TopLeft:
                    return new(0, 0, width, height);
                case CardLocation.TopMiddle:
                    return new(width + padding, 0, width, height);
                case CardLocation.TopRight:
                    return new(width * 2 + padding * 2, 0, width, height);
                case CardLocation.BottomLeft:
                    return new(shift, height + padding, width, height);
                case CardLocation.BottomRight:
                    return new(shift + width + padding, height + padding, width, height);
            }
            throw new Exception("Invalid card ID");
        }

        private static OpenCvSharp.Mat Extract(OpenCvSharp.Mat input, CardLocation cardLocation)
        {
            return new OpenCvSharp.Mat(input, GetRect(cardLocation));
        }

        private static OpenCvSharp.Mat ClipboardToMat()
        {
            using (MemoryStream stream = new MemoryStream())
            {
                System.Windows.Media.Imaging.BitmapEncoder encoder = new System.Windows.Media.Imaging.BmpBitmapEncoder();
                encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
                encoder.Save(stream);
                System.Drawing.Bitmap bitmap = new(stream);
                return OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap);
            }
        }

        private void LoadTemplates()
        {
            foreach (Pack pack in Enum.GetValues(typeof(Pack)))
            {
                string[] filenames = Directory.GetFiles($"data/{pack}", "*.webp", SearchOption.AllDirectories);
                foreach (string filename in filenames)
                {
                    CardTemplate template = new(pack, Path.GetFileNameWithoutExtension(filename), OpenCvSharp.Cv2.ImRead(filename));
                    templates.Add(template);
                }
            }
        }

        private void LoadCsv()
        {
            string[] filenames = Directory.GetFiles($"config", "*.csv", SearchOption.AllDirectories);
            var config = new CsvHelper.Configuration.CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture);
            foreach (string filename in filenames)
            {
                using (var reader = new StreamReader(filename))
                using (var csv = new CsvReader(reader, config))
                {
                    for (int i = 0; i < 5; i++)
                    {
                        csv.Read();
                    }

                    csv.ReadHeader();

                    while (csv.Read())
                    {
                        string? cardID = csv.GetField("Card ID");
                        string? copies = csv.GetField("Copies");
                        if (cardID == null || copies == null)
                        {
                            throw new Exception($"Invalid csv : {filename}");
                        }

                        foreach (CardTemplate template in templates)
                        {
                            if (template.ID == cardID)
                            {
                                template.informationToDisplay = copies;
                            }
                        }
                    }
                }
            }
        }

        private void ReadGP()
        {
            OpenCvSharp.Mat input = ClipboardToMat();
            InputImage.Source = input.ToBitmapSource();

            foreach (CardLocation cardLocation in Enum.GetValues(typeof(CardLocation)))
            {
                OpenCvSharp.Mat card = Extract(input, cardLocation);
                //card.SaveImage($"{id.ToString()}.png");

                foreach (CardTemplate template in templates)
                {
                    OpenCvSharp.Mat result = card.MatchTemplate(template.template, OpenCvSharp.TemplateMatchModes.CCoeffNormed);
                    OpenCvSharp.Cv2.MinMaxLoc(result, out _, out double maxVal, out _, out OpenCvSharp.Point maxLoc);

                    if (maxVal > 0.9)
                    {
                        switch (cardLocation)
                        {
                            case CardLocation.TopLeft:
                                TopLeftText.Text = template.informationToDisplay;
                                TopLeftImage.Source = template.template.ToBitmapSource();
                                break;
                            case CardLocation.TopMiddle:
                                TopMiddleText.Text = template.informationToDisplay;
                                TopMiddleImage.Source = template.template.ToBitmapSource();
                                break;
                            case CardLocation.TopRight:
                                TopRightText.Text = template.informationToDisplay;
                                TopRightImage.Source = template.template.ToBitmapSource();
                                break;
                            case CardLocation.BottomLeft:
                                BottomLeftText.Text = template.informationToDisplay;
                                BottomLeftImage.Source = template.template.ToBitmapSource();
                                break;
                            case CardLocation.BottomRight:
                                BottomRightText.Text = template.informationToDisplay;
                                BottomRightImage.Source = template.template.ToBitmapSource();
                                break;
                        }
                        continue;
                    }
                }
            }
        }

        public MainWindow()
        {
            InitializeComponent();
            SourceInitialized += MainWindow_SourceInitialized;

            LoadTemplates();
            LoadCsv();
        }
    }
}