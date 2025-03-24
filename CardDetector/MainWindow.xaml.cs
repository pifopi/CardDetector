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

    class CardTemplate
    {
        public string ID { get; private set; } = "";

        public OpenCvSharp.Mat image { get; private set; } = new();

        public OpenCvSharp.Mat descriptors { get; private set; } = new();

        public string informationToDisplay = "";

        public CardTemplate(string ID, OpenCvSharp.Mat image, OpenCvSharp.Mat descriptors)
        {
            this.ID = ID;
            this.image = image;
            this.descriptors = descriptors;
        }
    };

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        private readonly List<CardTemplate> templates = new();
        private readonly OpenCvSharp.ORB orb = OpenCvSharp.ORB.Create();

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
            const int WM_CLIPBOARDUPDATE = 0x031D;

            if (msg == WM_CLIPBOARDUPDATE && System.Windows.Clipboard.ContainsImage())
            {
                ReadGP();
            }
            return IntPtr.Zero;
        }

        private static OpenCvSharp.Rect GetRect(CardLocation cardLocation)
        {
            const int width = 367;
            const int height = 512;
            const int padding = 20;
            const int shift = 193;

            return cardLocation switch
            {
                CardLocation.TopLeft => new(0, 0, width, height),
                CardLocation.TopMiddle => new(width + padding, 0, width, height),
                CardLocation.TopRight => new(width * 2 + padding * 2, 0, width, height),
                CardLocation.BottomLeft => new(shift, height + padding, width, height),
                CardLocation.BottomRight => new(shift + width + padding, height + padding, width, height),
                _ => throw new ArgumentException("Invalid card location", nameof(cardLocation)),
            };
        }

        private static OpenCvSharp.Mat Extract(OpenCvSharp.Mat input, CardLocation cardLocation)
        {
            return new OpenCvSharp.Mat(input, GetRect(cardLocation));
        }

        private static OpenCvSharp.Mat ClipboardToMat()
        {
            using MemoryStream stream = new();
            System.Windows.Media.Imaging.BitmapEncoder encoder = new System.Windows.Media.Imaging.BmpBitmapEncoder();
            encoder.Frames.Add(System.Windows.Media.Imaging.BitmapFrame.Create(System.Windows.Clipboard.GetImage()));
            encoder.Save(stream);
            using System.Drawing.Bitmap bitmap = new(stream);
            return OpenCvSharp.Extensions.BitmapConverter.ToMat(bitmap);
        }

        private void LoadTemplates()
        {
            string[] filenames = Directory.GetFiles("data", "*.webp", SearchOption.AllDirectories);
            foreach (string filename in filenames)
            {
                OpenCvSharp.Mat image = OpenCvSharp.Cv2.ImRead(filename);
                OpenCvSharp.Mat descriptors = new();
                orb.DetectAndCompute(image, null, out _, descriptors);

                templates.Add(new(Path.GetFileNameWithoutExtension(filename), image, descriptors));
            }
        }

        private void LoadCsv()
        {
            string[] filenames = Directory.GetFiles("config", "*.csv", SearchOption.AllDirectories);
            var config = new CsvHelper.Configuration.CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture);
            foreach (string filename in filenames)
            {
                using var reader = new StreamReader(filename);
                using var csv = new CsvHelper.CsvReader(reader, config);

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
                        throw new Exception($"Invalid csv: {filename}");
                    }

                    foreach (CardTemplate template in templates)
                    {
                        if (template.ID == cardID)
                        {
                            template.informationToDisplay = copies;
                            break;
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

                OpenCvSharp.Mat descriptors = new();
                orb.DetectAndCompute(card, null, out _, descriptors);

                int bestMatchCount = 0;
                CardTemplate bestMatch = templates[0];

                foreach (CardTemplate template in templates)
                {
                    OpenCvSharp.BFMatcher matcher = new(OpenCvSharp.NormTypes.Hamming);
                    OpenCvSharp.DMatch[] matches = matcher.Match(descriptors, template.descriptors);

                    int goodMatches = matches.Count(m => m.Distance < 32);

                    if (goodMatches > bestMatchCount)
                    {
                        bestMatchCount = goodMatches;
                        bestMatch = template;
                    }
                }

                switch (cardLocation)
                {
                    case CardLocation.TopLeft:
                        TopLeftText.Text = bestMatch.informationToDisplay;
                        TopLeftImage.Source = bestMatch.image.ToBitmapSource();
                        break;
                    case CardLocation.TopMiddle:
                        TopMiddleText.Text = bestMatch.informationToDisplay;
                        TopMiddleImage.Source = bestMatch.image.ToBitmapSource();
                        break;
                    case CardLocation.TopRight:
                        TopRightText.Text = bestMatch.informationToDisplay;
                        TopRightImage.Source = bestMatch.image.ToBitmapSource();
                        break;
                    case CardLocation.BottomLeft:
                        BottomLeftText.Text = bestMatch.informationToDisplay;
                        BottomLeftImage.Source = bestMatch.image.ToBitmapSource();
                        break;
                    case CardLocation.BottomRight:
                        BottomRightText.Text = bestMatch.informationToDisplay;
                        BottomRightImage.Source = bestMatch.image.ToBitmapSource();
                        break;
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
