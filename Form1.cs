using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Speech.Recognition;
using System.Threading;
using System.Windows.Forms;

namespace VoicePhasmo
{
    public partial class Form1 : Form
    {
        SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();

        [DllImport("user32")]
        public static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        public const int MOUSEEVENTF_MOVE = 0x0001; /* mouse move */
        public const int MOUSEEVENTF_LEFTDOWN = 0x0002; /* left button down */
        public const int MOUSEEVENTF_LEFTUP = 0x0004; /* left button up */
        public const int MOUSEEVENTF_RIGHTDOWN = 0x0008; /* right button down */

        public void ObjetsClick(int x, int y)
        {
            SetCursorPos(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, x, y, 0, 0);
            mouse_event(MOUSEEVENTF_LEFTUP, x, y, 0, 0);
            Thread.Sleep(120);
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void btnEnable_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsync(RecognizeMode.Multiple);
            btnDisable.Visible = true;
            richTextBox1.Text = "- Log - Voice Enabled";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Choices commands = new Choices();
            commands.Add(new string[] { "Coucou", "Salut", "mon nom", "ouvre fasmo", "Lance Phasmo", "mets les objets", "ajoute les objets", "partie privée", "Lance une partie privée", "arrête d'écouter", "Command", "test" });
            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Culture = new System.Globalization.CultureInfo("fr-FR");
            gBuilder.Append(commands);
            Grammar grammar = new Grammar(gBuilder);

            recEngine.LoadGrammarAsync(grammar);
            recEngine.SetInputToDefaultAudioDevice();
            recEngine.SpeechRecognized += recEngine_SpeechRecognized;

            string screenWidth = Screen.PrimaryScreen.Bounds.Width.ToString();
            string screenHeight = Screen.PrimaryScreen.Bounds.Height.ToString();
            label1.Text = "Resolution: " + screenWidth + " x " + screenHeight;


            //label2.Text = "Micro: " + device.DeviceFriendlyName;
        }

        void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
            var screen = Screen.PrimaryScreen.Bounds.Height;
            switch (e.Result.Text)
            {
                case "Command":
                    richTextBox1.Text += "\n[Help] " +
                        "arrête d'écouter : Disable voirce recognition. \n" +
                        "Coucou / Salut : Say Hellow in back.\n" +
                        "Mon Nom : return PsykoDev. \n" +
                        "Ouvre Phasmo / Lance Phasmo : Launch Phasmophobia from Steam cli. \n" +
                        "Mets les objets / Ajoute les objets : Put all item from yout inventory. \n" +
                        "Partie privée / Lance une partie privée : create a private game and add all items. \n";
                    break;

                case "arrête d'écouter":
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'arrête d'écouter'";
                    }
                    recEngine.RecognizeAsyncStop();
                    richTextBox1.Text = "- Log - Disabled Voice";
                    break;

                case "Coucou":
                case "Salut":
                    richTextBox1.Text += $"\n[Size] {screen}";
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'Coucou'";
                    }
                    richTextBox1.Text += "\n[LOG]Hellow !";
                    break;

                case "mon nom":
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'Mon Nom'";
                    }
                    richTextBox1.Text += "\n[LOG]PsykoDev";
                    break;

                case "ouvre fasmo":
                case "Lance Phasmo":
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'Ouvre Phasmo'";
                    }
                    richTextBox1.Text += "\n[LOG]Ouverture de Phasmophobia";
                    Process cmd = new Process();
                    cmd.StartInfo.FileName = @"cmd.exe";
                    cmd.StartInfo.WorkingDirectory = @"C:\Program Files (x86)\Steam\";
                    cmd.StartInfo.Arguments = @" /c steam -applaunch 739630 ";
                    cmd.Start();
                    break;

                case "mets les objets":
                case "ajoute les objets":
                    richTextBox1.Text += $"\n[Size] {screen}";
                    richTextBox1.Text += "\n[LOG]Placement de tout les items dispo";
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'Met les objets'";
                    }
                    if (screen != 1440)
                    {
                        ObjetHD();
                        richTextBox1.Text += "\n[LOG]Finit";
                        return;
                    }
                    Objet2k();
                    richTextBox1.Text += "\n[LOG]Finit";
                    break;

                case "partie privée":
                case "Lance une partie privée":
                    richTextBox1.Text += "\n[LOG]Lancement d'un partie privée";
                    richTextBox1.Text += $"\n[Size] {screen}";
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'Partie Privée'";
                    }
                    if (screen != 1440)
                    {
                        PvGameHD();
                        richTextBox1.Text += "\n[LOG]Finit";
                        return;
                    }
                    PvGame2k();
                    richTextBox1.Text += "\n[LOG]Finit";
                    break;
                case "test":
                    richTextBox1.Text += "\n[LOG]Lancement d'un Test";
                    Test();
                    break;
            }
        }

        public void Test()
        {
            ObjetsClick(1522, 963);
        }

        private void btnDisable_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsyncStop();
            btnEnable.Visible = true;
            richTextBox1.Text = "- Log - Voice Disabled";
        }

        public void PvGame2k()
        {
            ObjetsClick(1264, 484);//Play
            ObjetsClick(1315, 394);//Private Game   
            Thread.Sleep(2000);
            Objet2k();
        }

        public void Objet2k()
        {
            ObjetsClick(1522, 963);//ADD
            ObjetsClick(1204, 472);//EMf
            ObjetsClick(1204, 511);//Flash
            ObjetsClick(1204, 547);//Photo
            ObjetsClick(1204, 591);//Lighter
            ObjetsClick(1204, 629);//Candle
            ObjetsClick(1204, 668);//UV light
            ObjetsClick(1204, 706);//Crucifix
            ObjetsClick(1204, 746);//Video cam
            ObjetsClick(1204, 785);//Spirit box
            ObjetsClick(1204, 823);//Salt
            ObjetsClick(1204, 862);//Smudge stick
            ObjetsClick(1204, 908);//Tripod
            ObjetsClick(1204, 940);//Strong flash
            ObjetsClick(1204, 983);//Motion sensor
            ObjetsClick(1204, 1021);//Sound sensor
            ObjetsClick(1912, 469);//Thermo
            ObjetsClick(1912, 509);//Sanity pills
            ObjetsClick(1912, 553);//Ghost writing book
            ObjetsClick(1912, 592);//Infrared light
            ObjetsClick(1912, 631);//Parabolic micro
            ObjetsClick(1912, 667);//Glowstick
            ObjetsClick(1912, 708);//Head mounted camera
            ObjetsClick(1266, 1121);//Back
        }



        public void PvGameHD()
        {
            ObjetsClick(950, 370);//Play
            ObjetsClick(990, 300);//Private Game
            Thread.Sleep(2000);
            ObjetHD();
        }

        public void ObjetHD()
        {
            ObjetsClick(1150, 700); //ADD
            ObjetsClick(900, 350);//EMf
            ObjetsClick(900, 380);//Flash
            ObjetsClick(900, 410);//Photo
            ObjetsClick(900, 440);//Lighter
            ObjetsClick(900, 470);//Candle
            ObjetsClick(900, 500);//UV light
            ObjetsClick(900, 530);//Crucifix
            ObjetsClick(900, 560);//Video cam
            ObjetsClick(900, 590);//Spirit box
            ObjetsClick(900, 620);//Salt
            ObjetsClick(900, 650);//Smudge sticks
            ObjetsClick(900, 680);//Tripod
            ObjetsClick(900, 710);//Strong flash
            ObjetsClick(900, 740);//Motion sensor
            ObjetsClick(900, 770);//Sound sensor
            ObjetsClick(1435, 350);//Thermo
            ObjetsClick(1435, 380);//Sanity pills
            ObjetsClick(1435, 410);//Ghost writing book
            ObjetsClick(1435, 440);//Infrared light
            ObjetsClick(1435, 470);//Parabolic micro
            ObjetsClick(1435, 500);//Glowstick
            ObjetsClick(1435, 530);//Head mounted camera
            ObjetsClick(940, 830);//Back
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
        }
    }
}