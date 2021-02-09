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
        private SpeechRecognitionEngine recEngine = new SpeechRecognitionEngine();

        [DllImport("user32")]
        public static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int cButtons, int dwExtraInfo);

        private const int MOUSEEVENTF_MOVE = 0x0001; /* mouse move */
        private const int MOUSEEVENTF_LEFTDOWN = 0x0002; /* left button down */
        private const int MOUSEEVENTF_LEFTUP = 0x0004; /* left button up */
        private const int MOUSEEVENTF_RIGHTDOWN = 0x0008; /* right button down */

        public void ObjetsClick()
        {
            int x = 100;
            int y = 100;
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
            richTextBox1.Text = "- Log - Enabled Voice";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Choices commands = new Choices();
            commands.Add(new string[] { "Coucou", "Salut", "mon nom", "ouvre fasmo", "Lance Phasmo", "mets les objets", "ajoute les objets", "partie privée", "Lance une partie privée", "arrête d'écouter", "Command" });
            GrammarBuilder gBuilder = new GrammarBuilder();
            gBuilder.Append(commands);
            Grammar grammar = new Grammar(gBuilder);

            recEngine.LoadGrammarAsync(grammar);
            recEngine.SetInputToDefaultAudioDevice();
            recEngine.SpeechRecognized += recEngine_SpeechRecognized;
        }

        private void recEngine_SpeechRecognized(object sender, SpeechRecognizedEventArgs e)
        {
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

                case "arrete d'ecouter":
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'arrête d'écouter'";
                    }
                    recEngine.RecognizeAsyncStop();
                    richTextBox1.Text = "- Log - Disabled Voice";
                    break;

                case "Coucou":
                case "Salut":
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
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'Met les objets'";
                    }
                    richTextBox1.Text += "\n[LOG]Placement de tout les items dispo";
                    Objet();
                    richTextBox1.Text += "\n[LOG]Finit";
                    break;

                case "partie privée":
                case "Lance une partie privée":
                    if (checkBox1.Checked)
                    {
                        richTextBox1.Text += "\n[Recognized] 'Partie Privée'";
                    }
                    richTextBox1.Text += "\n[LOG]Lancement d'un partie privée";
                    PvGame();
                    break;
            }
        }

        private void btnDisable_Click(object sender, EventArgs e)
        {
            recEngine.RecognizeAsyncStop();
            btnEnable.Visible = true;
            richTextBox1.Text = "- Log - Disabled Voice";
        }

        public void PvGame()
        {
            Cursor.Position = new Point(1264, 484);//Play
            ObjetsClick();
            Cursor.Position = new Point(1315, 394);//Private Game
            ObjetsClick();
            Thread.Sleep(2000);
            Objet();
        }

        public void Objet()
        {
            Cursor.Position = new Point(1522, 963); //ADD
            ObjetsClick();
            Cursor.Position = new Point(1204, 472);//EMf
            ObjetsClick();
            Cursor.Position = new Point(1204, 511);//Flash
            ObjetsClick();
            Cursor.Position = new Point(1204, 547);//Photo
            ObjetsClick();
            Cursor.Position = new Point(1204, 591);//Lighter
            ObjetsClick();
            Cursor.Position = new Point(1204, 629);//Candle
            ObjetsClick();
            Cursor.Position = new Point(1204, 668);//UV light
            ObjetsClick();
            Cursor.Position = new Point(1204, 706);//Crucifix
            ObjetsClick();
            Cursor.Position = new Point(1204, 746);//Video cam
            ObjetsClick();
            Cursor.Position = new Point(1204, 785);//Spirit box
            ObjetsClick();
            Cursor.Position = new Point(1204, 823);//Salt
            ObjetsClick();
            Cursor.Position = new Point(1204, 862);//Smudge sticks
            ObjetsClick();
            Cursor.Position = new Point(1204, 908);//Tripod
            ObjetsClick();
            Cursor.Position = new Point(1204, 940);//Strong flash
            ObjetsClick();
            Cursor.Position = new Point(1204, 983);//Motion sensor
            ObjetsClick();
            Cursor.Position = new Point(1204, 1021);//Sound sensor
            ObjetsClick();
            Cursor.Position = new Point(1912, 469);//Thermo
            ObjetsClick();
            Cursor.Position = new Point(1912, 509);//Sanity pills
            ObjetsClick();
            Cursor.Position = new Point(1912, 553);//Ghost writing book
            ObjetsClick();
            Cursor.Position = new Point(1912, 592);//Infrared light
            ObjetsClick();
            Cursor.Position = new Point(1912, 631);//Parabolic micro
            ObjetsClick();
            Cursor.Position = new Point(1912, 667);//Glowstick
            ObjetsClick();
            Cursor.Position = new Point(1912, 708);//Head mounted camera
            ObjetsClick();
            Cursor.Position = new Point(1266, 1121);//Back
            ObjetsClick();
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
        }
    }
}