using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Speech.Synthesis;

namespace SAT_Prep
{
    public partial class Translate : Form
    {
        public int[] guessed;
        public int wordNum;
        public int words = 0;
        Random rnd = new Random();
        string speech;
        SpeechSynthesizer synth = new SpeechSynthesizer();
        public Translate()
        {
            InitializeComponent();
            synth.SetOutputToDefaultAudioDevice();
        }

        
        

        private void button2_Click(object sender, EventArgs e)
        {
            string startupPath = Environment.CurrentDirectory;
            Random rnd = new Random();
            Excel.Application xlApp;
            Excel.Workbook wb;
            Excel.Worksheet ws;
            string module = textBox2.Text;
            xlApp = new Excel.Application();
            wb = xlApp.Workbooks.Open(startupPath + @"\Words.xlsx");
            ws = (Excel.Worksheet)wb.Worksheets.get_Item("Module_" + module);
            string word;
            int i = 1;
            while ((word = ws.Cells[i, 1].Value) != null)
            {
                listBox1.Items.Add(word);
                i++;
            }
            word = "";
            i = 1;
            while ((word = ws.Cells[i, 2].Value) != null)
            {
                listBox2.Items.Add(word);
                i++;
            }
            wb.Close();
            wordNum = rnd.Next(0, listBox1.Items.Count - 1);
            label5.Text = "Remaining: " + listBox1.Items.Count;
            speech = listBox1.Items[wordNum].ToString();
            textBox1.Select();
            synth.Speak(speech);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Start s = new Start();
            s.Show();
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Select();
            label2.Text = "";
            label3.Text = "";
            if (textBox1.Text == listBox1.Items[wordNum].ToString())
            {
                synth.Speak("Correct");
                listBox1.Items.Remove(listBox1.Items[wordNum]);
                listBox2.Items.Remove(listBox2.Items[wordNum]);
                textBox1.Text = "";
                label5.Text = "Remaining: " + listBox1.Items.Count;
                if (listBox1.Items.Count > 0)
                {
                    wordNum = rnd.Next(0, listBox1.Items.Count - 1);
                    speech = listBox1.Items[wordNum].ToString();
                    synth.Speak(speech);
                }
                else
                {
                    MessageBox.Show("There are no more words.");
                    textBox1.Text = "";
                    textBox2.Text = "";
                    label1.Text = "";
                    label2.Text = "";
                    label3.Text = "";
                }
            }
            else
            {
                synth.Speak("Wrong");
                label2.Text = "WRONG!!!";
                label3.Text = listBox1.Items[wordNum].ToString();
                wordNum = rnd.Next(0, listBox1.Items.Count - 1);
                speech = listBox1.Items[wordNum].ToString();
                synth.Speak(speech);
                textBox1.Text = "";

            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            synth.Speak(speech);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
            
        }
    }
}
