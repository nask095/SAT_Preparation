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
using System.Text.RegularExpressions;

namespace SAT_Prep
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
           

        }

        public int[] guessed;
        public int wordNum;
        public int words=0;
        Random rnd = new Random();
        Regex regexPattern = new Regex(@"[0-9]*");
        Regex regexPattern2 = new Regex(@"[^0-9]*");


        private void button1_Click(object sender, EventArgs e)
        {
            textBox1.Select();
            int temp;
            string test;
            string number="";
            label2.Text = "";
            label3.Text = "";
            foreach(Match m in regexPattern.Matches(label5.Text))
            {
                number = m.Groups[0].Value;
                break;
            }
         if(number + textBox1.Text==listBox2.Items[wordNum].ToString())
            {
                label2.Text = "Correct!";
                listBox1.Items.Remove(listBox1.Items[wordNum]);
                listBox2.Items.Remove(listBox2.Items[wordNum]);
                textBox1.Text = "";
                label7.Text = "Remaining: " + listBox1.Items.Count;
                if (listBox1.Items.Count>0)
                {
                    wordNum = rnd.Next(0, listBox1.Items.Count - 1);
                    label5.Text = listBox1.Items[wordNum].ToString();
                    test = listBox1.Items[wordNum].ToString();
                    do
                    {
                        test = test.Substring(1);
                        temp = Convert.ToInt32(test[0]);
                    } while (temp == 48 || temp == 49 || temp == 50 || temp == 51 || temp == 52 || temp == 53 || temp == 54 || temp == 55 || temp == 56 || temp == 57);
                    label1.Text = test;
                }
                else
                {
                    MessageBox.Show("There are no more words.");
                    textBox1.Text = "";
                    textBox2.Text = "";
                    label5.Text = "";
                    label1.Text = "";
                    label2.Text = "";
                    label3.Text = "";
                }
            }
            else
            {
                label2.Text="WRONG!!!";
                test = listBox2.Items[wordNum].ToString();
                do
                {
                    test = test.Substring(1);
                    temp = Convert.ToInt32(test[0]);
                } while (temp == 48 || temp == 49 || temp == 50 || temp == 51 || temp == 52 || temp == 53 || temp == 54 || temp == 55 || temp == 56 || temp == 57);
                label3.Text = test;
                wordNum = rnd.Next(0, listBox1.Items.Count - 1);
                label5.Text = listBox1.Items[wordNum].ToString();
                test = listBox1.Items[wordNum].ToString();
                do
                {
                    test = test.Substring(1);
                    temp = Convert.ToInt32(test[0]);
                } while (temp == 48 || temp == 49 || temp == 50 || temp == 51 || temp == 52 || temp == 53 || temp == 54 || temp == 55 || temp == 56 || temp == 57);
                label1.Text = test;
                textBox1.Text = "";

            } 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string startupPath = Environment.CurrentDirectory;
            int number=0;
            string word;
            Random rnd = new Random();
            listBox1.Items.Clear();
            listBox2.Items.Clear();
            Excel.Application xlApp;
            Excel.Workbook wb;
            Excel.Worksheet ws;
            string module = textBox2.Text;
            xlApp = new Excel.Application();      
            wb = xlApp.Workbooks.Open(startupPath+@"\Words.xlsx");
            ws = (Excel.Worksheet) wb.Worksheets.get_Item("Module_"+module);
            string test;
            int temp;
            int i = 1;
            while((word=ws.Cells[i,1].Value)!=null)
            {
                listBox1.Items.Add(number.ToString()+word);
                number++;
                i++;
            }
            number = 0;
            word = "";
            i = 1;
            while ((word = ws.Cells[i, 2].Value) != null)
            {
                listBox2.Items.Add(number.ToString()+word);
                number++;
                i++;
            }
            wb.Close();

            wordNum = rnd.Next(0, listBox1.Items.Count-1);
            label5.Text = listBox1.Items[wordNum].ToString();
            test = listBox1.Items[wordNum].ToString();
            do
            {
                test = test.Substring(1);
                temp = Convert.ToInt32(test[0]);
            } while (temp == 48 || temp == 49 || temp == 50 || temp == 51 || temp == 52 || temp == 53 || temp == 54 || temp == 55 || temp == 56 || temp == 57);
            label1.Text = test;
            label6.Visible = true;
            textBox1.Visible = true;
            button1.Visible = true;
            label7.Visible = true;
            label7.Text = "Remaining: " + listBox1.Items.Count;
            textBox1.Select();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Start s = new Start();
            s.Show();
            this.Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }
    }
}
