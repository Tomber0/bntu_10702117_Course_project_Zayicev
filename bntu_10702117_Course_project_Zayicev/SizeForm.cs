using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace bntu_10702117_Course_project_Zayicev
{
    public partial class SizeForm : Form
    {
        public SizeForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                settings.CubeX = float.Parse(textBox1.Text);

            }
            catch (Exception)
            {
                MessageBox.Show($"Недопустимое значение для {label2.Text}"); return;

            }
            try
            {
                settings.CubeY = float.Parse(textBox2.Text) ;

            }
            catch (Exception)
            {
                MessageBox.Show($"Недопустимое значение для {label3.Text}"); return;

            }
            try
            {
                settings.CubeZ = float.Parse(textBox3.Text) ;


            }
            catch (Exception)
            {
                MessageBox.Show($"Недопустимое значение для {label4.Text}"); return;

            }
            try
            {
                settings.BassR = float.Parse(textBox4.Text) ;

            }
            catch (Exception)
            {
                MessageBox.Show($"Недопустимое значение для {label6.Text} основания"); return;

            }
            try
            {
                settings.CutR = float.Parse(textBox5.Text) ;

            }
            catch (Exception)
            {
                MessageBox.Show($"Недопустимое значение для {label7.Text} выреза");
                return;
            }
            float[] list = { settings.BassR, settings.CubeX, settings.CubeY, settings.CubeZ, settings.CutR };
            foreach (var item in list)
            {
                if (item<=0)
                {
                    MessageBox.Show($"Недопустимое значение");
                    return;
                }
            }

            form1Ref.UpdateSettings(settings);
            Hide();

        }
        public Settings settings;
        public Form1 form1Ref;
        private void SizeForm_Load(object sender, EventArgs e)
        {

            textBox1.Text = settings.CubeX.ToString();
            textBox2.Text = settings.CubeY.ToString();
            textBox3.Text = settings.CubeZ.ToString();
            textBox4.Text = settings.BassR.ToString();
            textBox5.Text = settings.CutR.ToString();


        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Hide();
        }
    }
}
