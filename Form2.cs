using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
namespace Classificador_de_Peças
{
    public partial class Form2 : Form
    {
        public string usuario = "PCP";
        public string senha = "admin";

        private Form1 _form1;

        public Form2(Form1 form1)
        {
            InitializeComponent();
            _form1 = form1;

            FormClosing += Form2_FormClosing;

            ToolTip tip = new ToolTip();
            tip.SetToolTip(picbxVerSenha, "Segure para ver a senha");

            picbxVerSenha.Image = Properties.Resources.Closed;
            picbxVerSenha.SizeMode = PictureBoxSizeMode.StretchImage;

            picbxVerSenha.MouseDown += picbxVerSenha_MouseDown;
            picbxVerSenha.MouseUp += picbxVerSenha_MouseUp;
            picbxVerSenha.MouseLeave += picbxVerSenha_MouseLeave;

            textBox2.UseSystemPasswordChar = true;
            picbxVerSenha.Image = Properties.Resources.Closed;
            
        }


        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void picbxVerSenha_MouseDown(object sender, MouseEventArgs e)
        {
            textBox2.UseSystemPasswordChar = false;
            picbxVerSenha.Image = Properties.Resources.Opened;
        }

        private void picbxVerSenha_MouseUp(object sender, MouseEventArgs e)
        {
            textBox2.UseSystemPasswordChar = true;
            picbxVerSenha.Image = Properties.Resources.Closed;
        }

        private void picbxVerSenha_MouseLeave(object sender, EventArgs e)
        {
            textBox2.UseSystemPasswordChar = true;
            picbxVerSenha.Image = Properties.Resources.Closed;
        }


        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == usuario && textBox2.Text == senha)
            {
                this.Hide();
                _form1.ModoAdmin();
                _form1.Show();

            }
            else
            {
                MessageBox.Show("Usuário ou senha inválidos.");
            }

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }
        private void button2_Click(object sender, EventArgs e)
        {

            this.Hide();

            _form1.Show();
        }
        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {

            Application.Exit();
        }
        private bool visivel = false;


        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void picbxVerSenha_Click(object sender, EventArgs e)
        {

        }
    }
}
