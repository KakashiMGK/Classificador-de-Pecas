using System.Windows.Forms;

namespace Classificador_de_Peças
{
    partial class Form2
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form2));
            label1 = new Label();
            textBox1 = new TextBox();
            label2 = new Label();
            textBox2 = new TextBox();
            pictureBox1 = new PictureBox();
            button1 = new Button();
            button2 = new Button();
            imageList1 = new ImageList(components);
            picbxVerSenha = new PictureBox();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)picbxVerSenha).BeginInit();
            SuspendLayout();
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.None;
            label1.Location = new Point(63, 294);
            label1.Name = "label1";
            label1.Size = new Size(50, 15);
            label1.TabIndex = 0;
            label1.Text = "Usuário:";
            // 
            // textBox1
            // 
            textBox1.Anchor = AnchorStyles.None;
            textBox1.Location = new Point(119, 291);
            textBox1.Name = "textBox1";
            textBox1.Size = new Size(292, 23);
            textBox1.TabIndex = 1;
            textBox1.TextChanged += textBox1_TextChanged;
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.None;
            label2.Location = new Point(71, 340);
            label2.Name = "label2";
            label2.Size = new Size(42, 15);
            label2.TabIndex = 2;
            label2.Text = "Senha:";
            // 
            // textBox2
            // 
            textBox2.Anchor = AnchorStyles.None;
            textBox2.Location = new Point(118, 336);
            textBox2.Name = "textBox2";
            textBox2.Size = new Size(293, 23);
            textBox2.TabIndex = 3;
            textBox2.WordWrap = false;
            textBox2.TextChanged += textBox2_TextChanged;
            // 
            // pictureBox1
            // 
            pictureBox1.Anchor = AnchorStyles.None;
            pictureBox1.BackgroundImageLayout = ImageLayout.Center;
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.Location = new Point(85, -20);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(340, 334);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 4;
            pictureBox1.TabStop = false;
            pictureBox1.Click += pictureBox1_Click;
            // 
            // button1
            // 
            button1.Anchor = AnchorStyles.None;
            button1.Location = new Point(119, 378);
            button1.Name = "button1";
            button1.Size = new Size(125, 28);
            button1.TabIndex = 5;
            button1.Text = "Logar";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.None;
            button2.Location = new Point(286, 378);
            button2.Name = "button2";
            button2.Size = new Size(125, 28);
            button2.TabIndex = 6;
            button2.Text = "Sair";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click;
            // 
            // imageList1
            // 
            imageList1.ColorDepth = ColorDepth.Depth32Bit;
            imageList1.ImageStream = (ImageListStreamer)resources.GetObject("imageList1.ImageStream");
            imageList1.TransparentColor = Color.Transparent;
            imageList1.Images.SetKeyName(0, "Closed.png");
            imageList1.Images.SetKeyName(1, "Opened.png");
            // 
            // picbxVerSenha
            // 
            picbxVerSenha.Anchor = AnchorStyles.None;
            picbxVerSenha.Location = new Point(428, 336);
            picbxVerSenha.Name = "picbxVerSenha";
            picbxVerSenha.Size = new Size(27, 23);
            picbxVerSenha.TabIndex = 7;
            picbxVerSenha.TabStop = false;
            picbxVerSenha.Click += picbxVerSenha_Click;
            // 
            // Form2
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(559, 433);
            Controls.Add(picbxVerSenha);
            Controls.Add(button2);
            Controls.Add(button1);
            Controls.Add(textBox2);
            Controls.Add(label2);
            Controls.Add(textBox1);
            Controls.Add(label1);
            Controls.Add(pictureBox1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MaximizeBox = false;
            MinimumSize = new Size(571, 463);
            Name = "Form2";
            Text = "Classificador de Peças - Criado Por: Sergio Lucio de Oliveira Junior ";
            Load += Form2_Load;
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ((System.ComponentModel.ISupportInitialize)picbxVerSenha).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Label label1;
        private TextBox textBox1;
        private Label label2;
        private TextBox textBox2;
        private PictureBox pictureBox1;
        private Button button1;
        private Button button2;
        private ImageList imageList1;
        private PictureBox picbxVerSenha;
    }
}