namespace Classificador_de_Peças
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            button1 = new Button();
            progressBar1 = new ProgressBar();
            chckbxNest15MM = new CheckBox();
            button4 = new Button();
            label1 = new Label();
            label2 = new Label();
            PainelSecundario = new RichTextBox();
            comboBox1 = new ComboBox();
            button5 = new Button();
            dataGridView1 = new DataGridView();
            ERP = new DataGridViewTextBoxColumn();
            RAZAOSOCIAL = new DataGridViewTextBoxColumn();
            ORDEMCOMPRA = new DataGridViewTextBoxColumn();
            PLANEJADOR = new DataGridViewTextBoxColumn();
            QNTD = new DataGridViewTextBoxColumn();
            ALTURA = new DataGridViewTextBoxColumn();
            LARGURA = new DataGridViewTextBoxColumn();
            ESPESSURA = new DataGridViewTextBoxColumn();
            CODIGOMATERIAL = new DataGridViewTextBoxColumn();
            DESCRICAOMATERIAL = new DataGridViewTextBoxColumn();
            LARGURACORTE = new DataGridViewTextBoxColumn();
            ALTURACORTE = new DataGridViewTextBoxColumn();
            IMAGEMMATERIAL = new DataGridViewTextBoxColumn();
            CODIGOPECA = new DataGridViewTextBoxColumn();
            COMPLEMENTO = new DataGridViewTextBoxColumn();
            DESCRICAOPECA = new DataGridViewTextBoxColumn();
            DESENHO1 = new DataGridViewTextBoxColumn();
            DESENHO2 = new DataGridViewTextBoxColumn();
            DESENHO3 = new DataGridViewTextBoxColumn();
            VEIOMATERIAL = new DataGridViewTextBoxColumn();
            BORDASUP = new DataGridViewTextBoxColumn();
            BORDAINF = new DataGridViewTextBoxColumn();
            BORDAESQ = new DataGridViewTextBoxColumn();
            BORDADIR = new DataGridViewTextBoxColumn();
            DESTINOIMPRESSAO = new DataGridViewTextBoxColumn();
            CODBARRAS = new DataGridViewTextBoxColumn();
            POSTOSOPERATIVOS = new DataGridViewTextBoxColumn();
            NUMLOTE = new DataGridViewTextBoxColumn();
            CODCLIENTE = new DataGridViewTextBoxColumn();
            MODULO = new DataGridViewTextBoxColumn();
            ORDEM = new DataGridViewTextBoxColumn();
            DATAENTREGA = new DataGridViewTextBoxColumn();
            BOX = new DataGridViewTextBoxColumn();
            ESPECIAL = new DataGridViewTextBoxColumn();
            checkBox1 = new CheckBox();
            timer1 = new System.Windows.Forms.Timer(components);
            label3 = new Label();
            pictureBox1 = new PictureBox();
            button2 = new Button();
            chckbxPinFolha = new CheckBox();
            chckbxRipa = new CheckBox();
            chckbx18MM = new CheckBox();
            chckbx25MM = new CheckBox();
            chckbxOutros = new CheckBox();
            chckbxImprimir = new CheckBox();
            chckbxExcluir = new CheckBox();
            chckbxNest12MM = new CheckBox();
            chckbxNest18MM = new CheckBox();
            chckbxNest25MM = new CheckBox();
            button3 = new Button();
            button6 = new Button();
            chckBoxEnvioEmail = new CheckBox();
            chckbxManterOrdens = new CheckBox();
            chckbxCorrecaoCodBarras = new CheckBox();
            chckbxSobreescrever = new CheckBox();
            btnAdmin = new Button();
            btnAbrirPerfil = new Button();
            lblAgMDF = new Label();
            lblAgPerfil = new Label();
            btnAtualizacoes = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
            SuspendLayout();
            // 
            // button1
            // 
            button1.Location = new Point(12, 6);
            button1.Name = "button1";
            button1.Size = new Size(119, 45);
            button1.TabIndex = 0;
            button1.Text = "Abrir CSV - MDF";
            button1.UseVisualStyleBackColor = true;
            button1.Click += button1_Click;
            // 
            // progressBar1
            // 
            progressBar1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            progressBar1.Location = new Point(12, 566);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(1309, 23);
            progressBar1.TabIndex = 3;
            progressBar1.Click += progressBar1_Click;
            // 
            // chckbxNest15MM
            // 
            chckbxNest15MM.AutoSize = true;
            chckbxNest15MM.Location = new Point(542, 40);
            chckbxNest15MM.Name = "chckbxNest15MM";
            chckbxNest15MM.Size = new Size(91, 19);
            chckbxNest15MM.TabIndex = 10;
            chckbxNest15MM.Text = "NEST 15MM";
            chckbxNest15MM.UseVisualStyleBackColor = true;
            chckbxNest15MM.CheckedChanged += chckbxNest15MM_CheckedChanged;
            // 
            // button4
            // 
            button4.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            button4.Location = new Point(12, 510);
            button4.Name = "button4";
            button4.Size = new Size(462, 50);
            button4.TabIndex = 13;
            button4.Text = "Iniciar";
            button4.UseVisualStyleBackColor = true;
            button4.Click += button4_Click;
            // 
            // label1
            // 
            label1.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            label1.AutoSize = true;
            label1.Location = new Point(498, 545);
            label1.Name = "label1";
            label1.Size = new Size(47, 15);
            label1.TabIndex = 14;
            label1.Text = "Tempo:";
            label1.Click += label1_Click;
            // 
            // label2
            // 
            label2.Anchor = AnchorStyles.Bottom | AnchorStyles.Left;
            label2.AutoSize = true;
            label2.Location = new Point(498, 510);
            label2.Name = "label2";
            label2.Size = new Size(42, 15);
            label2.TabIndex = 16;
            label2.Text = "Status:";
            label2.Click += label2_Click;
            // 
            // PainelSecundario
            // 
            PainelSecundario.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Right;
            PainelSecundario.BackColor = SystemColors.ControlLight;
            PainelSecundario.Location = new Point(939, 115);
            PainelSecundario.Name = "PainelSecundario";
            PainelSecundario.ScrollBars = RichTextBoxScrollBars.Vertical;
            PainelSecundario.Size = new Size(382, 335);
            PainelSecundario.TabIndex = 18;
            PainelSecundario.Text = "";
            PainelSecundario.TextChanged += richTextBox2_TextChanged;
            // 
            // comboBox1
            // 
            comboBox1.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            comboBox1.FormattingEnabled = true;
            comboBox1.Items.AddRange(new object[] { "01 - PINFOLHA", "02 - PINFOLHA_MO", "03 - RIPA", "04 - MO", "05 - PORTAFRENTE_MO", "06 - MATPRIMA", "07 - PORTAFRENTE", "08 - 18MM", "09 - 25MM", "10 - OUTROS", "11 - IMPRIMIR", "12 - MAOAMG", "13 - ESTRODAPE", "14 - TAPEÇARIA", "15 - EXCLUIR" });
            comboBox1.Location = new Point(939, 471);
            comboBox1.Name = "comboBox1";
            comboBox1.Size = new Size(271, 23);
            comboBox1.TabIndex = 19;
            comboBox1.SelectedIndexChanged += comboBox1_SelectedIndexChanged;
            // 
            // button5
            // 
            button5.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button5.Location = new Point(1216, 471);
            button5.Name = "button5";
            button5.Size = new Size(105, 23);
            button5.TabIndex = 20;
            button5.Text = "OK";
            button5.UseVisualStyleBackColor = true;
            button5.Click += button5_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Columns.AddRange(new DataGridViewColumn[] { ERP, RAZAOSOCIAL, ORDEMCOMPRA, PLANEJADOR, QNTD, ALTURA, LARGURA, ESPESSURA, CODIGOMATERIAL, DESCRICAOMATERIAL, LARGURACORTE, ALTURACORTE, IMAGEMMATERIAL, CODIGOPECA, COMPLEMENTO, DESCRICAOPECA, DESENHO1, DESENHO2, DESENHO3, VEIOMATERIAL, BORDASUP, BORDAINF, BORDAESQ, BORDADIR, DESTINOIMPRESSAO, CODBARRAS, POSTOSOPERATIVOS, NUMLOTE, CODCLIENTE, MODULO, ORDEM, DATAENTREGA, BOX, ESPECIAL });
            dataGridView1.Location = new Point(12, 115);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.ReadOnly = true;
            dataGridView1.Size = new Size(921, 379);
            dataGridView1.TabIndex = 21;
            // 
            // ERP
            // 
            ERP.HeaderText = "ERP";
            ERP.Name = "ERP";
            ERP.ReadOnly = true;
            // 
            // RAZAOSOCIAL
            // 
            RAZAOSOCIAL.HeaderText = "RAZÃO SOCIAL";
            RAZAOSOCIAL.Name = "RAZAOSOCIAL";
            RAZAOSOCIAL.ReadOnly = true;
            // 
            // ORDEMCOMPRA
            // 
            ORDEMCOMPRA.HeaderText = "ORDEM COMPRA";
            ORDEMCOMPRA.Name = "ORDEMCOMPRA";
            ORDEMCOMPRA.ReadOnly = true;
            // 
            // PLANEJADOR
            // 
            PLANEJADOR.HeaderText = "PLANEJADOR";
            PLANEJADOR.Name = "PLANEJADOR";
            PLANEJADOR.ReadOnly = true;
            // 
            // QNTD
            // 
            QNTD.HeaderText = "QNTD";
            QNTD.Name = "QNTD";
            QNTD.ReadOnly = true;
            // 
            // ALTURA
            // 
            ALTURA.HeaderText = "ALTURA";
            ALTURA.Name = "ALTURA";
            ALTURA.ReadOnly = true;
            // 
            // LARGURA
            // 
            LARGURA.HeaderText = "LARGURA";
            LARGURA.Name = "LARGURA";
            LARGURA.ReadOnly = true;
            // 
            // ESPESSURA
            // 
            ESPESSURA.HeaderText = "ESPESSURA";
            ESPESSURA.Name = "ESPESSURA";
            ESPESSURA.ReadOnly = true;
            // 
            // CODIGOMATERIAL
            // 
            CODIGOMATERIAL.HeaderText = "CODIGO MATERIAL";
            CODIGOMATERIAL.Name = "CODIGOMATERIAL";
            CODIGOMATERIAL.ReadOnly = true;
            // 
            // DESCRICAOMATERIAL
            // 
            DESCRICAOMATERIAL.HeaderText = "DESCRIÇÃO MATERIAL";
            DESCRICAOMATERIAL.Name = "DESCRICAOMATERIAL";
            DESCRICAOMATERIAL.ReadOnly = true;
            // 
            // LARGURACORTE
            // 
            LARGURACORTE.HeaderText = "LARGURA CORTE (NÃO UTILIZADA)";
            LARGURACORTE.Name = "LARGURACORTE";
            LARGURACORTE.ReadOnly = true;
            // 
            // ALTURACORTE
            // 
            ALTURACORTE.HeaderText = "ALTURA CORTE (NÃO UTILIZADA)";
            ALTURACORTE.Name = "ALTURACORTE";
            ALTURACORTE.ReadOnly = true;
            // 
            // IMAGEMMATERIAL
            // 
            IMAGEMMATERIAL.HeaderText = "IMAGEM MATERIAL (NÃO UTILIZADA)";
            IMAGEMMATERIAL.Name = "IMAGEMMATERIAL";
            IMAGEMMATERIAL.ReadOnly = true;
            // 
            // CODIGOPECA
            // 
            CODIGOPECA.HeaderText = "CÓDIGO PEÇA";
            CODIGOPECA.Name = "CODIGOPECA";
            CODIGOPECA.ReadOnly = true;
            // 
            // COMPLEMENTO
            // 
            COMPLEMENTO.HeaderText = "COMPLEMENTO (NÃO UTILIZADA)";
            COMPLEMENTO.Name = "COMPLEMENTO";
            COMPLEMENTO.ReadOnly = true;
            // 
            // DESCRICAOPECA
            // 
            DESCRICAOPECA.HeaderText = "DESCRIÇÃO PEÇA";
            DESCRICAOPECA.Name = "DESCRICAOPECA";
            DESCRICAOPECA.ReadOnly = true;
            // 
            // DESENHO1
            // 
            DESENHO1.HeaderText = "DESENHO 1";
            DESENHO1.Name = "DESENHO1";
            DESENHO1.ReadOnly = true;
            // 
            // DESENHO2
            // 
            DESENHO2.HeaderText = "DESENHO 2";
            DESENHO2.Name = "DESENHO2";
            DESENHO2.ReadOnly = true;
            // 
            // DESENHO3
            // 
            DESENHO3.HeaderText = "DESENHO 3";
            DESENHO3.Name = "DESENHO3";
            DESENHO3.ReadOnly = true;
            // 
            // VEIOMATERIAL
            // 
            VEIOMATERIAL.HeaderText = "VEIO MATERIAL (NÃO UTILIZADA)";
            VEIOMATERIAL.Name = "VEIOMATERIAL";
            VEIOMATERIAL.ReadOnly = true;
            // 
            // BORDASUP
            // 
            BORDASUP.HeaderText = "BORDA SUP";
            BORDASUP.Name = "BORDASUP";
            BORDASUP.ReadOnly = true;
            // 
            // BORDAINF
            // 
            BORDAINF.HeaderText = "BORDA INF";
            BORDAINF.Name = "BORDAINF";
            BORDAINF.ReadOnly = true;
            // 
            // BORDAESQ
            // 
            BORDAESQ.HeaderText = "BORDA ESQ";
            BORDAESQ.Name = "BORDAESQ";
            BORDAESQ.ReadOnly = true;
            // 
            // BORDADIR
            // 
            BORDADIR.HeaderText = "BORDA DIR";
            BORDADIR.Name = "BORDADIR";
            BORDADIR.ReadOnly = true;
            // 
            // DESTINOIMPRESSAO
            // 
            DESTINOIMPRESSAO.HeaderText = "DESTINO IMPRESSAO";
            DESTINOIMPRESSAO.Name = "DESTINOIMPRESSAO";
            DESTINOIMPRESSAO.ReadOnly = true;
            // 
            // CODBARRAS
            // 
            CODBARRAS.HeaderText = "CODIGO BARRAS";
            CODBARRAS.Name = "CODBARRAS";
            CODBARRAS.ReadOnly = true;
            // 
            // POSTOSOPERATIVOS
            // 
            POSTOSOPERATIVOS.HeaderText = "POSTOS OPERATIVOS";
            POSTOSOPERATIVOS.Name = "POSTOSOPERATIVOS";
            POSTOSOPERATIVOS.ReadOnly = true;
            // 
            // NUMLOTE
            // 
            NUMLOTE.HeaderText = "SEM - LOTE";
            NUMLOTE.Name = "NUMLOTE";
            NUMLOTE.ReadOnly = true;
            // 
            // CODCLIENTE
            // 
            CODCLIENTE.HeaderText = "CODIGO CLIENTE";
            CODCLIENTE.Name = "CODCLIENTE";
            CODCLIENTE.ReadOnly = true;
            // 
            // MODULO
            // 
            MODULO.HeaderText = "MÓDULO";
            MODULO.Name = "MODULO";
            MODULO.ReadOnly = true;
            // 
            // ORDEM
            // 
            ORDEM.HeaderText = "NÚM ORDEM";
            ORDEM.Name = "ORDEM";
            ORDEM.ReadOnly = true;
            // 
            // DATAENTREGA
            // 
            DATAENTREGA.HeaderText = "DATA ENTREGA";
            DATAENTREGA.Name = "DATAENTREGA";
            DATAENTREGA.ReadOnly = true;
            // 
            // BOX
            // 
            BOX.HeaderText = "BOX (NÃO UTILIZADA)";
            BOX.Name = "BOX";
            BOX.ReadOnly = true;
            // 
            // ESPECIAL
            // 
            ESPECIAL.HeaderText = "ESPECIAL";
            ESPECIAL.Name = "ESPECIAL";
            ESPECIAL.ReadOnly = true;
            // 
            // checkBox1
            // 
            checkBox1.Font = new Font("Segoe UI", 6.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
            checkBox1.Location = new Point(250, 6);
            checkBox1.Name = "checkBox1";
            checkBox1.Size = new Size(129, 35);
            checkBox1.TabIndex = 22;
            checkBox1.Text = "Único / Ripa / Excluir / Imprimir";
            checkBox1.UseVisualStyleBackColor = true;
            checkBox1.CheckedChanged += checkBox1_CheckedChanged;
            // 
            // timer1
            // 
            timer1.Enabled = true;
            timer1.Interval = 1;
            timer1.Tick += timer1_Tick;
            // 
            // label3
            // 
            label3.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            label3.AutoSize = true;
            label3.Location = new Point(939, 453);
            label3.Name = "label3";
            label3.Size = new Size(103, 15);
            label3.TabIndex = 23;
            label3.Text = "Selecione o plano:";
            // 
            // pictureBox1
            // 
            pictureBox1.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            pictureBox1.BackgroundImageLayout = ImageLayout.None;
            pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
            pictureBox1.Location = new Point(1190, -37);
            pictureBox1.Name = "pictureBox1";
            pictureBox1.Size = new Size(148, 144);
            pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
            pictureBox1.TabIndex = 24;
            pictureBox1.TabStop = false;
            // 
            // button2
            // 
            button2.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            button2.Location = new Point(1201, 510);
            button2.Name = "button2";
            button2.Size = new Size(120, 50);
            button2.TabIndex = 25;
            button2.Text = "Abrir Local dos Arquivos";
            button2.UseVisualStyleBackColor = true;
            button2.Click += button2_Click_1;
            // 
            // chckbxPinFolha
            // 
            chckbxPinFolha.AutoSize = true;
            chckbxPinFolha.Location = new Point(250, 40);
            chckbxPinFolha.Name = "chckbxPinFolha";
            chckbxPinFolha.Size = new Size(98, 19);
            chckbxPinFolha.TabIndex = 26;
            chckbxPinFolha.Text = "A - Pint Folha";
            chckbxPinFolha.UseVisualStyleBackColor = true;
            chckbxPinFolha.CheckedChanged += chckbxPinFolha_CheckedChanged;
            // 
            // chckbxRipa
            // 
            chckbxRipa.AutoSize = true;
            chckbxRipa.Location = new Point(250, 65);
            chckbxRipa.Name = "chckbxRipa";
            chckbxRipa.Size = new Size(67, 19);
            chckbxRipa.TabIndex = 27;
            chckbxRipa.Text = "B - Ripa";
            chckbxRipa.UseVisualStyleBackColor = true;
            chckbxRipa.CheckedChanged += chckbxRipa_CheckedChanged;
            // 
            // chckbx18MM
            // 
            chckbx18MM.AutoSize = true;
            chckbx18MM.Location = new Point(250, 90);
            chckbx18MM.Name = "chckbx18MM";
            chckbx18MM.Size = new Size(79, 19);
            chckbx18MM.TabIndex = 28;
            chckbx18MM.Text = "C - 18MM";
            chckbx18MM.UseVisualStyleBackColor = true;
            chckbx18MM.CheckedChanged += chckbx18MM_CheckedChanged;
            // 
            // chckbx25MM
            // 
            chckbx25MM.AutoSize = true;
            chckbx25MM.Location = new Point(398, 13);
            chckbx25MM.Name = "chckbx25MM";
            chckbx25MM.Size = new Size(79, 19);
            chckbx25MM.TabIndex = 29;
            chckbx25MM.Text = "D - 25MM";
            chckbx25MM.UseVisualStyleBackColor = true;
            chckbx25MM.CheckedChanged += chckbx25MM_CheckedChanged;
            // 
            // chckbxOutros
            // 
            chckbxOutros.AutoSize = true;
            chckbxOutros.Location = new Point(398, 40);
            chckbxOutros.Name = "chckbxOutros";
            chckbxOutros.Size = new Size(82, 19);
            chckbxOutros.TabIndex = 30;
            chckbxOutros.Text = "E - Outros ";
            chckbxOutros.UseVisualStyleBackColor = true;
            chckbxOutros.CheckedChanged += chckbxOutros_CheckedChanged;
            // 
            // chckbxImprimir
            // 
            chckbxImprimir.AutoSize = true;
            chckbxImprimir.Location = new Point(398, 65);
            chckbxImprimir.Name = "chckbxImprimir";
            chckbxImprimir.Size = new Size(89, 19);
            chckbxImprimir.TabIndex = 31;
            chckbxImprimir.Text = "F - Imprimir";
            chckbxImprimir.UseVisualStyleBackColor = true;
            chckbxImprimir.CheckedChanged += chckbxImprimir_CheckedChanged;
            // 
            // chckbxExcluir
            // 
            chckbxExcluir.AutoSize = true;
            chckbxExcluir.Location = new Point(398, 90);
            chckbxExcluir.Name = "chckbxExcluir";
            chckbxExcluir.Size = new Size(79, 19);
            chckbxExcluir.TabIndex = 32;
            chckbxExcluir.Text = "G - Excluir";
            chckbxExcluir.UseVisualStyleBackColor = true;
            chckbxExcluir.CheckedChanged += chckbxExcluir_CheckedChanged;
            // 
            // chckbxNest12MM
            // 
            chckbxNest12MM.AutoSize = true;
            chckbxNest12MM.Location = new Point(542, 13);
            chckbxNest12MM.Name = "chckbxNest12MM";
            chckbxNest12MM.Size = new Size(91, 19);
            chckbxNest12MM.TabIndex = 33;
            chckbxNest12MM.Text = "NEST 12MM";
            chckbxNest12MM.UseVisualStyleBackColor = true;
            chckbxNest12MM.CheckedChanged += chckbxNest12MM_CheckedChanged;
            // 
            // chckbxNest18MM
            // 
            chckbxNest18MM.AutoSize = true;
            chckbxNest18MM.Location = new Point(542, 65);
            chckbxNest18MM.Name = "chckbxNest18MM";
            chckbxNest18MM.Size = new Size(91, 19);
            chckbxNest18MM.TabIndex = 34;
            chckbxNest18MM.Text = "NEST 18MM";
            chckbxNest18MM.UseVisualStyleBackColor = true;
            chckbxNest18MM.CheckedChanged += chckbxNest18MM_CheckedChanged;
            // 
            // chckbxNest25MM
            // 
            chckbxNest25MM.AutoSize = true;
            chckbxNest25MM.Location = new Point(542, 90);
            chckbxNest25MM.Name = "chckbxNest25MM";
            chckbxNest25MM.Size = new Size(91, 19);
            chckbxNest25MM.TabIndex = 35;
            chckbxNest25MM.Text = "NEST 25MM";
            chckbxNest25MM.UseVisualStyleBackColor = true;
            chckbxNest25MM.CheckedChanged += chckbxNest25MM_CheckedChanged;
            // 
            // button3
            // 
            button3.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            button3.Location = new Point(725, 13);
            button3.Name = "button3";
            button3.Size = new Size(127, 45);
            button3.TabIndex = 36;
            button3.Text = "Marcar Todos";
            button3.UseVisualStyleBackColor = true;
            button3.Click += button3_Click;
            // 
            // button6
            // 
            button6.Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right;
            button6.Location = new Point(725, 64);
            button6.Name = "button6";
            button6.Size = new Size(127, 45);
            button6.TabIndex = 37;
            button6.Text = "Desmarcar Todos";
            button6.UseVisualStyleBackColor = true;
            button6.Click += button6_Click;
            // 
            // chckBoxEnvioEmail
            // 
            chckBoxEnvioEmail.AutoSize = true;
            chckBoxEnvioEmail.Enabled = false;
            chckBoxEnvioEmail.Location = new Point(939, 88);
            chckBoxEnvioEmail.Name = "chckBoxEnvioEmail";
            chckBoxEnvioEmail.Size = new Size(87, 19);
            chckBoxEnvioEmail.TabIndex = 38;
            chckBoxEnvioEmail.Text = "Envio Email";
            chckBoxEnvioEmail.UseVisualStyleBackColor = true;
            chckBoxEnvioEmail.CheckedChanged += chckBoxEnvioEmail_CheckedChanged;
            // 
            // chckbxManterOrdens
            // 
            chckbxManterOrdens.AutoSize = true;
            chckbxManterOrdens.Checked = true;
            chckbxManterOrdens.CheckState = CheckState.Checked;
            chckbxManterOrdens.Enabled = false;
            chckbxManterOrdens.Font = new Font("Segoe UI", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chckbxManterOrdens.Location = new Point(939, 13);
            chckbxManterOrdens.Name = "chckbxManterOrdens";
            chckbxManterOrdens.Size = new Size(117, 17);
            chckbxManterOrdens.TabIndex = 39;
            chckbxManterOrdens.Text = "TX_ManterOrdens";
            chckbxManterOrdens.UseVisualStyleBackColor = true;
            chckbxManterOrdens.CheckedChanged += chckbxManterOrdens_CheckedChanged;
            // 
            // chckbxCorrecaoCodBarras
            // 
            chckbxCorrecaoCodBarras.AutoSize = true;
            chckbxCorrecaoCodBarras.Checked = true;
            chckbxCorrecaoCodBarras.CheckState = CheckState.Checked;
            chckbxCorrecaoCodBarras.Enabled = false;
            chckbxCorrecaoCodBarras.Font = new Font("Segoe UI", 6.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
            chckbxCorrecaoCodBarras.Location = new Point(939, 40);
            chckbxCorrecaoCodBarras.Name = "chckbxCorrecaoCodBarras";
            chckbxCorrecaoCodBarras.Size = new Size(115, 16);
            chckbxCorrecaoCodBarras.TabIndex = 40;
            chckbxCorrecaoCodBarras.Text = "TX_CorreçãoCodBarras";
            chckbxCorrecaoCodBarras.UseVisualStyleBackColor = true;
            chckbxCorrecaoCodBarras.CheckedChanged += chckbxCorrecaoCodBarras_CheckedChanged;
            // 
            // chckbxSobreescrever
            // 
            chckbxSobreescrever.AutoSize = true;
            chckbxSobreescrever.Checked = true;
            chckbxSobreescrever.CheckState = CheckState.Checked;
            chckbxSobreescrever.Enabled = false;
            chckbxSobreescrever.Location = new Point(939, 63);
            chckbxSobreescrever.Name = "chckbxSobreescrever";
            chckbxSobreescrever.Size = new Size(215, 19);
            chckbxSobreescrever.TabIndex = 41;
            chckbxSobreescrever.Text = "Sobreescrever Arquivos de Máquina";
            chckbxSobreescrever.UseVisualStyleBackColor = true;
            chckbxSobreescrever.CheckedChanged += chckbxSobreescrever_CheckedChanged;
            // 
            // btnAdmin
            // 
            btnAdmin.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            btnAdmin.Location = new Point(830, 535);
            btnAdmin.Name = "btnAdmin";
            btnAdmin.Size = new Size(103, 25);
            btnAdmin.TabIndex = 42;
            btnAdmin.Text = "Administrador";
            btnAdmin.UseVisualStyleBackColor = true;
            btnAdmin.Click += btnAdmin_Click;
            // 
            // btnAbrirPerfil
            // 
            btnAbrirPerfil.Location = new Point(12, 62);
            btnAbrirPerfil.Name = "btnAbrirPerfil";
            btnAbrirPerfil.Size = new Size(119, 45);
            btnAbrirPerfil.TabIndex = 43;
            btnAbrirPerfil.Text = "Abrir CSV - Perfil";
            btnAbrirPerfil.UseVisualStyleBackColor = true;
            btnAbrirPerfil.Click += btnAbrirPerfil_Click;
            // 
            // lblAgMDF
            // 
            lblAgMDF.AutoSize = true;
            lblAgMDF.Location = new Point(137, 21);
            lblAgMDF.Name = "lblAgMDF";
            lblAgMDF.Size = new Size(82, 15);
            lblAgMDF.TabIndex = 44;
            lblAgMDF.Text = "Aguardando...";
            lblAgMDF.Click += lblAgMDF_Click;
            // 
            // lblAgPerfil
            // 
            lblAgPerfil.AutoSize = true;
            lblAgPerfil.Location = new Point(137, 77);
            lblAgPerfil.Name = "lblAgPerfil";
            lblAgPerfil.Size = new Size(82, 15);
            lblAgPerfil.TabIndex = 45;
            lblAgPerfil.Text = "Aguardando...";
            // 
            // btnAtualizacoes
            // 
            btnAtualizacoes.Location = new Point(1098, 3);
            btnAtualizacoes.Name = "btnAtualizacoes";
            btnAtualizacoes.Size = new Size(86, 27);
            btnAtualizacoes.TabIndex = 46;
            btnAtualizacoes.Text = "Atualizações";
            btnAtualizacoes.UseVisualStyleBackColor = true;
            btnAtualizacoes.Click += btnAtualizacoes_Click;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            BackColor = SystemColors.Control;
            ClientSize = new Size(1333, 598);
            Controls.Add(btnAtualizacoes);
            Controls.Add(lblAgPerfil);
            Controls.Add(lblAgMDF);
            Controls.Add(btnAbrirPerfil);
            Controls.Add(btnAdmin);
            Controls.Add(chckbxSobreescrever);
            Controls.Add(chckbxCorrecaoCodBarras);
            Controls.Add(chckbxManterOrdens);
            Controls.Add(chckBoxEnvioEmail);
            Controls.Add(button6);
            Controls.Add(button3);
            Controls.Add(chckbxNest25MM);
            Controls.Add(chckbxNest18MM);
            Controls.Add(chckbxNest12MM);
            Controls.Add(chckbxExcluir);
            Controls.Add(chckbxImprimir);
            Controls.Add(chckbxOutros);
            Controls.Add(chckbx25MM);
            Controls.Add(chckbx18MM);
            Controls.Add(chckbxRipa);
            Controls.Add(chckbxPinFolha);
            Controls.Add(button2);
            Controls.Add(PainelSecundario);
            Controls.Add(dataGridView1);
            Controls.Add(pictureBox1);
            Controls.Add(label3);
            Controls.Add(checkBox1);
            Controls.Add(button5);
            Controls.Add(comboBox1);
            Controls.Add(label2);
            Controls.Add(label1);
            Controls.Add(button4);
            Controls.Add(chckbxNest15MM);
            Controls.Add(progressBar1);
            Controls.Add(button1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            MinimumSize = new Size(1349, 637);
            Name = "Form1";
            Text = "[2.1] Classificador de Peças - Criado Por: Sergio Lucio de Oliveira Junior ";
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            ((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button button1;
        private ProgressBar progressBar1;
        private CheckBox chckbxNest15MM;
        private Button button4;
        private Label label1;
        private Label label2;
        private RichTextBox PainelSecundario;
        private ComboBox comboBox1;
        private Button button5;
        private DataGridView dataGridView1;
        private DataGridViewTextBoxColumn ERP;
        private DataGridViewTextBoxColumn RAZAOSOCIAL;
        private DataGridViewTextBoxColumn ORDEMCOMPRA;
        private DataGridViewTextBoxColumn PLANEJADOR;
        private DataGridViewTextBoxColumn QNTD;
        private DataGridViewTextBoxColumn ALTURA;
        private DataGridViewTextBoxColumn LARGURA;
        private DataGridViewTextBoxColumn ESPESSURA;
        private DataGridViewTextBoxColumn CODIGOMATERIAL;
        private DataGridViewTextBoxColumn DESCRICAOMATERIAL;
        private DataGridViewTextBoxColumn LARGURACORTE;
        private DataGridViewTextBoxColumn ALTURACORTE;
        private DataGridViewTextBoxColumn IMAGEMMATERIAL;
        private DataGridViewTextBoxColumn CODIGOPECA;
        private DataGridViewTextBoxColumn COMPLEMENTO;
        private DataGridViewTextBoxColumn DESCRICAOPECA;
        private DataGridViewTextBoxColumn DESENHO1;
        private DataGridViewTextBoxColumn DESENHO2;
        private DataGridViewTextBoxColumn DESENHO3;
        private DataGridViewTextBoxColumn VEIOMATERIAL;
        private DataGridViewTextBoxColumn BORDASUP;
        private DataGridViewTextBoxColumn BORDAINF;
        private DataGridViewTextBoxColumn BORDAESQ;
        private DataGridViewTextBoxColumn BORDADIR;
        private DataGridViewTextBoxColumn DESTINOIMPRESSAO;
        private DataGridViewTextBoxColumn CODBARRAS;
        private DataGridViewTextBoxColumn POSTOSOPERATIVOS;
        private DataGridViewTextBoxColumn NUMLOTE;
        private DataGridViewTextBoxColumn CODCLIENTE;
        private DataGridViewTextBoxColumn MODULO;
        private DataGridViewTextBoxColumn ORDEM;
        private DataGridViewTextBoxColumn DATAENTREGA;
        private DataGridViewTextBoxColumn BOX;
        private DataGridViewTextBoxColumn ESPECIAL;
        private CheckBox checkBox1;
        private System.Windows.Forms.Timer timer1;
        private Label label3;
        private PictureBox pictureBox1;
        private Button button2;
        private CheckBox chckbxPinFolha;
        private CheckBox chckbxRipa;
        private CheckBox chckbx18MM;
        private CheckBox chckbx25MM;
        private CheckBox chckbxOutros;
        private CheckBox chckbxImprimir;
        private CheckBox chckbxExcluir;
        private CheckBox chckbxNest12MM;
        private CheckBox chckbxNest18MM;
        private CheckBox chckbxNest25MM;
        private Button button3;
        private Button button6;
        private CheckBox chckBoxEnvioEmail;
        private CheckBox chckbxManterOrdens;
        private CheckBox chckbxCorrecaoCodBarras;
        private CheckBox chckbxSobreescrever;
        private Button btnAdmin;
        private Button btnAbrirPerfil;
        private Label lblAgMDF;
        private Label lblAgPerfil;
        private Button btnAtualizacoes;
    }
}
