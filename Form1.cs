using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic;
using System;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Net.NetworkInformation;
using System.Reflection.Metadata;
using System.Security.Cryptography;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Excel = ClosedXML;

#pragma warning disable CA1845


namespace Classificador_de_Peças
{



    public partial class Form1 : Form
    {

        public string codigoUltimaLeitura = "";
        public int qntdpecasgrid;
        public string versao = "[2.11]";

        public string caminhoPadraoInfoPecas = @"J:\PCP\InfoPecasPlanos\";
        //public string caminhoPadraoInfoPecas = @"C:\Users\sergi\Desktop\PCP\InfoPecasPlanos\";

        // Para fazer a contagem de peças de cada plano gerado em csv
        public int emailqntdPLPinFolha = 0;
        public int emailqntdPLRipa = 0;
        public int emailqntdPL18mm = 0;
        public int emailqntdPL25mm = 0;
        public int emailqntdPLOutros = 0;
        public int emailqntdPLImprimir = 0;
        public int emailqntdPLExcluir = 0;
        public int emailqntdPLNesting12mm = 0;
        public int emailqntdPLNesting15mm = 0;
        public int emailqntdPLNesting18mm = 0;
        public int emailqntdPLNesting25mm = 0;
        public int emailqntdMatNaoExiste = 0;
        public int emailqntdBruto = 0;
        public int emailqntdpecasCodigoBarrasErrado = 0;


        // Para fazer a contagem de peças dentro das listas de relatorios de peças
        public int emailqntdPecasLote = 0;
        public int emailqntdCMPLote = 0;
        public int emailqntdCOLote = 0;
        public int emailqntdBOLote = 0;
        public int emailqntdFULote = 0;
        public int emailqntdUSILote = 0;
        public int emailqntdCFLote = 0;
        public int emailqntdMOLote = 0;
        public int emailqntdMALote = 0;
        public int emailqntdMPLote = 0;
        public int emailqntdMELote = 0;
        public int emailqntdPINLote = 0;
        public int emailqntdFLLote = 0;
        public int emailqntdColagemLote = 0;
        public int emailqntdCQLote = 0;
        public int emailqntdFiletacaoLote = 0;

        // Para fazer a contagem de peças de cada plano gerado em csv
        public int PREqntdPLPinFolha = 0;
        public int PREqntdPLRipa = 0;
        public int PREqntdPL18mm = 0;
        public int PREqntdPL25mm = 0;
        public int PREqntdPLOutros = 0;
        public int PREqntdPLImprimir = 0;
        public int PREqntdPLExcluir = 0;
        public int PREqntdPLNesting12mm = 0;
        public int PREqntdPLNesting15mm = 0;
        public int PREqntdPLNesting18mm = 0;
        public int PREqntdPLNesting25mm = 0;

        public Stopwatch stopwatch = new();


        public Form1()
        {
            InitializeComponent();
            

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            stopwatch = new Stopwatch();
            this.Text = $"{versao} Classificador de Peças - Criado Por: Sergio Lucio de Oliveira Junior";
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        public string caminhocolado2 = "";
        public int valorProgresso = 0;
        public int primeiroValorProgresso = 0;
        private object httpClient;
        public bool bloquearFecha;
        public string caminhocsvPerfil;
        public bool csvPerfilCarregado;

        private void button1_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            PararProgresso();

            OpenFileDialog ofd = new OpenFileDialog();
            //ofd.InitialDirectory = @"J:\PCP\PCP 2025\LOTE - PROD";

            ofd.Filter = "CSV files (*.csv)|*.csv";

            // Optionally, set the default extension (e.g., if no filter is selected)
            ofd.DefaultExt = ".csv";
            DialogResult buscaArquivo = ofd.ShowDialog();

            ofd.RestoreDirectory = false;

            this.Text = $"{CultureInfo.CurrentCulture.Name} | " +
            $"{versao} Classificador de Peças - Criado Por: Sergio Lucio de Oliveira Junior - Arquivo: " + ofd.FileName;
            if (buscaArquivo == DialogResult.OK)
            {

                caminhocolado2 = ofd.FileName;
                try
                {
                    var lotecru = File.ReadAllLines(caminhocolado2).Skip(1);
                    qntdpecasgrid = 0;

                    foreach (var linha in lotecru)
                    {
                        Armazenado armazenar = linha;

                        if(string.IsNullOrWhiteSpace(armazenar.Altura) || string.IsNullOrWhiteSpace(armazenar.Largura)) // Se altura ou largura estiverem vazias ou forem apenas espaços em branco, exibe uma mensagem de erro indicando o código da peça e a ordem, e para a execução do código para evitar erros posteriores
                        {
                            MessageBox.Show($"A peça com Código: {armazenar.CodigoPeca} e Ordem:{armazenar.NumeroOrdem}  possui Altura ou Largura vazia. Verifique o arquivo CSV.", "Erro de Dimensões", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ResetaPrograma();
                            return; // Para o codigo
                        }
                        if (int.TryParse(armazenar.CodigoMaterial, out _) == false) // Tenta fazer a coversão para inteiro do código de material, se não conseguir é porque tem letra ou caractere especial, o que é um código de matéria-prima inválido
                        {
                            MessageBox.Show($"A peça com Código: {armazenar.CodigoPeca} e Ordem:{armazenar.NumeroOrdem} possui código de matéria-prima inválida.", "Erro de Matéria-Prima", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ResetaPrograma();
                            return; // Para o codigo
                        }
                        if (armazenar.Planejador.Length > 5)
                        {
                            if (armazenar.Planejador.Contains("PERFIL"))
                            {
                                MessageBox.Show($"A peça com Código: {armazenar.CodigoPeca} e Ordem:{armazenar.NumeroOrdem} possui Planejador de PERFIL. Verifique o arquivo CSV.", "Erro de Planejador", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                ResetaPrograma();
                                return; // Para o codigo
                            }
                        }
                           

                        // Normaliza a string, substituindo vírgula por ponto e removendo espaços
                        string alturaStr = armazenar.Altura.Trim().Replace(",", ".");
                        string larguraStr = armazenar.Largura.Trim().Replace(",", ".");

                        // Converte para decimal usando InvariantCulture
                        decimal altura = decimal.Parse(alturaStr, CultureInfo.InvariantCulture);
                        decimal largura = decimal.Parse(larguraStr, CultureInfo.InvariantCulture);

                        dataGridView1.Rows.Add(armazenar.ERP, armazenar.RazaoSocial, armazenar.PedidoAmbiente, armazenar.Planejador, armazenar.Quantidade, altura, largura, armazenar.Espessura, armazenar.CodigoMaterial, armazenar.DescricaoMaterial, armazenar.LarguraCorte, armazenar.AlturaCorte, armazenar.ImagemMaterial, armazenar.CodigoPeca, armazenar.Complemento, armazenar.DescricaoPeca, armazenar.DesenhoUm, armazenar.DesenhoDois, armazenar.DesenhoTres, armazenar.VeioMaterial, armazenar.BordaSup, armazenar.BordaInf, armazenar.BordaEsq, armazenar.BordaDir, armazenar.DestinoImpressao, armazenar.CodigoBarras, armazenar.PostosOperativos, armazenar.NumeroLote, armazenar.CodigoCliente, armazenar.Modulo, armazenar.NumeroOrdem, armazenar.DataEntrega, armazenar.Plano, armazenar.Especial);
                        qntdpecasgrid++;
                    }
                    dataGridView1.Rows.Add("", "", "", "Qntd Total Peças: ", qntdpecasgrid);

                    PreStart(caminhocolado2);

                    lblAgMDF.Text = "MDF Carregado.";
                    label2.Text = "Status: CSV de MDF Carregado.";



                    if (System.IO.Path.GetFileName(caminhocolado2).Contains("RETFAB", StringComparison.OrdinalIgnoreCase) && checkBox1.Checked == false)
                    {
                        DialogResult askRETFAB = MessageBox.Show("Esse csv aparenta ser RETFAB, deseja selecionar a opção de UNICO CSV?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (askRETFAB == DialogResult.Yes)
                        {
                            checkBox1.Checked = true;
                        }
                    }
                    if (System.IO.Path.GetFileName(caminhocolado2).Contains("ASM", StringComparison.OrdinalIgnoreCase) == true && checkBox1.Checked == false)
                    {
                        DialogResult askRETFAB = MessageBox.Show("Esse csv aparenta ser ASM, deseja selecionar a opção de UNICO CSV?", "Confirmação", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (askRETFAB == DialogResult.Yes)
                        {
                            checkBox1.Checked = true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erro ao ler o arquivo: " + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            if (buscaArquivo == DialogResult.Cancel)
            {
                MessageBox.Show("Cancelado a busca pelo usuário.");
            }
            //MessageBox.Show(pacific.ToString()); // Para Testes

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private async void button4_Click(object sender, EventArgs e)
        {

            if (!File.Exists(caminhocolado2))
            {
                MessageBox.Show("Nenhum arquivo .csv selecionado.", "Abra um CSV", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (checkBox1.Checked == false &&
                chckbxPinFolha.Checked == false &&
                chckbxRipa.Checked == false &&
                chckbx18MM.Checked == false &&
                chckbx25MM.Checked == false &&
                chckbxOutros.Checked == false &&
                chckbxNest15MM.Checked == false &&
                chckbxImprimir.Checked == false &&
                chckbxExcluir.Checked == false &&
                chckbxNest12MM.Checked == false &&
                chckbxNest18MM.Checked == false &&
                chckbxNest25MM.Checked == false)
            {
                MessageBox.Show("Nenhuma Flag selecionada. Selecione o tipo de separação que deseja.", "Selecionar plano de separação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            stopwatch.Reset();
            stopwatch.Start();
            var lotecru = File.ReadAllLines(caminhocolado2).Skip(1);




            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nLendo os arquivos .txt com informações dos códigos.");
            PainelSecundario.ScrollToCaret();

            var arqPL_PINFOLHA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHA.txt");         // 01
            var arqPL_PINFOLHAMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHAMO.txt");       // 01
            var arqPL_RIPA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_RIPA.txt");             // 02
            var arqPL_TAPECARIA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_TAPECARIA.txt");        // 03
            var arqPL_MO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MO.txt");               // 04
            var arqPL_PORTAFRENTEMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTEMO.txt");    // 04
            var arqPL_MATPRIMA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MATPRIMA.txt");         // 05
            var arqPL_PORTAFRENTE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTE.txt");      // 06
            var arqPL18MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_18MM.txt");             // 07
            var arqPL25MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_25MM.txt");             // 08
            var arqPLOUTROS = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_OUTROS.txt");           // 09
            var arqPLIMPRIMIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_IMPRIMIR.txt");         // 10
            var arqPLMAOAMG = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MAOAMG.txt");           // --
            var arqPLESTRODAPE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_ESTRODAPE.txt");        // --
            var arqPLEXCLUIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_EXCLUIR.txt");          // --


            // LEITURA DO ARQUIVO DE BUSCA POR QUAL SEMANA 
            var arqDATASEMANA = File.ReadAllLines(caminhoPadraoInfoPecas + "DATASEMANA.txt");         // TXT COM AS DATAS DAS SEMANA


            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nCriando as listas de informações dos códigos.");
            PainelSecundario.ScrollToCaret();

            List<string> plpinfolhatxtList = new List<string>() { };  // 01
            List<string> plpinfolhamotxtList = new List<string>() { };  // 01
            List<string> plripatxtList = new List<string>() { };  // 02
            List<string> pltapecariatxtList = new List<string>() { };  // 03
            List<string> plmotxtList = new List<string>() { };  // 04
            List<string> plportafrentemotxtList = new List<string>() { };  // 04
            List<string> plmatprimatxtList = new List<string>() { };  // 05
            List<string> plportafrentetxtList = new List<string>() { };  // 06
            List<string> pl18mmtxtList = new List<string>() { };  // 07
            List<string> pl25mmtxtList = new List<string>() { };  // 08
            List<string> ploutrostxtList = new List<string>() { };  // 09    
            List<string> plimprimirtxtList = new List<string>() { };  // 10
            List<string> plmaoamgtxtList = new List<string>() { };  // --
            List<string> plestrodapetxtList = new List<string>() { };  // --
            List<string> plexcluirtxtList = new List<string>() { };  // --

            List<string> datasemanatxtList = new List<string>() { };  // -- 


            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nInserindo itens do .txt para listas.");
            PainelSecundario.ScrollToCaret();

            // 01 - PL_PINFOLHA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPINFOLHA in arqPL_PINFOLHA)
                plpinfolhatxtList.Add(itemPLPINFOLHA);

            // 01 - PL_PINFOLHAMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPINFOLHAMO in arqPL_PINFOLHAMO)
                plpinfolhamotxtList.Add(itemPLPINFOLHAMO);

            // 02 - PL_RIPA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL_RIPA in arqPL_RIPA)
                plripatxtList.Add(itemPL_RIPA);

            // 03 - PL_TAPECARIA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLTAPECARIA in arqPL_TAPECARIA)
                pltapecariatxtList.Add(itemPLTAPECARIA);

            // 04 - PL_MO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMO in arqPL_MO)
                plmotxtList.Add(itemPLMO);

            // 04 - PL_PORTAFRENTEMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPORTAFRENTEMO in arqPL_PORTAFRENTEMO)
                plportafrentemotxtList.Add(itemPLPORTAFRENTEMO);

            // 05 - PL_MATPRIMA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMATPRIMA in arqPL_MATPRIMA)
                plmatprimatxtList.Add(itemPLMATPRIMA);

            // 06 - PL_PORTAFRENTE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPORTAFRENTE in arqPL_PORTAFRENTE)
                plportafrentetxtList.Add(itemPLPORTAFRENTE);

            // 07 - PL_18MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL18MM in arqPL18MM)
                pl18mmtxtList.Add(itemPL18MM);

            // 08 - PL_25MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL25MM in arqPL25MM)
                pl25mmtxtList.Add(itemPL25MM);

            // 09 - PL_OUTROS - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLOUTROS in arqPLOUTROS)
                ploutrostxtList.Add(itemPLOUTROS);

            // 10 - PL_IMPRIMIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLIMPRIMIR in arqPLIMPRIMIR)
                plimprimirtxtList.Add(itemPLIMPRIMIR);

            // -- - PL_MAOAMG - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMAOAMG in arqPLMAOAMG)
                plmaoamgtxtList.Add(itemPLMAOAMG);

            // -- - PL_ESTRODAPE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLESTRODAPE in arqPLESTRODAPE)
                plestrodapetxtList.Add(itemPLESTRODAPE);

            // -- - PL_EXCLUIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLEXCLUIR in arqPLEXCLUIR)
                plexcluirtxtList.Add(itemPLEXCLUIR);


            // DATA SEMANAS - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemDATASEMANA in arqDATASEMANA)
                datasemanatxtList.Add(itemDATASEMANA);

            // LISTA PARA INCLUIR OS ITENS SEM INFO
            List<string> itemSemInfoList = new List<string>() { };
            var itensSemInfoList = new List<string> { }; // ITENS SEM INFORMAÇÃO DE TIPO DE PLANO

            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nContando as peças sem informação de plano.");
            PainelSecundario.ScrollToCaret();
            var todasAsListas = new HashSet<string>(
                    plpinfolhatxtList
                    .Concat(plpinfolhamotxtList)
                    .Concat(plripatxtList)
                    .Concat(pltapecariatxtList)
                    .Concat(plmotxtList)
                    .Concat(plportafrentemotxtList)
                    .Concat(plmatprimatxtList)
                    .Concat(plportafrentetxtList)
                    .Concat(pl18mmtxtList)
                    .Concat(pl25mmtxtList)
                    .Concat(ploutrostxtList)
                    .Concat(plimprimirtxtList)
                    .Concat(plmaoamgtxtList)
                    .Concat(plestrodapetxtList)
                    .Concat(plexcluirtxtList)
                );
            foreach (var linha in lotecru)
            {
                Armazenado armazenar = linha;
                string codigoBase;

                int idx = armazenar.CodigoPeca.IndexOf('.');

                if (idx >= 0)
                {
                    codigoBase = $"{armazenar.CodigoPeca[..idx]}.{armazenar.Espessura}";
                }
                else
                {
                    codigoBase = $"{armazenar.CodigoPeca}.{armazenar.Espessura}";
                }
                

                if (!todasAsListas.Contains(codigoBase))
                {
                    codigoBase = armazenar.CodigoPeca?.Split('.', 2)[0];

                    if (!string.IsNullOrEmpty(codigoBase))
                    {
                        itensSemInfoList.Add($"{codigoBase}.{armazenar.Espessura}");
                    }
                }

            }
            if (itensSemInfoList.Count() > 0)
            {

                PecasSemInfo(caminhocolado2);
            }
            else
            {
                itensSemInfoList.Clear();
                codigoUltimaLeitura = "";
                await Start(caminhocolado2);
            }
        }
        public void PreStart(string caminhocolado2)
        {

            var arqPL_PINFOLHA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHA.txt");         // 01
            var arqPL_PINFOLHAMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHAMO.txt");       // 01
            var arqPL_RIPA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_RIPA.txt");             // 02
            var arqPL_TAPECARIA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_TAPECARIA.txt");        // 03
            var arqPL_MO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MO.txt");               // 04
            var arqPL_PORTAFRENTEMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTEMO.txt");    // 04
            var arqPL_MATPRIMA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MATPRIMA.txt");         // 05
            var arqPL_PORTAFRENTE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTE.txt");      // 06
            var arqPL18MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_18MM.txt");             // 07
            var arqPL25MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_25MM.txt");             // 08
            var arqPLOUTROS = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_OUTROS.txt");           // 09
            var arqPLIMPRIMIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_IMPRIMIR.txt");         // 10
            var arqPLMAOAMG = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MAOAMG.txt");           // --
            var arqPLESTRODAPE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_ESTRODAPE.txt");        // --
            var arqPLEXCLUIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_EXCLUIR.txt");          // --


            List<string> plpinfolhatxtList = new List<string>() { };  // 01
            List<string> plpinfolhamotxtList = new List<string>() { };  // 01
            List<string> plripatxtList = new List<string>() { };  // 02
            List<string> pltapecariatxtList = new List<string>() { };  // 03
            List<string> plmotxtList = new List<string>() { };  // 04
            List<string> plportafrentemotxtList = new List<string>() { };  // 04
            List<string> plmatprimatxtList = new List<string>() { };  // 05
            List<string> plportafrentetxtList = new List<string>() { };  // 06
            List<string> pl18mmtxtList = new List<string>() { };  // 07
            List<string> pl25mmtxtList = new List<string>() { };  // 08
            List<string> ploutrostxtList = new List<string>() { };  // 09    
            List<string> plimprimirtxtList = new List<string>() { };  // 10
            List<string> plmaoamgtxtList = new List<string>() { };  // --
            List<string> plestrodapetxtList = new List<string>() { };  // --
            List<string> plexcluirtxtList = new List<string>() { };  // --

            // 01 - PL_PINFOLHA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPINFOLHA in arqPL_PINFOLHA)
                plpinfolhatxtList.Add(itemPLPINFOLHA);

            // 01 - PL_PINFOLHAMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPINFOLHAMO in arqPL_PINFOLHAMO)
                plpinfolhamotxtList.Add(itemPLPINFOLHAMO);

            // 02 - PL_RIPA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL_RIPA in arqPL_RIPA)
                plripatxtList.Add(itemPL_RIPA);

            // 03 - PL_TAPECARIA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLTAPECARIA in arqPL_TAPECARIA)
                pltapecariatxtList.Add(itemPLTAPECARIA);

            // 04 - PL_MO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMO in arqPL_MO)
                plmotxtList.Add(itemPLMO);

            // 04 - PL_PORTAFRENTEMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPORTAFRENTEMO in arqPL_PORTAFRENTEMO)
                plportafrentemotxtList.Add(itemPLPORTAFRENTEMO);

            // 05 - PL_MATPRIMA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMATPRIMA in arqPL_MATPRIMA)
                plmatprimatxtList.Add(itemPLMATPRIMA);

            // 06 - PL_PORTAFRENTE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPORTAFRENTE in arqPL_PORTAFRENTE)
                plportafrentetxtList.Add(itemPLPORTAFRENTE);

            // 07 - PL_18MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL18MM in arqPL18MM)
                pl18mmtxtList.Add(itemPL18MM);

            // 08 - PL_25MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL25MM in arqPL25MM)
                pl25mmtxtList.Add(itemPL25MM);

            // 09 - PL_OUTROS - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLOUTROS in arqPLOUTROS)
                ploutrostxtList.Add(itemPLOUTROS);

            // 10 - PL_IMPRIMIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLIMPRIMIR in arqPLIMPRIMIR)
                plimprimirtxtList.Add(itemPLIMPRIMIR);

            // -- - PL_MAOAMG - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMAOAMG in arqPLMAOAMG)
                plmaoamgtxtList.Add(itemPLMAOAMG);

            // -- - PL_ESTRODAPE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLESTRODAPE in arqPLESTRODAPE)
                plestrodapetxtList.Add(itemPLESTRODAPE);

            // -- - PL_EXCLUIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLEXCLUIR in arqPLEXCLUIR)
                plexcluirtxtList.Add(itemPLEXCLUIR);

            try
            {

                var loteCru = File.ReadLines(caminhocolado2).Skip(1);

                // AS LISTAS PARA A EXPLOSAO DO CSV PARA OTIMIZACAO
                var plpinfolhaList = new List<Armazenado> { }; // PLANO A
                var plripaList = new List<Armazenado> { }; // PLANO B
                var pl18List = new List<Armazenado> { }; // PLANO C
                var pl25List = new List<Armazenado> { }; // PLANO D
                var ploutrosList = new List<Armazenado> { }; // PLANO E
                var plimprimirList = new List<Armazenado> { }; // PLANO F
                var plexcluirList = new List<Armazenado> { }; // PLANO G
                var plnesting12mmList = new List<Armazenado> { }; // PLANO NESTING 12MM
                var plnesting15mmList = new List<Armazenado> { }; // PLANO NESTING 15MM
                var plnesting18mmList = new List<Armazenado> { }; // PLANO NESTING 18MM
                var plnesting25mmList = new List<Armazenado> { }; // PLANO NESTING 25MM

                var listaAmbientesFolha = new List<string> { };
                var listaPinturaFolha = new List<string> { };


                // Materias-Primas que não podem ir para o Nesting
                List<string> matNaoNesting = new() {
                "6267" , "12267", "18267", "25267", "6268" , "12268", "18268",
                "25268", "6269" , "12269", "18269", "25269", "6273" , "12273",
                "18273", "25273", "6274" , "12274", "18274", "25274", "6275" ,
                "12275", "18275", "25275", "6270" , "12270", "18270", "25270",
                "6260" , "12260", "18260", "25260", "6271" , "12271", "18271",
                "25271",
                };

                // Materias-Primas que não existe, criar a lista para observar e alertar.
                List<string> matNaoExiste = new() {
                "12192", "12049", "12154", "12161", "12255", "12273", "12274", "12270",
                "12260", "12271", "12254", "6990" , "12990", "18990", "25990",
                "37990", "51990", "6991" , "12991", "18991", "25991", "37991",
                "51991", "6992" , "12992", "18992", "25992", "37992", "51992"
                };

                List<string> matTouch = new() {
                "25270", "25260", "25271", "25273", "25274", "25275"
                };



                foreach (var linha in loteCru) // Primeira Leitura do CSV
                {
                    Armazenado armazenar = linha;
                    if (armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL"))
                        listaPinturaFolha.Add(armazenar.CodigoMaterial);

                }

                foreach (var linha in loteCru)
                {
                    Armazenado armazenar = linha;

                    armazenar.Altura = armazenar.Altura.Replace(",", ".");
                    armazenar.Largura = armazenar.Largura.Replace(",", ".");

                    //Console.WriteLine(armazenar.CodigoBarras);
                    // CONSERTA AS ESPESSURAS NA COLUNA DE ACORDO COM O CODIGO DO MATERIAL
                    if (armazenar.CodigoMaterial.Length == 5)
                        armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 2);
                    if (armazenar.CodigoMaterial.Length == 4)
                        armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 1);

                    // ALTERACAO DOS CAMPOS DA COLUNA DE NUMERO DO LOTE
                    if (!armazenar.PostosOperativos.Contains("PIN") || !armazenar.PostosOperativos.Contains("FL"))
                    {
                        if (!plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            // SE HOUVER SILK NAVAL DE 25MM ALTERAR PARA BRANCO TX NAVAL O QUE FOR PINTURA
                            if (armazenar.CodigoMaterial == "25277")
                            {
                                armazenar.CodigoMaterial = "25049";
                                armazenar.DescricaoMaterial = "CHAPA 25MM - BRANCO TX NAVAL";

                            }
                            // SE HOUVER BRANCO TX NAVAL DE 12MM ALTERAR PARA BRANCO TX NORMAL 12MM
                            if (armazenar.CodigoMaterial == "12049")
                            {
                                armazenar.CodigoMaterial = "12036";
                                armazenar.DescricaoMaterial = "CHAPA 12MM - BRANCO TX";
                            }
                        }
                    }
                    // até aqui para remover o não compensa e classificar pro nesting mesmo assim -------

                    //REGRAS PARA POSTOS OPERATIVOS
                    // PREENCHIMENTO ALUMINIO QUE NÃO CONTEM MO
                    if (armazenar.DescricaoPeca.Contains("PREENCH") && armazenar.DescricaoPeca.Contains("ALUM") && !armazenar.PostosOperativos.Contains("MO"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MO";
                    // ESTOFAD QUE NAO TEM TP
                    if (armazenar.DescricaoPeca.Contains("ESTOFA") && !armazenar.PostosOperativos.Contains("TP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-TP";
                    // BIT QUE NAO TEM MA
                    if (armazenar.DescricaoPeca.Contains("BIT") && !armazenar.PostosOperativos.Contains("MA"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MA";
                    // FRENTE DISPENSER SEM MP
                    if (armazenar.DescricaoPeca.Contains("DISPENSER") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    // BASE CALCEIRO SEM MP
                    if (armazenar.DescricaoPeca.Contains("BASE CALCEIRO") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";

                    // ALTERACAO DOS CAMPOS DA COLUNA DE NUMERO DO LOTE
                    if (armazenar.PostosOperativos.Contains("PIN"))
                    {
                        if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-02";
                        }
                        else
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-01";
                            // SE HOUVER SILK NAVAL DE 25MM ALTERAR PARA BRANCO TX NAVAL O QUE FOR PINTURA
                            if (armazenar.CodigoMaterial == "25277")
                            {
                                armazenar.CodigoMaterial = "25049";
                                armazenar.DescricaoMaterial = "CHAPA 25MM - BRANCO TX NAVAL";
                                // using (var logtxt = new StreamWriter("C:/Users/IDM/Desktop/725_relacao/logtxt.txt"))
                                // {
                                //     logtxt.Write("Alteração de peça de SILK NAVAL DE 25MM PARA BRANCO TX NAVAL 25MM (PEÇA PINTURA)");
                                // }

                            }
                            if (armazenar.CodigoMaterial == "12049")
                            {
                                armazenar.CodigoMaterial = "12036";
                                armazenar.DescricaoMaterial = "CHAPA 12MM - BRANCO TX";
                                // using (var logtxt = new StreamWriter("C:/Users/IDM/Desktop/725_relacao/logtxt.txt"))
                                // {
                                //     //logtxt.Write("Alteração de peça de SILK NAVAL DE 25MM PARA BRANCO TX NAVAL 25MM (PEÇA PINTURA)");
                                // }

                            }
                        }
                    }
                    else // SE HOUVER SILK NAVAL DE 25MM ALTERAR PARA SILK 25MM
                    {
                        if (armazenar.CodigoMaterial == "25277")
                        {
                            armazenar.CodigoMaterial = "25244";
                            armazenar.DescricaoMaterial = "SILK 25MM";
                            // using (var logtxt = new StreamWriter("C:/Users/IDM/Desktop/725_relacao/logtxt.txt"))
                            // {
                            //     logtxt.Write("Alteração de peça de SILK NAVAL DE 25MM PARA SILK 25MM");
                            // }
                        }
                        if (armazenar.CodigoMaterial == "12049")
                        {
                            armazenar.CodigoMaterial = "12036";
                            armazenar.DescricaoMaterial = "CHAPA 12MM - BRANCO TX";
                            // using (var logtxt = new StreamWriter("C:/Users/IDM/Desktop/725_relacao/logtxt.txt"))
                            // {
                            //     logtxt.Write("Alteração de peça de SILK NAVAL DE 25MM PARA BRANCO TX NAVAL 25MM (PEÇA PINTURA)");
                            // }

                        }
                    }
                    bool condicoesUnicaNesting = false;


                    // Normaliza a string, substituindo vírgula por ponto e removendo espaços
                    string alturaStr = armazenar.Altura.Trim().Replace(",", ".");
                    string larguraStr = armazenar.Largura.Trim().Replace(",", ".");

                    // Converte para decimal usando InvariantCulture
                    decimal altura = decimal.Parse(alturaStr, CultureInfo.InvariantCulture);
                    decimal largura = decimal.Parse(larguraStr, CultureInfo.InvariantCulture);

                    if (
                        !matTouch.Contains(armazenar.CodigoMaterial) &&
                        !matNaoNesting.Contains(armazenar.CodigoMaterial) &&
                        !matNaoExiste.Contains(armazenar.CodigoMaterial) &&
                        Math.Min(altura, largura) > 149 &&
                        !armazenar.DescricaoPeca.Contains("BARROTE"))
                    {
                        if (armazenar.Espessura is "18" or "12" or "25" or "15")
                            condicoesUnicaNesting = true;
                    }



                    if (armazenar.DescricaoPeca.Substring(0, 9) == "MAO AMIGA" && altura < 451)
                    {
                        PREqntdPLExcluir++;
                    }
                    else if (plexcluirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                    {
                        PREqntdPLExcluir++;
                    }
                    else if (condicoesUnicaNesting == true && armazenar.Espessura == "12")
                    {
                        PREqntdPLNesting12mm++;
                    }
                    else if (condicoesUnicaNesting == true && armazenar.Espessura == "18")
                    {
                        PREqntdPLNesting18mm++;
                    }
                    else if (condicoesUnicaNesting == true && armazenar.Espessura == "25")
                    {
                        PREqntdPLNesting25mm++;
                    }
                    else if (condicoesUnicaNesting == true && armazenar.Espessura == "15")
                    {
                        PREqntdPLNesting15mm++;
                    }
                    else if (armazenar.Espessura == "18")
                    {
                        if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            PREqntdPLRipa++;
                        }
                        else if (armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL") || listaPinturaFolha.Contains(armazenar.CodigoMaterial))
                        {
                            PREqntdPLPinFolha++;
                        }
                        else
                        {
                            PREqntdPL18mm++;
                        }
                    }
                    else if (armazenar.Espessura == "25" && !matTouch.Contains(armazenar.CodigoMaterial))
                    {
                        if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            PREqntdPLRipa++;
                        }
                        else if (armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL") || listaPinturaFolha.Contains(armazenar.CodigoMaterial))
                        {
                            PREqntdPLPinFolha++;
                        }
                        else
                        {
                            PREqntdPL25mm++;
                        }
                    }
                    else if (armazenar.Espessura == "6" || armazenar.Espessura == "12" || armazenar.Espessura == "3" || armazenar.Espessura == "15")
                    {
                        if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            PREqntdPLRipa++;
                        }
                        else if (armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL") || listaPinturaFolha.Contains(armazenar.CodigoMaterial))
                        {
                            PREqntdPLPinFolha++;
                        }
                        else
                        {
                            PREqntdPLOutros++;
                        }
                    }
                    else if (armazenar.Espessura == "37" || armazenar.Espessura == "51" || matTouch.Contains(armazenar.CodigoMaterial))
                    {
                        PREqntdPLImprimir++;
                    }

                }
                chckbxPinFolha.Text = "A - Pint Folha: " + PREqntdPLPinFolha;
                chckbxRipa.Text = "B - Ripa: " + PREqntdPLRipa;
                chckbx18MM.Text = "C - 18MM : " + PREqntdPL18mm;
                chckbx25MM.Text = "D - 25MM: " + PREqntdPL25mm;
                chckbxOutros.Text = "E - Outros: " + PREqntdPLOutros;
                chckbxImprimir.Text = "F - Imprimir: " + PREqntdPLImprimir;
                chckbxExcluir.Text = "G - Excluir: " + PREqntdPLExcluir;
                chckbxNest12MM.Text = "NEST 12MM: " + PREqntdPLNesting12mm;
                chckbxNest18MM.Text = "NEST 18MM: " + PREqntdPLNesting18mm;
                chckbxNest25MM.Text = "NEST 25MM: " + PREqntdPLNesting25mm;



            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro ao fazer o Pré Start:" + ex);
                PararProgresso();
            }

        }
        private static readonly object _logLock = new object();

        public void PararProgresso()
        {
            stopwatch.Stop();
            stopwatch.Reset();
            progressBar1.Value = 0;
        }

        public async Task Start(string caminhocolado2) // BOTÃO INICIAR
        {

            if (caminhocsvPerfil == null)
            {
                // Abre uma caixa de diálogo para o usuário escolher se ele quiser que continue sem selecionar um arquivo
                if (DialogResult.No == MessageBox.Show("Nenhum CSV de perfil selecionado. Deseja continuar assim mesmo?", "Arquivo de Perfil Não Selecionado", MessageBoxButtons.YesNo, MessageBoxIcon.Warning))
                {
                    PararProgresso();
                    return; // retorna para o início do método, não executando o restante do código 

                }
            }
            // Caso o usuário selecione um arquivo de perfil ou decide continuar mesmo sem o arquivo selecionado, o processo continua normalmente 
            label2.Text = "Status: Processo de classificação iniciado.";

            progressBar1.Maximum = qntdpecasgrid * 3 + 14;
            valorProgresso = 0;

            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nLendo os arquivos .txt com informações dos códigos.");
            PainelSecundario.ScrollToCaret();

            var arqPL_PINFOLHA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHA.txt");         // 01
            var arqPL_PINFOLHAMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHAMO.txt");       // 01
            var arqPL_RIPA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_RIPA.txt");             // 02
            var arqPL_TAPECARIA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_TAPECARIA.txt");        // 03
            var arqPL_MO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MO.txt");               // 04
            var arqPL_PORTAFRENTEMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTEMO.txt");    // 04
            var arqPL_MATPRIMA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MATPRIMA.txt");         // 05
            var arqPL_PORTAFRENTE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTE.txt");      // 06
            var arqPL18MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_18MM.txt");             // 07
            var arqPL25MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_25MM.txt");             // 08
            var arqPLOUTROS = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_OUTROS.txt");           // 09
            var arqPLIMPRIMIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_IMPRIMIR.txt");         // 10
            var arqPLMAOAMG = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MAOAMG.txt");           // --
            var arqPLESTRODAPE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_ESTRODAPE.txt");        // --
            var arqPLEXCLUIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_EXCLUIR.txt");          // --




            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nCriando as listas de informações dos códigos.");
            PainelSecundario.ScrollToCaret();

            List<string> plpinfolhatxtList = new List<string>() { };  // 01
            List<string> plpinfolhamotxtList = new List<string>() { };  // 01
            List<string> plripatxtList = new List<string>() { };  // 02
            List<string> pltapecariatxtList = new List<string>() { };  // 03
            List<string> plmotxtList = new List<string>() { };  // 04
            List<string> plportafrentemotxtList = new List<string>() { };  // 04
            List<string> plmatprimatxtList = new List<string>() { };  // 05
            List<string> plportafrentetxtList = new List<string>() { };  // 06
            List<string> pl18mmtxtList = new List<string>() { };  // 07
            List<string> pl25mmtxtList = new List<string>() { };  // 08
            List<string> ploutrostxtList = new List<string>() { };  // 09    
            List<string> plimprimirtxtList = new List<string>() { };  // 10
            List<string> plmaoamgtxtList = new List<string>() { };  // --
            List<string> plestrodapetxtList = new List<string>() { };  // --
            List<string> plexcluirtxtList = new List<string>() { };  // --

            List<int> pecasInativar = new List<int>() { }; // Listas das peças queserão inativadas,
            List<long> pecasManterOrdem = new List<long>() { }; // Listas das peças que NAO serão inativadas

            List<Ordem> pecasEnvioMes = new();

            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nInserindo itens do .txt para listas.");
            PainelSecundario.ScrollToCaret();

            // 01 - PL_PINFOLHA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPINFOLHA in arqPL_PINFOLHA)
                plpinfolhatxtList.Add(itemPLPINFOLHA);

            // 01 - PL_PINFOLHAMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPINFOLHAMO in arqPL_PINFOLHAMO)
                plpinfolhamotxtList.Add(itemPLPINFOLHAMO);

            // 02 - PL_RIPA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL_RIPA in arqPL_RIPA)
                plripatxtList.Add(itemPL_RIPA);

            // 03 - PL_TAPECARIA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLTAPECARIA in arqPL_TAPECARIA)
                pltapecariatxtList.Add(itemPLTAPECARIA);

            // 04 - PL_MO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMO in arqPL_MO)
                plmotxtList.Add(itemPLMO);

            // 04 - PL_PORTAFRENTEMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPORTAFRENTEMO in arqPL_PORTAFRENTEMO)
                plportafrentemotxtList.Add(itemPLPORTAFRENTEMO);

            // 05 - PL_MATPRIMA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMATPRIMA in arqPL_MATPRIMA)
                plmatprimatxtList.Add(itemPLMATPRIMA);

            // 06 - PL_PORTAFRENTE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPORTAFRENTE in arqPL_PORTAFRENTE)
                plportafrentetxtList.Add(itemPLPORTAFRENTE);

            // 07 - PL_18MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL18MM in arqPL18MM)
                pl18mmtxtList.Add(itemPL18MM);

            // 08 - PL_25MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL25MM in arqPL25MM)
                pl25mmtxtList.Add(itemPL25MM);

            // 09 - PL_OUTROS - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLOUTROS in arqPLOUTROS)
                ploutrostxtList.Add(itemPLOUTROS);

            // 10 - PL_IMPRIMIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLIMPRIMIR in arqPLIMPRIMIR)
                plimprimirtxtList.Add(itemPLIMPRIMIR);

            // -- - PL_MAOAMG - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMAOAMG in arqPLMAOAMG)
                plmaoamgtxtList.Add(itemPLMAOAMG);

            // -- - PL_ESTRODAPE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLESTRODAPE in arqPLESTRODAPE)
                plestrodapetxtList.Add(itemPLESTRODAPE);

            // -- - PL_EXCLUIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLEXCLUIR in arqPLEXCLUIR)
                plexcluirtxtList.Add(itemPLEXCLUIR);


            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nCriandos as colunas da tabela da Matriz de Lotes.");
            PainelSecundario.ScrollToCaret();

            DataTable tabelaMatrizLotes = new DataTable();
            tabelaMatrizLotes.Columns.Add("DATA PROCESSO");
            tabelaMatrizLotes.Columns.Add("NOME ARQUIVO");
            tabelaMatrizLotes.Columns.Add("SEMANA");
            tabelaMatrizLotes.Columns.Add("LOTE");
            tabelaMatrizLotes.Columns.Add("QNTD PEÇAS");
            tabelaMatrizLotes.Columns.Add("CMP");
            tabelaMatrizLotes.Columns.Add("PORCENT");
            tabelaMatrizLotes.Columns.Add("CO");
            tabelaMatrizLotes.Columns.Add("BO");
            tabelaMatrizLotes.Columns.Add("FU");
            tabelaMatrizLotes.Columns.Add("USI");
            tabelaMatrizLotes.Columns.Add("CF");
            tabelaMatrizLotes.Columns.Add("MO");
            tabelaMatrizLotes.Columns.Add("MA");
            tabelaMatrizLotes.Columns.Add("MP");
            tabelaMatrizLotes.Columns.Add("ME");
            tabelaMatrizLotes.Columns.Add("PIN");
            tabelaMatrizLotes.Columns.Add("FL");
            tabelaMatrizLotes.Columns.Add("COLAGEM");
            tabelaMatrizLotes.Columns.Add("CQ");
            tabelaMatrizLotes.Columns.Add("FILETACAO");


            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nCriandos as colunas da tabela de Matérias Primas não existentes.");
            PainelSecundario.ScrollToCaret();

            DataTable tabelaMatPrimaNaoExiste = new DataTable();
            tabelaMatPrimaNaoExiste.Columns.Add("NÚMERO DO PEDIDO");
            tabelaMatPrimaNaoExiste.Columns.Add("RAZãO SOCIAL CLIENTE");
            tabelaMatPrimaNaoExiste.Columns.Add("ORDEM COMPRA PEDIDO");
            tabelaMatPrimaNaoExiste.Columns.Add("NOME PLANEJADOR");
            tabelaMatPrimaNaoExiste.Columns.Add("QUANTIDADE");
            tabelaMatPrimaNaoExiste.Columns.Add("ALTURA PEÇA");
            tabelaMatPrimaNaoExiste.Columns.Add("LARGURA PEÇA");
            tabelaMatPrimaNaoExiste.Columns.Add("ESPESSURA PEÇA");
            tabelaMatPrimaNaoExiste.Columns.Add("CÓDIGO MATERIAL");
            tabelaMatPrimaNaoExiste.Columns.Add("DESCRIÇãO MATERIAL");
            tabelaMatPrimaNaoExiste.Columns.Add("LARGURA CORTE MATERIAL");
            tabelaMatPrimaNaoExiste.Columns.Add("ALTURA CORTE MATERIAL");
            tabelaMatPrimaNaoExiste.Columns.Add("IMAGEM DO MATERIAL");
            tabelaMatPrimaNaoExiste.Columns.Add("CÓDIGO PEÇA");
            tabelaMatPrimaNaoExiste.Columns.Add("COMPLEMENTO");
            tabelaMatPrimaNaoExiste.Columns.Add("DESCRIÇãO PEÇA");
            tabelaMatPrimaNaoExiste.Columns.Add("DESENHO PROGRAMADO 1");
            tabelaMatPrimaNaoExiste.Columns.Add("DESENHO PROGRAMADO 2");
            tabelaMatPrimaNaoExiste.Columns.Add("DESENHO PROGRAMADO 3");
            tabelaMatPrimaNaoExiste.Columns.Add("VEIO MATERIAL");
            tabelaMatPrimaNaoExiste.Columns.Add("BORDA SUPERIOR");
            tabelaMatPrimaNaoExiste.Columns.Add("BORDA INFERIOR");
            tabelaMatPrimaNaoExiste.Columns.Add("BORDA ESQUERDA");
            tabelaMatPrimaNaoExiste.Columns.Add("BORDA DIREITA");
            tabelaMatPrimaNaoExiste.Columns.Add("DESTINO IMPRESSãO");
            tabelaMatPrimaNaoExiste.Columns.Add("ID");
            tabelaMatPrimaNaoExiste.Columns.Add("POSTOS OPERATIVOS");
            tabelaMatPrimaNaoExiste.Columns.Add("NÚMERO DO LOTE");
            tabelaMatPrimaNaoExiste.Columns.Add("CÓDIGO CLIENTE");
            tabelaMatPrimaNaoExiste.Columns.Add("MODULO_ID");
            tabelaMatPrimaNaoExiste.Columns.Add("NÚMERO DA ORDEM");
            tabelaMatPrimaNaoExiste.Columns.Add("DATA ENTREGA LOTE");
            tabelaMatPrimaNaoExiste.Columns.Add("BOX");
            tabelaMatPrimaNaoExiste.Columns.Add("ESPECIAL");

            DataTable tabelaDXF = new DataTable();
            tabelaDXF.Columns.Add("NÚMERO DO PEDIDO");
            tabelaDXF.Columns.Add("RAZãO SOCIAL CLIENTE");
            tabelaDXF.Columns.Add("ORDEM COMPRA PEDIDO");
            tabelaDXF.Columns.Add("NOME PLANEJADOR");
            tabelaDXF.Columns.Add("QUANTIDADE");
            tabelaDXF.Columns.Add("ALTURA PEÇA");
            tabelaDXF.Columns.Add("LARGURA PEÇA");
            tabelaDXF.Columns.Add("ESPESSURA PEÇA");
            tabelaDXF.Columns.Add("CÓDIGO MATERIAL");
            tabelaDXF.Columns.Add("DESCRIÇãO MATERIAL");
            tabelaDXF.Columns.Add("LARGURA CORTE MATERIAL");
            tabelaDXF.Columns.Add("ALTURA CORTE MATERIAL");
            tabelaDXF.Columns.Add("IMAGEM DO MATERIAL");
            tabelaDXF.Columns.Add("CÓDIGO PEÇA");
            tabelaDXF.Columns.Add("COMPLEMENTO");
            tabelaDXF.Columns.Add("DESCRIÇãO PEÇA");
            tabelaDXF.Columns.Add("DESENHO PROGRAMADO 1");
            tabelaDXF.Columns.Add("DESENHO PROGRAMADO 2");
            tabelaDXF.Columns.Add("DESENHO PROGRAMADO 3");
            tabelaDXF.Columns.Add("VEIO MATERIAL");
            tabelaDXF.Columns.Add("BORDA SUPERIOR");
            tabelaDXF.Columns.Add("BORDA INFERIOR");
            tabelaDXF.Columns.Add("BORDA ESQUERDA");
            tabelaDXF.Columns.Add("BORDA DIREITA");
            tabelaDXF.Columns.Add("DESTINO IMPRESSãO");
            tabelaDXF.Columns.Add("ID");
            tabelaDXF.Columns.Add("POSTOS OPERATIVOS");
            tabelaDXF.Columns.Add("NÚMERO DO LOTE");
            tabelaDXF.Columns.Add("CÓDIGO CLIENTE");
            tabelaDXF.Columns.Add("MODULO_ID");
            tabelaDXF.Columns.Add("NÚMERO DA ORDEM");
            tabelaDXF.Columns.Add("DATA ENTREGA LOTE");
            tabelaDXF.Columns.Add("BOX");
            tabelaDXF.Columns.Add("ESPECIAL");
            int contpecasDXF = 0;

            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nCriandos as colunas da tabela Bruto.");
            PainelSecundario.ScrollToCaret();

            DataTable tabelaBruto = new DataTable();
            tabelaBruto.Columns.Add("NÚMERO DO PEDIDO");
            tabelaBruto.Columns.Add("RAZãO SOCIAL CLIENTE");
            tabelaBruto.Columns.Add("ORDEM COMPRA PEDIDO");
            tabelaBruto.Columns.Add("NOME PLANEJADOR");
            tabelaBruto.Columns.Add("QUANTIDADE");
            tabelaBruto.Columns.Add("ALTURA PEÇA");
            tabelaBruto.Columns.Add("LARGURA PEÇA");
            tabelaBruto.Columns.Add("ESPESSURA PEÇA");
            tabelaBruto.Columns.Add("CÓDIGO MATERIAL");
            tabelaBruto.Columns.Add("DESCRIÇãO MATERIAL");
            tabelaBruto.Columns.Add("LARGURA CORTE MATERIAL");
            tabelaBruto.Columns.Add("ALTURA CORTE MATERIAL");
            tabelaBruto.Columns.Add("IMAGEM DO MATERIAL");
            tabelaBruto.Columns.Add("CÓDIGO PEÇA");
            tabelaBruto.Columns.Add("COMPLEMENTO");
            tabelaBruto.Columns.Add("DESCRIÇãO PEÇA");
            tabelaBruto.Columns.Add("DESENHO PROGRAMADO 1");
            tabelaBruto.Columns.Add("DESENHO PROGRAMADO 2");
            tabelaBruto.Columns.Add("DESENHO PROGRAMADO 3");
            tabelaBruto.Columns.Add("VEIO MATERIAL");
            tabelaBruto.Columns.Add("BORDA SUPERIOR");
            tabelaBruto.Columns.Add("BORDA INFERIOR");
            tabelaBruto.Columns.Add("BORDA ESQUERDA");
            tabelaBruto.Columns.Add("BORDA DIREITA");
            tabelaBruto.Columns.Add("DESTINO IMPRESSãO");
            tabelaBruto.Columns.Add("ID");
            tabelaBruto.Columns.Add("POSTOS OPERATIVOS");
            tabelaBruto.Columns.Add("NÚMERO DO LOTE");
            tabelaBruto.Columns.Add("CÓDIGO CLIENTE");
            tabelaBruto.Columns.Add("MODULO_ID");
            tabelaBruto.Columns.Add("NÚMERO DA ORDEM");
            tabelaBruto.Columns.Add("DATA ENTREGA LOTE");
            tabelaBruto.Columns.Add("BOX");
            tabelaBruto.Columns.Add("ESPECIAL");

            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nCriandos as colunas da tabela de Relação de Peças.");
            PainelSecundario.ScrollToCaret();

            // TABELA DOS RELATORIOS DE PEÇAS
            DataTable tabelaRelacaoPecas = new DataTable();
            tabelaRelacaoPecas.Columns.Add("NÚMERO DO PEDIDO");
            tabelaRelacaoPecas.Columns.Add("RAZãO SOCIAL CLIENTE");
            tabelaRelacaoPecas.Columns.Add("ORDEM COMPRA PEDIDO");
            tabelaRelacaoPecas.Columns.Add("NOME PLANEJADOR");
            tabelaRelacaoPecas.Columns.Add("QUANTIDADE");
            tabelaRelacaoPecas.Columns.Add("ALTURA PEÇA");
            tabelaRelacaoPecas.Columns.Add("LARGURA PEÇA");
            tabelaRelacaoPecas.Columns.Add("ESPESSURA PEÇA");
            tabelaRelacaoPecas.Columns.Add("CÓDIGO MATERIAL");
            tabelaRelacaoPecas.Columns.Add("DESCRIÇãO MATERIAL");
            tabelaRelacaoPecas.Columns.Add("LARGURA CORTE MATERIAL");
            tabelaRelacaoPecas.Columns.Add("ALTURA CORTE MATERIAL");
            tabelaRelacaoPecas.Columns.Add("IMAGEM DO MATERIAL");
            tabelaRelacaoPecas.Columns.Add("CÓDIGO PEÇA");
            tabelaRelacaoPecas.Columns.Add("COMPLEMENTO");
            tabelaRelacaoPecas.Columns.Add("DESCRIÇãO PEÇA");
            tabelaRelacaoPecas.Columns.Add("DESENHO PROGRAMADO 1");
            tabelaRelacaoPecas.Columns.Add("DESENHO PROGRAMADO 2");
            tabelaRelacaoPecas.Columns.Add("DESENHO PROGRAMADO 3");
            tabelaRelacaoPecas.Columns.Add("VEIO MATERIAL");
            tabelaRelacaoPecas.Columns.Add("BORDA SUPERIOR");
            tabelaRelacaoPecas.Columns.Add("BORDA INFERIOR");
            tabelaRelacaoPecas.Columns.Add("BORDA ESQUERDA");
            tabelaRelacaoPecas.Columns.Add("BORDA DIREITA");
            tabelaRelacaoPecas.Columns.Add("DESTINO IMPRESSãO");
            tabelaRelacaoPecas.Columns.Add("ID");
            tabelaRelacaoPecas.Columns.Add("POSTOS OPERATIVOS");
            tabelaRelacaoPecas.Columns.Add("NÚMERO DO LOTE");
            tabelaRelacaoPecas.Columns.Add("CÓDIGO CLIENTE");
            tabelaRelacaoPecas.Columns.Add("MODULO_ID");
            tabelaRelacaoPecas.Columns.Add("NÚMERO DA ORDEM");
            tabelaRelacaoPecas.Columns.Add("DATA ENTREGA LOTE");
            tabelaRelacaoPecas.Columns.Add("BOX");
            tabelaRelacaoPecas.Columns.Add("ESPECIAL");

            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nCriandos as colunas da tabela de Peças com Código de Barras Errado.");
            PainelSecundario.ScrollToCaret();

            DataTable tabelapecasCodigoBarrasErrado = new DataTable();
            tabelapecasCodigoBarrasErrado.Columns.Add("NÚMERO DO PEDIDO");
            tabelapecasCodigoBarrasErrado.Columns.Add("RAZãO SOCIAL CLIENTE");
            tabelapecasCodigoBarrasErrado.Columns.Add("ORDEM COMPRA PEDIDO");
            tabelapecasCodigoBarrasErrado.Columns.Add("NOME PLANEJADOR");
            tabelapecasCodigoBarrasErrado.Columns.Add("QUANTIDADE");
            tabelapecasCodigoBarrasErrado.Columns.Add("ALTURA PEÇA");
            tabelapecasCodigoBarrasErrado.Columns.Add("LARGURA PEÇA");
            tabelapecasCodigoBarrasErrado.Columns.Add("ESPESSURA PEÇA");
            tabelapecasCodigoBarrasErrado.Columns.Add("CÓDIGO MATERIAL");
            tabelapecasCodigoBarrasErrado.Columns.Add("DESCRIÇãO MATERIAL");
            tabelapecasCodigoBarrasErrado.Columns.Add("LARGURA CORTE MATERIAL");
            tabelapecasCodigoBarrasErrado.Columns.Add("ALTURA CORTE MATERIAL");
            tabelapecasCodigoBarrasErrado.Columns.Add("IMAGEM DO MATERIAL");
            tabelapecasCodigoBarrasErrado.Columns.Add("CÓDIGO PEÇA");
            tabelapecasCodigoBarrasErrado.Columns.Add("COMPLEMENTO");
            tabelapecasCodigoBarrasErrado.Columns.Add("DESCRIÇãO PEÇA");
            tabelapecasCodigoBarrasErrado.Columns.Add("DESENHO PROGRAMADO 1");
            tabelapecasCodigoBarrasErrado.Columns.Add("DESENHO PROGRAMADO 2");
            tabelapecasCodigoBarrasErrado.Columns.Add("DESENHO PROGRAMADO 3");
            tabelapecasCodigoBarrasErrado.Columns.Add("VEIO MATERIAL");
            tabelapecasCodigoBarrasErrado.Columns.Add("BORDA SUPERIOR");
            tabelapecasCodigoBarrasErrado.Columns.Add("BORDA INFERIOR");
            tabelapecasCodigoBarrasErrado.Columns.Add("BORDA ESQUERDA");
            tabelapecasCodigoBarrasErrado.Columns.Add("BORDA DIREITA");
            tabelapecasCodigoBarrasErrado.Columns.Add("DESTINO IMPRESSãO");
            tabelapecasCodigoBarrasErrado.Columns.Add("ID");
            tabelapecasCodigoBarrasErrado.Columns.Add("POSTOS OPERATIVOS");
            tabelapecasCodigoBarrasErrado.Columns.Add("NÚMERO DO LOTE");
            tabelapecasCodigoBarrasErrado.Columns.Add("CÓDIGO CLIENTE");
            tabelapecasCodigoBarrasErrado.Columns.Add("MODULO_ID");
            tabelapecasCodigoBarrasErrado.Columns.Add("NÚMERO DA ORDEM");
            tabelapecasCodigoBarrasErrado.Columns.Add("DATA ENTREGA LOTE");
            tabelapecasCodigoBarrasErrado.Columns.Add("BOX");
            tabelapecasCodigoBarrasErrado.Columns.Add("ESPECIAL");

            try
            {

                var loteCru = File.ReadLines(caminhocolado2).Skip(1);
                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                PainelSecundario.AppendText("\n\nLendo o CSV");
                PainelSecundario.ScrollToCaret();
                IEnumerable<string> listaPerfis = null;
                if (File.Exists(caminhocsvPerfil))
                    listaPerfis = File.ReadLines(caminhocsvPerfil).Skip(1); // Leitura do CSV de Perfil

                // AS LISTAS PARA A EXPLOSAO DO CSV PARA OTIMIZACAO
                var plpinfolhaList = new List<Armazenado> { }; // PLANO A
                var plripaList = new List<Armazenado> { }; // PLANO B
                var pl18List = new List<Armazenado> { }; // PLANO C
                var pl25List = new List<Armazenado> { }; // PLANO D
                var ploutrosList = new List<Armazenado> { }; // PLANO E
                var plimprimirList = new List<Armazenado> { }; // PLANO F
                var plexcluirList = new List<Armazenado> { }; // PLANO G
                var plnesting12mmList = new List<Armazenado> { }; // PLANO NESTING 12MM
                var plnesting15mmList = new List<Armazenado> { }; // PLANO NESTING 15MM
                var plnesting18mmList = new List<Armazenado> { }; // PLANO NESTING 18MM
                var plnesting25mmList = new List<Armazenado> { }; // PLANO NESTING 25MM

                var restantesList = new List<Armazenado> { }; // Restantes (Nao marcados checkbox)

                var matNaoExisteList = new List<Armazenado> { }; // Lista para arquivo XLSX de Materias que não existe
                var brutoList = new List<Armazenado> { }; // Lista para receber todas as peças para produzir o arquivo bruto .xlsx
                var brutoGeralList = new List<Armazenado> { }; // Lista para receber todas as peças para produzir o arquivo bruto .xlsx
                var pecasDXF = new List<Armazenado> { }; // Lista para receber todas as peças para produzir em arquivo DXF

                var pecasCodigoBarrasErradoList = new List<Armazenado> { }; // Lista das Peças com Códigos Errados

                var listaCalculoM2 = new List<Armazenado> { }; // Lista para receber todos os materiais para calculo do m2 geral de cada acabamento

                // AS LISTAS PARA EXPLOSÃO DO CSV PARA RELATORIO DO LOTE
                var MPList = new List<Armazenado> { }; // RELATORIO MONTAGEM DE PERFIL   - POSTO OPERATIVO QUE CONTEM MP
                var PortasColagemList = new List<Armazenado> { }; // RELATORIO PORTAS COLAGEM
                var MEList = new List<Armazenado> { }; // RELATORIO MONTAGEM DE ELETRICA - POSTO OPERATIVO QUE CONTEM ME
                var MOList = new List<Armazenado> { }; // RELATORIO MONTAGEM DE CAIXARIA - PLANO QUE CONTEM PL_MO, PL_PORTAFRENTE_MO OU PL_PINFOLHAMO
                var MAList = new List<Armazenado> { }; // RELATORIO MARCENARIA           - POSTO OPERATIVO QUE CONTEM MA
                var MAOAMGRODAPEList = new List<Armazenado> { }; // RELATORIO MAO AMIGA E RODAPE   - PLANO QUE CONTEM PL_MAOAMG OU PL_ESTRODAPE
                var COLAGEMList = new List<Armazenado> { }; // RELATORIO COLAGEM              - PLANO QUE CONTEM PL_IMPRIMIR OU TAMPO BORD, BORD
                var TAPECARIAList = new List<Armazenado> { }; // RELATORIO TAPEÇARIA
                var FILETACAOList = new List<Armazenado> { };    // RELATORIO FILETACAO MANUAL   - 
                var PINTURAList = new List<Armazenado> { }; // RELATORIO PINTURA LIQUIDA      - POSTO OPERATIVO QUE CONTEM PIN
                var FOLHAList = new List<Armazenado> { }; // RELATORIO FOLHA PINTURA        - POSTO OPERATIVO QUE CONTEM FL
                var MAFOLHAList = new List<Armazenado> { }; // RELATORIO FOLHA MARCENARIA     - AMBIENTES QUE CONTEM FL, PEGAR TUDO QUE É FL E MA DESSE AMBIENTE

                var contagemList = new List<Armazenado> { }; // Contagem para progresso

                var listaAmbientesFolha = new List<string> { };
                var listaPinturaFolha = new List<string> { };
                var listaAlt12To18 = new List<Armazenado>();



                // Materias-Primas que não podem ir para o Nesting
                List<string> matNaoNesting = new List<string>() {
                "6267" , "12267", "18267", "25267", "6268" , "12268", "18268",
                "25268", "6269" , "12269", "18269", "25269", "6273" , "12273",
                "18273", "25273", "6274" , "12274", "18274", "25274", "6275" ,
                "12275", "18275", "25275", "6270" , "12270", "18270", "25270",
                "6260" , "12260", "18260", "25260", "6271" , "12271", "18271",
                "25271",
                };

                // Materias-Primas que não existe, criar a lista para observar e alertar.
                List<string> matNaoExiste = new List<string>() {
                "12192", "12049", "12161", "12255", "12273", "12274", "12270",
                "12260", "12271", "12254", "6990" , "12990", "18990", "25990",
                "37990", "51990", "6991" , "12991", "18991", "25991", "37991",
                "51991", "6992" , "12992", "18992", "25992", "37992", "51992"
                };

                List<string> matTouch = new List<string>() {
                "25270", "25260", "25271", "25273", "25274", "25275"
                };

                // Para fazer a contagem de peças de cada plano gerado em csv
                int qntdPLPinFolha = 0;
                int qntdPLRipa = 0;
                int qntdPL18mm = 0;
                int qntdPL25mm = 0;
                int qntdPLOutros = 0;
                int qntdPLImprimir = 0;
                int qntdPLExcluir = 0;
                int qntdPLNesting12mm = 0;
                int qntdPLNesting15mm = 0;
                int qntdPLNesting18mm = 0;
                int qntdPLNesting25mm = 0;
                int qntdMatNaoExiste = 0;
                int qntdBruto = 0;
                int qntdpecasCodigoBarrasErrado = 0;
                

                int qntdrestantesList = 0;


                // Para fazer a contagem de peças dentro das listas de relatorios de peças
                int qntdPecasLote = 0;
                int qntdCMPLote = 0;
                int qntdCOLote = 0;
                int qntdBOLote = 0;
                int qntdFULote = 0;
                int qntdUSILote = 0;
                int qntdCFLote = 0;
                int qntdMOLote = 0;
                int qntdMALote = 0;
                int qntdMPLote = 0;
                int qntdMELote = 0;
                int qntdPINLote = 0;
                int qntdFLLote = 0;
                int qntdColagemLote = 0;
                int qntdCQLote = 0;
                int qntdFiletacaoLote = 0;
                int qntdBrutoGeral = 0;


                var numeroLote = "";




                int numSemana = 0;

                var unicoCSV = false;
                if (checkBox1.Checked == true)
                    unicoCSV = true;


                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                PainelSecundario.AppendText("\n\nAdicionando cabeçalho nas listas dos planos.");
                PainelSecundario.ScrollToCaret();

                string cabecalho = "NUMERO DO PEDIDO;RAZAO SOCIAL CLIENTE;ORDEM COMPRA PEDIDO;NOME PLANEJADOR;QUANTIDADE;ALTURA PECA;LARGURA PECA;ESPESSURA PECA;CODIGO MATERIAL;DESCRICAO MATERIAL;LARGURA CORTE MATERIAL;ALTURA CORTE MATERIAL;IMAGEM DO MATERIAL;CODIGO PECA;COMPLEMENTO;DESCRICAO PECA;DESENHO PROGRAMADO 1;DESENHO PROGRAMADO 2;DESENHO PROGRAMADO 3;VEIO MATERIAL;BORDA SUPERIOR;BORDA INFERIOR;BORDA ESQUERDA;BORDA DIREITA;DESTINO IMPRESSAO;ID;POSTOS OPERATIVOS;NUMERO DO LOTE;CODIGO CLIENTE;MODULO_ID;NUMERO DA ORDEM;DATA ENTREGA LOTE;PLANO;ESPECIAL";
                // ADICIONA O CABEÇALHO PARA TODOS OS PLANOS
                plpinfolhaList.Add(cabecalho);
                plripaList.Add(cabecalho);
                pl18List.Add(cabecalho);
                pl25List.Add(cabecalho);
                ploutrosList.Add(cabecalho);
                plimprimirList.Add(cabecalho);
                plexcluirList.Add(cabecalho);
                plnesting12mmList.Add(cabecalho);
                plnesting15mmList.Add(cabecalho);
                plnesting18mmList.Add(cabecalho);
                plnesting25mmList.Add(cabecalho);
                restantesList.Add(cabecalho);


                if (unicoCSV == true)
                {
                    brutoList.Add(cabecalho);
                }


                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                PainelSecundario.AppendText("\n\nAnalisando peças com o código errado.");
                PainelSecundario.ScrollToCaret();
                // ANALISANDO QUANTIDADE PARA INCLUSÃO NA MATRIZ DE LOTES
                PainelSecundario.AppendText("\n\nAnalisando quantidade de peças dos postos operativos. (Não contando Excluir e Mao Amiga Menor igual a 450mm");
                PainelSecundario.ScrollToCaret();
                PainelSecundario.AppendText("\n\nAnalisando se a peça já foi classificada.");
                PainelSecundario.ScrollToCaret();

                var qntdalteradasilknaval25mm = 0;
                var qntdalteradatxnaval12mm = 0;

                var lote = loteCru; // captura a variável local


                foreach (var linha in loteCru)
                {
                    Armazenado armazenar = linha;

                    // INSERE A ESPESSURA CORRETA POR GARANTIA
                    if (armazenar.CodigoMaterial.Length == 5)
                        armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 2);
                    if (armazenar.CodigoMaterial.Length == 4)
                        armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 1);

                    // Normaliza a string, substituindo vírgula por ponto e removendo espaços
                    string alturaStr = armazenar.Altura.Trim().Replace(",", ".");
                    string larguraStr = armazenar.Largura.Trim().Replace(",", ".");

                    // Converte para decimal usando InvariantCulture
                    decimal altura = decimal.Parse(alturaStr, CultureInfo.InvariantCulture);
                    decimal largura = decimal.Parse(larguraStr, CultureInfo.InvariantCulture);

                    // verifica se tem porta mpr de 12mm
                    if (armazenar.CodigoMaterial.Contains("12161") && armazenar.DescricaoPeca.Length >= 3 && armazenar.DescricaoPeca.Substring(0, 3) == "MPR" && armazenar.CodigoPeca.Contains("POPA")) // Caso seja uma materia prima de 12mm
                    {
                        PainelSecundario.AppendText("\n\nPeça com matéria prima de 12mm SnowMatt encontrada. Peça: " + armazenar.CodigoPeca + " - " + armazenar.DescricaoPeca);
                        PainelSecundario.ScrollToCaret();
                        armazenar.Espessura = "18";          // Já faz a alteração desta peça que está rodando neste for each
                        armazenar.CodigoMaterial = "18" + armazenar.CodigoMaterial.Substring(2); // Já faz a alteração desta peça que está rodando neste for each
                        armazenar.DescricaoMaterial = armazenar.DescricaoMaterial.Replace("12MM", "18MM");
                        armazenar.DescricaoPeca = armazenar.DescricaoPeca.Replace("X12X12X", "X18X18X");
                        armazenar.DesenhoUm = "Removido Usi"; // remove a usinagem por via das duvidas
                        armazenar.DescricaoMaterial = armazenar.DescricaoMaterial + "ALT 12MM PARA 18MM"; // adiciona na descrição do material a alteração feita para controle
                        foreach (var item in listaCalculoM2) // percorre pela lista calculada m2
                        {
                            Armazenado arm2 = item;
                            if (arm2.Modulo == armazenar.Modulo && armazenar.DescricaoPeca.Length >= 3 && arm2.DescricaoPeca.Substring(0, 3) == "MPR" && arm2.CodigoPeca.Contains("POPA"))
                            {
                                // caso localize adicione na lista para quando o foreach passar novamente ele fazer a troca de espessura no codigo seguinte
                                if (!listaAlt12To18.Any(p => p.NumeroOrdem == arm2.NumeroOrdem))
                                    listaAlt12To18.Add(arm2);
                                PainelSecundario.AppendText($"\n\nPeça parceira encontrada para alteração de 12mm para 18mm. Peça: {arm2.CodigoPeca} - {arm2.DescricaoPeca}\n");
                                PainelSecundario.ScrollToCaret();
                            }
                        }
                    }

                    // SE HOUVER SILK NAVAL DE 25MM ALTERAR PARA BRANCO TX NAVAL 
                    if (armazenar.PostosOperativos.Contains("PIN"))
                    {
                        if (armazenar.CodigoMaterial == "25277")
                        {
                            qntdalteradasilknaval25mm++;
                            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                            PainelSecundario.AppendText($"\n\n{qntdalteradasilknaval25mm} peça(s) alterado o acabamento de 25MM SILK NAVAL para 25MM BRANCO TX NAVAL de peça com PINTURA.");
                            PainelSecundario.ScrollToCaret();

                            armazenar.CodigoMaterial = "25049";
                            armazenar.DescricaoMaterial = "CHAPA 25MM - BRANCO TX NAVAL";

                        }
                    }

                    if (!armazenar.PostosOperativos.Contains("PIN") || !armazenar.PostosOperativos.Contains("FL"))
                    {
                        if (!plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            if (armazenar.CodigoMaterial == "12049")
                            {
                                qntdalteradatxnaval12mm++;
                                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                                PainelSecundario.AppendText($"\n\n{qntdalteradatxnaval12mm} peça(s) alterado o acabamento de 12MM BRANCO TX NAVAL PARA 12MM BRANCO TX de peça SEM PINTURA.");
                                PainelSecundario.ScrollToCaret();

                                armazenar.CodigoMaterial = "12036";
                                armazenar.DescricaoMaterial = "CHAPA 12MM - BRANCO TX";
                            }
                        }
                    }


                    // FAZ A ANALISE DE MATERIAS QUE CONTENHA PINTURA
                    if (armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL"))
                        listaPinturaFolha.Add(armazenar.CodigoMaterial);
                    // PEÇAS COM CODIGO ERRADO
                    if (armazenar.CodigoBarras.Length == 26)
                    {


                        if (!armazenar.CodigoPeca.Contains("FRMD00200") && !armazenar.CodigoPeca.Contains("ENGROSSO") && !armazenar.CodigoPeca.Contains("RIPA002001") && !armazenar.CodigoPeca.Contains("FIMA00200") && armazenar.DescricaoPeca.Substring(0, 3) != "MPR" && !armazenar.DescricaoPeca.Contains("TRASEIRO") && !armazenar.DescricaoPeca.Contains("LATERAL GAV") && !armazenar.DescricaoPeca.Contains("MAO AMIGA") && !armazenar.DescricaoPeca.Contains("CONTRA FRENTE") && altura > 30 && largura > 30)
                        {
                            pecasCodigoBarrasErradoList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                            qntdpecasCodigoBarrasErrado++;
                        }
                    }


                    // PARA NAO CONTAR MAO AMIGA MENOR QUE 451
                    if (armazenar.DescricaoPeca.Substring(0, 9) == "MAO AMIGA" && altura < 451)
                    {
                    }
                    // NAO CONTAR PEÇAS EXCLUIR
                    else if (!plexcluirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                    {

                        if (armazenar.CodigoPeca.Contains("CMP"))
                            qntdCMPLote++;
                        if (armazenar.PostosOperativos.Contains("CO"))
                            qntdCOLote++;
                        if (armazenar.PostosOperativos.Contains("BO"))
                            qntdBOLote++;

                        if (armazenar.PostosOperativos.Contains("FU"))
                            qntdFULote++;

                        if (armazenar.PostosOperativos.Contains("USI"))
                            qntdUSILote++;

                        if (armazenar.PostosOperativos.Contains("CF"))
                            qntdCFLote++;

                        if (armazenar.PostosOperativos.Contains("MO"))
                            qntdMOLote++;

                        if (armazenar.PostosOperativos.Contains("MA"))
                            qntdMALote++;

                        if (armazenar.PostosOperativos.Contains("MP"))
                            qntdMPLote++;

                        if (armazenar.PostosOperativos.Contains("ME"))
                            qntdMELote++;

                        if (armazenar.PostosOperativos.Contains("PIN") && !armazenar.PostosOperativos.Contains("FL"))
                            qntdPINLote++;

                        if (armazenar.PostosOperativos.Contains("FL"))
                            qntdFLLote++;
                        if (armazenar.PostosOperativos.Contains("CQ"))
                            qntdCQLote++;
                        if (armazenar.Espessura == "37" || armazenar.Espessura == "51")
                            qntdColagemLote++;
                        if (matTouch.Contains(armazenar.CodigoMaterial))
                            qntdColagemLote++;
                        if (armazenar.CodigoPeca.Contains("TAMPBORD"))
                            qntdColagemLote++;


                    }


                    qntdPecasLote++;
                    numeroLote = armazenar.NumeroLote;

                    // Se tiver "_", pega o que vem depois
                    if (numeroLote.Contains("_") && numeroLote.Contains("-"))
                    {
                        numeroLote = numeroLote.Substring(numeroLote.LastIndexOf("_") + 1);
                        numeroLote = numeroLote.Substring(0, numeroLote.IndexOf("-"));
                    }
                    else if (numeroLote.Contains("-")) // Se tiver "-", pega o que vem antes
                    {
                        numeroLote = numeroLote.Substring(0, numeroLote.IndexOf("-"));
                    }
                    else if (armazenar.NumeroLote.Length >= 5 && !string.IsNullOrEmpty(numeroLote))
                        numeroLote = new string(armazenar.NumeroLote.Where(char.IsDigit).ToArray()).Substring(0, 5);
                    else if (armazenar.NumeroLote.Length >= 4 && !string.IsNullOrEmpty(numeroLote))
                        numeroLote = new string(armazenar.NumeroLote.Where(char.IsDigit).ToArray()).Substring(0, 4);
                    else if (armazenar.NumeroLote.Length <= 3 && !string.IsNullOrEmpty(numeroLote))
                        numeroLote = new string(armazenar.NumeroLote.Where(char.IsDigit).ToArray()).Substring(0, 3);


                    // Se não sobrar nada, vira "0"
                    if (string.IsNullOrEmpty(numeroLote))
                    {
                        numeroLote = "";
                    }


                    armazenar.AlturaCorte = ((altura / 1000) * (largura / 1000)).ToString(); // Adiciona na coluna Imagem Material o M2 da peça 
                    listaCalculoM2.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                    valorProgresso++;
                    progressBar1.Value = valorProgresso;

                    if (armazenar.DescricaoPeca.Contains(" L ") || armazenar.DescricaoPeca.Contains(".L "))
                    {
                        if (!armazenar.DescricaoPeca.Contains("PRAT") && !armazenar.DescricaoPeca.Contains("PORTA"))
                            qntdFiletacaoLote++;
                    }
                    else if (armazenar.DescricaoPeca.Contains("FREE"))
                    {
                        if (armazenar.Espessura == "37" || armazenar.Espessura == "51")
                            qntdFiletacaoLote++;
                    }
                }
                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                PainelSecundario.AppendText("\n\nAnalisando a quantidade de peças por operação.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COMPONENTES
                PainelSecundario.AppendText($"\nAnalisando a quantidade de componentes (CMP). {qntdCMPLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM CORTE
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Corte (CO). {qntdCOLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM BORDO
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Bordo (BO). {qntdBOLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM FURACAO
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Furação (FU). {qntdFULote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM USINAGENS
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Usinagens (USI). {qntdUSILote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM CONFERENCIA
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Conferência (CF). {qntdCFLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM MONTAGEM CAIXARIA
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Montagem (MO). {qntdMOLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM MARCENARIA
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Marcenaria (MA). {qntdMALote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM MONTAGEM DE PERFIL
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Montagem de Perfil (MP). {qntdMPLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM MONTAGEM ELETRICA
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Montagem de Elétrica (ME). {qntdMELote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM PINTURA LIQUIDA
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Pintura Líquida (PIN). {qntdPINLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COM FOLHA NATURAL
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Folha (FL). {qntdFLLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // CONTAGEM PEÇAS COLAGEM
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Colagem. {qntdColagemLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Filetação Manual. {qntdFiletacaoLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                PainelSecundario.AppendText($"\nAnalisando a quantidade de peças com Controle de Qualidade (CQ). {qntdCQLote} peças encontradas.");
                PainelSecundario.ScrollToCaret();
                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                PainelSecundario.AppendText("\n\nAnalise de quantidades e alterações feitas.");
                PainelSecundario.ScrollToCaret();

                var cont = 0;

                foreach (var linha in listaCalculoM2)
                {

                    Armazenado armazenar = linha;


                    //Console.WriteLine(armazenar.CodigoBarras);
                    // CONSERTA AS ESPESSURAS NA COLUNA DE ACORDO COM O CODIGO DO MATERIAL
                    if (armazenar.CodigoMaterial.Length == 5)
                        armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 2);
                    if (armazenar.CodigoMaterial.Length == 4)
                        armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 1);                   

                    decimal altura = 0m;

                    decimal largura = 0m;
                    // Tente Converte para decimal usando InvariantCulture
                    try
                    {
                        // Normaliza a string, substituindo vírgula por ponto e removendo espaços
                        string alturaStr = armazenar.Altura.Trim().Replace(",", ".");
                        string larguraStr = armazenar.Largura.Trim().Replace(",", ".");
                        altura = decimal.Parse(alturaStr, CultureInfo.InvariantCulture);
                        largura = decimal.Parse(larguraStr, CultureInfo.InvariantCulture);
                    }
                    catch (Exception ex) 
                    {
                        MessageBox.Show("Erro ao converter altura ou largura para decimal. Verifique os dados de entrada. Detalhes do erro: " + ex.Message);
                    }

                    

                    // ALTERACAO DOS CAMPOS DA COLUNA DE NUMERO DO LOTE
                    if (!armazenar.PostosOperativos.Contains("PIN") || !armazenar.PostosOperativos.Contains("FL"))
                    {
                        if (!plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            // SE HOUVER SILK NAVAL DE 25MM ALTERAR PARA BRANCO TX NAVAL O QUE FOR PINTURA
                            if (armazenar.CodigoMaterial == "25277")
                            {
                                armazenar.CodigoMaterial = "25049";
                                armazenar.DescricaoMaterial = "CHAPA 25MM - BRANCO TX NAVAL";

                            }
                            // SE HOUVER BRANCO TX NAVAL DE 12MM ALTERAR PARA BRANCO TX NORMAL 12MM
                            if (armazenar.CodigoMaterial == "12049")
                            {
                                armazenar.CodigoMaterial = "12036";
                                armazenar.DescricaoMaterial = "CHAPA 12MM - BRANCO TX";
                            }
                        }
                    }

                    if (!armazenar.CodigoPeca.Contains("OP:"))
                        armazenar.CodigoPeca = armazenar.CodigoPeca + "-OP:" + armazenar.NumeroOrdem;

                    if (armazenar.Especial == "(P.CATEGORIA_PRODUTO)")
                        armazenar.Especial = "";






                    // até aqui para remover o não compensa e classificar pro nesting mesmo assim -------

                    // PROCURA A SEMANA CASO SEJA 2026

                    List<string> semanas = new List<string> { };

                    if (numeroLote != null)
                    {
                        if (numeroLote != "")
                        {
                            if (armazenar.DataEntrega.Contains("2026") && armazenar.DataEntrega != "02/01/2026")
                            {


                                semanas.Add("13/02/2026");
                                for (var i = 1; i <= 52; i++) // Quantas vezes a quantidade de semanas no ano
                                    semanas.Add(String.Format("{0:dd/MM/yyyy}", Convert.ToDateTime(semanas.LastOrDefault()).AddDays(7)).ToString());
                                numSemana = (int.Parse(semanas.IndexOf(armazenar.DataEntrega).ToString()) + 1);
                                // using (StreamWriter texto = new StreamWriter(caminhocolado.Substring(0, caminhocolado.LastIndexOf("/")) + @"\Semanas.txt"))
                                // {
                                //     var valorSem=0;
                                //     foreach (var a in semanas)
                                //     {
                                //         valorSem = valorSem + 1;
                                //         texto.Write(a + " - Sem: " + valorSem + "\n");
                                //     }

                                //// }
                                //PainelSecundario.AppendText("\n\nNUMERO DA SEMANA É:" + numSemana);
                                //PainelSecundario.ScrollToCaret();

                            }

                            // PROCURA A SEMANA CASO SEJA 2025
                            if (armazenar.DataEntrega.Contains("2025"))
                            {
                                var semana01 = Convert.ToDateTime("31/01/2025").DayOfYear;
                                var semanafinal = Convert.ToDateTime("26/12/2025").DayOfYear;


                                semanas.Add("31/01/2025");
                                for (var i = 1; i <= 51; i++)
                                    semanas.Add(String.Format("{0:dd/MM/yyyy}", Convert.ToDateTime(semanas.LastOrDefault()).AddDays(7)).ToString());

                                if (int.Parse(numeroLote) >= 1019)
                                    numSemana = int.Parse(semanas.IndexOf(armazenar.DataEntrega).ToString()) - 1;
                                else
                                    numSemana = int.Parse(semanas.IndexOf(armazenar.DataEntrega).ToString()) + 1;

                                // using (StreamWriter texto = new StreamWriter(@"C:\Users\IDM\Desktop\TESTE 1016\semanas.txt"))
                                // {
                                //     var valorSem=0;
                                //     foreach (var a in semanas)
                                //     {
                                //         valorSem = valorSem + 1;
                                //         texto.Write(a + " - Sem: " + valorSem + "\n");
                                //     }

                                // }                            


                            }
                            if (armazenar.DataEntrega == "02/01/2026")
                                numSemana = 47;
                            if (armazenar.DataEntrega == "09/01/2026")
                                numSemana = 48;
                            if (armazenar.DataEntrega == "16/01/2026")
                                numSemana = 49;
                            if (armazenar.DataEntrega == "23/01/2026")
                                numSemana = 50;
                            if (armazenar.DataEntrega == "30/01/2026")
                                numSemana = 51;
                            if (armazenar.DataEntrega == "06/02/2026")
                                numSemana = 52;
                        }
                    }
                    //REGRAS PARA POSTOS OPERATIVOS
                    // PREENCHIMENTO ALUMINIO QUE NÃO CONTEM MO
                    if (armazenar.DescricaoPeca.Contains("PREENCH") && armazenar.DescricaoPeca.Contains("ALUM") && !armazenar.PostosOperativos.Contains("MO"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MO";
                    // ESTOFAD QUE NAO TEM TP
                    if (armazenar.DescricaoPeca.Contains("ESTOFAD") && !armazenar.PostosOperativos.Contains("TP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-TP";
                    // BIT QUE NAO TEM MA
                    if (armazenar.DescricaoPeca.Contains("BIT") && !armazenar.PostosOperativos.Contains("MA"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MA";
                    // FRENTE DISPENSER SEM MP
                    if (armazenar.DescricaoPeca.Contains("DISPENSER") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    // BASE CALCEIRO SEM MP INSERE MP
                    if (armazenar.DescricaoPeca.Contains("BASE CALCEIRO") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    // PORTA/FRENTE CHAMFER SEM MP INSERE MP
                    if (armazenar.DescricaoPeca.Contains("CHAMFER") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    // PORTA/FRENTE HALF SEM MP INSERE MP
                    if (armazenar.DescricaoPeca.Contains("HALF") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    // PORTA/FRENTE ITALIAN SEM MP INSERE MP
                    if (armazenar.DescricaoPeca.Contains("ITALIAN") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    // PORTA/FRENTE SOTILLE SEM MP INSERE MP
                    if (armazenar.DescricaoPeca.Contains("SOTILLE") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    // PORTA/FRENTE CURVED SEM MP INSERE MP
                    if (armazenar.DescricaoPeca.Contains("CURVED") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    // PORTA/FRENTE OCULT SEM MP INSERE MP
                    if (armazenar.DescricaoPeca.Contains("OCULT") && !armazenar.PostosOperativos.Contains("MP"))
                        armazenar.PostosOperativos = armazenar.PostosOperativos + "-MP";
                    

                    // FILTRA A MATERIA PRIMA QUE SEJA '12161' SNOW MAT QUE NAO EXISTE, LOCALIZAR A OUTRA MATERIA PRIMA E ALTERAR AS DUAS PARA 18MM
                    // Encontra a peça parceira que terá que fazer a troca de 12 para 18mm
                    bool trocar12To18 = listaAlt12To18.Any(p => p.NumeroOrdem == armazenar.NumeroOrdem);
                    if (trocar12To18 == true)
                    {
                        armazenar.Espessura = "18";          // Já faz a alteração desta peça que está rodando neste for each
                        armazenar.CodigoMaterial = "18" + armazenar.CodigoMaterial.Substring(2); // Já faz a alteração desta peça que está rodando neste for each
                        armazenar.DescricaoMaterial = armazenar.DescricaoMaterial.Replace("12MM", "18MM");
                        armazenar.DescricaoPeca = armazenar.DescricaoPeca.Replace("X12X12X", "X18X18X");
                        armazenar.DescricaoMaterial = armazenar.DescricaoMaterial + " - ALT 12MM PARA 18MM";
                        armazenar.DesenhoUm = "Removido Usi"; // remove a usinagem por via das duvidas
                        PainelSecundario.AppendText($"\n\nPeça alterada de 12mm para 18mm. Peça: {armazenar.CodigoPeca} - {armazenar.DescricaoPeca}\n");
                        PainelSecundario.ScrollToCaret();
                    }

                    if (armazenar.CodigoPeca.Contains(".") == true && !armazenar.NumeroLote.Contains("_") && System.IO.Path.GetFileName(caminhocolado2).Contains("ASM", StringComparison.OrdinalIgnoreCase) == false)
                    {
                        if (!armazenar.NumeroLote.Contains("_"))
                            armazenar.NumeroLote = "S" + numSemana + "_" + armazenar.NumeroLote;

                        if ((armazenar.DescricaoPeca.Substring(0, 9) == "MAO AMIGA" && altura < 451) || (armazenar.CodigoPeca.Contains("POSY00100") && armazenar.Espessura == "37"))
                            armazenar.Plano = "PL_EXCLUIR";
                        else if (plpinfolhatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // 01 - PINFOLHA
                            armazenar.Plano = "PL_PINFOLHA";
                        else if (plpinfolhamotxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            armazenar.Plano = "PL_PINFOLHAMO";
                            MOList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                        }
                        else if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))  // 02 - RIPA
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-02";
                            armazenar.Plano = "PL_RIPA";
                        }
                        else if (pltapecariatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))  // 03 - TAPECARIA
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-03";
                            armazenar.Plano = "PL_TAPECARIA";
                        }
                        else if (plmotxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura) || armazenar.PostosOperativos.Contains("MO"))  // 04 - PL MO
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-04";
                            armazenar.Plano = "PL_MO";
                            MOList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                        }
                        else if (plportafrentemotxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // 04 - PL PORTA FRENTE MO
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-04";
                            armazenar.Plano = "PL_PORTAFRENTEMO";
                            MOList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                        }
                        else if (plmatprimatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // 05 - PL MAT PRIMA
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-05";
                            armazenar.Plano = "PL_MATPRIMA";
                        }
                        else if (plportafrentetxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // 06 - PORTA FRENTE
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-06";
                            armazenar.Plano = "PL_PORTAFRENTE";
                        }
                        else if (pl18mmtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura) && !armazenar.PostosOperativos.Contains("MO")) // 07 - PL 18MM
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-07";
                            armazenar.Plano = "PL_18MM";
                        }
                        else if (pl25mmtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // 08 - PL 25MM
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-08";
                            armazenar.Plano = "PL_25MM";
                        }
                        else if (ploutrostxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // 09 - PL OUTROS
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-09";
                            armazenar.Plano = "PL_OUTROS";
                        }
                        else if (plimprimirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // 10 - PL IMPRIMIR
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-10";
                            armazenar.Plano = "PL_IMPRIMIR";
                        }
                        else if (plmaoamgtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // -- MAO AMG
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote;
                            armazenar.Plano = "PL_MAOAMG";
                        }
                        else if (plestrodapetxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // -- ESTRUTURA RODAPE
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote;
                            armazenar.Plano = "PL_ESTRODAPE";
                            MAOAMGRODAPEList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                        }
                        else if (armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura == "CCFU01.6" && armazenar.DescricaoPeca.Contains(" L "))
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-09";
                            armazenar.Plano = "PL_OUTROS";
                        }
                        else if (plexcluirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)) // -- MAO AMG
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote;
                            armazenar.Plano = "PL_EXCLUIR";
                        }
                        else if (armazenar.CodigoPeca.Substring(0, 3) == "CMP")
                        {
                            if (armazenar.Espessura == "18")
                            {
                                armazenar.NumeroLote = armazenar.NumeroLote + "-07";
                                armazenar.Plano = "PL_18MM";
                            }
                            else if (armazenar.Espessura == "25" && !matTouch.Contains(armazenar.CodigoMaterial))
                            {
                                armazenar.NumeroLote = armazenar.NumeroLote + "-08";
                                armazenar.Plano = "PL_25MM";
                            }
                            else if (armazenar.Espessura == "37" || armazenar.Espessura == "51" || matTouch.Contains(armazenar.CodigoMaterial))
                            {
                                armazenar.NumeroLote = armazenar.NumeroLote + "-09";
                                armazenar.Plano = "PL_IMPRIMIR";
                            }
                            else if (armazenar.Espessura == "12" || armazenar.Espessura == "6" || armazenar.Espessura == "3")
                            {
                                armazenar.NumeroLote = armazenar.NumeroLote + "-10";
                                armazenar.Plano = "PL_OUTROS";
                            }
                        }
                        if (armazenar.CodigoMaterial == "25270" || armazenar.CodigoMaterial == "25260" || armazenar.CodigoMaterial == "25271" || armazenar.CodigoMaterial == "25273" || armazenar.CodigoMaterial == "25274" || armazenar.CodigoMaterial == "25275")
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-10";
                            armazenar.Plano = "PL_IMPRIMIR";
                        }

                    }
                    // INICIO SEPARAÇÃO DO RELATORIO DE LOTES 
                    // INSERÇÃO DA LISTA DE COLAGEM
                    if ((armazenar.CodigoPeca.Contains("BORD") || armazenar.Espessura == "37" || armazenar.Espessura == "51" || armazenar.DescricaoPeca.Substring(0, 3).Contains("MPR")) && !armazenar.CodigoPeca.Contains("POSY0010"))
                    {
                        COLAGEMList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                        if(listaPerfis != null) // caso nao seja nulo ou vazio
                        {
                            foreach (var i in listaPerfis)
                            {
                                Armazenado perfil = i;
                                if (perfil.Modulo == armazenar.Modulo && (perfil.DescricaoPeca.Contains("TRILHO GUIA") || perfil.DescricaoPeca.Contains("FREE")) && !COLAGEMList.Any(p => p.NumeroOrdem == perfil.NumeroOrdem))
                                    COLAGEMList.Add(perfil);
                            }
                        }
                        
                    }

                    if (armazenar.PostosOperativos.Contains("TP") || armazenar.DescricaoPeca.Contains("PREENCH") && armazenar.DescricaoPeca.Contains("ALUM"))
                    {
                        TAPECARIAList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                    }

                    if (armazenar.DescricaoPeca.Contains(" L ") || armazenar.DescricaoPeca.Contains(".L "))
                    {
                        if (!armazenar.DescricaoPeca.Contains("PRAT") && !armazenar.DescricaoPeca.Contains("PORTA"))
                        {
                            FILETACAOList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                        }

                    }
                    else if (armazenar.DescricaoPeca.Contains("FREE"))
                    {
                        if (armazenar.Espessura == "37" || armazenar.Espessura == "51")
                        {
                            FILETACAOList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                        }
                    }

                    // INSERÇÃO NA LISTA DE PINTURA LIQUIDA 
                    if (armazenar.PostosOperativos.Contains("PIN") && !armazenar.PostosOperativos.Contains("FL"))
                    {
                        PINTURAList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                    }
                    // INSERÇÃO NA LISTA DE FOLHA PINTURA
                    if (armazenar.PostosOperativos.Contains("FL"))
                    {
                        listaAmbientesFolha.Add(armazenar.ERP);
                        FOLHAList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                    }
                    // INSERÇÃO NA LISTA DE MARCENARIA 
                    if (armazenar.PostosOperativos.Contains("MA") && !plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                    {
                        MAList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                    }
                    var PortasFree = new[] { "12", "37", "51" };
                    // INSERÇÃO NA LISTA DE MONTAGEM DE PERFIL (MP) 
                    if (armazenar.PostosOperativos.Contains("MP") && !(armazenar.DescricaoPeca.Contains("FREE") && PortasFree.Contains(armazenar.Espessura)) && !armazenar.CodigoPeca.Contains("POPA") && !armazenar.CodigoPeca.Contains("POSY") && !armazenar.PostosOperativos.Contains("ME")) // Nao insere as portas free
                    {
                        MPList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                    }
                    // INSERÇÃO NA LISTA DE MONTAGEM DE ELETRICA (ME)
                    if (armazenar.PostosOperativos.Contains("ME"))
                    {
                        MEList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                    }
                    // COLOCA NA LISTA DE RELACAO DE PÇAS DE ESTRUTURA MAO AMG E RODAPE
                    if (armazenar.DescricaoPeca.Contains("MAO AMIGA") && altura > 450)
                    {
                        MAOAMGRODAPEList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                    }
                    // FIM SEPARAÇÃO DOS RELATORIOS DE LOTES 
                    // ALTERACAO DOS CAMPOS DA COLUNA DE NUMERO DO LOTE

                    if (armazenar.PostosOperativos.Contains("PIN") && System.IO.Path.GetFileName(caminhocolado2).Contains("ASM", StringComparison.OrdinalIgnoreCase) == false || armazenar.PostosOperativos.Contains("FL") && System.IO.Path.GetFileName(caminhocolado2).Contains("ASM", StringComparison.OrdinalIgnoreCase) == false)
                    {
                        try
                        {

                            if (armazenar.NumeroLote.Contains('-'))
                                armazenar.NumeroLote = armazenar.NumeroLote.Substring(0, armazenar.NumeroLote.IndexOf("-"));
                        }
                        catch
                        {
                        }
                        if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-02";
                        }
                        else
                        {
                            armazenar.NumeroLote = armazenar.NumeroLote + "-01";
                            // SE HOUVER SILK NAVAL DE 25MM ALTERAR PARA BRANCO TX NAVAL O QUE FOR PINTURA
                            if (armazenar.CodigoMaterial == "25277")
                            {
                                armazenar.CodigoMaterial = "25049";
                                armazenar.DescricaoMaterial = "CHAPA 25MM - BRANCO TX NAVAL";
                                // using (var logtxt = new StreamWriter("C:/Users/IDM/Desktop/725_relacao/logtxt.txt"))
                                // {
                                //     logtxt.Write("Alteração de peça de SILK NAVAL DE 25MM PARA BRANCO TX NAVAL 25MM (PEÇA PINTURA)");
                                // }

                            }
                            if (armazenar.CodigoMaterial == "12049")
                            {
                                armazenar.CodigoMaterial = "12036";
                                armazenar.DescricaoMaterial = "CHAPA 12MM - BRANCO TX";
                                // using (var logtxt = new StreamWriter("C:/Users/IDM/Desktop/725_relacao/logtxt.txt"))
                                // {
                                //     //logtxt.Write("Alteração de peça de SILK NAVAL DE 25MM PARA BRANCO TX NAVAL 25MM (PEÇA PINTURA)");
                                // }

                            }
                        }
                    }
                    else // SE HOUVER SILK NAVAL DE 25MM ALTERAR PARA SILK 25MM
                    {
                        if (armazenar.CodigoMaterial == "25277")
                        {
                            armazenar.CodigoMaterial = "25244";
                            armazenar.DescricaoMaterial = "SILK 25MM";
                            // using (var logtxt = new StreamWriter("C:/Users/IDM/Desktop/725_relacao/logtxt.txt"))
                            // {
                            //     logtxt.Write("Alteração de peça de SILK NAVAL DE 25MM PARA SILK 25MM");
                            // }
                        }
                        if (armazenar.CodigoMaterial == "12049")
                        {
                            armazenar.CodigoMaterial = "12036";
                            armazenar.DescricaoMaterial = "CHAPA 12MM - BRANCO TX";
                            // using (var logtxt = new StreamWriter("C:/Users/IDM/Desktop/725_relacao/logtxt.txt"))
                            // {
                            //     logtxt.Write("Alteração de peça de SILK NAVAL DE 25MM PARA BRANCO TX NAVAL 25MM (PEÇA PINTURA)");
                            // }

                        }
                    }

                    bool condicoesUnicaNesting = false;


                    if (
                        !matTouch.Contains(armazenar.CodigoMaterial) &&
                        !matNaoNesting.Contains(armazenar.CodigoMaterial) &&
                        !matNaoExiste.Contains(armazenar.CodigoMaterial) &&
                        Math.Min(altura, largura) > 149 &&
                        !armazenar.DescricaoPeca.Contains("BARROTE")
                    )
                    {
                        if (armazenar.Espessura is "18" or "12" or "25" or "15")
                            condicoesUnicaNesting = true;
                    }
                    
                    if (unicoCSV == true)
                    {
                        if (altura > 2700m && largura > 1800m)
                        {
                            pecasDXF.Add(armazenar);
                            contpecasDXF++;
                        }
                        else
                        {                            
                            if (armazenar.DescricaoPeca.Substring(0, 9) == "MAO AMIGA" && altura < 451)
                            {
                                armazenar.Complemento = "G_EXCLUIR";
                                plexcluirList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                qntdPLExcluir++;

                            }
                            else if (plexcluirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura) || (armazenar.CodigoPeca.Contains("POSY00100") && armazenar.Espessura == "37"))
                            {
                                armazenar.Complemento = "G_EXCLUIR";
                                plexcluirList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                qntdPLExcluir++;

                            }
                            else if (armazenar.Espessura == "12")
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLRipa++;
                                    armazenar.Complemento = "B_RIPA";

                                }

                            }
                            else if (armazenar.Espessura == "18")
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLRipa++;
                                    armazenar.Complemento = "B_RIPA";
                                }

                            }
                            else if (armazenar.Espessura == "25" && !matTouch.Contains(armazenar.CodigoMaterial))
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLRipa++;
                                    armazenar.Complemento = "B_RIPA";
                                }
                            }
                            else if (armazenar.Espessura == "15")
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLRipa++;
                                    armazenar.Complemento = "B_RIPA";
                                }
                            }
                            else if (armazenar.Espessura == "6" || armazenar.Espessura == "3")
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    if (chckbxRipa.Checked == true)
                                    {
                                        plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdPLRipa++;
                                        armazenar.Complemento = "B_RIPA";
                                    }
                                    else
                                    {
                                        armazenar.Complemento = "RESTANTES";
                                        restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdrestantesList++;
                                    }
                                }
                            }
                            else if (armazenar.Espessura == "37" || armazenar.Espessura == "51" || matTouch.Contains(armazenar.CodigoMaterial))
                            {
                                armazenar.Complemento = "F_IMPRIMIR";
                                plimprimirList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                qntdPLImprimir++;

                            }
                        }
                    }
                    else // Caso não esteja o UnicoCSV selecionado
                    {
                        if (altura > 2700m && largura > 1800m)
                        {
                            pecasDXF.Add(armazenar);
                            contpecasDXF++;
                        }
                        else
                        {
                            if (armazenar.DescricaoPeca.Substring(0, 9) == "MAO AMIGA" && altura < 451)
                            {
                                armazenar.Complemento = "G_EXCLUIR";
                                if (chckbxExcluir.Checked == true)
                                {
                                    plexcluirList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLExcluir++;
                                }
                                else
                                {
                                    restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdrestantesList++;
                                }
                            }
                            else if (plexcluirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                            {
                                armazenar.Complemento = "G_EXCLUIR";
                                if (chckbxExcluir.Checked == true)
                                {
                                    plexcluirList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLExcluir++;
                                }
                                else
                                {
                                    restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdrestantesList++;
                                }
                            }
                            else if (armazenar.Espessura == "12")
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    if (chckbxRipa.Checked == true)
                                    {
                                        plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdPLRipa++;
                                        armazenar.Complemento = "B_RIPA";
                                    }
                                    else
                                    {
                                        armazenar.Complemento = "RESTANTES";
                                        restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdrestantesList++;
                                    }
                                }
                                else if (condicoesUnicaNesting == true && chckbxNest12MM.Checked == true)
                                {
                                    armazenar.Complemento = "NESTING_12MM";
                                    plnesting12mmList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLNesting12mm++;
                                }
                                else if ((armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL") || listaPinturaFolha.Contains(armazenar.CodigoMaterial)) && chckbxPinFolha.Checked == true)
                                {
                                    armazenar.Complemento = "A_PINFOLHA";
                                    plpinfolhaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLPinFolha++;
                                }
                                else if (chckbxOutros.Checked == true)
                                {
                                    ploutrosList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLOutros++;
                                    armazenar.Complemento = "E_OUTROS";
                                }
                                else
                                {
                                    armazenar.Complemento = "RESTANTES";
                                    restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdrestantesList++;
                                }

                            }
                            else if (armazenar.Espessura == "18")
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    if (chckbxRipa.Checked == true)
                                    {
                                        plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdPLRipa++;
                                        armazenar.Complemento = "B_RIPA";
                                    }
                                    else
                                    {
                                        armazenar.Complemento = "RESTANTES";
                                        restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdrestantesList++;
                                    }
                                }
                                else if (condicoesUnicaNesting == true && chckbxNest18MM.Checked == true)
                                {
                                    plnesting18mmList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLNesting18mm++;
                                    armazenar.Complemento = "NESTING_18MM";
                                }
                                else if ((armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL") || listaPinturaFolha.Contains(armazenar.CodigoMaterial)) && chckbxPinFolha.Checked == true)
                                {
                                    plpinfolhaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLPinFolha++;
                                    armazenar.Complemento = "A_PINFOLHA";
                                }
                                else if (chckbx18MM.Checked == true)
                                {
                                    pl18List.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPL18mm++;
                                    armazenar.Complemento = "C_18MM";
                                }
                                else
                                {
                                    armazenar.Complemento = "RESTANTES";
                                    restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdrestantesList++;
                                }
                            }
                            else if (armazenar.Espessura == "25" && !matTouch.Contains(armazenar.CodigoMaterial))
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    if (chckbxRipa.Checked == true)
                                    {
                                        plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdPLRipa++;
                                        armazenar.Complemento = "B_RIPA";
                                    }
                                    else
                                    {
                                        armazenar.Complemento = "RESTANTES";
                                        restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdrestantesList++;
                                    }
                                }
                                else if (condicoesUnicaNesting == true && armazenar.Espessura == "25" && chckbxNest25MM.Checked == true)
                                {
                                    plnesting25mmList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLNesting25mm++;
                                    armazenar.Complemento = "NESTING_25MM";
                                }
                                else if ((armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL") || listaPinturaFolha.Contains(armazenar.CodigoMaterial)) && chckbxPinFolha.Checked == true)
                                {
                                    plpinfolhaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLPinFolha++;
                                    armazenar.Complemento = "A_PINFOLHA";
                                }
                                else if (chckbx25MM.Checked == true)
                                {
                                    pl25List.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPL25mm++;
                                    armazenar.Complemento = "D_25MM";
                                }
                                else
                                {
                                    armazenar.Complemento = "RESTANTES";
                                    restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdrestantesList++;
                                }
                            }
                            else if (armazenar.Espessura == "15")
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    if (chckbxRipa.Checked == true)
                                    {
                                        plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdPLRipa++;
                                        armazenar.Complemento = "B_RIPA";
                                    }
                                    else
                                    {
                                        armazenar.Complemento = "RESTANTES";
                                        restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdrestantesList++;
                                    }
                                }
                                else if (condicoesUnicaNesting == true && armazenar.Espessura == "15" && chckbxNest15MM.Checked == true)
                                {
                                    plnesting15mmList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLNesting15mm++;
                                    armazenar.Complemento = "NESTING_15MM";
                                }
                                else if ((armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL") || listaPinturaFolha.Contains(armazenar.CodigoMaterial)) && chckbxPinFolha.Checked == true)
                                {
                                    armazenar.Complemento = "A_PINFOLHA";
                                    plpinfolhaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLPinFolha++;
                                }
                                else if (chckbxOutros.Checked == true)
                                {
                                    ploutrosList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLOutros++;
                                    armazenar.Complemento = "E_OUTROS";
                                }
                                else
                                {
                                    armazenar.Complemento = "RESTANTES";
                                    restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdrestantesList++;
                                }
                            }
                            else if (armazenar.Espessura == "6" || armazenar.Espessura == "3")
                            {
                                if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                                {
                                    if (chckbxRipa.Checked == true)
                                    {
                                        plripaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdPLRipa++;
                                        armazenar.Complemento = "B_RIPA";
                                    }
                                    else
                                    {
                                        armazenar.Complemento = "RESTANTES";
                                        restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                        qntdrestantesList++;
                                    }
                                }
                                else if ((armazenar.PostosOperativos.Contains("PIN") || armazenar.PostosOperativos.Contains("FL") || listaPinturaFolha.Contains(armazenar.CodigoMaterial)) && chckbxPinFolha.Checked == true)
                                {
                                    armazenar.Complemento = "A_PINFOLHA";
                                    plpinfolhaList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLPinFolha++;
                                }
                                else if (chckbxOutros.Checked == true)
                                {
                                    ploutrosList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLOutros++;
                                    armazenar.Complemento = "E_OUTROS";
                                }
                                else
                                {
                                    armazenar.Complemento = "RESTANTES";
                                    restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdrestantesList++;
                                }

                            }
                            else if (armazenar.Espessura == "37" || armazenar.Espessura == "51" || matTouch.Contains(armazenar.CodigoMaterial))
                            {
                                armazenar.Complemento = "F_IMPRIMIR";
                                if (chckbxImprimir.Checked == true)
                                {
                                    plimprimirList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdPLImprimir++;
                                }
                                else
                                {
                                    restantesList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura.ToString() + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                                    qntdrestantesList++;
                                }
                            }// Fecha o Contem "." no Codigo peça
                        }
                    }
                    // FIM DA SEPARACAO

                    if (matNaoExiste.Contains(armazenar.CodigoMaterial))
                    {
                        matNaoExisteList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                        qntdMatNaoExiste++;
                    }

                    contagemList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + "PLIMPRIMIR" + ";" + armazenar.Especial);




                    if (armazenar.DescricaoPeca.Substring(0, 9) == "MAO AMIGA" && altura < 451)
                    {
                    }
                    else if (plexcluirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                    {
                    }
                    else if (plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))
                    {
                    }
                    else if (armazenar.Espessura != "37" && armazenar.Espessura != "51" && !matTouch.Contains(armazenar.CodigoMaterial))
                    {

                        brutoList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                        qntdBruto++;

                    }

                    brutoGeralList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);
                    qntdBrutoGeral++;

                    valorProgresso++;
                    progressBar1.Value = valorProgresso;

                    // Adiciona os itens que serão mantidas as ordens no MES em lista 
                    if (armazenar.Complemento == "G_EXCLUIR" || armazenar.Complemento == "F_IMPRIMIR")
                        pecasInativar.Add(int.Parse(armazenar.NumeroOrdem));
                    else
                        pecasManterOrdem.Add(long.Parse(armazenar.NumeroOrdem));

                    // Listagem das peças codigo errado envio para o mes
                    if (armazenar.CodigoBarras.Length != 0)
                    {
                        pecasEnvioMes.Add(new Ordem
                        {
                            id_ordem = int.Parse(armazenar.NumeroOrdem),
                            campo05 = armazenar.CodigoBarras
                        });
                    }

                }

                PainelSecundario.AppendText($"\n\n{contagemList.Count()} peça(s) classificada(s).");
                PainelSecundario.ScrollToCaret();

                PainelSecundario.AppendText("\n\nAdicionando as peças que contém Folha e operação de MA (Marcenaria) na listagem de MA FOLHA.");
                PainelSecundario.ScrollToCaret();

                bloquearFecha = true;

                foreach (var linha in loteCru)
                {

                    Armazenado armazenar = linha;

                    // CONSERTA AS ESPESSURAS NA COLUNA DE ACORDO COM O CODIGO DO MATERIAL
                    if (armazenar.CodigoMaterial.Length == 5)
                        armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 2);
                    if (armazenar.CodigoMaterial.Length == 4)
                        armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 1);



                    if (armazenar.CodigoPeca.Contains(".") == true)
                        if (listaAmbientesFolha.Contains(armazenar.ERP))
                            if (armazenar.PostosOperativos.Contains("FL") || armazenar.PostosOperativos.Contains("MA"))
                            {
                                MAFOLHAList.Add(armazenar.ERP + ";" + armazenar.RazaoSocial + ";" + armazenar.PedidoAmbiente + ";" + armazenar.Planejador + ";" + armazenar.Quantidade + ";" + armazenar.Altura + ";" + armazenar.Largura + ";" + armazenar.Espessura + ";" + armazenar.CodigoMaterial + ";" + armazenar.DescricaoMaterial + ";" + armazenar.LarguraCorte + ";" + armazenar.AlturaCorte + ";" + armazenar.ImagemMaterial + ";" + armazenar.CodigoPeca + ";" + armazenar.Complemento + ";" + armazenar.DescricaoPeca + ";" + armazenar.DesenhoUm + ";" + armazenar.DesenhoDois + ";" + armazenar.DesenhoTres + ";" + armazenar.VeioMaterial + ";" + armazenar.BordaSup + ";" + armazenar.BordaInf + ";" + armazenar.BordaEsq + ";" + armazenar.BordaDir + ";" + armazenar.DestinoImpressao + ";" + armazenar.CodigoBarras + ";" + armazenar.PostosOperativos + ";" + armazenar.NumeroLote + ";" + armazenar.CodigoCliente + ";" + armazenar.Modulo + ";" + armazenar.NumeroOrdem + ";" + armazenar.DataEntrega + ";" + armazenar.Plano + ";" + armazenar.Especial);

                            }

                    valorProgresso++;
                    progressBar1.Value = valorProgresso;

                }

                string nomedoarquivo = System.IO.Path.GetFileNameWithoutExtension(caminhocolado2);
                string nomeSubdoarquivo = "";
                if (nomedoarquivo.Contains("ASM") || nomedoarquivo.Contains("RETFAB"))
                {
                    int primeiro = 0;
                    if (nomedoarquivo.Contains('_'))
                        primeiro = nomedoarquivo.IndexOf('_');

                    if (primeiro != -1)
                    {
                        int segundo = nomedoarquivo.IndexOf('_', primeiro + 1);
                        if (segundo != -1)
                            nomeSubdoarquivo = nomedoarquivo.Substring(0, segundo);
                    }
                }
                else
                    nomeSubdoarquivo = numeroLote + "_SEM_" + numSemana.ToString();


                // SELECAO DA LISTA PARA TRANSFORMAR EM CSV OS ARQUVIOS DE OTIMIZACAO
                var dadosplpinfolhalist = plpinfolhaList.Select(plripatitulo => (string)plripatitulo);
                var dadosplripalist = plripaList.Select(plripatitulo => (string)plripatitulo);
                var dadospl18list = pl18List.Select(pl18titulo => (string)pl18titulo);
                var dadospl25list = pl25List.Select(pl25titulo => (string)pl25titulo);
                var dadosploutroslist = ploutrosList.Select(ploutrostitulo => (string)ploutrostitulo);
                var dadosplimprimirlist = plimprimirList.Select(plimprimirtitulo => (string)plimprimirtitulo);
                var dadosplexcluirlist = plexcluirList.Select(plexcluirtitulo => (string)plexcluirtitulo);
                var dadosplnesting12mmlist = plnesting12mmList.Select(plnesting12mm => (string)plnesting12mm);
                var dadosplnesting15mmlist = plnesting15mmList.Select(plnesting15mm => (string)plnesting15mm);
                var dadosplnesting18mmlist = plnesting18mmList.Select(plnesting18mm => (string)plnesting18mm);
                var dadosplnesting25mmlist = plnesting25mmList.Select(plnesting25mm => (string)plnesting25mm);
                var dadosbrutoList = brutoList.Select(brutoListCSV => (string)brutoListCSV);
                var dadosbrutoListGeral = brutoGeralList.Select(brutoListGeralCSV => (string)brutoListGeralCSV);
                var dadosRestantesList = restantesList.Select(restantespecas => (string)restantespecas);
                //var dadosplnestinglist = plnestingList.Select(plnestingtitulo => (string)plnestingtitulo);


                // SELECAO DA LISTA PARA TRANSFORMAR EM CSV OS ARQUVIOS DE RELATORIO DO LOTE
                // var dadosMPList = MPList.Select(mptitulo => (string)mptitulo);
                // var dadosMEList = MEList.Select(metitulo => (string)metitulo);
                var dadosMOList = MOList.Select(motitulo => (string)motitulo);
                var dadosMAList = MAList.Select(matitulo => (string)matitulo);
                var dadosCOLAGEMList = COLAGEMList.Select(colagemtitulo => (string)colagemtitulo);
                var dadosFILETACAOList = FILETACAOList.Select(filetacaotitulo => (string)filetacaotitulo);
                var dadosMAOAMG_ESTRODAPEList = MAOAMGRODAPEList.Select(maoamgrodapetitulo => (string)maoamgrodapetitulo);
                var dadosPINTURAList = PINTURAList.Select(pinturatitulo => (string)pinturatitulo);
                var dadosFOLHAList = FOLHAList.Select(folhatitulo => (string)folhatitulo);
                var dadosMAFOLHAList = MAFOLHAList.Select(mafolhatitulo => (string)mafolhatitulo);

                //var dadosMATNAOEXISTEList = matNaoExiste.Select(matNaoExistetitulo => (string)matNaoExistetitulo);

                if (!Directory.Exists(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\"))
                    Directory.CreateDirectory(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\");
                // CRIA O ARQUIVO .XLSX DAS PEÇAS QUE ESTÃO COM MATERIA-PRIMA INEXISTENTE 
                if (qntdMatNaoExiste > 0)
                {
                    matNaoExisteList.ForEach(pecas =>
                    {
                        tabelaMatPrimaNaoExiste.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });

                    using (XLWorkbook xlsMatPrimaNaoExiste = new XLWorkbook())
                    {
                        xlsMatPrimaNaoExiste.AddWorksheet(tabelaMatPrimaNaoExiste, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsMatPrimaNaoExiste.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + numeroLote + "_MAT-PRIMA_NAO_EXISTE_" + qntdMatNaoExiste + ".xlsx");
                        }

                    }
                }
                tabelaDXF.Rows.Clear(); 
                
                if (contpecasDXF > 0)
                {
                    pecasDXF.ForEach(pecas =>
                    {
                        tabelaDXF.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlspecasDXF = new XLWorkbook())
                    {
                        xlspecasDXF.AddWorksheet(tabelaDXF, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            var nomePecasDXF = "";
                            if (System.IO.Path.GetFileName(caminhocolado2).Contains("RETFAB", StringComparison.OrdinalIgnoreCase) == false)
                                xlspecasDXF.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_ATENCAO_USAR_DXF_" + contpecasDXF + ".xlsx");
                            else
                            {
                                nomePecasDXF = System.IO.Path.GetFileNameWithoutExtension(caminhocolado2);
                                xlspecasDXF.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_ATENCAO_USAR_DXF_" + contpecasDXF + ".xlsx");
                            }
                        }
                    }
                }
                valorProgresso++;
                progressBar1.Value = valorProgresso;
                if (qntdpecasCodigoBarrasErrado > 0)
                {
                    pecasCodigoBarrasErradoList.ForEach(pecas =>
                    {
                        tabelapecasCodigoBarrasErrado.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });

                    using (XLWorkbook xlspecasCodigoBarrasErrado = new XLWorkbook())
                    {
                        xlspecasCodigoBarrasErrado.AddWorksheet(tabelapecasCodigoBarrasErrado, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            var nomeCodPecasErrado = "";
                            if (System.IO.Path.GetFileName(caminhocolado2).Contains("RETFAB", StringComparison.OrdinalIgnoreCase) == false)
                                xlspecasCodigoBarrasErrado.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_PecasCodigoBarrasErrado_" + qntdpecasCodigoBarrasErrado + ".xlsx");
                            else
                            {
                                nomeCodPecasErrado = System.IO.Path.GetFileNameWithoutExtension(caminhocolado2);
                                xlspecasCodigoBarrasErrado.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_PecasCodigoBarrasErrado_" + qntdpecasCodigoBarrasErrado + ".xlsx");

                            }
                        }

                    }
                }

                valorProgresso++;
                progressBar1.Value = valorProgresso;
                string caminhoFinal = caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\"))
                      + @"\Relatorios de peças\BRUTO.xlsx";

                if (!File.Exists(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "BRUTO.xlsx"))
                {
                    string pastaRel = caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\"))
                  + @"\Relatorios de peças";

                    if (!Directory.Exists(pastaRel))
                        Directory.CreateDirectory(pastaRel);

                    brutoList.ForEach(pecas =>
                    {
                        tabelaBruto.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsBruto = new XLWorkbook())
                    {
                        xlsBruto.AddWorksheet(tabelaBruto, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsBruto.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "BRUTO.xlsx");
                        }

                    }

                    valorProgresso++;
                    progressBar1.Value = valorProgresso;
                    tabelaRelacaoPecas.Rows.Clear();
                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    MAList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsMA = new XLWorkbook())
                    {
                        xlsMA.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsMA.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "MA.xlsx");
                        }


                    }
                    valorProgresso++;
                    progressBar1.Value = valorProgresso;
                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();

                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    MAFOLHAList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsMAFOLHA = new XLWorkbook())
                    {
                        xlsMAFOLHA.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsMAFOLHA.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "MA_FOLHA.xlsx");
                        }

                    }
                    valorProgresso++;
                    progressBar1.Value = valorProgresso;
                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();

                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    PINTURAList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsPINTURA = new XLWorkbook())
                    {
                        xlsPINTURA.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsPINTURA.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "PINTURA.xlsx");
                        }

                    }
                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();
                    valorProgresso++;
                    progressBar1.Value = valorProgresso;
                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    FOLHAList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsRelacao = new XLWorkbook())
                    {
                        xlsRelacao.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsRelacao.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "FOLHA.xlsx");
                        }

                    }
                    valorProgresso++;
                    progressBar1.Value = valorProgresso;
                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();

                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    MAOAMGRODAPEList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsRelacao = new XLWorkbook())
                    {
                        xlsRelacao.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsRelacao.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "MAOAMG_RODAPE.xlsx");
                        }

                    }
                    valorProgresso++;
                    progressBar1.Value = valorProgresso;
                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();

                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    COLAGEMList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsRelacao = new XLWorkbook())
                    {
                        xlsRelacao.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsRelacao.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "COLAGEM.xlsx");
                        }

                    }
                    tabelaRelacaoPecas.Rows.Clear();

                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    TAPECARIAList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsRelacao = new XLWorkbook())
                    {
                        xlsRelacao.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsRelacao.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "TAPECARIA.xlsx");
                        }

                    }
                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();

                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    FILETACAOList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsRelacao = new XLWorkbook())
                    {
                        xlsRelacao.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsRelacao.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "FILETACAO.xlsx");
                        }

                    }
                    valorProgresso++;
                    progressBar1.Value = valorProgresso;
                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();

                    // CRIAÇÃO DOS XLSX DA RELACAO DE PECAS MP
                    MOList.ForEach(pecas =>
                    {
                        tabelaRelacaoPecas.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsRelacao = new XLWorkbook())
                    {
                        xlsRelacao.AddWorksheet(tabelaRelacaoPecas, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsRelacao.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\" + "MO.xlsx");
                        }

                    }

                    RelatorioMoPerfil(MPList);
                    RelatorioMoEletrica(MEList);

                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();
                    valorProgresso++;
                    progressBar1.Value = valorProgresso;

                }


                // Para fazer a contagem de peças de cada plano gerado em csv
                emailqntdPLPinFolha = qntdPLPinFolha;
                emailqntdPLRipa = qntdPLRipa;
                emailqntdPL18mm = qntdPL18mm;
                emailqntdPL25mm = qntdPL25mm;
                emailqntdPLOutros = qntdPLOutros;
                emailqntdPLImprimir = qntdPLImprimir;
                emailqntdPLExcluir = qntdPLExcluir;
                emailqntdPLNesting12mm = qntdPLNesting12mm;
                emailqntdPLNesting15mm = qntdPLNesting15mm;
                emailqntdPLNesting18mm = qntdPLNesting18mm;
                emailqntdPLNesting25mm = qntdPLNesting25mm;
                emailqntdMatNaoExiste = qntdMatNaoExiste;
                emailqntdBruto = qntdBruto;
                emailqntdpecasCodigoBarrasErrado = qntdpecasCodigoBarrasErrado;


                // Para fazer a contagem de peças dentro das listas de relatorios de peças
                emailqntdPecasLote = qntdPecasLote;
                emailqntdCMPLote = qntdCMPLote;
                emailqntdCOLote = qntdCOLote;
                emailqntdBOLote = qntdBOLote;
                emailqntdFULote = qntdFULote;
                emailqntdUSILote = qntdUSILote;
                emailqntdCFLote = qntdCFLote;
                emailqntdMOLote = qntdMOLote;
                emailqntdMALote = qntdMALote;
                emailqntdMPLote = qntdMPLote;
                emailqntdMELote = qntdMELote;
                emailqntdPINLote = qntdPINLote;
                emailqntdFLLote = qntdFLLote;
                emailqntdColagemLote = qntdColagemLote;
                emailqntdFiletacaoLote = qntdFiletacaoLote;
                emailqntdCQLote = qntdCQLote;




                if (unicoCSV == false)
                {

                    // LIMPA A TABELA PARA REUTILZAR
                    tabelaRelacaoPecas.Rows.Clear();

                    brutoGeralList.ForEach(pecas =>
                    {
                        tabelaBruto.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                    });
                    using (XLWorkbook xlsBruto = new XLWorkbook())
                    {
                        xlsBruto.AddWorksheet(tabelaBruto, "Sheet");
                        using (MemoryStream ns = new MemoryStream())
                        {
                            xlsBruto.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_BRUTO_" + qntdBrutoGeral + ".xlsx");
                            //xlsBruto.SaveAs(@"J:\MARCENARIA" + @"\" + numeroLote + "_BRUTO_" + qntdBrutoGeral + ".xlsx");

                        }

                    }

                    var letraNum = 0;
                    var letraAlf = "";
                    if (qntdPLPinFolha > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_PINFOLHA_" + qntdPLPinFolha + ".csv", dadosplpinfolhalist);

                    }
                    // SALVA TODOS OS CVS PARA OTIMIZACAO
                    letraAlf = "";
                    if (qntdPLRipa > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_RIPA_" + qntdPLRipa + ".csv", dadosplripalist);

                    }
                    letraAlf = "";
                    if (qntdPL18mm > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_18MM_" + qntdPL18mm + ".csv", dadospl18list);

                    }
                    letraAlf = "";
                    if (qntdPL25mm > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_25MM_" + qntdPL25mm + ".csv", dadospl25list);

                    }
                    letraAlf = "";
                    if (qntdPLOutros > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_OUTROS_" + qntdPLOutros + ".csv", dadosploutroslist);

                    }

                    letraAlf = "";
                    if (qntdPLImprimir > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_IMPRIMIR_" + qntdPLImprimir + ".csv", dadosplimprimirlist);


                    }
                    letraAlf = "";
                    if (qntdPLExcluir > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_EXCLUIR_" + qntdPLExcluir + ".csv", dadosplexcluirlist);


                    }

                    if (qntdPLNesting12mm > 0)
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_NEST12MM_" + qntdPLNesting12mm + ".csv", dadosplnesting12mmlist);
                    if (qntdPLNesting15mm > 0)
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_NEST15MM_" + qntdPLNesting15mm + ".csv", dadosplnesting15mmlist);
                    if (qntdPLNesting18mm > 0)
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_NEST18MM_" + qntdPLNesting18mm + ".csv", dadosplnesting18mmlist);
                    if (qntdPLNesting25mm > 0)
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_NEST25MM_" + qntdPLNesting25mm + ".csv", dadosplnesting25mmlist);
                    if (qntdrestantesList > 0)
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_RESTANTES_" + qntdrestantesList + ".csv", dadosRestantesList);


                    // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                    PainelSecundario.AppendText("\n-->  RESUMO  <--");
                    PainelSecundario.AppendText($"\n\nQntd de pçs      - PL_PINFOLHA:     {qntdPLPinFolha}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_RIPA:         {qntdPLRipa}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_18MM:         {qntdPL18mm}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_25MM:         {qntdPL25mm}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_OUTROS:       {qntdPLOutros}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_IMPRIMIR:     {qntdPLImprimir}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_EXCLUIR:      {qntdPLExcluir}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_NESTING_12MM: {qntdPLNesting12mm}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_NESTING_15MM: {qntdPLNesting15mm}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_NESTING_18MM: {qntdPLNesting18mm}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - PL_NESTING_25MM: {qntdPLNesting25mm}");
                    PainelSecundario.AppendText($"\nQntd de pçs      - Restantes CSV: {qntdrestantesList}");
                    PainelSecundario.AppendText($"\n\nQntd de pçs      - BRUTO:           {qntdBrutoGeral}");

                    if (qntdMatNaoExiste > 0)
                        PainelSecundario.AppendText($"Qntd de pçs com matéria-prima que não existe: {qntdMatNaoExiste}.");
                    if (contpecasDXF > 0)
                    {
                        PainelSecundario.AppendText($"Qntd de pçs que deverão usar DXF para cortar: {contpecasDXF}.");
                        MessageBox.Show($"Há {contpecasDXF} peças para usar o DXF, essas peças foram removidas dos planos de corte e associadas ao Excel: {nomeSubdoarquivo}_ATENCAO_USAR_DXF_{contpecasDXF}.xlsx");
                    }
                                      
                    PainelSecundario.ScrollToCaret();

                    PainelSecundario.AppendText($"\n\nSoma de todos os planos é: {qntdPLPinFolha + qntdPLRipa + qntdPL18mm + qntdPL25mm + qntdPLOutros + qntdPLImprimir + qntdPLExcluir + qntdPLNesting12mm + qntdPLNesting15mm + qntdPLNesting18mm + qntdPLNesting25mm} pçs" +
                        $"\n e {contpecasDXF} de peças para cortar em DXF.");
                    PainelSecundario.ScrollToCaret();
                    //var ultimom2 = "";

                    valorProgresso++;
                    progressBar1.Value = valorProgresso;
                }
                else // SE ESTÁ SELECIONADO O Unico CSV / Ripa / Imprimir / Excluir
                {
                    tabelaRelacaoPecas.Rows.Clear();

                    var letraNum = 0;
                    var letraAlf = "";
                    if (PREqntdPLPinFolha > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";

                    }
                    // SALVA TODOS OS CVS PARA OTIMIZACAO
                    letraAlf = "";
                    if (PREqntdPLRipa > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_RIPA_" + qntdPLRipa + ".csv", dadosplripalist);
                    }
                    letraAlf = "";
                    if (PREqntdPL18mm > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";

                    }
                    letraAlf = "";
                    if (PREqntdPL25mm > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";

                    }
                    letraAlf = "";
                    if (PREqntdPLOutros > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";

                    }
                    letraAlf = "";
                    if (PREqntdPLImprimir > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_IMPRIMIR_" + qntdPLImprimir + ".csv", dadosplimprimirlist);
                    }
                    letraAlf = "";
                    if (PREqntdPLExcluir > 0)
                    {
                        letraNum++;
                        if (letraNum == 1)
                            letraAlf = "_A";
                        if (letraNum == 2)
                            letraAlf = "_B";
                        if (letraNum == 3)
                            letraAlf = "_C";
                        if (letraNum == 4)
                            letraAlf = "_D";
                        if (letraNum == 5)
                            letraAlf = "_E";
                        if (letraNum == 6)
                            letraAlf = "_F";
                        if (letraNum == 7)
                            letraAlf = "_G";
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + letraAlf.ToString() + "_EXCLUIR_" + qntdPLExcluir + ".csv", dadosplexcluirlist);
                    }
                    var nomeArquivo = "";

                    if (System.IO.Path.GetFileNameWithoutExtension(caminhocolado2).Contains("RETFAB", StringComparison.OrdinalIgnoreCase))
                    {
                        nomeArquivo = System.IO.Path.GetFileNameWithoutExtension(caminhocolado2);
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeArquivo + "_CSVBRUTO_" + qntdBruto + ".csv", dadosbrutoList);
                    }
                    else
                        File.WriteAllLines(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\" + nomeSubdoarquivo + "_CSVBRUTO_" + qntdBruto + ".csv", dadosbrutoList);

                }

                tabelaRelacaoPecas.Rows.Clear();
                tabelaBruto.Rows.Clear();
                brutoGeralList.ForEach(pecas =>
                {
                    tabelaBruto.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
                });
                using (XLWorkbook xlsBruto = new XLWorkbook())
                {
                    xlsBruto.AddWorksheet(tabelaBruto, "Sheet");
                    using (MemoryStream ns = new MemoryStream())
                    {
                        var contemBruto = false;
                        var caminhoXLSBruto = System.IO.Path.GetDirectoryName(caminhocolado2) + @"\" + nomeSubdoarquivo + "_BRUTO_UNICOCSV_" + qntdBrutoGeral + ".xlsx";
                        foreach (var files in Directory.GetFiles(System.IO.Path.GetDirectoryName(caminhocolado2)))
                        {
                            if (System.IO.Path.GetFileName(files).Contains("BRUTO") && System.IO.Path.GetFileName(files).Contains(".xlsx"))
                                contemBruto = true;
                        }
                        if (contemBruto == false)
                        {
                            xlsBruto.SaveAs(caminhoXLSBruto);
                            PainelSecundario.AppendText($"\n\nO arquivo bruto (.xlsx) com todas as peças foi criado");
                        }
                        else
                        {
                            PainelSecundario.AppendText($"\n\nJá existe um bruto (.xlsx) criado na pasta! Não será criado outro.");
                        }
                        PainelSecundario.ScrollToCaret();
                    }
                }

                if (qntdMatNaoExiste > 0)
                    PainelSecundario.AppendText($"\n\nQNTD DE PEÇAS COM MATERIA-PRIMA NÃO EXISTENTE ENCONTRADO:  {qntdMatNaoExiste}\n");
                if (qntdpecasCodigoBarrasErrado > 0)
                    PainelSecundario.AppendText($"\n\nQNTD DE PEÇAS COM CODIGO DE BARRAS ERRADO ENCONTRADO:  {qntdpecasCodigoBarrasErrado}\n");
                PainelSecundario.ScrollToCaret();

                if (caminhocolado2.Contains("RETFAB", StringComparison.OrdinalIgnoreCase) == false)
                {
                    try
                    {
                        label2.Text = "Status: Alimentando a Matriz Lotes";
                        var caminhoXlsx = @"J:\PCP\PCP 2025\APONTAMENTO DE PRODUÇÃO\Analise_Lote\Matriz_Lotes.xlsx";

                        var matrizDadosList = new List<Matriz_Lotes>();

                        using (var wb = new XLWorkbook(caminhoXlsx))
                        {
                            var ws = wb.Worksheet("MatrizLotes");

                            foreach (var row in ws.RowsUsed().Skip(1))
                            {
                                matrizDadosList.Add(new Matriz_Lotes
                                {
                                    dataprocesso = row.Cell(1).GetString(),
                                    nomearquivo = row.Cell(2).GetString(),
                                    semana = row.Cell(3).GetString(),
                                    lote = row.Cell(4).GetString(),
                                    qntdPecas = row.Cell(5).GetString(),
                                    qntdCMP = row.Cell(6).GetString(),
                                    porcentCMP = row.Cell(7).GetString(),
                                    qntdCO = row.Cell(8).GetString(),
                                    qntdBO = row.Cell(9).GetString(),
                                    qntdFU = row.Cell(10).GetString(),
                                    qntdUSI = row.Cell(11).GetString(),
                                    qntdCF = row.Cell(12).GetString(),
                                    qntdMO = row.Cell(13).GetString(),
                                    qntdMA = row.Cell(14).GetString(),
                                    qntdMP = row.Cell(15).GetString(),
                                    qntdME = row.Cell(16).GetString(),
                                    qntdPIN = row.Cell(17).GetString(),
                                    qntdFL = row.Cell(18).GetString(),
                                    qntdColagem = row.Cell(19).GetString(),
                                    qntdCQ = row.Cell(20).GetString(),
                                    qntdFiletacao = row.Cell(21).GetString()
                                });
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Erro ao enviar para a matriz: {ex}");
                    }

                }


                PainelSecundario.AppendText("\n\nQuantidade peças para Inativar operação CO: " + pecasInativar.Count());
                PainelSecundario.ScrollToCaret();

                if (chckbxManterOrdens.Checked == true || chckbxCorrecaoCodBarras.Checked == true)
                {

                    // PROCESSO DE MANTER ORDENS NA API
                    var api = new ApiClient("http://192.168.3.40/");

                    // 1) LOGIN
                    string token = await api.LoginAsync("ppcpAPI", "ppcp@@ID%$");
                    //string token = await api.LoginAsync("admin", "benelson");


                    if (chckbxCorrecaoCodBarras.Checked == true)
                    {
                        label2.Text = "Status: Enviando Código de Barras para o MES";
                        var listaMes = pecasEnvioMes.ToList();

                        var response = await api.AtualizarOrdensAsync(
                            listaMes,
                            token
                        );

                        //PainelSecundario.AppendText(
                        //    JsonSerializer.Serialize(listaMes)
                        //);

                        //var body = new AtualizarOrdensRequest
                        //{
                        //    ordens = listaMes
                        //};

                        //var json = JsonSerializer.Serialize(body);


                        //PainelSecundario.AppendText("\nJSON enviado:");
                        //PainelSecundario.AppendText(json);



                        if (response.IsSuccessStatusCode)
                        {
                            string ok = await response.Content.ReadAsStringAsync();
                            PainelSecundario.AppendText("\nAtualização realizada com sucesso!");
                            PainelSecundario.AppendText("\nResposta: " + ok);
                        }
                        else
                        {
                            string erro = await response.Content.ReadAsStringAsync();
                            PainelSecundario.AppendText("\nERRO ao atualizar ordens");
                            PainelSecundario.AppendText("\nStatus: " + (int)response.StatusCode);
                            PainelSecundario.AppendText("\nResposta: " + erro);
                        }

                        PainelSecundario.AppendText($"\n----- ATUALIZAÇÃO MES - CODIGO BARRAS -----");
                        PainelSecundario.AppendText($"\n----- ATUALIZAÇÃO     - FINALIZADO    -----");
                        //PainelSecundario.AppendText(listaOrdens);
                        PainelSecundario.ScrollToCaret();

                    }
                    if (chckbxManterOrdens.Checked == true)
                    {
                        label2.Text = "Status: Enviando peças para manter ordem de roteiro de corte para o MES ";
                        // FOREACH PASSAR PELAS ORDENS QUE SERÃO MANTIDAS
                        PainelSecundario.AppendText("\nQuantidade em pecas para manter ordem: " + pecasManterOrdem.Count());
                        PainelSecundario.ScrollToCaret();


                        var listaOrdens = pecasManterOrdem.ToList();

                        // listaOrdens = listaOrdens + "]";

                        PainelSecundario.AppendText("\n\nIniciou processamento da lista de ordens que serão mantidas.");
                        PainelSecundario.ScrollToCaret();

                        var manterordem = await api.ManterOrdensDoLoteAsync(numeroLote, pecasManterOrdem, token, "consolidação PPCP");

                        PainelSecundario.AppendText("\n\n----- MANTER ROTEIRO CO ATIVO EM ORDEMS -----");
                        PainelSecundario.AppendText($"\n----- LOTE {numeroLote} -----");
                        //PainelSecundario.AppendText(listaOrdens);
                        PainelSecundario.ScrollToCaret();

                        PainelSecundario.AppendText("\nStatusCode: " + (int)manterordem.StatusCode + " - " + manterordem.StatusCode);
                        PainelSecundario.ScrollToCaret();

                        string resposta = await manterordem.Content.ReadAsStringAsync();
                        PainelSecundario.AppendText("\nResposta da API:");
                        PainelSecundario.AppendText(resposta);
                        PainelSecundario.AppendText("\n\n--------------------------");
                        PainelSecundario.ScrollToCaret();
                        // FIM DO PROCESSO DE MANTER ORDENS
                    }
                }

                if (chckBoxEnvioEmail.Checked == true)
                    EnvioEmail();
                valorProgresso++;

                var progress = new Progress<ProgressoCopia>(p =>
                {
                    label2.Text = $"Processando {p.Processados}/{p.Total}";
                    PainelSecundario.AppendText("\n" + p.Mensagem);
                    PainelSecundario.ScrollToCaret();
                });

                if (chckbxSobreescrever.Checked == false)
                {
                    PainelSecundario.AppendText("\nCOPIANDO ARQUIVOS DE MAQUINA (SOMENTE QUE NAO EXISTE)");
                    PainelSecundario.ScrollToCaret();

                    await CopiarArquivosMaquina(progress);
                }
                else
                {
                    PainelSecundario.AppendText("\nCOPIANDO ARQUIVOS DE MAQUINA (SOBREESCREVER)");
                    PainelSecundario.ScrollToCaret();

                    await CopiarArquivosMaquinaSobreescrever(progress);
                }
                bloquearFecha = false;
                // Forma correta e simplificada

                // Ou usando a interpolação direta (mais legível):
                var agora = DateTime.Now;
                var horarioFim = $"Data: {agora:dd/MM/yyyy} - Horário: {agora:HH}h{agora:mm}m{agora:ss}s";

                PainelSecundario.AppendText($"\n\nHorário de finalização da classificação:  {horarioFim}\n");


                stopwatch.Stop();
                var tempo = String.Format("{0:00}h {1:00}m {2:00}s {3:00}mls", stopwatch.Elapsed.Hours, stopwatch.Elapsed.Minutes, stopwatch.Elapsed.Seconds, stopwatch.Elapsed.Milliseconds);


                PainelSecundario.AppendText($"\nO processo geral da classificação, separação dos planos, envio mes e cópia dos arquivos de furação e usinagens levou: {tempo}\n\nTodos os processos concluídos!.");
                PainelSecundario.ScrollToCaret();

                progressBar1.Value = progressBar1.Maximum;
                label2.Text = "Status: Todos os processos finalizados! ✅";

                var msgFinal = MessageBox.Show("Todos os processos finalizados! ✅\n\n\nDeseja abrir a pasta dos arquivos?", "FIM!", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
                if (msgFinal == DialogResult.Yes)
                {
                    var caminhoarquivo = System.IO.Path.GetDirectoryName(caminhocolado2);
                    if (Directory.Exists(caminhoarquivo))
                        Process.Start("Explorer.exe", caminhoarquivo);
                }

                // Log
                var textoLog = PainelSecundario.Text;
                File.AppendAllText(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Classificador_Log.txt", textoLog + Environment.NewLine);


            }

            catch (Exception ex)
            {
                MessageBox.Show($"{ex.Message}");
                PararProgresso();
            }




        }

        public void RelatorioMoEletrica(List<Armazenado> ListaME)
        {
            var codigosME = File.ReadAllLines(@"J:\PCP\InfoPecasPlanos\PERFIL\MO ELETRICA.txt");
            //var codigosME = File.ReadAllLines(@"C:\Users\sergi\Desktop\PCP\InfoPecasPlanos\PERFIL\MO ELETRICA.txt");
            List<string> codigosMEList = new List<string>();
            foreach (var i in codigosME)
            {
                codigosMEList.Add(i);
            }

            DataTable tableMP = new DataTable();
            tableMP.Columns.Add("NÚMERO DO PEDIDO");
            tableMP.Columns.Add("RAZãO SOCIAL CLIENTE");
            tableMP.Columns.Add("ORDEM COMPRA PEDIDO");
            tableMP.Columns.Add("NOME PLANEJADOR");
            tableMP.Columns.Add("QUANTIDADE");
            tableMP.Columns.Add("ALTURA PEÇA");
            tableMP.Columns.Add("LARGURA PEÇA");
            tableMP.Columns.Add("ESPESSURA PEÇA");
            tableMP.Columns.Add("CÓDIGO MATERIAL");
            tableMP.Columns.Add("DESCRIÇãO MATERIAL");
            tableMP.Columns.Add("LARGURA CORTE MATERIAL");
            tableMP.Columns.Add("ALTURA CORTE MATERIAL");
            tableMP.Columns.Add("IMAGEM DO MATERIAL");
            tableMP.Columns.Add("CÓDIGO PEÇA");
            tableMP.Columns.Add("COMPLEMENTO");
            tableMP.Columns.Add("DESCRIÇãO PEÇA");
            tableMP.Columns.Add("DESENHO PROGRAMADO 1");
            tableMP.Columns.Add("DESENHO PROGRAMADO 2");
            tableMP.Columns.Add("DESENHO PROGRAMADO 3");
            tableMP.Columns.Add("VEIO MATERIAL");
            tableMP.Columns.Add("BORDA SUPERIOR");
            tableMP.Columns.Add("BORDA INFERIOR");
            tableMP.Columns.Add("BORDA ESQUERDA");
            tableMP.Columns.Add("BORDA DIREITA");
            tableMP.Columns.Add("DESTINO IMPRESSãO");
            tableMP.Columns.Add("ID");
            tableMP.Columns.Add("POSTOS OPERATIVOS");
            tableMP.Columns.Add("NÚMERO DO LOTE");
            tableMP.Columns.Add("CÓDIGO CLIENTE");
            tableMP.Columns.Add("MODULO_ID");
            tableMP.Columns.Add("NÚMERO DA ORDEM");
            tableMP.Columns.Add("DATA ENTREGA LOTE");
            tableMP.Columns.Add("BOX");
            tableMP.Columns.Add("ESPECIAL");



            IEnumerable<string> listaPerfis = null;
            if (File.Exists(caminhocsvPerfil))
                listaPerfis = File.ReadLines(caminhocsvPerfil).Skip(1); // Leitura do CSV de Perfil


            var listaNovaME = new List<Armazenado>();


            foreach (var mdf in ListaME)
            {
                Armazenado mdfME = mdf;
                if (!listaNovaME.Any(p => p.NumeroOrdem == mdfME.NumeroOrdem))
                    listaNovaME.Add(mdfME);
                if (listaPerfis != null)
                {
                    foreach (var perfil in listaPerfis)
                    {
                        Armazenado perfilME = perfil;
                        if (perfilME.Modulo == mdfME.Modulo && codigosMEList.Contains(perfilME.CodigoMaterial)) // verifica se é o mesmo modulo e verifica se o perfil está na lista de codigo definidos que vai para o ME
                        {
                            if (!listaNovaME.Any(p => p.NumeroOrdem == perfilME.NumeroOrdem))
                                listaNovaME.Add(perfilME);
                        }
                    }
                }

            }

            var listaOrdenada = listaNovaME.OrderBy(peca => peca.ERP)
                .ThenBy(peca => peca.Planejador)
                .ThenBy(peca => peca.DescricaoPeca)
                .ToList();

            listaOrdenada.ForEach(pecas =>
            {
                tableMP.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
            });

            using (XLWorkbook xlsRelatorioMP = new XLWorkbook())
            {
                xlsRelatorioMP.AddWorksheet(tableMP, "Sheet");
                using (MemoryStream ns = new MemoryStream())
                {
                    xlsRelatorioMP.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\ME.xlsx");
                }

            }

        }
        public void RelatorioMoPerfil(List<Armazenado> ListaMP)
        {           
            var codigosMP = File.ReadAllLines(@"J:\PCP\InfoPecasPlanos\PERFIL\MO PERFIL.txt");
            //var codigosMP = File.ReadAllLines(@"C:\Users\sergi\Desktop\PCP\InfoPecasPlanos\PERFIL\MO PERFIL.txt");
            List<string> codigosMPList = new List<string>();
            foreach (var i in codigosMP)
            {
                codigosMPList.Add(i);
            }

            DataTable tableMP = new DataTable();
            tableMP.Columns.Add("NÚMERO DO PEDIDO");
            tableMP.Columns.Add("RAZãO SOCIAL CLIENTE");
            tableMP.Columns.Add("ORDEM COMPRA PEDIDO");
            tableMP.Columns.Add("NOME PLANEJADOR");
            tableMP.Columns.Add("QUANTIDADE");
            tableMP.Columns.Add("ALTURA PEÇA");
            tableMP.Columns.Add("LARGURA PEÇA");
            tableMP.Columns.Add("ESPESSURA PEÇA");
            tableMP.Columns.Add("CÓDIGO MATERIAL");
            tableMP.Columns.Add("DESCRIÇãO MATERIAL");
            tableMP.Columns.Add("LARGURA CORTE MATERIAL");
            tableMP.Columns.Add("ALTURA CORTE MATERIAL");
            tableMP.Columns.Add("IMAGEM DO MATERIAL");
            tableMP.Columns.Add("CÓDIGO PEÇA");
            tableMP.Columns.Add("COMPLEMENTO");
            tableMP.Columns.Add("DESCRIÇãO PEÇA");
            tableMP.Columns.Add("DESENHO PROGRAMADO 1");
            tableMP.Columns.Add("DESENHO PROGRAMADO 2");
            tableMP.Columns.Add("DESENHO PROGRAMADO 3");
            tableMP.Columns.Add("VEIO MATERIAL");
            tableMP.Columns.Add("BORDA SUPERIOR");
            tableMP.Columns.Add("BORDA INFERIOR");
            tableMP.Columns.Add("BORDA ESQUERDA");
            tableMP.Columns.Add("BORDA DIREITA");
            tableMP.Columns.Add("DESTINO IMPRESSãO");
            tableMP.Columns.Add("ID");
            tableMP.Columns.Add("POSTOS OPERATIVOS");
            tableMP.Columns.Add("NÚMERO DO LOTE");
            tableMP.Columns.Add("CÓDIGO CLIENTE");
            tableMP.Columns.Add("MODULO_ID");
            tableMP.Columns.Add("NÚMERO DA ORDEM");
            tableMP.Columns.Add("DATA ENTREGA LOTE");
            tableMP.Columns.Add("BOX");
            tableMP.Columns.Add("ESPECIAL");




            var listSotilles = new List<string>();

            var listaNovaMP = new List<Armazenado>();

            IEnumerable<string> listaPerfis = null;
            if (File.Exists(caminhocsvPerfil))
                listaPerfis = File.ReadLines(caminhocsvPerfil).Skip(1); // Leitura do CSV de Perfil

            foreach (var lineMP in ListaMP)                      // Passa pela Lista de MP para verificar o que de perfil pertence a cada peça
            {
                int contagemSotille = 0;
                Armazenado armMP = lineMP;                       // Cria um objeto do tipo Armazenado para cada linha da ListaMP, facilitando o acesso às propriedades de cada peça
                
                // Adiciona MDF se não estiver na lista
                if (!listaNovaMP.Any(p => p.NumeroOrdem == armMP.NumeroOrdem))
                    listaNovaMP.Add(armMP);


                if ((armMP.DescricaoPeca.Contains("SOTILLE") || armMP.DescricaoPeca.Contains("SAPATEIRA") )&& !listSotilles.Contains(armMP.Modulo + armMP.ERP)) // Caso seja porta sotille ou sapateira e já nao foi feita a contagem
                {
                    listSotilles.Add(armMP.Modulo + armMP.ERP); // adiciona na lista de sotille o modulo + erp para nao refazer a contagem
                    foreach (var item in ListaMP) // percorre novamente a lista de mdf mp
                    {
                        if(item.Modulo== armMP.Modulo && item.ERP == armMP.ERP) // filtra para ver se é o mesmo modulo e mesmo erp
                        {
                            
                            if (item.DescricaoPeca.Contains(".DUP")) // Verifica se é porta com sotille duplo
                                contagemSotille += 2; // faz contagem acrescida de 2 para portas com sotille duplo
                            else
                                contagemSotille += 1;// faz contagem acrescida de 1 para portas com sotille simples
                            //PainelSecundario.AppendText($"Amb: {armMP.PedidoAmbiente} - Modulo-ERP: {armMP.Modulo}-{armMP.ERP} = add contagem: {contagemSotille}\n");
                            //PainelSecundario.ScrollToCaret();
                        }
                    }
                }
                
                if (listaPerfis != null)
                {
                    // Passa pela lista de perfis
                    foreach (var linePerfil in listaPerfis)          // Para cada linha do csv de perfil, verifica se o perfil pertence à peça da ListaMP
                    {
                        Armazenado armPerfil = linePerfil;           // Cria um objeto do tipo Armazenado para cada linha da lista de perfis, facilitando o acesso às propriedades de cada perfil
                        bool contemPecaMdf = false;
                        bool contemPecaPerfil = false;

                        foreach (var item in listaNovaMP)    // Verifica se na nova lista MP tem o ordem de corte do MDF e do perfil para não adicionar duplicado
                        {
                            if (item.NumeroOrdem == armMP.NumeroOrdem)       // Se o número da ordem do item da nova lista MP for igual ao número da ordem do MDF, então a variável contemPecaMdf recebe true
                                contemPecaMdf = true;
                            if (item.NumeroOrdem == armPerfil.NumeroOrdem)   // Se o número da ordem do item da nova lista MP for igual ao número da ordem do perfil, então a variável contemPecaPerfil recebe true
                                contemPecaPerfil = true;

                        }

                        // Para cada perfil que pertence ao MDF
                        if (armPerfil.Modulo == armMP.Modulo &&
                            armPerfil.ERP == armMP.ERP &&
                            codigosMPList.Contains(armPerfil.CodigoMaterial))
                        {                           
                            // Adiciona Cantoneira se atender à regra
                            if (armPerfil.DescricaoPeca.Contains("CANTONEIRA") &&
                                !armPerfil.DescricaoPeca.Contains("95,7") &&
                                !armPerfil.DescricaoPeca.Contains("NATURAL") &&
                                !listaNovaMP.Any(p => p.NumeroOrdem == armPerfil.NumeroOrdem))
                            {
                                listaNovaMP.Add(armPerfil);
                            }
                            else if (armPerfil.DescricaoPeca.Contains("TUBO ESTRIADO") && !armMP.DescricaoPeca.Contains("DISPENSER"))
                            {
                            }
                            // Se for SOTILLE
                            else if (armPerfil.DescricaoPeca.Contains("SOTILLE") && contagemSotille > 0)
                            {                               
                                // Priorizar SOTILLE CUSTOMIZADO
                                foreach (var item in listaPerfis)
                                {
                                    Armazenado armCustom = item;
                                    if (armCustom.DescricaoPeca.Contains("SOTILLE CUSTOMIZADO") &&
                                        armCustom.Modulo == armMP.Modulo &&
                                        armCustom.ERP == armMP.ERP &&
                                        !listaNovaMP.Any(p => p.NumeroOrdem == armCustom.NumeroOrdem))
                                    {
                                        listaNovaMP.Add(item);
                                        contagemSotille--;

                                    }
                                }

                                // Adiciona o SOTILLE normal se ainda não estiver na lista
                                if (!listaNovaMP.Any(p => p.NumeroOrdem == armPerfil.NumeroOrdem))
                                {
                                    listaNovaMP.Add(armPerfil);
                                    contagemSotille--;
                                }
                            }
                            // Perfis normais (não Cantoneira e não SOTILLE)
                            else if (!armPerfil.DescricaoPeca.Contains("SOTILLE") &&
                                     !armPerfil.DescricaoPeca.Contains("CANTONEIRA") &&
                                     !listaNovaMP.Any(p => p.NumeroOrdem == armPerfil.NumeroOrdem))
                            {
                                listaNovaMP.Add(armPerfil);
                            }
                        }

                        // Painel Canaletado e Perfis canaletados
                        if (armMP.DescricaoPeca.Contains("CANALETADO") && contemPecaMdf == false)
                        {
                            listaNovaMP.Add(armMP);
                        }
                        if (armPerfil.DescricaoPeca.Contains("CANALETADO") && contemPecaPerfil == false)
                        {
                            listaNovaMP.Add(armPerfil);
                        }
                    }
                }
                
            }
            var listaOrdenada = listaNovaMP.OrderBy(peca => peca.ERP)
                .ThenBy(peca => peca.Planejador)
                .ThenBy(peca => peca.DescricaoPeca)
                .ToList();

            listaOrdenada.ForEach(pecas =>
            {
                tableMP.Rows.Add(pecas.ERP, pecas.RazaoSocial, pecas.PedidoAmbiente, pecas.Planejador, pecas.Quantidade, pecas.Altura, pecas.Largura, pecas.Espessura, pecas.CodigoMaterial, pecas.DescricaoMaterial, pecas.LarguraCorte, pecas.AlturaCorte, pecas.ImagemMaterial, pecas.CodigoPeca, pecas.Complemento, pecas.DescricaoPeca, pecas.DesenhoUm, pecas.DesenhoDois, pecas.DesenhoTres, pecas.VeioMaterial, pecas.BordaSup, pecas.BordaInf, pecas.BordaEsq, pecas.BordaDir, pecas.DestinoImpressao, pecas.CodigoBarras, pecas.PostosOperativos, pecas.NumeroLote, pecas.CodigoCliente, pecas.Modulo, pecas.NumeroOrdem, pecas.DataEntrega, pecas.Plano, pecas.Especial);
            });

            using (XLWorkbook xlsRelatorioMP = new XLWorkbook())
            {
                xlsRelatorioMP.AddWorksheet(tableMP, "Sheet");
                using (MemoryStream ns = new MemoryStream())
                {
                    xlsRelatorioMP.SaveAs(caminhocolado2.Substring(0, caminhocolado2.LastIndexOf(@"\")) + @"\Relatorios de peças\MP.xlsx");
                }

            }
        }


        public class ProgressoCopia
        {
            public int Processados { get; set; }
            public int Total { get; set; }
            public string Mensagem { get; set; }
        }

        public async Task CopiarArquivosMaquinaSobreescrever(IProgress<ProgressoCopia> progress)
        {
            var loteCru = File.ReadLines(caminhocolado2).Skip(1).ToList();

            int total = loteCru.Count;
            int processados = 0;
            int arquivosCopiados = 0;

            var sw = Stopwatch.StartNew();

            await Parallel.ForEachAsync(
                loteCru,
                new ParallelOptions { MaxDegreeOfParallelism = 4 },
                async (linha, ct) =>
                {
                    try
                    {
                        Armazenado armazenar = linha;


                        // ========== BPP ==========
                        if (armazenar.PostosOperativos.Contains("USI"))
                        {
                            File.Copy(
                                $@"\\192.168.0.10\mcm\PromobERP\BancoBuilder\Programas\Biesse\Rover_A4\Maquina1\Default\PAI\{armazenar.DesenhoUm}.bpp",
                                $@"\\192.168.0.10\erp\BIESSE\{armazenar.CodigoBarras}.bpp",
                                true
                            );

                            Interlocked.Increment(ref arquivosCopiados);
                        }

                        // ========== TLF ==========
                        if (armazenar.PostosOperativos.Contains("USI"))
                        {
                            File.Copy(
                                $@"\\192.168.0.10\mcm\PromobERP\BancoBuilder\Programas\Masterwood\Project_450\Maquina1\Default\PAI\{armazenar.DesenhoUm}.tlf",
                                $@"\\192.168.0.10\erp\MASTERWORK\{armazenar.CodigoBarras}.tlf",
                                true
                            );

                            Interlocked.Increment(ref arquivosCopiados);
                        }

                        if (armazenar.PostosOperativos.Contains("FU") || armazenar.PostosOperativos.Contains("USI"))
                        {
                            var destScx = $@"P:\NANXING\{armazenar.CodigoBarras}.scx";

                            File.Copy(
                                $@"\\192.168.0.10\mcm\PromobERP\BancoBuilder\Programas\Nanxing\NCB612\Machine1\Setup\PAI\{armazenar.DesenhoUm}.scx",
                                destScx,
                                true
                            );

                            var scxContent = await File.ReadAllTextAsync(destScx);
                            scxContent = scxContent.Replace(
                                $"ID=\"{armazenar.DesenhoUm}\"",
                                $"ID=\"{armazenar.CodigoBarras}\""
                            );
                            await File.WriteAllTextAsync(destScx, scxContent);

                            Interlocked.Increment(ref arquivosCopiados);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogErro(ex, linha);
                    }
                    finally
                    {
                        int atual = Interlocked.Increment(ref processados);

                        // Atualiza UI a cada 10
                        if (atual % 10 == 0 || atual == total)
                        {
                            progress?.Report(new ProgressoCopia
                            {
                                Processados = atual,
                                Total = total,
                                Mensagem = $"Processando {atual}/{total}"
                            });
                        }
                    }
                }
            );


            sw.Stop();
            progress?.Report(new ProgressoCopia
            {
                Processados = total,
                Total = total,
                Mensagem = $"Finalizado o processo de copia em {sw.Elapsed}"
            });
        }


        public async Task CopiarArquivosMaquina(IProgress<ProgressoCopia> progress)
        {
            var loteCru = File.ReadLines(caminhocolado2).Skip(1).ToList();

            int total = loteCru.Count;
            int processados = 0;
            int arquivosCopiados = 0;

            var sw = Stopwatch.StartNew();

            await Parallel.ForEachAsync(
                loteCru,
                new ParallelOptions { MaxDegreeOfParallelism = 4 },
                async (linha, ct) =>
                {
                    try
                    {
                        Armazenado armazenar = linha;

                        // ========= BPP =========
                        if (armazenar.PostosOperativos.Contains("USI"))
                        {
                            var origemBpp =
                                $@"\\192.168.0.10\mcm\PromobERP\BancoBuilder\Programas\Biesse\Rover_A4\Maquina1\Default\PAI\{armazenar.DesenhoUm}.bpp";

                            var destinoBpp =
                                $@"\\192.168.0.10\erp\BIESSE\{armazenar.CodigoBarras}.bpp";

                            if (CopiarArquivoDefensivo(origemBpp, destinoBpp))
                            {
                                Interlocked.Increment(ref arquivosCopiados);
                            }
                        }

                        // ========= TLF =========
                        if (armazenar.PostosOperativos.Contains("USI"))
                        {
                            var origemTlf =
                                $@"\\192.168.0.10\mcm\PromobERP\BancoBuilder\Programas\Masterwood\Project_450\Maquina1\Default\PAI\{armazenar.DesenhoUm}.tlf";

                            var destinoTlf =
                                $@"\\192.168.0.10\erp\MASTERWORK\{armazenar.CodigoBarras}.tlf";

                            if (CopiarArquivoDefensivo(origemTlf, destinoTlf))
                            {
                                Interlocked.Increment(ref arquivosCopiados);
                            }
                        }

                        // ========= SCX =========
                        if (armazenar.PostosOperativos.Contains("FU") ||
                            armazenar.PostosOperativos.Contains("USI"))
                        {
                            var origemScx =
                                $@"\\192.168.0.10\mcm\PromobERP\BancoBuilder\Programas\Nanxing\NCB612\Machine1\Setup\PAI\{armazenar.DesenhoUm}.scx";

                            var destinoScx =
                                $@"P:\NANXING\{armazenar.CodigoBarras}.scx";

                            if (await CopiarScxDefensivoAsync(
                                    origemScx,
                                    destinoScx,
                                    armazenar.DesenhoUm,
                                    armazenar.CodigoBarras))
                            {
                                Interlocked.Increment(ref arquivosCopiados);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        // Só erros realmente inesperados caem aqui
                        LogErro(ex, linha);
                    }
                    finally
                    {
                        int atual = Interlocked.Increment(ref processados);

                        if (atual % 10 == 0 || atual == total)
                        {
                            progress?.Report(new ProgressoCopia
                            {
                                Processados = atual,
                                Total = total,
                                Mensagem = $"Processando {atual}/{total}"
                            });
                        }
                    }
                }
            );

            sw.Stop();

            progress?.Report(new ProgressoCopia
            {
                Processados = total,
                Total = total,
                Mensagem = $"Finalizado o processo de copia em {sw.Elapsed}"
            });
        }

        private bool CopiarArquivoDefensivo(string origem, string destino)
        {
            try
            {
                if (!File.Exists(origem))
                    return false;

                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destino)!);

                File.Copy(origem, destino, false);
                return true;
            }
            catch (IOException)
            {
                // Arquivo já existia ou estava em uso
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }

        private async Task<bool> CopiarScxDefensivoAsync(
        string origem,
        string destino,
        string idOriginal,
        string idNovo)
        {
            try
            {
                if (!File.Exists(origem))
                    return false;

                Directory.CreateDirectory(System.IO.Path.GetDirectoryName(destino)!);

                File.Copy(origem, destino, false);

                var conteudo = await File.ReadAllTextAsync(destino);
                conteudo = conteudo.Replace(
                    $"ID=\"{idOriginal}\"",
                    $"ID=\"{idNovo}\""
                );

                await File.WriteAllTextAsync(destino, conteudo);
                return true;
            }
            catch (IOException)
            {
                return false;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }


        private void LogErro(Exception ex, string linha)
        {
            try
            {
                Armazenado armazenar = linha;
                string logPath = @"C:\Logs\erro_maquinas.log";

                lock (_logLock)
                {
                    if (!Directory.Exists(@"C:\Logs"))
                        Directory.CreateDirectory(@"C:\Logs");

                    File.AppendAllText(
                        logPath,
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | " +
                        $"{armazenar.NumeroOrdem} | " +
                        $"{armazenar.CodigoBarras} | " +
                        $"{ex.Message}{Environment.NewLine}"
                    );
                }
            }
            catch
            {
                // nunca deixa log quebrar o processo
            }
        }

        public async void PecasSemInfo(string caminho)
        {

            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nLendo os arquivos .txt com informações dos códigos.");
            PainelSecundario.ScrollToCaret();

            var arqPL_PINFOLHA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHA.txt");         // 01
            var arqPL_PINFOLHAMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHAMO.txt");       // 01
            var arqPL_RIPA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_RIPA.txt");             // 02
            var arqPL_TAPECARIA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_TAPECARIA.txt");        // 03
            var arqPL_MO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MO.txt");               // 04
            var arqPL_PORTAFRENTEMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTEMO.txt");    // 04
            var arqPL_MATPRIMA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MATPRIMA.txt");         // 05
            var arqPL_PORTAFRENTE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTE.txt");      // 06
            var arqPL18MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_18MM.txt");             // 07
            var arqPL25MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_25MM.txt");             // 08
            var arqPLOUTROS = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_OUTROS.txt");           // 09
            var arqPLIMPRIMIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_IMPRIMIR.txt");         // 10
            var arqPLMAOAMG = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MAOAMG.txt");           // --
            var arqPLESTRODAPE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_ESTRODAPE.txt");        // --
            var arqPLEXCLUIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_EXCLUIR.txt");          // --




            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nCriando as listas de informações dos códigos.");
            PainelSecundario.ScrollToCaret();

            List<string> plpinfolhatxtList = new List<string>() { };  // 01
            List<string> plpinfolhamotxtList = new List<string>() { };  // 01
            List<string> plripatxtList = new List<string>() { };  // 02
            List<string> pltapecariatxtList = new List<string>() { };  // 03
            List<string> plmotxtList = new List<string>() { };  // 04
            List<string> plportafrentemotxtList = new List<string>() { };  // 04
            List<string> plmatprimatxtList = new List<string>() { };  // 05
            List<string> plportafrentetxtList = new List<string>() { };  // 06
            List<string> pl18mmtxtList = new List<string>() { };  // 07
            List<string> pl25mmtxtList = new List<string>() { };  // 08
            List<string> ploutrostxtList = new List<string>() { };  // 09    
            List<string> plimprimirtxtList = new List<string>() { };  // 10
            List<string> plmaoamgtxtList = new List<string>() { };  // --
            List<string> plestrodapetxtList = new List<string>() { };  // --
            List<string> plexcluirtxtList = new List<string>() { };  // --





            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nInserindo itens do .txt para listas.");
            PainelSecundario.ScrollToCaret();

            // 01 - PL_PINFOLHA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPINFOLHA in arqPL_PINFOLHA)
                plpinfolhatxtList.Add(itemPLPINFOLHA);

            // 01 - PL_PINFOLHAMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPINFOLHAMO in arqPL_PINFOLHAMO)
                plpinfolhamotxtList.Add(itemPLPINFOLHAMO);

            // 02 - PL_RIPA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL_RIPA in arqPL_RIPA)
                plripatxtList.Add(itemPL_RIPA);

            // 03 - PL_TAPECARIA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLTAPECARIA in arqPL_TAPECARIA)
                pltapecariatxtList.Add(itemPLTAPECARIA);

            // 04 - PL_MO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMO in arqPL_MO)
                plmotxtList.Add(itemPLMO);

            // 04 - PL_PORTAFRENTEMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPORTAFRENTEMO in arqPL_PORTAFRENTEMO)
                plportafrentemotxtList.Add(itemPLPORTAFRENTEMO);

            // 05 - PL_MATPRIMA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMATPRIMA in arqPL_MATPRIMA)
                plmatprimatxtList.Add(itemPLMATPRIMA);

            // 06 - PL_PORTAFRENTE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLPORTAFRENTE in arqPL_PORTAFRENTE)
                plportafrentetxtList.Add(itemPLPORTAFRENTE);

            // 07 - PL_18MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL18MM in arqPL18MM)
                pl18mmtxtList.Add(itemPL18MM);

            // 08 - PL_25MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPL25MM in arqPL25MM)
                pl25mmtxtList.Add(itemPL25MM);

            // 09 - PL_OUTROS - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLOUTROS in arqPLOUTROS)
                ploutrostxtList.Add(itemPLOUTROS);

            // 10 - PL_IMPRIMIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLIMPRIMIR in arqPLIMPRIMIR)
                plimprimirtxtList.Add(itemPLIMPRIMIR);

            // -- - PL_MAOAMG - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLMAOAMG in arqPLMAOAMG)
                plmaoamgtxtList.Add(itemPLMAOAMG);

            // -- - PL_ESTRODAPE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLESTRODAPE in arqPLESTRODAPE)
                plestrodapetxtList.Add(itemPLESTRODAPE);

            // -- - PL_EXCLUIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
            foreach (var itemPLEXCLUIR in arqPLEXCLUIR)
                plexcluirtxtList.Add(itemPLEXCLUIR);

            var loteCru = File.ReadLines(caminhocolado2).Skip(1);
            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nLendo o CSV");
            PainelSecundario.ScrollToCaret();

            var itensSemInfoList = new List<string> { }; // ITENS SEM INFORMAÇÃO DE TIPO DE PLANO
            var itemencontrado = false;
            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
            PainelSecundario.AppendText("\n\nContando as peças sem informação de plano.");
            PainelSecundario.ScrollToCaret();
            itensSemInfoList.Clear();
            foreach (var linha in loteCru)
            {


                Armazenado arm = linha;
                if ((plpinfolhatxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plpinfolhamotxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plripatxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || pltapecariatxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plmotxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plportafrentemotxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plmatprimatxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plportafrentetxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || pl18mmtxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || pl25mmtxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || ploutrostxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plimprimirtxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plmaoamgtxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plestrodapetxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)
                    || plexcluirtxtList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura)))
                {
                }
                else if (!itensSemInfoList.Contains(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura))
                    itensSemInfoList.Add(arm.CodigoPeca.Substring(0, arm.CodigoPeca.IndexOf(".")) + "." + arm.Espessura);
            }
            if (primeiroValorProgresso == 0)
                progressBar1.Maximum = int.Parse(itensSemInfoList.Count().ToString());
            primeiroValorProgresso = int.Parse(itensSemInfoList.Count().ToString());

            // Se houver algum item sem informação ele irá gerar a pesquisa e informará
            if (itensSemInfoList.Count() > 0)
            {

                // Correção espessura
                PainelSecundario.AppendText($"\nQntd item: {itensSemInfoList.Count()}\nCorrigindo espessura...");
                PainelSecundario.ScrollToCaret();

                foreach (var lin in loteCru)
                {
                    Armazenado armazenar = lin;

                    if (itemencontrado == false)
                    {

                        if (armazenar.CodigoMaterial.Length == 5)
                            armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 2);
                        if (armazenar.CodigoMaterial.Length == 4)
                            armazenar.Espessura = armazenar.CodigoMaterial.Substring(0, 1);



                        if ((plpinfolhatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plpinfolhamotxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plripatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || pltapecariatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plmotxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plportafrentemotxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plmatprimatxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plportafrentetxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || pl18mmtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || pl25mmtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || ploutrostxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plimprimirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plmaoamgtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plestrodapetxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura)
                            || plexcluirtxtList.Contains(armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura))) { }

                        else if (armazenar.CodigoPeca.Substring(0, 3) != "CMP")
                        {
                            // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                            PainelSecundario.AppendText("\n\nItem sem informação de plano encontrado.");
                            PainelSecundario.ScrollToCaret();
                            itemencontrado = true;
                            PainelSecundario.AppendText(
                                $"\nCódigo Item: '{armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura}'\n" +
                                $"Peça: {armazenar.DescricaoPeca} \nMódulo ID: {armazenar.Modulo}\nPostos Operativos: {armazenar.PostosOperativos}\n\n" +
                                $"Este item não existe em nenhuma lista.\n" +
                                $"\nHá {itensSemInfoList.Count()} itens sem informação de plano.\n" +
                                $"Por favor, selecione em qual lista deseja adicionar abaixo e clique em 'OK'\n"
                                );
                            PainelSecundario.ScrollToCaret();


                            codigoUltimaLeitura = armazenar.CodigoPeca.Substring(0, armazenar.CodigoPeca.IndexOf(".")) + "." + armazenar.Espessura;

                            valorProgresso++;
                            progressBar1.Value = valorProgresso;

                        }

                    }
                }
            }
            await Start(caminhocolado2);
        }



        private void richTextBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private async void button5_Click(object sender, EventArgs e)
        {
            if (!File.Exists(caminhocolado2))
            {
                MessageBox.Show("Nenhum arquivo .csv selecionado.", "Abra um CSV", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (checkBox1.Checked == false &&
                chckbxPinFolha.Checked == false &&
                chckbxRipa.Checked == false &&
                chckbx18MM.Checked == false &&
                chckbx25MM.Checked == false &&
                chckbxOutros.Checked == false &&
                chckbxNest15MM.Checked == false &&
                chckbxImprimir.Checked == false &&
                chckbxExcluir.Checked == false &&
                chckbxNest12MM.Checked == false &&
                chckbxNest18MM.Checked == false &&
                chckbxNest25MM.Checked == false)
            {
                MessageBox.Show("Nenhuma Flag selecionada. Selecione o tipo de separação que deseja.", "Selecionar plano de separação", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (codigoUltimaLeitura == "")
            {
                PainelSecundario.AppendText("Nenhum código para gravar encontrado.");
                PainelSecundario.ScrollToCaret();
                await Start(caminhocolado2);
            }
            else
            {
                label2.Text = "Status: Peça(s) pendente(s) para classificação.";

                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                PainelSecundario.AppendText("\n\nLendo os arquivos .txt com informações dos códigos.");
                PainelSecundario.ScrollToCaret();

                var arqPL_PINFOLHA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHA.txt");         // 01
                var arqPL_PINFOLHAMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PINFOLHAMO.txt");       // 01
                var arqPL_RIPA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_RIPA.txt");             // 02
                var arqPL_TAPECARIA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_TAPECARIA.txt");        // 03
                var arqPL_MO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MO.txt");               // 04
                var arqPL_PORTAFRENTEMO = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTEMO.txt");    // 04
                var arqPL_MATPRIMA = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MATPRIMA.txt");         // 05
                var arqPL_PORTAFRENTE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_PORTAFRENTE.txt");      // 06
                var arqPL18MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_18MM.txt");             // 07
                var arqPL25MM = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_25MM.txt");             // 08
                var arqPLOUTROS = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_OUTROS.txt");           // 09
                var arqPLIMPRIMIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_IMPRIMIR.txt");         // 10
                var arqPLMAOAMG = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_MAOAMG.txt");           // --
                var arqPLESTRODAPE = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_ESTRODAPE.txt");        // --
                var arqPLEXCLUIR = File.ReadAllLines(caminhoPadraoInfoPecas + "PL_EXCLUIR.txt");          // --




                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                PainelSecundario.AppendText("\n\nCriando as listas de informações dos códigos.");
                PainelSecundario.ScrollToCaret();

                List<string> plpinfolhatxtList = new List<string>() { };  // 01
                List<string> plpinfolhamotxtList = new List<string>() { };  // 01
                List<string> plripatxtList = new List<string>() { };  // 02
                List<string> pltapecariatxtList = new List<string>() { };  // 03
                List<string> plmotxtList = new List<string>() { };  // 04
                List<string> plportafrentemotxtList = new List<string>() { };  // 04
                List<string> plmatprimatxtList = new List<string>() { };  // 05
                List<string> plportafrentetxtList = new List<string>() { };  // 06
                List<string> pl18mmtxtList = new List<string>() { };  // 07
                List<string> pl25mmtxtList = new List<string>() { };  // 08
                List<string> ploutrostxtList = new List<string>() { };  // 09    
                List<string> plimprimirtxtList = new List<string>() { };  // 10
                List<string> plmaoamgtxtList = new List<string>() { };  // --
                List<string> plestrodapetxtList = new List<string>() { };  // --
                List<string> plexcluirtxtList = new List<string>() { };  // --


                // Descrição do que está fazendo atualmente no Painel de Visualização Secundário
                PainelSecundario.AppendText("\n\nInserindo itens do .txt para listas.");
                PainelSecundario.ScrollToCaret();

                // 01 - PL_PINFOLHA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLPINFOLHA in arqPL_PINFOLHA)
                    plpinfolhatxtList.Add(itemPLPINFOLHA);

                // 01 - PL_PINFOLHAMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLPINFOLHAMO in arqPL_PINFOLHAMO)
                    plpinfolhamotxtList.Add(itemPLPINFOLHAMO);

                // 02 - PL_RIPA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPL_RIPA in arqPL_RIPA)
                    plripatxtList.Add(itemPL_RIPA);

                // 03 - PL_TAPECARIA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLTAPECARIA in arqPL_TAPECARIA)
                    pltapecariatxtList.Add(itemPLTAPECARIA);

                // 04 - PL_MO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLMO in arqPL_MO)
                    plmotxtList.Add(itemPLMO);

                // 04 - PL_PORTAFRENTEMO - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLPORTAFRENTEMO in arqPL_PORTAFRENTEMO)
                    plportafrentemotxtList.Add(itemPLPORTAFRENTEMO);

                // 05 - PL_MATPRIMA - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLMATPRIMA in arqPL_MATPRIMA)
                    plmatprimatxtList.Add(itemPLMATPRIMA);

                // 06 - PL_PORTAFRENTE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLPORTAFRENTE in arqPL_PORTAFRENTE)
                    plportafrentetxtList.Add(itemPLPORTAFRENTE);

                // 07 - PL_18MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPL18MM in arqPL18MM)
                    pl18mmtxtList.Add(itemPL18MM);

                // 08 - PL_25MM - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPL25MM in arqPL25MM)
                    pl25mmtxtList.Add(itemPL25MM);

                // 09 - PL_OUTROS - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLOUTROS in arqPLOUTROS)
                    ploutrostxtList.Add(itemPLOUTROS);

                // 10 - PL_IMPRIMIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLIMPRIMIR in arqPLIMPRIMIR)
                    plimprimirtxtList.Add(itemPLIMPRIMIR);

                // -- - PL_MAOAMG - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLMAOAMG in arqPLMAOAMG)
                    plmaoamgtxtList.Add(itemPLMAOAMG);

                // -- - PL_ESTRODAPE - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLESTRODAPE in arqPLESTRODAPE)
                    plestrodapetxtList.Add(itemPLESTRODAPE);

                // -- - PL_EXCLUIR - LEITURA DO ARQUIVO .TXT COM OS CODIGOS E ADICIONA NA LIST<>
                foreach (var itemPLEXCLUIR in arqPLEXCLUIR)
                    plexcluirtxtList.Add(itemPLEXCLUIR);


                var escolhaPlano = comboBox1.SelectedIndex;
                switch (escolhaPlano)
                {
                    // caso escolha PL_PINFOLHA
                    case 0:
                        {

                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_PINFOLHA.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_PINFOLHA... reclassificando em... {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plpinfolhatxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }

                                arquivo.Write(codigoUltimaLeitura);

                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_PINFOLHAMO
                    case 1:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_PINFOLHAMO.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_PINFOLHAMO... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plpinfolhamotxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_RIPA
                    case 2:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_RIPA.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_RIPA... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plripatxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_MO
                    case 3:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_MO.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_MO... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plmotxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_PORTAFRENTEMO
                    case 4:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_PORTAFRENTEMO.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_PORTAFRENTEMO... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plportafrentemotxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_MATPRIMA
                    case 5:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_MATPRIMA.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_MATPRIMA... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plmatprimatxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_PORTAFRENTE
                    case 6:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_PORTAFRENTE.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {

                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_PORTAFRENTE... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plportafrentetxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_18MM
                    case 7:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_18MM.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_18MM... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in pl18mmtxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_25MM
                    case 8:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_25MM.txt"))
                            {
                                //CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_25MM... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in pl25mmtxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_OUTROS
                    case 9:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_OUTROS.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_OUTROS... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in ploutrostxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_IMPRIMIR
                    case 10:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_IMPRIMIR.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_IMPRIMIR... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plimprimirtxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_MAOAMG
                    case 11:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_MAOAMG.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_MAOAMG... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plmaoamgtxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    // CASO ESCOLHA PL_ESTRODAPE
                    case 12:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_ESTRODAPE.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_ESTRODAPE... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plestrodapetxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    case 13:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_TAPECARIA.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_TAPECARIA... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in pltapecariatxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    case 14:
                        {
                            // abertura do arquivo 
                            using (var arquivo = new StreamWriter(caminhoPadraoInfoPecas + "PL_EXCLUIR.txt"))
                            {
                                // CONTAGEM DE 3 segundos para ir para o menu
                                for (var a = 3; a >= 0; a--)
                                {
                                    PainelSecundario.AppendText($"Incluindo o item '{codigoUltimaLeitura}' em PL_EXCLUIR... Reclassificando em {a} segundos.");
                                    PainelSecundario.ScrollToCaret();
                                    Thread.Sleep(1000);
                                }
                                // Lê a lista, recoloca ela toda (porque quando abre o arquivo zera os itens do.txt) e adiciona o item novo
                                foreach (var i in plexcluirtxtList)
                                {
                                    arquivo.Write($"{i}\n");
                                }
                                arquivo.Write(codigoUltimaLeitura);
                            }
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                    default:
                        {
                            PainelSecundario.AppendText($"Nenhum valor válido selecionado. Voltando a reclassicação em 3 segundos.");
                            PainelSecundario.ScrollToCaret();
                            Thread.Sleep(3000);
                            PecasSemInfo(caminhocolado2);
                            break;
                        }
                }
            }
        }

        public void EnvioEmail()
        {
            string unicoCSV = "Não";

            if (checkBox1.Checked)
                unicoCSV = "Sim";

            TimeZoneInfo tzi = TimeZoneInfo.FindSystemTimeZoneById("America/Sao_Paulo");

            // it's a simple one-liner
            DateTime pacific = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, tzi);

            using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
            {
                //smtp.Host = "smtp.gmail.com";
                smtp.EnableSsl = true;
                smtp.Credentials = new NetworkCredential("", "");

                using (MailMessage msg = new MailMessage())
                {
                    var nomedoarquivo = System.IO.Path.GetFileNameWithoutExtension(caminhocolado2);
                    msg.From = new MailAddress("");
                    msg.To.Add(new MailAddress(""));
                    msg.Subject = $"Classificação de Peças - {nomedoarquivo} - {qntdpecasgrid} peças";
                    msg.Body = $"Olá,\n" +
                        $"Segue relatório da classificação do arquivo: {nomedoarquivo}\n\n" +
                        $"Data / Horário: {pacific}\n" +
                        $"Caminho do Arquivo: {caminhocolado2}\n" +
                        $"Quantidade total de peças: {qntdpecasgrid}\n" +
                        $"Classificação Para Unico Arquivo: {unicoCSV}\n" +
                        $"\nQuantidade de Peças em PL_PINFOLHA:     {emailqntdPLPinFolha}" +
                        $"\nQuantidade de Peças em PL_RIPA:         {emailqntdPLRipa}" +
                        $"\nQuantidade de Peças em PL_18MM:         {emailqntdPL18mm}" +
                        $"\nQuantidade de Peças em PL_25MM:         {emailqntdPL25mm}" +
                        $"\nQuantidade de Peças em PL_OUTROS:       {emailqntdPLOutros}" +
                        $"\nQuantidade de Peças em PL_IMPRIMIR:     {emailqntdPLImprimir}" +
                        $"\nQuantidade de Peças em PL_EXCLUIR:      {emailqntdPLExcluir}" +
                        $"\nQuantidade de Peças em PL_NESTING_12MM: {emailqntdPLNesting12mm}" +
                        $"\nQuantidade de Peças em PL_NESTING_18MM: {emailqntdPLNesting15mm}" +
                        $"\nQuantidade de Peças em PL_NESTING_18MM: {emailqntdPLNesting18mm}" +
                        $"\nQuantidade de Peças em PL_NESTING_25MM: {emailqntdPLNesting25mm}" +
                        $"\n" +
                        $"\nQuantidade de componentes (CMP): {emailqntdCMPLote} " +
                        $"\nQuantidade de peças com Corte (CO): {emailqntdCOLote} " +
                        $"\nQuantidade de peças com Bordo (BO): {emailqntdBOLote} " +
                        $"\nQuantidade de peças com Furação (FU): {emailqntdFULote} " +
                        $"\nQuantidade de peças com Usinagens (USI): {emailqntdUSILote} " +
                        $"\nQuantidade de peças com Conferência (CF): {emailqntdCFLote} " +
                        $"\nQuantidade de peças com Montagem (MO): {emailqntdMOLote} " +
                        $"\nQuantidade de peças com Marcenaria (MA): {emailqntdMALote} " +
                        $"\nQuantidade de peças com Montagem de Perfil (MP): {emailqntdMPLote} " +
                        $"\nQuantidade de peças com Montagem de Elétrica (ME): {emailqntdMELote} " +
                        $"\nQuantidade de peças com Pintura Líquida (PIN): {emailqntdPINLote} " +
                        $"\nQuantidade de peças com Folha (FL): {emailqntdFLLote} " +
                        $"\nQuantidade de peças com Colagem: {emailqntdColagemLote} " +
                        $"\nQuantidade de peças com Controle de Qualidade (CQ): {emailqntdCQLote} " +
                        $"\nQuantidade de peças com Filetação Manual: {emailqntdFiletacaoLote} ";

                    try
                    {
                        smtp.Send(msg);
                        MessageBox.Show("Email enviado com sucesso!");
                    }
                    catch (System.Exception ex)
                    {
                        MessageBox.Show("Ocorreu um erro ao enviar ao email: " + ex.Message);
                    }
                }

            }
        }
        private void chckbxNest15MM_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            var tempo = String.Format("{0:00}h {1:00}m {2:00}s {3:00}mls", stopwatch.Elapsed.Hours, stopwatch.Elapsed.Minutes, stopwatch.Elapsed.Seconds, stopwatch.Elapsed.Milliseconds);
            label1.Text = "Tempo: " + tempo;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void progressBar1_Click(object sender, EventArgs e)
        {

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            if (caminhocolado2 != "")
            {
                var caminhoarquivo = System.IO.Path.GetDirectoryName(caminhocolado2);
                if (Directory.Exists(caminhoarquivo))
                    Process.Start("Explorer.exe", caminhoarquivo);
                else
                {
                    MessageBox.Show("Diretório não encontrado.");
                }
            }
            else
            {
                MessageBox.Show("Diretório não encontrado.");
            }

        }

        private void chckbxPinFolha_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL PINFOHLA
        {

        }

        private void chckbxNest18MM_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL NEST18MM
        {

        }

        private void button3_Click(object sender, EventArgs e)
        {
            checkBox1.Checked = false;
            chckbxPinFolha.Checked = true;
            chckbxRipa.Checked = true;
            chckbx18MM.Checked = true;
            chckbx25MM.Checked = true;
            chckbxOutros.Checked = true;
            chckbxNest15MM.Checked = true;
            chckbxImprimir.Checked = true;
            chckbxExcluir.Checked = true;
            chckbxNest12MM.Checked = true;
            chckbxNest18MM.Checked = true;
            chckbxNest25MM.Checked = true;

        }

        private void chckbxRipa_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL RIPA
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            chckbxPinFolha.Checked = false;
            chckbxRipa.Checked = false;
            chckbx18MM.Checked = false;
            chckbx25MM.Checked = false;
            chckbxOutros.Checked = false;
            chckbxNest15MM.Checked = false;
            chckbxImprimir.Checked = false;
            chckbxExcluir.Checked = false;
            chckbxNest12MM.Checked = false;
            chckbxNest18MM.Checked = false;
            chckbxNest25MM.Checked = false;
        }

        private void chckbxOutros_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL OUTROS 3,6,12,15MM
        {

        }

        private void chckbxImprimir_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL IMPRIMIR
        {

        }

        private void chckbx25MM_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL 25MM
        {

        }

        private void chckbx18MM_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL 18MM
        {

        }

        private void chckbxExcluir_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL EXCLUIR
        {

        }

        private void chckbxNest12MM_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL NEST12MM
        {

        }
        private void chckbxNest25MM_CheckedChanged(object sender, EventArgs e) // CHECK BOX PL NEST25MM
        {

        }

        private void chckBoxEnvioEmail_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                chckbxPinFolha.Checked = false;
                chckbxRipa.Checked = false;
                chckbx18MM.Checked = false;
                chckbx25MM.Checked = false;
                chckbxOutros.Checked = false;
                chckbxNest15MM.Checked = false;
                chckbxImprimir.Checked = false;
                chckbxExcluir.Checked = false;
                chckbxNest12MM.Checked = false;
                chckbxNest18MM.Checked = false;
                chckbxNest25MM.Checked = false;
            }

        }

        public void ModoAdmin()
        {
            chckbxSobreescrever.Enabled = true;
            chckbxManterOrdens.Enabled = true;
            chckbxCorrecaoCodBarras.Enabled = true;
            chckBoxEnvioEmail.Enabled = true;
        }

        private void chckbxManterOrdens_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chckbxCorrecaoCodBarras_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void chckbxSobreescrever_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void btnAdmin_Click(object sender, EventArgs e)
        {
            Form2 login = new Form2(this);
            this.Hide();

            login.Show();

        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // Verifica se o usuário clicou no "X" ou deu Alt+F4
            if (e.CloseReason == CloseReason.UserClosing && bloquearFecha == true)
            {
                // Cancela o fechamento
                e.Cancel = true;
                MessageBox.Show("Aguarde para poder fechar.");
            }
        }

        private void lblAgMDF_Click(object sender, EventArgs e)
        {

        }

        private void btnAbrirPerfil_Click(object sender, EventArgs e)
        {
            caminhocsvPerfil = ""; // reseta o caminho perfil
            lblAgPerfil.Text = "Aguardando...";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV files (*.csv)|*.csv";

            // Optionally, set the default extension (e.g., if no filter is selected)
            ofd.DefaultExt = ".csv";
            DialogResult buscaArquivo = ofd.ShowDialog();

            ofd.RestoreDirectory = false;

            if (buscaArquivo == DialogResult.OK)
            {
                caminhocsvPerfil = ofd.FileName;
                try
                {
                    var csvPerfil = File.ReadAllLines(caminhocsvPerfil).Skip(1);
                    qntdpecasgrid = 0;

                    foreach (var linha in csvPerfil)
                    {
                        Armazenado armazenar = linha;

                        if (string.IsNullOrWhiteSpace(armazenar.Altura) || string.IsNullOrWhiteSpace(armazenar.Largura)) // Se altura ou largura estiverem vazias ou forem apenas espaços em branco, exibe uma mensagem de erro indicando o código da peça e a ordem, e para a execução do código para evitar erros posteriores
                        {
                            MessageBox.Show($"A peça com Código: {armazenar.CodigoPeca} e Ordem:{armazenar.NumeroOrdem}  possui Altura ou Largura vazia. Verifique o arquivo CSV.", "Erro de Dimensões", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ResetaPrograma();
                            return; // Para o codigo
                        }
                        if (int.TryParse(armazenar.CodigoMaterial, out _) == true) // Tenta fazer a coversão para inteiro do código de material, se não conseguir é porque tem letra ou caractere especial, o que é um código de matéria-prima inválido
                        {
                            MessageBox.Show($"A peça com Código: {armazenar.CodigoPeca} e Ordem:{armazenar.NumeroOrdem} possui código de matéria-prima inválida.", "Erro de Matéria-Prima", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ResetaPrograma();
                            return; // Para o codigo
                        }
                        if (armazenar.Planejador.Length > 6)
                        {
                            if (armazenar.Planejador.Contains("FABRICA"))
                            {
                                MessageBox.Show($"A peça com Código: {armazenar.CodigoPeca} e Ordem:{armazenar.NumeroOrdem} possui Planejador de FABRICA. Verifique o arquivo CSV.", "Erro de Planejador", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                ResetaPrograma();
                                return; // Para o codigo
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Erro no processamento: {ex.Message}");
                }

                lblAgPerfil.Text = "Perfil Carregado";
                label2.Text = "Status: CSV de Perfil Carregado.";
                csvPerfilCarregado = true;

            }
        }

        private void btnAtualizacoes_Click(object sender, EventArgs e)
        {
            var formAtt = new FormAtualizacoes();
            formAtt.Show();
        }
        public void ResetaPrograma()
        {
            dataGridView1.Rows.Clear(); // Limpa as celulas do Data Grid
            PararProgresso();

            caminhocolado2 = "";
            caminhocsvPerfil = "";
            // Reseta quantidade de cada int
            emailqntdPLPinFolha = 0;
            emailqntdPLRipa = 0;
            emailqntdPL18mm = 0;
            emailqntdPL25mm = 0;
            emailqntdPLOutros = 0;
            emailqntdPLImprimir = 0;
            emailqntdPLExcluir = 0;
            emailqntdPLNesting12mm = 0;
            emailqntdPLNesting15mm = 0;
            emailqntdPLNesting18mm = 0;
            emailqntdPLNesting25mm = 0;
            emailqntdMatNaoExiste = 0;
            emailqntdBruto = 0;
            emailqntdpecasCodigoBarrasErrado = 0;

            emailqntdPecasLote = 0;
            emailqntdCMPLote = 0;
            emailqntdCOLote = 0;
            emailqntdBOLote = 0;
            emailqntdFULote = 0;
            emailqntdUSILote = 0;
            emailqntdCFLote = 0;
            emailqntdMOLote = 0;
            emailqntdMALote = 0;
            emailqntdMPLote = 0;
            emailqntdMELote = 0;
            emailqntdPINLote = 0;
            emailqntdFLLote = 0;
            emailqntdColagemLote = 0;
            emailqntdCQLote = 0;
            emailqntdFiletacaoLote = 0;

            PREqntdPLPinFolha = 0;
            PREqntdPLRipa = 0;
            PREqntdPL18mm = 0;
            PREqntdPL25mm = 0;
            PREqntdPLOutros = 0;
            PREqntdPLImprimir = 0;
            PREqntdPLExcluir = 0;
            PREqntdPLNesting12mm = 0;
            PREqntdPLNesting15mm = 0;
            PREqntdPLNesting18mm = 0;
            PREqntdPLNesting25mm = 0;

            lblAgMDF.Text = "Aguardando...";
            lblAgPerfil.Text = "Aguardando...";

            chckbxPinFolha.Text = "A - Pint Folha: " + PREqntdPLPinFolha;
            chckbxRipa.Text = "B - Ripa: " + PREqntdPLRipa;
            chckbx18MM.Text = "C - 18MM : " + PREqntdPL18mm;
            chckbx25MM.Text = "D - 25MM: " + PREqntdPL25mm;
            chckbxOutros.Text = "E - Outros: " + PREqntdPLOutros;
            chckbxImprimir.Text = "F - Imprimir: " + PREqntdPLImprimir;
            chckbxExcluir.Text = "G - Excluir: " + PREqntdPLExcluir;
            chckbxNest12MM.Text = "NEST 12MM: " + PREqntdPLNesting12mm;
            chckbxNest18MM.Text = "NEST 18MM: " + PREqntdPLNesting18mm;
            chckbxNest25MM.Text = "NEST 25MM: " + PREqntdPLNesting25mm;

        }
    }

    public class Ordem
    {
        public int id_ordem { get; set; }
        public string campo05 { get; set; }
    }

    public class AtualizarOrdensRequest
    {
        public List<Ordem> ordens { get; set; }
    }

    public class ApiClient
    {
        private readonly HttpClient _http;

        public ApiClient(string baseUrl)
        {
            _http = new HttpClient
            {
                BaseAddress = new Uri(baseUrl),
                Timeout = TimeSpan.FromMinutes(5)
            };
        }

        // ---------------------------
        // LOGIN FUNCTION

        // ---------------------------
        public async Task<string> LoginAsync(string username, string password)
        {
            var endpoint = "api/auth/login";

            var body = new
            {
                user = username,       // <-- use o campo correto aceito pela API
                password = password
            };

            var json = JsonSerializer.Serialize(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            var response = await _http.PostAsync(endpoint, content);
            var responseContent = await response.Content.ReadAsStringAsync();

            // Console.WriteLine("Resposta do login:");
            // Console.WriteLine(responseContent);

            response.EnsureSuccessStatusCode();

            using var doc = JsonDocument.Parse(responseContent);


            // AQUI PEGAMOS APENAS O KEY
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
            string token = doc.RootElement
                            .GetProperty("retorno")
                            .GetProperty("key")
                            .GetString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.

#pragma warning disable CS8603 // Possible null reference return.
            return token;
#pragma warning restore CS8603 // Possible null reference return.
            // Console.Clear();
        }

        // ---------------------------
        // ATIVAR ORDEM
        // ---------------------------
        // public async Task<HttpResponseMessage> AtivarOrdemAsync(int orderId, string token, string motivo)
        // {
        //     var endpoint = $"api/cliente/ordem/{orderId}/ativar";

        //     var body = new
        //     {
        //         motivo = motivo
        //     };

        //     var json = JsonSerializer.Serialize(body);
        //     var content = new StringContent(json, Encoding.UTF8, "application/json");

        //     _http.DefaultRequestHeaders.Clear();
        //     _http.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
        //     _http.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", token);

        //     return await _http.PostAsync(endpoint, content);
        // }

        // ---------------------------
        // INATIVAR ORDEM
        // ---------------------------
        
        public async Task<HttpResponseMessage> ManterOrdensDoLoteAsync(
     string codigoLote,
     IEnumerable<long> ordensParaManter,
     string token,
     string motivo)
        {
            var endpoint = "api/cliente/lote/inativar-roteiros";

            var body = new
            {
                motivo = motivo,
                codigo_lote = codigoLote,
                ignorar_ordens = ordensParaManter, // nome conforme Swagger
                codigo_operacao = "CO",            // sempre fixo
                sequencia_operacao = 10             // sempre fixo
            };

            var json = JsonSerializer.Serialize(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            _http.DefaultRequestHeaders.Clear();
            _http.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
            _http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            return await _http.PostAsync(endpoint, content);
        }

        public async Task<HttpResponseMessage> AtualizarOrdensAsync(
    IEnumerable<Ordem> ordens,
    string token)
        {
            var endpoint = "api/cliente/ordem";

            var body = new AtualizarOrdensRequest
            {
                ordens = ordens.ToList()
            };

            var json = JsonSerializer.Serialize(body);
            var content = new StringContent(json, Encoding.UTF8, "application/json");

            _http.DefaultRequestHeaders.Clear();
            _http.DefaultRequestHeaders.Accept.Add(
                new MediaTypeWithQualityHeaderValue("application/json"));
            _http.DefaultRequestHeaders.Authorization =
                new AuthenticationHeaderValue("Bearer", token);

            return await _http.PutAsync(endpoint, content);
        }

    }
    public class Armazenado
    {
        public string ERP { get; set; }
        public string RazaoSocial { get; set; }
        public string PedidoAmbiente { get; set; }
        public string Planejador { get; set; }
        public string Quantidade { get; set; }
        public string Altura { get; set; }
        public string Largura { get; set; }
        public string Espessura { get; set; }
        public string CodigoMaterial { get; set; }
        public string DescricaoMaterial { get; set; }
        public string LarguraCorte { get; set; }
        public string AlturaCorte { get; set; }
        public string ImagemMaterial { get; set; }
        public string CodigoPeca { get; set; }
        public string Complemento { get; set; }
        public string DescricaoPeca { get; set; }
        public string DesenhoUm { get; set; }
        public string DesenhoDois { get; set; }
        public string DesenhoTres { get; set; }
        public string VeioMaterial { get; set; }
        public string BordaSup { get; set; }
        public string BordaInf { get; set; }
        public string BordaEsq { get; set; }
        public string BordaDir { get; set; }
        public string DestinoImpressao { get; set; }
        public string CodigoBarras { get; set; }
        public string PostosOperativos { get; set; }
        public string NumeroLote { get; set; }
        public string CodigoCliente { get; set; }
        public string Modulo { get; set; }
        public string NumeroOrdem { get; set; }
        public string DataEntrega { get; set; }
        public string Plano { get; set; }
        public string Especial { get; set; }

        public static implicit operator string(Armazenado armazenar)
            => $"{armazenar.ERP};{armazenar.RazaoSocial};{armazenar.PedidoAmbiente};{armazenar.Planejador};{armazenar.Quantidade};{armazenar.Altura};{armazenar.Largura};{armazenar.Espessura};{armazenar.CodigoMaterial};{armazenar.DescricaoMaterial};{armazenar.LarguraCorte};{armazenar.AlturaCorte};{armazenar.ImagemMaterial};{armazenar.CodigoPeca};{armazenar.Complemento};{armazenar.DescricaoPeca};{armazenar.DesenhoUm};{armazenar.DesenhoDois};{armazenar.DesenhoTres};{armazenar.VeioMaterial};{armazenar.BordaSup};{armazenar.BordaInf};{armazenar.BordaEsq};{armazenar.BordaDir};{armazenar.DestinoImpressao};{armazenar.CodigoBarras};{armazenar.PostosOperativos};{armazenar.NumeroLote};{armazenar.CodigoCliente};{armazenar.Modulo};{armazenar.NumeroOrdem};{armazenar.DataEntrega};{armazenar.Plano};{armazenar.Especial}";

        public static implicit operator Armazenado(string line)
        {
            var data = line.Split(";");
            return new Armazenado(
            data[0],
            data[1],
            data[2],
            data[3],
            data[4],
            data[5],
            data[6],
            data[7],
            data[8],
            data[9],
            data[10],
            data[11],
            data[12],
            data[13],
            data[14],
            data[15],
            data[16],
            data[17],
            data[18],
            data[19],
            data[20],
            data[21],
            data[22],
            data[23],
            data[24],
            data[25],
            data[26],
            data[27],
            data[28],
            data[29],
            data[30],
            data[31],
            data[32],
            data[33]);
        }
        public Armazenado(string erpXLS, string razaosocialXLS, string pedidoambienteXLS, string planejadorXLS, string quantidadeXLS, string alturaXLS, string larguraXLS, string espessuraXLS, string codigomaterialXLS, string descricaomaterialXLS, string larguracorteXLS, string alturacorteXLS, string imagemmaterialXLS, string codigopecaXLS, string complementoXLS, string descricaopecaXLS, string desenhoumXLS, string desenhodoisXLS, string desenhotresXLS, string veiomaterialXLS, string bordasupXLS, string bordainfXLS, string bordaesqXLS, string bordadirXLS, string destinoimpressaoXLS, string codigobarrasXLS, string postosoperativosXLS, string numeroloteXLS, string codigoclienteXLS, string moduloXLS, string numeroordemXLS, string dataentregaXLS, string planoXLS, string especialXLS)
        {
            ERP = erpXLS;
            RazaoSocial = razaosocialXLS;
            PedidoAmbiente = pedidoambienteXLS;
            Planejador = planejadorXLS;
            Quantidade = quantidadeXLS;
            Altura = alturaXLS;
            Largura = larguraXLS;
            Espessura = espessuraXLS;
            CodigoMaterial = codigomaterialXLS;
            DescricaoMaterial = descricaomaterialXLS;
            LarguraCorte = larguracorteXLS;
            AlturaCorte = alturacorteXLS;
            ImagemMaterial = imagemmaterialXLS;
            CodigoPeca = codigopecaXLS;
            Complemento = complementoXLS;
            DescricaoPeca = descricaopecaXLS;
            DesenhoUm = desenhoumXLS;
            DesenhoDois = desenhodoisXLS;
            DesenhoTres = desenhotresXLS;
            VeioMaterial = veiomaterialXLS;
            BordaSup = bordasupXLS;
            BordaInf = bordainfXLS;
            BordaEsq = bordaesqXLS;
            BordaDir = bordadirXLS;
            DestinoImpressao = destinoimpressaoXLS;
            CodigoBarras = codigobarrasXLS;
            PostosOperativos = postosoperativosXLS;
            NumeroLote = numeroloteXLS;
            CodigoCliente = codigoclienteXLS;
            Modulo = moduloXLS;
            NumeroOrdem = numeroordemXLS;
            DataEntrega = dataentregaXLS;
            Plano = planoXLS;
            Especial = especialXLS;
        }
    }
    public class Matriz_Lotes
    {
        public string dataprocesso { get; set; }
        public string nomearquivo { get; set; }
        public string semana { get; set; }
        public string lote { get; set; }

        public string qntdPecas { get; set; }
        public string qntdCMP { get; set; }
        public string porcentCMP { get; set; }
        public string qntdCO { get; set; }
        public string qntdBO { get; set; }
        public string qntdFU { get; set; }
        public string qntdUSI { get; set; }
        public string qntdCF { get; set; }
        public string qntdMO { get; set; }
        public string qntdMA { get; set; }
        public string qntdMP { get; set; }
        public string qntdME { get; set; }
        public string qntdPIN { get; set; }
        public string qntdFL { get; set; }
        public string qntdColagem { get; set; }
        public string qntdCQ { get; set; }
        public string qntdFiletacao { get; set; }

        // CSV explícito (controle total)
        public string ToCsv()
            => string.Join(";",
                dataprocesso,
                nomearquivo,
                semana,
                lote,
                qntdPecas,
                qntdCMP,
                porcentCMP,
                qntdCO,
                qntdBO,
                qntdFU,
                qntdUSI,
                qntdCF,
                qntdMO,
                qntdMA,
                qntdMP,
                qntdME,
                qntdPIN,
                qntdFL,
                qntdColagem,
                qntdCQ,
                qntdFiletacao
            );
    }
}
#pragma warning restore CA1845
