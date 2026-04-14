using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Classificador_de_Peças
{
    public partial class FormAtualizacoes : Form
    {
        // RichTextBox já definido no designer
        private RichTextBox rtbAtualizacoes;
        private Button btnFechar;

        public FormAtualizacoes()
        {
            InitializeComponent();
            InicializarComponentes();
            CarregarAtualizacoes();
        }

        private void InicializarComponentes()
        {
            var form1 = new Form1();
            string versao = form1.versao;
            this.Text = $"{versao} Notas de Atualização";
            this.Size = new Size(500, 400);
            this.StartPosition = FormStartPosition.CenterParent;

            // RichTextBox
            rtbAtualizacoes = new RichTextBox();
            rtbAtualizacoes.Dock = DockStyle.Top;
            rtbAtualizacoes.Height = 320;
            rtbAtualizacoes.ReadOnly = true;
            rtbAtualizacoes.ScrollBars = RichTextBoxScrollBars.Vertical;
            rtbAtualizacoes.BackColor = Color.White;
            rtbAtualizacoes.Font = new Font("Consolas", 10);

            // Botão Fechar
            btnFechar = new Button();
            btnFechar.Text = "Fechar";
            btnFechar.Width = 100;
            btnFechar.Height = 30;
            btnFechar.Top = rtbAtualizacoes.Bottom + 10;
            btnFechar.Left = (this.ClientSize.Width - btnFechar.Width) / 2;
            btnFechar.Anchor = AnchorStyles.Bottom;
            btnFechar.Click += (s, e) => this.Close();

            this.Controls.Add(rtbAtualizacoes);
            this.Controls.Add(btnFechar);
        }

        private void CarregarAtualizacoes()
        {
            // Lista de atualizações
            var lista = new List<Atualizacao>
            {
                //new Atualizacao //  Versão 1.0
                //{
                    
                //},
                new Atualizacao // Versão 2.0
                {
                    Versao = "2.0",
                    Data = DateTime.Parse("2026-01-01"),
                    Mudancas = new List<string>
                    {
                        " * Mais intuitivo, acessibilidade e melhoria visual de dados.\n" +
                        " * Antes era Console.",
                    }
                },                
                new Atualizacao // Versão 
                {
                    Versao = "2.1",
                    Data = DateTime.Parse("2026-03-28"),
                    Mudancas = new List<string>
                    {
                        " * Inserção: Inserido os puxadores free e trilhos guia na listagem de colagem.\n" +
                        " * Inserção: Relatórios de Montagem de Perfil e Montagem de Elétrica.\n" +
                        " * Ajustado: Porta Passagem MPR 12mm SnowMatt - alteração automatica para 18mm e remoção da usinagem, o mesmo para a peça parceira.\n" +
                        " * Ajustado: Relatórios de Montagem de Perfil removido as portas de passagem free (12mm matéria-prima, 37mm colagem, 51mm colagem).\n" +
                        " * Ajustado: Relatórios de Montagem de Perfil removido os itens que são ME.\n" +
                        " * Ajustado Adaptação ao Sotille para que saia apenas a contagem de acordo com a quantidade de MDF, priorizando o 'CUSTOMIZADO' que é referência à sapateira.", // Placeholder para futuras atualizações.
                    }
                },
                new Atualizacao // Versão 
                {
                    Versao = "2.11",
                    Data = DateTime.Parse("2026-03-30"),
                    Mudancas = new List<string>
                    {
                        " * Ajustado: Caminho do csv de perfil exception removido, agora o usuário pode classificar sem selecionar o csv de perfil com uma confirmação de seguir o processo.\n" +
                        " * Ajustado: MP e ME não geravam nem o MDF caso não tenha o csv de perfil carregado.\n" +
                        " * Ajustado: Lista de Colagem quebrando quando não encontrava a listagem de perfis. Agora gera somente MDF caso não tenha o csv de perfil.\n" +
                        " * Melhoria: Melhoria visual das atualizações e fixado o valor mínimo para a janela.\n" +
                        " * Ajustado: Quando selecionado Unico CSV, já pré calcular a letra de cada csv que irá gerar.\n" +
                        " * Melhoria Código: Adicionado variável versao, ajustado para todas as janelas, alteração somente nesta variável reduzindo a alteração em vários form's.\n" +
                        " * Ajustado: Ajustado os Perfis duplicando na lista de colagem."

                    }                    
                },
                new Atualizacao
                {
                    Versao = "2.12",
                    Data = DateTime.Parse("2026-04-14"),
                    Mudancas = new List<string>
                    {
                        " * Ajustado: Regra no Base Calceiro, Chamfer, Half, Italian, Ocult, Sotille e Curved que não contém MP faz a inserção no Posto Operativo.\n" +
                        " * Melhoria do código: Diminuição de código para melhor performance.\n" +
                        " * Ajustado: Tubo estriado entrando no Relatório de MP quando é outro MDF puxando o mesmo módulo, inserção da regra para que o MDF contenha 'DISPENSER'.\n" +
                        " * Ajustado: Caso tente carregar csv de MDF em perfil ou vice-versa, ou caso não seja um formato válido de csv, o programa reseta e limpa tudo.\n" 



                    }
                }
                //new Atualizacao // Versão 
                //{
                //    Versao = "2.2 (Em produção)",
                //    Data = DateTime.Parse("2026-03-28"),
                //    Mudancas = new List<string>
                //    {
                //        "Relatório de Perfil (Pendencia)\n" +
                //        " --> Preciso fazer o mesmo do Sottile para o Curved, diferença que o Curved não vai na sapateira.\n"
                        

                //    }
                //},   
            };

            PreencherRichTextBox(lista);
        }
        private void rtbAtualizacoes_TextChanged(object sender, EventArgs e)
        {

        }

        private void PreencherRichTextBox(List<Atualizacao> lista)
        {
            rtbAtualizacoes.Clear();

            var ordenada = lista.OrderBy(a => a.Data).ToList();
            

            foreach (var at in ordenada)
            {
                // Título: Versão + Data
                rtbAtualizacoes.SelectionFont = new Font("Consolas", 10, FontStyle.Bold);
                rtbAtualizacoes.SelectionColor = Color.DarkBlue;
                rtbAtualizacoes.AppendText($"Versão {at.Versao} - {at.Data:dd/MM/yyyy}\n");

                // Separador
                rtbAtualizacoes.SelectionFont = new Font("Consolas", 9, FontStyle.Italic);
                rtbAtualizacoes.SelectionColor = Color.Gray;
                rtbAtualizacoes.AppendText("------------------------\n");

                // Mudanças
                rtbAtualizacoes.SelectionFont = new Font("Consolas", 9, FontStyle.Regular);
                rtbAtualizacoes.SelectionColor = Color.Black;

                foreach (var item in at.Mudancas)
                {
                    rtbAtualizacoes.AppendText($"{item}\n");
                }

                rtbAtualizacoes.AppendText("\n");
            }
        }
    }

    // Classe para manter cada atualização organizada
    public class Atualizacao
    {
        public string Versao { get; set; }
        public DateTime Data { get; set; }
        public List<string> Mudancas { get; set; } = new List<string>();
    }
}


