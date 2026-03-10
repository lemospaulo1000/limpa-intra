using System;
using System.Windows.Forms;
using SIGEFES.LimpaIntra.LeituraExcel;
using SIGEFES.LimpaIntra.Processamento;
using SIGEFES.LimpaIntra.Modelos;
using System.Collections.Generic;
using System.Linq;

namespace LimparIntras
{
    internal class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "Excel (*.xls)|*.xls";
                dialog.Title = "Selecionar balancete";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                string caminho = dialog.FileName;

                var leitor = new LeitorBalanceteExcel();

                var balancete = leitor.Ler(caminho);

                Dictionary<long, ContaContabil> original =
                    balancete.Contas.ToDictionary(
                        x => x.Key,
                        x => new ContaContabil
                        {
                            Codigo = x.Value.Codigo,
                            Descricao = x.Value.Descricao,
                            SaldoInicial = x.Value.SaldoInicial,
                            Debito = x.Value.Debito,
                            Credito = x.Value.Credito,
                            SaldoAtual = x.Value.SaldoAtual,
                            DC = x.Value.DC
                        });

                List<ContaContabil> intra = new List<ContaContabil>(balancete.ContasIntra);

                LimpadorIntraOffs.Processar(balancete);

                Console.WriteLine("Processamento concluído.");
                Console.WriteLine("Contas processadas: " + balancete.Contas.Count);

                MessageBox.Show("Processamento concluído com sucesso.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}