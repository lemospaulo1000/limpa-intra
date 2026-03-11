using System.Collections.Generic;
using System.Linq;
using SIGEFES.LimpaIntra.LeituraExcel;
using SIGEFES.LimpaIntra.Processamento;
using SIGEFES.LimpaIntra.Modelos;

namespace SIGEFES.LimpaIntra.Servicos
{
    public class MotorLimpaIntra
    {
        public ResultadoMotor Processar(string caminho)
        {
            LeitorBalanceteExcel leitor = new LeitorBalanceteExcel();
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

            // executa algoritmo contábil
            LimpadorIntraOffs.Processar(balancete);

            return new ResultadoMotor
            {
                Original = original,
                Processado = balancete.Contas,
                Intra = intra
            };
        }
    }
}