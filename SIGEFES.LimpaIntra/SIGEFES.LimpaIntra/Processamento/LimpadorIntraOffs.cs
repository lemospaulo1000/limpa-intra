using System.Collections.Generic;
using SIGEFES.LimpaIntra.Modelos;
using SIGEFES.LimpaIntra.Servicos;
using SIGEFES.LimpaIntra.Util;

namespace SIGEFES.LimpaIntra.Processamento
{
    public class LimpadorIntraOffs
    {
        public static void Processar(Balancete balancete)
        {
            HashSet<long> processados = new HashSet<long>();

            foreach (ContaContabil conta in balancete.ContasIntra)
            {
                long codigo = conta.Codigo;

                if (processados.Contains(codigo))
                    continue;

                processados.Add(codigo);

                List<long> superiores = HierarquiaConta.GerarSuperiores(codigo);

                foreach (long sup in superiores)
                {
                    if (balancete.Contas.ContainsKey(sup))
                    {
                        ContaContabil s = balancete.Contas[sup];

                        // mesma lógica que já funcionava para SaldoAtual
                        s.SaldoInicial -= conta.SaldoInicial;
                        s.Debito -= conta.Debito;
                        s.Credito -= conta.Credito;
                        s.SaldoAtual -= conta.SaldoAtual;
                    }
                }
            }

            foreach (ContaContabil conta in balancete.ContasIntra)
            {
                if (balancete.Contas.ContainsKey(conta.Codigo))
                    balancete.Contas.Remove(conta.Codigo);
            }
        }
    }
}