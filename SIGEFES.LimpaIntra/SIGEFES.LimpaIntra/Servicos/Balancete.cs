using System.Collections.Generic;
using SIGEFES.LimpaIntra.Modelos;

namespace SIGEFES.LimpaIntra.Servicos
{
    public class Balancete
    {
        public Dictionary<long, ContaContabil> Contas { get; } =
            new Dictionary<long, ContaContabil>();

        public List<ContaContabil> ContasIntra { get; } =
            new List<ContaContabil>();
    }
}