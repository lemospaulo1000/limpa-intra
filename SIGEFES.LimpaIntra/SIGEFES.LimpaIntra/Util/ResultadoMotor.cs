using System.Collections.Generic;
using SIGEFES.LimpaIntra.Modelos;

namespace SIGEFES.LimpaIntra.Servicos
{
    public class ResultadoMotor
    {
        public Dictionary<long, ContaContabil> Original;
        public Dictionary<long, ContaContabil> Processado;
        public List<ContaContabil> Intra;
    }
}