using System.Collections.Generic;

namespace SIGEFES.LimpaIntra.Util
{
    public class HierarquiaConta
    {
        public static List<long> GerarSuperiores(long codigo)
        {
            List<long> lista = new List<long>();

            long titulo = (codigo / 100000) * 100000;
            long subgrupo = (codigo / 1000000) * 1000000;
            long grupo = (codigo / 10000000) * 10000000;
            long classe = (codigo / 100000000) * 100000000;

            lista.Add(titulo);
            lista.Add(subgrupo);
            lista.Add(grupo);
            lista.Add(classe);

            return lista;
        }
    }
}