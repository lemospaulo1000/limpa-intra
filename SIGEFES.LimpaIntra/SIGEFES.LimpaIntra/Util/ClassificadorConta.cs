namespace SIGEFES.LimpaIntra.Util
{
    public class ClassificadorConta
    {
        public static bool EhIntra(long codigo)
        {
            return ((codigo / 10000) % 10) == 2;
        }
    }
}