namespace SIGEFES.LimpaIntra.Modelos
{
    public class ContaContabil
    {
        public long Codigo { get; set; }

        public string Descricao { get; set; } = "";

        public decimal SaldoInicial { get; set; }

        public decimal Debito { get; set; }

        public decimal Credito { get; set; }

        public decimal SaldoAtual { get; set; }

        public char DC { get; set; }
    }
}