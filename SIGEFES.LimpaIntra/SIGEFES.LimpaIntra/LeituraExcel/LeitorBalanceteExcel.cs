using System;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using SIGEFES.LimpaIntra.Modelos;
using SIGEFES.LimpaIntra.Servicos;
using SIGEFES.LimpaIntra.Util;

namespace SIGEFES.LimpaIntra.LeituraExcel
{
    public class LeitorBalanceteExcel
    {
        public Balancete Ler(string caminho)
        {
            Balancete balancete = new Balancete();

            using (FileStream fs = new FileStream(caminho, FileMode.Open, FileAccess.Read))
            {
                HSSFWorkbook workbook = new HSSFWorkbook(fs);

                ISheet sheet = workbook.GetSheetAt(0);

                int linha = 9;

                while (true)
                {
                    IRow row = sheet.GetRow(linha);

                    if (row == null)
                        break;

                    ICell celulaConta = row.GetCell(0);

                    if (celulaConta == null)
                        break;

                    string textoConta = celulaConta.ToString().Trim();

                    if (string.IsNullOrWhiteSpace(textoConta))
                        break;

                    ContaContabil conta = CriarConta(row, textoConta);

                    if (conta == null)
                    {
                        linha++;
                        continue;
                    }

                    balancete.Contas[conta.Codigo] = conta;

                    if (ClassificadorConta.EhIntra(conta.Codigo))
                        balancete.ContasIntra.Add(conta);

                    linha++;
                }
            }

            return balancete;
        }

        private ContaContabil CriarConta(IRow row, string textoConta)
        {
            string[] partes = textoConta.Split(new string[] { " - " }, StringSplitOptions.None);

            long codigo = 0;
            long.TryParse(partes[0].Trim(), out codigo);

            // manter apenas até subtítulo
            if (codigo % 10000 != 0)
                return null;

            string descricao = partes.Length > 1 ? partes[1].Trim() : "";

            decimal saldoInicial = LerDecimal(row.GetCell(1));
            decimal debito = LerDecimal(row.GetCell(2));
            decimal credito = LerDecimal(row.GetCell(3));
            decimal saldoAtual = LerDecimal(row.GetCell(4));

            char dc = LerChar(row.GetCell(5));

            if (dc == 'D')
                saldoAtual = -saldoAtual;

            ContaContabil conta = new ContaContabil();

            conta.Codigo = codigo;
            conta.Descricao = descricao;
            conta.SaldoInicial = saldoInicial;
            conta.Debito = debito;
            conta.Credito = credito;
            conta.SaldoAtual = saldoAtual;
            conta.DC = dc;

            return conta;
        }

        private decimal LerDecimal(ICell cell)
        {
            if (cell == null)
                return 0;

            if (cell.CellType == CellType.Numeric)
                return (decimal)cell.NumericCellValue;

            decimal valor;
            decimal.TryParse(cell.ToString(), out valor);

            return valor;
        }

        private char LerChar(ICell cell)
        {
            if (cell == null)
                return ' ';

            if (cell.CellType == CellType.String)
                return cell.StringCellValue.Trim()[0];

            if (cell.CellType == CellType.Numeric)
            {
                double v = cell.NumericCellValue;

                if (v == 68) return 'D';
                if (v == 67) return 'C';
            }

            string txt = cell.ToString().Trim();

            if (txt.Length == 0)
                return ' ';

            return txt[0];
        }
    }
}