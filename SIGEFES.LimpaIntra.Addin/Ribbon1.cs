using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using SIGEFES.LimpaIntra.Servicos;
using SIGEFES.LimpaIntra.Modelos;

namespace SIGEFES.LimpaIntra.Addin
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
        }

        private void BtnLimparIntra_Click(object sender, RibbonControlEventArgs e)
        {
            Excel.Application app = Globals.ThisAddIn.Application;

            bool calcAnterior = app.Calculation == Excel.XlCalculation.xlCalculationAutomatic;

            try
            {
                app.ScreenUpdating = false;
                app.DisplayAlerts = false;
                app.Calculation = Excel.XlCalculation.xlCalculationManual;
                app.EnableEvents = false;

                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Filter = "Excel (*.xls)|*.xls";

                if (dialog.ShowDialog() != DialogResult.OK)
                    return;

                string caminho = dialog.FileName;

                MotorLimpaIntra motor = new MotorLimpaIntra();

                ResultadoMotor resultado = motor.Processar(caminho);

                var original = resultado.Original;
                var processado = resultado.Processado;
                var intra = resultado.Intra;

                decimal somaIntraAtivo = intra.Where(x => (x.Codigo / 100000000) == 1).Sum(x => x.SaldoAtual);
                decimal somaIntraPassivo = intra.Where(x => (x.Codigo / 100000000) == 2).Sum(x => x.SaldoAtual);
                decimal somaIntraClasse3 = intra.Where(x => (x.Codigo / 100000000) == 3).Sum(x => x.SaldoAtual);
                decimal somaIntraClasse4 = intra.Where(x => (x.Codigo / 100000000) == 4).Sum(x => x.SaldoAtual);

                decimal ativoOriginal = original[100000000].SaldoAtual;
                decimal classe2Original = original[200000000].SaldoAtual;
                decimal classe3Original = original[300000000].SaldoAtual;
                decimal classe4Original = original[400000000].SaldoAtual;

                decimal ativoFinal = processado[100000000].SaldoAtual;
                decimal classe2Final = processado[200000000].SaldoAtual;

                Excel.Workbook wb = app.Workbooks.Add();

                Excel.Worksheet wsOriginal = wb.Worksheets[1];
                wsOriginal.Name = "Balancete_Original";

                Excel.Worksheet wsSemIntra = wb.Worksheets.Add();
                wsSemIntra.Name = "Balancete_Sem_Intra";

                Excel.Worksheet wsDiff = wb.Worksheets.Add();
                wsDiff.Name = "Diferenca";

                Excel.Worksheet wsIntra = wb.Worksheets.Add();
                wsIntra.Name = "Intra_OFSS";

                Excel.Worksheet wsResumo = wb.Worksheets.Add();
                wsResumo.Name = "Resumo_Auditoria";

                Excel.Worksheet wsResumoGrupo = wb.Worksheets.Add();
                wsResumoGrupo.Name = "Resumo_Grupos";

                EscreverCabecalho(wsOriginal);
                EscreverCabecalho(wsSemIntra);
                EscreverCabecalho(wsDiff);
                EscreverCabecalho(wsIntra);

                int l1 = 2;
                int l2 = 2;
                int l3 = 2;
                int l4 = 2;

                foreach (var c in original.Values.OrderBy(x => x.Codigo))
                    EscreverConta(wsOriginal, l1++, c);

                foreach (var c in processado.Values.OrderBy(x => x.Codigo))
                    EscreverConta(wsSemIntra, l2++, c);

                foreach (var c in original.Values.OrderBy(x => x.Codigo))
                {
                    if (processado.ContainsKey(c.Codigo))
                    {
                        var novo = processado[c.Codigo];

                        ContaContabil diff = new ContaContabil
                        {
                            Codigo = c.Codigo,
                            Descricao = c.Descricao,
                            SaldoInicial = novo.SaldoInicial - c.SaldoInicial,
                            Debito = novo.Debito - c.Debito,
                            Credito = novo.Credito - c.Credito,
                            SaldoAtual = novo.SaldoAtual - c.SaldoAtual,
                            DC = c.DC
                        };

                        EscreverConta(wsDiff, l3++, diff);
                    }
                }

                foreach (var c in intra.OrderBy(x => x.Codigo))
                    EscreverConta(wsIntra, l4++, c);

                FormatarPlanilha(wsOriginal);
                FormatarPlanilha(wsSemIntra);
                FormatarPlanilha(wsDiff);
                FormatarPlanilha(wsIntra);

                wsIntra.UsedRange.Interior.Color = 0xFFFF99;

                GerarResumo(
                    wsResumo,
                    ativoOriginal,
                    somaIntraAtivo,
                    ativoFinal,
                    classe2Original,
                    somaIntraPassivo,
                    classe2Final,
                    classe3Original,
                    somaIntraClasse3,
                    classe4Original,
                    somaIntraClasse4);

                GerarResumoGrupos(wsResumoGrupo, original, intra);

                MessageBox.Show("Processamento concluído.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erro: " + ex.Message);
            }
            finally
            {
                app.ScreenUpdating = true;
                app.DisplayAlerts = true;
                app.EnableEvents = true;

                if (calcAnterior)
                    app.Calculation = Excel.XlCalculation.xlCalculationAutomatic;
            }
        }

        private void EscreverCabecalho(Excel.Worksheet ws)
        {
            ws.Cells[1, 1] = "Conta";
            ws.Cells[1, 2] = "Descrição";
            ws.Cells[1, 3] = "Saldo Inicial";
            ws.Cells[1, 4] = "Débito";
            ws.Cells[1, 5] = "Crédito";
            ws.Cells[1, 6] = "Saldo Atual";
            ws.Cells[1, 7] = "D/C";

            ws.Range["A1:G1"].Font.Bold = true;
        }

        private void EscreverConta(Excel.Worksheet ws, int linha, ContaContabil c)
        {
            ws.Cells[linha, 1] = c.Codigo;
            ws.Cells[linha, 2] = c.Descricao;
            ws.Cells[linha, 3] = c.SaldoInicial;
            ws.Cells[linha, 4] = c.Debito;
            ws.Cells[linha, 5] = c.Credito;
            ws.Cells[linha, 6] = c.SaldoAtual;
            ws.Cells[linha, 7] = c.DC.ToString();
        }

        private void FormatarPlanilha(Excel.Worksheet ws)
        {
            Excel.Range used = ws.UsedRange;

            Excel.Range numeros = ws.Range["C2:F" + used.Rows.Count];
            numeros.NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-";

            used.Columns.AutoFit();

            ws.Application.ActiveWindow.SplitRow = 1;
            ws.Application.ActiveWindow.FreezePanes = true;

            int linhas = used.Rows.Count;

            for (int i = 2; i <= linhas; i++)
            {
                if (i % 2 == 0)
                    ws.Range["A" + i, "G" + i].Interior.Color = 0xF2F2F2;

                long codigo = Convert.ToInt64(ws.Cells[i, 1].Value);

                int nivel = NivelConta(codigo);

                if (nivel == 1 || nivel == 2)
                    ws.Range["A" + i, "G" + i].Font.Bold = true;
            }
        }

        private int NivelConta(long codigo)
        {
            if (codigo % 100000000 == 0) return 1;
            if (codigo % 10000000 == 0) return 2;
            if (codigo % 1000000 == 0) return 3;
            if (codigo % 100000 == 0) return 4;
            return 5;
        }
        private void GerarResumo(
    Excel.Worksheet ws,
    decimal ativoOriginal,
    decimal intraAtivo,
    decimal ativoFinal,
    decimal passivoOriginal,
    decimal intraPassivo,
    decimal passivoFinal,
    decimal classe3Original,
    decimal intraClasse3,
    decimal classe4Original,
    decimal intraClasse4)
        {
            decimal ativoOriginalAbs = Math.Abs(ativoOriginal);
            decimal intraAtivoAbs = Math.Abs(intraAtivo);
            decimal ativoFinalAbs = Math.Abs(ativoFinal);

            decimal passivoOriginalAbs = Math.Abs(passivoOriginal);
            decimal intraPassivoAbs = Math.Abs(intraPassivo);
            decimal passivoFinalAbs = Math.Abs(passivoFinal);

            decimal classe3OriginalAbs = Math.Abs(classe3Original);
            decimal intraClasse3Abs = Math.Abs(intraClasse3);
            decimal classe3FinalAbs = Math.Abs(classe3Original - intraClasse3);

            decimal classe4OriginalAbs = Math.Abs(classe4Original);
            decimal intraClasse4Abs = Math.Abs(intraClasse4);
            decimal classe4FinalAbs = Math.Abs(classe4Original - intraClasse4);

            ws.Cells[1, 1] = "Resumo contas INTRA-OFSS";
            ws.Range["A1:D1"].Font.Bold = true;

            ws.Cells[3, 2] = "Original";
            ws.Cells[3, 3] = "INTRA";
            ws.Cells[3, 4] = "Após ajuste";

            ws.Range["B3:D3"].Font.Bold = true;

            ws.Cells[4, 1] = "ATIVO";
            ws.Cells[4, 2] = ativoOriginalAbs;
            ws.Cells[4, 3] = intraAtivoAbs;
            ws.Cells[4, 4] = ativoFinalAbs;

            ws.Cells[5, 1] = "Passivo e Patrimônio Líquido";
            ws.Cells[5, 2] = passivoOriginalAbs;
            ws.Cells[5, 3] = intraPassivoAbs;
            ws.Cells[5, 4] = passivoFinalAbs;

            ws.Cells[6, 1] = "diferença:";
            ws.Cells[6, 3] = intraAtivoAbs - intraPassivoAbs;

            ws.Cells[6, 1].Font.Color = 255;
            ws.Cells[6, 3].Font.Color = 255;

            ws.Cells[8, 1] = "Variação Patrimonial Diminutiva";
            ws.Cells[8, 2] = classe3OriginalAbs;
            ws.Cells[8, 3] = intraClasse3Abs;
            ws.Cells[8, 4] = classe3FinalAbs;

            ws.Cells[9, 1] = "Variação Patrimonial Aumentativa";
            ws.Cells[9, 2] = classe4OriginalAbs;
            ws.Cells[9, 3] = intraClasse4Abs;
            ws.Cells[9, 4] = classe4FinalAbs;

            ws.Cells[10, 1] = "diferença:";
            ws.Cells[10, 3] = intraClasse3Abs - intraClasse4Abs;

            ws.Cells[10, 1].Font.Color = 255;
            ws.Cells[10, 3].Font.Color = 255;

            ws.Range["B4:D10"].NumberFormat =
                "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-";

            Excel.Range tabela = ws.Range["A3:D10"];
            tabela.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            ws.Columns.AutoFit();
        }
        private void GerarResumoGrupos(
    Excel.Worksheet ws,
    Dictionary<long, ContaContabil> contasOriginais,
    List<ContaContabil> contasIntra)
        {
            ws.Cells[1, 1] = "Grupo";
            ws.Cells[1, 2] = "Descrição";
            ws.Cells[1, 3] = "Saldo Original";
            ws.Cells[1, 4] = "Intra";
            ws.Cells[1, 5] = "Saldo Consolidado";

            ws.Range["A1:E1"].Font.Bold = true;

            Dictionary<long, decimal> originalPorGrupo = new Dictionary<long, decimal>();
            Dictionary<long, decimal> intraPorGrupo = new Dictionary<long, decimal>();

            foreach (var c in contasOriginais.Values)
            {
                if (c.Codigo % 10000000 != 0) continue;
                if (c.Codigo % 100000000 == 0) continue;

                int classe = (int)(c.Codigo / 100000000);
                if (classe > 4) continue;

                long grupo = c.Codigo / 10000000;
                originalPorGrupo[grupo] = c.SaldoAtual;
            }

            foreach (var c in contasIntra)
            {
                int classe = (int)(c.Codigo / 100000000);
                if (classe > 4) continue;

                long grupo = c.Codigo / 10000000;

                if (!intraPorGrupo.ContainsKey(grupo))
                    intraPorGrupo[grupo] = 0;

                intraPorGrupo[grupo] += c.SaldoAtual;
            }

            var grupos = originalPorGrupo.Keys
                .Union(intraPorGrupo.Keys)
                .OrderBy(x => x);

            int linha = 2;

            foreach (var g in grupos)
            {
                decimal original = originalPorGrupo.ContainsKey(g) ? originalPorGrupo[g] : 0;
                decimal intra = intraPorGrupo.ContainsKey(g) ? intraPorGrupo[g] : 0;

                long codigoGrupo = g * 10000000;

                string descricao = contasOriginais.ContainsKey(codigoGrupo)
                    ? contasOriginais[codigoGrupo].Descricao
                    : "";

                ws.Cells[linha, 1] = g;
                ws.Cells[linha, 2] = descricao;
                ws.Cells[linha, 3] = original;
                ws.Cells[linha, 4] = intra;
                ws.Cells[linha, 5] = original - intra;

                linha++;
            }

            ws.Range["C2:E" + (linha - 1)].NumberFormat =
                "_-* #,##0.00_-;-* #,##0.00_-;_-* \"-\"??_-;_-@_-";

            ws.Columns.AutoFit();
        }
    }
}