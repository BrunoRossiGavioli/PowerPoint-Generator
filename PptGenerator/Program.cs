using Microsoft.Office.Interop.PowerPoint;
using PptGenerator.Extensions;
using System.Globalization;

namespace PptGenerator
{
    public class Program
    {
        private static void Main(string[] args)
        {
            Console.WriteLine("Quantos clientes vão ser adicionados?");
            if (!int.TryParse(Console.ReadLine(), out var linhas))
            {
                Console.WriteLine("Caracter inválido, encerrando aplicação...");
                return;
            }

            var presentation = PptExtensions.IniciarPowerPoint("ModeloMonitoramento");
            var slide = presentation.Slides.AdicionarSlide();
            var table = slide.Shapes.AddTable(linhas+1, 4).Table;

            Dictionary<int, string> headers = new() { { 1, "Nome" }, { 2, "Idade" }, { 3, "Renda Mensal" }, { 4, "Renda Anual" } };
            for (int i = 1; i <= table.Rows[1].Cells.Count; i++)
                table.Rows[1].Cells[i].DefinirValor(headers[i]);

            int clienteNum = 1;
            foreach (Row row in table.Rows)
            {
                if (row == table.Rows[1])
                    continue;

                Console.Clear();
                Console.WriteLine($"Cliente - {clienteNum++}\n");
                for (int i = 1; i <= row.Cells.Count; i++)
                {
                    if (i == 1)
                    {
                        Console.WriteLine("Informe o nome do cliente");
                        row.Cells[i].DefinirValor(Console.ReadLine());
                    }
                    else if (i == 2)
                    {
                        Console.WriteLine("Informe a idade do cliente");
                        row.Cells[i].DefinirValor(Console.ReadLine());
                    }
                    else if (i == 3)
                    {
                        Console.WriteLine("Informe a renda mensal do cliente");
                        if (!decimal.TryParse(Console.ReadLine(), out var valor))
                            continue;
                        row.Cells[i].DefinirValor(valor.ToString("C", CultureInfo.CurrentCulture));
                        row.Cells[i+1].DefinirValor((valor * 12).ToString("C", CultureInfo.CurrentCulture));
                    }
                }
            }

            presentation.SalvarComo();
        }
    }
}