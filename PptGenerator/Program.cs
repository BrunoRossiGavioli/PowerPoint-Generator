using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;


internal class Program
{
    private static void Main(string[] args)
    {
        var basePath = @"D:\Desenvolvimento\Aprendizado\Office.Interop\PptGenerator\PptGenerator";

        var modelosPath = Path.Combine(basePath, "FileModel");
        var processadosPath = Path.Combine(basePath, "Processados");
        if (!Directory.Exists(processadosPath))
            Directory.CreateDirectory(processadosPath);
        if (!Directory.Exists(modelosPath))
            Directory.CreateDirectory(modelosPath);

        var baseFilePath = Path.Combine(modelosPath, "ModeloMonitoramento.pptx");
        var ppt = new Application
        {
            Visible = MsoTriState.msoTrue,
            WindowState = PpWindowState.ppWindowNormal
        };
        var presentation = ppt.Presentations.Open(baseFilePath, WithWindow: MsoTriState.msoCTrue);
        var customLayout = presentation.SlideMaster.CustomLayouts[1];
        var slide = presentation.Slides.AddSlide(1, customLayout);
        var table = slide.Shapes.AddTable(5, 8).Table;
        var col = 1;
        foreach (Cell cell in table.Rows[1].Cells)
        {
            cell.Shape.TextFrame.TextRange.Text = col++.ToString();
        }
        presentation.SaveAs(Path.Combine(processadosPath, $"Apresentacao_{DateTime.Now:ddMMyyyyHHmmss}"));
    }
}