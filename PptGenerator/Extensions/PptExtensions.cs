using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PptGenerator.Extensions
{
    public static class PptExtensions
    {
        private static readonly string _basePath = @"D:\Desenvolvimento\Aprendizado\Office.Interop\PptGenerator\PptGenerator";
        private static readonly string _modelosPath = Path.Combine(_basePath, "FileModel");
        private static readonly string _processadosPath = Path.Combine(_basePath, "Processados");
        private static CustomLayout _customLayout = default!;

        public static Presentation IniciarPowerPoint(string modeloPpt)
        {
            if (!Directory.Exists(_processadosPath))
                Directory.CreateDirectory(_processadosPath);

            if (!Directory.Exists(_modelosPath))
                Directory.CreateDirectory(_modelosPath);

            var baseFilePath = Path.Combine(_modelosPath, modeloPpt + ".pptx");
            var ppt = new Application();
            var presentation = ppt.Presentations.Open(baseFilePath, WithWindow: MsoTriState.msoCTrue);
            presentation.Application.WindowState = PpWindowState.ppWindowMinimized;
            _customLayout = presentation.SlideMaster.CustomLayouts[1];
            return presentation;
        }

        public static void SalvarComo(this Presentation presentation)
        {
            presentation.Application.WindowState = PpWindowState.ppWindowNormal;
            presentation.SaveAs(Path.Combine(_processadosPath, $"Apresentacao_{DateTime.Now:ddMMyyyyHHmmss}"));
        }

        public static Slide AdicionarSlide(this Slides slides, int? index = null, CustomLayout? pCustomLayout = null)
        {
            return slides.AddSlide(index ?? slides.Count + 1, pCustomLayout ?? _customLayout);
        }

        public static void DefinirValor(this Cell cell, object? value)
        {
            try
            {
                cell.Shape.TextFrame.TextRange.Text = value?.ToString() ?? string.Empty;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                cell.Shape.TextFrame.TextRange.Text = string.Empty;
            }
        }
    }
}
