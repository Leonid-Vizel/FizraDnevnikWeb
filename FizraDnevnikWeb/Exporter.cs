using ClosedXML.Excel;
using FizraDnevnikWeb.Pages;

namespace FizraDnevnikWeb;

public sealed class Exporter
{
    public byte[] Export(IndexModel model)
    {
        using (XLWorkbook book = new XLWorkbook("pattern.xlsx"))
        {
            var sheet = book.Worksheet(3);
            if (model.AA17 != null)
            {
                sheet.Cell(17, 27).Value = model.AA17;
            }
            if (model.Group != null)
            {
                sheet.Cell(19, 27).Value = model.Group;
            }

            if (model.Height != null)
            {
                sheet.Cell(9, 78).Value = model.Height;
            }
            if (model.Weight != null)
            {
                sheet.Cell(12, 78).Value = model.Weight;
            }
            if (model.FatPercent != null)
            {
                sheet.Cell(15, 78).Value = model.FatPercent;
            }
            if (model.MusclePercent != null)
            {
                sheet.Cell(18, 78).Value = model.MusclePercent;
            }
            if (model.IMT != null)
            {
                sheet.Cell(21, 78).Value = model.IMT;
            }

            if (model.PulsePressureDate != null)
            {
                sheet.Cell(124, 50).Value = model.PulsePressureDate.Value.ToString("dd.MM.yyyy");
            }
            if (model.Pressure != null)
            {
                sheet.Cell(132, 50).Value = model.Pressure;
            }
            if (model.Pulse != null)
            {
                sheet.Cell(138, 50).Value = model.Pulse;
            }

            if (model.OrtostatDate != null)
            {
                sheet.Cell(36, 50).Value = model.OrtostatDate.Value.ToString("dd.MM.yyyy");
            }
            if (model.OrtostatLayingPulse != null)
            {
                sheet.Cell(40, 50).Value = model.OrtostatLayingPulse;
            }
            if (model.OrtostatStandingPulse != null)
            {
                sheet.Cell(44, 50).Value = model.OrtostatStandingPulse;
            }
            if (model.OrtostatResult != null)
            {
                sheet.Cell(48, 50).Value = model.OrtostatResult;
            }

            if (model.RuffieDate != null)
            {
                sheet.Cell(68, 9).Value = model.RuffieDate.Value.ToString("dd.MM.yyyy");
            }
            if (model.RuffieBefore != null)
            {
                sheet.Cell(71, 9).Value = model.RuffieBefore;
            }
            if (model.RuffieAfter != null)
            {
                sheet.Cell(74, 9).Value = model.RuffieAfter;
            }
            if (model.RuffieRelax != null)
            {
                sheet.Cell(77, 9).Value = model.RuffieRelax;
            }

            if (model.Norma1 != null)
            {
                sheet.Cell(160, 73).Value = model.Norma1;
            }
            if (model.Norma2 != null)
            {
                sheet.Cell(165, 73).Value = model.Norma2;
            }
            if (model.Norma3 != null)
            {
                sheet.Cell(170, 73).Value = model.Norma3;
            }

            using (MemoryStream memStream = new MemoryStream())
            {
                book.SaveAs(memStream);
                return memStream.ToArray();
            }
        }
    }
}
