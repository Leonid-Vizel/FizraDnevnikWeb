using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using Weasel.Export.AspNetCore;

namespace FizraDnevnikWeb.Pages;

[ValidateAntiForgeryToken]
[ResponseCache(VaryByHeader = "User-Cached", Duration = 60)]
public sealed class IndexModel : PageModel
{
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Ячейка АА17")]
    public string? AA17 { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Номер группы")]
    public string? Group { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Рост (см)")]
    public string? Height { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Вес (кг)")]
    public string? Weight { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("% жировой ткани")]
    public string? FatPercent { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("% мышечной ткани")]
    public string? MusclePercent { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Индекс массы тела (IMT)")]
    public string? IMT { get; set; }

    [BindProperty]
    [DisplayName("Дата")]
    public DateOnly? OrtostatDate { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Пульс лежа за 1 мин")]
    public string? OrtostatLayingPulse { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Пульс стоя за 1 мин")]
    public string? OrtostatStandingPulse { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Результат (разница между пульсом стоя и пульсом лежа)")]
    public string? OrtostatResult { get; set; }

    [BindProperty]
    [DisplayName("Дата")]
    public DateOnly? RuffieDate { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Пульс перед нагрузкой (за 15 сек)")]
    public string? RuffieBefore { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Пульс после нагрузки (за 15 секунд)")]
    public string? RuffieAfter { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Пульс за последние 15 сек в первые минуты отдыха")]
    public string? RuffieRelax { get; set; }

    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Поднимания туловища из положения лежа (кол-во повторов)")]
    public string? Norma1 { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Отжимания (кол-во повторов)")]
    public string? Norma2 { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Гибкость (см)")]
    public string? Norma3 { get; set; }

    [BindProperty]
    [DisplayName("Дата")]
    public DateOnly? PulsePressureDate { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Пульс")]
    public string? Pulse { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "Максимальная длина - {0} символов!")]
    [DisplayName("Артериальное давление")]
    public string? Pressure { get; set; }

    public Task<IActionResult> OnPost()
    {
        IActionResult result;
        if (!ModelState.IsValid)
        {
            result = Page(); 
        }
        else
        {
            var exporter = new Exporter();
            result = this.XlFile(exporter.Export(this), $"Дневник самоконтроля ({AA17}, {Group})");
        }
        return Task.FromResult(result);
    }
}
