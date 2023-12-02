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
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("������ ��17")]
    public string? AA17 { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("����� ������")]
    public string? Group { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("���� (��)")]
    public string? Height { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("��� (��)")]
    public string? Weight { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("% ������� �����")]
    public string? FatPercent { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("% �������� �����")]
    public string? MusclePercent { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("������ ����� ���� (IMT)")]
    public string? IMT { get; set; }

    [BindProperty]
    [DisplayName("����")]
    public DateOnly? OrtostatDate { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("����� ���� �� 1 ���")]
    public string? OrtostatLayingPulse { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("����� ���� �� 1 ���")]
    public string? OrtostatStandingPulse { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("��������� (������� ����� ������� ���� � ������� ����)")]
    public string? OrtostatResult { get; set; }

    [BindProperty]
    [DisplayName("����")]
    public DateOnly? RuffieDate { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("����� ����� ��������� (�� 15 ���)")]
    public string? RuffieBefore { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("����� ����� �������� (�� 15 ������)")]
    public string? RuffieAfter { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("����� �� ��������� 15 ��� � ������ ������ ������")]
    public string? RuffieRelax { get; set; }

    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("���������� �������� �� ��������� ���� (���-�� ��������)")]
    public string? Norma1 { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("��������� (���-�� ��������)")]
    public string? Norma2 { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("�������� (��)")]
    public string? Norma3 { get; set; }

    [BindProperty]
    [DisplayName("����")]
    public DateOnly? PulsePressureDate { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("�����")]
    public string? Pulse { get; set; }
    [BindProperty]
    [MaxLength(200, ErrorMessage = "������������ ����� - {0} ��������!")]
    [DisplayName("������������ ��������")]
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
            result = this.XlFile(exporter.Export(this), $"������� ������������ ({AA17}, {Group})");
        }
        return Task.FromResult(result);
    }
}
