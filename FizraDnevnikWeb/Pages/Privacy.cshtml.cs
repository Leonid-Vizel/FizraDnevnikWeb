using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;

namespace FizraDnevnikWeb.Pages;

[IgnoreAntiforgeryToken]
[ResponseCache(VaryByHeader = "User-Cached", Duration = 60)]
public sealed class PrivacyModel : PageModel { }
