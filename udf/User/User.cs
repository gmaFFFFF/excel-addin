using System.DirectoryServices.AccountManagement;
using System.Security.Principal;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace gmafffff.excel.udf.User;

public class User {
    public User() {
        using var user = WindowsIdentity.GetCurrent();

        try {
            using PrincipalContext adContextGeneral = new(ContextType.Domain);
            using var              userPrincipal    = UserPrincipal.FindByIdentity(adContextGeneral, user.Name);

            ОтображаемоеИмя = userPrincipal?.DisplayName;
            Фамилия         = userPrincipal?.Surname;
            Имя             = userPrincipal?.GivenName;
            Имя2            = userPrincipal?.Name;
            Отчество        = userPrincipal?.MiddleName;
            AdИмя           = userPrincipal?.UserPrincipalName;
            УчетнаяЗапись   = user.Name;
            Телефон         = userPrincipal?.VoiceTelephoneNumber;
            Email           = userPrincipal?.EmailAddress;
        }
        catch (PlatformNotSupportedException) {
            Ошибка = true;
            ОшибкаТекст = "Неподдерживаемая платформа. "
                        + "Попробуйте рядом с *.xll-файлом add-in разместить следующие библиотеки:\n"
                        + "— System.DirectoryServices.dll\n"
                        + "— System.DirectoryServices.AccountManagement.dll\n"
                        + "— System.DirectoryServices.Protocols.dll";
        }
        catch (PrincipalServerDownException) {
            Ошибка      = true;
            ОшибкаТекст = "Недоступен сервер ActiveDirectory";
        }
        catch (Exception ex) {
            Ошибка      = true;
            ОшибкаТекст = ex.Message;
        }
    }

    public string? ОтображаемоеИмя { get; }
    public string? Фамилия { get; }
    public string? Имя { get; }
    public string? Имя2 { get; }
    public string? Отчество { get; }
    public string? AdИмя { get; }
    public string УчетнаяЗапись { get; }
    public string? Телефон { get; }
    public string? Email { get; }
    public bool Ошибка { get; }
    public string? ОшибкаТекст { get; }

    public string Json() {
        JsonSerializerOptions настройкиСохранения = new() {
            WriteIndented = true, AllowTrailingCommas = true, Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };

        return JsonSerializer.Serialize(this, настройкиСохранения);
    }
}