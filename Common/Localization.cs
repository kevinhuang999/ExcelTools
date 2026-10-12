namespace ExcelTools.Common;

using System.Globalization;
using ExcelDna.Integration;

/// <summary>
/// Follows the display language of the Excel client the add-in is loaded into.
///
/// Chinese Excel  -> Chinese UI
/// English Excel  -> English UI
/// anything else  -> English
///
/// The language is resolved lazily on first use, because the ribbon XML is
/// requested before the add-in can be sure the Excel object model is ready.
/// If reading Excel's UI language fails for any reason we fall back to the
/// process UI culture, and finally to English.
/// </summary>
internal static class Loc
{
    /// <summary>msoLanguageIDUI - the language of the Excel user interface.</summary>
    private const int MsoLanguageIdUi = 2;

    /// <summary>Primary language id for Chinese (any script/region).</summary>
    private const int PrimaryLanguageChinese = 0x04;

    private static bool _initialized;
    private static bool _chinese;

    public static bool IsChinese
    {
        get
        {
            EnsureInitialized();
            return _chinese;
        }
    }

    /// <summary>Pick the Chinese or the English variant of a string.</summary>
    public static string T(string chinese, string english)
        => IsChinese ? chinese : english;

    private static void EnsureInitialized()
    {
        if (_initialized) return;

        int languageId = ReadExcelUiLanguageId();
        if (languageId == 0)
            languageId = ReadProcessUiLanguageId();

        if (languageId == 0)
        {
            // Could not determine the language yet (the Excel object model may
            // not be ready this early). Answer English for now, but do not
            // cache, so a later call can resolve it properly.
            _chinese = false;
            return;
        }

        _initialized = true;

        // The low 10 bits are the primary language id.
        _chinese = (languageId & 0x3FF) == PrimaryLanguageChinese;
    }

    private static int ReadExcelUiLanguageId()
    {
        try
        {
            dynamic app = ExcelDnaUtil.Application;
            if (app == null) return 0;

            // Application.LanguageSettings.LanguageID(msoLanguageIDUI)
            dynamic settings = app.LanguageSettings;
            return Convert.ToInt32(settings.LanguageID(MsoLanguageIdUi));
        }
        catch
        {
            return 0;
        }
    }

    private static int ReadProcessUiLanguageId()
    {
        try
        {
            return CultureInfo.CurrentUICulture.LCID;
        }
        catch
        {
            return 0;
        }
    }
}
