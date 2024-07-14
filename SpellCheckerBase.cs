using System.Runtime.InteropServices;

namespace SpellChecker;

public class SpellCheckerBase
{
    protected enum CORRECTIVE_ACTION
    {
        CORRECTIVE_ACTION_NONE,
        CORRECTIVE_ACTION_GET_SUGGESTIONS,
        CORRECTIVE_ACTION_REPLACE,
        CORRECTIVE_ACTION_DELETE,
    }

    [Guid("B7C82D61-FBE8-4B47-9B27-6C0D2E0DE0A3")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    protected interface ISpellingError
    {
        uint StartIndex { get; }
        uint Length { get; }
        CORRECTIVE_ACTION CorrectiveAction { get; }
        string Replacement { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("803E3BD4-2828-4410-8290-418D1D73C762")]
    [ComImport]
    protected interface IEnumSpellingError
    {
        [return: MarshalAs(UnmanagedType.Interface)]
        ISpellingError Next();
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("00000101-0000-0000-C000-000000000046")]
    [ComImport]
    protected interface IEnumString
    {
        void Next([In] uint celt, [MarshalAs(UnmanagedType.LPWStr)] out string rgelt, out uint pceltFetched);
        void Skip([In] uint celt);
        void Reset();
        void Clone([MarshalAs(UnmanagedType.Interface)] out IEnumString ppenum);
    }

    [Guid("432E5F85-35CF-4606-A801-6F70277E1D7A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    protected interface IOptionDescription
    {
        string Id { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
        string Heading { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
        string Description { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
        IEnumString Labels { [return: MarshalAs(UnmanagedType.Interface)] get; }
    }

    [Guid("0B83A5B0-792F-4EAB-9799-ACF52C5ED08A")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    protected interface ISpellCheckerChangedEventHandler
    {
        void Invoke([MarshalAs(UnmanagedType.Interface), In] ISpellChecker sender);
    }

    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [Guid("B6FD0B71-E2BC-4653-8D05-F197E412770B")]
    [ComImport]
    protected interface ISpellChecker
    {
        string languageTag { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumSpellingError Check([MarshalAs(UnmanagedType.LPWStr), In] string text);
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumString Suggest([MarshalAs(UnmanagedType.LPWStr), In] string word);
        void Add([MarshalAs(UnmanagedType.LPWStr), In] string word);
        void Ignore([MarshalAs(UnmanagedType.LPWStr), In] string word);
        void AutoCorrect([MarshalAs(UnmanagedType.LPWStr), In] string from, [MarshalAs(UnmanagedType.LPWStr), In] string to);
        byte GetOptionValue([MarshalAs(UnmanagedType.LPWStr), In] string optionId);
        IEnumString OptionIds { [return: MarshalAs(UnmanagedType.Interface)] get; }
        string Id { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
        string LocalizedName { [return: MarshalAs(UnmanagedType.LPWStr)] get; }
        uint add_SpellCheckerChanged([MarshalAs(UnmanagedType.Interface), In] ISpellCheckerChangedEventHandler handler);
        void remove_SpellCheckerChanged([In] uint eventCookie);
        [return: MarshalAs(UnmanagedType.Interface)]
        IOptionDescription GetOptionDescription([MarshalAs(UnmanagedType.LPWStr), In] string optionId);
        [return: MarshalAs(UnmanagedType.Interface)]
        IEnumSpellingError ComprehensiveCheck([MarshalAs(UnmanagedType.LPWStr), In] string text);
    }

    [Guid("8E018A9D-2415-4677-BF08-794EA61F94BB")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    [ComImport]
    protected interface ISpellCheckerFactory
    {
        IEnumString SupportedLanguages { [return: MarshalAs(UnmanagedType.Interface)] get; }
        int IsSupported([MarshalAs(UnmanagedType.LPWStr), In] string languageTag);
        [return: MarshalAs(UnmanagedType.Interface)]
        ISpellChecker CreateSpellChecker([MarshalAs(UnmanagedType.LPWStr), In] string languageTag);
    }

    [Guid("7AB36653-1796-484B-BDFA-E74F1DB7C1DC")]
    [ComImport]
    protected class SpellCheckerFactoryClass
    {
    }
}
