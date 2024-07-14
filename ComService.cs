using System.Runtime.InteropServices;
using System.Text.Json;
using System.Text.RegularExpressions;

namespace SpellChecker;

[ComVisible(true)]
[Guid("fe103d6e-e71b-414c-80bf-982f18f23237")]
public class ComService
{
    public string SpellCheck(string text)
    {
        var spellCheckResults = SpellCheckerAPI.SpellCheck(text);
        string x = JsonSerializer.Serialize(spellCheckResults);
        string jsonString = Regex.Unescape(x);

        return jsonString;
    }
}