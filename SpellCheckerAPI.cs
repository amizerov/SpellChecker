using System.Runtime.InteropServices;

namespace SpellChecker;

public record SpellCheckResult(string Word, string Action, string Replacement, List<string> Suggestions);

public class SpellCheckerAPI : SpellCheckerBase
{
    public static List<SpellCheckResult> SpellCheck(string s)
    {
        SpellCheckerFactoryClass? factory = null;
        ISpellCheckerFactory? ifactory = null;
        ISpellChecker? checker = null;
        ISpellingError? error = null;
        IEnumSpellingError? errors = null;
        IEnumString? suggestions = null;

        List<SpellCheckResult> spellCheckResults = new();

        try
        {
            factory = new SpellCheckerFactoryClass();
            ifactory = (ISpellCheckerFactory)factory;

            //проверим поддержку русского языка
            int res = ifactory.IsSupported("ru-RU");
            if (res == 0) { throw new Exception("Fatal error: russian language not supported!"); }

            checker = ifactory.CreateSpellChecker("ru-RU");

            errors = checker.Check(s);
            while (true)
            {
                //получаем ошибку
                if (error != null) { Marshal.ReleaseComObject(error); error = null; }
                error = errors.Next();
                if (error == null) break;

                //получаем слово с ошибкой
                string word = s.Substring((int)error.StartIndex, (int)error.Length);
                string action = "";
                string replac = error.Replacement;
                List<string> sugges = new();

                //получаем рекомендуемое действие
                switch (error.CorrectiveAction)
                {
                    case CORRECTIVE_ACTION.CORRECTIVE_ACTION_DELETE:
                        action = "удалить";
                        break;

                    case CORRECTIVE_ACTION.CORRECTIVE_ACTION_REPLACE:
                        action = "заменить";
                        break;

                    case CORRECTIVE_ACTION.CORRECTIVE_ACTION_GET_SUGGESTIONS:
                        action = "заменить на одно из";

                        if (suggestions != null) { Marshal.ReleaseComObject(suggestions); suggestions = null; }

                        //получаем список слов, предложенных для замены
                        suggestions = checker.Suggest(word);

                        while (true)
                        {
                            string suggestion;
                            uint count = 0;
                            suggestions.Next(1, out suggestion, out count);
                            if (count == 1) sugges.Add(suggestion);
                            else break;
                        }
                        break;
                }
                  
                if(replac == "") replac = sugges.Count > 0 ? sugges[0] : "";
                spellCheckResults.Add(new SpellCheckResult(word, action, replac, sugges));
            }
        }
        finally
        {
            if (suggestions != null) { Marshal.ReleaseComObject(suggestions); }
            if (factory != null) { Marshal.ReleaseComObject(factory); }
            if (ifactory != null) { Marshal.ReleaseComObject(ifactory); }
            if (checker != null) { Marshal.ReleaseComObject(checker); }
            if (error != null) { Marshal.ReleaseComObject(error); }
            if (errors != null) { Marshal.ReleaseComObject(errors); }
        }

        return spellCheckResults;
    }
}