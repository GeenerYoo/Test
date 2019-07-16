using Eplan.EplApi.Base;

namespace ApiExt.EplAddIn.Sample.Extensions
{
    public static class StringExtensions
    {
        public static MultiLangString GetMultiLangString(this string text)
        {
            MultiLangString mlString = new MultiLangString();

            if (string.IsNullOrWhiteSpace(text))
                mlString.SetAsString(string.Empty);
            else
                mlString.SetAsString(text);

            return mlString;
        }
    }
}
