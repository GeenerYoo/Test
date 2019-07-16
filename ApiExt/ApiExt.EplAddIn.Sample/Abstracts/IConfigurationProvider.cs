using System.Configuration;

namespace ApiExt.EplAddIn.Sample.Abstracts
{
    public interface IConfigurationProvider
    {
        Configuration Configuration { get; }

        KeyValueConfigurationCollection Appsettings { get; }

        ConnectionStringSettingsCollection ConnectionStrings { get; }

        string ConfigurationFilePath { get; }

        string GetConnectionString(string name);
    }
}
