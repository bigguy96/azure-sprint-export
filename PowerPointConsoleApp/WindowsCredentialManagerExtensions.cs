using Microsoft.Extensions.Configuration;

namespace PowerPointConsoleApp;

    public static class WindowsCredentialManagerExtensions
    {
        public static IConfigurationBuilder AddWindowsCredentialManager(this IConfigurationBuilder builder, string[] keys)
        {
            var dict = new Dictionary<string, string?>();
            foreach (var key in keys)
            {
                var value = WinCred.GetCredential(key);
                if (!string.IsNullOrEmpty(value))
                {
                    dict[key] = value;
                }
            }
            builder.AddInMemoryCollection(dict!);
            return builder;
        }
    }