using Microsoft.Extensions.Configuration;

namespace PowerPointConsoleApp.Infrastructure;

// IConfigurationBuilder extension that populates configuration keys from Windows Credential Manager.
// Allows secrets such as the Azure DevOps PAT to be stored securely on developer machines
// without committing them to source control.
public static class WindowsCredentialManagerExtensions
{
    public static IConfigurationBuilder AddWindowsCredentialManager(
        this IConfigurationBuilder builder, string[] keys)
    {
        var dict = new Dictionary<string, string?>();
        foreach (var key in keys)
        {
            var value = WinCred.GetCredential(key);
            if (!string.IsNullOrEmpty(value))
                dict[key] = value;
        }

        builder.AddInMemoryCollection(dict!);
        return builder;
    }
}
