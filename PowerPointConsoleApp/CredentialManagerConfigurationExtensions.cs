using Microsoft.Extensions.Configuration;
using System.Runtime.InteropServices;

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

    public static class WinCred
    {
        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool CredRead(string target, int type, int reservedFlag, out IntPtr credentialPtr);


        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern void CredFree(IntPtr buffer);


        public static string? GetCredential(string target)
        {
            if (CredRead(target, 1 /* CRED_TYPE_GENERIC */, 0, out IntPtr credPtr))
            {
                var cred = (CREDENTIAL)Marshal.PtrToStructure(credPtr, typeof(CREDENTIAL))!;
                string password = Marshal.PtrToStringUni(cred.CredentialBlob, (int)cred.CredentialBlobSize / 2)!;
                CredFree(credPtr);
                return password;
            }
            return null;
        }


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct CREDENTIAL
        {
            public int Flags;
            public int Type;
            public string TargetName;
            public string Comment;
            public long LastWritten;
            public int CredentialBlobSize;
            public IntPtr CredentialBlob;
            public int Persist;
            public int AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }
    }