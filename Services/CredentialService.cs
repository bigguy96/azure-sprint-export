using System;
using System.Runtime.InteropServices;
using System.Text;

namespace SprintItemsApp.Services
{
    public class CredentialService
    {
        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct CREDENTIAL
        {
            public int Flags;
            public int Type;
            public string TargetName;
            public string Comment;
            public System.Runtime.InteropServices.ComTypes.FILETIME LastWritten;
            public int CredentialBlobSize;
            public IntPtr CredentialBlob;
            public int Persist;
            public int AttributeCount;
            public IntPtr Attributes;
            public string TargetAlias;
            public string UserName;
        }

        [DllImport("advapi32.dll", SetLastError = true, CharSet = CharSet.Unicode)]
        private static extern bool CredRead(string target, int type, int flags, out IntPtr credentialPtr);

        [DllImport("advapi32.dll", SetLastError = true)]
        private static extern void CredFree(IntPtr credentialPtr);

        public bool TryGetCredential(string targetName, out string username, out string password)
        {
            username = null;
            password = null;

            if (CredRead(targetName, 1 /* CRED_TYPE_GENERIC */, 0, out IntPtr credentialPtr))
            {
                try
                {
                    CREDENTIAL credential = Marshal.PtrToStructure<CREDENTIAL>(credentialPtr);
                    username = credential.UserName;
                    password = Marshal.PtrToStringUni(credential.CredentialBlob, credential.CredentialBlobSize / 2);
                    return true;
                }
                finally
                {
                    CredFree(credentialPtr);
                }
            }
            return false;
        }
    }
}