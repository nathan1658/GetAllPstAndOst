using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace Main
{
    // Registry class for loading NTUser.dat file and 
    // temporary hive key and then unloading temporary hive key.

    public class RegistryInterop
    {
        [StructLayout(LayoutKind.Sequential)]
        public struct LUID
        {
            public int LowPart;
            public int HighPart;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct TOKEN_PRIVILEGES
        {
            public LUID Luid;
            public int Attributes;
            public int PrivilegeCount;
        }

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern int OpenProcessToken(int ProcessHandle, int DesiredAccess, ref int tokenhandle);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        public static extern int GetCurrentProcess();

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern int LookupPrivilegeValue(string lpsystemname, string lpname, [MarshalAs(UnmanagedType.Struct)] ref LUID lpLuid);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto)]
        public static extern int AdjustTokenPrivileges(int tokenhandle, int disableprivs, [MarshalAs(UnmanagedType.Struct)]ref TOKEN_PRIVILEGES Newstate, int bufferlength, int PreivousState, int Returnlength);

        [DllImport("advapi32.dll")]
        public static extern int RegLoadKey(uint hKey, string lpSubKey, string lpFile);

        [DllImport("advapi32.dll")]
        public static extern int RegUnLoadKey(uint hKey, string lpSubKey);

        public const int TOKEN_ADJUST_PRIVILEGES = 0x00000020;
        public const int TOKEN_QUERY = 0x00000008;
        public const int SE_PRIVILEGE_ENABLED = 0x00000002;
        public const string SE_RESTORE_NAME = "SeRestorePrivilege";
        public const string SE_BACKUP_NAME = "SeBackupPrivilege";
        public const uint HKEY_USERS = 0x80000003;

        //temporary hive key
        public const string HIVE_SUBKEY = "Test";


        static private Boolean gotPrivileges = false;

        static private void GetPrivileges()
        {


            int token = 0;

            int retval = 0;

            TOKEN_PRIVILEGES TP = new TOKEN_PRIVILEGES();

            TOKEN_PRIVILEGES TP2 = new TOKEN_PRIVILEGES();

            LUID RestoreLuid = new LUID();

            LUID BackupLuid = new LUID();

            retval = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES | TOKEN_QUERY, ref token);

            retval = LookupPrivilegeValue(null, SE_RESTORE_NAME, ref RestoreLuid);

            retval = LookupPrivilegeValue(null, SE_BACKUP_NAME, ref BackupLuid);

            TP.PrivilegeCount = 1;

            TP.Attributes = SE_PRIVILEGE_ENABLED;

            TP.Luid = RestoreLuid;

            TP2.PrivilegeCount = 1;

            TP2.Attributes = SE_PRIVILEGE_ENABLED;

            TP2.Luid = BackupLuid;

            retval = AdjustTokenPrivileges(token, 0, ref TP, 1024, 0, 0);

            retval = AdjustTokenPrivileges(token, 0, ref TP2, 1024, 0, 0);

            gotPrivileges = true;
        }

        static public string Load(string file)
        {
            if (!gotPrivileges)
                GetPrivileges();
            var result = RegLoadKey(HKEY_USERS, HIVE_SUBKEY, file);
            Console.WriteLine("Load return " + result);
            if (result != 0)//Fail
                return null;
            return HIVE_SUBKEY;
            //RegLoadKey(HKEY_USERS, HIVE_SUBKEY, file);
            //return HIVE_SUBKEY;
        }

        static public void Unload()
        {
            if (!gotPrivileges)
                GetPrivileges();
            int output = RegUnLoadKey(HKEY_USERS, HIVE_SUBKEY);
            Console.WriteLine("Unload return " + output);
        }
    }
}
