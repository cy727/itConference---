using System;
using System.Collections.Generic;
using System.Text;
using System.Management;
using System.Security.Cryptography;
using System.IO;

namespace itConference
{
    class chrislock
    {
        //private string strSerial;

        public string getDiskSerial()
        {
            String HardDiskID = "";
            ManagementClass cimobject = new ManagementClass("Win32_DiskDrive");
            ManagementObjectCollection moc = cimobject.GetInstances();
            try
            {
                foreach (ManagementObject mo in moc)
                {
                    HardDiskID = (string)mo.Properties["Model"].Value;
                    mo.Dispose();
                }
            }
            catch
            {
                HardDiskID = "";
            }
            return HardDiskID;
        }

        public string getMacAddress()
        {
            string MacAddress = "";
            ManagementClass mc = new ManagementClass("Win32_NetworkAdapterConfiguration");
            ManagementObjectCollection moc2 = mc.GetInstances();
            foreach (ManagementObject mo in moc2)
            {
                if ((bool)mo["IPEnabled"] == true)
                    MacAddress = mo["MacAddress"].ToString();
                mo.Dispose();
            }
            return MacAddress; 
        }

        //加密数据
        public string getLockEnData(string strData, string strPass)
        {
            if (strPass == "" || strData == "") return "";
            byte[] enIV ={ 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF};
            byte[] enKey = Encoding.UTF8.GetBytes(strPass);

            DESCryptoServiceProvider enProvider = new DESCryptoServiceProvider();
            byte[] enData = Encoding.UTF8.GetBytes(strData);
            MemoryStream enMemory = new MemoryStream();
            CryptoStream enCrypt = new CryptoStream(enMemory, enProvider.CreateEncryptor(enKey, enIV), CryptoStreamMode.Write);

            enCrypt.Write(enData,0,enData.Length);
            enCrypt.FlushFinalBlock();

            return Convert.ToBase64String(enMemory.ToArray());
        }

        //解密数据
        public string getUnlockEnData(string strData, string strPass)
        {
            if (strPass == "" || strData == "") return "";

            byte[] enIV ={ 0x12, 0x34, 0x56, 0x78, 0x90, 0xAB, 0xCD, 0xEF };


            try
            {
                byte[] enKey = System.Text.Encoding.Default.GetBytes(strPass);
                byte[] enBytes = Convert.FromBase64String(strData);
                MemoryStream deMemory = new MemoryStream();
                DESCryptoServiceProvider deProvider = new DESCryptoServiceProvider();
                CryptoStream deCrypt = new CryptoStream(deMemory, deProvider.CreateDecryptor(enKey, enIV), CryptoStreamMode.Write);

                deCrypt.Write(enBytes, 0, enBytes.Length);
                deCrypt.FlushFinalBlock();
                System.Text.Encoding deEncoding = new System.Text.UTF8Encoding();
                return deEncoding.GetString(deMemory.ToArray());
            }
            catch
            {
                return "";
            }
        }




    }
}
