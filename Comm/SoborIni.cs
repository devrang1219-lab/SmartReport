using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;


namespace AIRoboticsWCS.Lib
{
    public class SoborIni
    {
        [DllImport("kernel32.dll")]
        private static extern int GetPrivateProfileString(
            string section, string key, string def, StringBuilder retVal, int size, string filePath);

        [DllImport("kernel32.dll")]
        private static extern long WritePrivateProfileString(
            string section, string key, string value, string filePath);

        public string DirectoryName = @".";
        public string FileName = "./main.ini";
        public string Message = "";

        public bool IsExistFile = false;

        public string DefaultSection = "Config";

        public string Temp = "";


        /// <summary>
        /// filename의 ini 파일 생성 (ex) "./config.ini"
        /// </summary>
        /// <param name="fileName"></param>
        public SoborIni(string fileName = "")
        {
            if (fileName != "")
            {
                FileName = fileName;
            }

            FileInfo fi = new FileInfo(FileName);
            IsExistFile = fi.Exists;
        }

        public string Write(string section, string key, string value)
        {
            try
            {
                if (section == "") section = DefaultSection;
                WritePrivateProfileString(section, key, value, FileName);
                return value;
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                return "";
            }
        }
        public string Write(string key, string value)
        {
            Write("Config", key, value);
            return value;
        }

        public string WriteCrypt(string key, string value)
        {
            string cryptKey = "MoornmoIoT";
            Write("Config", key, Encrypt(value, cryptKey));
            return value;
        }
        public string WriteCrypt(string section, string key, string value)
        {
            string cryptKey = "MoornmoIoT";
            Write(section, key, Encrypt(value, cryptKey));
            return value;
        }
        //
        public int WriteInt(string section, string key, int value)
        {
            Write(section, key, value.ToString());
            return value;
        }
        //
        public int WriteInt(string key, int value)
        {
            Write("", key, value.ToString());
            return value;
        }
        //
        public bool WriteBool(string section, string key, bool value)
        {
            Write(section, key, value.ToString());
            return value;
        }
        //
        public bool WriteBool(string key, bool value)
        {
            Write("", key, value.ToString());
            return value;
        }
        //
        public string[] WriteArray(string section, string key, string[] value, string delimiter = "|")
        {
            Write(section, key, string.Join(delimiter, value));
            return value;
        }
        //
        public string[] WriteArray(string key, string[] value)
        {
            Write("", key, string.Join("|", value));
            return value;
        }
        //
        public string Write(TextBox obj)
        {
            string result = obj.Text;
            Write(obj.Name, obj.Text);
            return result;
        }
        //
        public string[] WriteList(TextBox obj)
        {
            WriteArray(obj.Name + "_list", obj.Lines);
            return obj.Lines;
        }
        //
        public int Write(NumericUpDown obj)
        {
            int result = (int)obj.Value;
            Write(obj.Name, obj.Value.ToString());
            return result;
        }
        //
        public bool Write(CheckBox obj)
        {
            bool result = obj.Checked;
            WriteBool(obj.Name, obj.Checked);
            return result;
        }
        //
        public bool Write(RadioButton obj)
        {
            bool result = obj.Checked;
            WriteBool(obj.Name, obj.Checked);
            return result;
        }
        //
        public void WriteList(ComboBox obj)
        {
            List<string> slist = new List<string>();
            slist.AddRange(obj.Items.Cast<string>());
            Write(obj.Name + "_list", String.Join("|", slist));
            Write(obj.Name, obj.Text);
        }
        //
        public string Write(ComboBox obj)
        {
            string result = obj.Text;
            Write(obj.Name, obj.Text);
            return result;
        }
        //
        public void WriteList(ListBox obj)
        {
            List<string> slist = new List<string>();
            slist.AddRange(obj.Items.Cast<string>());
            Write(obj.Name + "_list", String.Join("|", slist));
            Write(obj.Name, obj.Text);
        }
        //
        public string Write(ListBox obj)
        {
            string result = obj.Text;
            Write(obj.Name, obj.Text);
            return result;
        }
        //
        public string ReadDefault(string section, string key, string defaultValue)
        {
            try
            {
                if (section == "") section = "Config";

                StringBuilder temp = new StringBuilder(2000);
                int i = GetPrivateProfileString(section, key, "", temp, 2000, FileName);
                if (i > 0)
                {
                    return temp.ToString();
                }
                return defaultValue;
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                return defaultValue;
            }
        }

        public string ReadDefault(string key, string defaultValue)
        {
            string section = "Config";
            return ReadDefault(section, key, defaultValue);
        }
        public string Read(string section, string key)
        {
            try
            {
                if (section == "") section = "Config";

                StringBuilder temp = new StringBuilder(2000);
                int i = GetPrivateProfileString(section, key, "", temp, 2000, FileName);
                return temp.ToString();
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                return "";
            }
        }
        //
        public string Read(string key)
        {
            return Read("", key);
        }
        //
        public string ReadCrypt(string key)
        {
            string cryptKey = "MoornmoIoT";
            return Decrypt(Read("", key), cryptKey);
        }
        public string ReadCrypt(string section, string key)
        {
            string cryptKey = "MoornmoIoT";
            return Decrypt(Read(section, key), cryptKey);
        }
        //
        public int ReadInt(string section, string key, int defaultValue)
        {
            int result = defaultValue;
            try
            {
                result = Convert.ToInt32(Read(section, key));
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                result = defaultValue;
            }

            return result;
        }
        //
        public int ReadInt(string section, string key)
        {
            return ReadInt(section, key, 0);
        }
        //
        public int ReadIntD(string key, int defaultValue)
        {
            int result = defaultValue;
            try
            {
                result = ReadInt(DefaultSection, key);
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                result = defaultValue;
            }
            return result;
        }
        //
        public int ReadInt(string key)
        {
            return ReadInt("", key);
        }
        //
        public bool ReadBoolD(string key, bool defaultValue = false)
        {
            bool result = defaultValue;
            try
            {
                result = ReadBool(DefaultSection, key);
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                result = defaultValue;
            }
            return result;
        }
        public bool ReadBool(string section, string key, bool defaultValue = false)
        {
            bool result = defaultValue;
            try
            {
                result = ReadBool(section, key);
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                result = defaultValue;
            }

            return result;
        }
        //
        public bool ReadBool(string section, string key)
        {
            bool result = false;

            try
            {
                result = Convert.ToBoolean(Read(section, key));
            }
            catch (Exception ex)
            {
                Message = ex.Message;
            }

            return result;
        }
        //
        public bool ReadBool(string key, bool defaultValue = false)
        {
            return ReadBool("", key, defaultValue);
        }
        //
        public bool ReadBool(string key)
        {
            return ReadBool("", key);
        }
        //
        public string[] ReadArray(string section, string key, string delimiter = "|")
        {
            try
            {
                string resultString = Read(section, key);
                string[] resultArray;

                if (resultString == "")
                {
                    resultArray = new string[0] { };
                }
                else
                {
                    resultArray = resultString.Split(delimiter[0]);
                }

                return resultArray;
            }
            catch (Exception ex)
            {
                Message = ex.Message;
                return null;
            }
        }
        //
        public string[] ReadArray(string key)
        {
            return ReadArray("", key, "|");
        }

        public int[] ReadArrayInt(string key)
        {
            string[] data = ReadArray(key);
            List<int> ret = new List<int>();

            try
            {
                foreach (string d in data)
                {
                    ret.Add(Convert.ToInt32(d));
                }

                return ret.ToArray();
            }
            catch
            {
                return ret.ToArray();
            }
        }
        //
        public string[] ReadList(TextBox obj)
        {
            obj.Lines = ReadArray("", obj.Name + "_list", "|");
            return obj.Lines;
        }
        //
        public string Read(TextBox obj)
        {
            string result = Read(obj.Name);
            obj.Text = result;
            return result;
        }
        //
        public int Read(NumericUpDown obj)
        {
            int result = ReadInt(obj.Name);
            obj.Value = result;
            return result;
        }
        //
        public bool Read(CheckBox obj)
        {
            bool result = ReadBool(obj.Name);
            obj.Checked = result;
            return result;
        }
        //
        public bool Read(RadioButton obj)

        {
            bool result = ReadBool(obj.Name);
            obj.Checked = result;
            return result;
        }
        //
        public string Read(ComboBox obj)
        {
            string result = Read(obj.Name);
            obj.Text = result;
            return result;
        }
        //
        public void ReadList(ComboBox obj)
        {
            try
            {
                obj.Items.Clear();
                obj.Items.AddRange(Read(obj.Name + "_list").Split('|'));
            }
            catch (Exception ex)
            {
                Message = ex.Message;
            }
        }
        //
        public string[] ReadList(string list)
        {
            return Read(list).Split('|');
        }
        //
        public string ReadText(ComboBox obj)
        {
            string result = Read(obj.Name);
            obj.Text = result;
            return result;
        }
        //
        public void Read(ListBox obj)
        {
            try
            {
                obj.Items.Clear();
                obj.Items.AddRange(Read(obj.Name + "_list").Split('|'));
            }
            catch (Exception ex)
            {
                Message = ex.Message;
            }
        }

        public string Encrypt(string textToEncrypt, string key)
        {
            RijndaelManaged rijndaelCipher = new RijndaelManaged();
            rijndaelCipher.Mode = CipherMode.CBC;
            rijndaelCipher.Padding = PaddingMode.PKCS7;

            rijndaelCipher.KeySize = 128;
            rijndaelCipher.BlockSize = 128;

            byte[] pwdBytes = Encoding.UTF8.GetBytes(key);
            byte[] keyBytes = new byte[16];
            int len = pwdBytes.Length;

            if (len > keyBytes.Length)
            {
                len = keyBytes.Length;
            }

            Array.Copy(pwdBytes, keyBytes, len);
            rijndaelCipher.Key = keyBytes;
            rijndaelCipher.IV = keyBytes;
            ICryptoTransform transform = rijndaelCipher.CreateEncryptor();
            byte[] plainText = Encoding.UTF8.GetBytes(textToEncrypt);

            return Convert.ToBase64String(transform.TransformFinalBlock(plainText, 0, plainText.Length));
        }

        public string Decrypt(string textToDecrypt, string key = "MoornmoIoT")
        {
            try
            {
                RijndaelManaged rijndaelCipher = new RijndaelManaged();
                rijndaelCipher.Mode = CipherMode.CBC;
                rijndaelCipher.Padding = PaddingMode.PKCS7;

                rijndaelCipher.KeySize = 128;
                rijndaelCipher.BlockSize = 128;

                byte[] encryptedData = Convert.FromBase64String(textToDecrypt);
                byte[] pwdBytes = Encoding.UTF8.GetBytes(key);
                byte[] keyBytes = new byte[16];
                int len = pwdBytes.Length;

                if (len > keyBytes.Length)
                {
                    len = keyBytes.Length;
                }

                Array.Copy(pwdBytes, keyBytes, len);
                rijndaelCipher.Key = keyBytes;
                rijndaelCipher.IV = keyBytes;
                byte[] plainText = rijndaelCipher.CreateDecryptor().TransformFinalBlock(encryptedData, 0, encryptedData.Length);

                return Encoding.UTF8.GetString(plainText);
            }
            catch
            {
                return "";
            }
        }

    }
}
