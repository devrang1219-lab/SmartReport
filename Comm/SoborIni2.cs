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
    public class SoborIni2
    {
        [DllImport("kernel32.dll")]
        private static extern int GetPrivateProfileString(
           string section, string key, string def, StringBuilder retVal, int size, string filePath);

        [DllImport("kernel32.dll")]
        private static extern long WritePrivateProfileString(
            string section, string key, string value, string filePath);

        public string DirectoryName = @".";
        public string FileName = "config.ini";
        public string Section = "Config";
        public string InfoSection = "Info";
        public string CryptKey = "MoornmoIoT";
        public string Message = "";
        public bool ChangeObjName = false;
        public string Temp = "";
        public string[] Temps = new string[10];
        public List<string> ErrorLog = new List<string>();

        public string FilePath
        {
            get
            {
                return DirectoryName + @"\" + FileName;
            }
        }


        public int ErrorCount
        {
            get
            {
                return ErrorLog.Count;
            }
        }

        public bool IsExist
        {
            get
            {
                FileInfo fi = new FileInfo(FilePath);
                return fi.Exists;
            }
        }


        public SoborIni2(string dirName, string fileName)
        {
            if (dirName != "")
            {
                DirectoryName = dirName;
                DirectoryName = dirName.Replace("/", @"\");
                if (DirectoryName.StartsWith(".") == false)
                {
                    DirectoryName = @".\" + DirectoryName;
                }
            }

            if (fileName != "")
            {
                FileName = fileName;
                if (FileName.Contains(".ini") == false)
                {
                    FileName = FileName + ".ini";
                }
            }

            string directoryName = Path.GetDirectoryName(FilePath);
            DirectoryInfo di = new DirectoryInfo(directoryName);
            if (di.Exists == false)
            {
                di.Create();
            }

            SWrite(InfoSection, "StartTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            SWrite(InfoSection, "FileName", FilePath);
        }

        /// <summary>
        /// filePath = "main", "main.ini" 
        /// </summary>
        /// <param name="fileName"></param>
        public SoborIni2(string fileName = "")
        {
            if (fileName != "")
            {
                FileName = fileName;
                if (FileName.Contains(".ini") == false)
                {
                    FileName = FileName + ".ini";
                }
            }
            else
            {
                FileName = "config.ini";
            }

            string directoryName = Path.GetDirectoryName(FilePath);
            DirectoryInfo di = new DirectoryInfo(directoryName);
            if (di.Exists == false)
            {
                di.Create();
            }

            SWrite(InfoSection, "AccessTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            SWrite(InfoSection, "FileName", di.FullName);
        }

        private void AddLog(string text)
        {
            ErrorLog.Add(text);
            if (ErrorCount > 1000)
            {
                ErrorLog.RemoveAt(0);
            }

            SWrite(InfoSection, "LastErrorMessage", text);
            SWrite(InfoSection, "LastErrorTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
        }

        private void AddLog(Exception ex)
        {
            AddLog(ex.Message);
        }

        #region Convert

        public int[] StringToInt(string[] items, int errorValue = 0)
        {
            if (items == null || items.Length == 0)
            {
                return null;
            }

            int[] result = new int[items.Length];
            for (int i = 0; i < items.Length; i++)
            {
                try
                {
                    result[i] = Convert.ToInt32(items[i]);
                }
                catch
                {
                    result[i] = errorValue;
                }
            }
            return result;
        }

        public string[] IntToString(int[] items)
        {
            string[] result = new string[items.Length];
            for (int i = 0; i < items.Length; i++)
            {
                result[i] = items[i].ToString();
            }
            return result;
        }

        public double[] StringToDouble(string[] items, double errorValue = 0)
        {
            if (items == null || items.Length == 0)
            {
                return null;
            }

            double[] result = new double[items.Length];
            for (int i = 0; i < items.Length; i++)
            {
                try
                {
                    result[i] = Convert.ToDouble(items[i]);
                }
                catch
                {
                    result[i] = errorValue;
                }
            }
            return result;
        }

        public string[] DoubleToString(double[] items)
        {
            string[] result = new string[items.Length];
            for (int i = 0; i < items.Length; i++)
            {
                result[i] = items[i].ToString();
            }
            return result;
        }

        public string Encrypt(string textToEncrypt, string cryptKey)
        {
            RijndaelManaged rijndaelCipher = new RijndaelManaged();
            rijndaelCipher.Mode = CipherMode.CBC;
            rijndaelCipher.Padding = PaddingMode.PKCS7;

            rijndaelCipher.KeySize = 128;
            rijndaelCipher.BlockSize = 128;

            byte[] pwdBytes = Encoding.UTF8.GetBytes(cryptKey);
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

        public string Decrypt(string textToDecrypt, string cryptKey, string dvalue = "")
        {
            try
            {
                RijndaelManaged rijndaelCipher = new RijndaelManaged();
                rijndaelCipher.Mode = CipherMode.CBC;
                rijndaelCipher.Padding = PaddingMode.PKCS7;

                rijndaelCipher.KeySize = 128;
                rijndaelCipher.BlockSize = 128;

                byte[] encryptedData = Convert.FromBase64String(textToDecrypt);
                byte[] pwdBytes = Encoding.UTF8.GetBytes(cryptKey);
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
                return dvalue;
            }
        }

        #endregion

        #region Read Base
        private string _Read(string section, string key)
        {
            string result;
            if (section == "") section = Section;
            try
            {
                StringBuilder temp = new StringBuilder(2000);
                int i = GetPrivateProfileString(section, key, "", temp, 2000, FilePath);
                result = temp.ToString();
                return result;
            }
            catch (Exception ex)
            {
                AddLog(ex);
                return "";
            }
        }

        private string _Read(string section, string key, string dvalue)
        {
            string result;
            result = _Read(section, key);

            if (result == null || result == "")
            {
                result = dvalue;
            }
            return result;
        }

        private int _Read(string section, string key, int dvalue)
        {
            int result;
            string temp = _Read(section, key);
            try
            {
                result = int.Parse(temp);
            }
            catch
            {
                result = dvalue;
            }
            return result;
        }

        private double _Read(string section, string key, double dvalue)
        {
            double result;
            string temp = _Read(section, key);
            try
            {
                result = double.Parse(temp);
            }
            catch
            {
                result = dvalue;
            }
            return result;
        }

        private bool _Read(string section, string key, bool dvalue)
        {
            bool result;
            string temp = _Read(section, key);
            try
            {
                result = bool.Parse(temp);
            }
            catch
            {
                result = dvalue;
            }
            return result;
        }

        private string[] _ReadArray(string section, string key, string splitter = "|")
        {
            string[] result;
            string temp = _Read(section, key);

            if (temp != null && temp != "")
            {
                result = temp.Split(splitter[0]);
            }
            else
            {
                result = null;
            }

            return result;
        }

        private List<string> _ReadList(string section, string key, string splitter = "|")
        {
            string[] temp = _ReadArray(section, key, splitter);

            if (temp != null)
            {
                return temp.ToList();
            }
            else
            {
                return null;
            }
        }

        private string _ReadCrypt(string section, string key, string cryptKey = "", string dvalue = "")
        {
            if (cryptKey == "")
            {
                cryptKey = CryptKey;
            }
            string temp = _Read(section, key);
            string result = Decrypt(temp, cryptKey, dvalue);
            return result;
        }
        #endregion

        #region Read Default Section
        public string Read(string key)
        {
            return _Read(Section, key);
        }

        public string Read(string key, string dvalue)
        {
            return _Read(Section, key, dvalue);
        }

        public string Read(TextBox obj, string key = "", string dvalue = "")
        {
            string result = "";
            if (key == "")
            {
                result = _Read(Section, obj.Name, dvalue);
            }
            else
            {
                result = _Read(Section, key, dvalue);
            }
            obj.Text = result;
            return result;
        }

        public string Read(ComboBox obj, string key = "", string dvalue = "")
        {
            string result = "";
            if (key == "")
            {
                result = _Read(Section, obj.Name, dvalue);
            }
            else
            {
                result = _Read(Section, key, dvalue);
            }
            obj.Text = result;
            return result;
        }

        public string Read(DateTimePicker obj, string key = "", string dvalue = "")
        {
            string result = "";
            if (key == "")
            {
                result = _Read(Section, obj.Name, dvalue);
            }
            else
            {
                result = _Read(Section, key, dvalue);
            }

            try
            {
                obj.Text = result;
            }
            catch (Exception)
            {
                return obj.Text;
            }
            return result;
        }

        public int Read(NumericUpDown obj, string key = "", int dvalue = 0)
        {
            int result = 0;
            if (key == "")
            {
                result = _Read(Section, obj.Name, dvalue);
            }
            else
            {
                result = _Read(Section, key, dvalue);
            }
            try
            {
                obj.Value = result;
            }
            catch
            {
                return (int)obj.Value;
            }
            return result;
        }

        public bool Read(CheckBox obj, string key = "", bool dvalue = false)
        {
            bool result = false;
            if (key == "")
            {
                result = _Read(Section, obj.Name, dvalue);
            }
            else
            {
                result = _Read(Section, key, dvalue);
            }
            obj.Checked = result;
            return result;
        }


        public bool Read(RadioButton obj, string key = "", bool dvalue = false)
        {
            bool result = false;
            if (key == "")
            {
                result = _Read(Section, obj.Name, dvalue);
            }
            else
            {
                result = _Read(Section, key, dvalue);
            }
            obj.Checked = result;
            return result;
        }

        public bool ReadBool(string key, bool dvalue = false)
        {
            return _Read(Section, key, dvalue);
        }

        public int ReadInt(string key, int dvalue = 0)
        {
            return _Read(Section, key, dvalue);
        }

        public double ReadDouble(string key, double dvalue = 0)
        {
            return _Read(Section, key, dvalue);
        }

        public string[] ReadArray(string key, string splitter = "|")
        {
            return _ReadArray(Section, key, splitter);
        }

        public int[] ReadArrayInt(string key, string splitter = "|")
        {
            string[] temp = ReadArray(key, splitter);
            int[] result = StringToInt(temp, 0);
            return result;
        }

        public List<string> ReadList(string key, string splitter = "|")
        {
            return _ReadList(Section, key, splitter);
        }

        public List<int> ReadListInt(string key, string splitter = "|")
        {
            List<int> result = new List<int>();
            result.AddRange(ReadArrayInt(key, splitter));
            return result;
        }


        public double[] ReadArrayDouble(string key, string splitter = "|")
        {
            string[] temp = ReadArray(key, splitter);
            double[] result = StringToDouble(temp, 0);
            return result;
        }

        public string[] ReadItems(TextBox obj)
        {
            string[] result = _ReadArray(Section, obj.Name + "_items");
            obj.Lines = result;
            return result;
        }

        public List<string> ReadItemList(TextBox obj)
        {
            string[] result = _ReadArray(Section, obj.Name + "_items");
            obj.Lines = result;
            if (result != null && result.Length > 0)
            {
                return result.ToList();
            }
            else
            {
                return null;
            }
        }

        public string[] ReadItems(ComboBox obj)
        {
            string[] result = _ReadArray(Section, obj.Name + "_items");
            obj.Items.AddRange(result);
            return result;
        }

        public string[] ReadItems(ListBox obj)
        {
            string[] result = _ReadArray(Section, obj.Name + "_items");
            obj.Items.AddRange(result);
            return result;
        }

        public List<string> ReadItemList(ListBox obj)
        {
            string[] result = _ReadArray(Section, obj.Name + "_items");
            obj.Items.AddRange(result);
            return result.ToList();
        }

        public string ReadCrypt(string key, string cryptKey = "", string dvalue = "")
        {
            if (cryptKey == "")
            {
                cryptKey = CryptKey;
            }
            string result = _ReadCrypt(Section, key, cryptKey, dvalue);
            return result;
        }

        #endregion

        #region Read User Section
        public string R(string section, string key)
        {
            return _Read(section, key);
        }

        public string R(string section, string key, string dvalue)
        {
            return _Read(section, key, dvalue);
        }

        public string SRead(string section, string key)
        {
            return _Read(section, key);
        }

        public string SRead(string section, string key, string dvalue)
        {
            return _Read(section, key, dvalue);
        }

        public string SRead(string section, TextBox obj, string dvalue = "")
        {
            string result = _Read(section, obj.Name, dvalue);
            obj.Text = result;
            return result;
        }

        public string SRead(string section, ComboBox obj, string dvalue = "")
        {
            string result = _Read(section, obj.Name, dvalue);
            obj.Text = result;
            return result;
        }

        public string SRead(string section, DateTimePicker obj, string dvalue = "")
        {
            string result = _Read(section, obj.Name, dvalue);
            try
            {
                obj.Text = result;
            }
            catch (Exception)
            {
                return obj.Text;
            }
            return result;
        }

        public int SRead(string section, NumericUpDown obj, int dvalue = 0)
        {
            int result = _Read(section, obj.Name, dvalue);
            try
            {
                obj.Value = result;
            }
            catch
            {
                return (int)obj.Value;
            }
            return result;
        }

        public bool SRead(string section, CheckBox obj, bool dvalue = false)
        {
            bool result = _Read(section, obj.Name, dvalue);
            obj.Checked = result;
            return result;
        }

        public bool SRead(string section, RadioButton obj, bool dvalue = false)
        {
            bool result = _Read(section, obj.Name, dvalue);
            obj.Checked = result;
            return result;
        }

        public bool SReadBool(string section, string key, bool dvalue = false)
        {
            return _Read(section, key, dvalue);
        }

        public int SReadInt(string section, string key, int dvalue = 0)
        {
            return _Read(section, key, dvalue);
        }

        public double SReadDouble(string section, string key, double dvalue = 0)
        {
            return _Read(section, key, dvalue);
        }

        public string[] SReadArray(string section, string key, string splitter = "|")
        {
            return _ReadArray(section, key, splitter);
        }


        public int[] SReadArrayInt(string section, string key, string splitter = "|")
        {
            string[] temp = SReadArray(section, key, splitter);
            int[] result = StringToInt(temp, 0);
            return result;
        }

        public List<string> SReadList(string section, string key, string splitter = "|")
        {
            return _ReadList(section, key, splitter);
        }


        public List<int> SReadListInt(string section, string key, string splitter = "|")
        {
            List<int> result = new List<int>();
            result.AddRange(SReadArrayInt(section, key, splitter));
            return result;
        }

        public double[] SReadArrayDouble(string section, string key, string splitter = "|")
        {
            string[] temp = SReadArray(section, key, splitter);
            double[] result = StringToDouble(temp, 0);
            return result;
        }

        public string[] SReadItems(string section, ComboBox obj)
        {
            string[] result = _ReadArray(section, obj.Name + "_items");
            obj.Items.AddRange(result);
            return result;
        }

        public string[] SReadItems(string section, ListBox obj)
        {
            string[] result = _ReadArray(section, obj.Name + "_items");
            obj.Items.AddRange(result);
            return result;
        }
        #endregion

        #region Write
        #region Write Base
        private string _Write(string section, string key, string value)
        {
            string result;
            if (section == "") section = Section;
            try
            {
                WritePrivateProfileString(section, key, value, FilePath);
                result = value;
            }
            catch
            {
                result = "";
            }

            return result;
        }

        private int _Write(string section, string key, int value)
        {
            _Write(section, key, value.ToString());
            return value;
        }

        private double _Write(string section, string key, double value)
        {
            _Write(section, key, value.ToString());
            return value;
        }

        private bool _Write(string section, string key, bool value)
        {
            _Write(section, key, value.ToString());
            return value;
        }

        private string[] _WriteArray(string section, string key, string[] items, string splitter = "|")
        {
            _Write(section, key, string.Join(splitter, items));
            return items;
        }

        private List<string> _WriteList(string section, string key, List<string> items, string splitter = "|")
        {
            _Write(section, key, string.Join(splitter, items));
            return items;
        }

        private string _WriteCrypt(string section, string key, string value, string cryptKey = "")
        {
            if (cryptKey == "")
            {
                cryptKey = CryptKey;
            }
            string temp = Encrypt(value, cryptKey);
            _Write(section, key, temp);
            return value;
        }
        #endregion

        #region Write Default Section
        public string Write(string key, string value)
        {
            return _Write(Section, key, value);
        }

        public int Write(string key, int value)
        {
            return _Write(Section, key, value);
        }

        public double Write(string key, double value)
        {
            return _Write(Section, key, value);
        }

        public bool Write(string key, bool value)
        {
            return _Write(Section, key, value);
        }

        // Input Field
        public string Write(TextBox obj, string key = "")
        {
            if (key == "")
            {
                _Write(Section, obj.Name, obj.Text);
            }
            else
            {
                _Write(Section, key, obj.Text);
            }
            return obj.Text;
        }

        public string Write(ComboBox obj, string key = "")
        {
            if (key == "")
            {
                _Write(Section, obj.Name, obj.Text);
            }
            else
            {
                _Write(Section, key, obj.Text);
            }
            return obj.Text;
        }

        public string Write(ListBox obj, string key = "")
        {
            if (key == "")
            {
                _Write(Section, obj.Name, obj.Text);
            }
            else
            {
                _Write(Section, key, obj.Text);
            }
            return obj.Text;
        }

        public string Write(DateTimePicker obj, string key = "")
        {
            if (key == "")
            {
                _Write(Section, obj.Name, obj.Text);
            }
            else
            {
                _Write(Section, key, obj.Text);
            }
            return obj.Text;
        }

        public int Write(NumericUpDown obj, string key = "")
        {
            if (key == "")
            {
                _Write(Section, obj.Name, (int)obj.Value);
            }
            else
            {
                _Write(Section, key, (int)obj.Value);
            }
            return (int)obj.Value;
        }

        public bool Write(CheckBox obj, string key = "")
        {
            if (key == "")
            {
                _Write(Section, obj.Name, obj.Checked);
            }
            else
            {
                _Write(Section, key, obj.Checked);
            }
            return obj.Checked;
        }

        public bool Write(RadioButton obj, string key = "")
        {
            if (key == "")
            {
                _Write(Section, obj.Name, obj.Checked);
            }
            else
            {
                _Write(Section, key, obj.Checked);
            }
            return obj.Checked;
        }

        // Data Type
        public int WriteInt(string key, int value)
        {
            return _Write(Section, key, value);
        }

        public double WriteDouble(string key, double value)
        {
            return _Write(Section, key, value);
        }

        public bool WriteBool(string key, bool value)
        {
            return _Write(Section, key, value);
        }

        public string[] WriteArray(string key, string[] items, string splitter = "|")
        {
            return _WriteArray(Section, key, items, splitter);
        }

        public List<string> WriteList(string key, List<string> items, string splitter = "|")
        {
            return _WriteList(Section, key, items, splitter);
        }

        public string[] WriteItems(TextBox obj)
        {
            string[] items = obj.Lines;
            _WriteArray(Section, obj.Name + "_items", items);
            return items;
        }

        public List<string> WriteItemList(TextBox obj)
        {
            string[] items = obj.Lines;
            _WriteArray(Section, obj.Name + "_items", items);
            return items.ToList();
        }

        public string[] WriteItems(ComboBox obj)
        {
            //List<string> items = new List<string>();
            //items.AddRange(obj.Items.Cast<string>());
            string[] items = new string[obj.Items.Count];
            obj.Items.CopyTo(items, 0);
            _WriteArray(Section, obj.Name + "_items", items);
            return items;
        }

        public List<string> WriteItemList(ComboBox obj)
        {
            //List<string> items = new List<string>();
            //items.AddRange(obj.Items.Cast<string>());
            string[] items = new string[obj.Items.Count];
            obj.Items.CopyTo(items, 0);
            _WriteArray(Section, obj.Name + "_items", items);
            return items.ToList();
        }

        public string[] WriteItems(ListBox obj)
        {
            string[] items = new string[obj.Items.Count];
            obj.Items.CopyTo(items, 0);
            _WriteArray(Section, obj.Name + "_items", items);
            return items;
        }

        public string WriteCrypt(string key, string value)
        {
            _WriteCrypt(Section, key, value);
            return value;
        }

        #endregion

        #region Write User Section
        public string W(string section, string key, string value)
        {
            return _Write(section, key, value);
        }

        public string SWrite(string section, string key, string value)
        {
            return _Write(section, key, value);
        }

        public int SWrite(string section, string key, int value)
        {
            return _Write(section, key, value);
        }

        public double SWrite(string section, string key, double value)
        {
            return _Write(section, key, value);
        }

        public bool SWrite(string section, string key, bool value)
        {
            return _Write(section, key, value);
        }

        public string SWrite(string section, TextBox obj)
        {
            _Write(section, obj.Name, obj.Text);
            return obj.Text;
        }

        public string SWrite(string section, ComboBox obj)
        {
            _Write(section, obj.Name, obj.Text);
            return obj.Text;
        }

        public string SWrite(string section, DateTimePicker obj)
        {
            _Write(section, obj.Name, obj.Text);
            return obj.Text;
        }

        public int SWrite(string section, NumericUpDown obj)
        {
            _Write(section, obj.Name, (int)obj.Value);
            return (int)obj.Value;
        }

        public bool SWrite(string section, CheckBox obj)
        {
            _Write(section, obj.Name, obj.Checked);
            return obj.Checked;
        }

        public bool SWrite(string section, RadioButton obj)
        {
            _Write(section, obj.Name, obj.Checked);
            return obj.Checked;
        }

        public int SWriteInt(string section, string key, int value)
        {
            return _Write(section, key, value);
        }

        public double SWriteDouble(string section, string key, double value)
        {
            return _Write(section, key, value);
        }

        public bool SWriteBool(string section, string key, bool value)
        {
            return _Write(section, key, value);
        }

        public string[] SWriteArray(string section, string key, string[] items)
        {
            return _WriteArray(section, key, items);
        }

        public List<string> SWriteList(string section, string key, List<string> items)
        {
            return _WriteList(section, key, items);
        }

        public string[] SWriteItems(string section, ComboBox obj)
        {
            string[] items = new string[obj.Items.Count];
            obj.Items.CopyTo(items, 0);
            _WriteArray(section, obj.Name + "_items", items);
            return items;
        }

        public string[] SWriteItems(string section, ListBox obj)
        {
            string[] items = new string[obj.Items.Count];
            obj.Items.CopyTo(items, 0);
            _WriteArray(section, obj.Name + "_items", items);
            return items;
        }
        #endregion

        #endregion
    }
}
