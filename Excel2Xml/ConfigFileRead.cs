using System.IO;
using System.Text;

namespace Excel2Xml
{
    class ConfigFileRead
    {
        static string content = "";
        StringBuilder strbuffer = new StringBuilder(2048);

        public StringBuilder StrBuffer
        {
            get { return strbuffer; }
        }

        public void Read(string config)
        {
            content = File.ReadAllText(config);
            strbuffer.Append(content);
        }

        public void Replace(string collumnName, string value)
        {
            strbuffer.Replace(collumnName, value);
        }

        public void Reset()
        {
            strbuffer.Remove(0, strbuffer.Length);
            strbuffer.Append(content);
        }
    }
}
