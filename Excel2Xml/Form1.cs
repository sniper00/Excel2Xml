using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace Excel2Xml
{
    enum OutputType
    {
        XML,
        JSON
    }

    public partial class Form1 : Form
    {
        string ExcelPath = "";
        string ConfigPath = "";
        delegate void DoDataDelegate();

        bool bProcessing = false;

        OutputType OType;

        public Form1()
        {
            InitializeComponent();
        }

        private void Label_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Excel文件|*.xls;*.xlsx";
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                ExcelPath = System.IO.Path.GetDirectoryName(dialog.FileName);
            
                textBox1.Text = dialog.FileName;

                ExcelRead er = new ExcelRead();
                er.OpenFile(textBox1.Text, true);
                er.CloseExcel();
                comboBox1.Items.Clear();
                comboBox1.Text = "";

                foreach (var name in  er.SheetNameList)
                {
                    comboBox1.Items.Add(name);
                }
                comboBox1.SelectedIndex = 0;

            }
        }


        private void button3_Click(object sender, EventArgs e)
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.InitialDirectory = ExcelPath;
            dialog.Filter = "xml文件(*.xml)|*.xml|json文件(*.json)|*.json";
            dialog.AddExtension = true;
            if (textBox1.Text.Length > 0)
            {
                var saveName = Path.GetFileNameWithoutExtension(textBox1.Text);
                dialog.FileName = saveName;
            }

            if (dialog.ShowDialog() == DialogResult.OK)
            {
                string ext = Path.GetExtension(dialog.FileName);
                if(ext.Contains("xml"))
                {
                    OType = OutputType.XML;
                }
                else if(ext.Contains("json"))
                {
                    OType = OutputType.JSON;
                }          
                textBox3.Text = dialog.FileName;
            }
        }

        void DoData()
        {
            bProcessing = true;


            if (progressBar1.InvokeRequired)
            {
                DoDataDelegate d = DoData;
                progressBar1.Invoke(d);
            }
            else
            {
                ExcelRead er = new ExcelRead();
                er.OpenFile(textBox1.Text, true);

                string sheetName = comboBox1.SelectedItem.ToString();
                DataTable dt = er.ReadTable(sheetName);
                er.CloseExcel();

                progressBar1.Maximum = dt.Rows.Count;

                List <string> columnsName = new List<string>();
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    columnsName.Add(dt.Columns[j].ColumnName);
                }

                ConfigFileRead c = new ConfigFileRead();
                c.Read(ConfigPath);

                StreamWriter sw = new StreamWriter(textBox3.Text, false, Encoding.GetEncoding("UTF-8"));

                if(OType == OutputType.XML)
                {
                    var saveName = Path.GetFileNameWithoutExtension(textBox1.Text);
                    sw.Write("<?xml version=\"1.0\" encoding=\"utf-8\" ?>");
                    sw.Write("\r\n");
                    sw.Write("<"+ saveName+">");
                    sw.Write("\r\n");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {

                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            c.Replace(string.Format("#{0}#", columnsName[j]), dt.Rows[i][j].ToString());
                        }

                        progressBar1.Value = i + 1;
                        Application.DoEvents();

                        c.StrBuffer.Append("\r\n");
                        sw.Write(c.StrBuffer.ToString());
                        c.Reset();
                    }
                    sw.Write("</" + saveName + ">");
                }
                else if(OType == OutputType.JSON)
                {
                    var saveName = Path.GetFileNameWithoutExtension(textBox1.Text);
                    sw.Write("{\""+ saveName+"\":[");
                    for (int i = 0; i < dt.Rows.Count; i++)
                    {       
                        if(i>0)
                        {
                            sw.Write(",");
                        }
                        for (int j = 0; j < dt.Columns.Count; j++)
                        {
                            c.Replace(string.Format("#{0}#", columnsName[j]), dt.Rows[i][j].ToString());
                        }
                        progressBar1.Value = i + 1;
                        Application.DoEvents();
                        sw.Write(c.StrBuffer.ToString().Trim());
                        c.Reset();
                    }
                    sw.Write("]}");
                }

             
                sw.Flush();
                sw.Close();

                bProcessing = false;
                MessageBox.Show("处理完成！", "ok", MessageBoxButtons.OK);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ConfigPath = ExcelPath+"\\"+ Path.GetFileNameWithoutExtension(textBox1.Text);
            ConfigPath += ".ec";

            if (!File.Exists(this.textBox1.Text) || !File.Exists(ConfigPath))
            {
                MessageBox.Show("Excel 文件 或 配置文件 不存在！","Error", MessageBoxButtons.OK);
                return;
            }

            if(bProcessing)
            {
                MessageBox.Show("任务正在处理", "Warning", MessageBoxButtons.OK);
                return;
            }

            progressBar1.Value = 0;
            Thread thread = new Thread(DoData);
            thread.IsBackground = true;
            thread.Start();
        }
    }
}
