using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WinFormBlackListDump
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void deleteFile(string filepath)
        {
            if (File.Exists(filepath))
            {
                try
                {
                    File.Delete(filepath);
                }
                catch (Exception ex) { MessageBox.Show(ex.ToString(), "Ошибка удаления файла"); }
            }
        
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<string> filepath;
            openFileDialog1.Filter = "xml files (*.xml)|*.xml";
            
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //FileInfo fi = new FileInfo(openFileDialog1.FileName);  
                
                toolStripStatusLabel1.Text = "Обрабатываю файл";
                Application.DoEvents();
                filepath = DumpHelperLibrary.Txt.CreateTxtFile(DumpHelperLibrary.Txt.ReadXmlFile(openFileDialog1.FileName));
                deleteFile(openFileDialog1.FileName);
                
                toolStripStatusLabel1.Text = "Сохраняю.";
                Application.DoEvents();
                foreach (string f in filepath)
                {
                    DumpHelperLibrary.Excel.ExportTxtToExcel(f);
                }
                
                toolStripStatusLabel1.Text = "Сохранил.";
                foreach (string f in filepath)
                {
                    deleteFile(f);
                }
            }
            
        }        
    }
}

