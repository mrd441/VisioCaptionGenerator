using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using IVisio = Microsoft.Office.Interop.Visio;
using System.IO;

namespace VisioCaptionGenerator
{
    public partial class Form1 : Form
    {
        public struct fileListElement
        {
            public string fileName { get; set; }
            public string make { get; set; }
            public string makeDate { get; set; }
            public string check { get; set; }
            public string checkDate { get; set; }
            public string tpType { get; set; }
        }

        List<fileListElement> fileList;

        public Form1()
        {
            InitializeComponent();

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);

            fileList = new List<fileListElement>();
            fileListElementBindingSource.DataSource = fileList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                string templateFileName = Directory.GetCurrentDirectory() + "\\template.vss";
                string newFullfileName = Directory.GetCurrentDirectory() + "\\template1.vss";
                if (!File.Exists(templateFileName))                               
                    throw new Exception("Не найден шаблон выходного файла: " + templateFileName);

                //string fileName = resultList.Last().fileName;
                //string fullFileName = listBox1.Items[resulFileNumber].ToString();
                //string fileExt = fileName.Split('.').Last();
                //string newFileName = fileName.Replace("_schema", "");
                //string newFilePath = fullFileName.Replace(fileName, "result\\");
                //string newFullfileName = newFilePath + newFileName;

                IVisio.ApplicationClass visapp = new IVisio.ApplicationClass();
                IVisio.Document doc = visapp.Documents.Open(templateFileName);
                IVisio.Page page = doc.Pages[1];
                IVisio.Master visioRectMaster = doc.Masters.get_ItemU("Sheet.1");
                
                IVisio.Shape visioRectShape = page.Drop(visioRectMaster, 4.25, 5.5);
                visioRectShape.Text = @"Rectangle text.";

                doc.SaveAs(newFullfileName);
                doc.Close();
                visapp.Quit();

            }
            catch(Exception ex)
            {
                LogTextEvent(Color.Red, ex.Message);
            }
                        
        }

        public void LogTextEvent( Color TextColor, string EventText)
        {
            RichTextBox TextEventLog = logBox;
            if (TextEventLog.InvokeRequired)
            {
                TextEventLog.BeginInvoke(new Action(delegate {
                    LogTextEvent(TextColor, EventText);
                }));
                return;
            }

            string nDateTime = DateTime.Now.ToString("hh:mm:ss tt") + " - ";

            // color text.
            TextEventLog.SelectionStart = TextEventLog.Text.Length;
            TextEventLog.SelectionColor = TextColor;

            // newline if first line, append if else.
            if (TextEventLog.Lines.Length == 0)
            {
                TextEventLog.AppendText(nDateTime + EventText);
                TextEventLog.ScrollToCaret();
                TextEventLog.AppendText(System.Environment.NewLine);
            }
            else
            {
                TextEventLog.AppendText(nDateTime + EventText + System.Environment.NewLine);
                TextEventLog.ScrollToCaret();
            }
        }

        void Form1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        void Form1_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
            {
                if (file.Contains(".kml"))
                {
                    string fileName = file;
                    int slashPos = file.LastIndexOf('\\');
                    if (slashPos != -1)
                        fileName = file.Substring(slashPos + 1, file.Length - slashPos - 1);
                    fileListElement fle = new fileListElement();
                    fle.fileName = fileName;
                    if (!fileList.Contains(fle))
                    {
                        fileList.Add(fle);
                        LogTextEvent(Color.Black, "Добавлен файл " + file);
                    }
                    else
                        LogTextEvent(Color.Red, "Файл с таким название уже добавлен: " + file);
                }
                else
                    LogTextEvent(Color.Red, "Расширение файла должно быть KLM: " + file);
            }
            fileListElementBindingSource.ResetBindings(true);
        }
    }
}
