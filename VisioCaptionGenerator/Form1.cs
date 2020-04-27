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
            public string filePath { get; set; }
        }

        List<fileListElement> fileList;
        IVisio.InvisibleApp visapp;
        public Form1()
        {
            InitializeComponent();

            this.AllowDrop = true;
            this.DragEnter += new DragEventHandler(Form1_DragEnter);
            this.DragDrop += new DragEventHandler(Form1_DragDrop);

            fileList = new List<fileListElement>();
            fileListElementBindingSource.DataSource = fileList;
            //dataGridView1.DataSource = fileList;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GenerateVSS();
        }

        public void GenerateVSS()
        {
            visapp = new IVisio.InvisibleApp();            
            visapp.Visible = false;
            foreach (fileListElement fileEl in fileList)
            {
                try
                {
                    string templateFileName = Directory.GetCurrentDirectory() + "\\template.vss";
                    string newFullfileName = fileEl.filePath + "\\result\\" + fileEl.fileName.Replace(".kml", ".vss");
                    
                    if (!File.Exists(templateFileName))
                        throw new Exception("Не найден шаблон выходного файла: " + templateFileName);

                    Directory.CreateDirectory(fileEl.filePath + "\\result\\");

                    string fileName = fileEl.fileName;
                    string caption = "Поопорная схема ВЛ 0,4 кВ от ";
                    string firstTmp = "";
                    int tpPos = fileName.IndexOf("ТП");
                    if (tpPos >= 0)
                    {
                        firstTmp = fileName.Substring(tpPos, fileName.Length - tpPos).Replace('_', '/').Replace(".kml", " ") + "кВа ";
                        string secondTmp = fileName.Substring(0, tpPos - 1);
                        fileName = firstTmp + secondTmp;
                    }
                    else
                        throw new Exception("Не корректное название файла: " + fileName);
                    caption = caption + fileName;

                    IVisio.Document doc = visapp.Documents.Open(templateFileName );// (short)IVisio.VisOpenSaveArgs.visAddHidden + (short)IVisio.VisOpenSaveArgs.visOpenNoWorkspace);
                    IVisio.Page page = doc.Pages[1];                    

                    IVisio.Shape visioRectMaster = page.Shapes.get_ItemU("Sheet.1");
                    visioRectMaster.Text = caption;
                    
                    visioRectMaster = page.Shapes.get_ItemU("Sheet.509");
                    visioRectMaster.Text = caption;

                    visioRectMaster = page.Shapes.get_ItemU("Sheet.496");
                    visioRectMaster.Text = fileEl.make;

                    visioRectMaster = page.Shapes.get_ItemU("Sheet.508");
                    visioRectMaster.Text = fileEl.makeDate;

                    visioRectMaster = page.Shapes.get_ItemU("Sheet.495");
                    visioRectMaster.Text = fileEl.check;

                    visioRectMaster = page.Shapes.get_ItemU("Sheet.507");
                    visioRectMaster.Text = fileEl.checkDate;

                    IVisio.Master aMaster;
                    switch (fileEl.tpType)
                    {
                        case "Столбовая":
                            aMaster = doc.Masters.get_ItemU(@"СТП");
                            visioRectMaster = page.Drop(aMaster, 6.3400314961, 5.4108622047);
                            //visioRectMaster.tra
                            visioRectMaster.Shapes[2].Text = "С" + firstTmp;
                            break;
                        case "Мачтовая":
                            aMaster = doc.Masters.get_ItemU(@"МТП");
                            visioRectMaster = page.Drop(aMaster, 6.3976378, 5.4108622047);
                            visioRectMaster.Shapes[2].Text = "М" + firstTmp;
                            break;
                        case "Закрытая":
                            aMaster = doc.Masters.get_ItemU(@"ЗТП");
                            visioRectMaster = page.Drop(aMaster, 6.313976378, 5.23622);
                            visioRectMaster.Shapes[2].Text = "З" + firstTmp;
                            break;
                        case "Комплектная":
                            aMaster = doc.Masters.get_ItemU(@"КТП");
                            visioRectMaster = page.Drop(aMaster, 6.4422795276, 5.4108622047);
                            visioRectMaster.Shapes[2].Text = "К" + firstTmp;
                            break;
                    }
                   
                    doc.SaveAs(newFullfileName);
                    doc.Close();
                    //visapp.Quit();

                }
                catch (Exception ex)
                {
                    LogTextEvent(Color.Red, ex.Message);                    
                }
            }
            foreach (IVisio.Document aDoc in visapp.Documents)
                aDoc.Close();
            visapp.Quit();
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
                    string filePath = "";
                    int slashPos = file.LastIndexOf('\\');
                    if (slashPos != -1)
                    {
                        fileName = file.Substring(slashPos + 1, file.Length - slashPos - 1);
                        filePath = file.Substring(0,slashPos);
                    }
                    fileListElement fle = new fileListElement();
                    fle.fileName = fileName;
                    fle.filePath = filePath;
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

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            var value = dataGridView1.CurrentCell.Value;
            
            if (e.RowIndex >= 0 & value != null)
            {
                fileListElement fileEl = fileList[e.RowIndex];
                string valStr = value.ToString();
                switch (e.ColumnIndex)
                {
                    case 1:
                        fileEl.make = valStr;
                        break;
                    case 2:
                        fileEl.makeDate = valStr;
                        break;
                    case 3:
                        fileEl.check = valStr;
                        break;
                    case 4:
                        fileEl.checkDate = valStr;
                        break;
                    case 5:
                        fileEl.tpType = valStr;
                        break;
                }
                fileList[e.RowIndex] = fileEl;
            }
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex>=0)
            {
                var value = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value;
            }
        }

        private void dataGridView1_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (e.RowIndex >= 0 & fileList.Count> e.RowIndex)
            {
                fileListElement fileEl = fileList[e.RowIndex];
                string valStr;
                switch (e.ColumnIndex)
                {
                    case 1:
                        valStr = fileEl.make;
                        break;
                    case 2:
                        valStr = fileEl.makeDate;
                        break;
                    case 3:
                        valStr = fileEl.check;
                        break;
                    case 4:
                        valStr = fileEl.checkDate;
                        break;
                    case 5:
                        valStr = fileEl.tpType;
                        break;
                }
                dataGridView1.CurrentCell.Value = fileEl;
            }
        }          
    }
}
