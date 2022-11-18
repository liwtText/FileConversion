using Common;
using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Controls;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace FileConversion
{
    public partial class Form1 : Form
    {
        private static int threadNum, number, successNum, failNum;
        private Queue<string> filePath = new Queue<string>();
        public Form1()
        {
            InitializeComponent();
            ConfigIni.GetIniVal();
            GetConfig();
        }

        #region 选项框
        private void textBoxDropDown1_Click(object sender, EventArgs e)
        {
            BoxClick(textBoxDropDown1, labelX4);
        }
        private void textBoxDropDown5_Click(object sender, EventArgs e)
        {
            BoxClick(textBoxDropDown5, null);
        }
        private void textBoxDropDown6_Click(object sender, EventArgs e)
        {
            BoxClick(textBoxDropDown6, null);
        }
        private void textBoxDropDown7_Click(object sender, EventArgs e)
        {
            BoxClick(textBoxDropDown7, labelX16, "pdf");
        }
        private void textBoxDropDown4_Click(object sender, EventArgs e)
        {
            BoxClick(textBoxDropDown4, null);
        }
        private void textBoxDropDown2_Click(object sender, EventArgs e)
        {
            BoxClick(textBoxDropDown2, null);
        }
        public void BoxClick(TextBoxDropDown textBox, LabelX label, string format = "")
        {
            FolderBrowserDialog ofd = new FolderBrowserDialog();
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                textBox.Text = ofd.SelectedPath;
                if (label == null) return;
                DirectoryInfo dir = new DirectoryInfo(ofd.SelectedPath);
                int number = 0;
                if (dir.Exists)
                {
                    FileInfo[] files = dir.GetFiles().Where(a => a.Extension.Contains("ppt") || a.Extension.Contains("pptx")).ToArray();
                    number = files.Length;
                }
                label.Text = number.ToString();
            }
        }
        #endregion

        #region 按钮
        private void buttonX3_Click(object sender, EventArgs e)
        {
            SetConfig();
            MessageBox.Show("修改成功");
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            number = 0;
            successNum = 0;
            failNum = 0;
            threadNum = int.Parse(textBoxX7.Text);
            if (true)
            {
                AsposeHelp.min = int.Parse(textBoxX2.Text);
                AsposeHelp.max = int.Parse(textBoxX3.Text);
                AsposeHelp.page = int.Parse(textBoxX1.Text);

                if (AsposeHelp.max < AsposeHelp.min) MessageBox.Show("随机数区间设置错误");
                FileList();
            }
        }

        private void buttonX5_Click(object sender, EventArgs e)
        {
            number = 0;
            successNum = 0;
            failNum = 0;
            threadNum = int.Parse(textBoxX7.Text);
            progressBarX3.Visible = true;
            int size = int.Parse(textBoxX4.Text);
            MultiThreading(size);

        }

        #endregion
        #region 文档处理
        private void ProcessOne()
        {
            while (true)
            {
                if (filePath.Count == 0) break;
                string path = filePath.Dequeue();
                try
                {
                    string fileName = Path.GetFileName(path);
                    string newPath = Path.Combine(ConfigIni.pptSusPath, fileName);
                    AsposeHelp aspose = new AsposeHelp();
                    if (aspose.Help(path, newPath))
                    {
                        successNum++;
                        FileInfo fil = new FileInfo(path);
                        fil.MoveTo(Path.Combine(ConfigIni.originalPath, fileName));
                    }
                    else failNum++;
                }
                catch (Exception ex)
                {
                    LoggerHelper.WriteLog(typeof(Form1), $"文件路径{path}" + ex.ToString());
                }
            }
        }
        public void FileList()
        {
            GetFiles(ConfigIni.pptPath, "ppt");
            progressBarX1.Visible = true;
            InitProgressBar(progressBarX1, 0, filePath.Count);
            if (filePath.Count < threadNum) threadNum = 1;
            for (int i = 0; i < threadNum; i++)
            {
                if (filePath.Count > 0)
                {
                    Task.Factory.StartNew(() => { ProcessOne(); });
                }
            }
            CloseProgress(progressBarX1, labelX13, labelX2);
        }
        public void MultiThreading(int size)
        {
            GetFiles(ConfigIni.pdfPath, "pdf", size);
            InitProgressBar(progressBarX3, 0, filePath.Count);
            if (filePath.Count < threadNum) threadNum = 1;
            for (int i = 0; i < threadNum; i++)
            {
                Task.Factory.StartNew(() => { PdfOrDoc(); });
            }
            CloseProgress(progressBarX3, labelX19, labelX18);
        }
        public void PdfOrDoc()
        {

        }
        public void GetFiles(string path, string format, int size = 0)
        {
            DirectoryInfo dir = new DirectoryInfo(path);
            if (dir.Exists)
            {
                FileInfo[] files = dir.GetFiles().Where(a => a.Extension.Contains(format) || a.Extension.Contains(format + "x")).OrderBy(a => a.Length).ToArray();
                foreach (var item in files)
                {
                    double fileSize = item.Length / 1024 / 1024;
                    if (size > 0 && fileSize >= size)
                    {
                        item.Delete();
                        continue;
                    }
                    filePath.Enqueue(item.FullName);
                }
                number = filePath.Count;
            }
        }
        #endregion

        #region 进度条
        public delegate void dele_showLoad(ProgressBarX progressBar, LabelX lable, LabelX lable1, bool type);
        public void LoadDisplay(ProgressBarX progressBar, LabelX lable, LabelX lable1, bool type)
        {
            if (this.InvokeRequired)
            {
                dele_showLoad delege = new dele_showLoad(LoadDisplay);
                this.Invoke(delege, type);
            }
            else
            {
                if (type)
                {
                    StartProgressNar(progressBar, successNum + failNum, lable);
                    lable1.Text = successNum.ToString();
                }
                else
                {
                    lable.Text = "";
                    progressBar.Visible = false;
                    GC.Collect();
                }
            }
        }
        /// <summary>
        /// 初始化
        /// </summary>
        /// <param name="progressBar"></param>
        /// <param name="minValue"></param>
        /// <param name="maxValue"></param>
        private void InitProgressBar(ProgressBarX progressBar, int minValue, int maxValue)
        {
            if (progressBar == null || minValue < 0 || maxValue < 0 || minValue > maxValue) return;
            progressBar.Value = 0;
            progressBar.Minimum = minValue;
            progressBar.Maximum = maxValue;
        }
        /// <summary>
        /// 加载中
        /// </summary>
        /// <param name="progressBar"></param>
        /// <param name="value"></param>
        /// <param name="lable"></param>
        private void StartProgressNar(ProgressBarX progressBar, int value, LabelX lable)
        {
            try
            {
                if (progressBar == null || lable == null) return;
                //System.Windows.Forms.Application.DoEvents();
                progressBar.Value = value;
                int tmp = value * 100 / progressBar.Maximum;
                lable.Text = "(" + tmp + "%)";
                lable.Refresh();
            }
            catch
            {
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            AsposeHelp.PdfOrDoc(@"E:\Img\新建文件夹\3333.pdf", @"E:\Img\新建文件夹\3333.docx");
        }

        /// <summary>
        /// 任务完成
        /// </summary>
        public void CloseProgress(ProgressBarX progressBar, LabelX lable, LabelX lable1)
        {
            System.Diagnostics.Stopwatch watch = new System.Diagnostics.Stopwatch();
            watch.Start();//开始计时
            Task.Factory.StartNew(() =>
            {
                while (true)
                {
                    LoadDisplay(progressBar, lable, lable1, true);
                    if (successNum + failNum == number)
                    {
                        LoadDisplay(progressBar, lable, lable1, false);
                        break;
                    }
                    Thread.Sleep(1000);
                }
                watch.Stop();//停止计时
                MessageBox.Show("任务完成  耗时:" + (watch.ElapsedMilliseconds / 1000) + "秒");
            });
        }
        #endregion
        #region 添加配置
        public void SetConfig()
        {
            ConfigIni.pptPath = textBoxDropDown1.Text;
            ConfigIni.pptSusPath = textBoxDropDown5.Text;
            ConfigIni.originalPath = textBoxDropDown6.Text;
            ConfigIni.pdfPath = textBoxDropDown7.Text;
            ConfigIni.docPath = textBoxDropDown4.Text;
            ConfigIni.originalPathTwo = textBoxDropDown2.Text;
            ConfigIni.SetIniVal();
        }
        public void GetConfig()
        {
            try
            {
                int number = 0;
                if (!string.IsNullOrEmpty(ConfigIni.pptPath))
                {
                    textBoxDropDown1.Text = ConfigIni.pptPath;
                    DirectoryInfo dir = new DirectoryInfo(ConfigIni.pptPath);
                    if (dir.Exists)
                    {
                        FileInfo[] files = dir.GetFiles().Where(a => a.Extension.Contains("ppt") || a.Extension.Contains("pptx")).ToArray();
                        number = files.Length;
                    }
                    labelX4.Text = number.ToString();
                }
                if (!string.IsNullOrEmpty(ConfigIni.pptSusPath)) textBoxDropDown5.Text = ConfigIni.pptSusPath;
                if (!string.IsNullOrEmpty(ConfigIni.originalPath)) textBoxDropDown6.Text = ConfigIni.originalPath;
                if (!string.IsNullOrEmpty(ConfigIni.pdfPath))
                {
                    textBoxDropDown7.Text = ConfigIni.pdfPath;
                    DirectoryInfo dir = new DirectoryInfo(ConfigIni.pdfPath);
                    if (dir.Exists)
                    {
                        FileInfo[] files = dir.GetFiles().Where(a => a.Extension.Contains("pdf")).ToArray();
                        number = files.Length;
                    }
                    labelX16.Text = number.ToString();
                }
                if (!string.IsNullOrEmpty(ConfigIni.docPath)) textBoxDropDown4.Text = ConfigIni.docPath;
                if (!string.IsNullOrEmpty(ConfigIni.originalPathTwo)) textBoxDropDown2.Text = ConfigIni.originalPathTwo;
            }
            catch (Exception ex)
            {
                LoggerHelper.WriteLog(typeof(Form1), ex.ToString());
            }
        }
        #endregion
    }
}
