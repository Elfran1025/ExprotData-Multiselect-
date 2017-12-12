using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ExprotData
{
    public partial class Form1 : Form
    {
        public static Form1 form1;
        string exportPath = System.AppDomain.CurrentDomain.SetupInformation.ApplicationBase+ "Excel导出文件夹" + "\\"; //生成CSV路径
                //exportPath = exportPath + ;
        public Form1()
        {
            InitializeComponent();
            form1 = this;

            //Shown += new EventHandler(Form1_Shown);

            // To report progress from the background worker we need to set this property
            backgroundWorker1.WorkerReportsProgress = true;

            // This event will be raised on the worker thread when the worker starts
            backgroundWorker1.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);

            // This event will be raised when we call ReportProgress
            backgroundWorker1.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            CheckForIllegalCrossThreadCalls = false;
        }
        void Form1_Shown(object sender, EventArgs e)
        {
            // Start the background worker
          
        }

        // On worker thread so do our thing!
        void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            // Your background task goes here
            ExportJSON ej = new ExportJSON();
            OpenFileDialog ofd =(OpenFileDialog) e.Argument;
            string[] files = ofd.FileNames;
            string[] safename = ofd.SafeFileNames;
            progressBar1.Maximum = files.Length;
            progressBar1.Value = 0;
            for (int i = 0; i < files.Length; i++)
            {
                if (!files[i].Equals(null) && !files[i].Equals(""))
                {
                    label1.Text = "文件总数" + files.Length + "个，" + "正在导出第" + (i + 1) + "个，" + "剩余" + (files.Length - i - 1) + "个";
                    Console.WriteLine(DateTime.Now + "文件总数" + files.Length + "个，" + "正在导出第" + (i + 1) + "个，" + "剩余" + (files.Length - i - 1) + "个");
                    string ExportPath = ej.ExportJSON_Method(files[i], safename[i]);
                    listBox1.Items.Add(DateTime.Now + "\t" + safename[i] + "\t导出完成");
                    progressBar1.Value++;
                    //listBox1.Items.Add(ExportPath);

                }


            }
            label1.Text = "文件总数" + files.Length + "个，" + "已全部完成";
            label2.Text = "全部导出完成";
            Console.WriteLine(DateTime.Now + "文件总数" + files.Length + "个，" + "已全部完成");





            //for (int i = 0; i <= 100; i++)
            //{
            //    // Report progress to 'UI' thread
            //    backgroundWorker1.ReportProgress(i);
            //    // Simulate long task
            //    System.Threading.Thread.Sleep(100);
            //}
        }

        // Back on the 'UI' thread so we can update the progress bar
        void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            // The progress percentage is a property of e
            progressBar1.Value = e.ProgressPercentage;
        }


        private void button1_Click(object sender, EventArgs e)
        {
          
            //string path = "";
            string multiple;
            //path = Application.StartupPath + @"\Documents";
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV|*.csv";
            ofd.DefaultExt = ".csv";
            ofd.Multiselect = true;
            //ofd.InitialDirectory = path;
            ofd.ShowDialog();
            string[] files= ofd.FileNames;
            string[] safename = ofd.SafeFileNames;
            //path = ofd.FileName;
            textBox1.Text = "共" + files.Length + "个文件";
            backgroundWorker1.RunWorkerAsync(ofd);

            








            //listBox1.Items.Add(ofd.SafeFileName+"\t导出完成"+"位置："+ej.ExportJSON_Method(path));


        }

        private void button2_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(exportPath);
        }
    }
}
