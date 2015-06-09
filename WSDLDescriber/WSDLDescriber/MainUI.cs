using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace WSDLDescriber
{
    public partial class MainUI : Form
    {
        private int statusInt;
        private string statusString;
        private Thread thread;
        public MainUI()
        {
            InitializeComponent();
            statusInt = 0;
            statusString = "Beginning the program";
            timer.Interval = 100;
        }

        private void generatorButton_Click(object sender, EventArgs e)
        {
            saveFileDialog.Filter = "Word Document (*.docx, *.doc)| *.docx; *.doc";
            saveFileDialog.DefaultExt = "docx";
            saveFileDialog.ShowDialog();
        }

        private void saveFileDialog_FileOk(object sender, CancelEventArgs e)
        {
            saveFileDialog.Dispose();
            thread = new Thread(StartProcess);
            timer.Start();
            if (timer.Enabled)
            {
                thread.Start();
            }
        }
        private void StartProcess()
        {
            statusInt = 0;
            
            ApplicationManager manager = new ApplicationManager();
            manager.DescribeWSDLInDocument(urlBox.Text, saveFileDialog.FileName, authorNameBox.Text, ref statusInt, ref statusString);
            statusInt = 0;
            
        }

        private void timer_Tick(object sender, EventArgs e)
        {
            progressBar.Value = statusInt;
            statusBox.Text = statusString;
            generatorButton.Enabled = false;
            if (!thread.IsAlive)
            {
                generatorButton.Enabled = true;
                timer.Stop();
            }
        }
    }
}
