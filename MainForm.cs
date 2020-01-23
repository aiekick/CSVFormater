/*
MIT License

Copyright (c) 2020 Aiekick

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/


using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Threading;

namespace CSVFormater
{
    #region Public Delegates

    // delegates used to call MainForm functions from worker thread
    public delegate void DelegateAddConsoleMsg(String s);
    public delegate void DelegateSetProgressValue(int val);
    public delegate void DelegateSetProgressBisValue(int val);
    public delegate void DelegateSetProgressCalcul1Value(int val);
    public delegate void DelegateSetProgressCalcul2Value(int val);
    public delegate void DelegateThreadFinished();

    #endregion

    /// <summary>
	/// Description of MainForm.
	/// </summary>
	public partial class MainForm : Form
	{
        #region Members For Thread

        // worker thread
        Thread m_WorkerThread;

        // events used to stop worker thread
        ManualResetEvent m_EventStopThread;
        ManualResetEvent m_EventThreadStopped;

        string configBoxText;
        string pathBoxText;
        string[] pathArr;
        
        // Delegate instances used to cal user interface functions 
        // from worker thread:
        public DelegateAddConsoleMsg m_DelegateAddConsoleMsg;
        public DelegateSetProgressValue m_DelegateSetProgressValue;
        public DelegateSetProgressBisValue m_DelegateSetProgressBisValue;
        public DelegateSetProgressCalcul1Value m_DelegateSetProgressCalcul1Value;
	    public DelegateSetProgressCalcul1Value m_DelegateSetProgressCalcul2Value;
	    public DelegateThreadFinished m_DelegateThreadFinished;

        StreamWriter Log;
        
        #endregion

        public ConfigLua cLua;

		public MainForm()
		{
			//
			// The InitializeComponent() call is required for Windows Forms designer support.
			//
			InitializeComponent();

            UpdateConfigBox();

            // initialize delegates
            m_DelegateAddConsoleMsg = new DelegateAddConsoleMsg(this.AddConsoleMsg);
            m_DelegateSetProgressValue = new DelegateSetProgressValue(this.SetProgressValue);
            m_DelegateSetProgressBisValue = new DelegateSetProgressBisValue(this.SetProgressBisValue);
            m_DelegateSetProgressCalcul1Value = new DelegateSetProgressCalcul1Value(this.SetProgressCalcul1Value);
           	m_DelegateSetProgressCalcul2Value = new DelegateSetProgressCalcul1Value(this.SetProgressCalcul2Value);
           	m_DelegateThreadFinished = new DelegateThreadFinished(this.ThreadFinished);

            // initialize events
            m_EventStopThread = new ManualResetEvent(false);
            m_EventThreadStopped = new ManualResetEvent(false);
            
            cLua = new ConfigLua(m_EventStopThread, m_EventThreadStopped, this);
		}

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void LaunchThread_Click(object sender, EventArgs e)
        {
        	if ( m_WorkerThread == null ) 
        		RunBtn.Text = "GO";
            
        	if (RunBtn.Text == "GO") // on lance le thread
            {
                RunBtn.Text = "STOP";
                RunBtn.Update();
                
                listView1.Items.Clear();

                // reset events
                m_EventStopThread.Reset();
                m_EventThreadStopped.Reset();

                configBoxText = this.ConfigBox.Text;
                pathBoxText = this.PathBox.Text;

                // create worker thread instance
                m_WorkerThread = new Thread(new ThreadStart(this.WorkerThreadFunction));

                m_WorkerThread.Name = "Worker Thread Config Lua";	// looks nice in Output window

                m_WorkerThread.Start();
            }
            else
                if (RunBtn.Text == "STOP") // on stop le thread
                { 
                    RunBtn.Text = "GO";

                    StopThread();
                }
        }

        // OpenFileDialog
        private void OpenFiles_Click(object sender, EventArgs e)
        {
            OpenFileDialog fDialog = new OpenFileDialog();
            
            fDialog.Title = "Open File";
            fDialog.Filter = "CSV Files|*.csv|Office SpreadSheet XMl Files|*.xml";
            fDialog.AddExtension = true;
            fDialog.CheckFileExists = true;
            fDialog.ShowHelp = true;
            fDialog.CheckFileExists = true;
            fDialog.Multiselect = true;
            pathArr = null;
            
            if (fDialog.ShowDialog() == DialogResult.OK)
            {
            	int count = fDialog.FileNames.Length;
            	if ( count > 1 )
            	{
            		string str = count.ToString() + " Files to Parse";
            		PathBox.Text = str;
            		pathArr = fDialog.FileNames;
            		cLua.xlsFormater.MultipleFiles = true;
            	}
            	else 
            	{
                	cLua.xlsFormater.MultipleFiles = false;
            		PathBox.Text = fDialog.FileName.ToString();
            	}
            }
            
        }

        public void UpdateConfigBox()
        {
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            string scriptPath = appPath + "\\scripts";

            DirectoryInfo di = new DirectoryInfo(scriptPath);
            FileInfo[] rgFiles = di.GetFiles("*.lua");

            foreach (FileInfo fi in rgFiles)
            {
                ConfigBox.Items.Add(fi.Name);
            }

            ConfigBox.SelectedIndex = 0;
        }

        public void AddConsoleMsg(string msg)
        {
        	listView1.Items.Add(msg);
        }

        public void SetProgressValue(int val)
        {
            progressBar.Value = val;
        }

        public void SetProgressBisValue(int val)
        {
            progressBarBis.Value = val;
        }
        
        public void SetProgressCalcul1Value(int val)
        {
            progressCalcul1Bar.Value = val;
        }
        
        public void SetProgressCalcul2Value(int val)
        {
            progressCalcul2Bar.Value = val;
        }
        
        // Worker thread function.
        // Called indirectly from btnStartThread_Click
        private void WorkerThreadFunction()
        {
            if ( pathArr != null )
            {
            	cLua.RunConfig(configBoxText, pathArr);
            }
            else
            {
           		cLua.RunConfig(configBoxText, pathBoxText);
            }
        }

        // Stop worker thread if it is running.
        // Called when user presses Stop button of form is closed.
        private void StopThread()
        {
            if (m_WorkerThread != null && m_WorkerThread.IsAlive)  // thread is active
            {
                // set event "Stop"
                m_EventStopThread.Set();

                // wait when thread  will stop or finish
                while (m_WorkerThread.IsAlive)
                {
                    // We cannot use here infinite wait because our thread
                    // makes syncronous calls to main form, this will cause deadlock.
                    // Instead of this we wait for event some appropriate time
                    // (and by the way give time to worker thread) and
                    // process events. These events may contain Invoke calls.
                    if (WaitHandle.WaitAll(
                        (new ManualResetEvent[] { m_EventThreadStopped }),
                        100,
                        true))
                    {
                        break;
                    }

                    Application.DoEvents();
                }
            }

            ThreadFinished();		// set initial state of buttons
        }

        // Set initial state of controls.
        // Called from worker thread using delegate and Control.Invoke
        private void ThreadFinished()
        {
            RunBtn.Text = "GO";
            m_WorkerThread = null;
        }

        private void progressBar_Click(object sender, EventArgs e)
        {

        }

        private void EditFile_Click(object sender, EventArgs e)
        {
            string appPath = Path.GetDirectoryName(Application.ExecutablePath);
            string scriptPath = appPath + "\\scripts\\";
            string filepath = scriptPath + ConfigBox.Text;

            System.Diagnostics.Process.Start(filepath);
        }
		
        public bool IsAppendOnCurrentExcelFile()
        {
        	return checkBox1.Checked;
        }
        
        // sauve le contenu de la liste box dnas un fichier
		void Save_Click(object sender, EventArgs e)
		{
			int countMsg = listView1.Items.Count;
			if ( countMsg == 0 ) return;
			
			string LogFilePath = "Save " + DateTime.Now.ToString() + ".log";
			LogFilePath = LogFilePath.Replace('/', '-');
			LogFilePath = LogFilePath.Replace(' ', '_');
			LogFilePath = LogFilePath.Replace(':', '-');

			if (!File.Exists(LogFilePath))
            {
                Log = new StreamWriter(LogFilePath);
            }
            else
            {
                Log = File.AppendText(LogFilePath);
            }

            for ( int i=0; i< countMsg; i++ )
            {
            	Log.WriteLine(listView1.Items[i].Text);
            }
         
            Log.Close();
		}
		
		void ListView1RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
		{
			
		}
	}
}
