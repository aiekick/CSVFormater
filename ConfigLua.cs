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
using LuaInterface;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using System.Xml;

namespace CSVFormater
{
	/// <summary>
	/// Description of ConfigLua.
	/// </summary>
	public class ConfigLua
	{
        #region Members For Threads

        // Main thread sets this event to stop worker thread:
        ManualResetEvent m_EventStop;

        // Worker thread sets this event when it is stopped:
        ManualResetEvent m_EventStopped;

        // Reference to main form used to make syncronous user interface calls:
        MainForm m_form;

        #endregion

        public Lua luaInterpret;
        public XLSFormater xlsFormater;

        // LUA VARIABLE
        string currentLineName;
        string lastLineName;
        string Author;
        string CSVSeparator;
        string FonctionForEachLine;
        string FonctionForEndFile;
        
        int currentLineIndex;

        StreamWriter Log;
        string LogFilePath;

        bool ErrorProcessus;
        string ErrorProcessString;

        // on va verifer le code lua
        public ConfigLua(ManualResetEvent eventStop, ManualResetEvent eventStopped, MainForm form)
        {
            luaInterpret = new Lua(); // ce charge d'interpreter le fichier lua selctionné
            xlsFormater = new XLSFormater(); // ce charge de manipuler excel

            m_EventStop = eventStop;
            m_EventStopped = eventStopped;

            m_form = form;
		}

        private void DeleteLogFile()
        {
        	string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            LogFilePath = appPath + "\\logfile.txt";
            
            if (File.Exists(LogFilePath))
        		System.IO.File.Delete(LogFilePath);
        }

        private void AddToLogFile(string value)
        {
            if (!File.Exists(LogFilePath))
            {
                Log = new StreamWriter(LogFilePath);
            }
            else
            {
                Log = File.AppendText(LogFilePath);
            }

            // Write to the file:
            string msg = DateTime.Now.ToString() + " +> " + value;
            Log.WriteLine(msg);
        
            // Ajoute le message dnas la console
            // Make synchronous call to main form.
            // MainForm.AddConsoleMsg function runs in main thread.
            // To make asynchronous call use BeginInvoke
            m_form.Invoke(m_form.m_DelegateAddConsoleMsg, new Object[] { msg });

            // Close the stream:
            Log.Close();
        }

        public int RunConfig(string CurrentConfigBoxName, string[] CSVFilePathArr)
        {
        	//////////////////////////////////////////////////////////////////
            // FILTRAGE DES ERREURS //////////////////////////////////////////
            //////////////////////////////////////////////////////////////////
            if (CurrentConfigBoxName.Length == 0)
            {
                MessageBox.Show("Aucune Config n'est sélèctionnée");
                return -1;
            }

            if (CSVFilePathArr.Length == 0)
            {
                MessageBox.Show("Aucun fichier CSV n'est défini");
                return -1;
            }

            int countFiles = CSVFilePathArr.Length;
            int idx = 1;
            int res = 0;
            foreach(string filepath in CSVFilePathArr)
            {
            	res += RunConfig(CurrentConfigBoxName, filepath);
            	
            	// MAJ DE LA PROGRESSBAR
                int percent = (int)((idx * 100 ) / countFiles);
                // Make asynchronous call to main form
                // to set progressbar
                m_form.Invoke(m_form.m_DelegateSetProgressBisValue, new Object[] { percent });
                
                idx++;
            }
            
            m_form.Invoke(m_form.m_DelegateSetProgressBisValue, new Object[] { 100 });
                
            return res;
        }
        
        public int RunConfig(string CurrentConfigBoxName, string CSVFilePath)
        {
            //////////////////////////////////////////////////////////////////
            // FILTRAGE DES ERREURS //////////////////////////////////////////
            //////////////////////////////////////////////////////////////////
            if (CurrentConfigBoxName.Length == 0)
            {
                MessageBox.Show("Aucune Config n'est sélèctionnée");
                return -1;
            }

            if (CSVFilePath.Length == 0)
            {
                MessageBox.Show("Aucun fichier CSV n'est défini");
                return -1;
            }

            if (System.IO.File.Exists(CSVFilePath) == false)
            {
                MessageBox.Show("Le fichier CSV n'existe pas sous ce chemin");
                return -1;
            }

            //////////////////////////////////////////////////////////////////
            // INTIALISATION DU PROCESSUS ////////////////////////////////////
            //////////////////////////////////////////////////////////////////

            // ERROR VARIABLE
            ErrorProcessus = false;
            ErrorProcessString = "";

            // DELETE LOG FILE
        	DeleteLogFile();

            // CLEAR CONSOLE BOX
            // on efface la consolebox dans la mainform dnas la fonction private void button2_Click(object sender, EventArgs e)

            int result = InitLuaFunctions(); // on init les fontions qui vont etre utilisées dans lua

            string appPath = Path.GetDirectoryName(Application.ExecutablePath);

            string luaScriptFile = appPath + "\\scripts\\" + CurrentConfigBoxName;

            try
            {
                luaInterpret.DoFile(luaScriptFile);
            }
            catch (Exception e)
            {
                AddToLogFile(e.ToString());
                
                // Make asynchronous call to main form
                // to inform it that thread finished
                m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                StopProcess(e.ToString());
                return -1;
            }

            try
            {
                luaInterpret.DoString("Init();");
            }
            catch (Exception e)
            {
                AddToLogFile(e.ToString());
                
                // Make asynchronous call to main form
                // to inform it that thread finished
                m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                StopProcess(e.ToString());
                return -1;
            } 

            //////////////////////////////////////////////////////////////////
            // TAILLE DU FICHIER CSV /////////////////////////////////////////
            //////////////////////////////////////////////////////////////////
            long FileSize = 0;
            try
            {
                FileInfo f = new FileInfo(CSVFilePath);
                FileSize = f.Length;
            }
            catch (Exception e)
            {
                AddToLogFile(e.ToString());
                
                // Make asynchronous call to main form
                // to inform it that thread finished
                m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                StopProcess(e.ToString());
                return -1;
            } 

            // get extention
            string ext = Path.GetExtension(CSVFilePath);

            if (ext == "csv")
            {
            	//////////////////////////////////////////////////////////////////
	            // LECTURE DU FICHIER CSV ////////////////////////////////////////
	            //////////////////////////////////////////////////////////////////
	            
	            // On va ouvir le fichier CSV
	            currentLineIndex = 0;
	            string line;
	            long sizeLine = 0;
	
	            System.IO.StreamReader file = null;
	            
	            try
	            {
		            // Read the file and display it line by line.
		            file = new System.IO.StreamReader(CSVFilePath, System.Text.Encoding.Default);
	            }
	            catch (Exception e)
	            {
	                AddToLogFile(e.ToString());
	                
	                // Make asynchronous call to main form
	                // to inform it that thread finished
	                m_form.Invoke(m_form.m_DelegateThreadFinished, null);
	                StopProcess(e.ToString());
	                return -1;
	            } 
	            
	            while ((line = file.ReadLine()) != null)
	            {
	
	                // check if thread is cancelled
	                if (m_EventStop.WaitOne(0, true))
	                {
	                    // clean-up operations may be placed here
	                    // ...
	
	                    file.Close();
	
	                    // inform main thread that this thread stopped
	                    m_EventStopped.Set();
	
	                    return -1;
	                } 
	                
	                if (ErrorProcessus == true)
	                {
	                    MessageBox.Show(ErrorProcessString, "Processus stoppé");
	                    break;
	                }
	
	                // Set CurrentLine
	                result = ParseCSVLineToLuaTable(currentLineName, line);
	
	                // Exec Function for Each Row
	                ExecFonctionForEachLine();
	
	                // Set LastLine
	                result = ParseCSVLineToLuaTable(lastLineName, line);
	
	                // MAJ DE LA PROGRESSBAR
	                sizeLine += (long)line.Length;
	                int percent = (int)((sizeLine * 100 ) / FileSize);
	                // Make asynchronous call to main form
	                // to set progressbar
	                m_form.Invoke(m_form.m_DelegateSetProgressValue, new Object[] { percent });
	
	                currentLineIndex++;
	            }
	
	            // Exec Function for End File
	            ExecFonctionForEndFile();
	            
	            file.Close();
            }
            else if (ext == "xml")
            {
            	//////////////////////////////////////////////////////////////////
	            // LECTURE DU FICHIER XML de type Office SpreadSheet /////////////
	            //////////////////////////////////////////////////////////////////
	            
	            // On va ouvir le fichier XML
	            currentLineIndex = 0;
                string line = "";
	            long sizeLine = 0;
	
	            System.IO.StreamReader file = null;
	            
	            XmlReader xmlReader;
	            
	            try
	            {
	           		xmlReader = XmlReader.Create("http://www.ecb.int/stats/eurofxref/eurofxref-daily.xml");
	            }
	            catch (Exception e)
	            {
	                AddToLogFile(e.ToString());
	                
	                // Make asynchronous call to main form
	                // to inform it that thread finished
	                m_form.Invoke(m_form.m_DelegateThreadFinished, null);
	                StopProcess(e.ToString());
	                return -1;
	            } 
	            
	            while (xmlReader.Read())
	            {
	                // check if thread is cancelled
	                if (m_EventStop.WaitOne(0, true))
	                {
	                    xmlReader.Close();
	
	                    // inform main thread that this thread stopped
	                    m_EventStopped.Set();
	
	                    return -1;
	                } 
	                
	                if (ErrorProcessus == true)
	                {
	                    MessageBox.Show(ErrorProcessString, "Processus stoppé");
	                    break;
	                }
	
	                if(xmlReader.NodeType == XmlNodeType.Element)
	                {
	                	if (xmlReader.Name == "Workbook")
	                	{
	                		if (xmlReader.GetAttribute("xmlns") != "urn:schemas-microsoft-com:office:spreadsheet")
	                			break;// pas le bon format de fichier
	                	}
	                	
	                	if (xmlReader.Name == "Row")
	                	{
	                		
	                	}
	                }
	                
	                // Set CurrentLine
	                result = ParseCSVLineToLuaTable(currentLineName, line);
	
	                // Exec Function for Each Row
	                ExecFonctionForEachLine();
	
	                // Set LastLine
	                result = ParseCSVLineToLuaTable(lastLineName, line);
	
	                // MAJ DE LA PROGRESSBAR
	                sizeLine += (long)line.Length;
	                int percent = (int)((sizeLine * 100 ) / FileSize);
	                // Make asynchronous call to main form
	                // to set progressbar
	                m_form.Invoke(m_form.m_DelegateSetProgressValue, new Object[] { percent });
	
	                currentLineIndex++;
	            }
	
	            // Exec Function for End File
	            ExecFonctionForEndFile();
	            
	            file.Close();
            }

            m_form.Invoke(m_form.m_DelegateSetProgressValue, new Object[] { 100 });
            
            // Make asynchronous call to main form
            // to inform it that thread finished
            m_form.Invoke(m_form.m_DelegateThreadFinished, null);

            return 0;
        }
		
        public void SetProgressCalcul1ValueForLua(int val)
        {
        	if ( val >= 0 && val <= 100 ) 
        		m_form.Invoke(m_form.m_DelegateSetProgressCalcul1Value, new Object[] { val });
        }
        
        public void SetProgressCalcul2ValueForLua(int val)
        {
        	if ( val >= 0 && val <= 100 ) 
        		m_form.Invoke(m_form.m_DelegateSetProgressCalcul2Value, new Object[] { val });
        }
        
        
        // On enregistre les fonctions excel a utilisé dans LUA
        private int InitLuaFunctions()
        {
        	// Info => MainFrom
        	luaInterpret.RegisterFunction("IsAppendOnCurrentExcelFile", m_form, m_form.GetType().GetMethod("IsAppendOnCurrentExcelFile"));
            
            // Info -> this
            luaInterpret.RegisterFunction("SetAuthor", this, this.GetType().GetMethod("SetAuthor"));
            luaInterpret.RegisterFunction("SetSeparator", this, this.GetType().GetMethod("SetSeparator"));
            luaInterpret.RegisterFunction("SetBufferForCurrentLine", this, this.GetType().GetMethod("SetBufferForCurrentLine"));
            luaInterpret.RegisterFunction("SetBufferForLastLine", this, this.GetType().GetMethod("SetBufferForLastLine"));
            luaInterpret.RegisterFunction("SetFunctionForEachLine", this, this.GetType().GetMethod("SetFunctionForEachLine"));
            luaInterpret.RegisterFunction("SetFunctionForEndFile", this, this.GetType().GetMethod("SetFunctionForEndFile"));
            
            // Info Reading CSV File
            luaInterpret.RegisterFunction("GetCurrentRowIndex", this, this.GetType().GetMethod("GetCurrentRowIndex"));
            
            // Progress Bar Control for calcul sup
            luaInterpret.RegisterFunction("SetProgressCalcul1Value", this, this.GetType().GetMethod("SetProgressCalcul1ValueForLua"));
            luaInterpret.RegisterFunction("SetProgressCalcul2Value", this, this.GetType().GetMethod("SetProgressCalcul2ValueForLua"));
                                     
            // Debug Lua
            luaInterpret.RegisterFunction("LogValue", this, this.GetType().GetMethod("LogValue"));
            luaInterpret.RegisterFunction("StopProcess", this, this.GetType().GetMethod("StopProcess"));
            
            // Infos => File
            luaInterpret.RegisterFunction("GetCurrentFileNameWithoutExt", xlsFormater, xlsFormater.GetType().GetMethod("GetCurrentFileNameWithoutExt"));
            
            // Sheet -> xlsFormater
            luaInterpret.RegisterFunction("OpenExcelApp", xlsFormater, xlsFormater.GetType().GetMethod("OpenExcelApp"));
            luaInterpret.RegisterFunction("OpenXlsFile", xlsFormater, xlsFormater.GetType().GetMethod("OpenXlsFile"));
            luaInterpret.RegisterFunction("AddSheet", xlsFormater, xlsFormater.GetType().GetMethod("AddSheet"));
            luaInterpret.RegisterFunction("RenameSheet", xlsFormater, xlsFormater.GetType().GetMethod("RenameSheet"));
            luaInterpret.RegisterFunction("SetActiveSheet", xlsFormater, xlsFormater.GetType().GetMethod("RenameSheet"));

            // Cell -> xlsFormater
            luaInterpret.RegisterFunction("AddCell", xlsFormater, xlsFormater.GetType().GetMethod("AddCell"));
		    luaInterpret.RegisterFunction("GetCell", xlsFormater, xlsFormater.GetType().GetMethod("GetCell"));

            // Insert -> xlsFormater
            luaInterpret.RegisterFunction("InsertColAfter", xlsFormater, xlsFormater.GetType().GetMethod("InsertColAfter"));
            luaInterpret.RegisterFunction("InsertRowAfter", xlsFormater, xlsFormater.GetType().GetMethod("InsertRowAfter"));

            // Size -> xlsFormater
            luaInterpret.RegisterFunction("SetSizeOfCol", xlsFormater, xlsFormater.GetType().GetMethod("SetSizeOfCol"));
            
            // Sort -> xlsFormater
            luaInterpret.RegisterFunction("SortRangeByOneColInOrder", xlsFormater, xlsFormater.GetType().GetMethod("SortRangeByOneColInOrder"));
            luaInterpret.RegisterFunction("SortRangeByTwoColInOrder", xlsFormater, xlsFormater.GetType().GetMethod("SortRangeByTwoColInOrder"));
            luaInterpret.RegisterFunction("SortRangeByThreeColInOrder", xlsFormater, xlsFormater.GetType().GetMethod("SortRangeByThreeColInOrder"));

            // Replace -> xlsFormater
            luaInterpret.RegisterFunction("Replace", xlsFormater, xlsFormater.GetType().GetMethod("Replace"));
            
            // SetRowColor SetColColor SetRangeColor GetColorIndexInNewExcelSheet -> xlsFormater
            luaInterpret.RegisterFunction("SetRowColor", xlsFormater, xlsFormater.GetType().GetMethod("SetRowColor"));
            luaInterpret.RegisterFunction("SetColColor", xlsFormater, xlsFormater.GetType().GetMethod("SetColColor"));
            luaInterpret.RegisterFunction("SetRangeColor", xlsFormater, xlsFormater.GetType().GetMethod("SetRangeColor"));
            luaInterpret.RegisterFunction("GetColorIndexInNewExcelSheet", xlsFormater, xlsFormater.GetType().GetMethod("GetColorIndexInNewExcelSheet"));
            
            // SetRangeBordure SetRowBordure SetColBordure -> xlsFormater
            luaInterpret.RegisterFunction("SetRangeBordure", xlsFormater, xlsFormater.GetType().GetMethod("SetRangeBordure"));
            luaInterpret.RegisterFunction("SetRowBordure", xlsFormater, xlsFormater.GetType().GetMethod("SetRowBordure"));
            luaInterpret.RegisterFunction("SetColBordure", xlsFormater, xlsFormater.GetType().GetMethod("SetColBordure"));
            
            // AutoFitCols AutoFitRows -> xlsFormater
            luaInterpret.RegisterFunction("AutoFitCols", xlsFormater, xlsFormater.GetType().GetMethod("AutoFitCols"));
            luaInterpret.RegisterFunction("AutoFitRows", xlsFormater, xlsFormater.GetType().GetMethod("AutoFitRows"));
            
            // SetRangeAlignement SetColAlignement SetRowAlignement -> xlsFormater
            luaInterpret.RegisterFunction("SetRangeAlignement", xlsFormater, xlsFormater.GetType().GetMethod("SetRangeAlignement"));
            luaInterpret.RegisterFunction("SetColsAlignement", xlsFormater, xlsFormater.GetType().GetMethod("SetColsAlignement"));
            luaInterpret.RegisterFunction("SetRowsAlignement", xlsFormater, xlsFormater.GetType().GetMethod("SetRowsAlignement"));
            
            return 0;
        }

        // Exec Function for Each Row
        private int ExecFonctionForEachLine()
        { 
            if (FonctionForEachLine == null)
            {
                MessageBox.Show("ERROR : ConfigLua.ExecFonctionForEachLine.FonctionForEachLine == 0");
                return -1;
            }
            string str = FonctionForEachLine + "();";

            try
            {
                luaInterpret.DoString(str);
            }
            catch (Exception e)

            {
                string msg = e.ToString() + "/n INTERDICTION DE TOUCHER A ESCEL PENDANT LA GENERATION SOUS PEINE DE PERDRE LA CONNECTION AVEC EXCEL";
                AddToLogFile(e.ToString()); 

                // Make asynchronous call to main form
                // to inform it that thread finished
                m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                StopProcess(e.ToString());
            }
            return 0;
        }

        private int ExecFonctionForEndFile()
        {
        	if (FonctionForEndFile == null)
            {
                MessageBox.Show("ERROR : ConfigLua.ExecFonctionForEndFile.FonctionForEndFile == 0");
                return -1;
            }
            string str = FonctionForEndFile + "();";

            try
            {
                luaInterpret.DoString(str);
            }
            catch (Exception e)

            {
                string msg = e.ToString() + "/n INTERDICTION DE TOUCHER A ESCEL PENDANT LA GENERATION SOUS PEINE DE PERDRE LA CONNECTION AVEC EXCEL";
                AddToLogFile(e.ToString()); 

                // Make asynchronous call to main form
                // to inform it that thread finished
                m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                StopProcess(e.ToString());
            }
            return 0;
        }
        
        private int ParseCSVLineToLuaTable(string tableName, string line)
        {
            if (tableName.Length == 0)
            {
                MessageBox.Show("ERROR : ConfigLua.ParseCSVLineToLuaTable.tableName == 0");
                return -1;
            }

            if (line.Length == 0)
            {
                MessageBox.Show("ERROR : ConfigLua.ParseCSVLineToLuaTable.line == 0");
                return -1;
            }
        
            // parse line

            char separator = CSVSeparator[0];
            string[] words = line.Split(separator);

            // creation de la table lua
                // reset de la table
                try
                {
                    luaInterpret.DoString(tableName + " = {};");
                }
                catch (Exception e)
                {
                    AddToLogFile(e.ToString());
                    
                    // Make asynchronous call to main form
                    // to inform it that thread finished
                    m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                    StopProcess(e.ToString());
                } 
               

                // formatage de la table
                foreach (string s in words)
                {
                    string mot = s;
                    if (mot != "")
                    {
                        if (s[0] == '\"' && mot.Length > 0)
                            mot = mot.Remove(0, 1);
                        if (s[s.Length - 1] == '\"' && mot.Length > 0)
                            mot = mot.Remove(s.Length - 2, 1);
                        // on verifie qu'il n'y a pas de \" dans le mot si on en trouve on le rempalce par un \' 
                        mot = mot.Replace('\"', '\'');
                        mot = mot.Replace("''", "\'");
                    }
                    string str = "table.insert(" + tableName + ",\"" + mot + "\");";
                    
                    try
                    {
                        luaInterpret.DoString(str);
                    }
                    catch (Exception e)
                    {
                        AddToLogFile(e.ToString());
                        
                        // Make asynchronous call to main form
                        // to inform it that thread finished
                        m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                        StopProcess(e.ToString());
                    } 
                    
                }
            
            return 0;
        }

        /*private int ParseXMLRowToLuaTable(string vTableName, XmlReader vReader)
        {
            if (vTableName.Length == 0)
            {
                MessageBox.Show("ERROR : ConfigLua.ParseXMLRowToLuaTable.tableName == 0");
                return -1;
            }

            if (line.Length == 0)
            {
                MessageBox.Show("ERROR : ConfigLua.ParseXMLRowToLuaTable.line == 0");
                return -1;
            }
        
            // parse line

            char separator = CSVSeparator[0];
            string[] words = line.Split(separator);

            // creation de la table lua
                // reset de la table
                try
                {
                    luaInterpret.DoString(vTableName + " = {};");
                }
                catch (Exception e)
                {
                    AddToLogFile(e.ToString());
                    
                    // Make asynchronous call to main form
                    // to inform it that thread finished
                    m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                    StopProcess(e.ToString());
                } 
               

                // formatage de la table
                foreach (string s in words)
                {
                    string mot = s;
                    if (mot != "")
                    {
                        if (s[0] == '\"' && mot.Length > 0)
                            mot = mot.Remove(0, 1);
                        if (s[s.Length - 1] == '\"' && mot.Length > 0)
                            mot = mot.Remove(s.Length - 2, 1);
                        // on verifie qu'il n'y a pas de \" dans le mot si on en trouve on le rempalce par un \' 
                        mot = mot.Replace('\"', '\'');
                        mot = mot.Replace("''", "\'");
                    }
                    string str = "table.insert(" + vTableName + ",\"" + mot + "\");";
                    
                    try
                    {
                        luaInterpret.DoString(str);
                    }
                    catch (Exception e)
                    {
                        AddToLogFile(e.ToString());
                        
                        // Make asynchronous call to main form
                        // to inform it that thread finished
                        m_form.Invoke(m_form.m_DelegateThreadFinished, null);
                        StopProcess(e.ToString());
                    } 
                    
                }
            
            return 0;
        }*/
        
        ///////////////////////////////////////////////////////////////////////////////////
        // FONCTION LUA ///////////////////////////////////////////////////////////////////
        ///////////////////////////////////////////////////////////////////////////////////
        public void SetAuthor(string str)
        {
            if (str.Length == 0)
            {
                MessageBox.Show("ERROR : ConfigLua.SetAuthor.str == 0");
            }
            Author = str;
        }
        
        public void SetSeparator(string str)
        {
            if (str.Length == 0)
            {
                MessageBox.Show("ERROR : Le separateur de champ ne peut pas etre nul");
            }
            if (str.Length > 1)
            {
                MessageBox.Show("ERROR : Le separateur de champ doit etre un simple charatere ex: SetSeparator(\";\"); ");
            }

            CSVSeparator = str;
        }
        
        public void SetBufferForCurrentLine(string str)
        {
            if (str.Length == 0)
            {
                MessageBox.Show("ERROR : ConfigLua.SetBufferForCurrentLine.str == 0");
            }
            currentLineName = str;
        }
        
        public void SetBufferForLastLine(string str)
        {
            if (str.Length == 0)
           	{
                MessageBox.Show("ERROR : ConfigLua.SetBufferForLastLine.str == 0");
            }
            lastLineName = str;
        }

        public void SetFunctionForEachLine(string str)
        {
            if (str.Length == 0)
            {
                MessageBox.Show("ERROR : ConfigLua.SetFunctionForEachLine.str == 0");
            }
            FonctionForEachLine = str;
        }
        
        public void SetFunctionForEndFile(string str)
        {
            if (str.Length == 0)
            {
                MessageBox.Show("ERROR : ConfigLua.SetFunctionForEndFile.str == 0");
            }
            FonctionForEndFile = str;
        }
        
        public int GetCurrentRowIndex()
        {
        	return currentLineIndex;
        }
        
        // Log dans un fichier
        public void LogValue(string value)
        {
            AddToLogFile(value);
        }

        public void StopProcess(string erreur)
        {
            ErrorProcessus = true;
            ErrorProcessString = erreur;
        }
	}
}
