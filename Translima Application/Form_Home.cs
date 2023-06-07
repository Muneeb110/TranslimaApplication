using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Serialization;
using IniParser.Model;
using IniParser.Parser;
using MetroFramework;
using MetroFramework.Forms;
using Logging_Framework;

namespace Translima_Application
{
    public partial class Form_Home : MetroForm
    {
        Logger logger;
        private bool Proceed = true;

        private Timer TimerProcess;

        private static Form_Home Home = null;

        private delegate void EnableDelegate(string[] Values);

        public Form_Home()
        {
            InitializeComponent();
            Home = this;
        }

        private void Form_Home_Load(object sender, EventArgs e)
        {
            CheckForIllegalCrossThreadCalls = false;

            if (!Properties.Settings.Default.INIFile.Equals("None"))
            {
                Txt_INIFile.Text = Properties.Settings.Default.INIFile;
                ReadINIFile(Properties.Settings.Default.INIFile);
            }
        }

        private void Btn_INIFile_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    InitialDirectory = @"C:\",
                    Title = "Browse INI Files",
                    CheckFileExists = true,
                    CheckPathExists = true,
                    DefaultExt = "ini",
                    Filter = "ini files (*.ini)|*.ini",
                    FilterIndex = 2,
                    RestoreDirectory = true,
                    ShowReadOnly = true
                };

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    Properties.Settings.Default.INIFile = openFileDialog.FileName;
                    Properties.Settings.Default.Save();

                    Txt_INIFile.Text = openFileDialog.FileName;
                    ReadINIFile(openFileDialog.FileName);
                }
            }
            catch(Exception ex)
            {
               
                MetroMessageBox.Show(this, ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReadINIFile(string INIFile)
        {
            try
            {
                IniDataParser parser = new IniDataParser();
                parser.Configuration.CommentString = "#";

                IniData parsedData;

                using (FileStream fileStream = File.Open(INIFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                {
                    using (StreamReader streamReader = new StreamReader(fileStream, System.Text.Encoding.UTF8))
                    {
                        parsedData = parser.Parse(streamReader.ReadToEnd());
                    }
                }

                Txt_PollingFolder.Text = parsedData["Settings"]["PollingFolder"];
                Txt_ProcessedFolder.Text = parsedData["Settings"]["ProcessedFolder"];
                Txt_Timer.Text = parsedData["Settings"]["Timer"];
                Txt_ConnectionString.Text = parsedData["Settings"]["ConnectionString"];
                Txt_LogPath.Text = parsedData["Settings"]["LogPath"];
                Txt_LogLevel.Text = parsedData["Settings"]["LogLevel"];
                Txt_FileName.Text = parsedData["Settings"]["FileName"];
                txt_pdfFolderPath.Text = parsedData["Settings"]["PDFFolder"];
                txtPDFSize.Text = parsedData["Settings"]["PDFSizeLimit"];
                switch (Txt_LogLevel.Text)
                {
                    case "Info":
                        logger = Logger.GetLogger(Txt_LogPath.Text, Txt_FileName.Text, LogLevels.Info);
                        break;
                    case "Debug":
                        logger = Logger.GetLogger(Txt_LogPath.Text, Txt_FileName.Text, LogLevels.Debug);
                        break;
                    default:
                        break;
                }
            }
            catch(Exception ex)
            {
                
                MetroMessageBox.Show(this, ex.Message, "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ReadXML()
        {
            if(!System.IO.Directory.Exists(Txt_PollingFolder.Text.Trim()))
            {
                Proceed = true;

                this.Invoke((MethodInvoker)delegate
                {
                    Btn_Start.Text = "Start";
                    Btn_Start.BackColor = Color.Lime;

                    MetroMessageBox.Show(this, "Polling folder does not exist!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
            else if (!System.IO.Directory.Exists(Txt_ProcessedFolder.Text.Trim()))
            {
                Proceed = true;

                this.Invoke((MethodInvoker)delegate
                {
                    Btn_Start.Text = "Start";
                    Btn_Start.BackColor = Color.Lime;

                    MetroMessageBox.Show(this, "Proccessed folder does not exist!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
            else
            {
                Proceed = false;

                Database database = new Database(Txt_ConnectionString.Text, logger);

                string errorFolder = @"D:\HARDDISK\INTFILES\AEB2TRANSLIMA\Polling\Error\";
                string processedFolder = Txt_ProcessedFolder.Text.Trim() + DateTime.Now.ToString("yyyy-MM-dd") + @"\";

                if (!System.IO.Directory.Exists(processedFolder))
                {
                    System.IO.Directory.CreateDirectory(processedFolder);
                }

                if (database.ConnectionStringValidity.Equals("Success"))
                //if(true)
                {
                    string[] valueStart = { "true" };
                    UpdateStartProcess(valueStart);

                    string[] files = Directory.GetFiles(Txt_PollingFolder.Text.Trim(), "*.xml", SearchOption.TopDirectoryOnly);

                    foreach (string file in files)
                    {
                        logger.Log(LogLevels.Info, "File found: " + file);
                        BrokerFileDTO brokerFileDTO = null;

                        XmlSerializer serializer = new XmlSerializer(typeof(BrokerFileDTO));

                        StreamReader reader = new StreamReader(file);
                        brokerFileDTO = (BrokerFileDTO)serializer.Deserialize(reader);
                        reader.Close();

                        database.Start(brokerFileDTO, this, txt_pdfFolderPath.Text, txtPDFSize.Text);
                        logger.Log(LogLevels.Info, "File processed: " + file);
                        try
                        {
                            File.Move(file, file.Replace(Txt_PollingFolder.Text.Trim(), processedFolder));
                        }
                        catch
                        {
                            if (!System.IO.Directory.Exists(errorFolder))
                            {
                                System.IO.Directory.CreateDirectory(errorFolder);
                            }
                            File.Move(file, file.Replace(Txt_PollingFolder.Text.Trim(), errorFolder));

                        }
                    }

                    database.CloseCon();

                    string[] valueEnd = { "false" };
                    UpdateStartProcess(valueEnd);
                }
                else if (database.ConnectionStringValidity.Equals("Error"))
                {
                    Proceed = true;

                    this.Invoke((MethodInvoker)delegate
                    {
                        Btn_Start.Text = "Start";
                        Btn_Start.BackColor = Color.Lime;

                        MetroMessageBox.Show(this, "Invalid Connection String!\nCheck your SQL connection again!", "Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            }
        }

        private async void TimerProcess_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("Timer End: " + DateTime.Now);
            logger.Log(LogLevels.Info, "Timer End: " + DateTime.Now);
            TimerProcess.Stop();

            await Task.Run(() =>
            {
                ReadXML();
            });
        }

        private async void Btn_Start_Click(object sender, EventArgs e)
        {
            switch (Txt_LogLevel.Text)
            {
                case "Info":
                    logger = Logger.GetLogger(Txt_LogPath.Text, Txt_FileName.Text, LogLevels.Info);
                    break;
                case "Debug":
                    logger = Logger.GetLogger(Txt_LogPath.Text, Txt_FileName.Text, LogLevels.Debug);
                    break;
                default:
                    break;
            }
            if (Proceed == false)
            {
                Proceed = true;

                Btn_Start.Text = "Start";
                Btn_Start.BackColor = Color.Lime;
                logger.Log(LogLevels.Info, "Stopping Timer.");
                TimerProcess.Stop();
            }
            else
            {
                await Task.Run(() =>
                {
                    logger.Log(LogLevels.Info, "Going to Read XML.");
                    ReadXML();
                });
            }
        }

        public static void UpdateHistory(string[] Values)
        {
            if (Home != null)
                Home.GridHistory(Values);
        }

        public static void UpdateCommercial(string[] Values)
        {
            if (Home != null)
                Home.GridCommercial(Values);
        }

        public static void UpdateCommercialExtra(string[] Values)
        {
            if (Home != null)
                Home.GridCommercialExtra(Values);
        }

        public static void UpdateCustomsOffices(string[] Values)
        {
            if (Home != null)
                Home.GridCustomsOffices(Values);
        }

        public static void UpdateInvoices(string[] Values)
        {
            if (Home != null)
                Home.GridInvoices(Values);
        }

        public static void UpdateItems(string[] Values)
        {
            if (Home != null)
                Home.GridItems(Values);
        }

        public static void UpdateItemsExtra(string[] Values)
        {
            if (Home != null)
                Home.GridItemsExtra(Values);
        }

        public static void UpdateParties(string[] Values)
        {
            if (Home != null)
                Home.GridParties(Values);
        }

        private void GridHistory(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(GridHistory), new object[] { Values });
                return;
            }

            Grid_History.Rows.Add(Values[0], Values[1], Values[2], Values[3], Values[4]);
        }

        private void GridCommercial(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(GridCommercial), new object[] { Values });
                return;
            }

            Grid_Commercial.Rows.Add(Values[0], Values[1], Values[2], Values[3], Values[4], Values[5], Values[6], Values[7], Values[8], Values[9], Values[10], Values[11], Values[12], Values[13], Values[14], Values[15], Values[16], Values[17], Values[18], Values[19], Values[20], Values[21]);
        }

        private void GridCommercialExtra(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(GridCommercialExtra), new object[] { Values });
                return;
            }

            Grid_CommercialExtra.Rows.Add(Values[0], Values[1], Values[2], Values[3]);
        }

        private void GridCustomsOffices(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(GridCustomsOffices), new object[] { Values });
                return;
            }

            Grid_CustomsOffices.Rows.Add(Values[0], Values[1], Values[2], Values[3]);
        }

        private void GridInvoices(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(GridInvoices), new object[] { Values });
                return;
            }

            Grid_Invoices.Rows.Add(Values[0], Values[1], Values[2], Values[3], Values[4]);
        }

        private void GridItems(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(GridItems), new object[] { Values });
                return;
            }

            Grid_Items.Rows.Add(Values[0], Values[1], Values[2], Values[3], Values[4], Values[5], Values[6], Values[7], Values[8], Values[9], Values[10], Values[11], Values[12], Values[13], Values[14], Values[15]);
        }

        private void GridItemsExtra(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(GridItemsExtra), new object[] { Values });
                return;
            }

            Grid_ItemsExtra.Rows.Add(Values[0], Values[1], Values[2]);
        }

        private void GridParties(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(GridParties), new object[] { Values });
                return;
            }

            Grid_Parties.Rows.Add(Values[0], Values[1], Values[2], Values[3], Values[4], Values[5], Values[6], Values[7], Values[8], Values[9], Values[10]);
        }

        public static void UpdateDashboard(string[] Values)
        {
            if (Home != null)
                Home.Dashboard(Values);
        }

        private void Dashboard(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(Dashboard), new object[] { Values });
                return;
            }

            Tile_TotalDeclarations.Text = "TOTAL DECLARATIONS\n" + Values[0];
            Tile_TotalDeclarationsExtra.Text = "TOTAL DECLARATIONS EXTRA\n" + Values[1];
            Tile_TotalCustomOffices.Text = "TOTAL CUSTOM OFFICES\n" + Values[2];
            Tile_TotalInvoices.Text = "TOTAL INVOICES\n" + Values[3];
            Tile_TotalItems.Text = "TOTAL ITEMS\n" + Values[4];
            Tile_TotalItemsExtra.Text = "TOTAL ITEMS EXTRA\n" + Values[5];
            Tile_TotalParties.Text = "TOTAL PARTIES\n" + Values[6];
        }

        private static void UpdateStartProcess(string[] Values)
        {
            if (Home != null)
                Home.StartProcess(Values);
        }

        private void StartProcess(string[] Values)
        {
            if (InvokeRequired)
            {
                this.Invoke(new EnableDelegate(StartProcess), new object[] { Values });
                return;
            }

            if (Values[0].Equals("true"))
            {
                Btn_Start.Enabled = false;
                PS_SavingData.Visible = true;
            }
            else
            {
                Btn_Start.Text = "Stop";
                Btn_Start.BackColor = Color.Red;

                Btn_Start.Enabled = true;
                PS_SavingData.Visible = false;

                TimerProcess = new Timer
                {
                    Interval = (Convert.ToInt32(Txt_Timer.Text.Trim()) * 1000)
                };

                TimerProcess.Tick += new EventHandler(TimerProcess_Tick);
                TimerProcess.Start();
                logger.Log(LogLevels.Info, "Timer Start: " + DateTime.Now);
                Console.WriteLine("Timer Start: " + DateTime.Now);
            }
        }
        
    }
}
