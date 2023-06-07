using Logging_Framework;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data;
using Spire.Pdf;
using System.Drawing.Imaging;
using System.Drawing;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace Translima_Application
{
    class Database
    {
        private SqlConnection Con = null;
        private BrokerFileDTO Data = null;
        Logger logger;
        public string ConnectionStringValidity = "";

        public Database(string Connection_String, Logging_Framework.Logger logger)
        {
            this.logger = logger;
            try
            {
                Con = new SqlConnection(Connection_String);
                Con.Open();
                this.logger.Log(LogLevels.Info, "Database connection Open");
                ConnectionStringValidity = "Success";
            }
            catch (Exception ex)
            {
                Con = null;
                Console.WriteLine("An error occured while closing the connection!\nReason: " + ex.ToString());
                this.logger.Log(LogLevels.Error, "An error occured while closing the connection:" + ex.ToString());
                ConnectionStringValidity = "Error";

            }
        }

        public void CloseCon()
        {
            try
            {
                Con.Close();
                Con.Dispose();
                logger.Log(LogLevels.Info, "Database Connection Closed.");
            }
            catch (Exception ex)
            {
                logger.Log(LogLevels.Error, "An error occured while closing the connection: " + ex.ToString());
                Console.WriteLine("An error occured while closing the connection!\nReason: " + ex.ToString());
            }
        }

        public void Start(BrokerFileDTO BrokerFileDTOData, Form_Home Home, string pdfFolderPath, string pdfSize)
        {
            Data = BrokerFileDTOData;
            logger.Log(LogLevels.Info, "Going to InsertCommercial");

            //ConvertPDf(pdfFolderPath, pdfSize, "AY3CGC00B31000368935");

            bool isCommercialDataInserted = InsertCommercial();
            string localReference = "";
            if (isCommercialDataInserted)
            {

                logger.Log(LogLevels.Info, "Going to InsertCommercialExtra");
                InsertCommercialExtra(ref localReference);
                logger.Log(LogLevels.Info, "Going to InsertCustomsOffices");
                InsertCommercialAdditionalFields();
                logger.Log(LogLevels.Info, "Going to InsertAdditionalFields");
                InsertCustomsOffices();
                logger.Log(LogLevels.Info, "Going to InsertInvoices");
                InsertInvoices();
                logger.Log(LogLevels.Info, "Going to InsertItems");
                InsertItems();
                logger.Log(LogLevels.Info, "Going to InsertParties");
                InsertParties();
                logger.Log(LogLevels.Info, "Going to InsertCosts");
                InsertCosts();
                logger.Log(LogLevels.Info, "Going to call SP:[SP_CHECKS]");
                CallStoredProcedure(localReference);
                logger.Log(LogLevels.Info, "Going to create the txt file for Peach");
                CreatePeachTxt();
                ConvertPDf(pdfFolderPath, pdfSize, localReference);
            }
            else
            {
                logger.Log(LogLevels.Info, "Commercial insert failed, not going to insert other tables data as well.");
            }
            logger.Log(LogLevels.Info, "Going to SetDashboard");
            SetDashboard();



        }

        private void CallStoredProcedure(string localReference)
        {
            try
            {



                using (var command = new SqlCommand("SP_CHECKS", Con))
                {
                    command.CommandType = CommandType.StoredProcedure;
                    command.Parameters.Add("@Status", SqlDbType.VarChar).Value = "";
                    command.Parameters.Add("@localReference", SqlDbType.VarChar).Value = localReference;
                    command.ExecuteNonQuery();
                }


            }
            catch (Exception ex)
            {
                logger.Log(LogLevels.Error, "An error occured while calling stored procedure SP_CHECKS.\nReason: " + ex.ToString());
                Console.WriteLine("An error occured while calling stored procedure SP_CHECKS.\nReason:" + ex.ToString());
                InsertHistory(Data.Declaration.LocalReference, "SP_CHECKS", "Error", "An error occured while calling stored procedure SP_CHECKS.Reason:" + ex.ToString());
            }
        }

        private void SplitAndSaveInterval(string pdfFilePath, string outputPath, int startPage, int interval, string pdfFileName)
        {
            using (PdfReader reader = new PdfReader(pdfFilePath))
            {
                Document document = new Document();
                PdfCopy copy = new PdfCopy(document, new FileStream(outputPath + "\\" + pdfFileName + ".pdf", FileMode.Create));
                document.Open();

                for (int pagenumber = startPage; pagenumber < (startPage + interval); pagenumber++)
                {
                    if (reader.NumberOfPages >= pagenumber)
                    {
                        copy.AddPage(copy.GetImportedPage(reader, pagenumber));
                    }
                    else
                    {
                        break;
                    }

                }

                document.Close();
            }
        }

        //public void MergeTiff(string[] sourceFiles, string outputFile)
        //{
        //    string[] sa = sourceFiles;
        //    //get the codec for tiff files
        //    ImageCodecInfo info = null;
        //    foreach (ImageCodecInfo ice in ImageCodecInfo.GetImageEncoders())
        //        if (ice.MimeType == "image/tiff")
        //            info = ice;

        //    //use the save encoder
        //    System.Drawing.Imaging.Encoder enc = System.Drawing.Imaging.Encoder.SaveFlag;

        //    EncoderParameters ep = new EncoderParameters(1);
        //    ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);

        //    Bitmap pages = null;

        //    int frame = 0;

        //    foreach (string s in sa)
        //    {
        //        if (frame == 0)
        //        {
        //            MemoryStream ms = new MemoryStream(File.ReadAllBytes(s));
        //            pages = (Bitmap)Image.FromStream(ms);

        //            //save the first frame
        //            pages.Save(outputFile, info, ep);
        //        }
        //        else
        //        {
        //            //save the intermediate frames
        //            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);

        //            try
        //            {
        //                MemoryStream mss = new MemoryStream(File.ReadAllBytes(s));
        //                Bitmap bm = (Bitmap)Image.FromStream(mss);
        //                pages.SaveAdd(bm, ep);
        //            }
        //            catch (Exception e)
        //            {
        //                // LogError(e, s);
        //            }
        //        }

        //        if (frame == sa.Length - 1)
        //        {
        //            //flush and close.
        //            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.Flush);
        //            pages.SaveAdd(ep);

        //        }

        //        frame++;
        //    }

        //}

        private void ConvertPDf(string pdfFolderPath, string pdfSize, string localReference)
        {
            logger.Log(LogLevels.Debug, "Going to convert pdf files for local reference:" + localReference);
            string directoryPath = pdfFolderPath + "\\" + localReference + "\\PDF";

            if (Directory.Exists(directoryPath))
            {
                string backupDirectoryPath = pdfFolderPath + "\\" + localReference + "\\Backup";
                Directory.CreateDirectory(backupDirectoryPath);
                string[] files = Directory.GetFiles(directoryPath);
                foreach (var item in files)
                {
                    string extension = Path.GetFileName(item);
                    string fileName = Path.GetFileNameWithoutExtension(item);

                    if (extension.Contains("pdf"))
                    {
                        FileInfo fileInfo = new FileInfo(item);
                        var fileLengthInKB = fileInfo.Length / 1024;
                        if (fileLengthInKB > int.Parse(pdfSize))
                        {
                            PdfReader reader = new PdfReader(item);
                            int interval = 2;
                            int pageNameSuffix = 0;
                            string pdfFileName = fileName + "-";

                            if (reader.NumberOfPages > 1)
                            {
                                for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber += interval)
                                {
                                    pageNameSuffix++;
                                    string newPdfFileName = string.Format(pdfFileName + "{0}", pageNameSuffix);
                                    SplitAndSaveInterval(item, directoryPath, pageNumber, interval, newPdfFileName);
                                }


                                //string[] fileNames = new string[pdfDocument.Pages.Count];
                                //for (int i = 0; i < pdfDocument.Pages.Count; i++)
                                //{

                                //    bmp = pdfDocument.SaveAsImage(i, 96, 96);
                                //    bmp.Save(directoryPath + "\\" + string.Format(fileName + "{0}.tiff", i), ImageFormat.Tiff);
                                //    Rectangle rect = new Rectangle(0, 0, bmp.Width, bmp.Height);
                                //    //Original image  
                                //    Bitmap _img = new Bitmap(bmp.Width, bmp.Height);
                                //    // for cropinf image  
                                //    Graphics g = Graphics.FromImage(_img);
                                //    // create graphics  
                                //    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                //    g.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
                                //    g.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                                //    //set image attributes  
                                //    g.DrawImage(bmp, 0, -50, rect, GraphicsUnit.Pixel);
                                //    _img.Save(directoryPath + "\\" + string.Format(fileName + "{0}.tiff", i), ImageFormat.Tiff);
                                //    fileNames[i] = directoryPath + "\\" + string.Format(fileName + "{0}.tiff",i);
                                //}
                                ////MergeTiff(fileNames, directoryPath + "\\" + string.Format(fileName + ".tiff"));
                                //foreach (var item1 in fileNames)
                                //{
                                //    File.Delete(item1);
                                //}
                                // JoinTiffImages(images, directoryPath + "\\" + string.Format(fileName + ".tiff"), EncoderValue.CompressionLZW);
                            }
                            else
                            {
                                reader.Close();
                                Spire.Pdf.PdfDocument pdfDocument = new Spire.Pdf.PdfDocument(item);
                                System.Drawing.Image bmp;

                                bmp = pdfDocument.SaveAsImage(0);
                                bmp.Save(directoryPath + "\\" + fileName + ".jpeg", ImageFormat.Jpeg);

                            }
                            reader.Close();
                            File.Move(item, backupDirectoryPath + "\\" + extension);
                            InsertHistory(localReference, "Commercial", "Success", "PDF conversion performed for filename: " + item);
                        }
                        else
                        {
                            logger.Log(LogLevels.Debug, "Not converting file:" + extension);
                        }
                    }
                }
            }
            else
                logger.Log(LogLevels.Debug, "PDF folder not found.");

        }

        private ImageCodecInfo GetEncoderInfo(string mimeType)

        {

            ImageCodecInfo[] encoders = ImageCodecInfo.GetImageEncoders();

            for (int j = 0; j < encoders.Length; j++)

            {

                if (encoders[j].MimeType == mimeType)

                    return encoders[j];

            }
            logger.Log(LogLevels.Error, mimeType + " mime type not found in ImageCodecInfo");
            return null;
            //throw new Exception(mimeType + " mime type not found in ImageCodecInfo");

        }


        //public void JoinTiffImages(Image[] images, string outFile, EncoderValue compressEncoder)

        //{

        //    //use the save encoder

        //    System.Drawing.Imaging.Encoder enc = System.Drawing.Imaging.Encoder.SaveFlag;

        //    EncoderParameters ep = new EncoderParameters(2);

        //    ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.MultiFrame);

        //    ep.Param[1] = new EncoderParameter(System.Drawing.Imaging.Encoder.Compression, (long)compressEncoder);

        //    Image pages = images[0];

        //    int frame = 0;

        //    ImageCodecInfo info = GetEncoderInfo("image/tiff");

        //    foreach (Image img in images)

        //    {

        //        if (frame == 0)

        //        {

        //            pages = img;

        //            //save the first frame

        //            pages.Save(outFile, info, ep);

        //        }

        //        else

        //        {

        //            //save the intermediate frames

        //            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.FrameDimensionPage);

        //            pages.SaveAdd(img, ep);

        //        }

        //        if (frame == images.Length - 1)

        //        {

        //            //flush and close.

        //            ep.Param[0] = new EncoderParameter(enc, (long)EncoderValue.Flush);

        //            pages.SaveAdd(ep);

        //        }

        //        frame++;

        //    }

        //}


        private void InsertCosts()
        {
            for (int c = 0; c < Data.Declaration.costs.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO costs " +
                            "(amount," +
                             "type," +
                             "currencyIso," +
                             "localReference)" +
                            "VALUES " +
                            "(@amount, " +
                            "@type, " +
                            "@currencyIso," +
                            "@localReference);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);

                    Cmd.Parameters.AddWithValue("@amount", Data.Declaration.costs[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].costsamount.valueamount) ? (object)"No Value" : (object)Data.Declaration.costs[c].costsamount.valueamount);
                    Cmd.Parameters.AddWithValue("@currencyIso", Data.Declaration.costs[c].costsamount.currencyIsoamount == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].costsamount.currencyIsoamount) ? (object)"No Value" : (object)Data.Declaration.costs[c].costsamount.currencyIsoamount);
                    Cmd.Parameters.AddWithValue("@type", Data.Declaration.costs[c].type == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].type) ? (object)"No Value" : (object)Data.Declaration.costs[c].type);
                    Cmd.Parameters.AddWithValue("@localReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);

                    logger.Log(LogLevels.Debug, "INSERT INTO costs " +
                           "(amount," +
                            "type," +
                            "currencyIso," +
                            "localReference)" +
                           "VALUES " +
                           "(" + (Data.Declaration.costs[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].costsamount.valueamount) ? (object)"No Value" : (object)Data.Declaration.costs[c].costsamount.valueamount) + ", " +
                           (Data.Declaration.costs[c].type == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].type) ? (object)"No Value" : (object)Data.Declaration.costs[c].type) + ", " +
                           (Data.Declaration.costs[c].costsamount.currencyIsoamount == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].costsamount.currencyIsoamount) ? (object)"No Value" : (object)Data.Declaration.costs[c].costsamount.currencyIsoamount) + "," +
                           (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference) + ");");


                    Cmd.ExecuteNonQuery();

                    Cmd.Dispose();

                    string[] values = {
                        Data.Declaration.costs[c].costsamount.valueamount == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].costsamount.valueamount) ? "No Value" : Data.Declaration.costs[c].costsamount.valueamount,
                        Data.Declaration.costs[c].costsamount.currencyIsoamount == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].costsamount.currencyIsoamount) ? "No Value" : Data.Declaration.costs[c].costsamount.currencyIsoamount,
                        Data.Declaration.costs[c].type == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.costs[c].type) ? "No Value" : Data.Declaration.costs[c].type

                    };
                    InsertHistory(Data.Declaration.LocalReference, "Costs", "Success", "'Costs' table inserted successfully!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occured while inserting to 'invoices' table!\nReason: " + ex.ToString());
                    logger.Log(LogLevels.Error, "An error occured while inserting to 'invoices' table!Reason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "Costs", "Error", "An error occured while inserting to 'Costs' table!\nReason: " + ex.ToString());
                }
            }
        }

        private void InsertHistory(string Key, string Table, string Status, string Message)
        {
            try
            {
                string Query = "INSERT INTO history " +
                        "(dateTime, " +
                        "[key], " +
                        "[table], " +
                        "status, " +
                        "message) " +
                        "VALUES " +
                        "(@dateTime, " +
                        "@key, " +
                        "@table, " +
                        "@status, " +
                        "@message);";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                Cmd.Parameters.AddWithValue("@dateTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
                Cmd.Parameters.AddWithValue("@key", Key);
                Cmd.Parameters.AddWithValue("@table", Table);
                Cmd.Parameters.AddWithValue("@status", Status);
                Cmd.Parameters.AddWithValue("@message", Message);


                logger.Log(LogLevels.Debug, "INSERT INTO history " +
                        "(dateTime, " +
                        "[key], " +
                        "[table], " +
                        "status, " +
                        "message) " +
                        "VALUES " +
                        "(" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + ", " +
                        Key + ", " +
                        Table + ", " +
                        Status + ", " +
                        Message + ");");

                Cmd.ExecuteNonQuery();

                Cmd.Dispose();


                string[] values = {
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    Key,
                    Table,
                    Status,
                    Message
                };

                Form_Home.UpdateHistory(values);
            }
            catch (Exception ex)
            {
                logger.Log(LogLevels.Error, "An error occured while inserting to 'history' table!Reason: " + ex.ToString());
                Console.WriteLine("An error occured while inserting to 'history' table!\nReason: " + ex.ToString());
            }
        }

        private bool InsertCommercial()
        {
            bool isCommercialDataInserted;
            try
            {
                string Query = "INSERT INTO commercial " +
                        "(clientCode, " +
                        "createDate, " +
                        "commercialReference, " +
                        "destinationCountryCode, " +
                        "dispatchCountryCode, " +
                        "incotermCode, " +
                        "incotermPlace, " +
                        "personInChargeName, " +
                        "totalGrossWeight, " +
                        "totalInvoicePrice, " +
                        "totalNetWeight, " +
                        "transportMeansIdentity, " +
                        "transportMeansType, " +
                        "transportMeansModeType, " +
                        "nationality, " +
                        "declaringCountryCode, " +
                        "process, " +
                        "version, " +
                        "status, " +
                        "processCode, " +
                        "ContactDetailsName, " +
                        "ContactDetailsTelephone," +
                        "instructions," +
                        "personInChargePhone," +
                        "personInChargeEmail," +
                        "localReference) " +
                        "VALUES " +
                        "(@clientCode, " +
                        "@createDate, " +
                        "@commercialReference, " +
                        "@destinationCountryCode, " +
                        "@dispatchCountryCode, " +
                        "@incotermCode, " +
                        "@incotermPlace, " +
                        "@personInChargeName, " +
                        "@totalGrossWeight, " +
                        "@totalInvoicePrice, " +
                        "@totalNetWeight, " +
                        "@transportMeansIdentity, " +
                        "@transportMeansType, " +
                        "@transportMeansModeType, " +
                        "@nationality, " +
                        "@declaringCountryCode, " +
                        "@process, " +
                        "@version, " +
                        "@status, " +
                        "@processCode, " +
                        "@ContactDetailsName, " +
                        "@ContactDetailsTelephone," +
                        "@instructions," +
                        "@personInChargePhone," +
                        "@personInChargeEmail," +
                        "@localReference);";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                Cmd.Parameters.AddWithValue("@clientCode", Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.ClientCode) ? (object)"No Value" : (object)Data.ClientCode);
                Cmd.Parameters.AddWithValue("@createDate", Data.CreateDate == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.CreateDate.DateInTimezone + " " + Data.CreateDate.Timezone) ? (object)"No Value" : (object)Data.CreateDate.DateInTimezone + " " + Data.CreateDate.Timezone);
                Cmd.Parameters.AddWithValue("@commercialReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? (object)"No Value" : (object)Data.Declaration.CommercialReference);
                Cmd.Parameters.AddWithValue("@destinationCountryCode", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.DestinationCountryCode) ? (object)"No Value" : (object)Data.Declaration.DestinationCountryCode);
                Cmd.Parameters.AddWithValue("@dispatchCountryCode", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.DispatchCountryCode) ? (object)"No Value" : (object)Data.Declaration.DispatchCountryCode);
                Cmd.Parameters.AddWithValue("@incotermCode", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.IncotermCode) ? (object)"No Value" : (object)Data.Declaration.IncotermCode);
                Cmd.Parameters.AddWithValue("@incotermPlace", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.IncotermPlace) ? (object)"No Value" : (object)Data.Declaration.IncotermPlace);
                Cmd.Parameters.AddWithValue("@localReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);
                Cmd.Parameters.AddWithValue("@personInChargeName", Data.Declaration.PersonInCharge == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.PersonInCharge.Name) ? (object)"No Value" : (object)Data.Declaration.PersonInCharge.Name);
                Cmd.Parameters.AddWithValue("@totalGrossWeight", Data.Declaration.TotalGrossWeight == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalGrossWeight.Value + " " + Data.Declaration.TotalGrossWeight.Unit) ? (object)"No Value" : (object)Data.Declaration.TotalGrossWeight.Value + " " + Data.Declaration.TotalGrossWeight.Unit);
                Cmd.Parameters.AddWithValue("@totalInvoicePrice", Data.Declaration.TotalInvoicePrice == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalInvoicePrice.Value + " " + Data.Declaration.TotalInvoicePrice.CurrencyIso) ? (object)"No Value" : (object)Data.Declaration.TotalInvoicePrice.Value + " " + Data.Declaration.TotalInvoicePrice.CurrencyIso);
                Cmd.Parameters.AddWithValue("@totalNetWeight", Data.Declaration.TotalNetWeight == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalNetWeight.Value + " " + Data.Declaration.TotalNetWeight.Unit) ? (object)"No Value" : (object)Data.Declaration.TotalNetWeight.Value + " " + Data.Declaration.TotalNetWeight.Unit);
                Cmd.Parameters.AddWithValue("@transportMeansIdentity", Data.Declaration.TransportMeans == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.Identity) ? (object)"No Value" : (object)Data.Declaration.TransportMeans.Identity);
                Cmd.Parameters.AddWithValue("@transportMeansType", Data.Declaration.TransportMeans == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.MeansType) ? (object)"No Value" : (object)Data.Declaration.TransportMeans.MeansType);
                Cmd.Parameters.AddWithValue("@transportMeansModeType", Data.Declaration.TransportMeans == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.ModeType) ? (object)"No Value" : (object)Data.Declaration.TransportMeans.ModeType);
                Cmd.Parameters.AddWithValue("@nationality", Data.Declaration.TransportMeans == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.Nationality) ? (object)"No Value" : (object)Data.Declaration.TransportMeans.Nationality);
                Cmd.Parameters.AddWithValue("@declaringCountryCode", Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.DeclaringCountryCode) ? (object)"No Value" : (object)Data.DeclaringCountryCode);
                Cmd.Parameters.AddWithValue("@process", Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Process) ? (object)"No Value" : (object)Data.Process);
                Cmd.Parameters.AddWithValue("@version", Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Version) ? (object)"No Value" : (object)Data.Version);
                Cmd.Parameters.AddWithValue("@status", "NEW");
                Cmd.Parameters.AddWithValue("@processCode", "DEFRA");
                Cmd.Parameters.AddWithValue("@ContactDetailsName", "ChannelPorts Ltd");
                Cmd.Parameters.AddWithValue("@ContactDetailsTelephone", "004412345567889");
                Cmd.Parameters.AddWithValue("@instructions", Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Instructions) ? (object)"No Value" : (object)Data.Declaration.Instructions);
                Cmd.Parameters.AddWithValue("@personInChargePhone", Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.PersonInCharge.Phone) ? (object)"No Value" : (object)Data.Declaration.PersonInCharge.Phone);
                Cmd.Parameters.AddWithValue("@personInChargeEmail", Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.PersonInCharge.Email) ? (object)"No Value" : (object)Data.Declaration.PersonInCharge.Email);


                logger.Log(LogLevels.Debug, "INSERT INTO commercial " +
                        "(clientCode, " +
                        "createDate, " +
                        "commercialReference, " +
                        "destinationCountryCode, " +
                        "dispatchCountryCode, " +
                        "incotermCode, " +
                        "incotermPlace, " +
                        "personInChargeName, " +
                        "totalGrossWeight, " +
                        "totalInvoicePrice, " +
                        "totalNetWeight, " +
                        "transportMeansIdentity, " +
                        "transportMeansType, " +
                        "transportMeansModeType, " +
                        "nationality, " +
                        "declaringCountryCode, " +
                        "process, " +
                        "version, " +
                        "status, " +
                        "processCode, " +
                        "ContactDetailsName, " +
                        "ContactDetailsTelephone," +
                        "localReference) " +
                        "VALUES " +
                        "(" + (Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.ClientCode) ? (object)"No Value" : (object)Data.ClientCode) + ", " +
                        (Data.CreateDate == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.CreateDate.DateInTimezone + " " + Data.CreateDate.Timezone) ? (object)"No Value" : (object)Data.CreateDate.DateInTimezone + " " + Data.CreateDate.Timezone) + ", " +
                        (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? (object)"No Value" : (object)Data.Declaration.CommercialReference) + ", " +
                        (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.DestinationCountryCode) ? (object)"No Value" : (object)Data.Declaration.DestinationCountryCode) + ", " +
                        (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.DispatchCountryCode) ? (object)"No Value" : (object)Data.Declaration.DispatchCountryCode) + ", " +
                        (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.IncotermCode) ? (object)"No Value" : (object)Data.Declaration.IncotermCode) + ", " +
                        (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.IncotermPlace) ? (object)"No Value" : (object)Data.Declaration.IncotermPlace) + ", " +
                        (Data.Declaration.PersonInCharge == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.PersonInCharge.Name) ? (object)"No Value" : (object)Data.Declaration.PersonInCharge.Name) + ", " +
                        (Data.Declaration.TotalGrossWeight == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalGrossWeight.Value + " " + Data.Declaration.TotalGrossWeight.Unit) ? (object)"No Value" : (object)Data.Declaration.TotalGrossWeight.Value + " " + Data.Declaration.TotalGrossWeight.Unit) + ", " +
                        (Data.Declaration.TotalInvoicePrice == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalInvoicePrice.Value + " " + Data.Declaration.TotalInvoicePrice.CurrencyIso) ? (object)"No Value" : (object)Data.Declaration.TotalInvoicePrice.Value + " " + Data.Declaration.TotalInvoicePrice.CurrencyIso) + ", " +
                        (Data.Declaration.TotalNetWeight == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalNetWeight.Value + " " + Data.Declaration.TotalNetWeight.Unit) ? (object)"No Value" : (object)Data.Declaration.TotalNetWeight.Value + " " + Data.Declaration.TotalNetWeight.Unit) + ", " +
                        (Data.Declaration.TransportMeans == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.Identity) ? (object)"No Value" : (object)Data.Declaration.TransportMeans.Identity) + ", " +
                        (Data.Declaration.TransportMeans == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.MeansType) ? (object)"No Value" : (object)Data.Declaration.TransportMeans.MeansType) + ", " +
                        (Data.Declaration.TransportMeans == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.ModeType) ? (object)"No Value" : (object)Data.Declaration.TransportMeans.ModeType) + ", " +
                        (Data.Declaration.TransportMeans == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.Nationality) ? (object)"No Value" : (object)Data.Declaration.TransportMeans.Nationality) + ", " +
                        (Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.DeclaringCountryCode) ? (object)"No Value" : (object)Data.DeclaringCountryCode) + ", " +
                        (Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Process) ? (object)"No Value" : (object)Data.Process) + ", " +
                        (Data == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Version) ? (object)"No Value" : (object)Data.Version) + ", " +
                        "NEW, " +
                        "DEFRA, " +
                        "ChannelPorts Ltd, " +
                        "004412345567889," +
                        (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference) + ")");

                Cmd.ExecuteNonQuery();

                Cmd.Dispose();

                string[] values = {
                    Data == null ? "No Value" : String.IsNullOrWhiteSpace(Data.ClientCode) ? "No Value" : Data.ClientCode,
                    Data.CreateDate == null ? "No Value" : String.IsNullOrWhiteSpace(Data.CreateDate.DateInTimezone + " " + Data.CreateDate.Timezone) ? "No Value" : Data.CreateDate.DateInTimezone + " " + Data.CreateDate.Timezone,
                    Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? "No Value" : Data.Declaration.CommercialReference,
                    Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.DestinationCountryCode) ? "No Value" : Data.Declaration.DestinationCountryCode,
                    Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.DispatchCountryCode) ? "No Value" : Data.Declaration.DispatchCountryCode,
                    //Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.GoodsDescription) ? "No Value" : Data.Declaration.GoodsDescription,
                    Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.IncotermCode) ? "No Value" : Data.Declaration.IncotermCode,
                    Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.IncotermPlace) ? "No Value" : Data.Declaration.IncotermPlace,
                    Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? "No Value" : Data.Declaration.LocalReference,
                    Data.Declaration.PersonInCharge == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.PersonInCharge.Name) ? "No Value" : Data.Declaration.PersonInCharge.Name,
                    Data.Declaration.TotalGrossWeight == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalGrossWeight.Value + " " + Data.Declaration.TotalGrossWeight.Unit) ? "No Value" : Data.Declaration.TotalGrossWeight.Value + " " + Data.Declaration.TotalGrossWeight.Unit,
                    Data.Declaration.TotalInvoicePrice == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalInvoicePrice.Value + " " + Data.Declaration.TotalInvoicePrice.CurrencyIso) ? "No Value" : Data.Declaration.TotalInvoicePrice.Value + " " + Data.Declaration.TotalInvoicePrice.CurrencyIso,
                    Data.Declaration.TotalNetWeight == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TotalNetWeight.Value + " " + Data.Declaration.TotalNetWeight.Unit) ? "No Value" : Data.Declaration.TotalNetWeight.Value + " " + Data.Declaration.TotalNetWeight.Unit,
                    Data.Declaration.TransportMeans == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.Identity) ? "No Value" : Data.Declaration.TransportMeans.Identity,
                    Data.Declaration.TransportMeans == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.MeansType) ? "No Value" : Data.Declaration.TransportMeans.MeansType,
                    Data.Declaration.TransportMeans == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.ModeType) ? "No Value" : Data.Declaration.TransportMeans.ModeType,
                    Data.Declaration.TransportMeans == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.TransportMeans.Nationality) ? "No Value" : Data.Declaration.TransportMeans.Nationality,
                    Data == null ? "No Value" : String.IsNullOrWhiteSpace(Data.DeclaringCountryCode) ? "No Value" : Data.DeclaringCountryCode,
                    Data == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Process) ? "No Value" : Data.Process,
                    Data == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Version) ? "No Value" : Data.Version,
                    "NEW",
                    "ChannelPorts Ltd",
                    "004412345567889",
                    Data.Declaration.Instructions == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Instructions) ? "No Value" : Data.Declaration.Instructions,
                    Data.Declaration.PersonInCharge.Phone == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.PersonInCharge.Phone) ? "No Value" : Data.Declaration.PersonInCharge.Phone,
                    Data.Declaration.PersonInCharge.Email == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.PersonInCharge.Email) ? "No Value" : Data.Declaration.PersonInCharge.Email
                };

                Form_Home.UpdateCommercial(values);

                InsertHistory(Data.Declaration.LocalReference, "commercial", "Success", "'commercial' table inserted successfully!");
                isCommercialDataInserted = true;

            }
            catch (Exception ex)
            {
                logger.Log(LogLevels.Error, "An error occured while inserting to 'commercial' table!Reason: " + ex.ToString());
                Console.WriteLine("An error occured while inserting to 'commercial' table!\nReason: " + ex.ToString());
                InsertHistory(Data.Declaration.LocalReference, "commercial", "Error", "An error occured while inserting to 'commercial' table!\nReason: " + ex.ToString());
                isCommercialDataInserted = false;
            }
            return isCommercialDataInserted;
        }

        private void InsertCommercialExtra(ref string localReference)
        {
            for (int c = 0; c < Data.Declaration.ExtraFields.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO commercial_extra " +
                            "([key], " +
                            "value, " +
                            "commercialReference, " +
                            "localReference) " +
                            "VALUES " +
                            "(@key, " +
                            "@value, " +
                            "@commercialReference, " +
                            "@localReference);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);

                    Cmd.Parameters.AddWithValue("@key", Data.Declaration.ExtraFields[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.ExtraFields[c].Key) ? (object)"No Value" : (object)Data.Declaration.ExtraFields[c].Key);
                    Cmd.Parameters.AddWithValue("@value", Data.Declaration.ExtraFields[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.ExtraFields[c].Value) ? (object)"No Value" : (object)Data.Declaration.ExtraFields[c].Value);
                    Cmd.Parameters.AddWithValue("@commercialReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? (object)"No Value" : (object)Data.Declaration.CommercialReference);
                    Cmd.Parameters.AddWithValue("@localReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);

                    logger.Log(LogLevels.Debug, "INSERT INTO commercial_extra " +
                            "([key], " +
                            "value, " +
                            "commercialReference, " +
                            "localReference) " +
                            "VALUES " +
                            "(" + (Data.Declaration.ExtraFields[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.ExtraFields[c].Key) ? (object)"No Value" : (object)Data.Declaration.ExtraFields[c].Key) + ", " +
                            (Data.Declaration.ExtraFields[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.ExtraFields[c].Value) ? (object)"No Value" : (object)Data.Declaration.ExtraFields[c].Value) + ", " +
                            (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? (object)"No Value" : (object)Data.Declaration.CommercialReference) + ", " +
                            (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference) + ");");
                    Cmd.ExecuteNonQuery();


                    Cmd.Dispose();

                    string[] values = {
                        Data.Declaration.ExtraFields[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.ExtraFields[c].Key) ? "No Value" : Data.Declaration.ExtraFields[c].Key,
                        Data.Declaration.ExtraFields[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.ExtraFields[c].Value) ? "No Value" : Data.Declaration.ExtraFields[c].Value,
                        Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? "No Value" : Data.Declaration.CommercialReference,
                        Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? "No Value" : Data.Declaration.LocalReference
                    };

                    Form_Home.UpdateCommercialExtra(values);

                    InsertHistory(Data.Declaration.LocalReference, "commercial_extra", "Success", "'commercial_extra' table inserted successfully!");
                    localReference = Data.Declaration.LocalReference;
                }
                catch (Exception ex)
                {
                    logger.Log(LogLevels.Error, "An error occured while inserting to 'commercial_extra' table!Reason: " + ex.ToString());
                    Console.WriteLine("An error occured while inserting to 'commercial_extra' table!\nReason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "commercial_extra", "Error", "An error occured while inserting to 'commercial_extra' table!\nReason: " + ex.ToString());
                }
            }
        }

        private void InsertCommercialAdditionalFields()
        {
            for (int c = 0; c < Data.Declaration.AdditionalReferences.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO commercial_additionalref " +
                            "(itemNumber, " +
                            "reference, " +
                            "referenceType, " +
                            "issueDate, " +
                            "localReference) " +
                            "VALUES " +
                            "(@itemNumber, " +
                            "@referenceType, " +
                            "@reference, " +
                            "@issueDate, " +
                            "@localReference);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);

                    Cmd.Parameters.AddWithValue("@itemNumber", Data.Declaration.AdditionalReferences[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].itemNumber) ? (object)"No Value" : (object)Data.Declaration.AdditionalReferences[c].itemNumber);
                    Cmd.Parameters.AddWithValue("@reference", Data.Declaration.AdditionalReferences[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].referenceType) ? (object)"No Value" : (object)Data.Declaration.AdditionalReferences[c].referenceType);
                    Cmd.Parameters.AddWithValue("@referenceType", Data.Declaration.AdditionalReferences[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].reference) ? (object)"No Value" : (object)Data.Declaration.AdditionalReferences[c].reference);
                    Cmd.Parameters.AddWithValue("@issueDate", Data.Declaration.AdditionalReferences[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].IssueDate.dateInTimezone) ? (object)"No Value" : (object)Data.Declaration.AdditionalReferences[c].IssueDate.dateInTimezone);
                    Cmd.Parameters.AddWithValue("@localReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);

                    logger.Log(LogLevels.Debug, "INSERT INTO commercial_additionalref " +
                            "(itemNumber, " +
                            "reference, " +
                            "referenceType, " +
                            "issueDate, " +
                            "localReference) " +
                            "VALUES " +
                            "(" + (Data.Declaration.AdditionalReferences[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].itemNumber) ? "No Value" : Data.Declaration.AdditionalReferences[c].itemNumber) + ", " +
                            (Data.Declaration.AdditionalReferences[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].referenceType) ? "No Value" : Data.Declaration.AdditionalReferences[c].referenceType) + ", " +
                            (Data.Declaration.AdditionalReferences[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].reference) ? "No Value" : Data.Declaration.AdditionalReferences[c].reference) + ", " +
                            (Data.Declaration.AdditionalReferences[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].IssueDate.dateInTimezone) ? "No Value" : Data.Declaration.AdditionalReferences[c].IssueDate.dateInTimezone) + ", " +
                            (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference) + ");");
                    Cmd.ExecuteNonQuery();


                    Cmd.Dispose();

                    string[] values = {
                        Data.Declaration.AdditionalReferences[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].itemNumber) ? "No Value" : Data.Declaration.AdditionalReferences[c].itemNumber,
                        Data.Declaration.AdditionalReferences[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].referenceType) ? "No Value" : Data.Declaration.AdditionalReferences[c].referenceType,
                        Data.Declaration.AdditionalReferences[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].reference) ? "No Value" : Data.Declaration.AdditionalReferences[c].reference,
                        Data.Declaration.AdditionalReferences[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.AdditionalReferences[c].IssueDate.dateInTimezone) ? "No Value" : Data.Declaration.AdditionalReferences[c].IssueDate.dateInTimezone,
                        Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? "No Value" : Data.Declaration.LocalReference
                    };

                    InsertHistory(Data.Declaration.LocalReference, "commercial_additionalref", "Success", "'commercial_additionalref' table inserted successfully!");
                }
                catch (Exception ex)
                {
                    logger.Log(LogLevels.Error, "An error occured while inserting to 'commercial_additionalref' table!Reason: " + ex.ToString());
                    Console.WriteLine("An error occured while inserting to 'commercial_additionalref' table!\nReason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "commercial_additionalref", "Error", "An error occured while inserting to 'commercial_additionalref' table!\nReason: " + ex.ToString());
                }
            }
        }

        private void InsertCustomsOffices()
        {
            for (int c = 0; c < Data.Declaration.CustomsOffices.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO customs_offices " +
                            "(identCCode, " +
                            "officeType, " +
                            "localReference) " +
                            "VALUES " +
                            "(@identCCode, " +
                            "@officeType, " +
                            "@localReference);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);

                    Cmd.Parameters.AddWithValue("@identCCode", Data.Declaration.CustomsOffices[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CustomsOffices[c].IdentCcode) ? (object)"No Value" : (object)Data.Declaration.CustomsOffices[c].IdentCcode);
                    Cmd.Parameters.AddWithValue("@officeType", Data.Declaration.CustomsOffices[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CustomsOffices[c].OfficeType) ? (object)"No Value" : (object)Data.Declaration.CustomsOffices[c].OfficeType);
                    Cmd.Parameters.AddWithValue("@localReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);


                    logger.Log(LogLevels.Debug, "INSERT INTO customs_offices " +
                            "(identCCode, " +
                            "officeType, " +
                            "localReference) " +
                            "VALUES " +
                            "(" + (Data.Declaration.CustomsOffices[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CustomsOffices[c].IdentCcode) ? (object)"No Value" : (object)Data.Declaration.CustomsOffices[c].IdentCcode) + ", " +
                            (Data.Declaration.CustomsOffices[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CustomsOffices[c].OfficeType) ? (object)"No Value" : (object)Data.Declaration.CustomsOffices[c].OfficeType) + ", " +
                            (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference) + ");");

                    Cmd.ExecuteNonQuery();

                    Cmd.Dispose();

                    string[] values = {
                        Data.Declaration.CustomsOffices[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CustomsOffices[c].IdentCcode) ? "No Value" : Data.Declaration.CustomsOffices[c].IdentCcode,
                        Data.Declaration.CustomsOffices[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CustomsOffices[c].OfficeType) ? "No Value" : Data.Declaration.CustomsOffices[c].OfficeType,
                        Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? "No Value" : Data.Declaration.LocalReference
                    };

                    Form_Home.UpdateCustomsOffices(values);

                    InsertHistory(Data.Declaration.LocalReference, "customs_offices", "Success", "'customs_offices' table inserted successfully!");
                }
                catch (Exception ex)
                {
                    logger.Log(LogLevels.Error, "An error occured while inserting to 'customs_offices' table!Reason: " + ex.ToString());
                    Console.WriteLine("An error occured while inserting to 'customs_offices' table!\nReason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "customs_offices", "Error", "An error occured while inserting to 'customs_offices' table!\nReason: " + ex.ToString());
                }
            }
        }

        private void InsertInvoices()
        {
            for (int c = 0; c < Data.Declaration.Invoices.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO invoices " +
                            "(invoiceDate, " +
                            "invoiceNumber, " +
                            "invoicePrice, " +
                            "localReference) " +
                            "VALUES " +
                            "(@invoiceDate, " +
                            "@invoiceNumber, " +
                            "@invoicePrice, " +
                            "@localReference);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);



                    Cmd.Parameters.AddWithValue("@invoiceDate", Data.Declaration.Invoices[c].InvoiceDate == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoiceDate.DateInTimezone + " " + Data.Declaration.Invoices[c].InvoiceDate.Timezone) ? (object)"No Value" : (object)Data.Declaration.Invoices[c].InvoiceDate.DateInTimezone + " " + Data.Declaration.Invoices[c].InvoiceDate.Timezone);
                    Cmd.Parameters.AddWithValue("@invoiceNumber", Data.Declaration.Invoices[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoiceNumber) ? (object)"No Value" : (object)Data.Declaration.Invoices[c].InvoiceNumber);
                    Cmd.Parameters.AddWithValue("@invoicePrice", Data.Declaration.Invoices[c].InvoicePrice == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoicePrice.Value + " " + Data.Declaration.Invoices[c].InvoicePrice.CurrencyIso) ? (object)"No Value" : (object)Data.Declaration.Invoices[c].InvoicePrice.Value + " " + Data.Declaration.Invoices[c].InvoicePrice.CurrencyIso);
                    Cmd.Parameters.AddWithValue("@localReference", Data.Declaration.LocalReference == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);

                    logger.Log(LogLevels.Debug, "INSERT INTO invoices " +
                            "(invoiceDate, " +
                            "invoiceNumber, " +
                            "invoicePrice, " +
                            "localReference) " +
                            "VALUES " +
                            "(" + (Data.Declaration.Invoices[c].InvoiceDate == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoiceDate.DateInTimezone + " " + Data.Declaration.Invoices[c].InvoiceDate.Timezone) ? (object)"No Value" : (object)Data.Declaration.Invoices[c].InvoiceDate.DateInTimezone + " " + Data.Declaration.Invoices[c].InvoiceDate.Timezone) + ", " +
                            (Data.Declaration.Invoices[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoiceNumber) ? (object)"No Value" : (object)Data.Declaration.Invoices[c].InvoiceNumber) + ", " +
                            (Data.Declaration.Invoices[c].InvoicePrice == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoicePrice.Value + " " + Data.Declaration.Invoices[c].InvoicePrice.CurrencyIso) ? (object)"No Value" : (object)Data.Declaration.Invoices[c].InvoicePrice.Value + " " + Data.Declaration.Invoices[c].InvoicePrice.CurrencyIso) + ", " +
                            (Data.Declaration.LocalReference == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference) + ");");
                    Cmd.ExecuteNonQuery();

                    Cmd.Dispose();

                    string[] values = {
                        Data.Declaration.Invoices[c].InvoiceDate == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoiceDate.DateInTimezone + " " + Data.Declaration.Invoices[c].InvoiceDate.Timezone) ? "No Value" : Data.Declaration.Invoices[c].InvoiceDate.DateInTimezone + " " + Data.Declaration.Invoices[c].InvoiceDate.Timezone,
                        Data.Declaration.Invoices[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoiceNumber) ? "No Value" : Data.Declaration.Invoices[c].InvoiceNumber,
                        Data.Declaration.Invoices[c].InvoicePrice == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Invoices[c].InvoicePrice.Value + " " + Data.Declaration.Invoices[c].InvoicePrice.CurrencyIso) ? "No Value" : Data.Declaration.Invoices[c].InvoicePrice.Value + " " + Data.Declaration.Invoices[c].InvoicePrice.CurrencyIso,
                        Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? "No Value" : Data.Declaration.CommercialReference,
                        Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? "No Value" : Data.Declaration.LocalReference
                    };

                    Form_Home.UpdateInvoices(values);

                    InsertHistory(Data.Declaration.LocalReference, "invoices", "Success", "'invoices' table inserted successfully!");
                }
                catch (Exception ex)
                {
                    logger.Log(LogLevels.Error, "An error occured while inserting to 'invoices' table!Reason: " + ex.ToString());
                    Console.WriteLine("An error occured while inserting to 'invoices' table!\nReason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "invoices", "Error", "An error occured while inserting to 'invoices' table!\nReason: " + ex.ToString());
                }
            }
        }

        //private void SortItems()
        //{
        //    try
        //    {
        //        List<int> indexOfY = new List<int>();
        //        indexOfY.Clear();

        //        List<int> indexOfN = new List<int>();
        //        indexOfN.Clear();

        //        List<int> indexOfS = new List<int>();
        //        indexOfS.Clear();

        //        List<DUCR> sortedItems = new List<DUCR>();
        //        sortedItems.Clear();

        //        for (int c = 0; c < Data.Declaration.Items.Count; c++)
        //        {
        //            for (int i = 0; i < Data.Declaration.Items[c].ExtraFields.Count; i++)
        //            {
        //                if (Data.Declaration.Items[c].ExtraFields[i].Key == "PEACH_PLANTPASSPORT" && Data.Declaration.Items[c].ExtraFields[i].Value == "Y")
        //                {
        //                    indexOfY.Add(c);
        //                }
        //                else
        //                {
        //                    indexOfN.Add(c);
        //                }
        //            }
        //        }

        //        for (int c = 0; c < indexOfY.Count; c++)
        //        {
        //            indexOfS.Add(indexOfY[c]);
        //        }

        //        for (int c = 0; c < indexOfN.Count; c++)
        //        {
        //            indexOfS.Add(indexOfN[c]);
        //        }

        //        for (int c = 0; c < indexOfS.Count; c++)
        //        {
        //            sortedItems.Add(new DUCR
        //            {
        //                packageIdClientSystem = Data.Declaration.Items[indexOfS[c]].Packages.packageIdClientSystem == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].Packages.packageIdClientSystem,
        //                GoodsDescription = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].GoodsDescription,
        //                GrossWeight = Data.Declaration.Items[indexOfS[c]].GrossWeight == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].GrossWeight.Value + " " + Data.Declaration.Items[indexOfS[c]].GrossWeight.Unit,
        //                InvoiceNumbers = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].InvoiceNumbers,
        //                ItemNumber = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].ItemNumber,
        //                MaterialNumber = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].MaterialNumber,
        //                NetPrice = Data.Declaration.Items[indexOfS[c]].NetPrice == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].NetPrice.Value + " " + Data.Declaration.Items[indexOfS[c]].NetPrice.CurrencyIso,
        //                NetWeight = Data.Declaration.Items[indexOfS[c]].NetWeight == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].NetWeight.Value + " " + Data.Declaration.Items[indexOfS[c]].NetWeight.Unit,
        //                OriginCountryCode = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].OriginCountryCode,
        //                PackageMarks = Data.Declaration.Items[indexOfS[c]].Packages == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].Packages.Marks,
        //                PackageNumber = Data.Declaration.Items[indexOfS[c]].Packages == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].Packages.Number,
        //                PackageType = Data.Declaration.Items[indexOfS[c]].Packages == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].Packages.PackageType,
        //                packedItemQuantity = Data.Declaration.Items[indexOfS[c]].Packages.packedItemQuantity == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].Packages.packedItemQuantity,
        //                PreferentialCountryCode = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].PreferentialCountryCode,
        //                SequenceNumber = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].SequenceNumber,
        //                TariffNumber = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].TariffNumber,
        //                StatisticalValue = Data.Declaration.Items[indexOfS[c]].StatisticalValue == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].StatisticalValue.Value + " " + Data.Declaration.Items[indexOfS[c]].StatisticalValue.CurrencyIso,
        //                //DUCRNumber = GenerateDUCRNumber(c, Data.Declaration == null ? "No Value" : Data.Declaration.LocalReference),
        //                DUCRNumber = Data.Declaration.Items[indexOfS[c]] == null ? "No Value" : Data.Declaration.Items[indexOfS[c]].TariffNumber,
        //                CommercialReference = Data.Declaration == null ? "No Value" : Data.Declaration.CommercialReference,
        //                LocalReference = Data.Declaration == null ? "No Value" : Data.Declaration.LocalReference,
        //                Index = indexOfS[c]
        //            });
        //        }

        //        InsertItems(sortedItems);
        //    }
        //    catch(Exception ex)
        //    {
        //        Console.WriteLine("An error occured while sorting 'items'!\nReason: " + ex.ToString());
        //    }
        //}

        private void InsertItems()
        {
            for (int c = 0; c < Data.Declaration.Items.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO items " +
                            "(goodsDescription, " +
                            "grossWeight, " +
                            "invoiceNumbers, " +
                            "itemNumber, " +
                            "materialNumber, " +
                            "netPrice, " +
                            "netPriceCurrency, " +
                            "netWeight, " +
                            "originCountryCode, " +
                            "preferentialCountryCode, " +
                            "sequenceNumber, " +
                            "tariffNumber, " +
                            "statisticalValue, " +
                            "quantity," +
                            "quantityUOM," +
                            "itemStatus," +
                            "localReference) " +
                            "output INSERTED.itemsID VALUES " +
                            "(@goodsDescription, " +
                            "@grossWeight, " +
                            "@invoiceNumbers, " +
                            "@itemNumber, " +
                            "@materialNumber, " +
                            "@netPrice, " +
                            "@netPriceCurrency, " +
                            "@netWeight, " +
                            "@originCountryCode, " +
                            "@preferentialCountryCode, " +
                            "@sequenceNumber, " +
                            "@tariffNumber, " +
                            "@statisticalValue, " +
                            "@quantity, " +
                            "@quantityUOM," +
                            "@itemStatus," +
                            "@localReference);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);

                    Cmd.Parameters.AddWithValue("@goodsDescription", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].GoodsDescription) ? (object)"No Value" : (object)Data.Declaration.Items[c].GoodsDescription);
                    Cmd.Parameters.AddWithValue("@grossWeight", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].GrossWeight.Value) ? (object)"No Value" : (object)Data.Declaration.Items[c].GrossWeight.Value);
                    Cmd.Parameters.AddWithValue("@invoiceNumbers", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].InvoiceNumbers) ? (object)"No Value" : (object)Data.Declaration.Items[c].InvoiceNumbers);
                    Cmd.Parameters.AddWithValue("@itemNumber", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].ItemNumber) ? (object)"No Value" : (object)Data.Declaration.Items[c].ItemNumber);
                    Cmd.Parameters.AddWithValue("@materialNumber", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].MaterialNumber) ? (object)"No Value" : (object)Data.Declaration.Items[c].MaterialNumber);
                    Cmd.Parameters.AddWithValue("@netPrice", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].NetPrice.Value) ? (object)"No Value" : (object)Data.Declaration.Items[c].NetPrice.Value);
                    Cmd.Parameters.AddWithValue("@netPriceCurrency", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].NetPrice.CurrencyIso) ? (object)"No Value" : (object)Data.Declaration.Items[c].NetPrice.CurrencyIso);
                    Cmd.Parameters.AddWithValue("@netWeight", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].NetWeight.Value) ? (object)"No Value" : (object)Data.Declaration.Items[c].NetWeight.Value);
                    Cmd.Parameters.AddWithValue("@originCountryCode", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].OriginCountryCode) ? (object)"No Value" : (object)Data.Declaration.Items[c].OriginCountryCode);
                    Cmd.Parameters.AddWithValue("@preferentialCountryCode", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].PreferentialCountryCode) ? (object)"No Value" : (object)Data.Declaration.Items[c].PreferentialCountryCode);
                    Cmd.Parameters.AddWithValue("@sequenceNumber", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].SequenceNumber) ? (object)"No Value" : (object)Data.Declaration.Items[c].SequenceNumber);
                    Cmd.Parameters.AddWithValue("@tariffNumber", Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].TariffNumber) ? (object)"No Value" : (object)Data.Declaration.Items[c].TariffNumber);
                    Cmd.Parameters.AddWithValue("@statisticalValue", Data.Declaration.Items[c].StatisticalValue == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].StatisticalValue.Value + " " + Data.Declaration.Items[c].StatisticalValue.CurrencyIso) ? (object)"No Value" : (object)Data.Declaration.Items[c].StatisticalValue.Value + " " + Data.Declaration.Items[c].StatisticalValue.CurrencyIso);
                    Cmd.Parameters.AddWithValue("@quantity", Data.Declaration.Items[c].Quantity.Unit == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].Quantity.Value) ? (object)"No Value" : (object)Data.Declaration.Items[c].Quantity.Value);
                    Cmd.Parameters.AddWithValue("@quantityUOM", Data.Declaration.Items[c].Quantity.Value == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].Quantity.Unit) ? (object)"No Value" : (object)Data.Declaration.Items[c].Quantity.Unit);
                    Cmd.Parameters.AddWithValue("@itemStatus", "NEW");
                    Cmd.Parameters.AddWithValue("@localReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);

                    logger.Log(LogLevels.Debug, "INSERT INTO items " +
                            "(goodsDescription, " +
                            "grossWeight, " +
                            "invoiceNumbers, " +
                            "itemNumber, " +
                            "materialNumber, " +
                            "netPrice, " +
                            "netWeight, " +
                            "originCountryCode, " +
                            "preferentialCountryCode, " +
                            "sequenceNumber, " +
                            "tariffNumber, " +
                            "statisticalValue, " +
                            "itemStatus, " +
                            "localReference) " +
                            "output INSERTED.itemsID VALUES " +
                            "(" + (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].GoodsDescription) ? (object)"No Value" : (object)Data.Declaration.Items[c].GoodsDescription) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].GrossWeight.Value) ? (object)"No Value" : (object)Data.Declaration.Items[c].GrossWeight.Value) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].InvoiceNumbers) ? (object)"No Value" : (object)Data.Declaration.Items[c].InvoiceNumbers) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].ItemNumber) ? (object)"No Value" : (object)Data.Declaration.Items[c].ItemNumber) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].MaterialNumber) ? (object)"No Value" : (object)Data.Declaration.Items[c].MaterialNumber) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].NetPrice.Value) ? (object)"No Value" : (object)Data.Declaration.Items[c].NetPrice.Value) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].NetWeight.Value) ? (object)"No Value" : (object)Data.Declaration.Items[c].NetWeight.Value) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].OriginCountryCode) ? (object)"No Value" : (object)Data.Declaration.Items[c].OriginCountryCode) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].PreferentialCountryCode) ? (object)"No Value" : (object)Data.Declaration.Items[c].PreferentialCountryCode) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].SequenceNumber) ? (object)"No Value" : (object)Data.Declaration.Items[c].SequenceNumber) + ", " +
                            (Data.Declaration.Items[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].TariffNumber) ? (object)"No Value" : (object)Data.Declaration.Items[c].TariffNumber) + ", " +
                            (Data.Declaration.Items[c].StatisticalValue == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].StatisticalValue.Value + " " + Data.Declaration.Items[c].StatisticalValue.CurrencyIso) ? (object)"No Value" : (object)Data.Declaration.Items[c].StatisticalValue.Value + " " + Data.Declaration.Items[c].StatisticalValue.CurrencyIso) + ", " +
                            "NEW, " +
                            (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference) + ");");


                    int itemsID = (int)Cmd.ExecuteScalar();

                    Cmd.Dispose();

                    string[] values = {
                        itemsID.ToString(),
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].GoodsDescription) ? "No Value" : Data.Declaration.Items[c].GoodsDescription,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].GrossWeight.Value) ? "No Value" : Data.Declaration.Items[c].GrossWeight.Value,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].InvoiceNumbers) ? "No Value" : Data.Declaration.Items[c].InvoiceNumbers,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].ItemNumber) ? "No Value" : Data.Declaration.Items[c].ItemNumber,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].MaterialNumber) ? "No Value" : Data.Declaration.Items[c].MaterialNumber,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].NetPrice.Value) ? "No Value" : Data.Declaration.Items[c].NetPrice.Value,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].NetPrice.CurrencyIso) ? "No Value" : Data.Declaration.Items[c].NetPrice.CurrencyIso,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].NetWeight.Value) ? "No Value" : Data.Declaration.Items[c].NetWeight.Value,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].OriginCountryCode) ? "No Value" : Data.Declaration.Items[c].OriginCountryCode,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].PreferentialCountryCode) ? "No Value" : Data.Declaration.Items[c].PreferentialCountryCode,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].SequenceNumber) ? "No Value" : Data.Declaration.Items[c].SequenceNumber,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].TariffNumber) ? "No Value" : Data.Declaration.Items[c].TariffNumber,
                         Data.Declaration.Items[c].StatisticalValue == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].StatisticalValue.Value + " " + Data.Declaration.Items[c].StatisticalValue.CurrencyIso) ? "No Value" : Data.Declaration.Items[c].StatisticalValue.Value + " " + Data.Declaration.Items[c].StatisticalValue.CurrencyIso,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].Quantity.Value) ? "No Value" : Data.Declaration.Items[c].Quantity.Value,
                         Data.Declaration.Items[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[c].Quantity.Unit) ? "No Value" : Data.Declaration.Items[c].Quantity.Unit,
                         Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? "No Value" : Data.Declaration.LocalReference,
                    };

                    Form_Home.UpdateItems(values);

                    InsertHistory(Data.Declaration.LocalReference, "items", "Success", "'items' table inserted successfully!");

                    InsertItemsExtra(c, itemsID);
                    InsertItemsPackages(c, itemsID);
                }
                catch (Exception ex)
                {
                    logger.Log(LogLevels.Error, "An error occured while inserting to 'items' table!Reason: " + ex.ToString());
                    Console.WriteLine("An error occured while inserting to 'items' table!\nReason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "items", "Error", "An error occured while inserting to 'items' table!\nReason: " + ex.ToString());
                }
            }
        }

        private void InsertItemsExtra(int Index, int ItemsID)
        {
            for (int c = 0; c < Data.Declaration.Items[Index].ExtraFields.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO items_extra " +
                            "([key], " +
                            "value, " +
                            "itemsID) " +
                            "VALUES " +
                            "(@key, " +
                            "@value, " +
                            "@itemsID);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);

                    Cmd.Parameters.AddWithValue("@key", Data.Declaration.Items[Index].ExtraFields[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].ExtraFields[c].Key) ? (object)"No Value" : (object)Data.Declaration.Items[Index].ExtraFields[c].Key);
                    Cmd.Parameters.AddWithValue("@value", Data.Declaration.Items[Index].ExtraFields[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].ExtraFields[c].Value) ? (object)"No Value" : (object)Data.Declaration.Items[Index].ExtraFields[c].Value);
                    Cmd.Parameters.AddWithValue("@itemsID", ItemsID);
                    logger.Log(LogLevels.Debug, "INSERT INTO items_extra " +
                            "([key], " +
                            "value, " +
                            "itemsID) " +
                            "VALUES " +
                            "(@" + (Data.Declaration.Items[Index].ExtraFields[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].ExtraFields[c].Key) ? (object)"No Value" : (object)Data.Declaration.Items[Index].ExtraFields[c].Key) + ", " +
                            (Data.Declaration.Items[Index].ExtraFields[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].ExtraFields[c].Value) ? (object)"No Value" : (object)Data.Declaration.Items[Index].ExtraFields[c].Value) + ", " +
                            ItemsID + ");");
                    Cmd.ExecuteNonQuery();

                    Cmd.Dispose();

                    string[] values = {
                        Data.Declaration.Items[Index].ExtraFields[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].ExtraFields[c].Key) ? "No Value" : Data.Declaration.Items[Index].ExtraFields[c].Key,
                        Data.Declaration.Items[Index].ExtraFields[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].ExtraFields[c].Value) ? "No Value" : Data.Declaration.Items[Index].ExtraFields[c].Value,
                        ItemsID.ToString()
                    };

                    Form_Home.UpdateItemsExtra(values);

                    InsertHistory(Data.Declaration.LocalReference, "items_extra", "Success", "'items_extra' table inserted successfully!");
                }
                catch (Exception ex)
                {
                    logger.Log(LogLevels.Error, "An error occured while inserting to 'items_extra' table!Reason: " + ex.ToString());
                    Console.WriteLine("An error occured while inserting to 'items_extra' table!\nReason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "items_extra", "Error", "An error occured while inserting to 'items_extra' table!\nReason: " + ex.ToString());
                }
            }
        }

        private void InsertItemsPackages(int Index, int ItemsID)
        {
            for (int c = 0; c < Data.Declaration.Items[Index].Packages.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO items_packages " +
                            "(packagesMarks, " +
                            "packagesNumber, " +
                            "packageType, " +
                            "packedItemQuantity, " +
                            "packagesGrossMass, " +
                            "packagesGrossMassUnit, " +
                            "packageIdClientSystem, " +
                            "localReference, " +
                            "itemsID) " +
                            "VALUES " +
                            "(@packagesMarks, " +
                            "@packagesNumber, " +
                            "@packageType, " +
                            "@packedItemQuantity, " +
                            "@packagesGrossMass, " +
                            "@packagesGrossMassUnit, " +
                            "@packageIdClientSystem, " +
                            "@localReference, " +
                            "@itemsID);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);

                    Cmd.Parameters.AddWithValue("@packagesMarks", Data.Declaration.Items[Index] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].Marks) ? (object)"No Value" : (object)Data.Declaration.Items[Index].Packages[c].Marks);
                    Cmd.Parameters.AddWithValue("@packagesNumber", Data.Declaration.Items[Index] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].Number) ? (object)"No Value" : (object)Data.Declaration.Items[Index].Packages[c].Number);
                    Cmd.Parameters.AddWithValue("@packageType", Data.Declaration.Items[Index] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].PackageType) ? (object)"No Value" : (object)Data.Declaration.Items[Index].Packages[c].PackageType);
                    Cmd.Parameters.AddWithValue("@packedItemQuantity", Data.Declaration.Items[Index] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].packedItemQuantity) ? (object)"No Value" : (object)Data.Declaration.Items[Index].Packages[c].packedItemQuantity);
                    Cmd.Parameters.AddWithValue("@packagesGrossMass", Data.Declaration.Items[Index] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].PackageGrossMass.Value) ? (object)"No Value" : (object)Data.Declaration.Items[Index].Packages[c].PackageGrossMass.Value);
                    Cmd.Parameters.AddWithValue("@packagesGrossMassUnit", Data.Declaration.Items[Index] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].PackageGrossMass.Unit) ? (object)"No Value" : (object)Data.Declaration.Items[Index].Packages[c].PackageGrossMass.Unit);
                    Cmd.Parameters.AddWithValue("@packageIdClientSystem", Data.Declaration.Items[Index] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].packageIdClientSystem) ? (object)"No Value" : (object)Data.Declaration.Items[Index].Packages[c].packageIdClientSystem);
                    Cmd.Parameters.AddWithValue("@localReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);
                    Cmd.Parameters.AddWithValue("@itemsID", ItemsID);

                    Cmd.ExecuteNonQuery();

                    Cmd.Dispose();

                    string[] values = {
                        Data.Declaration.Items[Index].Packages[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].Marks) ? "No Value" : Data.Declaration.Items[Index].Packages[c].Marks,
                        Data.Declaration.Items[Index].Packages[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].Number) ? "No Value" : Data.Declaration.Items[Index].Packages[c].Number,
                        Data.Declaration.Items[Index].Packages[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].PackageType) ? "No Value" : Data.Declaration.Items[Index].Packages[c].PackageType,
                        Data.Declaration.Items[Index].Packages[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].packedItemQuantity) ? "No Value" : Data.Declaration.Items[Index].Packages[c].packedItemQuantity,
                        Data.Declaration.Items[Index].Packages[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].PackageGrossMass.Value) ? "No Value" : Data.Declaration.Items[Index].Packages[c].PackageGrossMass.Value,
                        Data.Declaration.Items[Index].Packages[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].PackageGrossMass.Unit) ? "No Value" : Data.Declaration.Items[Index].Packages[c].PackageGrossMass.Unit,
                        Data.Declaration.Items[Index].Packages[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Items[Index].Packages[c].packageIdClientSystem) ? "No Value" : Data.Declaration.Items[Index].Packages[c].packageIdClientSystem,
                        ItemsID.ToString()
                    };

                    InsertHistory(Data.Declaration.LocalReference, "items_packages", "Success", "'items_packages' table inserted successfully!");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occured while inserting to 'items_packages' table!\nReason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "item_packages", "Error", "An error occured while inserting to 'items_packages' table!\nReason: " + ex.ToString());
                }
            }
        }

        private void InsertParties()
        {
            for (int c = 0; c < Data.Declaration.Parties.Count; c++)
            {
                try
                {
                    string Query = "INSERT INTO parties " +
                            "(city, " +
                            "countryCode, " +
                            "name, " +
                            "partyType, " +
                            "postalCode, " +
                            "street, " +
                            "traderID, " +
                            "vatID, " +
                            "companyCode, " +
                            "commercialReference, " +
                            "localReference) " +
                            "VALUES " +
                            "(@city, " +
                            "@countryCode, " +
                            "@name, " +
                            "@partyType, " +
                            "@postalCode, " +
                            "@street, " +
                            "@traderID, " +
                            "@vatID, " +
                            "@companyCode, " +
                            "@commercialReference, " +
                            "@localReference);";

                    SqlCommand Cmd = new SqlCommand(Query, Con);

                    Cmd.Parameters.AddWithValue("@city", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].City) ? (object)"No Value" : (object)Data.Declaration.Parties[c].City);
                    Cmd.Parameters.AddWithValue("@countryCode", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].CountryCode) ? (object)"No Value" : (object)Data.Declaration.Parties[c].CountryCode);
                    Cmd.Parameters.AddWithValue("@name", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].Name) ? (object)"No Value" : (object)Data.Declaration.Parties[c].Name);
                    Cmd.Parameters.AddWithValue("@partyType", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].PartyType) ? (object)"No Value" : (object)Data.Declaration.Parties[c].PartyType);
                    Cmd.Parameters.AddWithValue("@postalCode", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].PostalCode) ? (object)"No Value" : (object)Data.Declaration.Parties[c].PostalCode);
                    Cmd.Parameters.AddWithValue("@street", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].Street) ? (object)"No Value" : (object)Data.Declaration.Parties[c].Street);
                    Cmd.Parameters.AddWithValue("@traderID", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].TraderId) ? (object)"No Value" : (object)Data.Declaration.Parties[c].TraderId);
                    Cmd.Parameters.AddWithValue("@vatID", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].VatId) ? (object)"No Value" : (object)Data.Declaration.Parties[c].VatId);
                    Cmd.Parameters.AddWithValue("@companyCode", Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].CompanyCode) ? (object)"No Value" : (object)Data.Declaration.Parties[c].CompanyCode);
                    Cmd.Parameters.AddWithValue("@commercialReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? (object)"No Value" : (object)Data.Declaration.CommercialReference);
                    Cmd.Parameters.AddWithValue("@localReference", Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference);

                    logger.Log(LogLevels.Debug, "INSERT INTO parties " +
                            "(city, " +
                            "countryCode, " +
                            "name, " +
                            "partyType, " +
                            "postalCode, " +
                            "street, " +
                            "traderID, " +
                            "vatID, " +
                            "companyCode, " +
                            "commercialReference, " +
                            "localReference) " +
                            "VALUES " +
                            "(" + (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].City) ? (object)"No Value" : (object)Data.Declaration.Parties[c].City) + ", " +
                            (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].CountryCode) ? (object)"No Value" : (object)Data.Declaration.Parties[c].CountryCode) + ", " +
                            (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].Name) ? (object)"No Value" : (object)Data.Declaration.Parties[c].Name) + ", " +
                            (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].PartyType) ? (object)"No Value" : (object)Data.Declaration.Parties[c].PartyType) + ", " +
                            (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].PostalCode) ? (object)"No Value" : (object)Data.Declaration.Parties[c].PostalCode) + ", " +
                            (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].Street) ? (object)"No Value" : (object)Data.Declaration.Parties[c].Street) + ", " +
                            (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].TraderId) ? (object)"No Value" : (object)Data.Declaration.Parties[c].TraderId) + ", " +
                            (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].VatId) ? (object)"No Value" : (object)Data.Declaration.Parties[c].VatId) + ", " +
                            (Data.Declaration.Parties[c] == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].CompanyCode) ? (object)"No Value" : (object)Data.Declaration.Parties[c].CompanyCode) + ", " +
                            (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? (object)"No Value" : (object)Data.Declaration.CommercialReference) + ", " +
                            (Data.Declaration == null ? (object)"No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? (object)"No Value" : (object)Data.Declaration.LocalReference) + ");");

                    Cmd.ExecuteNonQuery();

                    Cmd.Dispose();

                    string[] values = {
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].City) ? "No Value" : Data.Declaration.Parties[c].City,
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].CountryCode) ? "No Value" : Data.Declaration.Parties[c].CountryCode,
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].Name) ? "No Value" : Data.Declaration.Parties[c].Name,
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].PartyType) ? "No Value" : Data.Declaration.Parties[c].PartyType,
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].PostalCode) ? "No Value" : Data.Declaration.Parties[c].PostalCode,
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].Street) ? "No Value" : Data.Declaration.Parties[c].Street,
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].TraderId) ? "No Value" : Data.Declaration.Parties[c].TraderId,
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].VatId) ? "No Value" : Data.Declaration.Parties[c].VatId,
                        Data.Declaration.Parties[c] == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.Parties[c].CompanyCode) ? "No Value" : Data.Declaration.Parties[c].CompanyCode,
                        Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.CommercialReference) ? "No Value" : Data.Declaration.CommercialReference,
                        Data.Declaration == null ? "No Value" : String.IsNullOrWhiteSpace(Data.Declaration.LocalReference) ? "No Value" : Data.Declaration.LocalReference
                    };

                    Form_Home.UpdateParties(values);

                    InsertHistory(Data.Declaration.LocalReference, "parties", "Success", "'parties' table inserted successfully!");
                }
                catch (Exception ex)
                {
                    logger.Log(LogLevels.Error, "An error occured while inserting to 'parties' table!Reason: " + ex.ToString());
                    Console.WriteLine("An error occured while inserting to 'parties' table!\nReason: " + ex.ToString());
                    InsertHistory(Data.Declaration.LocalReference, "parties", "Error", "An error occured while inserting to 'parties' table!\nReason: " + ex.ToString());
                }
            }
        }
        public void CreatePeachTxt()
        {

            try
            {
                string queryFileName = "select TotalNetWeight,PHYTOcert from vw_export2peachtxt_filename where vw_export2peachtxt_filename.localReference =" + "'" + Data.Declaration.LocalReference + "'";
                List<string> tempList = new List<string>();

                SqlCommand command = new SqlCommand(queryFileName, Con);

                SqlDataReader reader = command.ExecuteReader();

                while (reader.Read())
                {
                    if (!reader.IsDBNull(0))
                    {
                        tempList.Add(reader[0].ToString() + "KG" + " " + reader[1].ToString());
                    }
                }
                reader.Close();
                string[] fileName = tempList.ToArray();



                //Declare Variables and provide values
                string FileNamePart = fileName[0];
                string DestinationFolder = @"D:\HARDDISK\INTFILES\AEB2TRANSLIMA\Attachments\" + Data.Declaration.LocalReference + "\\PDF ";
                string FileDelimiter = ","; // , seperated
                string FileExtension = ".txt"; //Extension




                //Read data from table or view to data table
                string query = "SELECT genus,species,quantity,identifier FROM vw_export2peachtxt_content WHERE vw_export2peachtxt_content.localReference =" + "'" + Data.Declaration.LocalReference + "'";
                SqlCommand cmd = new SqlCommand(query, Con);
                DataTable d_table = new DataTable();
                d_table.Load(cmd.ExecuteReader());

                //Prepare the file path 
                string FileFullPath = DestinationFolder + "\\" + FileNamePart + FileExtension;

                StreamWriter sw = null;
                sw = new StreamWriter(FileFullPath, false);

                // Write the Header Row to File
                int ColumnCount = d_table.Columns.Count;
                for (int ic = 0; ic < ColumnCount; ic++)
                {
                    sw.Write(d_table.Columns[ic]);
                    if (ic < ColumnCount - 1)
                    {
                        sw.Write(FileDelimiter);
                    }
                }
                sw.Write(sw.NewLine);

                // Write All Rows to the File
                foreach (DataRow dr in d_table.Rows)
                {
                    for (int ir = 0; ir < ColumnCount; ir++)
                    {
                        if (!Convert.IsDBNull(dr[ir]))
                        {
                            sw.Write(dr[ir].ToString());
                        }
                        if (ir < ColumnCount - 1)
                        {
                            sw.Write(FileDelimiter);
                        }
                    }
                    sw.Write(sw.NewLine);

                }

                sw.Close();

            }
            catch (Exception ex)
            {
                logger.Log(LogLevels.Error, "An error occured while trying to create the Peach CSV file: " + ex.ToString());
                Console.WriteLine("An error occured while trying to create the Peach CSV file\nReason: " + ex.ToString());
                InsertHistory(Data.Declaration.LocalReference, "stam_PEACHidentifier", "Error", "An error occured while trying to create the Peach CSV file!");
            }
        }

        private void SetDashboard()
        {
            string[] values = {
                GetCommercialCount().ToString(),
                GetCommercialExtraCount().ToString(),
                GetCustomsOfficesCount().ToString(),
                GetInvoicesCount().ToString(),
                GetItemsCount().ToString(),
                GetItemsExtraCount().ToString(),
                GetPartiesCount().ToString()
            };

            Form_Home.UpdateDashboard(values);
        }

        private int GetCommercialCount()
        {
            try
            {
                string Query = "SELECT COUNT(*) FROM commercial;";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                int count = (Int32)Cmd.ExecuteScalar();

                Cmd.Dispose();

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured while getting 'commercial' row count!\nReason: " + ex.ToString());

                return 0;
            }
        }

        private int GetCommercialExtraCount()
        {
            try
            {
                string Query = "SELECT COUNT(*) FROM commercial_extra;";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                int count = (Int32)Cmd.ExecuteScalar();

                Cmd.Dispose();

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured while getting 'commercial_extra' row count!\nReason: " + ex.ToString());

                return 0;
            }
        }

        private int GetCustomsOfficesCount()
        {
            try
            {
                string Query = "SELECT COUNT(*) FROM customs_offices;";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                int count = (Int32)Cmd.ExecuteScalar();

                Cmd.Dispose();

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured while getting 'customs_offices' row count!\nReason: " + ex.ToString());

                return 0;
            }
        }

        private int GetInvoicesCount()
        {
            try
            {
                string Query = "SELECT COUNT(*) FROM invoices;";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                int count = (Int32)Cmd.ExecuteScalar();

                Cmd.Dispose();

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured while getting 'invoices' row count!\nReason: " + ex.ToString());

                return 0;
            }
        }

        private int GetItemsCount()
        {
            try
            {
                string Query = "SELECT COUNT(*) FROM items;";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                int count = (Int32)Cmd.ExecuteScalar();

                Cmd.Dispose();

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured while getting 'items' row count!\nReason: " + ex.ToString());

                return 0;
            }
        }

        private int GetItemsExtraCount()
        {
            try
            {
                string Query = "SELECT COUNT(*) FROM items_extra;";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                int count = (Int32)Cmd.ExecuteScalar();

                Cmd.Dispose();

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured while getting 'items_extra' row count!\nReason: " + ex.ToString());

                return 0;
            }
        }

        private int GetPartiesCount()
        {
            try
            {
                string Query = "SELECT COUNT(*) FROM parties;";

                SqlCommand Cmd = new SqlCommand(Query, Con);

                int count = (Int32)Cmd.ExecuteScalar();

                Cmd.Dispose();

                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occured while getting 'parties' row count!\nReason: " + ex.ToString());

                return 0;
            }
        }
    }

    class DUCR
    {
        public string GoodsDescription { get; set; }
        public string GrossWeight { get; set; }
        public string InvoiceNumbers { get; set; }
        public string ItemNumber { get; set; }
        public string MaterialNumber { get; set; }
        public string NetPrice { get; set; }
        public string NetWeight { get; set; }
        public string OriginCountryCode { get; set; }
        public string PackageMarks { get; set; }
        public string PackageNumber { get; set; }
        public string PackageType { get; set; }
        public string packedItemQuantity { get; set; }
        public string PreferentialCountryCode { get; set; }
        public string SequenceNumber { get; set; }
        public string TariffNumber { get; set; }
        public string StatisticalValue { get; set; }
        public string DUCRNumber { get; set; }
        public string CommercialReference { get; set; }
        public string LocalReference { get; set; }
        public int Index { get; set; }

        public string packageIdClientSystem { get; set; }

        public string amount { get; set; }
    }
}
