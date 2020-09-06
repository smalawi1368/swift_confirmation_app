using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OracleClient;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Timers;


//final swift confiramation Service by Alawi 8-25-2020
namespace SWIFT_Conf_Services
{
    public partial class Service1 : ServiceBase
    {
        //public values
        static string CustomeEmail;
        static string BankName;
        static string cust_no;
        static string strConnString = ConfigurationManager.ConnectionStrings["OracleCon"].ConnectionString;
        public static OracleConnection con = new OracleConnection(strConnString);
        OracleDataAdapter orada;
        DataTable FileData_tbl = new DataTable();

        private Timer tm = null;
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            tm = new Timer();
            tm.Interval = 10000;
            tm.Elapsed += new ElapsedEventHandler(this.Timer_tick);
            tm.Enabled = true;
            try
            {
                SwiftConfClass.WriteAppLogs("SWIFT Confirmation App Started");
            }
            catch (Exception exc)
            {
                SwiftConfClass.WriteAppLogs(exc.Message);
            }
        }

        protected override void OnStop()
        {
            SwiftConfClass.WriteAppLogs("SWIFT Confirmation App Stopped");
        }

        private void Timer_tick(object sender, ElapsedEventArgs e)
        {
            SwiftConfClass.WriteAppLogs("<<<<<>>>>>> THIS MESSAGE PASS TO SYSTEM AFTER 10 MIN <<<<<>>>>>");
            //create a data store---
            DataTable FileData_tbl = new DataTable();
            FileData_tbl.Columns.Add("FileName", typeof(string));
            DataRow row = FileData_tbl.NewRow();
            //----------------------
            //Date format
            DateTime dt = DateTime.Now;
            string month = DateTime.Now.Month.ToString();
            int MonthConvert = Int16.Parse(month);
            int PerMonth = MonthConvert;
            string monthName = (CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(PerMonth)).ToLower();
            string dateFormat = "swift_" + DateTime.Now.Year + "_" + monthName.Substring(0, 3) + "_" + DateTime.Now.Day;

            DirectoryInfo dir = new DirectoryInfo(@"D:\EMS\telex\Telex_pending");//
            FileInfo[] Files = dir.GetFiles("*.out");

            foreach (FileInfo file in Files)
            {
                string[] separatingStrings = { "{1:" };
                string contents = File.ReadAllText(@"D:\EMS\telex\Telex_pending\" + file.Name);

                string[] words = contents.Split(separatingStrings, System.StringSplitOptions.RemoveEmptyEntries);
                string FolderName = file.Name.Substring(1, 8);
                FileData_tbl.Rows.Add(FolderName);
                //----------------------------------
                for (int i = 0; i <= words.Length - 1; i++)
                {
                    string strVal = words[i];

                    //COPY STRVALTO I.TXT AND SAVE
                    Directory.CreateDirectory(@"D:\EMS\Telex\" + dateFormat + @"\" + FolderName);
                    File.WriteAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + FolderName + @"\" + i + ".txt", "{1:"+strVal);
                    //EMPTY STRVAL FOR NEXT VALUE
                    strVal = String.Empty;
                }
                try { File.Move(@"D:\EMS\telex\Telex_pending\" + file.Name, @"D:\EMS\Telex\Processed\" + file.Name); }
                catch (Exception exc)
                {
                    if (exc.Message.Contains("Cannot create a file when that file already exists"))
                    {
                        SwiftConfClass.WriteAppLogs(string.Format("file: {0} already exist in processed folder. moving to duplicate folder ...", file.Name));
                        try { File.Move(@"D:\EMS\Telex\" + file.Name, @"D:\EMS\Telex\Processed\duplicate\" + file.Name); }
                        catch (Exception _exc)
                        {
                            if (exc.Message.Contains("Cannot create a file when that file already exists"))
                            {
                                SwiftConfClass.WriteAppLogs(string.Format("file: {0} already exist in duplicate folder. deleting file ...", file.Name));
                                file.Delete();
                                return;
                            }
                            else
                                SwiftConfClass.WriteAppLogs(_exc.Message);
                        }
                        return;
                    }
                    else
                        SwiftConfClass.WriteAppLogs(exc.Message);
                }

                File.Delete(@"D:\EMS\Telex\" + dateFormat + @"\" + FolderName + @"\0.txt");
            }//ENDOFFOREEACH
             //-----------------------------------------------

            //step #2
            foreach (DataRow dr in FileData_tbl.Rows)
            {
                //FIND NUMBER OF FILE IN A DIRECTORY
                string path = @"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString();
                int fCount = Directory.GetFiles(path, "*", SearchOption.TopDirectoryOnly).Length;
                SwiftConfClass.WriteAppLogs("<<<       FILE " + dr["FileName"].ToString() + "out" + " IS PROGRESSING      >>>");
                //-----------------------WHEN THERE IS A FILE (only on file)---------------------------
                if (fCount < 2)
                {

                    //------------------------------------
                    //string RectDeclimator = "{2:";
                    string RecString = File.ReadAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
                    //int RecIndex = RecString.LastIndexOf(RectDeclimator);
                    int MT103position = RecString.IndexOf("I103");
                    string bic_code = RecString.Substring(MT103position + 4, 11);
                    //------------------------------------READ BANK NAME FROM FLEXCUBE----------------------------

                    if (con == null || con.State != ConnectionState.Open)
                    {
                        con.Open();
                    }
                    //              
                    string queryR = "SELECT bank_name FROM fccprod.ISTM_BIC_DIRECTORY@fc WHERE bic_code like '" + bic_code + "'";

                    OracleCommand cmdR = new OracleCommand(queryR, con);
                    OracleDataReader datardR = cmdR.ExecuteReader();
                    DataTable Bank_name_dataTbl = new DataTable();
                    Bank_name_dataTbl.Load(datardR);
                    foreach (DataRow dtrow in Bank_name_dataTbl.Rows)
                    {
                        BankName = dtrow["bank_name"].ToString();
                    }

                    //-------------------------------------
                    string accountDeclimator = ":50K:/";
                    string accountString = File.ReadAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
                    int accIndex = accountString.LastIndexOf(accountDeclimator);
                    string account = accountString.Substring(accIndex + 6, 16);
                    //-------------------------------------
                    string murDeclimator = "{108:";
                    string murString = File.ReadAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
                    int murIndex = murString.LastIndexOf(murDeclimator);
                    string mur = murString.Substring(murIndex + 5, 16);
                    //------------------------------FIND UTER-------------------------------------------
                    string uterDeclimator = "{121:";
                    string uterString = File.ReadAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
                    int uterIndex = uterString.LastIndexOf(uterDeclimator);
                    string uter = uterString.Substring(uterIndex + 5, 36);
                    //---------------------------------------------------------------------------
                    cust_no = account.Substring(7, 7);

                    if (con == null || con.State != ConnectionState.Open)
                    {
                        con.Open();
                    }
                    string query = "SELECT * from fccprod.sttm_customer@fc WHERE customer_no ='" + cust_no + "'";

                    OracleCommand cmd = new OracleCommand(query, con);
                    OracleDataReader datard = cmd.ExecuteReader();
                    DataTable udf_3_dataTbl = new DataTable();
                    udf_3_dataTbl.Load(datard);
                    foreach (DataRow dtrow in udf_3_dataTbl.Rows)
                    {
                        CustomeEmail = dtrow["udf_3"].ToString();
                    }
                    path = @"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt";
                    //generate pdf file
                    string LongMessage;
                    using (var streamReader = new StreamReader(path))
                    {
                        LongMessage = streamReader.ReadToEnd();
                    }
                    string emailBody = Environment.NewLine +
                    "<<<< Message Header >>>>"
                    + Environment.NewLine +
                    "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
                    "SWIFT INPUT:                        FIN 103 Single Customer Credit Trasfer" + Environment.NewLine + Environment.NewLine +
                    "                                    AFBIBAFKAXXX    " + Environment.NewLine + Environment.NewLine + Environment.NewLine +
                    "                                    AFGHANISTAN INTERNATIONAL BANK - KABUL" + Environment.NewLine + Environment.NewLine +
                                                                                                      Environment.NewLine +
                    "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
                   "Reciever:                           '" + bic_code + "'" + Environment.NewLine +
                    "                                   '" + BankName + "'" + Environment.NewLine +

                    "MUR:                               '" + mur + "'" + Environment.NewLine +
                    "UETR:                              '" + uter + "'" + Environment.NewLine + Environment.NewLine + Environment.NewLine +
                    "<<<< MESSAGE TEXT >>>>"
                      + Environment.NewLine +
                     "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
                     LongMessage + Environment.NewLine +
                     "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
                    "Thank you for using AIB" + Environment.NewLine +
                    DateTime.Today.ToLongDateString();

                    MigraDoc.DocumentObjectModel.Document document = CreateDocument(emailBody);
                    document.UseCmykColor = true;
                    const bool unicode = false;
                    // const PdfFontEmbedding embedding = PdfFontEmbedding.Always;
                    PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode);
                    pdfRenderer.Document = document;
                    pdfRenderer.RenderDocument();
                    // Save the document...
                    string filename = @"D:\EMS\Telex\pdf\" + @"\" + uter + ".pdf";
                    try
                    {
                        pdfRenderer.PdfDocument.Save(filename);
                    }
                    catch (Exception _exc)
                    {
                        SwiftConfClass.WriteAppLogs(_exc.Message);
                    }

 

                    //-----------------------------------EMAIL sending portion-------------------------------------------------
                    string Msg = SwiftConfClass.PopulateBody (path, account, uter, mur);
                    path = String.Empty;
                    //SendEmail(ToEmail, Subj, Message, account, uetr, mur,address,account)
                   SwiftConfClass.SendEmail(CustomeEmail, "SWIFT Confirmation", Msg, mur, filename, account, dr["FileName"].ToString());

                    //--------------------------------------
                } // end of if(file<2)
                else
                {
                    // I WILL WORK ON IT SOON

                    //Step #1 count number of file in the folder

                    for (int i = 1; i <= fCount; i++)
                    {

                        string RecString = File.ReadAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
                        //int RecIndex = RecString.LastIndexOf(RectDeclimator);
                        int MT103position = RecString.IndexOf("I103");
                        string bic_code = RecString.Substring(MT103position + 4, 11);
                        //------------------------------------READ BANK NAME FROM FLEXCUBE----------------------------

                        if (con == null || con.State != ConnectionState.Open)
                        {
                            con.Open();
                        }
                        //              
                        string queryR = "select bank_name from fccprod.ISTM_BIC_DIRECTORY@fc where bic_code like '" + bic_code + "'";

                        OracleCommand cmdR = new OracleCommand(queryR, con);
                        OracleDataReader datardR = cmdR.ExecuteReader();
                        DataTable Bank_name_dataTbl = new DataTable();
                        Bank_name_dataTbl.Load(datardR);
                        foreach (DataRow dtrow in Bank_name_dataTbl.Rows)
                        {
                            BankName = dtrow["bank_name"].ToString();
                        }
                        //-------------------------------------------------------
                        string accountDeclimator = ":50K:/";
                        string accountString = File.ReadAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\" + i + ".txt");
                        int accIndex = accountString.LastIndexOf(accountDeclimator);
                        string account = accountString.Substring(accIndex + 6, 16);
                        //-------------------------------------
                        string murDeclimator = "{108:";
                        string murString = File.ReadAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\" + i + ".txt");
                        int murIndex = murString.LastIndexOf(murDeclimator);
                        string mur = murString.Substring(murIndex + 5, 16);
                        //------------------------------FIND UTER-------------------------------------------
                        string uterDeclimator = "{121:";
                        string uterString = File.ReadAllText(@"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\" + i + ".txt");
                        int uterIndex = uterString.LastIndexOf(uterDeclimator);
                        string uter = uterString.Substring(uterIndex + 5, 36);
                        cust_no = account.Substring(7, 7);

                        //-------------------------------------------------
                        if (con == null || con.State != ConnectionState.Open)
                        {
                            con.Open();
                        }
                        string query = "SELECT * from fccprod.sttm_customer@fc where customer_no ='" + cust_no + "'";

                        // string query = "select UDF_3 from sttm_cust_account@fc where upper(ac_desc) like  '%" + txtSearch.Text.ToUpper() + "%'";
                        OracleCommand cmd = new OracleCommand(query, con);
                        OracleDataReader datard = cmd.ExecuteReader();
                        DataTable udf_3_dataTbl = new DataTable();
                        udf_3_dataTbl.Load(datard);
                        foreach (DataRow dtrow in udf_3_dataTbl.Rows)
                        {
                            CustomeEmail = dtrow["udf_3"].ToString();
                        }
                        path = @"D:\EMS\Telex\" + dateFormat + @"\" + dr["FileName"].ToString() + @"\" + i + ".txt";
                        //generate pdf file
                        string LongMessage;
                        using (var streamReader = new StreamReader(path))
                        {
                            LongMessage = streamReader.ReadToEnd();
                        }
                        string emailBody = Environment.NewLine +
                        "<<< Message Header >>>"
                        + Environment.NewLine +
                        "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
                        "SWIFT INPUT:                        FIN 103 Single Customer Credit Trasfer" + Environment.NewLine + Environment.NewLine +
                        "                                    AFBIBAFKAXXX    " + Environment.NewLine + Environment.NewLine + Environment.NewLine +
                        "                                    AFGHANISTAN INTERNATIONAL BANK - KABUL" + Environment.NewLine + Environment.NewLine +
                                                                                                          Environment.NewLine +
                        "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine +
                        "Reciever:                          '" + bic_code + "'" + Environment.NewLine +
                        "                                   '" + BankName + "'" + Environment.NewLine +

                        "MUR:                               '" + mur + "'" + Environment.NewLine +
                        "UETR:                              '" + uter + "'" + Environment.NewLine + Environment.NewLine + Environment.NewLine +
                        "<<< MESSAGE TEXT >>>"
                          + Environment.NewLine +
                         "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ " + Environment.NewLine + Environment.NewLine +
                        "{1:" + LongMessage + Environment.NewLine +
                         "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
                        "Thank you for using AIB" + Environment.NewLine +
                        DateTime.Today.ToLongDateString();

                        MigraDoc.DocumentObjectModel.Document document = CreateDocument(emailBody);
                        document.UseCmykColor = true;
                        const bool unicode = false;
                        // const PdfFontEmbedding embedding = PdfFontEmbedding.Always;
                        PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode);
                        pdfRenderer.Document = document;
                        pdfRenderer.RenderDocument();
                        // Save the document...
                        string filename = @"D:\EMS\Telex\pdf\" + @"\" + uter + ".pdf";
                        try
                        {
                            pdfRenderer.PdfDocument.Save(filename);
                        }
                        catch (Exception _exc)
                        {
                            SwiftConfClass.WriteAppLogs(_exc.Message);
                        }
                        //-----------------------------------EMAIL sending portion-------------------------------------------------
                        string Msg = SwiftConfClass.PopulateBody(path, account, uter, mur);
                        path = String.Empty;
                        //SendEmail(ToEmail, Subj, Message, account, uetr, mur,address,account)// prototype
                         SwiftConfClass.SendEmail(CustomeEmail, "SWIFT Confirmation", Msg, mur, filename, account, dr["FileName"].ToString());
                    }//
                }
            }
        }
        public static MigraDoc.DocumentObjectModel.Document CreateDocument(string input)
        {
            MigraDoc.DocumentObjectModel.Document document = new MigraDoc.DocumentObjectModel.Document();
            document.DefaultPageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Portrait;
            document.DefaultPageSetup.HeaderDistance = 1;
            Section section = document.AddSection();
            section.PageSetup.TopMargin = 90;
            section.PageSetup.BottomMargin = 20;
            section.PageSetup.RightMargin = 5;
            section.PageSetup.LeftMargin = 50;
            Paragraph header = section.Headers.Primary.AddParagraph();
            MigraDoc.DocumentObjectModel.Shapes.Image img = new MigraDoc.DocumentObjectModel.Shapes.Image();
            header.AddImage(@"D:\EMS\Telex\white.png");
            Paragraph paragraph = section.AddParagraph();
            MigraDoc.DocumentObjectModel.Shapes.Image img2 = new MigraDoc.DocumentObjectModel.Shapes.Image();
            header.AddImage(@"D:\EMS\Telex\logo.png");
            Paragraph paragraph2 = section.AddParagraph();
            paragraph.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 0, 0, 100);
            paragraph.Format.Font.Name = "Courier New";
            paragraph.Format.Font.Size = 9;
            paragraph.AddText(input.Replace(" ", " "));
            return document;
        }//--------------------------------------



    }
}
