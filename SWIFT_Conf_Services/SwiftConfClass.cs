using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Timers;
using System.IO;
using System.Configuration;
using System.Data.OracleClient;
using MigraDoc.DocumentObjectModel;
using System.Globalization;
using System.Data;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using System.Net.Mail;
using System.Net.Mime;
using System.Net;

namespace SWIFT_Conf_Services
{
    class SwiftConfClass
    {
        static string strVal;


        public static void  WriteAppLogss(Exception ex)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\Logfile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + ex.Source.ToString().Trim() + ";" + ex.Message.ToString());
                sw.Flush();
                sw.Close();
            }
            catch
            {

            }
        }
        public static void WriteAppLogs(string Msg)
        {
            StreamWriter sw = null;
            try
            {
                sw = new StreamWriter(AppDomain.CurrentDomain.BaseDirectory + "\\Logfile.txt", true);
                sw.WriteLine(DateTime.Now.ToString() + ": " + Msg);
                sw.Flush();
                sw.Close();
            }
            catch
            {

            }
        }
        //public static void LongMessageGenerator()
        //{





        //}
        public static string PopulateBody(string path, string account, string uetr, string mur)
        {
            string body = string.Empty;
            string[] separatingStrings = { "{4:" };
            string contents = File.ReadAllText(path);
            string[] words = contents.Split(separatingStrings, System.StringSplitOptions.RemoveEmptyEntries);
            for (int i = 0; i <= words.Length - 1; i++)
            {
                strVal = words[i];
                //File.WriteAllText(path+ @"\" + i + ".txt", strVal);
            }
            body = strVal;
            string HTMLBody = "<div style='width:500px; margin:0 auto; font-size:9px'><span style='float:left; height: 15px; margin-top:5px; padding-bottom:30px;'><img mc:edit='logo1' src='https://www.aib.af/aib_public/images/AIB%20Logo_new.png ' width='70' height='40' alt='Afghanistan International Bank' /></span><div style='line-height:50px; background-color:#1d3c7f; height:50px; padding-left:25px ; color:#fff;font-family:Bahnschrift'> Afghanistan International Bank</div><div style='line-height:-10px; background-color:#1d3c7f; height:10px; padding:5px; margin-top:3px; ;color:#fff;font-family:Bahnschrift'>SWIFT Confirmation Message</div><div><h1 style='font-family:Bahnschrift; font-size:14px; padding: 5px 10px;'>Dear Customer,</h1><p style='font-family:Bahnschrift; font-size:14px; padding: 10px 19px; text-align: justify; text-justify: inter-word;'>Please find attached, copy of swift message for your international tranfer. For enquiries or clarifications, kindly contact customer care on +93(0) 20 255 0 255 or <a href='mailto:payment.order@aib.af'> email </a> us.</p><p style='font-family:Bahnschrift; font-size:10px; padding: 10px 19px; '>Afghanistan International Bank<br>Your partner to growth</p><hr style='border-bottom:35px solid #00A94F' /></div></div><style> .card { display: inline; margin: 0 auto; text-align: center; }</style><div style='text-align:center'> <a href='https://www.aib.af/aib_public/pspb-debit-Card.asp'><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/SmallBanners/Debit-Card.jpg' width='140' height='100' alt='Afghanistan International Bank'/></a> <a href='https://www.aib.af/aib_public/pspb-titanium-Card.asp'><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/SmallBanners/card-tcc.jpg' width='140' height='100' alt='Afghanistan International Bank' /></a> <a href='https://www.aib.af/aib_public/pspb-Credit-Card.asp'><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/SmallBanners/Credit-Card.jpg' width='140' height='100' alt='Afghanistan International Bank' /></a> <div> <table cellpadding='5' style='margin:0 auto; text-align:center'> <tr><td><a href='https://www.facebook.com/AfghanistanIntBank/'> <i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/facebook.png' width='20' height='20' alt='Afghanistan International Bank' /></i></a></td> <td><a href='https://www.linkedin.com/company/afghanistan-international-bank'><i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/linkedin.png' width='20' height='20' alt='Afghanistan International Bank' /></i></a></td> <td><a href='https://twitter.com/AfgIntBank'> <i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/twitter.png' height='20' alt='Afghanistan International Bank' /></i></a></li></td> <td><a href='https://www.instagram.com/AIB.Bank/'><i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/instagram.png' width='20' height='20' alt='Afghanistan International Bank' /></i></a></td> <td><a href='https://wa.me/93700001111'><i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/whatsapp.png' width='20' height='20' alt='Afghanistan International Bank' /></i></a></td> <td><a href='https://www.youtube.com/channel/UCzy1u-1kfApizvLTXnVmF7w'> <i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/youtube.png' width='22' height='24' alt='Afghanistan International Bank'/></i></a></td> </tr> </table> </div> <div style='padding-left:10px;'> <br /> <p dir=ltr style='font-family:Lateef,Bahnschrift,Arial; margin-top:-20px; font-size:12px;color:#1d3c7f;'> Address : Airport Road, BiBi Mahro P.O.Box No. 2074, Kabul Afghanistan.<a href='https://www.google.com/maps/place//@34.5444337,69.1972505,646m/am=t/data=!3m1!1e3!4m16!1m15!4m14!1m6!1m2!1s0x38d16d4de2854935:0x1ced01add7a3c724!2sAIB+-+Afghanistan+International+Bank+HO,+Airport+Road,+Kabul!2m2!1d69.1993447!2d34.5436522!1m6!1m2!1s0x38d16c18e5e8135f:0xebf50fc6d4d79ee9!2zQXppemkgUGxhemEg2LnYstuM2LLbjCDZvtmE2KfYstinLCA0dGggTWFjcm9yeWFuIE1haW4gUm9hZCwgS2Fib2wsIEFmZ2hhbmlzdGFu!2m2!1d69.2002654!2d34.5451926!5m1!1e1?hl=en' target='_blank'>Google Map</a> <br />Tel: +93(0) 20 255 0 255 <br /> <a style='font-family:Lateef' href='https://www.aib.af/'>Online Customer Care</a> </p> </div></div>";
            //string HTMLBody = "<div style='width:500px; margin:0 auto; font-size:12px'><div style='background-color:#1d3c7f; height:10px; padding:5px ; color:#fff;font-family:Bahnschrift'>Instance Type and Transmistion</div><table style='font-family:Bahnschrift; font-size:12px; padding-left:20px;'><tr><td>Orginal</td><td></td></tr><tr><td style='padding-right:90px;'>Priority / Delivery</td><td>Normal</td></tr></table><br><div style=' background-color:#1d3c7f; height:10px; padding:5px ;color:#fff;font-family:Bahnschrift'>Message Header</div><table style='font-family:Bahnschrift; font-size:12px; padding-left:20px;'><tr><td>SWIFT Input</td><td style='font-family:Bahnschrift; font-size:12px;'>FIN 103 Single Customer Credit Trasfer</td></tr><tr style='text-align:right'><td></td><td>AFGHANISTAN INTERNATIONAL BANK</td></tr><tr><td style='padding-right:90px;'>Sender</td><td>AFIBAFKAXXX</td></tr><td style='padding-right:90px;'>Receiver</td><td>CRASGB2LXXX</td></tr><tr style='text-align:right; text-align:left'><td></td><td>LONDON GB</td></tr><tr><td style='padding-right:90px;'>MUR</td><td>" + mur + "</td></tr><tr><td style='padding-right:90px;'>UETR</td><td>" + uetr + "</td></tr></table><div style=' background-color:#1d3c7f; height:10px; padding:5px; color:#fff;font-family:Bahnschrift'>Message Text</div><p style='padding:10px 10px; font-family:Bahnschrift;white-space:pre'>" + body + "</p><hr style='background-color:#00a94f; height:50px' /></div><style>.card {display: inline; margin: 0 auto; text-align: center;}</style><div style='text-align:center'> <a href='https://www.aib.af/aib_public/pspb-debit-Card.asp'><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/SmallBanners/Debit-Card.jpg' width='140' height='100' alt='Afghanistan International Bank' /></a> <a href='https://www.aib.af/aib_public/pspb-titanium-Card.asp'><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/SmallBanners/card-tcc.jpg' width='140' height='100' alt='Afghanistan International Bank' /></a> <a href='https://www.aib.af/aib_public/pspb-Credit-Card.asp'><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/SmallBanners/Credit-Card.jpg' width='140' height='100' alt='Afghanistan International Bank' /></a> <div> <table cellpadding='5' style='margin:0 auto; text-align:center'> <tr> <td><a href='https://www.facebook.com/AfghanistanIntBank/'> <i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/facebook.png' width='20' height='20' alt='Afghanistan International Bank' /></i></a></td> <td><a href='https://www.linkedin.com/company/afghanistan-international-bank'><i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/linkedin.png' width='20' height='20' alt='Afghanistan International Bank' /></i></a></td> <td><a href='https://twitter.com/AfgIntBank'> <i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/twitter.png' height='20' alt='Afghanistan International Bank' /></i></a></li></td> <td><a href='https://www.instagram.com/AIB.Bank/'><i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/instagram.png' width='20' height='20' alt='Afghanistan International Bank' /></i></a></td> <td><a href=' https://wa.me/93700001111'><i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/whatsapp.png' width='20' height='20' alt='Afghanistan International Bank' /></i></a></td> <td><a href='https://www.youtube.com/channel/UCzy1u-1kfApizvLTXnVmF7w'> <i><img class='card' mc:edit='logo1' src='https://www.aib.af/aib_public/images/social_media_alerts/youtube.png' width='22' height='24' alt='Afghanistan International Bank' /></i></a></td> </tr> </table> </div> <div style='padding-left:10px;'> <br /> <p dir=ltr style='font-family:Lateef,Bahnschrift,Arial; margin-top:-20px; font-size:12px;color:#1d3c7f;'> Address : Airport Road, BiBi Mahro P.O.Box No. 2074, Kabul Afghanistan.<a href='https://www.google.com/maps/place//@34.5444337,69.1972505,646m/am=t/data=!3m1!1e3!4m16!1m15!4m14!1m6!1m2!1s0x38d16d4de2854935:0x1ced01add7a3c724!2sAIB+-+Afghanistan+International+Bank+HO,+Airport+Road,+Kabul!2m2!1d69.1993447!2d34.5436522!1m6!1m2!1s0x38d16c18e5e8135f:0xebf50fc6d4d79ee9!2zQXppemkgUGxhemEg2LnYstuM2LLbjCDZvtmE2KfYstinLCA0dGggTWFjcm9yeWFuIE1haW4gUm9hZCwgS2Fib2wsIEFmZ2hhbmlzdGFu!2m2!1d69.2002654!2d34.5451926!5m1!1e1?hl=en' target='_blank'>Google Map</a> <br />Tel: +93(0) 20 255 255 <br /> <a style='font-family:Lateef' href='https://www.aib.af/'>Online Customer Care</a> </p> </div></div> ";
            strVal = String.Empty;
            return HTMLBody;
        }
        //public static MigraDoc.DocumentObjectModel.Document CreateDocument(string input)
        //{
        //    MigraDoc.DocumentObjectModel.Document document = new MigraDoc.DocumentObjectModel.Document();
        //    document.DefaultPageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Portrait;
        //    document.DefaultPageSetup.HeaderDistance = 1;
        //    Section section = document.AddSection();
        //    section.PageSetup.TopMargin = 90;
        //    section.PageSetup.BottomMargin = 20;
        //    section.PageSetup.RightMargin = 5;
        //    section.PageSetup.LeftMargin = 50;
        //    Paragraph header = section.Headers.Primary.AddParagraph();
        //    MigraDoc.DocumentObjectModel.Shapes.Image img = new MigraDoc.DocumentObjectModel.Shapes.Image();
        //    header.AddImage(@"D:\EMS\Telex\white.png");
        //    Paragraph paragraph = section.AddParagraph();
        //    MigraDoc.DocumentObjectModel.Shapes.Image img2 = new MigraDoc.DocumentObjectModel.Shapes.Image();
        //    header.AddImage(@"D:\EMS\Telex\logo.png");
        //    Paragraph paragraph2 = section.AddParagraph();
        //    paragraph.Format.Font.Color = MigraDoc.DocumentObjectModel.Color.FromCmyk(0, 0, 0, 100);
        //    paragraph.Format.Font.Name = "Courier New";
        //    paragraph.Format.Font.Size = 9;
        //    paragraph.AddText(input.Replace(" ", " "));
        //    return document;
        //}
        //public static void FileCatagorizer()
        //{
        //    //create a data store---
        //    //create a data store---
        //    DataTable FileData_tbl = new DataTable();
        //    FileData_tbl.Columns.Add("FileName", typeof(string));
        //    DataRow row = FileData_tbl.NewRow();
        //    //----------------------
        //    //Date format
        //    DateTime dt = DateTime.Now;
        //    string month = DateTime.Now.Month.ToString();
        //    int MonthConvert = Int16.Parse(month);
        //    int PerMonth = MonthConvert;
        //    string monthName = (CultureInfo.CurrentCulture.DateTimeFormat.GetMonthName(PerMonth)).ToLower();
        //    string dateFormat = "swift_" + DateTime.Now.Year + "_" + monthName.Substring(0, 3) + "_" + DateTime.Now.Day;

        //    DirectoryInfo dir = new DirectoryInfo(ConfigurationManager.AppSettings["TelexPath"]);//
        //    FileInfo[] Files = dir.GetFiles("*.out");

        //    foreach (FileInfo file in Files)
        //    {
        //        string[] separatingStrings = { "{1:" };
        //        string contents = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + file.Name);

        //        string[] words = contents.Split(separatingStrings, System.StringSplitOptions.RemoveEmptyEntries);
        //        string FolderName = file.Name.Substring(1, 8);
        //        FileData_tbl.Rows.Add(FolderName);
        //        //----------------------------------
        //        for (int i = 0; i <= words.Length - 1; i++)
        //        {
        //            string strVal = words[i];

        //            //COPY STRVALTO I.TXT AND SAVE
        //            Directory.CreateDirectory(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + FolderName);
        //            File.WriteAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + FolderName + @"\" + i + ".txt", strVal);
        //            //EMPTY STRVAL FOR NEXT VALUE
        //            strVal = String.Empty;
        //        }
        //        try { File.Move(ConfigurationManager.AppSettings["TelexPath"] + file.Name, ConfigurationManager.AppSettings["ProccedPath"] + file.Name); }
        //        catch (Exception exc)
        //        {
        //            if (exc.Message.Contains("Cannot create a file when that file already exists"))
        //            {
        //                 WriteAppLogs(string.Format("file: {0} already exist in processed folder. moving to duplicate folder ...", file.Name));
        //                try { File.Move(ConfigurationManager.AppSettings["TelexPath"] + file.Name, ConfigurationManager.AppSettings["DupPath"] + file.Name); }
        //                catch (Exception _exc)
        //                {
        //                    if (exc.Message.Contains("Cannot create a file when that file already exists"))
        //                    {
        //                         WriteAppLogs(string.Format("file: {0} already exist in duplicate folder. deleting file ...", file.Name));
        //                        file.Delete();
        //                        return;
        //                    }
        //                    else
        //                         WriteAppLogs(_exc.Message);
        //                }
        //                return;
        //            }

        //            else
        //                 WriteAppLogs(exc.Message);
        //        }

        //        File.Delete(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + FolderName + @"\0.txt");
        //    }//ENDOFFOREEACH
        //     //-----------------------------------------------

        //    //step #2
        //    foreach (DataRow dr in FileData_tbl.Rows)
        //    {
        //        //FIND NUMBER OF FILE IN A DIRECTORY
        //        string path = ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString();
        //        int fCount = Directory.GetFiles(path, "*", SearchOption.TopDirectoryOnly).Length;
        //         WriteAppLogs("<<<       FILE " + dr["FileName"].ToString() + "out" + " IS PROGRESSING      >>>");
        //        //-----------------------WHEN THERE IS A FILE (only on file)---------------------------
        //        if (fCount < 2)
        //        {

        //            //------------------------------------
        //            //string RectDeclimator = "{2:";
        //            string RecString = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
        //            //int RecIndex = RecString.LastIndexOf(RectDeclimator);
        //            int MT103position = RecString.IndexOf("I103");
        //            string bic_code = RecString.Substring(MT103position + 4, 11);
        //            //------------------------------------READ BANK NAME FROM FLEXCUBE----------------------------

        //            if (con == null || con.State != ConnectionState.Open)
        //            {
        //                con.Open();
        //            }
        //            //              
        //            string queryR = "select bank_name from fccprod.ISTM_BIC_DIRECTORY@fc where bic_code like '" + bic_code + "'";

        //            OracleCommand cmdR = new OracleCommand(queryR, con);
        //            OracleDataReader datardR = cmdR.ExecuteReader();
        //            DataTable Bank_name_dataTbl = new DataTable();
        //            Bank_name_dataTbl.Load(datardR);
        //            foreach (DataRow dtrow in Bank_name_dataTbl.Rows)
        //            {
        //                BankName = dtrow["bank_name"].ToString();
        //            }


        //            //-------------------------------------
        //            string accountDeclimator = ":50K:/";
        //            string accountString = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
        //            int accIndex = accountString.LastIndexOf(accountDeclimator);
        //            string account = accountString.Substring(accIndex + 6, 16);
        //            //-------------------------------------
        //            string murDeclimator = "{108:";
        //            string murString = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
        //            int murIndex = murString.LastIndexOf(murDeclimator);
        //            string mur = murString.Substring(murIndex + 5, 16);
        //            //------------------------------FIND UTER-------------------------------------------
        //            string uterDeclimator = "{121:";
        //            string uterString = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
        //            int uterIndex = uterString.LastIndexOf(uterDeclimator);
        //            string uter = uterString.Substring(uterIndex + 5, 36);
        //            //---------------------------------------------------------------------------
        //            cust_no = account.Substring(7, 7);

        //            if (con == null || con.State != ConnectionState.Open)
        //            {
        //                con.Open();
        //            }
        //            string query = "SELECT * from fccprod.sttm_customer@fc where customer_no ='" + cust_no + "'";

        //            OracleCommand cmd = new OracleCommand(query, con);
        //            OracleDataReader datard = cmd.ExecuteReader();
        //            DataTable udf_3_dataTbl = new DataTable();
        //            udf_3_dataTbl.Load(datard);
        //            foreach (DataRow dtrow in udf_3_dataTbl.Rows)
        //            {
        //                CustomeEmail = dtrow["udf_3"].ToString();
        //            }
        //            path = ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt";
        //            //generate pdf file
        //            string LongMessage;
        //            using (var streamReader = new StreamReader(path))
        //            {
        //                LongMessage = streamReader.ReadToEnd();
        //            }
        //            string emailBody = Environment.NewLine +
        //            "<<< Message Header >>>"
        //            + Environment.NewLine +
        //            "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
        //            "SWIFT INPUT:                        FIN 103 Single Customer Credit Trasfer" + Environment.NewLine + Environment.NewLine +
        //            "                                    AFBIBAFKAXXX    " + Environment.NewLine + Environment.NewLine + Environment.NewLine +
        //            "                                    AFGHANISTAN INTERNATIONAL BANK - KABUL" + Environment.NewLine + Environment.NewLine +
        //                                                                                              Environment.NewLine +
        //            "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine +
        //            "Reciever:                          '" + bic_code + "'" + Environment.NewLine +
        //            "                                   '" + BankName + "'" + Environment.NewLine +

        //            "MUR:                               '" + mur + "'" + Environment.NewLine +
        //            "UETR:                              '" + uter + "'" + Environment.NewLine + Environment.NewLine + Environment.NewLine +
        //            "<<< MESSAGE TEXT >>>"
        //              + Environment.NewLine +
        //             "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ " + Environment.NewLine + Environment.NewLine +
        //            "{1:" + LongMessage + Environment.NewLine +
        //             "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
        //            "Thank you for using AIB" + Environment.NewLine +
        //            DateTime.Today.ToLongDateString();

        //            MigraDoc.DocumentObjectModel.Document document = CreateDocument(emailBody);
        //            document.UseCmykColor = true;
        //            const bool unicode = false;
        //            // const PdfFontEmbedding embedding = PdfFontEmbedding.Always;
        //            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode);
        //            pdfRenderer.Document = document;
        //            pdfRenderer.RenderDocument();
        //            // Save the document...
        //            string filename = ConfigurationManager.AppSettings["pdfPath"] + @"\" + uter + ".pdf";
        //            try
        //            {
        //                pdfRenderer.PdfDocument.Save(filename);
        //            }
        //            catch (Exception _exc)
        //            {
        //                WriteAppLogs(_exc.Message);
        //            }



        //            //-----------------------------------EMAIL sending portion-------------------------------------------------
        //            string Msg = PopulateBody(path, account, uter, mur);
        //            path = String.Empty;
        //            //SendEmail(ToEmail, Subj, Message, account, uetr, mur,address,account)
        //            SendEmail(CustomeEmail, "SWIFT Confirmation", Msg, mur, filename, account, dr["FileName"].ToString());

        //            //--------------------------------------
        //        } // end of if(file<2)
        //        else
        //        {
        //            // case of more than one file

        //            //Step #1 count number of file in the folder

        //            for (int i = 1; i <= fCount; i++)
        //            {

        //                string RecString = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\1.txt");
        //                //int RecIndex = RecString.LastIndexOf(RectDeclimator);
        //                int MT103position = RecString.IndexOf("I103");
        //                string bic_code = RecString.Substring(MT103position + 4, 11);
        //                //------------------------------------READ BANK NAME FROM FLEXCUBE----------------------------

        //                if (con == null || con.State != ConnectionState.Open)
        //                {
        //                    con.Open();
        //                }
        //                //              
        //                string queryR = "select bank_name from fccprod.ISTM_BIC_DIRECTORY@fc where bic_code like '" + bic_code + "'";

        //                OracleCommand cmdR = new OracleCommand(queryR, con);
        //                OracleDataReader datardR = cmdR.ExecuteReader();
        //                DataTable Bank_name_dataTbl = new DataTable();
        //                Bank_name_dataTbl.Load(datardR);
        //                foreach (DataRow dtrow in Bank_name_dataTbl.Rows)
        //                {
        //                    BankName = dtrow["bank_name"].ToString();
        //                }
        //                //-------------------------------------------------------
        //                string accountDeclimator = ":50K:/";
        //                string accountString = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\" + i + ".txt");
        //                int accIndex = accountString.LastIndexOf(accountDeclimator);
        //                string account = accountString.Substring(accIndex + 6, 16);
        //                //-------------------------------------
        //                string murDeclimator = "{108:";
        //                string murString = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\" + i + ".txt");
        //                int murIndex = murString.LastIndexOf(murDeclimator);
        //                string mur = murString.Substring(murIndex + 5, 16);
        //                //------------------------------FIND UTER-------------------------------------------
        //                string uterDeclimator = "{121:";
        //                string uterString = File.ReadAllText(ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\" + i + ".txt");
        //                int uterIndex = uterString.LastIndexOf(uterDeclimator);
        //                string uter = uterString.Substring(uterIndex + 5, 36);
        //                cust_no = account.Substring(7, 7);

        //                //-------------------------------------------------
        //                if (con == null || con.State != ConnectionState.Open)
        //                {
        //                    con.Open();
        //                }
        //                string query = "SELECT * from fccprod.sttm_customer@fc where customer_no ='" + cust_no + "'";

        //                // string query = "select UDF_3 from sttm_cust_account@fc where upper(ac_desc) like  '%" + txtSearch.Text.ToUpper() + "%'";
        //                OracleCommand cmd = new OracleCommand(query, con);
        //                OracleDataReader datard = cmd.ExecuteReader();
        //                DataTable udf_3_dataTbl = new DataTable();
        //                udf_3_dataTbl.Load(datard);
        //                foreach (DataRow dtrow in udf_3_dataTbl.Rows)
        //                {
        //                    CustomeEmail = dtrow["udf_3"].ToString();
        //                }
        //                path = ConfigurationManager.AppSettings["TelexPath"] + dateFormat + @"\" + dr["FileName"].ToString() + @"\" + i + ".txt";
        //                //generate pdf file
        //                string LongMessage;
        //                using (var streamReader = new StreamReader(path))
        //                {
        //                    LongMessage = streamReader.ReadToEnd();
        //                }
        //                string emailBody = Environment.NewLine +
        //                "<<< Message Header >>>"
        //                + Environment.NewLine +
        //                "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
        //                "SWIFT INPUT:                        FIN 103 Single Customer Credit Trasfer" + Environment.NewLine + Environment.NewLine +
        //                "                                    AFBIBAFKAXXX    " + Environment.NewLine + Environment.NewLine + Environment.NewLine +
        //                "                                    AFGHANISTAN INTERNATIONAL BANK - KABUL" + Environment.NewLine + Environment.NewLine +
        //                                                                                                  Environment.NewLine +
        //                "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine +
        //                "Reciever:                          '" + bic_code + "'" + Environment.NewLine +
        //                "                                   '" + BankName + "'" + Environment.NewLine +

        //                "MUR:                               '" + mur + "'" + Environment.NewLine +
        //                "UETR:                              '" + uter + "'" + Environment.NewLine + Environment.NewLine + Environment.NewLine +
        //                "<<< MESSAGE TEXT >>>"
        //                  + Environment.NewLine +
        //                 "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ " + Environment.NewLine + Environment.NewLine +
        //                "{1:" + LongMessage + Environment.NewLine +
        //                 "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
        //                "Thank you for using AIB" + Environment.NewLine +
        //                DateTime.Today.ToLongDateString();

        //                MigraDoc.DocumentObjectModel.Document document = CreateDocument(emailBody);
        //                document.UseCmykColor = true;
        //                const bool unicode = false;
        //                // const PdfFontEmbedding embedding = PdfFontEmbedding.Always;
        //                PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode);
        //                pdfRenderer.Document = document;
        //                pdfRenderer.RenderDocument();
        //                // Save the document...
        //                string filename = ConfigurationManager.AppSettings["TelexPath"] + @"\" + uter + ".pdf";
        //                try
        //                {
        //                    pdfRenderer.PdfDocument.Save(filename);
        //                }
        //                catch (Exception _exc)
        //                {
        //                    WriteAppLogs(_exc.Message);
        //                }
        //                //-----------------------------------EMAIL sending portion-------------------------------------------------
        //                string Msg = PopulateBody(path, account, uter, mur);
        //                path = String.Empty;
        //                //SendEmail(ToEmail, Subj, Message, account, uetr, mur,address,account)// prototype
        //                SendEmail(CustomeEmail, "SWIFT Confirmation", Msg, mur, filename, account, dr["FileName"].ToString());
        //            }//
        //        }
        //    }
        //}
        //public string  pdfBody(string bic_code, string Bankname, string mur, string uter,string LongMessage,string account,string filename, string path)
        //{
        //    string emailBody = Environment.NewLine +
        //             "<<< Message Header >>>"
        //             + Environment.NewLine +
        //             "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _"    + Environment.NewLine + Environment.NewLine +
        //             "SWIFT INPUT:                        FIN 103 Single Customer Credit Trasfer"       + Environment.NewLine + Environment.NewLine +
        //             "                                    AFBIBAFKAXXX    " + Environment.NewLine       + Environment.NewLine + Environment.NewLine +
        //             "                                    AFGHANISTAN INTERNATIONAL BANK - KABUL"       + Environment.NewLine + Environment.NewLine +
        //                                                                                                  Environment.NewLine +
        //             "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine +
        //             "Reciever:                          '" + bic_code + "'" + Environment.NewLine +
        //             "                                   '" + BankName + "'" + Environment.NewLine +

        //             "MUR:                               '" + mur + "'" + Environment.NewLine +
        //             "UETR:                              '" + uter + "'" + Environment.NewLine + Environment.NewLine + Environment.NewLine +
        //             "<<< MESSAGE TEXT >>>"
        //               + Environment.NewLine +
        //              "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ " + Environment.NewLine + Environment.NewLine +
        //             "{1:" + LongMessage + Environment.NewLine +
        //              "_ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _ _ __ _ _" + Environment.NewLine + Environment.NewLine +
        //             "Thank you for using AIB" + Environment.NewLine +
        //             DateTime.Today.ToLongDateString();

        //    MigraDoc.DocumentObjectModel.Document document = CreateDocument(emailBody);
        //    document.UseCmykColor = true;
        //    const bool unicode = false;
        //    PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(unicode);
        //    pdfRenderer.Document = document;
        //    pdfRenderer.RenderDocument();
        //    // Save the document...
        //    string Pdffilename = @"D:\EMS\Telex\pdf\" + @"\" + uter + ".pdf";
        //    try
        //    {
        //        pdfRenderer.PdfDocument.Save(Pdffilename);
        //    }
        //    catch (Exception _exc)
        //    {
        //        WriteAppLogs(_exc.Message);
        //    }

        //    //-----------------------------------EMAIL sending portion-------------------------------------------------
        //    string Msg = PopulateBody(path, account, uter, mur);
        //    path = String.Empty;
        //    //SendEmail(ToEmail, Subj, Message, account, uetr, mur,address,account)
        //    SendEmail(CustomeEmail, "SWIFT Confirmation", Msg, mur, Pdffilename, account, filename);


        //    return longMessage;
        //}

        public static void SendEmail(String ToEmail, String Subj, string Message, string mur, string attach, string account, string file_name)
        {
            //Reading sender Email credential from web.config file  

            try
            {
                string HostAdd = ConfigurationManager.AppSettings["Host"].ToString();//192.0.0.4
                string FromEmailid = ConfigurationManager.AppSettings["FromEmail"].ToString();//alrts@bankaib.af
                string Pass = ConfigurationManager.AppSettings["Pass"].ToString();//Aib@123$
                MailMessage mailMessage = new MailMessage();
                mailMessage.From = new MailAddress(FromEmailid); //From Email Id  
                mailMessage.Subject = Subj;
                mailMessage.Body = Message;
                Attachment data = new Attachment(
                attach, MediaTypeNames.Application.Octet); mailMessage.IsBodyHtml = true;
                mailMessage.Attachments.Add(data);
                string[] ToMuliId = ToEmail.Split(',',';');

                foreach (string ToEMailId in ToMuliId)
                {
                    mailMessage.To.Add(new MailAddress(ToEMailId)); //adding multiple TO Email Id  
                }
                //network and security related credentials
                NetworkCredential NetworkCred = new NetworkCredential();
                NetworkCred.UserName = mailMessage.From.Address;
                NetworkCred.Password = Pass;
                SmtpClient sC = new SmtpClient(HostAdd);
                sC.EnableSsl = false;
                sC.Port = 25;
                sC.Credentials = new NetworkCredential(FromEmailid, Pass);
                //sC.Send(mailMessage);
                sC.Send(mailMessage);
                //sending Email  
                WriteAppLogs("EMAIL HAS SENT TO: " + ToEmail + " by MUR: " + mur + " to Account:" + account + " from file: " + file_name);
            }
            catch (Exception ex)
            {
                WriteAppLogs("EMAIL HAS NOT SENT TO: " + ToEmail + " from file: " + file_name + " \n Error: " + ex);
            }


            }


        }
}
