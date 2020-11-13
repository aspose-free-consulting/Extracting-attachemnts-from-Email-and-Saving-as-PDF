using Aspose.Cells;
using Aspose.Email;
using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExtractEmailAttachment
{
    class Program
    {
        static int EmailIndex = 0;
        public static void LoadLicense(String LicensePath)
        {
            String licPath = @"C:\Aspose Data\Aspose.Total.Product.Family.lic";
            licPath = LicensePath;
            //Load Email License

            try
            {
                Aspose.Email.License licEmail = new Aspose.Email.License();
                licEmail.SetLicense(licPath);
            }
            catch(Exception e)
            {
                Console.WriteLine("Error loading licnse file for Aspose.Email...");
                Console.WriteLine("Stack Trace: " + e.Message);
            }
            //Load Words License
            try{
                Aspose.Words.License licWords = new Aspose.Words.License();
                licWords.SetLicense(licPath);
                }
            catch (Exception e)
            {
                Console.WriteLine("Error loading licnse file for Aspose.Words...");
                Console.WriteLine("Stack Trace: " + e.Message);
            }

            //Load PDF License
            try
            {
                Aspose.Pdf.License licPDF = new Aspose.Pdf.License();
                licPDF.SetLicense(licPath);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error loading licnse file for Aspose.PDF...");
                Console.WriteLine("Stack Trace: " + e.Message);
            }
            //Load Cells License
            try
            {
                Aspose.Cells.License licCells = new Aspose.Cells.License();
                licCells.SetLicense(licPath);
            }
            catch (Exception e)
            {
                Console.WriteLine("Error loading licnse file for Aspose.Cells...");
                Console.WriteLine("Stack Trace: " + e.Message);
            }

        }
        public static Stream GenerateStreamFromString(string s)
        {
            var stream = new MemoryStream();
            var writer = new StreamWriter(stream);
            writer.Write(s);
            writer.Flush();
            stream.Position = 0;
            return stream;
        }
        public static Aspose.Pdf.Document LoadEmail(FileInfo MsgFile)
        {
            try
            {
                MailMessage msg = MailMessage.Load(MsgFile.FullName);
                #region Generate Body PDF
                var HtmlBody = msg.HtmlBody;
                //Creating the Body PDF:
                MemoryStream PDFStream = new MemoryStream();
                MemoryStream stream = new MemoryStream();
                msg.Save(stream, EmlSaveOptions.DefaultMhtml);

                Document bodyContent = new Document(stream);
                bodyContent.Save(PDFStream, Aspose.Words.SaveFormat.Pdf);
                /*using (var stream = GenerateStreamFromString(msg.HtmlBody))
                {
                   Document bodyContent = new Document(stream);
                   bodyContent.Save(PDFStream, Aspose.Words.SaveFormat.Pdf);
                }*/

                #endregion
                #region Exrtracting attachements, saving as PDF and appending in main doc
                Aspose.Pdf.Document MainPdf = new Aspose.Pdf.Document(PDFStream);
                Aspose.Pdf.Document TempPdf = null;
                MemoryStream TempPdfStream = null;
                int AttachmentCount = msg.Attachments.Count;
                for (int i = 0; i < AttachmentCount; i++)
                {
                    Attachment attachment = msg.Attachments[i];
                    //Fetching data
                    Stream AttachmentStream = attachment.ContentStream;
                    TempPdfStream = new MemoryStream();
                    FileInfo AttachmentInfo = new FileInfo(attachment.Name);

                    switch (AttachmentInfo.Extension)
                    {
                        case ".docx":
                            {
                                Document doc = new Document(AttachmentStream);
                                // Save the document in PDF format.
                                doc.Save(TempPdfStream, Aspose.Words.SaveFormat.Pdf);
                                TempPdf = new Aspose.Pdf.Document(TempPdfStream);
                                MainPdf.Pages.Add(TempPdf.Pages);
                                break;
                            }
                        case ".doc":
                            {
                                Document doc = new Document(AttachmentStream);
                                // Save the document in PDF format.
                                doc.Save(TempPdfStream, Aspose.Words.SaveFormat.Pdf);
                                TempPdf = new Aspose.Pdf.Document(TempPdfStream);
                                MainPdf.Pages.Add(TempPdf.Pages);
                                break;
                            }

                        case ".pdf":
                            {
                                Aspose.Pdf.Document pdfdoc = new Aspose.Pdf.Document(AttachmentStream);
                                // Save the document in PDF format.
                                MainPdf.Pages.Add(pdfdoc.Pages);
                                break;
                            }

                        case ".xlsx":
                            {
                                Workbook wb = new Workbook(AttachmentStream);
                                // Save the document in PDF format.
                                wb.Save(TempPdfStream, Aspose.Cells.SaveFormat.Pdf);
                                TempPdf = new Aspose.Pdf.Document(TempPdfStream);
                                MainPdf.Pages.Add(TempPdf.Pages);
                                break;
                            }
                        default:
                            {
                                break;
                            }
                    }
                }
                #endregion
                #region Save PDF
                using (MemoryStream EmailPDF = new MemoryStream())
                {
                    MainPdf.Save(EmailPDF, Aspose.Pdf.SaveFormat.Pdf);
                }
                #endregion
                return MainPdf;
            }
           catch (Exception e)
            {
                Console.WriteLine("Error extracting the information from MSG file..");
                Console.WriteLine("Stack Trace: " + e.Message);
                return null;
            }
        }
        static void Main(string[] args)
        {
            String MailFolderPath = @"C:\Users\mudas\Desktop\TestAttachment\MSG\";
            String MsgPath = "";
            String TargetPDFName = "OutputPDF.pdf";
            String LicensePath = "";
            if (args.Length < 3)
            {
                Console.WriteLine("Please enter a arguments for source files path, License File and genrated PDF Name");
                Environment.Exit(0);
            }
            else
            {
                MsgPath = args[0].ToString();
                LicensePath = args[1].ToString();
                TargetPDFName = args[2].ToString();

            }

            //Load License files
            LoadLicense(LicensePath);


            MailFolderPath = MsgPath;
            //Finding MSG files
            string[] filePaths = Directory.GetFiles(MailFolderPath, "*.msg");

            if (filePaths.Count() > 0)
            {
                //Readind individual files and generating PDF
                Aspose.Pdf.Document FinalPDF = null;
                Aspose.Pdf.Document TempFinalPDF = null;

                int iCounter = 1;
                for (int index = 0; index < filePaths.Count(); index++)
                {
                    FileInfo MsgFile = new FileInfo(filePaths[index]);
                    TempFinalPDF = LoadEmail(MsgFile);
                    if (TempFinalPDF != null)
                    {
                        if (iCounter <= 1)
                        {
                            FinalPDF = TempFinalPDF;
                            iCounter = 1;
                        }
                        else
                        {
                            FinalPDF.Pages.Add(TempFinalPDF.Pages);
                        }
                        iCounter++;
                    }
                    else
                    {
                        iCounter--;
                    }
                }
                //Generating Final PDF
                FinalPDF.Save(MailFolderPath + @"\" + TargetPDFName, Aspose.Pdf.SaveFormat.Pdf);
            }
            else
            {
                Console.WriteLine("No Message files found on specified path. Please check the path again..");
            }
        }
    }
}
