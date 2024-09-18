# Extracting Email Content and attachemnts from MSG files and Saving as single PDF

This [free consulting project](https://aspose-free-consulting.github.io/) uses [Aspose.Email](https://products.aspose.com/email), [Aspose.Cells](https://products.aspose.com/cells) and [Aspose.Words](https://products.aspose.com/words) to extrat diferent attachments (Excel, Word, PDF and images) from email and saving them as single PDF file using simple API calls.

In this sample project, Aspose.Email API has been used to process one or multiple MSG files having attachments or no attachments. The API shall process each MSG file and perform following operation on each MSG file:

Extract Body Headers and Body 
Extract Attachment files from MSG
Save the extracted information as PDF

The extracted attachments are then processed and saved to PDF using Aspose.Words in case of Word file attachemnt, Aspose.Cells in case of Excel file attachment and Aspose.PDF in case of PDF file attachments.

Finally, Aspose.PDF has been used to save the extracted PDF and also merging all different types of attachment PDF files to a single PDF


# Guidelines for Usage

**Install latest API version from Nuget:** In order to use the sample project, you need to install the latest availble version of [Aspose.Email for .NET from NUGET](https://www.nuget.org/packages/Aspose.Email/), [Aspose.Cells for .NET from NUGET](https://www.nuget.org/packages/Aspose.Cells/) and [Aspose.Words for .NET from NUGET](https://www.nuget.org/packages/Aspose.Words/).

**Set License File:** 
It's advised that you must have a valid license file before using the application. In case you do not have one, you can still use the trial version of API (It limits the API usage features) but suggested approach is to use the API with license. You can buy our product license or may request a [Free trial license](https://purchase.aspose.com/temporary-license) to evaluate the API with complete features on your end.

**Compiling the Sample project**
You need to compile the project by resolving all dependencies. In the end the **ExtractEmailAttachment.exe** will be gnerated.

**Executing the project**
This project has been configured to execute by providing command line arguments on execution time. You will be using following Power shell command to execute the exe. Obviously, first of all you will set path to your exe file in Power shell.

**.\ExtractEmailAttachment.exe "C:\Users\mudas\Desktop\TestAttachment\MSG" "C:\Aspose Data\Aspose.Total.Product.Family.lic" "TestOutput.pdf"**

There are 3 arguments used:

1: "C:\Users\mudas\Desktop\TestAttachment\MSG" : It is path to MSG files directory

2: "C:\Aspose Data\Aspose.Total.Product.Family.lic": It is License file name including its path. If license file is place along side the exe file, you will only be adding the license file name here.

3: "TestOutput.pdf": Desired output PDF file name


![GitHub Logo](https://user-images.githubusercontent.com/3595481/99031834-111d3f80-2546-11eb-8db5-5cfc136297de.png)

## Interested in Aspose free consulting project?
[If you are also interested in a free consulting project by Aspose team then please view details on this page](https://aspose-free-consulting.github.io/)

If you have any questions about Aspose APIs, please feel free to [post your query in Aspose file format APIs Forums](https://forum.aspose.com/). 

Also, you can keep in touch with the latest developments in [file format APIs offered by Aspose at our Blog](https://blog.aspose.com/).

## This free consluting project is based on the following issue:

I want to create a console app that extract attachments from email (msg) and convert/merge into one PDF: github.com/aspose-free-consulting/projects/issues/49
