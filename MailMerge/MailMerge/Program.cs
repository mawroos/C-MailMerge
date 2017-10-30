using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
namespace WordAuto
{
    class MailMerge
    {
        static void Main(string[] args)
        {
            //OBJECT OF MISSING "NULL VALUE"

            Object oMissing = System.Reflection.Missing.Value;
            Console.WriteLine("Enter Template Path: ");
            string templatePath = Console.ReadLine();
            Object oTemplatePath = templatePath;
            string repeat;
            do
            {
                repeat = string.Empty;
                Application wordApp = new Application();
                Document wordDoc = new Document();
                wordDoc = MergeMail(ref oMissing, ref oTemplatePath, wordApp);
                Console.WriteLine($"{wordDoc.Name} is done, do you need to do another Mail Merge? type Yes");
                wordApp.Application.Quit();
                repeat = Console.ReadLine();
            } while (repeat == "Yes");
            Console.WriteLine("Thanks, we are done:-), press Enter to Exit");
            Console.ReadLine()
        }

        private static Document MergeMail(ref object oMissing, ref object oTemplatePath, Application wordApp)
        {
            Document wordDoc = wordApp.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            foreach (Field myMergeField in wordDoc.Fields)
            {


                Range rngFieldCode = myMergeField.Code;

                String fieldText = rngFieldCode.Text;



                // ONLY GETTING THE MAILMERGE FIELDS

                if (fieldText.StartsWith(" MERGEFIELD"))
                {

                    // THE TEXT COMES IN THE FORMAT OF

                    // MERGEFIELD  MyFieldName  \\* MERGEFORMAT

                    // THIS HAS TO BE EDITED TO GET ONLY THE FIELDNAME "MyFieldName"

                    Int32 endMerge = fieldText.IndexOf("\\");

                    Int32 fieldNameLength = fieldText.Length - endMerge;

                    String fieldName = fieldText.Substring(11, fieldNameLength - 12);

                    // GIVES THE FIELDNAMES AS THE USER HAD ENTERED IN .dot FILE

                    fieldName = fieldName.Trim();

                    // **** FIELD REPLACEMENT IMPLEMENTATION GOES HERE ****//


                    Console.WriteLine($"Enter the Value for Field {fieldName}:");
                    string fieldValue = Console.ReadLine();

                    myMergeField.Select();

                    wordApp.Selection.TypeText(fieldValue);


                }

            }
            Console.WriteLine("ALL fields merged, please input the generated file path/name:");
            string generatedFilenName = Console.ReadLine();
            wordDoc.SaveAs(generatedFilenName+".docx");
            Console.WriteLine("Do you want to generate it as PDF as well, type Yes");
            string pdfAnswer = Console.ReadLine();
            if (pdfAnswer == "Yes") wordDoc.SaveAs2(generatedFilenName+".pdf", WdSaveFormat.wdFormatPDF);
            return wordDoc;
        }
    }
}