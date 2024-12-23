using System;
using System.Collections.Generic;
using System.Diagnostics.SymbolStore;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
class Program
{
    static void Main(string[] args)
    {
        // Paths to template, JSON file, and output directory
        string wordTemplate = "data/BankTemplate.docx";
        string jsonFilePath = "data/broker_data.json";
        string outputDirectory = "data/Output";

        // Ensure output directory exists
        Directory.CreateDirectory(outputDirectory);

        // Load JSON data
        var brokers = LoadBrokerJson(jsonFilePath);
        if (brokers is not null)
        {
            // Process each broker
            foreach (var broker in brokers)
            {

                GenerateDocs(wordTemplate, outputDirectory, broker);
            }
            Console.WriteLine("Docs generated successfully!");
        }

    }

    static void GenerateDocs(string wordTemplate, string outputDirectory, BrokerInformation broker)
    {

        // Create a document.   
        using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(wordTemplate, true))
        {
            string? docText = null;
            // Pdf File Name
            string newPdfFile = outputDirectory + $"/{broker.BrokerName}.pdf";

            // Create new doc by copying the template
            string newFileName = outputDirectory + $"/{broker.BrokerName}.docx";
            
            using var newDocument = (WordprocessingDocument)wordDocument.Clone(newFileName, true);

            if (newDocument.MainDocumentPart is not null && broker is not null)
            {
                using (StreamReader sr = new StreamReader(newDocument.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();

                }

                Regex regexText = new Regex("BrokerName");
                docText = regexText.Replace(docText, broker.BrokerName ?? "");

                regexText = new Regex("TINNumber");
                docText = regexText.Replace(docText, broker.TINNumber ?? "");

                regexText = new Regex("NPNNumber");
                docText = regexText.Replace(docText, broker.NPNNumber ?? "");

                regexText = new Regex("BrokerContactName");
                docText = regexText.Replace(docText, broker.BrokerContactName ?? "");

                regexText = new Regex("BrokerAddress");
                docText = regexText.Replace(docText, broker.BrokerAddress ?? "");

                regexText = new Regex("BrokerCity");
                docText = regexText.Replace(docText, broker.BrokerCity ?? "");

                regexText = new Regex("BrokerState");
                docText = regexText.Replace(docText, broker.BrokerState ?? "");

                regexText = new Regex("BrokerZip");
                docText = regexText.Replace(docText, broker.BrokerZip ?? "");


                using (StreamWriter sw = new StreamWriter(newDocument.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                    sw.Close();
                }
                newDocument.Save();
            }
            newDocument.Dispose();
            

            ConvertWordToPdf(newFileName, newPdfFile);

            //SaveAsPdf(newDocument, newPdfFile);

        }


    }

    static void FindAndReplace(string wordTemplate, string outputDirectory, User user)
    {

        // Create a document.   
        using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(wordTemplate, true))
        {
            string? docText = null;
            // Create new doc by copying the template
            string newFileName = outputDirectory + $"/{user.Name}.docx";
            using var newDocument = (WordprocessingDocument)wordDocument.Clone(newFileName, true);

            if (newDocument.MainDocumentPart is not null && user is not null)
            {
                using (StreamReader sr = new StreamReader(newDocument.MainDocumentPart.GetStream()))
                {
                    docText = sr.ReadToEnd();
                }

                Regex regexText = new Regex("{{Name}}");
                docText = regexText.Replace(docText, user.Name ?? "");
                regexText = new Regex("{{Email}}");
                docText = regexText.Replace(docText, user.Email ?? "");

                using (StreamWriter sw = new StreamWriter(newDocument.MainDocumentPart.GetStream(FileMode.Create)))
                {
                    sw.Write(docText);
                }
                newDocument.Save();

                // Create Pdf
                string newPdfFile = outputDirectory + $"/{user.Name}.pdf";
                SaveAsPdf(newDocument, newPdfFile);
            }

        }


    }
    static void SaveAsPdf(WordprocessingDocument doc, string pdfPath)
    {
        //To be impliment
    }

    static void ConvertWordToPdf(string docPath, string pdfPath)
    {
        var document = new Aspose.Words.Document(docPath);
        document.Save(pdfPath, Aspose.Words.SaveFormat.Pdf);
    }

    static List<BrokerInformation>? LoadBrokerJson(string path)
    {
        string jsonContent = File.ReadAllText(path);
        return JsonSerializer.Deserialize<List<BrokerInformation>>(jsonContent);
    }
    static List<User>? LoadJson(string path)
    {
        string jsonContent = File.ReadAllText(path);
        return JsonSerializer.Deserialize<List<User>>(jsonContent);
    }
    // User class for JSON data mapping
    public class User
    {
        public string? Name { get; set; }
        public string? Email { get; set; }
    }

    // BrokerInformation Class
    public class BrokerInformation
    {
        public string? BrokerName { get; set; }
        public string? TINNumber { get; set; }
        public string? NPNNumber { get; set; }
        public string? BrokerContactName { get; set; }
        public string? BrokerAddress { get; set; }
        public string? BrokerCity { get; set; }
        public string? BrokerState { get; set; }
        public string? BrokerZip { get; set; }
    }
}