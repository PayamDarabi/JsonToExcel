using ClosedXML.Excel;
using JsonFromOrToExcel;
using JsonFromOrToExcel.Objects;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace JsonToExcel
{
    class Program
    {
        //Change Files Path To Your Directory Of Json Files

        static string filesFormat = "*.json";
        static string filesPath = @"<YOUR DRIVE>:\<DIRECTORY IN THIS FORMAT: C:\AA\BB\CC\";
        static string sheetName = "Sheet";
        static void Main(string[] args)
        {

            Console.WriteLine("Starting");
            if(filesPath== @"<YOUR DRIVE>:\<DIRECTORY IN THIS FORMAT: C:\AA\BB\CC\")
            {
                Console.WriteLine("Your filesPath Not Correct... Please Put Your Root Path In --filesPath-- Variable");
                return;
            }
            Console.WriteLine("FilesPath: {0}", filesPath);
            Console.WriteLine("Files Format: {0}", filesFormat);


            string[] filePaths = Directory.GetFiles(filesPath, filesFormat);
            foreach (var file in filePaths)
            {
                // Get File Name And Set It To SheetName
                sheetName = Path.GetFileNameWithoutExtension(file);
                Console.WriteLine("Reading File: {0}", sheetName);

                using (StreamReader r = new StreamReader(file))
                {
                    string jsonContent = r.ReadToEnd();

                    //JsonObject: Root Object Of Json
                    var jsonObj = JsonConvert.DeserializeObject<JsonObject>(jsonContent);

                    List<Vocab> vocabsList = ToList(jsonObj.Vocabs);

                    var fileName = file.Split('\\').Last().Split('.')[0].Trim().ToString();

                    JsonToExcel(vocabsList, fileName);

                }
            }

        }
        public static void JsonToExcel(List<Vocab> vocabs, string filename)
        {
            try
            {
                XLWorkbook wb = new XLWorkbook();

                wb.Worksheets.Add(sheetName);

                if (filename.Contains("."))
                {
                    int IndexOfLastFullStop = filename.LastIndexOf('.');

                    filename = filename.Substring(0, IndexOfLastFullStop) + ".xlsx";

                }
                int row = 1;
                foreach (Vocab vocab in vocabs)
                {
                    var ws = wb.Worksheet(sheetName);

                    ws.Cell("A" + row.ToString()).Value = vocab.English.ToString();
                    ws.Cell("B" + row.ToString()).Value = vocab.Persian.ToString();
                    row++;
                }
                filename = filename + ".xlsx";

                wb.SaveAs(filename);
                Console.WriteLine("Excel file saved SuccessFully In Path: {0}", filename);
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                + ex.Message);
            }

        }
        public static List<Vocab> ToList(List<Vocab> vocabsList)
        {
            var vocabs = new List<Vocab>();

            foreach (var vocab in vocabsList)
            {
                vocabs.Add(new Vocab { English = vocab.English.Trim(), Persian = vocab.Persian.Trim() });
            }
            return vocabs;
        }
    }
}
