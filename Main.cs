internal class Program
{
    private static void Main(string[] args)
    {
        string source = null;
        string destination = null;

        // Check args and assign source and destination
        if (args.Length > 0)
        {
            source = args[0];
        }
        if (args.Length > 1)
        {
            destination = args[1];
        }

        // If source is not provided, check the current directory for a single XLSX file
        if (source == null)
        {
            var currentDirectory = Directory.GetCurrentDirectory();
            var xlsxFiles = Directory.GetFiles(currentDirectory, "*.xlsx");

            if (xlsxFiles.Length == 1)
            {
                source = xlsxFiles[0];
            }
            else
            {
                Console.WriteLine("No source file specified, and no single XLSX file found in the current directory.");
                return;
            }
        }

        // If destination is not provided, set it to "output.docx" in the current directory
        if (destination == null)
        {
            destination = Path.Combine(Directory.GetCurrentDirectory(), "output.docx");
        }

        // Check if the source file exists
        if (!File.Exists(source))
        {
            Console.WriteLine($"The source file '{source}' does not exist.");
            return;
        }

        // Process lesson plans and create the document
        try
        {
            var lg = new Generator();
            var plans = Loader.GetLessonPlansFromXlsx(source);
            lg.CreateDocument(destination, plans);
            Console.WriteLine($"Created {plans.Count} lessons in '{destination}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }



//         LessonDocumentGenerator lg = new LessonDocumentGenerator();
//         //var lgdata = new lpdemo();

//         // 2024-03-26 creating new lesson plan
//         // string fileName = @"c:\temp\demooutput.docx";
//         // string fileName2 = @"C:\Users\seaml\OneDrive\learning\ED5966_ED5967\ED5966_ED5967_AT3\source.xlsx";

//         string fileName = @"c:\temp\demooutput.docx";
//         string fileName2 = @"C:\Users\seaml\OneDrive\learning\ED5976\AT3\source.xlsx";
//         //fileName2 = @"C:\Users\seaml\OneDrive\learning\ED5976\AT3\source2.xlsx";
//         fileName2 = @"C:\Users\seaml\OneDrive\learning\profex\placement03\source.xlsx";
//         fileName2 = @"C:\Users\seaml\OneDrive\learning\profex\placement03\dataset.xlsx";

//         if(!File.Exists(fileName2))
//         {
//             Console.WriteLine($"File '{fileName2}' does not exist.");
//             return;
//         }


//         var lessonplans = lpdemo.GetLessonPlansFromXlsx(fileName2);
// // C:\Users\seaml\OneDrive\learning\ED5956_ED5957\ED5956_ED5957_A3
//         lg.CreateDocument(fileName, lessonplans);
//         Console.WriteLine($"Created {lessonplans.Count} lessons in '{fileName}'.");
//         //System.Diagnostics.Process.Start("winword.exe", fileName);
//

}