using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Timers; // Explicitly specify System.Timers
using ADODB;

class Program
{
    private static int totalRecords = 0;
    private static int processedRecords = 0;
    private static Timer? progressTimer;
    private static string filename = string.Empty;
    private const int MaxRecordsPerRun = 1000000;

    static void Main(string[] args)
    {
        if (args.Length < 2 || args.Length > 3)
        {
            Console.WriteLine("Usage: SASToCSV <path-to-SAS-dataset> <output-CSV-file> [<start-record-number>]");
            return;
        }

        string fileToProcess = args[0].Trim('\'', '"');
        string outputCsvFile = args[1].Trim('\'', '"');
        int startRecord = args.Length == 3 ? int.Parse(args[2]) : 0;

        if (!File.Exists(fileToProcess))
        {
            Console.WriteLine($"{fileToProcess} does not exist.");
            return;
        }

        string filePath = Path.GetDirectoryName(fileToProcess) ?? ".";
        filename = Path.GetFileNameWithoutExtension(fileToProcess);

        const CursorTypeEnum adOpenDynamic = CursorTypeEnum.adOpenDynamic;
        const LockTypeEnum adLockOptimistic = LockTypeEnum.adLockOptimistic;
        const int adCmdTableDirect = 512;

        Connection objConnection = new Connection();
        Recordset objRecordset = new Recordset();

        try
        {
            objConnection.Open($"Provider=SAS.LocalProvider;Data Source=\"{filePath}\";");
            objRecordset.ActiveConnection = objConnection;
            objRecordset.Properties["SAS Formats"].Value = "_ALL_";

            objRecordset.Open(filename, Type.Missing, adOpenDynamic, adLockOptimistic, adCmdTableDirect);
            objRecordset.MoveFirst();

            totalRecords = objRecordset.RecordCount;
            processedRecords = 0;

            // Move to the start record if specified
            if (startRecord > 0)
            {
                objRecordset.Move(startRecord);
            }

            // Set up a timer to report progress every 10 seconds
            progressTimer = new Timer(10000);
            progressTimer.Elapsed += ReportProgress;
            progressTimer.Start();

            using (StreamWriter writer = new StreamWriter(outputCsvFile, startRecord > 0, Encoding.UTF8))
            {
                // Write the header if starting from the first record
                if (startRecord == 0)
                {
                    for (int i = 0; i < objRecordset.Fields.Count; i++)
                    {
                        writer.Write($"\"{objRecordset.Fields[i].Name}\"");
                        if (i < objRecordset.Fields.Count - 1)
                            writer.Write(",");
                    }
                    writer.WriteLine();
                }

                // Write the data
                while (!objRecordset.EOF && processedRecords < MaxRecordsPerRun)
                {
                    for (int i = 0; i < objRecordset.Fields.Count; i++)
                    {
                        writer.Write($"\"{objRecordset.Fields[i].Value}\"");
                        if (i < objRecordset.Fields.Count - 1)
                            writer.Write(",");
                    }
                    writer.WriteLine();
                    objRecordset.MoveNext();
                    processedRecords++;
                }
            }

            progressTimer.Stop();

            if (processedRecords >= MaxRecordsPerRun)
            {
                // Start a new instance of the program to process the next batch
                Process.Start("SASToCSV.exe", $"\"{fileToProcess}\" \"{outputCsvFile}\" {startRecord + processedRecords}");
                Process.GetCurrentProcess().Kill();
            }
            else
            {
                Console.WriteLine($"Data successfully written to {outputCsvFile}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unable to process {fileToProcess}");
            Console.WriteLine(ex.Message);
        }
        finally
        {
            if (objRecordset.State != 0)
                objRecordset.Close();
            if (objConnection.State != 0)
                objConnection.Close();
        }
    }

    private static void ReportProgress(object? sender, ElapsedEventArgs e) // Allow sender to be nullable
    {
        if (totalRecords > 0)
        {
            double progress = (double)processedRecords / totalRecords * 100;
            Console.WriteLine($"{filename}: {progress:F2}%");
        }
    }
}
