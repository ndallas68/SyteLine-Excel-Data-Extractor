using System;

namespace ExtractAndWriteData
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 8)
            {
                throw new Exception("Not enough arguments supplied.");
            }
            ExtractAndWriteData extractAndWriteData;
            if (args.Length == 8)
            {
                extractAndWriteData = new ExtractAndWriteData(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7]);
            }
            else if (args.Length == 9)
            {
                extractAndWriteData = new ExtractAndWriteData(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8]);
            }
            else
            {
                extractAndWriteData = new ExtractAndWriteData(args[0], args[1], args[2], args[3], args[4], args[5], args[6], args[7], args[8], args[9]);
            }

            extractAndWriteData.ExportToCsv();
            extractAndWriteData.session.client.CloseSession();
        }
    }
}
