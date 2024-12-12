using System;

namespace WorkCenterDemandExtract
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length < 6)
            {
                throw new Exception("Not enough arguments supplied.");
            }
            WorkCenterDemandExtract workCenterDemandExtract = new WorkCenterDemandExtract(args[0], args[1], args[2], args[3], args[4]);
            workCenterDemandExtract.GenerateDataTable();

            if (args[5] == "EXCEL")
            {
                workCenterDemandExtract.ExportToExcel();
            }
            else
            {
                workCenterDemandExtract.ExportToCsv();
            }
            workCenterDemandExtract.session.client.CloseSession();
        }
    }
}
