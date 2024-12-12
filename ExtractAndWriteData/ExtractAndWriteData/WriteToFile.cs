using System;
using System.IO;

namespace ExtractAndWriteData
{
    internal class WriteToFile
    {
        private String fileName;

        public WriteToFile(string fileName)
        {
            this.fileName = fileName;
        }

        public void CreateFile(String[] text)
        {
            File.WriteAllLines(fileName, text);
        }
    }
}