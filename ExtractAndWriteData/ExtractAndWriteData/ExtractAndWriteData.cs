using Mongoose.IDO.Protocol;
using System;

namespace ExtractAndWriteData
{
    internal class ExtractAndWriteData
    {
        public string fileName;
        public string idoName;
        public string properties;
        public string recordCap;
        public string filter;
        public string orderBy;
        public Session session;

        public ExtractAndWriteData(string url, string username, string password, string configuration, string fileName, string idoName, string properties, string recordCap)
        {
            this.idoName = idoName;
            this.properties = properties;
            this.recordCap = recordCap;
            this.fileName = fileName;
            filter = string.Empty;
            orderBy = string.Empty;
            session = new Session();
            session.Login(url, username, password, configuration);
        }
        public ExtractAndWriteData(string url, string username, string password, string configuration, string fileName, string idoName, string properties, string recordCap, string filter)
        {
            this.idoName = idoName;
            this.properties = properties;
            this.recordCap = recordCap;
            this.filter = filter;
            this.fileName = fileName;
            orderBy = string.Empty;
            session = new Session();
            session.Login(url, username, password, configuration);
        }
        public ExtractAndWriteData(string url, string username, string password, string configuration, string fileName, string idoName, string properties, string recordCap, string filter, string orderBy)
        {
            this.idoName = idoName;
            this.properties = properties;
            this.recordCap = recordCap;
            this.filter = filter;
            this.orderBy = orderBy;
            this.fileName = fileName;
            session = new Session();
            session.Login(url, username, password, configuration);
        }
        public PropertyList createPropertyList()
        {
            string list = string.Empty;
            if (properties == "*")
            {
                GetPropertyInfoResponseData propertyResponse = default;
                propertyResponse = session.client.GetPropertyInfo(idoName);
                bool first = true;
                foreach (PropertyInfo property in propertyResponse.Properties)
                {
                    if (property.SubcollectionProgID != string.Empty)
                    {
                        continue;
                    }
                    if (first)
                    {
                        list = property.Name;
                        first = false;
                    }
                    else
                    {
                        list += "," + property.Name;
                    }
                }
                return new PropertyList(list);
            }
            else
            {
                return new PropertyList(properties);
            }
        }

        public void ExportToCsv()
        {
            LoadCollectionResponseData response = session.client.LoadCollection(idoName, properties, filter, orderBy, int.Parse(recordCap));

            String[] textForFile = new string[response.Items.Count + 1];

            PropertyList propertyList = createPropertyList();

            textForFile[0] = propertyList.ToString();

            WriteToFile writeToFile = new WriteToFile(fileName + " " + DateTime.Now.ToString("yyyy-MM-dd") + ".csv");

            String concatLineData = String.Empty;
            int i = 1;

            foreach (IDOItem item in response.Items)
            {
                bool first = true;
                foreach (IDOPropertyValue property in item.PropertyValues)
                {
                    if (first)
                    {
                        concatLineData = property.Value.Replace(",", " ");
                        first = false;
                    }
                    else
                    {
                        concatLineData += "," + property.Value.Replace(",", " ");
                    }
                }
                textForFile[i] = concatLineData;
                i++;
            }
            writeToFile.CreateFile(textForFile);
        }
    }
}
