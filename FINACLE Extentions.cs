
using System;
using System.IO;
using System.Xml;
using System.Text;
using System.Collections.Specialized;
using Microsoft.MetadirectoryServices;
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Net;
using System.Data;

namespace FimSync_Ezma
{
    public class EzmaExtension :
    IMAExtensible2CallExport,
    IMAExtensible2CallImport,
    //IMAExtensible2FileImport,
    //IMAExtensible2FileExport,
    //IMAExtensible2GetHierarchy,
    IMAExtensible2GetSchema,
    IMAExtensible2GetCapabilities,
    IMAExtensible2GetParameters
    //IMAExtensible2GetPartitions
    {
        private int m_importDefaultPageSize = 1200;
        private int m_importMaxPageSize = 1500;
        private int m_exportDefaultPageSize = 10;
        private int m_exportMaxPageSize = 20;

        public string USERNAME;
        public string PASSWORD;
        public string EMPLOYEE_ID;
        public string FIRST_NAME;
        public string LAST_NAME;
        public string EMAIL;
        public string EMPLOYEE_STATUS;
        public string ROLE_ASSIGNMENT;
        public string WORKPHONE;

        public string webServicePassword;
        public string webServiceUsername;
        public string imUrl;
        public string exUrl;
        public string enableLogging;
        public string loggingLevel;
        //
        // Constructor
        //
        public EzmaExtension()
        {
            //
            // TODO: Add constructor logic here
            //
        }

        public MACapabilities Capabilities
        {
            get
            {
                MACapabilities myCapabilities = new MACapabilities();

                myCapabilities.ConcurrentOperation = true;
                myCapabilities.ObjectRename = false;
                myCapabilities.DeleteAddAsReplace = true;
                myCapabilities.DeltaImport = true;
                myCapabilities.DistinguishedNameStyle = MADistinguishedNameStyle.None;
                myCapabilities.ExportType = MAExportType.AttributeUpdate;
                myCapabilities.NoReferenceValuesInFirstExport = false;
                myCapabilities.Normalizations = MANormalizations.None;

                return myCapabilities;
            }
        }

        public int ImportDefaultPageSize
        {
            get
            {
                return m_importDefaultPageSize;
            }
        }

        public int ImportMaxPageSize
        {
            get
            {
                return m_importMaxPageSize;
            }
        }

        public int ExportDefaultPageSize
        {
            get
            {
                return m_exportDefaultPageSize;
            }

            set
            {
                m_exportDefaultPageSize = value;
            }
        }

        public int ExportMaxPageSize
        {
            get
            {
                return m_exportMaxPageSize;
            }
            set
            {
                m_exportMaxPageSize = value;
            }
        }

        public IList<ConfigParameterDefinition> GetConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            List<ConfigParameterDefinition> configParametersDefinitions = new List<ConfigParameterDefinition>();
            switch (page)
            {
                case ConfigParameterPage.Connectivity:
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("imUrl", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("exUrl", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("webServiceUsername", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("webServicePassword", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("enableLogging", ""));
                    configParametersDefinitions.Add(ConfigParameterDefinition.CreateStringParameter("loggingLevel", ""));
                    break;
                    //case ConfigParameterPage.Global:
                    //    break;
                    //case ConfigParameterPage.Partition:
                    //    break;
                    //case ConfigParameterPage.RunStep:
                    //    break;
            }
            return configParametersDefinitions;
        }

        public CloseImportConnectionResults CloseImportConnection(CloseImportConnectionRunStep importRunStep)
        {
            return new CloseImportConnectionResults();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="importRunStep"></param>
        /// <returns></returns>
        public GetImportEntriesResults GetImportEntries(GetImportEntriesRunStep importRunStep)
        {
            HttpWebRequest request = WebRequest.Create(this.imUrl) as HttpWebRequest;
            // Get response
            //string URLToSend = LiteralURLTrackerAddress.Text.ToString();
            //HttpWebRequest request = (HttpWebRequest)WebRequest.Create(URLToSend.ToString()); request.Method = "GET"; request.KeepAlive = true;
            using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
            {
                List<CSEntryChange> csentries = new List<CSEntryChange>();
                GetImportEntriesResults importReturnInfo;
                StreamReader responseStream = new StreamReader(response.GetResponseStream());
                string webResponseStream = responseStream.ReadToEnd();
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.LoadXml(webResponseStream);
                //load data into dataset
                DataSet da = new DataSet();
                da.ReadXml(response.GetResponseStream());
                XmlNodeList nodeList = xmlDoc.DocumentElement.SelectNodes("/Table/Product");
                   foreach (XmlNode node in nodeList)
                {                    
                    USERNAME= node.SelectSingleNode("USERNAME").InnerText;
                    PASSWORD= node.SelectSingleNode("PASSWORD").InnerText;
                    EMPLOYEE_ID= node.SelectSingleNode("EMPLOYEE_ID").InnerText;
                    FIRST_NAME= node.SelectSingleNode("FIRST_NAME").InnerText;
                    LAST_NAME= node.SelectSingleNode("LAST_NAME").InnerText;
                    EMAIL= node.SelectSingleNode("EMAIL").InnerText;
                    EMPLOYEE_STATUS= node.SelectSingleNode("EMPLOYEE_STATUS").InnerText;
                    ROLE_ASSIGNMENT= node.SelectSingleNode("ROLE_ASSIGNMENT").InnerText;
                    WORKPHONE= node.SelectSingleNode("WORKPHONE").InnerText;

                    
                    CSEntryChange csentry1 = CSEntryChange.Create();

                    csentry1.ObjectModificationType = ObjectModificationType.Add;
                    csentry1.ObjectType = "Person";

                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("USERNAME", USERNAME));
                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("PASSWORD", PASSWORD));
                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("EMPLOYEE_ID", EMPLOYEE_ID));
                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("FIRST_NAME", FIRST_NAME));
                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("LAST_NAME", LAST_NAME));
                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("EMAIL", EMAIL));
                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("EMPLOYEE_STATUS", EMPLOYEE_STATUS));
                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("ROLE_ASSIGNMENT", ROLE_ASSIGNMENT));
                    csentry1.AttributeChanges.Add(AttributeChange.CreateAttributeAdd("WORKPHONE", WORKPHONE));

                    csentries.Add(csentry1);
                }

                importReturnInfo = new GetImportEntriesResults();
                importReturnInfo.MoreToImport = false;
                importReturnInfo.CSEntries = csentries;
                return importReturnInfo;
            }
        }
        public Schema GetSchema(KeyedCollection<string, ConfigParameter> configParameters)
        {
            Microsoft.MetadirectoryServices.SchemaType personType = Microsoft.MetadirectoryServices.SchemaType.Create("Person", false);

            //myname = configParameters["myname"].Value;

            string myattribute1 = "USERNAME";
            string myattribute2 = "PASSWORD";
            string myattribute3 = "EMPLOYEE_ID";
            string myattribute4 = "FIRST_NAME";
            string myattribute5 = "LAST_NAME";
            string myattribute6 = "EMAIL";
            string myattribute7 = "EMPLOYEE_STATUS";
            string myattribute8 = "ROLE_ASSIGNMENT";
            string myattribute9 = "WORKPHONE";

            if (myattribute1 == "USERNAME")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute1, AttributeType.String));
            }

            if (myattribute2 == "PASSWORD")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute2, AttributeType.String));
            }

            if (myattribute3 == "EMPLOYEE_ID")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute3, AttributeType.String));
            }

            if (myattribute4 == "FIRST_NAME")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute4, AttributeType.String));
            }

            if (myattribute5 == "LAST_NAME")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute5, AttributeType.String));
            }

            if (myattribute6 == "EMAIL")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute6, AttributeType.String));
            }

            if (myattribute7 == "EMPLOYEE_STATUS")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute7, AttributeType.String));
            }

            if (myattribute8 == "ROLE_ASSIGNMENT")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute8, AttributeType.String));
            }

            if (myattribute9 == "WORKPHONE")
            {
                personType.Attributes.Add(SchemaAttribute.CreateAnchorAttribute(myattribute9, AttributeType.String));
            }

            Schema schema = Schema.Create();
            schema.Types.Add(personType);

            return schema;
        }

        public OpenImportConnectionResults OpenImportConnection(KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenImportConnectionRunStep importRunStep)
        {
            this.imUrl = configParameters["imUrl"].Value;
            this.webServiceUsername = configParameters["webServiceUsername"].Value;
            this.webServicePassword = configParameters["webServicePassword"].Value;
            this.enableLogging = configParameters["enableLogging"].Value;
            this.loggingLevel = configParameters["loggingLevel"].Value;
            return new OpenImportConnectionResults();
        }



        //Export Region
        public void OpenExportConnection(KeyedCollection<string, ConfigParameter> configParameters, Schema types, OpenExportConnectionRunStep exportRunStep)
        {
            this.exUrl = configParameters["exUrl"].Value;
            this.webServiceUsername = configParameters["webServiceUsername"].Value;
            this.webServicePassword = configParameters["webServicePassword"].Value;
            this.enableLogging = configParameters["enableLogging"].Value;
            this.loggingLevel = configParameters["loggingLevel"].Value;
        }


        public ParameterValidationResult ValidateConfigParameters(KeyedCollection<string, ConfigParameter> configParameters, ConfigParameterPage page)
        {
            ParameterValidationResult myResults = new ParameterValidationResult();
            return myResults;
        }

        public PutExportEntriesResults PutExportEntries(IList<CSEntryChange> csentries)
        {
            string exportwebservicepath;
            int i = 0;
            foreach (CSEntryChange csentryChange in csentries)
            {
                EMPLOYEE_ID = csentries[i].DN.ToString();
                if (csentryChange.ObjectType == "Person")
                {
                    #region Add
                    if (csentryChange.ObjectModificationType == ObjectModificationType.Add)
                    {
                        #region a
                        foreach (string attrib in csentryChange.ChangedAttributeNames)
                        {
                            switch (attrib)
                            {                               
                                case "USERNAME":
                                    USERNAME = csentryChange.AttributeChanges["USERNAME"].ValueChanges[0].Value.ToString();
                                    break;
                                case "PASSWORD":
                                    PASSWORD = csentryChange.AttributeChanges["PASSWORD"].ValueChanges[0].Value.ToString();
                                    break;
                                case "LAST_NAME":
                                    LAST_NAME = csentryChange.AttributeChanges["LAST_NAME"].ValueChanges[0].Value.ToString();
                                    break;

                                case "FIRST_NAME":
                                    FIRST_NAME = csentryChange.AttributeChanges["FIRST_NAME"].ValueChanges[0].Value.ToString();
                                    break;

                                case "EMAIL":
                                    EMAIL = csentryChange.AttributeChanges["EMAIL"].ValueChanges[0].Value.ToString();
                                    break;

                                case "EMPLOYEE_STATUS":
                                    EMPLOYEE_STATUS = csentryChange.AttributeChanges["EMPLOYEE_STATUS"].ValueChanges[0].Value.ToString();
                                    break;

                                case "ROLE_ASSIGNMENT":
                                    ROLE_ASSIGNMENT = csentryChange.AttributeChanges["ROLE_ASSIGNMENT"].ValueChanges[0].Value.ToString();
                                    break;

                                case "WORKPHONE":
                                    WORKPHONE = csentryChange.AttributeChanges["WORKPHONE"].ValueChanges[0].Value.ToString();
                                    break;
                            }
                        }
                        #endregion

                        //call the webservice to update
                        exportwebservicepath = this.exUrl; //+ "?employeeid=" + myEmpID + "&email=" + attribValue + "&accountName=" + attribValue2;
                        HttpWebRequest request = WebRequest.Create(exportwebservicepath) as HttpWebRequest;
                        HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                    }

                    #endregion
                    #region Delete
                    if (csentryChange.ObjectModificationType == ObjectModificationType.Delete)
                    {

                        //call the webservice to update
                        exportwebservicepath = this.exUrl; //+ "?employeeid=" + myEmpID + "&email=" + attribValue + "&accountName=" + attribValue2;
                        HttpWebRequest request = WebRequest.Create(exportwebservicepath) as HttpWebRequest;
                        HttpWebResponse response = request.GetResponse() as HttpWebResponse;


                    }
                    #endregion

                    #region Update
                    if (csentryChange.ObjectModificationType == ObjectModificationType.Update)
                    {

                        foreach (string attribName in csentryChange.ChangedAttributeNames)
                        {


                            if (csentryChange.AttributeChanges[attribName].ModificationType == AttributeModificationType.Add)
                            {
                                EMPLOYEE_ID = csentryChange.AnchorAttributes[0].Value.ToString();
                                string attribValue = csentryChange.AttributeChanges[attribName].ValueChanges[0].Value.ToString();
                                //string cmdText = "Update" + myTable + "SET" + attribName + " = '" + attribValue + "' Where EmployeeID = '" + myEmpID + "'";
                                //cmd.CommandText = cmdText;
                                //cmd.Connection = conn;
                                //cmd.ExecuteNonQuery();


                                //call the webservice to update
                                exportwebservicepath = this.exUrl; //+ "?employeeid=" + myEmpID + "&email=" + attribValue + "&accountName=" + attribValue2;
                                HttpWebRequest request = WebRequest.Create(exportwebservicepath) as HttpWebRequest;
                                HttpWebResponse response = request.GetResponse() as HttpWebResponse;

                            }
                            else if (csentryChange.AttributeChanges[attribName].ModificationType == AttributeModificationType.Delete)
                            {

                                EMPLOYEE_ID = csentryChange.AnchorAttributes[0].Value.ToString();
                                //string cmdText = "Update " + myTable + " SET " + attribName + " = 'NULL' Where EmployeeID = '" + myEmpID + "'";
                                //cmd.CommandText = cmdText;
                                //cmd.Connection = conn;
                                //cmd.ExecuteNonQuery();

                                //call the webservice to update
                                exportwebservicepath = this.exUrl; //+ "?employeeid=" + myEmpID + "&email=" + attribValue + "&accountName=" + attribValue2;
                                HttpWebRequest request = WebRequest.Create(exportwebservicepath) as HttpWebRequest;
                                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                            }
                            else if (csentryChange.AttributeChanges[attribName].ModificationType == AttributeModificationType.Replace)
                            {
                                //EMPLOYEE_ID = csentryChange.AnchorAttributes[0].Value.ToString();
                                //string attribValue = csentryChange.AttributeChanges[attribName].ValueChanges[0].Value.ToString();
                                //string cmdText = "Update " + myTable + " SET " + attribName + " = '" + attribValue + "' Where EmployeeID = '" + myEmpID + "'";
                                //cmd.CommandText = cmdText;
                                //cmd.Connection = conn;
                                //cmd.ExecuteNonQuery();

                                //call the webservice to update
                                exportwebservicepath = this.exUrl; //+ "?employeeid=" + myEmpID + "&email=" + attribValue + "&accountName=" + attribValue2;
                                HttpWebRequest request = WebRequest.Create(exportwebservicepath) as HttpWebRequest;
                                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                            }
                            else if (csentryChange.AttributeChanges[attribName].ModificationType == AttributeModificationType.Update)
                            {
                                //EMPLOYEE_ID = csentryChange.AnchorAttributes[0].Value.ToString();
                                //string attribValue = csentryChange.AttributeChanges[attribName].ValueChanges[0].Value.ToString();
                                //string cmdText = "Update " + myTable + " SET " + attribName + " = '" + attribValue + "' Where EmployeeID = '" + myEmpID + "'";
                                //cmd.CommandText = cmdText;
                                //cmd.Connection = conn;
                                //cmd.ExecuteNonQuery();

                                //call the webservice to update
                                exportwebservicepath = this.exUrl; //+ "?employeeid=" + myEmpID + "&email=" + attribValue + "&accountName=" + attribValue2;
                                HttpWebRequest request = WebRequest.Create(exportwebservicepath) as HttpWebRequest;
                                HttpWebResponse response = request.GetResponse() as HttpWebResponse;
                            }




                        }




                    }

                    #endregion


                }

                i++;

            }




            PutExportEntriesResults exportEntriesResults = new PutExportEntriesResults();

            return exportEntriesResults;
        }


        public void CloseExportConnection(CloseExportConnectionRunStep exportRunStep)
        {
            //conn.Close();
        }

    };
}
