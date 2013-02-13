using System;
using System.Xml;
using System.Collections.Generic;
using System.Text;
using Tridion.ContentManager.Templating;
using Tridion.ContentManager.Publishing.Rendering;
using Tridion.ContentManager.Templating.Assembly;
using Tridion.ContentManager.ContentManagement;
using Tridion.ContentManager;
using Tridion.ContentManager.CommunicationManagement;
using System.Data.OleDb;
using System.Data;
using System.IO;
using System.Web;

namespace com.klm.tridion.iss
{
    /// <summary>
    /// DisplayExcelAsJSON converts an Excel sheet into JSON output. Every row is processed as a JSON {element}, with key/value pairs.
    /// For every column value in  a row, the column name is used as the property name, and the column value is the value.
    /// </summary>
    [TcmTemplateTitle("Convert Excel sheet to JSON")]
    public class DisplayExcelAsJSON : ITemplate
    {
        private const String TEMP_FILE_LOCATION = @"D:\Temp\Tridion\";

        TemplatingLogger _logger = TemplatingLogger.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public void Transform(Engine engine, Package package)
        {
            // store Excel sheet on filesystem
            TcmUri compUri = new TcmUri(package.GetValue("Component.ID"));
            Component comp = (Component)engine.GetSession().GetObject(compUri);
            Binary binary = engine.PublishingContext.RenderedItem.AddBinary(comp);
            _logger.Info(binary.FilePath);
            String strTmpFileName = String.Format("{0}{1}.xls", TEMP_FILE_LOCATION, comp.Title.Replace(":", "_"));

            String result = String.Empty;
            try
            {
                bool storeFile = ByteArrayToFile(strTmpFileName, comp.BinaryContent.GetByteArray());
                _logger.Debug(String.Format("Storing file {0}, result = {1}", strTmpFileName, storeFile));
                // TODO: make sheet name a parameter (template parameter, component metadata?)
                result = ProcessSheet(strTmpFileName, "Sheet1");
            }
            catch (Exception ex)
            {
                _logger.Error(ex.Message);
            }
            finally
            {
                // delete file
                File.Delete(strTmpFileName);
                _logger.Info(String.Format("File {0} deleted successfully", strTmpFileName));
            }
            Item output = package.CreateStringItem(ContentType.Text, result);
            package.PushItem(Package.OutputName, output);
        }

        private String ProcessSheet(String strFileName, String sheetName)
        {
            // TODO: have connection string depend on XLS/XLSX?
            String connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFileName + ";Extended Properties='Excel 8.0;IMEX=1;Persist Security Info=False;HDR=YES'";

            StringBuilder builder = new StringBuilder();
            using (OleDbConnection cn = new OleDbConnection(connString))
            {
                cn.Open();
                //Get All Sheets Name
                DataTable sheetsName = cn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "Table" });
                //Get the First Sheet Name
                string firstSheetName = sheetsName.Rows[0][2].ToString();
                _logger.Info(String.Format("Using first sheet name {0}", firstSheetName));
                using (OleDbDataAdapter adapter = new OleDbDataAdapter(String.Format("SELECT * FROM [{0}]", firstSheetName), cn))
                {
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);
                    int rows = dt.Rows.Count;
                    int colums = dt.Columns.Count;
                    builder.Append("[");
                    int rowNum = 1;
                    foreach (DataRow dr in dt.Rows)
                    {
                        bool previousValueDisplayed = false;
                        builder.Append("{");
                        int colNum = 1;
                        foreach (DataColumn dc in dt.Columns)
                        {
                            String colName = dc.ToString().Trim();
                            String colVal = dr[dc.Ordinal].ToString().Trim();
                            if (colVal.Equals("✓"))
                            {
                                colVal = "true";
                            }
                            if (!String.IsNullOrEmpty(colVal))
                            {
                                // must a comma be displayed first?
                                if (previousValueDisplayed)
                                {
                                    builder.Append(",\n");
                                }
                                builder.Append(String.Format("\"{0}\" :\"{1}\"", colName, ToJSONString(colVal)));
                                previousValueDisplayed = true;
                            }
                            colNum++;
                        }
                        builder.Append("}");

                        if (rowNum < rows)
                        {
                            builder.Append(",");
                        }
                        rowNum++;

                    }
                    builder.Append("]");
                }
            }
            return builder.ToString();
        }

        /// <summary>
        /// Function to save byte array to a file
        /// </summary>
        /// <param name="_FileName">File name to save byte array</param>
        /// <param name="_ByteArray">Byte array to save to external file</param>
        /// <returns>Return true if byte array save successfully, if not return false</returns>
        public bool ByteArrayToFile(string _FileName, byte[] _ByteArray)
        {
            try
            {
                // check if file exists and if so, delete it first
                if (File.Exists(_FileName))
                {
                    File.Delete(_FileName);
                    _logger.Info(String.Format("File {0} already exists, removing it", _FileName));
                }

                // Open file for reading
                System.IO.FileStream _FileStream = new System.IO.FileStream(_FileName, System.IO.FileMode.Create, System.IO.FileAccess.Write);

                // Writes a block of bytes to this stream using data from a byte array.
                _FileStream.Write(_ByteArray, 0, _ByteArray.Length);

                // close file stream
                _FileStream.Close();

                return true;
            }
            catch (Exception _Exception)
            {
                // Error
                Console.WriteLine("Exception caught in process: {0}", _Exception.ToString());
            }

            // error occured, return false
            return false;
        }

        /// <summary>
        /// Evaluates all characters in a string and returns a new string,
        /// properly formatted for JSON compliance and bounded by double-quotes.
        /// </summary>
        /// <param name="text">string to be evaluated</param>
        /// <returns>new string, in JSON-compliant form</returns>
        public string ToJSONString(string text)
        {
            char[] charArray = text.ToCharArray();
            List<string> output = new List<string>();
            foreach (char c in charArray)
            {
                if (((int)c) == 8)              //Backspace
                    output.Add("\\b");
                else if (((int)c) == 9)         //Horizontal tab
                    output.Add("\\t");
                else if (((int)c) == 10)        //Newline
                    output.Add("\\n");
                else if (((int)c) == 12)        //Formfeed
                    output.Add("\\f");
                else if (((int)c) == 13)        //Carriage return
                    output.Add("\\n");
                else if (((int)c) == 34)        //Double-quotes (")
                    output.Add("\\" + c.ToString());
                else if (((int)c) == 47)        //Solidus   (/)
                    output.Add("\\" + c.ToString());
                else if (((int)c) == 92)        //Reverse solidus   (\)
                    output.Add("\\" + c.ToString());
                else if (((int)c) > 31)
                    output.Add(c.ToString());
                //TODO: add support for hexadecimal
            }
            return string.Join("", output.ToArray());
        }
    }
}
