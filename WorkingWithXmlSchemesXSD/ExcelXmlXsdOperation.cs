using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Text;

namespace WorkingWithXmlSchemesXSD
{

    public static class ExcelXmlXsdOperation
    {
        private readonly static string XmlHeader = @"<?xml version=""1.0"" encoding=""UTF-8""?>
         <MyClass xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance""  xsi:noNamespaceSchemaLocation=""Book.xsd"" >";
        private readonly static string XmlFooter = "</MyClass>";
        private readonly static string XmlNodeSchema =
            @"
                <book>
                   <title> {0} </title> 
                   <price> {1} <price> 
                   <releaseDate> {2} </releaseDate>
                   <author>
                         <yas>  {3} </yas>   
                         <name> {4} </name>       
                         <gender> {5} </gender>
                         <dateOfBirth> {6} </dateOfBirth>
                         <placeOfBirth> {7} </placeOfBirth>
                         <address> {8} </address>
                   </author>
                </book>
               ";
        public static DataTable GetExcelDataInDataTableFormat(string excelPath, string sheetName)
        {
            DataTable dataTable = new DataTable();
            string connectionString = string.Empty;
            DirectoryInfo directoryInfo = new DirectoryInfo(excelPath);

            FileInfo fileInfo = directoryInfo.GetFiles("*.xlsx").FirstOrDefault();
            connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileInfo.FullName + ";Extended Properties='Excel 8.0;HDR=NO;IMEX=1';";

            try
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();
                    string query = string.Format("SELECT * FROM [{0}]", sheetName);
                    OleDbDataAdapter dataAdapter = new OleDbDataAdapter(query, connection);
                    dataAdapter.Fill(dataTable);
                }
            }
            catch (Exception ex) {
                Console.WriteLine("Hata + " + ex.Message.ToString());
                throw;
            }
            return dataTable;
        }

        public static string ConvertDataTableDataToString(DataTable dataTable)
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.Append(XmlHeader);
            stringBuilder.Append(GetBodyXml(dataTable));
            stringBuilder.Append(System.Environment.NewLine);
            stringBuilder.Append(XmlFooter);
            return stringBuilder.ToString();
        }

        public static StringBuilder GetBodyXml(DataTable dataTable)
        {
            StringBuilder stringBuilder = new StringBuilder();

            for (int rowNumber = 1; rowNumber < dataTable.Rows.Count; rowNumber++)
            {
                int columnNumber = 0;
                Book book = new Book();
                book.author = new Author();

                book.Title = dataTable.Rows[rowNumber][columnNumber++].ToString().Trim();
                book.Price = Convert.ToDecimal(dataTable.Rows[rowNumber][columnNumber++].ToString().Trim());
                book.ReleaseDate = Convert.ToDateTime(dataTable.Rows[rowNumber][columnNumber++].ToString().Trim());
                book.author.Yas = Convert.ToByte(dataTable.Rows[rowNumber][columnNumber++].ToString().Trim());
                book.author.Name = dataTable.Rows[rowNumber][columnNumber++].ToString().Trim();
                book.author.Gender = Convert.ToChar(dataTable.Rows[rowNumber][columnNumber++].ToString().Trim());
                book.author.DateOfBirth = Convert.ToDateTime(dataTable.Rows[rowNumber][columnNumber++].ToString().Trim());
                book.author.PlaceOfBirth = dataTable.Rows[rowNumber][columnNumber++].ToString().Trim();
                book.author.Address = dataTable.Rows[rowNumber][columnNumber++].ToString().Trim();

                stringBuilder.Append(CombineXmlElements(book));
            }
            return stringBuilder;
        }

        public static string CombineXmlElements(Book book) 
        {
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendFormat(XmlNodeSchema,
                book.Title,
                book.Price,
                book.ReleaseDate,
                book.author.Yas,
                book.author.Name,
                book.author.Gender,
                book.author.DateOfBirth,
                book.author.PlaceOfBirth,
                book.author.Address
                );
            return stringBuilder.ToString();
        }

        public static bool ConvertExcelToXml()
        {
            try
            {
                string excelPath = @"C:\Users\...\Desktop\Excel_Xml";
                DataTable dataTable = ExcelXmlXsdOperation.GetExcelDataInDataTableFormat(excelPath, "Book$");
                string xmlString = ConvertDataTableDataToString(dataTable);
                string date = DateTime.Now.ToString("yyyyMMdd-HHmmss", System.Globalization.CultureInfo.InvariantCulture);
                string xmlFilePath = excelPath + "XML_" + date + ".xml";

                if (!Directory.Exists(excelPath))
                {
                    Directory.CreateDirectory(excelPath);
                }

                using (StreamWriter writer = new StreamWriter(xmlFilePath, true, Encoding.UTF8, 65536))
                {
                    writer.Write(xmlString);
                }

            }
            catch (Exception ex )
            {
                Console.WriteLine("Hata " + ex.Message.ToString()) ;
                throw;
            }


           return true;
        }

    }

}
