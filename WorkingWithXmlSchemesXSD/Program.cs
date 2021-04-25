using System;
using System.IO;
using System.Reflection;

namespace WorkingWithXmlSchemesXSD
{
    class Program
    {
        static void Main(string[] args)
        {

            //ExcelXmlXsdOperation.ConvertExcelToXml();
            ExcelXmlXsdOperation.ValidateXmlWithXsd();

            Console.ReadKey();
            Console.WriteLine("Hello World!");
        }
    }
}
