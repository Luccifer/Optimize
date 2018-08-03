using ExcelDna.Integration;
using System.Runtime.InteropServices;

/*
namespace Cognos
{
    [ComVisible(true)]
    public class Writer : XlCall
    {
        [ExcelCommand(Name = "writeInto")]
        public static void writeInto(ExcelReference refer, string value)
        {
            string text = "Hello World";
            Excel(xlcAlert, text);

            
            // Via COM works fine.
            dynamic Application = ExcelDnaUtil.Application;
            Application.Range[$"{refer.RowFirst},{refer.ColumnFirst}"].Value = text;
            
        }
    }
}
*/
