using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelUtility
{
    public class MemoryStreamExcelHelper
    {
        /// <summary>
        /// Update Excel file
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="spreadSheetName"></param>
        /// <param name="cellAddressRange"></param>
        /// <param name="cellValue"></param>
        /// <returns></returns>
        public static MemoryStream UpdateExcel(Stream stream, string spreadSheetName, Dictionary<string, string> updateDictionary)
        {
            XLWorkbook workbook = new XLWorkbook(stream);
            IXLWorksheet worksheet = workbook.Worksheet(spreadSheetName);

            foreach (string cellAddressRange in updateDictionary.Keys)
            {
                // key is the cell address range, value is the cell value
                worksheet.Cell(cellAddressRange).Value = updateDictionary[cellAddressRange];
            }

            MemoryStream ms = new MemoryStream();
            workbook.SaveAs(ms);

            return ms;
        }
    }
}
