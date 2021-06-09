using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace ExcelUtility
{
    internal class ExcelFileComponent
    {
        private SpreadsheetDocument _document;
        internal ExcelFileComponent(SpreadsheetDocument document)
        {
            _document = document;
            WorkbookPart = _document.WorkbookPart;
            Sheet= (Sheet)_document.WorkbookPart.Workbook.Sheets.ChildElements.GetItem(0);
            Worksheet= ((WorksheetPart)WorkbookPart.GetPartById(Sheet.Id)).Worksheet;
            SheetData= (SheetData)Worksheet.ChildElements.GetItem(4);
        }
        public WorkbookPart WorkbookPart { get; set; }
        public Sheet Sheet { get; set; }
        public Worksheet Worksheet { get; set; }
        public SheetData SheetData { get; set; }
    }
    public class ExcelUtility
    {
        public static List<T> ReadAndConvertToObject<T>(Stream sm, Dictionary<string, string> columnHeaderMappings=null) where T : class
        {
            SpreadsheetDocument doc = SpreadsheetDocument.Open(sm, false);
            ExcelFileComponent fileComponent = new ExcelFileComponent(doc);
            try
            {
                List<T> fileData = new List<T>();
                for (int i = 1; i < fileComponent.SheetData.ChildElements.Count; i++)
                {
                    Row currentrow = (Row)fileComponent.SheetData.ChildElements.GetItem(i);
                    T filitem = (T)Activator.CreateInstance(typeof(T));
                    for (int j = 0; j < currentrow.ChildElements.Count; j++)
                    {
                        string columnHeader = GetColumnHeader(fileComponent.SheetData, fileComponent.WorkbookPart, j);
                        string currentcellvalue = GetCellValue(fileComponent.SheetData, fileComponent.WorkbookPart, j, i);
                        var propName = columnHeaderMappings != null ? filitem.GetType().GetProperty(columnHeaderMappings[columnHeader]) :
                                            filitem.GetType().GetProperty(columnHeader);
                        if(!string.IsNullOrEmpty(currentcellvalue))
                        filitem.GetType().GetProperty(propName.Name).SetValue(filitem, Convert.ChangeType(currentcellvalue, propName.PropertyType));
                    }
                    fileData.Add(filitem);
                }
                return fileData;
            }
            catch  (Exception ex)
            {
                throw ex;
            }
            finally
            {
                doc.Dispose();
                sm.Dispose();
            }
        }
        private static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
        private static string GetCellValue(SheetData sheet,WorkbookPart wb,int  ColIndex,int RowIndex)
        {
            var node = (Cell)sheet.ChildElements.GetItem(RowIndex).ChildElements.GetItem(ColIndex);
           
            string cellValue = "";
            if (node.DataType != null)
            {
                if(node.DataType== CellValues.SharedString)
                {
                    var sharedItem = GetSharedStringItemById(wb, Int32.Parse(node.InnerText));
                    if (sharedItem.Text != null)
                    {
                        //code to take the string value  
                        cellValue = sharedItem.Text.Text;
                    }
                    else if (sharedItem.InnerText != null)
                    {
                        cellValue = sharedItem.InnerText;
                    }
                    else if (sharedItem.InnerXml != null)
                    {
                        cellValue = sharedItem.InnerXml;
                    }
                }
                else
                {
                    cellValue = node.InnerText;
                }
               
            }
            else
            {
                cellValue = node.InnerText;
            }
            return cellValue;
        }
        private static string GetColumnHeader(SheetData sheet, WorkbookPart wb, int ColIndex)
        {
           return GetCellValue(sheet, wb, ColIndex, 0);
        }
    }
}
