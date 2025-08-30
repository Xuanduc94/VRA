using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Viettel_Report_Automation.Utils
{
    public class ExcelCellUpdater
    {
        /// <summary>
        /// Cập nhật giá trị của một ô trong file Excel
        /// </summary>
        /// <param name="filePath">Đường dẫn đến file Excel</param>
        /// <param name="worksheetName">Tên worksheet (để trống sẽ lấy worksheet đầu tiên)</param>
        /// <param name="cellReference">Tham chiếu ô (VD: "A1", "B2", "C10")</param>
        /// <param name="value">Giá trị mới cần cập nhật</param>
        public static void UpdateCellValue(string filePath, string worksheetName, string cellReference, object value)
        {
            SpreadsheetDocument document = null;
            try
            {
                using ( document = SpreadsheetDocument.Open(filePath, true))
                {
                    WorkbookPart workbookPart = document.WorkbookPart;

                    // Tìm worksheet
                    WorksheetPart worksheetPart = GetWorksheetPart(workbookPart, worksheetName);

                    if (worksheetPart == null)
                    {
                        throw new ArgumentException($"Không tìm thấy worksheet: {worksheetName}");
                    }

                    Worksheet worksheet = worksheetPart.Worksheet;
                    SheetData sheetData = worksheet.GetFirstChild<SheetData>();

                    // Tìm hoặc tạo ô
                    Cell cell = GetOrCreateCell(worksheet, cellReference);

                    // Cập nhật giá trị
                    SetCellValue(cell, value, workbookPart);

                    // Lưu thay đổi
                    worksheetPart.Worksheet.Save();
                    
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Lỗi khi cập nhật ô {cellReference}: {ex.Message}", ex);
            }
        }


        /// <summary>
        /// Lấy WorksheetPart theo tên hoặc lấy worksheet đầu tiên
        /// </summary>
        /// 

        public static bool saveAsFile(string filePath, string path)
        {

            try
            {
                File.Copy(filePath, path);

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("Đã xảy ra lỗi : "+ ex.Message +" Vui lòng thử lại sau");
            }

        }

        /// <summary>
        /// Lấy WorksheetPart theo tên hoặc lấy worksheet đầu tiên
        /// </summary>
        private static WorksheetPart GetWorksheetPart(WorkbookPart workbookPart, string worksheetName)
        {
            if (string.IsNullOrEmpty(worksheetName))
            {
                // Lấy worksheet đầu tiên
                return workbookPart.WorksheetParts.FirstOrDefault();
            }

            // Tìm worksheet theo tên
            Sheet sheet = workbookPart.Workbook.Descendants<Sheet>()
                .FirstOrDefault(s => s.Name == worksheetName);

            if (sheet == null) return null;

            return (WorksheetPart)workbookPart.GetPartById(sheet.Id);
        }

        /// <summary>
        /// Tìm hoặc tạo ô mới
        /// </summary>
        private static Cell GetOrCreateCell(Worksheet worksheet, string cellReference)
        {
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();

            // Tách row và column từ cell reference (VD: A1 -> row=1, col=A)
            string columnName = GetColumnName(cellReference);
            uint rowIndex = GetRowIndex(cellReference);

            // Tìm hoặc tạo row
            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
            //if (row == null)
            //{
            //    row = new Row() { RowIndex = rowIndex };
            //    sheetData.Append(row);
            //}

            // Tìm hoặc tạo cell
            Cell cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference.Value == cellReference);
            //if (cell == null)
            //{
            //    cell = new Cell() { CellReference = cellReference };

            //    // Chèn cell vào đúng vị trí theo thứ tự column
            //    Cell refCell = row.Elements<Cell>()
            //        .FirstOrDefault(c => string.Compare(c.CellReference.Value, cellReference, true) > 0);

            //    if (refCell != null)
            //    {
            //        row.InsertBefore(cell, refCell);
            //    }
            //    else
            //    {
            //        row.Append(cell);
            //    }
            //}

            return cell;
        }

        /// <summary>
        /// Cập nhật giá trị cho ô
        /// </summary>
        private static void SetCellValue(Cell cell, object value, WorkbookPart workbookPart)
        {

            cell.CellValue = new CellValue(value.ToString());
            //if (value == null)
            //{
            //    cell.CellValue = new CellValue("");
            //    cell.DataType = CellValues.String;
            //    return;
            //}

            //string valueString = value.ToString();

            // Xử lý theo kiểu dữ liệu
            //if (value is DateTime dateTime)
            //{
            //    // Chuyển DateTime thành số serial date của Excel
            //    double oaDate = dateTime.ToOADate();
            //    cell.CellValue = new CellValue(oaDate.ToString());
            //    cell.DataType = CellValues.Number;
            //}
            //else if (IsNumeric(valueString))
            //{
            //    // Số
            //    cell.CellValue = new CellValue(valueString);
            //    cell.DataType = CellValues.Number;
            //}
            //else if (bool.TryParse(valueString, out bool boolValue))
            //{
            //    // Boolean
            //    cell.CellValue = new CellValue(boolValue ? "1" : "0");
            //    cell.DataType = CellValues.Boolean;
            //}
            //else
            //{
            //    // Chuỗi - sử dụng SharedStringTable để tối ưu
            //    int stringIndex = GetSharedStringIndex(workbookPart, valueString);
            //    cell.CellValue = new CellValue(stringIndex.ToString());
            //    cell.DataType = CellValues.SharedString;
            //}
        }

        /// <summary>
        /// Lấy hoặc tạo index trong SharedStringTable
        /// </summary>
        private static int GetSharedStringIndex(WorkbookPart workbookPart, string text)
        {
            SharedStringTablePart shareStringPart = workbookPart.SharedStringTablePart ??
                workbookPart.AddNewPart<SharedStringTablePart>();

            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int index = 0;
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return index;
                }
                index++;
            }

            // Thêm string mới
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
            shareStringPart.SharedStringTable.Save();

            return index;
        }

        /// <summary>
        /// Lấy tên cột từ cell reference (VD: A1 -> A)
        /// </summary>
        private static string GetColumnName(string cellReference)
        {
            string columnName = "";
            foreach (char c in cellReference)
            {
                if (char.IsLetter(c))
                {
                    columnName += c;
                }
                else
                {
                    break;
                }
            }
            return columnName;
        }

        /// <summary>
        /// Lấy số hàng từ cell reference (VD: A1 -> 1)
        /// </summary>
        private static uint GetRowIndex(string cellReference)
        {
            string rowIndex = "";
            foreach (char c in cellReference)
            {
                if (char.IsDigit(c))
                {
                    rowIndex += c;
                }
            }
            return uint.Parse(rowIndex);
        }

        /// <summary>
        /// Kiểm tra chuỗi có phải số không
        /// </summary>
        private static bool IsNumeric(string value)
        {
            return double.TryParse(value, out _);
        }
    }

}
