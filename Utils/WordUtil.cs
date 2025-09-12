using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace Viettel_Report_Automation.Utils
{
    public static class WordUtil
    {
        public static List<DataTable> ReadAllTablesAsDataTable(DocX doc)
        {
            var result = new List<DataTable>();

            foreach (var table in doc.Tables)
            {
                var dt = new DataTable();

                if (table.Rows.Count == 0) continue;

                // Header: dùng dòng đầu tiên
                foreach (var cell in table.Rows[0].Cells)
                {
                    string colName = cell.Paragraphs[0].Text.Trim();
                    dt.Columns.Add(string.IsNullOrEmpty(colName) ? "Column" + dt.Columns.Count : colName);
                }

                // Các dòng dữ liệu
                for (int i = 1; i < table.Rows.Count; i++)
                {
                    var row = dt.NewRow();
                    for (int j = 0; j < table.Rows[i].Cells.Count; j++)
                    {
                        row[j] = table.Rows[i].Cells[j].Paragraphs[0].Text.Trim();
                    }
                    dt.Rows.Add(row);
                }

                result.Add(dt);
            }

            return result;
        }
    }
}
