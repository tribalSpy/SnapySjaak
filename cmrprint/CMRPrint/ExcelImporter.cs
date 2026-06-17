using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;

namespace CMRPrint
{
    public static class ExcelImporter
    {
        public static List<Customer> ImportCustomers(string filePath)
        {
            var customers = new List<Customer>();

            if (!File.Exists(filePath))
            {
                return customers;
            }

            using var workbook = new XLWorkbook(filePath);
            var worksheet = workbook.Worksheet(1);
            var firstRow = worksheet.FirstRowUsed();
            if (firstRow == null)
            {
                return customers;
            }

            var headers = new Dictionary<string, int>(System.StringComparer.OrdinalIgnoreCase);
            foreach (var cell in firstRow.CellsUsed())
            {
                headers[cell.GetString().Trim()] = cell.Address.ColumnNumber;
            }

            for (var row = firstRow.RowBelow(); !row.IsEmpty(); row = row.RowBelow())
            {
                var customer = new Customer
                {
                    Name = GetValue(row, headers, "Name"),
                    Address = GetValue(row, headers, "Address"),
                    City = GetValue(row, headers, "City"),
                    Country = GetValue(row, headers, "Country"),
                    VatNumber = GetValue(row, headers, "VatNumber"),
                };

                if (!string.IsNullOrWhiteSpace(customer.Name))
                {
                    customers.Add(customer);
                }
            }

            return customers;
        }

        private static string GetValue(IXLRow row, Dictionary<string, int> headers, string header)
        {
            return headers.TryGetValue(header, out var column)
                ? row.Cell(column).GetString().Trim()
                : string.Empty;
        }
    }
}
