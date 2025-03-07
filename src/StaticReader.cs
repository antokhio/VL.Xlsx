using ExcelDataReader;
using System.Data;
using VL.Lib.Collections;
using Path = VL.Lib.IO.Path;

namespace Xlsx;

public static class Experimental
{
    public static DataSet ReadExcelFile(Path? input)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);

        using (var stream = File.OpenRead(input))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                return reader.AsDataSet();

            }
        }
    }

    public static DataTable ReadTable(DataSet input, int tableIndex, out string tableName)
    {
        var table = input.Tables[tableIndex % input.Tables.Count];
        tableName = table.TableName;
        return table;
    }

    public static Spread<T> ReadTableRows<T>(DataTable input, RowParser<T> parser, int skipRows = 1)
    {
        var rows = input?.Rows;

        var builder = new SpreadBuilder<T>();

        for (int i = skipRows; i < rows?.Count; i++)
        {
            T row = parser.Invoke(rows[i]?.ItemArray.ToSpread() ?? Spread<object?>.Empty);

            builder.Add(row);
        }

        return builder.ToSpread();
    }
}
