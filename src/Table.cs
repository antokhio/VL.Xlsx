using System.Data;
using VL.Core.Import;

namespace Xlsx;

[ProcessNode()]
public class Table
{
    private DataSet? _input;
    public DataSet? Input
    {
        internal get => _input;
        set
        {
            if (_input != value)
            {
                _input = value;

                Read();
            }
        }
    }

    private int _tableIndex;
    public int TableIndex
    {
        internal get => _tableIndex;
        set
        {
            if (_tableIndex != value)
            {
                _tableIndex = value;

                Read();
            }
        }
    }

    public void Read()
    {
        if (_input != null)
        {
            Output = _input.Tables[_tableIndex % _input.Tables.Count];
            TableName = Output.TableName;

            return;
        }
    }

    public DataTable Output { get; private set; }
    public string TableName { get; private set; }
}
