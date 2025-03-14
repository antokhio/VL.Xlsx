using System.Data;
using VL.Core.Import;
using VL.Lib.Collections;

namespace Xlsx;

public delegate T RowParser<T>(Spread<object?> row);

[ProcessNode()]
public class TableRowParser<T>
{
    private DataTable? _input;
    public DataTable? Input
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

    private int _skipRows;
    public int SkipRows
    {
        internal get => _skipRows;
        set
        {
            if (_skipRows != value)
            {
                _skipRows = value;

                Read();
            }
        }
    }

    private RowParser<T>? _parser;
    public RowParser<T>? Parser
    {
        internal get => _parser;
        set
        {
            if (_parser != value)
            {
                _parser = value;

                Read();
            }
        }
    }

    public Spread<T> Output { get; private set; }

    public void Read() => Output = Utils.ReadTableRows<T>(_input, _parser, _skipRows);
}
