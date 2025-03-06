using System.Data;
using VL.Core.Import;
using VL.Lib.Collections;

namespace Xlsx;

[ProcessNode()]
public class TableRows
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

    private Spread<Spread<string>>? _output;

    public void Read()
    {
        var rows = _input?.Rows;

        if (rows == null)
        {
            _output = Spread<Spread<string>>.Empty;
            return;
        }

        var groupBuilder = new SpreadBuilder<SpreadBuilder<string>>();

        var firstRowItems = rows[0].ItemArray;

        firstRowItems.ForEach(x =>
        {
            groupBuilder.Add(new SpreadBuilder<string>());
        });

        for (int i = _skipRows; i < rows.Count; i++)
        {
            var row = rows[i];

            for (int j = 0; j < row.ItemArray.Count(); j++)
            {
                groupBuilder[j].Add(row.ItemArray[j]?.ToString() ?? string.Empty);
            }
        }

        _output = groupBuilder.Select(group => group.ToSpread()).ToSpread();
    }

    public void Update([Pin(PinGroupKind = VL.Model.PinGroupKind.Collection, PinGroupDefaultCount = 2)] out Spread<Spread<string>> output)
    {
        output = _output ?? Spread<Spread<string>>.Empty;
    }
}
