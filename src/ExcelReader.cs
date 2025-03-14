using ExcelDataReader;
using System.Data;
using VL.Core.Import;
using Path = VL.Lib.IO.Path;

namespace Xlsx;

[ProcessNode()]
public class ExcelReader
{
    private Path _input = Path.Default;
    public Path Input
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

    public ExcelReader()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    private DataSet _output;
    public DataSet Output { get => _output; internal set => _output = value; }

    public void Read()
    {
        // TODO: VALIDATION
        bool isValid = _input.Exists;

        try
        {
            Output = Utils.ReadExcelFile(_input);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);

            Output = null;
        }
    }
}
