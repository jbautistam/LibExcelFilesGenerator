using ClosedXML.Excel;

namespace Bau.Libraries.LibExcelFilesGenerator;

/// <summary>
///		Generador de Excel
/// </summary>
public class ExcelBuilder
{ 
	/// <summary>
	///		Alineación horizontal
	/// </summary>
	public enum HorizontalAlignment
	{
		/// <summary>Sin alineación</summary>
		Unknown,
		/// <summary>Izquierda</summary>
		Left,
		/// <summary>Centrado</summary>
		Center,
		/// <summary>Derecha</summary>
		Right
	}
	/// <summary>
	///		Alineación vertical
	/// </summary>
	public enum VerticalAlignment
	{
		/// <summary>Sin alineación</summary>
		Unknown,
		/// <summary>Superior</summary>
		Top,
		/// <summary>Centrado</summary>
		Center,
		/// <summary>Inferior</summary>
		Bottom
	}
	/// <summary>
	///		Tipo de datos
	/// </summary>
	public enum DataType
	{
		/// <summary>Cadena</summary>
		String,
		/// <summary>Fecha</summary>
		Date,
		/// <summary>Número</summary>
		Integer,
		/// <summary>Decimal</summary>
		Decimal,
		/// <summary>Valor lógico</summary>
		Boolean,
		/// <summary>Hora</summary>
		Time
	}
	/// <summary>
	///		Estilo de los bordes
	/// </summary>
	public enum BorderStyle
	{
		None,
		DashDot,
		DashDotDot,
		Dashed,
		Dotted,
		Double,
		Hair,
		Medium,
		MediumDashDot,
		MediumDashDotDot,
		MediumDashed,
		SlantDashDot,
		Thick,
		Thin
	}
    /// <summary>
    ///		Patrón de relleno
    /// </summary>
    public enum Pattern
    {
        None,
        Solid,
        DarkDown,
        DarkGray,
        DarkGrid,
        DarkHorizontal,
        DarkTrellis,
        DarkUp,
        DarkVertical,
        Gray0625,
        Gray125,
        LightDown,
        LightGray,
        LightGrid,
        LightHorizontal,
        LightTrellis,
        LightUp,
        LightVertical,
        MediumGray
    }
    /// <summary>
    ///		Patrón de línea
    /// </summary>
    public enum LinePattern
    {
		None,
		DashDot,
		DashDotDot,
		Dashed,
		Dotted,
		Double,
		Hair,
		Medium,
		MediumDashDot,
		MediumDashDotDot,
		MediumDashed,
		SlantDashDot,
		Thick,
		Thin
	}

	// Variables privadas
	private XLWorkbook? _excelFile = null;

	public ExcelBuilder(string fileName)
	{ 
		// Asigna las propiedades
		FileName = fileName;
		// Asigna los generadores
		SheetBuilder = new ExcelSheetBuilder(this);
	}

	/// <summary>
	///		Carga un archivo
	/// </summary>
	public void Load()
	{
		ExcelFile = new XLWorkbook(FileName);
	}

	/// <summary>
	///		Carga un archivo a partir de un stream en memoria
	/// </summary>
	public void Load(MemoryStream stream)
	{
		ExcelFile = new XLWorkbook(stream);
	}

	/// <summary>
	///		Obtiene un stream en memoria con el contenido del archivo
	/// </summary>
	public MemoryStream GetStream()
	{
		MemoryStream stream = new();

			// Grabamos el archivo sobre el stream
			ExcelFile.SaveAs(stream);
			// Lo colocamos en la posición 0
			stream.Position = 0;
			// y devolvemos el stream
			return stream;
	}

	/// <summary>
	///		Obtiene los bytes con el contenido del archivo
	/// </summary>
	public byte[] GetBytes() => GetStream().ToArray();

	/// <summary>
	///		Graba el archivo Excel
	/// </summary>
	public void Save(bool validate = false)
	{
		const string Extension = ".xlsx";
		bool mustRename = false;

			// La librería sólo permite grabar con extensión XLSX
			if (!FileName.EndsWith(Extension, StringComparison.CurrentCultureIgnoreCase))
			{
				mustRename = true;
				FileName = FileName + Extension;
			}
			// Graba el archivo
			ExcelFile.SaveAs(FileName, validate);
			// ... y en su caso, vuelve a dejar el nombre como estaba
			if (mustRename)
				File.Move(FileName, FileName[..^Extension.Length]);
	}

	/// <summary>
	///		Cierra el archivo Excel (y libera su memoria)
	/// </summary>
	public void Close()
	{
		ExcelFile.Dispose();
		_excelFile = null;
	}

	/// <summary>
	///		Comprueba si existe una hoja
	/// </summary>
	public bool ExistsWorkSheet(string sheetName) => ExcelFile.Worksheets.TryGetWorksheet(sheetName, out IXLWorksheet _);

	/// <summary>
	///		Obtiene el nombre de la hoja activa
	/// </summary>
	public string GetWorkSheetName() => SheetBuilder.SheetName;

	/// <summary>
	///		Obtiene el generador de la hoja actual
	/// </summary>
	public ExcelSheetBuilder WithWorkSheet() => SheetBuilder;

	/// <summary>
	///		Obtiene un generador para una hoja
	/// </summary>
	public ExcelSheetBuilder WithWorkSheet(string sheetName)
	{ 
		// Obtiene el generador de hoja
		SheetBuilder.SheetName = sheetName;
		// Devuelve el generador
		return SheetBuilder;
	}

	/// <summary>
	///		Obtiene un generador para una hoja por su índice
	/// </summary>
	public ExcelSheetBuilder WithWorkSheet(int sheetIndex) => SheetBuilder.WithSheet(sheetIndex);

	/// <summary>
	///		Obtiene un generador de celda
	/// </summary>
	public ExcelCellBuilder WithCell(int row, int column) => SheetBuilder.WithCell(row, column);

	/// <summary>
	///		Obtiene un generador de fila
	/// </summary>
	public ExcelRowBuilder WithRow(int row) => SheetBuilder.WithRow(row);

	/// <summary>
	///		Obtiene un generador de columna
	/// </summary>
	public ExcelColumnBuilder WithColumn(int column) => SheetBuilder.WithColumn(column);

	/// <summary>
	///		Obtiene un generador de rango
	/// </summary>
	public ExcelRangeBuilder WithRange(int startRow, int startColumn, int endRow, int endColumn) => SheetBuilder.WithRange(startRow, startColumn, endRow, endColumn);

	/// <summary>
	///		Archivo Excel
	/// </summary>
	internal XLWorkbook ExcelFile 
	{
		set { _excelFile = value; }
		get
		{
			// Crea el objeto del archivo si no existía
			if (_excelFile is null)
				_excelFile = new XLWorkbook();
			// Devuelve el archivo
			return _excelFile;
		}
	}

	/// <summary>
	///		Nombre de archivo
	/// </summary>
	public string FileName { get; set; }

	/// <summary>
	///		Generador de hoja actual
	/// </summary>
	private ExcelSheetBuilder SheetBuilder { get; }
}