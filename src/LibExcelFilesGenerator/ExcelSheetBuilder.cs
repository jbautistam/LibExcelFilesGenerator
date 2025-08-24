using ClosedXML.Excel;

namespace Bau.Libraries.LibExcelFilesGenerator;

/// <summary>
///		Generador de hojas
/// </summary>
public class ExcelSheetBuilder
{ 
	// Variables privadas
	private string _sheetName = default!;

	public ExcelSheetBuilder(ExcelBuilder excelBuilder)
	{ 
		// Guarda el generador
		ExcelBuilder = excelBuilder;
		// Asigna los generadores
		CellBuilder = new ExcelCellBuilder(this);
		ColumnBuilder = new ExcelColumnBuilder(this);
		RowBuilder = new ExcelRowBuilder(this);
		RangeBuilder = new ExcelRangeBuilder(this);
	}

	/// <summary>
	///		Crea o activa una hoja
	/// </summary>
	public ExcelSheetBuilder WithSheet(string sheetName)
	{ 
		// Cambia la hoja activa
		if (!string.IsNullOrEmpty(sheetName))
		{
			IXLWorksheet workSheet;

				// Cambia el nombre de hoja de cálculo activa
				SheetName = sheetName;
				// Activa la hoja de cálculo (crea una nueva hoja o activa una de las existentes)
				if (!ExcelBuilder.ExcelFile.Worksheets.TryGetWorksheet(sheetName, out workSheet))
					ActiveWorkSheet = ExcelBuilder.ExcelFile.Worksheets.Add(sheetName);
				else
					ActiveWorkSheet = workSheet;
		}
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene una hoja por su índice
	/// </summary>
	public ExcelSheetBuilder WithSheet(int sheetIndex)
	{
		IEnumerator<IXLWorksheet> enumerator = ExcelBuilder.ExcelFile.Worksheets.GetEnumerator();
		int workSheet = 0;

			// Busca la hoja con ese índice
			while (enumerator.MoveNext())
			{
				// Actualiza la hoja actual
				if (workSheet == sheetIndex)
				{
					SheetName = enumerator.Current.Name;
					ActiveWorkSheet = enumerator.Current;
				}
				// Incrementa el índice
				workSheet++;
			}
			// Devuelve el generador
			return this;
	}

	///// <summary>
	/////		Asigna una contraseña a la hoja de cálculo actual
	///// </summary>
	//public ExcelSheetBuilder SetPassword(string lockPassword, bool canInsertRows, bool canInsertColumns,
	//									 bool canDeleteRows, bool canDeleteColumns)
	//{
	//	IXLSheetProtection protection = ActiveWorkSheet.Protect(lockPassword);

	//		// Indica si se pueden añadir filas y/o columnas
	//		protection.InsertRows = canInsertRows;
	//		protection.InsertColumns = canInsertColumns;
	//		protection.DeleteRows = canDeleteRows;
	//		protection.DeleteColumns = canDeleteColumns;
	//		// Deja el resto de protección activa
	//		protection.FormatCells = true;
	//		protection.FormatColumns = true;
	//		protection.FormatRows = true;
	//		protection.AutoFilter = true;
	//		// Devuelve el generador
	//		return this;
	//}

	/// <summary>
	///		Inmoviliza las filas
	/// </summary>
	public ExcelSheetBuilder WithFreezeRows(int indexRows)
	{ 
		// Asigna las filas inmovilizadas
		ActiveWorkSheet.SheetView.FreezeRows(indexRows);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Inmoviliza las columnas
	/// </summary>
	public ExcelSheetBuilder WithFreezeColumns(int indexColumns)
	{ 
		// Asigna las columnas inmovilizadas
		ActiveWorkSheet.SheetView.FreezeColumns(indexColumns);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Crea autofiltro sobre las columnas
	/// </summary>
	public ExcelSheetBuilder WithAutoFilter()
	{ 
		// Añade un autofiltro a la columna
		ActiveWorkSheet.RangeUsed().SetAutoFilter();
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene la última fila escrita de una hoja de cálculo
	/// </summary>
	public int GetLastWritenRow() => ActiveWorkSheet.LastRowUsed().RowNumber() + 1;

	/// <summary>
	///		Obtiene la última columna escrita de una hoja de cálculo
	/// </summary>
	public int GetLastWritenColumn() => ActiveWorkSheet.LastColumnUsed().ColumnNumber() + 1;

	/// <summary>
	///		Obtiene un generador de celda
	/// </summary>
	public ExcelCellBuilder WithCell(int row, int column)
	{ 
		// Asigna los datos de la celda
		CellBuilder.Row = row;
		CellBuilder.Column = column;
		// Devuelve el generador
		return CellBuilder;
	}

	/// <summary>
	///		Obtiene un generador de fila
	/// </summary>
	public ExcelRowBuilder WithRow(int row)
	{ 
		// Asigna la fila
		RowBuilder.Row = row;
		// Devuelve el generador
		return RowBuilder;
	}

	/// <summary>
	///		Obtiene un generador para una columna
	/// </summary>
	public ExcelColumnBuilder WithColumn(int column)
	{ 
		// Asigna la columna
		ColumnBuilder.Column = column;
		// Devuelve el generador
		return ColumnBuilder;
	}

	/// <summary>
	///		Obtiene un generador de rango
	/// </summary>
	public ExcelRangeBuilder WithRange(int startRow, int startColumn, int endRow, int endColumn)
	{ 
		// Asigna los datos del rango
		RangeBuilder.StartRow = startRow;
		RangeBuilder.StartColumn = startColumn;
		RangeBuilder.EndRow = endRow;
		RangeBuilder.EndColumn = endColumn;
		// Devuelve el generador
		return RangeBuilder;
	}

	/// <summary>
	///		Actualiza todas las columnas
	/// </summary>
	public ExcelSheetBuilder WithAutoAdjustColumnsToContents()
	{
		// Ajusta las columnas al contenido
		for (int column = 1; column < GetLastWritenColumn(); column++)
			ActiveWorkSheet.Column(column).AdjustToContents();
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Auto ajusta las columnas
	/// </summary>
	public ExcelSheetBuilder WithAutoAdjustColumns()
	{
		// Auto ajusta las columnas
		ActiveWorkSheet.Columns().AdjustToContents();
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Auto ajusta las filas
	/// </summary>
	public ExcelSheetBuilder WithAutoAdjustRows()
	{
		// Auto ajusta las filas
		ActiveWorkSheet.Rows().AdjustToContents();
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Generador del archivo Excel
	/// </summary>
	internal ExcelBuilder ExcelBuilder { get; }

	/// <summary>
	///		Generador de fila
	/// </summary>
	private ExcelRowBuilder RowBuilder { get; }

	/// <summary>
	///		Generador de columna
	/// </summary>
	private ExcelColumnBuilder ColumnBuilder { get; }

	/// <summary>
	///		Generador de celdas
	/// </summary>
	private ExcelCellBuilder CellBuilder { get; }

	/// <summary>
	///		Generador de rangos
	/// </summary>
	private ExcelRangeBuilder RangeBuilder { get; }

	/// <summary>
	///		Hoja del libro activa
	/// </summary>
	internal IXLWorksheet ActiveWorkSheet { get; private set; } = default!;

	/// <summary>
	///		Nombre de hoja
	/// </summary>
	internal string SheetName 
	{
		get { return _sheetName; }
		set 
		{
			if (!string.IsNullOrEmpty(value) && !value.Equals(_sheetName, StringComparison.CurrentCultureIgnoreCase))
			{
				_sheetName = value;
				WithSheet(_sheetName);
			}
		}
	}
}