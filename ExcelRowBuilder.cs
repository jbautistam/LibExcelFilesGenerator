using ClosedXML.Excel;

namespace Bau.Libraries.LibExcelFilesGenerator;

/// <summary>
///		Generador de filas
/// </summary>
public class ExcelRowBuilder
{
	public ExcelRowBuilder(ExcelSheetBuilder sheetBuilder)
	{
		SheetBuilder = sheetBuilder;
	}

	/// <summary>
	///		Asigna el alto de una fila
	/// </summary>
	public ExcelRowBuilder WithHeight(double height)
	{ 
		// Asigna la altura
		ActiveRow.Height = height;
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Crea un generador de estilo
	/// </summary>
	public ExcelStyleBuilder WithStyle() => new ExcelStyleBuilder(SheetBuilder, ActiveRow.Style);

	/// <summary>
	///		Borra una fila
	/// </summary>
	public void Delete()
	{
		ActiveRow.Delete();
	}

	/// <summary>
	///		Generador de hoja
	/// </summary>
	public ExcelSheetBuilder SheetBuilder { get; }

	/// <summary>
	///		Fila activa
	/// </summary>
	private IXLRow ActiveRow => SheetBuilder.ActiveWorkSheet.Row(Row);

	/// <summary>
	///		Fila
	/// </summary>
	public int Row { get; internal set; } = 1;
}
