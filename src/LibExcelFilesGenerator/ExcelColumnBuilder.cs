using ClosedXML.Excel;

namespace Bau.Libraries.LibExcelFilesGenerator;

/// <summary>
///		Generador de columnas
/// </summary>
public class ExcelColumnBuilder
{
	public ExcelColumnBuilder(ExcelSheetBuilder sheetBuilder)
	{
		SheetBuilder = sheetBuilder;
	}

	/// <summary>
	///		Asigna el ancho de una columna
	/// </summary>
	public ExcelColumnBuilder WithWidth(double width)
	{ 
		// Asigna la altura
		ActiveColumn.Width = width;
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Ajusta una columna a su contenido
	/// </summary>
	public void WithAdjustToContent()
	{
		ActiveColumn.AdjustToContents();
	}

	/// <summary>
	///		Oculta una columna
	/// </summary>
	public ExcelColumnBuilder WithHide()
	{ 
		// Oculta la columna
		ActiveColumn.Hide();
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Bloquea una columna
	/// </summary>
	public ExcelColumnBuilder WithLock(bool isLocked = true)
	{ 
		// Bloquea / desbloquea la columna
		ActiveColumn.Style.Protection.SetLocked(isLocked);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Crea un generador de estilo
	/// </summary>
	public ExcelStyleBuilder WithStyle() => new(SheetBuilder, ActiveColumn.Style);

	/// <summary>
	///		Borra la columna
	/// </summary>
	public void Delete()
	{
		ActiveColumn.Delete();
	}

	/// <summary>
	///		Generador de hoja
	/// </summary>
	private ExcelSheetBuilder SheetBuilder { get; }

	/// <summary>
	///		Columna activa
	/// </summary>
	private IXLColumn ActiveColumn => SheetBuilder.ActiveWorkSheet.Column(Column);

	/// <summary>
	///		Columna
	/// </summary>
	public int Column { get; internal set; } = 1;
}
