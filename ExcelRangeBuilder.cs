using ClosedXML.Excel;

namespace Bau.Libraries.LibExcelFilesGenerator;

/// <summary>
///		Generador para rangos
/// </summary>
public class ExcelRangeBuilder
{
	public ExcelRangeBuilder(ExcelSheetBuilder sheetBuilder)
	{
		SheetBuilder = sheetBuilder;
	}

	/// <summary>
	///		Genera una validación numérica sobre el rango definido
	/// </summary>
	public ExcelRangeBuilder WithNumericValidation(double minimum, double maximum)
	{ 
		// Asigna la validación
		Range.CreateDataValidation().Decimal.Between(minimum, maximum);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Genera una validación para datos rellenos
	/// </summary>
	public ExcelRangeBuilder WithNoEmptyValidation()
	{   
		// Asigna la validación
		//	Range.SetDataValidation().AllowedValues = XLAllowedValues.
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Crea una validación por fechas en el rango
	/// </summary>
	public ExcelRangeBuilder WithDateValidation(DateTime? minimum, DateTime? maximum)
	{ 
		// Asigna la validación
		if (maximum is null && minimum is not null)
			Range.CreateDataValidation().Date.EqualOrGreaterThan(minimum.Value);
		else if (minimum is null && maximum is not null)
			Range.CreateDataValidation().Date.EqualOrGreaterThan(maximum.Value);
		else if (minimum is not null && maximum is not null)
			Range.CreateDataValidation().Date.Between(minimum.Value, maximum.Value);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Crea una validación para un tipo de datos
	/// </summary>
	public ExcelRangeBuilder WithTypeValidation(ExcelBuilder.DataType type)
	{   
		// Asigna el tipo permitido
		switch (type)
		{
			case ExcelBuilder.DataType.Integer:
					Range.CreateDataValidation().AllowedValues = XLAllowedValues.WholeNumber;
				break;
			case ExcelBuilder.DataType.Decimal:
					Range.CreateDataValidation().AllowedValues = XLAllowedValues.Decimal;
				break;
			case ExcelBuilder.DataType.String:
					Range.CreateDataValidation().AllowedValues = XLAllowedValues.AnyValue;
				break;
			case ExcelBuilder.DataType.Date:
					Range.CreateDataValidation().AllowedValues = XLAllowedValues.Date;
				break;
			case ExcelBuilder.DataType.Time:
					Range.CreateDataValidation().AllowedValues = XLAllowedValues.Time;
				break;
			default:
				throw new NotImplementedException("No se puede asignar este tipo a una validación");
		}
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Combina el rango de celdas
	/// </summary>
	public ExcelRangeBuilder WithMerge(bool merge = true)
	{
		// Combina las columnas
		Range.Merge(merge);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene el generador de celda en la primera fila del rango
	/// </summary>
	public ExcelCellBuilder WithCell() => SheetBuilder.WithCell(StartRow, StartColumn);

	/// <summary>
	///		Crea un generador de estilos
	/// </summary>
	public ExcelStyleBuilder WithStyle() => new ExcelStyleBuilder(SheetBuilder, Range.Style);

	/// <summary>
	///		Vuelve al generador padre
	/// </summary>
	public ExcelSheetBuilder Back() => SheetBuilder;

	/// <summary>
	///		Obtiene el rango definido
	/// </summary>
	private IXLRange Range => SheetBuilder.ActiveWorkSheet.Range(StartRow, StartColumn, EndRow, EndColumn);

	/// <summary>
	///		Fila inicial del rango
	/// </summary>
	public int StartRow { get; internal set; }

	/// <summary>
	///		Columna inicial de rango
	/// </summary>
	public int StartColumn { get; internal set; }

	/// <summary>
	///		Fila final del rango
	/// </summary>
	public int EndRow { get; internal set; }

	/// <summary>
	///		Columna final del rango
	/// </summary>
	public int EndColumn { get; internal set; }

	/// <summary>
	///		Generador de hoja del que depende el rango
	/// </summary>
	private ExcelSheetBuilder SheetBuilder { get; }
}
