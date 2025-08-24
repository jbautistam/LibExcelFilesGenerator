using System.Drawing;
using ClosedXML.Excel;

namespace Bau.Libraries.LibExcelFilesGenerator;

/// <summary>
///		Generador de celdas para Excel
/// </summary>
public class ExcelCellBuilder
{
	public ExcelCellBuilder(ExcelSheetBuilder builder)
	{
		SheetBuilder = builder;
	}

	/// <summary>
	///		Obtiene una celda (está en el generador de celdas para que no haya que volver atrás cada vez que se rellena una celda)
	/// </summary>
	public ExcelCellBuilder WithCell(int row, int column)
	{ 
		// Cambia fila y columna
		Row = row;
		Column = column;
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Asigna un valor a una celda
	/// </summary>
	public ExcelCellBuilder WithValue(object? value)
	{ 
		// Asigna un valor a una celda
		if (value is not null)
			switch (value)
			{
				case double val:
						Cell.Value = val;
					break;
				case decimal val:
						Cell.Value = val;
					break;
				case float val:
						Cell.Value = val;
					break;
				case int val:
						Cell.Value = val;
					break;
				case short val:
						Cell.Value = val;
					break;
				case byte val:
						Cell.Value = val;
					break;
				case long val:
						Cell.Value = val;
					break;
				case bool val:
						Cell.Value = val;
					break;
				case DateTime val:
						Cell.Value = val;
					break;
				case string val:
						Cell.Value = val;
					break;
				default:
						Cell.Value = value.ToString();
					break;
			}
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Asigna un hipervínculo a una celda
	/// </summary>
	public ExcelCellBuilder WithHyperlink(string text, string url, string? toolTip = null)
	{ 
		// Introduce el vínculo
		Cell.Value = text;
		if (!string.IsNullOrWhiteSpace(url))
			Cell.SetHyperlink(new XLHyperlink(url.Replace('\\', '/'), toolTip));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Cambia el tamaño de fuente de la celda
	/// </summary>
	public ExcelCellBuilder WithSize(double size)
	{ 
		// Indica el formato de negrita
		Cell.Style.Font.SetFontSize(size);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene el tamaño de fuente de la celda
	/// </summary>
	public double GetSize() => Cell.Style.Font.FontSize;

	/// <summary>
	///		Cambia la negrita de la celda
	/// </summary>
	public ExcelCellBuilder WithBold(bool isBold = true)
	{ 
		// Indica el formato de negrita
		Cell.Style.Font.SetBold(isBold);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Comprueba si una celda está en negrita
	/// </summary>
	public bool CheckIsBold() => Cell.Style.Font.Bold;

	/// <summary>
	///		Cambia la cursiva de la celda
	/// </summary>
	public ExcelCellBuilder WithItalic(bool isItalic = true)
	{ 
		// Indica el formato de cursiva
		Cell.Style.Font.SetItalic(isItalic);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Comprueba si una celda está en cursiva
	/// </summary>
	public bool CheckIsItalic() => Cell.Style.Font.Italic;

	/// <summary>
	///		Cambia el color de fondo de la celda
	/// </summary>
	public ExcelCellBuilder WithBackground(System.Drawing.Color? background = null)
	{ 
		// Indica el color de fondo de la celda
		if (background != null && background != System.Drawing.Color.Transparent)
			Cell.Style.Fill.SetBackgroundColor(Tools.ExcelTools.ConvertColor(background));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene el color de fondo de la celda
	/// </summary>
	public System.Drawing.Color GetBackground() => Tools.ExcelTools.ConvertColor(Cell.Style.Fill.BackgroundColor);

	/// <summary>
	///		Cambia el color de texto de la celda
	/// </summary>
	public ExcelCellBuilder WithColor(System.Drawing.Color? color = null)
	{ 
		// Indica el color de texto de la celda
		if (color != null && color != System.Drawing.Color.Transparent)
			Cell.Style.Font.SetFontColor(Tools.ExcelTools.ConvertColor(color));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene el color de fondo de la celda
	/// </summary>
	public System.Drawing.Color Getcolor() => Tools.ExcelTools.ConvertColor(Cell.Style.Font.FontColor);

	/// <summary>
	///		Rellena una celda con una fórmula
	/// </summary>
	public ExcelCellBuilder WithFormula(string formula)
	{
		// Añade la fórmula
		if (!string.IsNullOrWhiteSpace(formula))
		{
			// Normaliza la fórmula
			formula = formula.Trim();
			if (!formula.StartsWith("="))
				formula = "=" + formula;
			// Asigna la fórmula
			Cell.FormulaA1 = formula;
		}
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Indica si se debe partir el texto de la celda por palabras
	/// </summary>
	public ExcelCellBuilder WithWrapText(bool wrapText = true)
	{ 
		// Indica si se va a partir el texto
		Cell.Style.Alignment.WrapText = wrapText;
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene el valor que indica si se debe partir el texto de la celda
	/// </summary>
	public bool CheckWrapText() => Cell.Style.Alignment.WrapText;

	/// <summary>
	///		Asigna la alineacion horizontal a una celda
	/// </summary>
	public ExcelCellBuilder WithHorizontalAlign(ExcelBuilder.HorizontalAlignment align)
	{	
		// Asigna la alineación horizontal
		if (align != ExcelBuilder.HorizontalAlignment.Unknown)
			Cell.Style.Alignment.SetHorizontal(Tools.ExcelTools.ConvertAlign(align));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene la alineación horizontal
	/// </summary>
	public ExcelBuilder.HorizontalAlignment GetHorizontalAlign() => Tools.ExcelTools.ConvertAlign(Cell.Style.Alignment.Horizontal);

	/// <summary>
	///		Asigna la alineacion vertical a una celda
	/// </summary>
	public ExcelCellBuilder WithVerticalAlign(ExcelBuilder.VerticalAlignment align)
	{ 
		// Asigna la alineación horizontal
		if (align != ExcelBuilder.VerticalAlignment.Unknown)
			Cell.Style.Alignment.SetVertical(Tools.ExcelTools.ConvertAlign(align));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene la alineación horizontal
	/// </summary>
	public ExcelBuilder.VerticalAlignment GetVerticalAlign() => Tools.ExcelTools.ConvertAlign(Cell.Style.Alignment.Vertical);

	/// <summary>
	///		Une varias celdas
	/// </summary>
	public ExcelCellBuilder WithMergeTo(int rowTo, int cellTo)
	{   
		// Une las celdas
		SheetBuilder.ActiveWorkSheet.Range(Cell, SheetBuilder.ActiveWorkSheet.Cell(rowTo, cellTo)).Merge();
		// Devuelve el generador
		return this;
	}
	/// <summary>
	///		Asigna el patrón al fondo de una celda
	/// </summary>
	public ExcelCellBuilder WithPattern(ExcelBuilder.Pattern pattern)
	{
		// Asigna el patrón
		if (pattern != ExcelBuilder.Pattern.None)
			Cell.Style.Fill.SetPatternType(Tools.ExcelTools.ConvertPattern(pattern));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Define el borde superior
	/// </summary>
	public ExcelCellBuilder WithBorderTop(Color? color, ExcelBuilder.LinePattern pattern)
	{
		// Asigna el borde superior
		Cell.Style.Border.TopBorderColor = Tools.ExcelTools.ConvertColor(color);
		Cell.Style.Border.TopBorder = Tools.ExcelTools.ConvertLinePattern(pattern);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Define el borde inferior
	/// </summary>
	public ExcelCellBuilder WithBorderBottom(Color? color, ExcelBuilder.LinePattern pattern)
	{
		// Asigna el borde superior
		Cell.Style.Border.BottomBorderColor = Tools.ExcelTools.ConvertColor(color);
		Cell.Style.Border.BottomBorder = Tools.ExcelTools.ConvertLinePattern(pattern);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Define el borde izquierdo
	/// </summary>
	public ExcelCellBuilder WithBorderLeft(Color? color, ExcelBuilder.LinePattern pattern)
	{
		// Asigna el borde superior
		Cell.Style.Border.LeftBorderColor = Tools.ExcelTools.ConvertColor(color);
		Cell.Style.Border.LeftBorder = Tools.ExcelTools.ConvertLinePattern(pattern);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Define el borde derecho
	/// </summary>
	public ExcelCellBuilder WithBorderRight(Color? color, ExcelBuilder.LinePattern pattern)
	{
		// Asigna el borde superior
		Cell.Style.Border.RightBorderColor = Tools.ExcelTools.ConvertColor(color);
		Cell.Style.Border.RightBorder = Tools.ExcelTools.ConvertLinePattern(pattern);
		// Devuelve el generador
		return this;
	}


	/// <summary>
	///		Obtiene el valor de una celda
	/// </summary>
	public object GetValue() => Cell.Value;

	/// <summary>
	///		Obtiene el valor de una celda en formato cadena
	/// </summary>
	public string GetValueString()
	{
		if (Cell.IsEmpty())
			return string.Empty;
		else
			return Cell.Value.ToString();
	}

	/// <summary>
	///		Crea un generador de estilo
	/// </summary>
	public ExcelStyleBuilder WithStyle() => new ExcelStyleBuilder(SheetBuilder, Cell.Style);

	/// <summary>
	///		Vuelve al generador de hojas
	/// </summary>
	public ExcelSheetBuilder Back() => SheetBuilder;

	/// <summary>
	///		Generador de hojas
	/// </summary>
	private ExcelSheetBuilder SheetBuilder { get; }

	/// <summary>
	///		Celda actual
	/// </summary>
	private IXLCell Cell => SheetBuilder.ActiveWorkSheet.Cell(Row, Column);

	/// <summary>
	///		Fila
	/// </summary>
	public int Row { get; internal set; } = 1;

	/// <summary>
	///		Columna
	/// </summary>
	public int Column { get; internal set; } = 1;
}