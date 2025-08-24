using ClosedXML.Excel;

namespace Bau.Libraries.LibExcelFilesGenerator;

/// <summary>
///		Generador de Excel
/// </summary>
public class ExcelStyleBuilder
{
	public ExcelStyleBuilder(ExcelSheetBuilder builder, IXLStyle style)
	{
		SheetBuilder = builder;
		Style = style;
	}

	/// <summary>
	///		Cambia el tamaño de fuente de la celda
	/// </summary>
	public ExcelStyleBuilder WithSize(double size)
	{ 
		// Indica el formato de negrita
		Style.Font.SetFontSize(size);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene el tamaño de fuente de la celda
	/// </summary>
	public double GetSize() => Style.Font.FontSize;

	/// <summary>
	///		Cambia la negrita de la celda
	/// </summary>
	public ExcelStyleBuilder WithBold(bool isBold = true)
	{ 
		// Indica el formato de negrita
		Style.Font.SetBold(isBold);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Comprueba si una celda está en negrita
	/// </summary>
	public bool CheckIsBold() => Style.Font.Bold;

	/// <summary>
	///		Cambia la cursiva de la celda
	/// </summary>
	public ExcelStyleBuilder WithItalic(bool isItalic = true)
	{ 
		// Indica el formato de cursiva
		Style.Font.SetItalic(isItalic);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Comprueba si una celda está en cursiva
	/// </summary>
	public bool CheckIsItalic() => Style.Font.Italic;

	/// <summary>
	///		Cambia el color de fondo de la celda
	/// </summary>
	public ExcelStyleBuilder WithBackground(System.Drawing.Color? background = null)
	{ 
		// Indica el color de fondo de la celda
		if (background != null && background != System.Drawing.Color.Transparent)
			Style.Fill.SetBackgroundColor(Tools.ExcelTools.ConvertColor(background));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene el color de fondo de la celda
	/// </summary>
	public System.Drawing.Color GetBackground() => Tools.ExcelTools.ConvertColor(Style.Fill.BackgroundColor);

	/// <summary>
	///		Cambia el color de texto de la celda
	/// </summary>
	public ExcelStyleBuilder WithColor(System.Drawing.Color? color = null)
	{ 
		// Indica el color de texto de la celda
		if (color != null && color != System.Drawing.Color.Transparent)
			Style.Font.SetFontColor(Tools.ExcelTools.ConvertColor(color));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene el color de fondo de la celda
	/// </summary>
	public System.Drawing.Color Getcolor() => Tools.ExcelTools.ConvertColor(Style.Font.FontColor);

	/// <summary>
	///		Asigna la alineacion horizontal a una celda
	/// </summary>
	public ExcelStyleBuilder WithHorizontalAlign(ExcelBuilder.HorizontalAlignment align)
	{	
		// Asigna la alineación horizontal
		if (align != ExcelBuilder.HorizontalAlignment.Unknown)
			Style.Alignment.SetHorizontal(Tools.ExcelTools.ConvertAlign(align));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene la alineación horizontal
	/// </summary>
	public ExcelBuilder.HorizontalAlignment GetHorizontalAlign() => Tools.ExcelTools.ConvertAlign(Style.Alignment.Horizontal);

	/// <summary>
	///		Asigna la alineacion vertical a una celda
	/// </summary>
	public ExcelStyleBuilder WithVerticalAlign(ExcelBuilder.VerticalAlignment align)
	{ 
		// Asigna la alineación horizontal
		if (align != ExcelBuilder.VerticalAlignment.Unknown)
			Style.Alignment.SetVertical(Tools.ExcelTools.ConvertAlign(align));
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Obtiene la alineación horizontal
	/// </summary>
	public ExcelBuilder.VerticalAlignment GetVerticalAlign() => Tools.ExcelTools.ConvertAlign(Style.Alignment.Vertical);

	/// <summary>
	///		Asigna un borde completo de la celda
	/// </summary>
	public ExcelStyleBuilder WithBorder(ExcelBuilder.BorderStyle border, System.Drawing.Color? color = null)
	{
		// Asigna los bordes
		WithTopBorder(border, color);
		WithBottomBorder(border, color);
		WithLeftBorder(border, color);
		WithRightBorder(border, color);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Asigna un borde a la parte superior de la celda
	/// </summary>
	public ExcelStyleBuilder WithTopBorder(ExcelBuilder.BorderStyle border, System.Drawing.Color? color = null)
	{
		// Asigna el borde
		Style.Border.TopBorder = Tools.ExcelTools.ConvertBorder(border);
		Style.Border.TopBorderColor = Tools.ExcelTools.ConvertColor(color, System.Drawing.Color.Black);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Asigna un borde a la parte superior de la celda
	/// </summary>
	public ExcelStyleBuilder WithBottomBorder(ExcelBuilder.BorderStyle border, System.Drawing.Color? color = null)
	{
		// Asigna el borde
		Style.Border.BottomBorder = Tools.ExcelTools.ConvertBorder(border);
		Style.Border.BottomBorderColor = Tools.ExcelTools.ConvertColor(color, System.Drawing.Color.Black);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Asigna un borde a la parte iquierda de la celda
	/// </summary>
	public ExcelStyleBuilder WithLeftBorder(ExcelBuilder.BorderStyle border, System.Drawing.Color? color = null)
	{
		// Asigna el borde
		Style.Border.LeftBorder = Tools.ExcelTools.ConvertBorder(border);
		Style.Border.LeftBorderColor = Tools.ExcelTools.ConvertColor(color, System.Drawing.Color.Black);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Asigna un borde a la parte derecha de la celda
	/// </summary>
	public ExcelStyleBuilder WithRightBorder(ExcelBuilder.BorderStyle border, System.Drawing.Color? color = null)
	{
		// Asigna el borde
		Style.Border.RightBorder = Tools.ExcelTools.ConvertBorder(border);
		Style.Border.RightBorderColor = Tools.ExcelTools.ConvertColor(color, System.Drawing.Color.Black);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Asigna un borde a la parte derecha de la celda
	/// </summary>
	public ExcelStyleBuilder WithInsideBorder(ExcelBuilder.BorderStyle border, System.Drawing.Color? color = null)
	{
		// Asigna el borde
		Style.Border.InsideBorder = Tools.ExcelTools.ConvertBorder(border);
		Style.Border.InsideBorderColor = Tools.ExcelTools.ConvertColor(color, System.Drawing.Color.Black);
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Asigna un patrón de relleno
	/// </summary>
	public ExcelStyleBuilder WithPattern(ExcelBuilder.Pattern pattern)
	{
		// Devuelve el generador
		return this;
	}

	/// <summary>
	///		Vuelve al generador de hojas
	/// </summary>
	public ExcelSheetBuilder Back() => SheetBuilder;

	/// <summary>
	///		Generador de hojas
	/// </summary>
	private ExcelSheetBuilder SheetBuilder { get; }

	/// <summary>
	///		Estilo actual
	/// </summary>
	private IXLStyle Style { get; }
}