using System.Drawing;

using ClosedXML.Excel;

namespace Bau.Libraries.LibExcelFilesGenerator.Tools;

/// <summary>
///		Funciones de ayuda para tratamiento de Excel
/// </summary>
internal static class ExcelTools
{
	/// <summary>
	///		Convierte la alineación horizontal al enumerado de Excel
	/// </summary>
	internal static XLAlignmentHorizontalValues ConvertAlign(ExcelBuilder.HorizontalAlignment horizontalAlign)
	{
		return horizontalAlign switch
					{
						ExcelBuilder.HorizontalAlignment.Center => XLAlignmentHorizontalValues.Center,
						ExcelBuilder.HorizontalAlignment.Left => XLAlignmentHorizontalValues.Left,
						ExcelBuilder.HorizontalAlignment.Right => XLAlignmentHorizontalValues.Right,
						_ => XLAlignmentHorizontalValues.General
					};
	}

	/// <summary>
	///		Convierte la alineación vertical al enumerado del generador
	/// </summary>
	internal static ExcelBuilder.HorizontalAlignment ConvertAlign(XLAlignmentHorizontalValues horizontalAlign)
	{
		return horizontalAlign switch
					{
						XLAlignmentHorizontalValues.Center => ExcelBuilder.HorizontalAlignment.Center,
						XLAlignmentHorizontalValues.Left => ExcelBuilder.HorizontalAlignment.Left,
						XLAlignmentHorizontalValues.Right => ExcelBuilder.HorizontalAlignment.Right,
						_ => ExcelBuilder.HorizontalAlignment.Unknown
					};
	}

	/// <summary>
	///		Convierte la alineación vertical al enumerado de Excel
	/// </summary>
	internal static XLAlignmentVerticalValues ConvertAlign(ExcelBuilder.VerticalAlignment verticalAlign)
	{
		return verticalAlign switch
					{
						ExcelBuilder.VerticalAlignment.Center => XLAlignmentVerticalValues.Center,
						ExcelBuilder.VerticalAlignment.Top => XLAlignmentVerticalValues.Top,
						ExcelBuilder.VerticalAlignment.Bottom => XLAlignmentVerticalValues.Bottom,
						_ => XLAlignmentVerticalValues.Top
					};
	}

	/// <summary>
	///		Convierte la alineación vertical al enumerado del generador
	/// </summary>
	internal static ExcelBuilder.VerticalAlignment ConvertAlign(XLAlignmentVerticalValues verticalAlign)
	{
		return verticalAlign switch
					{
						XLAlignmentVerticalValues.Center => ExcelBuilder.VerticalAlignment.Center,
						XLAlignmentVerticalValues.Top => ExcelBuilder.VerticalAlignment.Top,
						XLAlignmentVerticalValues.Bottom => ExcelBuilder.VerticalAlignment.Bottom,
						_ => ExcelBuilder.VerticalAlignment.Unknown
					};
	}

	/// <summary>
	///		Convierte la alineación vertical al enumerado del generador
	/// </summary>
	internal static XLBorderStyleValues ConvertBorder(ExcelBuilder.BorderStyle border)
	{
		return border switch
					{
						ExcelBuilder.BorderStyle.DashDot => XLBorderStyleValues.DashDot,
						ExcelBuilder.BorderStyle.DashDotDot => XLBorderStyleValues.DashDotDot,
						ExcelBuilder.BorderStyle.Dashed => XLBorderStyleValues.Dashed,
						ExcelBuilder.BorderStyle.Dotted => XLBorderStyleValues.Dotted,
						ExcelBuilder.BorderStyle.Double => XLBorderStyleValues.Double,
						ExcelBuilder.BorderStyle.Hair => XLBorderStyleValues.Hair,
						ExcelBuilder.BorderStyle.Medium => XLBorderStyleValues.Medium,
						ExcelBuilder.BorderStyle.MediumDashDot => XLBorderStyleValues.MediumDashDot,
						ExcelBuilder.BorderStyle.MediumDashDotDot => XLBorderStyleValues.MediumDashDotDot,
						ExcelBuilder.BorderStyle.MediumDashed => XLBorderStyleValues.MediumDashed,
						ExcelBuilder.BorderStyle.SlantDashDot => XLBorderStyleValues.SlantDashDot,
						ExcelBuilder.BorderStyle.Thick => XLBorderStyleValues.Thick,
						ExcelBuilder.BorderStyle.Thin => XLBorderStyleValues.Thin,
						_ => XLBorderStyleValues.None
					};
	}

	/// <summary>
	///		Convierte un color de <see cref="Color"/> a <see cref="XLColor"/>
	/// </summary>
	internal static XLColor ConvertColor(Color? color) => ConvertColor(color, Color.Transparent);

	/// <summary>
	///		Convierte un color de <see cref="Color"/> a <see cref="XLColor"/>
	/// </summary>
	internal static XLColor ConvertColor(Color? color, Color defaultColor)
	{
		Color clrColor = color ?? defaultColor; // ... por quitar el nulo

			return XLColor.FromArgb(clrColor.A, clrColor.R, clrColor.G, clrColor.B);
	}

	/// <summary>
	///		Convierte un color de Excel a RGB
	/// </summary>
	internal static Color ConvertColor(XLColor color) => Color.FromArgb(color.Color.A, color.Color.R, color.Color.G, color.Color.B);
	
	/// <summary>
	///		Convierte un patrón de línea
	/// </summary>
	internal static XLBorderStyleValues ConvertLinePattern(ExcelBuilder.LinePattern pattern)
	{
		return pattern switch
				{
					ExcelBuilder.LinePattern.DashDot => XLBorderStyleValues.DashDot,
					ExcelBuilder.LinePattern.DashDotDot => XLBorderStyleValues.DashDotDot,
					ExcelBuilder.LinePattern.Dashed => XLBorderStyleValues.Dashed,
					ExcelBuilder.LinePattern.Dotted => XLBorderStyleValues.Dotted,
					ExcelBuilder.LinePattern.Double => XLBorderStyleValues.Double,
					ExcelBuilder.LinePattern.Hair => XLBorderStyleValues.Hair,
					ExcelBuilder.LinePattern.Medium => XLBorderStyleValues.Medium,
					ExcelBuilder.LinePattern.MediumDashDot => XLBorderStyleValues.MediumDashDot,
					ExcelBuilder.LinePattern.MediumDashDotDot => XLBorderStyleValues.MediumDashDotDot,
					ExcelBuilder.LinePattern.MediumDashed => XLBorderStyleValues.MediumDashed,
					ExcelBuilder.LinePattern.SlantDashDot => XLBorderStyleValues.SlantDashDot,
					ExcelBuilder.LinePattern.Thick => XLBorderStyleValues.Thick,
					ExcelBuilder.LinePattern.Thin => XLBorderStyleValues.Thin,
					_ => XLBorderStyleValues.None
				};
	}

	/// <summary>
	///		Convierte un patrón de relleno
	/// </summary>
	internal static XLFillPatternValues ConvertPattern(ExcelBuilder.Pattern pattern)
	{
		return pattern switch
				{
					ExcelBuilder.Pattern.Solid => XLFillPatternValues.Solid,
					ExcelBuilder.Pattern.DarkDown => XLFillPatternValues.DarkDown,
					ExcelBuilder.Pattern.DarkGray => XLFillPatternValues.DarkGray,
					ExcelBuilder.Pattern.DarkGrid => XLFillPatternValues.DarkGrid,
					ExcelBuilder.Pattern.DarkHorizontal => XLFillPatternValues.DarkHorizontal,
					ExcelBuilder.Pattern.DarkTrellis => XLFillPatternValues.DarkTrellis,
					ExcelBuilder.Pattern.DarkUp => XLFillPatternValues.DarkUp,
					ExcelBuilder.Pattern.DarkVertical => XLFillPatternValues.DarkVertical,
					ExcelBuilder.Pattern.Gray0625 => XLFillPatternValues.Gray0625,
					ExcelBuilder.Pattern.Gray125 => XLFillPatternValues.Gray125,
					ExcelBuilder.Pattern.LightDown => XLFillPatternValues.LightDown,
					ExcelBuilder.Pattern.LightGray => XLFillPatternValues.LightGray,
					ExcelBuilder.Pattern.LightGrid => XLFillPatternValues.LightGrid,
					ExcelBuilder.Pattern.LightHorizontal => XLFillPatternValues.LightHorizontal,
					ExcelBuilder.Pattern.LightTrellis => XLFillPatternValues.LightTrellis,
					ExcelBuilder.Pattern.LightUp => XLFillPatternValues.LightUp,
					ExcelBuilder.Pattern.LightVertical => XLFillPatternValues.LightVertical,
					ExcelBuilder.Pattern.MediumGray => XLFillPatternValues.MediumGray,
					_ => XLFillPatternValues.None
				};
	}
}