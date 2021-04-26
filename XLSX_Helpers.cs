using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;
using X15 = DocumentFormat.OpenXml.Office2013.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using X15ac = DocumentFormat.OpenXml.Office2013.ExcelAc;
using Ap = DocumentFormat.OpenXml.ExtendedProperties;
using Vt = DocumentFormat.OpenXml.VariantTypes;
using System.Text.RegularExpressions;
using System.Globalization;

namespace XlsAutoReportWindowsService.XLSX
{
    class XLSX_Helpers
    {
        //Метод убирает из строки запрещенные спец символы.
        //Если не использовать, то при наличии в строке таких символов, вылетит ошибка.
        public static string ReplaceHexadecimalSymbols(string txt)
        {
            string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            if (txt != null)
                return Regex.Replace(txt, r, "", RegexOptions.Compiled);
            else
                return "---";
        }
        public static string GetCellReference(int colIndex, uint rowInd)
        {
            int div = colIndex;
            string colLetter = String.Empty;
            int mod = 0;

            while (div > 0)
            {
                mod = (div - 1) % 26;
                colLetter = (char)(65 + mod) + colLetter;
                div = (int)((div - mod) / 26);
            }
            return colLetter + rowInd.ToString();
        }
        /*Метод добавления текста в sharedStringTable*/
        public static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }
        /*
         Метод добавления ячейки в строку
         на входе: строка, столбец, значение, тип значения, необязательный параметр - стиль
         */
        public static int InsertCellToRow(Row row, int cellNum, string val, SharedStringTablePart sstp, uint style = 0U)
        {
            if (val != null)
            {
                string cellRef = GetCellReference(cellNum, row.RowIndex);
                Cell refCell = null;
                Cell newCell = new Cell() { CellReference = cellRef, StyleIndex = style };
                row.InsertBefore(newCell, refCell);
                decimal tmp = 0;
                int tmp2 = 0;
                DateTime tmp3 = DateTime.Now;
                if (decimal.TryParse(val.Replace(',', '.'), out tmp))
                {
                    newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    newCell.CellValue = new CellValue(val);
                }
                else if (int.TryParse(val.Replace(',', '.'), out tmp2))
                {
                    newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                    newCell.CellValue = new CellValue(val);
                }
                /* else if (DateTime.TryParseExact(val, "dd.MM.yyyy", CultureInfo.CurrentCulture, DateTimeStyles.None, out tmp3)) {
                     newCell.DataType = new EnumValue<CellValues>(CellValues.Date);
                     newCell.CellValue = new CellValue(val);
                 }
                 else if (DateTime.TryParseExact(val, "HH:mm:ss", CultureInfo.CurrentCulture, DateTimeStyles.None, out tmp3))
                 {
                     newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                     newCell.CellValue = new CellValue(val);
                 } */
                else
                {
                    newCell.CellValue = new CellValue(InsertSharedStringItem(val, sstp).ToString());
                    newCell.DataType = new EnumValue<CellValues>(CellValues.SharedString);
                }

                cellNum++;
            }
            return cellNum;
        }

        internal static ThemePart GenerateThemePartContent(ThemePart themePart)
        {
            A.Theme theme = new A.Theme() { Name = "Тема Office" };
            theme.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements = new A.ThemeElements();

            A.ColorScheme colorScheme = new A.ColorScheme() { Name = "Стандартная" };

            A.Dark1Color dark1Color = new A.Dark1Color();
            A.SystemColor systemColor1 = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color.Append(systemColor1);

            A.Light1Color light1Color = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color.Append(systemColor2);

            A.Dark2Color dark2Color = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color.Append(rgbColorModelHex1);

            A.Light2Color light2Color = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent1Color.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent5Color.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor.Append(rgbColorModelHex10);

            colorScheme.Append(dark1Color);
            colorScheme.Append(light1Color);
            colorScheme.Append(dark2Color);
            colorScheme.Append(light2Color);
            colorScheme.Append(accent1Color);
            colorScheme.Append(accent2Color);
            colorScheme.Append(accent3Color);
            colorScheme.Append(accent4Color);
            colorScheme.Append(accent5Color);
            colorScheme.Append(accent6Color);
            colorScheme.Append(hyperlink);
            colorScheme.Append(followedHyperlinkColor);

            A.FontScheme fontScheme2 = new A.FontScheme() { Name = "Стандартная" };

            A.MajorFont majorFont = new A.MajorFont();
            majorFont.Append(new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" });
            majorFont.Append(new A.EastAsianFont() { Typeface = "" });
            majorFont.Append(new A.ComplexScriptFont() { Typeface = "" });
            majorFont.Append(new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" });
            majorFont.Append(new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" });
            majorFont.Append(new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" });
            majorFont.Append(new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" });
            majorFont.Append(new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" });
            majorFont.Append(new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" });
            majorFont.Append(new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" });
            majorFont.Append(new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" });
            majorFont.Append(new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" });
            majorFont.Append(new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" });
            majorFont.Append(new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" });
            majorFont.Append(new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" });
            majorFont.Append(new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" });
            majorFont.Append(new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" });
            majorFont.Append(new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" });
            majorFont.Append(new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" });
            majorFont.Append(new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" });
            majorFont.Append(new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" });
            majorFont.Append(new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" });
            majorFont.Append(new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" });
            majorFont.Append(new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" });
            majorFont.Append(new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" });
            majorFont.Append(new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" });
            majorFont.Append(new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" });
            majorFont.Append(new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" });
            majorFont.Append(new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" });
            majorFont.Append(new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" });
            majorFont.Append(new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" });
            majorFont.Append(new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" });
            majorFont.Append(new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" });

            A.MinorFont minorFont = new A.MinorFont();
            minorFont.Append(new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" });
            minorFont.Append(new A.EastAsianFont() { Typeface = "" });
            minorFont.Append(new A.ComplexScriptFont() { Typeface = "" });
            minorFont.Append(new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック" });
            minorFont.Append(new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" });
            minorFont.Append(new A.SupplementalFont() { Script = "Hans", Typeface = "等线" });
            minorFont.Append(new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" });
            minorFont.Append(new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" });
            minorFont.Append(new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" });
            minorFont.Append(new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" });
            minorFont.Append(new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" });
            minorFont.Append(new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" });
            minorFont.Append(new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" });
            minorFont.Append(new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" });
            minorFont.Append(new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" });
            minorFont.Append(new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" });
            minorFont.Append(new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" });
            minorFont.Append(new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" });
            minorFont.Append(new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" });
            minorFont.Append(new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" });
            minorFont.Append(new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" });
            minorFont.Append(new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" });
            minorFont.Append(new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" });
            minorFont.Append(new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" });
            minorFont.Append(new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" });
            minorFont.Append(new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" });
            minorFont.Append(new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" });
            minorFont.Append(new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" });
            minorFont.Append(new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" });
            minorFont.Append(new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" });
            minorFont.Append(new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" });
            minorFont.Append(new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" });
            minorFont.Append(new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" });

            fontScheme2.Append(majorFont);
            fontScheme2.Append(minorFont);

            A.FormatScheme formatScheme = new A.FormatScheme() { Name = "Стандартная" };

            A.FillStyleList fillStyleList = new A.FillStyleList();

            A.SolidFill solidFill1 = new A.SolidFill();
            A.SchemeColor schemeColor1 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill1.Append(schemeColor1);

            A.GradientFill gradientFill1 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList1 = new A.GradientStopList();

            A.GradientStop gradientStop1 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation1 = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation1 = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint1 = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation1);
            schemeColor2.Append(saturationModulation1);
            schemeColor2.Append(tint1);

            gradientStop1.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList1.Append(gradientStop1);
            gradientStopList1.Append(gradientStop2);
            gradientStopList1.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill1 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill1.Append(gradientStopList1);
            gradientFill1.Append(linearGradientFill1);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList.Append(solidFill1);
            fillStyleList.Append(gradientFill1);
            fillStyleList.Append(gradientFill2);

            A.LineStyleList lineStyleList = new A.LineStyleList();

            A.Outline outline1 = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash1 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter1 = new A.Miter() { Limit = 800000 };

            outline1.Append(solidFill2);
            outline1.Append(presetDash1);
            outline1.Append(miter1);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList.Append(outline1);
            lineStyleList.Append(outline2);
            lineStyleList.Append(outline3);

            A.EffectStyleList effectStyleList = new A.EffectStyleList();

            A.EffectStyle effectStyle1 = new A.EffectStyle();
            A.EffectList effectList1 = new A.EffectList();

            effectStyle1.Append(effectList1);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow1 = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow1.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow1);

            effectStyle3.Append(effectList3);

            effectStyleList.Append(effectStyle1);
            effectStyleList.Append(effectStyle2);
            effectStyleList.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList.Append(solidFill5);
            backgroundFillStyleList.Append(solidFill6);
            backgroundFillStyleList.Append(gradientFill3);

            formatScheme.Append(fillStyleList);
            formatScheme.Append(lineStyleList);
            formatScheme.Append(effectStyleList);
            formatScheme.Append(backgroundFillStyleList);

            themeElements.Append(colorScheme);
            themeElements.Append(fontScheme2);
            themeElements.Append(formatScheme);
            A.ObjectDefaults objectDefaults = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList = new A.ExtraColorSchemeList();
            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList = new A.OfficeStyleSheetExtensionList();

            theme.Append(themeElements);
            theme.Append(objectDefaults);
            theme.Append(extraColorSchemeList);
            theme.Append(officeStyleSheetExtensionList);

            themePart.Theme = theme;

            return themePart;
        }

        internal static ThemePart GenerateThemePartWeeklyReportContent(ThemePart themePart)
        {
            A.Theme theme = new A.Theme() { Name = "Тема Office" };
            theme.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            A.ThemeElements themeElements = new A.ThemeElements();

            A.ColorScheme colorScheme = new A.ColorScheme() { Name = "Стандартная" };

            A.Dark1Color dark1Color = new A.Dark1Color();
            A.SystemColor systemColor = new A.SystemColor() { Val = A.SystemColorValues.WindowText, LastColor = "000000" };

            dark1Color.Append(systemColor);

            A.Light1Color light1Color = new A.Light1Color();
            A.SystemColor systemColor2 = new A.SystemColor() { Val = A.SystemColorValues.Window, LastColor = "FFFFFF" };

            light1Color.Append(systemColor2);

            A.Dark2Color dark2Color = new A.Dark2Color();
            A.RgbColorModelHex rgbColorModelHex1 = new A.RgbColorModelHex() { Val = "44546A" };

            dark2Color.Append(rgbColorModelHex1);

            A.Light2Color light2Color = new A.Light2Color();
            A.RgbColorModelHex rgbColorModelHex2 = new A.RgbColorModelHex() { Val = "E7E6E6" };

            light2Color.Append(rgbColorModelHex2);

            A.Accent1Color accent1Color = new A.Accent1Color();
            A.RgbColorModelHex rgbColorModelHex3 = new A.RgbColorModelHex() { Val = "5B9BD5" };

            accent1Color.Append(rgbColorModelHex3);

            A.Accent2Color accent2Color = new A.Accent2Color();
            A.RgbColorModelHex rgbColorModelHex4 = new A.RgbColorModelHex() { Val = "ED7D31" };

            accent2Color.Append(rgbColorModelHex4);

            A.Accent3Color accent3Color = new A.Accent3Color();
            A.RgbColorModelHex rgbColorModelHex5 = new A.RgbColorModelHex() { Val = "A5A5A5" };

            accent3Color.Append(rgbColorModelHex5);

            A.Accent4Color accent4Color = new A.Accent4Color();
            A.RgbColorModelHex rgbColorModelHex6 = new A.RgbColorModelHex() { Val = "FFC000" };

            accent4Color.Append(rgbColorModelHex6);

            A.Accent5Color accent5Color = new A.Accent5Color();
            A.RgbColorModelHex rgbColorModelHex7 = new A.RgbColorModelHex() { Val = "4472C4" };

            accent5Color.Append(rgbColorModelHex7);

            A.Accent6Color accent6Color = new A.Accent6Color();
            A.RgbColorModelHex rgbColorModelHex8 = new A.RgbColorModelHex() { Val = "70AD47" };

            accent6Color.Append(rgbColorModelHex8);

            A.Hyperlink hyperlink = new A.Hyperlink();
            A.RgbColorModelHex rgbColorModelHex9 = new A.RgbColorModelHex() { Val = "0563C1" };

            hyperlink.Append(rgbColorModelHex9);

            A.FollowedHyperlinkColor followedHyperlinkColor = new A.FollowedHyperlinkColor();
            A.RgbColorModelHex rgbColorModelHex10 = new A.RgbColorModelHex() { Val = "954F72" };

            followedHyperlinkColor.Append(rgbColorModelHex10);

            colorScheme.Append(dark1Color);
            colorScheme.Append(light1Color);
            colorScheme.Append(dark2Color);
            colorScheme.Append(light2Color);
            colorScheme.Append(accent1Color);
            colorScheme.Append(accent2Color);
            colorScheme.Append(accent3Color);
            colorScheme.Append(accent4Color);
            colorScheme.Append(accent5Color);
            colorScheme.Append(accent6Color);
            colorScheme.Append(hyperlink);
            colorScheme.Append(followedHyperlinkColor);

            A.FontScheme fontScheme = new A.FontScheme() { Name = "Стандартная" };

            A.MajorFont majorFont = new A.MajorFont();
            majorFont.Append(new A.LatinFont() { Typeface = "Calibri Light", Panose = "020F0302020204030204" });
            majorFont.Append(new A.EastAsianFont() { Typeface = "" });
            majorFont.Append(new A.ComplexScriptFont() { Typeface = "" });
            majorFont.Append(new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック Light" });
            majorFont.Append(new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" });
            majorFont.Append(new A.SupplementalFont() { Script = "Hans", Typeface = "等线 Light" });
            majorFont.Append(new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" });
            majorFont.Append(new A.SupplementalFont() { Script = "Arab", Typeface = "Times New Roman" });
            majorFont.Append(new A.SupplementalFont() { Script = "Hebr", Typeface = "Times New Roman" });
            majorFont.Append(new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" });
            majorFont.Append(new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" });
            majorFont.Append(new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" });
            majorFont.Append(new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" });
            majorFont.Append(new A.SupplementalFont() { Script = "Khmr", Typeface = "MoolBoran" });
            majorFont.Append(new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" });
            majorFont.Append(new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" });
            majorFont.Append(new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" });
            majorFont.Append(new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" });
            majorFont.Append(new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" });
            majorFont.Append(new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" });
            majorFont.Append(new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" });
            majorFont.Append(new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" });
            majorFont.Append(new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" });
            majorFont.Append(new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" });
            majorFont.Append(new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" });
            majorFont.Append(new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" });
            majorFont.Append(new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" });
            majorFont.Append(new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" });
            majorFont.Append(new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" });
            majorFont.Append(new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" });
            majorFont.Append(new A.SupplementalFont() { Script = "Viet", Typeface = "Times New Roman" });
            majorFont.Append(new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" });
            majorFont.Append(new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" });

            A.MinorFont minorFont = new A.MinorFont();
            minorFont.Append(new A.LatinFont() { Typeface = "Calibri", Panose = "020F0502020204030204" });
            minorFont.Append(new A.EastAsianFont() { Typeface = "" });
            minorFont.Append(new A.ComplexScriptFont() { Typeface = "" });
            minorFont.Append(new A.SupplementalFont() { Script = "Jpan", Typeface = "游ゴシック" });
            minorFont.Append(new A.SupplementalFont() { Script = "Hang", Typeface = "맑은 고딕" });
            minorFont.Append(new A.SupplementalFont() { Script = "Hans", Typeface = "等线" });
            minorFont.Append(new A.SupplementalFont() { Script = "Hant", Typeface = "新細明體" });
            minorFont.Append(new A.SupplementalFont() { Script = "Arab", Typeface = "Arial" });
            minorFont.Append(new A.SupplementalFont() { Script = "Hebr", Typeface = "Arial" });
            minorFont.Append(new A.SupplementalFont() { Script = "Thai", Typeface = "Tahoma" });
            minorFont.Append(new A.SupplementalFont() { Script = "Ethi", Typeface = "Nyala" });
            minorFont.Append(new A.SupplementalFont() { Script = "Beng", Typeface = "Vrinda" });
            minorFont.Append(new A.SupplementalFont() { Script = "Gujr", Typeface = "Shruti" });
            minorFont.Append(new A.SupplementalFont() { Script = "Khmr", Typeface = "DaunPenh" });
            minorFont.Append(new A.SupplementalFont() { Script = "Knda", Typeface = "Tunga" });
            minorFont.Append(new A.SupplementalFont() { Script = "Guru", Typeface = "Raavi" });
            minorFont.Append(new A.SupplementalFont() { Script = "Cans", Typeface = "Euphemia" });
            minorFont.Append(new A.SupplementalFont() { Script = "Cher", Typeface = "Plantagenet Cherokee" });
            minorFont.Append(new A.SupplementalFont() { Script = "Yiii", Typeface = "Microsoft Yi Baiti" });
            minorFont.Append(new A.SupplementalFont() { Script = "Tibt", Typeface = "Microsoft Himalaya" });
            minorFont.Append(new A.SupplementalFont() { Script = "Thaa", Typeface = "MV Boli" });
            minorFont.Append(new A.SupplementalFont() { Script = "Deva", Typeface = "Mangal" });
            minorFont.Append(new A.SupplementalFont() { Script = "Telu", Typeface = "Gautami" });
            minorFont.Append(new A.SupplementalFont() { Script = "Taml", Typeface = "Latha" });
            minorFont.Append(new A.SupplementalFont() { Script = "Syrc", Typeface = "Estrangelo Edessa" });
            minorFont.Append(new A.SupplementalFont() { Script = "Orya", Typeface = "Kalinga" });
            minorFont.Append(new A.SupplementalFont() { Script = "Mlym", Typeface = "Kartika" });
            minorFont.Append(new A.SupplementalFont() { Script = "Laoo", Typeface = "DokChampa" });
            minorFont.Append(new A.SupplementalFont() { Script = "Sinh", Typeface = "Iskoola Pota" });
            minorFont.Append(new A.SupplementalFont() { Script = "Mong", Typeface = "Mongolian Baiti" });
            minorFont.Append(new A.SupplementalFont() { Script = "Viet", Typeface = "Arial" });
            minorFont.Append(new A.SupplementalFont() { Script = "Uigh", Typeface = "Microsoft Uighur" });
            minorFont.Append(new A.SupplementalFont() { Script = "Geor", Typeface = "Sylfaen" });


            fontScheme.Append(majorFont);
            fontScheme.Append(minorFont);

            A.FormatScheme formatScheme = new A.FormatScheme() { Name = "Стандартная" };

            A.FillStyleList fillStyleList = new A.FillStyleList();

            A.SolidFill solidFill = new A.SolidFill();
            A.SchemeColor schemeColor = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill.Append(schemeColor);

            A.GradientFill gradientFill = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList = new A.GradientStopList();

            A.GradientStop gradientStop = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor2 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation = new A.LuminanceModulation() { Val = 110000 };
            A.SaturationModulation saturationModulation = new A.SaturationModulation() { Val = 105000 };
            A.Tint tint = new A.Tint() { Val = 67000 };

            schemeColor2.Append(luminanceModulation);
            schemeColor2.Append(saturationModulation);
            schemeColor2.Append(tint);

            gradientStop.Append(schemeColor2);

            A.GradientStop gradientStop2 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor3 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation2 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation2 = new A.SaturationModulation() { Val = 103000 };
            A.Tint tint2 = new A.Tint() { Val = 73000 };

            schemeColor3.Append(luminanceModulation2);
            schemeColor3.Append(saturationModulation2);
            schemeColor3.Append(tint2);

            gradientStop2.Append(schemeColor3);

            A.GradientStop gradientStop3 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor4 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation3 = new A.LuminanceModulation() { Val = 105000 };
            A.SaturationModulation saturationModulation3 = new A.SaturationModulation() { Val = 109000 };
            A.Tint tint3 = new A.Tint() { Val = 81000 };

            schemeColor4.Append(luminanceModulation3);
            schemeColor4.Append(saturationModulation3);
            schemeColor4.Append(tint3);

            gradientStop3.Append(schemeColor4);

            gradientStopList.Append(gradientStop);
            gradientStopList.Append(gradientStop2);
            gradientStopList.Append(gradientStop3);
            A.LinearGradientFill linearGradientFill = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill.Append(gradientStopList);
            gradientFill.Append(linearGradientFill);

            A.GradientFill gradientFill2 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList2 = new A.GradientStopList();

            A.GradientStop gradientStop4 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor5 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation4 = new A.SaturationModulation() { Val = 103000 };
            A.LuminanceModulation luminanceModulation4 = new A.LuminanceModulation() { Val = 102000 };
            A.Tint tint4 = new A.Tint() { Val = 94000 };

            schemeColor5.Append(saturationModulation4);
            schemeColor5.Append(luminanceModulation4);
            schemeColor5.Append(tint4);

            gradientStop4.Append(schemeColor5);

            A.GradientStop gradientStop5 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor6 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.SaturationModulation saturationModulation5 = new A.SaturationModulation() { Val = 110000 };
            A.LuminanceModulation luminanceModulation5 = new A.LuminanceModulation() { Val = 100000 };
            A.Shade shade1 = new A.Shade() { Val = 100000 };

            schemeColor6.Append(saturationModulation5);
            schemeColor6.Append(luminanceModulation5);
            schemeColor6.Append(shade1);

            gradientStop5.Append(schemeColor6);

            A.GradientStop gradientStop6 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor7 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.LuminanceModulation luminanceModulation6 = new A.LuminanceModulation() { Val = 99000 };
            A.SaturationModulation saturationModulation6 = new A.SaturationModulation() { Val = 120000 };
            A.Shade shade2 = new A.Shade() { Val = 78000 };

            schemeColor7.Append(luminanceModulation6);
            schemeColor7.Append(saturationModulation6);
            schemeColor7.Append(shade2);

            gradientStop6.Append(schemeColor7);

            gradientStopList2.Append(gradientStop4);
            gradientStopList2.Append(gradientStop5);
            gradientStopList2.Append(gradientStop6);
            A.LinearGradientFill linearGradientFill2 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill2.Append(gradientStopList2);
            gradientFill2.Append(linearGradientFill2);

            fillStyleList.Append(solidFill);
            fillStyleList.Append(gradientFill);
            fillStyleList.Append(gradientFill2);

            A.LineStyleList lineStyleList = new A.LineStyleList();

            A.Outline outline = new A.Outline() { Width = 6350, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill2 = new A.SolidFill();
            A.SchemeColor schemeColor8 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill2.Append(schemeColor8);
            A.PresetDash presetDash = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter = new A.Miter() { Limit = 800000 };

            outline.Append(solidFill2);
            outline.Append(presetDash);
            outline.Append(miter);

            A.Outline outline2 = new A.Outline() { Width = 12700, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill3 = new A.SolidFill();
            A.SchemeColor schemeColor9 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill3.Append(schemeColor9);
            A.PresetDash presetDash2 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter2 = new A.Miter() { Limit = 800000 };

            outline2.Append(solidFill3);
            outline2.Append(presetDash2);
            outline2.Append(miter2);

            A.Outline outline3 = new A.Outline() { Width = 19050, CapType = A.LineCapValues.Flat, CompoundLineType = A.CompoundLineValues.Single, Alignment = A.PenAlignmentValues.Center };

            A.SolidFill solidFill4 = new A.SolidFill();
            A.SchemeColor schemeColor10 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill4.Append(schemeColor10);
            A.PresetDash presetDash3 = new A.PresetDash() { Val = A.PresetLineDashValues.Solid };
            A.Miter miter3 = new A.Miter() { Limit = 800000 };

            outline3.Append(solidFill4);
            outline3.Append(presetDash3);
            outline3.Append(miter3);

            lineStyleList.Append(outline);
            lineStyleList.Append(outline2);
            lineStyleList.Append(outline3);

            A.EffectStyleList effectStyleList = new A.EffectStyleList();

            A.EffectStyle effectStyle = new A.EffectStyle();
            A.EffectList effectList = new A.EffectList();

            effectStyle.Append(effectList);

            A.EffectStyle effectStyle2 = new A.EffectStyle();
            A.EffectList effectList2 = new A.EffectList();

            effectStyle2.Append(effectList2);

            A.EffectStyle effectStyle3 = new A.EffectStyle();

            A.EffectList effectList3 = new A.EffectList();

            A.OuterShadow outerShadow = new A.OuterShadow() { BlurRadius = 57150L, Distance = 19050L, Direction = 5400000, Alignment = A.RectangleAlignmentValues.Center, RotateWithShape = false };

            A.RgbColorModelHex rgbColorModelHex11 = new A.RgbColorModelHex() { Val = "000000" };
            A.Alpha alpha1 = new A.Alpha() { Val = 63000 };

            rgbColorModelHex11.Append(alpha1);

            outerShadow.Append(rgbColorModelHex11);

            effectList3.Append(outerShadow);

            effectStyle3.Append(effectList3);

            effectStyleList.Append(effectStyle);
            effectStyleList.Append(effectStyle2);
            effectStyleList.Append(effectStyle3);

            A.BackgroundFillStyleList backgroundFillStyleList = new A.BackgroundFillStyleList();

            A.SolidFill solidFill5 = new A.SolidFill();
            A.SchemeColor schemeColor11 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };

            solidFill5.Append(schemeColor11);

            A.SolidFill solidFill6 = new A.SolidFill();

            A.SchemeColor schemeColor12 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint5 = new A.Tint() { Val = 95000 };
            A.SaturationModulation saturationModulation7 = new A.SaturationModulation() { Val = 170000 };

            schemeColor12.Append(tint5);
            schemeColor12.Append(saturationModulation7);

            solidFill6.Append(schemeColor12);

            A.GradientFill gradientFill3 = new A.GradientFill() { RotateWithShape = true };

            A.GradientStopList gradientStopList3 = new A.GradientStopList();

            A.GradientStop gradientStop7 = new A.GradientStop() { Position = 0 };

            A.SchemeColor schemeColor13 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint6 = new A.Tint() { Val = 93000 };
            A.SaturationModulation saturationModulation8 = new A.SaturationModulation() { Val = 150000 };
            A.Shade shade3 = new A.Shade() { Val = 98000 };
            A.LuminanceModulation luminanceModulation7 = new A.LuminanceModulation() { Val = 102000 };

            schemeColor13.Append(tint6);
            schemeColor13.Append(saturationModulation8);
            schemeColor13.Append(shade3);
            schemeColor13.Append(luminanceModulation7);

            gradientStop7.Append(schemeColor13);

            A.GradientStop gradientStop8 = new A.GradientStop() { Position = 50000 };

            A.SchemeColor schemeColor14 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Tint tint7 = new A.Tint() { Val = 98000 };
            A.SaturationModulation saturationModulation9 = new A.SaturationModulation() { Val = 130000 };
            A.Shade shade4 = new A.Shade() { Val = 90000 };
            A.LuminanceModulation luminanceModulation8 = new A.LuminanceModulation() { Val = 103000 };

            schemeColor14.Append(tint7);
            schemeColor14.Append(saturationModulation9);
            schemeColor14.Append(shade4);
            schemeColor14.Append(luminanceModulation8);

            gradientStop8.Append(schemeColor14);

            A.GradientStop gradientStop9 = new A.GradientStop() { Position = 100000 };

            A.SchemeColor schemeColor15 = new A.SchemeColor() { Val = A.SchemeColorValues.PhColor };
            A.Shade shade5 = new A.Shade() { Val = 63000 };
            A.SaturationModulation saturationModulation10 = new A.SaturationModulation() { Val = 120000 };

            schemeColor15.Append(shade5);
            schemeColor15.Append(saturationModulation10);

            gradientStop9.Append(schemeColor15);

            gradientStopList3.Append(gradientStop7);
            gradientStopList3.Append(gradientStop8);
            gradientStopList3.Append(gradientStop9);
            A.LinearGradientFill linearGradientFill3 = new A.LinearGradientFill() { Angle = 5400000, Scaled = false };

            gradientFill3.Append(gradientStopList3);
            gradientFill3.Append(linearGradientFill3);

            backgroundFillStyleList.Append(solidFill5);
            backgroundFillStyleList.Append(solidFill6);
            backgroundFillStyleList.Append(gradientFill3);

            formatScheme.Append(fillStyleList);
            formatScheme.Append(lineStyleList);
            formatScheme.Append(effectStyleList);
            formatScheme.Append(backgroundFillStyleList);

            themeElements.Append(colorScheme);
            themeElements.Append(fontScheme);
            themeElements.Append(formatScheme);
            A.ObjectDefaults objectDefaults = new A.ObjectDefaults();
            A.ExtraColorSchemeList extraColorSchemeList = new A.ExtraColorSchemeList();
            A.OfficeStyleSheetExtensionList officeStyleSheetExtensionList = new A.OfficeStyleSheetExtensionList();

            theme.Append(themeElements);
            theme.Append(objectDefaults);
            theme.Append(extraColorSchemeList);
            theme.Append(officeStyleSheetExtensionList);

            themePart.Theme = theme;
            return themePart;
        }

        // Generates content of workbookStylesPart1.
        internal static WorkbookStylesPart GenerateWorkbookStylesPartContent(WorkbookStylesPart workbookStylesPart)
        {
            Stylesheet stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            /* NumberingFormats numberingFormats = new NumberingFormats() { Count = (UInt32Value)3U };
             numberingFormats.Append(new NumberingFormat() { NumberFormatId = (UInt32Value)165U, FormatCode = "dd/mm/yy;@" });
             numberingFormats.Append(new NumberingFormat() { NumberFormatId = (UInt32Value)166U, FormatCode = "h:mm;@" });
             numberingFormats.Append(new NumberingFormat() { NumberFormatId = (UInt32Value)167U, FormatCode = "[h]:mm:ss;@" });      */

            Fonts fonts = new Fonts() { Count = (UInt32Value)1U, KnownFonts = true };

            Font font1 = new Font();
            FontSize fontSize1 = new FontSize() { Val = 11D };
            Color color1 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName1 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering1 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet1 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme1 = new FontScheme() { Val = FontSchemeValues.Minor };

            font1.Append(fontSize1);
            font1.Append(color1);
            font1.Append(fontName1);
            font1.Append(fontFamilyNumbering1);
            font1.Append(fontCharSet1);
            font1.Append(fontScheme1);

            fonts.Append(font1);

            Fills fills = new Fills() { Count = (UInt32Value)3U };

            Fill fill1 = new Fill();
            PatternFill patternFill1 = new PatternFill() { PatternType = PatternValues.None };

            fill1.Append(patternFill1);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Theme = (UInt32Value)0U, Tint = -0.14999847407452621D };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            fills.Append(fill1);
            fills.Append(fill2);
            fills.Append(fill3);

            Borders borders = new Borders() { Count = (UInt32Value)1U };

            Border border1 = new Border();
            LeftBorder leftBorder1 = new LeftBorder();
            RightBorder rightBorder1 = new RightBorder();
            TopBorder topBorder1 = new TopBorder();
            BottomBorder bottomBorder1 = new BottomBorder();
            DiagonalBorder diagonalBorder1 = new DiagonalBorder();

            border1.Append(leftBorder1);
            border1.Append(rightBorder1);
            border1.Append(topBorder1);
            border1.Append(bottomBorder1);
            border1.Append(diagonalBorder1);

            borders.Append(border1);

            CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat1 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats.Append(cellFormat1);

            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)4U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment1 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat3.Append(alignment1);

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center  };

            cellFormat4.Append(alignment2);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat5.Append(alignment3);
            /*
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)165U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment4 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat6.Append(alignment4);

            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)166U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment5 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat7.Append(alignment5);

            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)167U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment6 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat8.Append(alignment6);  */

            cellFormats.Append(cellFormat2);
            cellFormats.Append(cellFormat3);
            cellFormats.Append(cellFormat4);
            cellFormats.Append(cellFormat5);
            //cellFormats.Append(cellFormat6);
            //cellFormats.Append(cellFormat7);
            //cellFormats.Append(cellFormat8);

            CellStyles cellStyles = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle1 = new CellStyle() { Name = "Обычный", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles.Append(cellStyle1);
            DifferentialFormats differentialFormats = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            Colors colors = new Colors();

            MruColors mruColors = new MruColors();
            Color color2 = new Color() { Rgb = "FF00FFFF" };
            Color color3 = new Color() { Rgb = "FFFF66CC" };
            Color color4 = new Color() { Rgb = "FF33CCFF" };
            Color color5 = new Color() { Rgb = "FFFF9933" };
            Color color6 = new Color() { Rgb = "FFFF6600" };

            mruColors.Append(color2);
            mruColors.Append(color3);
            mruColors.Append(color4);
            mruColors.Append(color5);
            mruColors.Append(color6);

            colors.Append(mruColors);

            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension.Append(slicerStyles);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles);

            stylesheetExtensionList.Append(stylesheetExtension);
            stylesheetExtensionList.Append(stylesheetExtension2);

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles);
            stylesheet.Append(differentialFormats);
            stylesheet.Append(tableStyles);
            stylesheet.Append(colors);
            stylesheet.Append(stylesheetExtensionList);

            workbookStylesPart.Stylesheet = stylesheet;

            return workbookStylesPart;
        }

        internal static WorkbookStylesPart GenerateWorkbookWeeklyReportStylesPartContent(WorkbookStylesPart part)
        {
            Stylesheet stylesheet = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            Fonts fonts = new Fonts() { Count = (UInt32Value)3U, KnownFonts = true };

            Font font = new Font();
            FontSize fontSize = new FontSize() { Val = 11D };
            Color color = new Color() { Theme = (UInt32Value)1U };
            FontName fontName = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet = new FontCharSet() { Val = 204 };
            FontScheme fontScheme = new FontScheme() { Val = FontSchemeValues.Minor };

            font.Append(fontSize);
            font.Append(color);
            font.Append(fontName);
            font.Append(fontFamilyNumbering);
            font.Append(fontCharSet);
            font.Append(fontScheme);

            Font font2 = new Font();
            Bold bold1 = new Bold();
            FontSize fontSize2 = new FontSize() { Val = 11D };
            Color color2 = new Color() { Theme = (UInt32Value)1U };
            FontName fontName2 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering2 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet2 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme2 = new FontScheme() { Val = FontSchemeValues.Minor };

            font2.Append(bold1);
            font2.Append(fontSize2);
            font2.Append(color2);
            font2.Append(fontName2);
            font2.Append(fontFamilyNumbering2);
            font2.Append(fontCharSet2);
            font2.Append(fontScheme2);

            Font font3 = new Font();
            Bold bold2 = new Bold();
            FontSize fontSize3 = new FontSize() { Val = 11D };
            Color color3 = new Color() { Theme = (UInt32Value)0U };
            FontName fontName3 = new FontName() { Val = "Calibri" };
            FontFamilyNumbering fontFamilyNumbering3 = new FontFamilyNumbering() { Val = 2 };
            FontCharSet fontCharSet3 = new FontCharSet() { Val = 204 };
            FontScheme fontScheme3 = new FontScheme() { Val = FontSchemeValues.Minor };

            font3.Append(bold2);
            font3.Append(fontSize3);
            font3.Append(color3);
            font3.Append(fontName3);
            font3.Append(fontFamilyNumbering3);
            font3.Append(fontCharSet3);
            font3.Append(fontScheme3);

            fonts.Append(font);
            fonts.Append(font2);
            fonts.Append(font3);

            Fills fills = new Fills() { Count = (UInt32Value)6U };

            Fill fill = new Fill();
            PatternFill patternFill = new PatternFill() { PatternType = PatternValues.None };

            fill.Append(patternFill);

            Fill fill2 = new Fill();
            PatternFill patternFill2 = new PatternFill() { PatternType = PatternValues.Gray125 };

            fill2.Append(patternFill2);

            Fill fill3 = new Fill();

            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Theme = (UInt32Value)0U, Tint = -0.14999847407452621D };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);

            fill3.Append(patternFill3);

            Fill fill4 = new Fill();

            PatternFill patternFill4 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor2 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.79998168889431442D };
            BackgroundColor backgroundColor2 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill4.Append(foregroundColor2);
            patternFill4.Append(backgroundColor2);

            fill4.Append(patternFill4);

            Fill fill5 = new Fill();

            PatternFill patternFill5 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor3 = new ForegroundColor() { Theme = (UInt32Value)4U };
            BackgroundColor backgroundColor3 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill5.Append(foregroundColor3);
            patternFill5.Append(backgroundColor3);

            fill5.Append(patternFill5);

            Fill fill6 = new Fill();

            PatternFill patternFill6 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor4 = new ForegroundColor() { Theme = (UInt32Value)4U, Tint = 0.59999389629810485D };
            BackgroundColor backgroundColor4 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill6.Append(foregroundColor4);
            patternFill6.Append(backgroundColor4);

            fill6.Append(patternFill6);

            fills.Append(fill);
            fills.Append(fill2);
            fills.Append(fill3);
            fills.Append(fill4);
            fills.Append(fill5);
            fills.Append(fill6);

            Borders borders = new Borders() { Count = (UInt32Value)8U };

            Border border = new Border();
            LeftBorder leftBorder = new LeftBorder();
            RightBorder rightBorder = new RightBorder();
            TopBorder topBorder = new TopBorder();
            BottomBorder bottomBorder = new BottomBorder();
            DiagonalBorder diagonalBorder = new DiagonalBorder();

            border.Append(leftBorder);
            border.Append(rightBorder);
            border.Append(topBorder);
            border.Append(bottomBorder);
            border.Append(diagonalBorder);

            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color4);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color5 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color5);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color6 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color6);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color7 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color7);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            Border border3 = new Border();

            LeftBorder leftBorder3 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color8 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder3.Append(color8);

            RightBorder rightBorder3 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color9 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder3.Append(color9);

            TopBorder topBorder3 = new TopBorder() { Style = BorderStyleValues.Medium };
            Color color10 = new Color() { Indexed = (UInt32Value)64U };

            topBorder3.Append(color10);

            BottomBorder bottomBorder3 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color11 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder3.Append(color11);
            DiagonalBorder diagonalBorder3 = new DiagonalBorder();

            border3.Append(leftBorder3);
            border3.Append(rightBorder3);
            border3.Append(topBorder3);
            border3.Append(bottomBorder3);
            border3.Append(diagonalBorder3);

            Border border4 = new Border();

            LeftBorder leftBorder4 = new LeftBorder() { Style = BorderStyleValues.Medium };
            Color color12 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder4.Append(color12);

            RightBorder rightBorder4 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color13 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder4.Append(color13);

            TopBorder topBorder4 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color14 = new Color() { Indexed = (UInt32Value)64U };

            topBorder4.Append(color14);

            BottomBorder bottomBorder4 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color15 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder4.Append(color15);
            DiagonalBorder diagonalBorder4 = new DiagonalBorder();

            border4.Append(leftBorder4);
            border4.Append(rightBorder4);
            border4.Append(topBorder4);
            border4.Append(bottomBorder4);
            border4.Append(diagonalBorder4);

            Border border5 = new Border();

            LeftBorder leftBorder5 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color16 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder5.Append(color16);

            RightBorder rightBorder5 = new RightBorder() { Style = BorderStyleValues.Medium };
            Color color17 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder5.Append(color17);

            TopBorder topBorder5 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color18 = new Color() { Indexed = (UInt32Value)64U };

            topBorder5.Append(color18);

            BottomBorder bottomBorder5 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color19 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder5.Append(color19);
            DiagonalBorder diagonalBorder5 = new DiagonalBorder();

            border5.Append(leftBorder5);
            border5.Append(rightBorder5);
            border5.Append(topBorder5);
            border5.Append(bottomBorder5);
            border5.Append(diagonalBorder5);

            Border border6 = new Border();
            LeftBorder leftBorder6 = new LeftBorder();

            RightBorder rightBorder6 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color20 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder6.Append(color20);
            TopBorder topBorder6 = new TopBorder();

            BottomBorder bottomBorder6 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color21 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder6.Append(color21);
            DiagonalBorder diagonalBorder6 = new DiagonalBorder();

            border6.Append(leftBorder6);
            border6.Append(rightBorder6);
            border6.Append(topBorder6);
            border6.Append(bottomBorder6);
            border6.Append(diagonalBorder6);

            Border border7 = new Border();
            LeftBorder leftBorder7 = new LeftBorder();

            RightBorder rightBorder7 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color22 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder7.Append(color22);

            TopBorder topBorder7 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color23 = new Color() { Indexed = (UInt32Value)64U };

            topBorder7.Append(color23);

            BottomBorder bottomBorder7 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color24 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder7.Append(color24);
            DiagonalBorder diagonalBorder7 = new DiagonalBorder();

            border7.Append(leftBorder7);
            border7.Append(rightBorder7);
            border7.Append(topBorder7);
            border7.Append(bottomBorder7);
            border7.Append(diagonalBorder7);

            Border border8 = new Border();
            LeftBorder leftBorder8 = new LeftBorder();

            RightBorder rightBorder8 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color25 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder8.Append(color25);
            TopBorder topBorder8 = new TopBorder();

            BottomBorder bottomBorder8 = new BottomBorder() { Style = BorderStyleValues.Medium };
            Color color26 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder8.Append(color26);
            DiagonalBorder diagonalBorder8 = new DiagonalBorder();

            border8.Append(leftBorder8);
            border8.Append(rightBorder8);
            border8.Append(topBorder8);
            border8.Append(bottomBorder8);
            border8.Append(diagonalBorder8);

            borders.Append(border);
            borders.Append(border2);
            borders.Append(border3);
            borders.Append(border4);
            borders.Append(border5);
            borders.Append(border6);
            borders.Append(border7);
            borders.Append(border8);

            CellStyleFormats cellStyleFormats = new CellStyleFormats() { Count = (UInt32Value)1U };
            CellFormat cellFormat = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U };

            cellStyleFormats.Append(cellFormat);

            CellFormats cellFormats = new CellFormats() { Count = (UInt32Value)12U };
            CellFormat cellFormat2 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U };

            CellFormat cellFormat3 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyAlignment = true };
            Alignment alignment = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat3.Append(alignment);

            CellFormat cellFormat4 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)2U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment2 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat4.Append(alignment2);

            CellFormat cellFormat5 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)0U, FormatId = (UInt32Value)0U, ApplyFill = true, ApplyAlignment = true };
            Alignment alignment3 = new Alignment() { Horizontal = HorizontalAlignmentValues.Center };

            cellFormat5.Append(alignment3);
            CellFormat cellFormat6 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)1U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat7 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)2U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat8 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)4U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat9 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)0U, FillId = (UInt32Value)0U, BorderId = (UInt32Value)3U, FormatId = (UInt32Value)0U, ApplyBorder = true };
            CellFormat cellFormat10 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)2U, FillId = (UInt32Value)4U, BorderId = (UInt32Value)7U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };
            CellFormat cellFormat11 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)3U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };
            CellFormat cellFormat12 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)5U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };
            CellFormat cellFormat13 = new CellFormat() { NumberFormatId = (UInt32Value)0U, FontId = (UInt32Value)1U, FillId = (UInt32Value)5U, BorderId = (UInt32Value)6U, FormatId = (UInt32Value)0U, ApplyFont = true, ApplyFill = true, ApplyBorder = true };

            cellFormats.Append(cellFormat2);
            cellFormats.Append(cellFormat3);
            cellFormats.Append(cellFormat4);
            cellFormats.Append(cellFormat5);
            cellFormats.Append(cellFormat6);
            cellFormats.Append(cellFormat7);
            cellFormats.Append(cellFormat8);
            cellFormats.Append(cellFormat9);
            cellFormats.Append(cellFormat10);
            cellFormats.Append(cellFormat11);
            cellFormats.Append(cellFormat12);
            cellFormats.Append(cellFormat13);

            CellStyles cellStyles = new CellStyles() { Count = (UInt32Value)1U };
            CellStyle cellStyle = new CellStyle() { Name = "Обычный", FormatId = (UInt32Value)0U, BuiltinId = (UInt32Value)0U };

            cellStyles.Append(cellStyle);
            DifferentialFormats differentialFormats = new DifferentialFormats() { Count = (UInt32Value)0U };
            TableStyles tableStyles1 = new TableStyles() { Count = (UInt32Value)0U, DefaultTableStyle = "TableStyleMedium2", DefaultPivotStyle = "PivotStyleLight16" };

            Colors colors = new Colors();

            MruColors mruColors = new MruColors();
            Color color27 = new Color() { Rgb = "FF00FFFF" };
            Color color28 = new Color() { Rgb = "FFFF66CC" };
            Color color29 = new Color() { Rgb = "FF33CCFF" };
            Color color30 = new Color() { Rgb = "FFFF9933" };
            Color color31 = new Color() { Rgb = "FFFF6600" };

            mruColors.Append(color27);
            mruColors.Append(color28);
            mruColors.Append(color29);
            mruColors.Append(color30);
            mruColors.Append(color31);

            colors.Append(mruColors);

            StylesheetExtensionList stylesheetExtensionList = new StylesheetExtensionList();

            StylesheetExtension stylesheetExtension = new StylesheetExtension() { Uri = "{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" };
            stylesheetExtension.AddNamespaceDeclaration("x14", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main");
            X14.SlicerStyles slicerStyles = new X14.SlicerStyles() { DefaultSlicerStyle = "SlicerStyleLight1" };

            stylesheetExtension.Append(slicerStyles);

            StylesheetExtension stylesheetExtension2 = new StylesheetExtension() { Uri = "{9260A510-F301-46a8-8635-F512D64BE5F5}" };
            stylesheetExtension2.AddNamespaceDeclaration("x15", "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main");
            X15.TimelineStyles timelineStyles = new X15.TimelineStyles() { DefaultTimelineStyle = "TimeSlicerStyleLight1" };

            stylesheetExtension2.Append(timelineStyles);

            stylesheetExtensionList.Append(stylesheetExtension);
            stylesheetExtensionList.Append(stylesheetExtension2);

            stylesheet.Append(fonts);
            stylesheet.Append(fills);
            stylesheet.Append(borders);
            stylesheet.Append(cellStyleFormats);
            stylesheet.Append(cellFormats);
            stylesheet.Append(cellStyles);
            stylesheet.Append(differentialFormats);
            stylesheet.Append(tableStyles1);
            stylesheet.Append(colors);
            stylesheet.Append(stylesheetExtensionList);

            part.Stylesheet = stylesheet;

            return part;
        }

        internal static ExtendedFilePropertiesPart GenerateExtendedFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            Ap.Properties properties = new Ap.Properties();
            properties.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application = new Ap.Application();
            application.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity = new Ap.DocumentSecurity();
            documentSecurity.Text = "0";
            Ap.ScaleCrop scaleCrop = new Ap.ScaleCrop();
            scaleCrop.Text = "false";

            Ap.HeadingPairs headingPairs = new Ap.HeadingPairs();

            Vt.VTVector vTVector = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR = new Vt.VTLPSTR();
            vTLPSTR.Text = "Листы";

            variant.Append(vTLPSTR);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt321 = new Vt.VTInt32();
            vTInt321.Text = "1";

            variant2.Append(vTInt321);

            vTVector.Append(variant);
            vTVector.Append(variant2);

            headingPairs.Append(vTVector);

            Ap.TitlesOfParts titlesOfParts = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)1U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Макеты";

            vTVector2.Append(vTLPSTR2);

            titlesOfParts.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate = new Ap.LinksUpToDate();
            linksUpToDate.Text = "false";
            Ap.SharedDocument sharedDocument = new Ap.SharedDocument();
            sharedDocument.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged = new Ap.HyperlinksChanged();
            hyperlinksChanged.Text = "false";
            Ap.ApplicationVersion applicationVersion = new Ap.ApplicationVersion();
            applicationVersion.Text = "15.0300";

            properties.Append(application);
            properties.Append(documentSecurity);
            properties.Append(scaleCrop);
            properties.Append(headingPairs);
            properties.Append(titlesOfParts);
            properties.Append(linksUpToDate);
            properties.Append(sharedDocument);
            properties.Append(hyperlinksChanged);
            properties.Append(applicationVersion);

            extendedFilePropertiesPart.Properties = properties;

            return extendedFilePropertiesPart;
        }

        internal static ExtendedFilePropertiesPart GenerateExtendedWeeklyFilePropertiesPartContent(ExtendedFilePropertiesPart extendedFilePropertiesPart)
        {
            Ap.Properties properties = new Ap.Properties();
            properties.AddNamespaceDeclaration("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");
            Ap.Application application = new Ap.Application();
            application.Text = "Microsoft Excel";
            Ap.DocumentSecurity documentSecurity = new Ap.DocumentSecurity();
            documentSecurity.Text = "0";
            Ap.ScaleCrop scaleCrop = new Ap.ScaleCrop();
            scaleCrop.Text = "false";

            Ap.HeadingPairs headingPairs = new Ap.HeadingPairs();

            Vt.VTVector vTVector = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Variant, Size = (UInt32Value)2U };

            Vt.Variant variant = new Vt.Variant();
            Vt.VTLPSTR vTLPSTR = new Vt.VTLPSTR();
            vTLPSTR.Text = "Листы";

            variant.Append(vTLPSTR);

            Vt.Variant variant2 = new Vt.Variant();
            Vt.VTInt32 vTInt32 = new Vt.VTInt32();
            vTInt32.Text = "2";

            variant2.Append(vTInt32);

            vTVector.Append(variant);
            vTVector.Append(variant2);

            headingPairs.Append(vTVector);

            Ap.TitlesOfParts titlesOfParts = new Ap.TitlesOfParts();

            Vt.VTVector vTVector2 = new Vt.VTVector() { BaseType = Vt.VectorBaseValues.Lpstr, Size = (UInt32Value)2U };
            Vt.VTLPSTR vTLPSTR2 = new Vt.VTLPSTR();
            vTLPSTR2.Text = "Расчет";
            Vt.VTLPSTR vTLPSTR3 = new Vt.VTLPSTR();
            vTLPSTR3.Text = "Выгрузка";

            vTVector2.Append(vTLPSTR2);
            vTVector2.Append(vTLPSTR3);

            titlesOfParts.Append(vTVector2);
            Ap.LinksUpToDate linksUpToDate = new Ap.LinksUpToDate();
            linksUpToDate.Text = "false";
            Ap.SharedDocument sharedDocument = new Ap.SharedDocument();
            sharedDocument.Text = "false";
            Ap.HyperlinksChanged hyperlinksChanged = new Ap.HyperlinksChanged();
            hyperlinksChanged.Text = "false";
            Ap.ApplicationVersion applicationVersion = new Ap.ApplicationVersion();
            applicationVersion.Text = "15.0300";

            properties.Append(application);
            properties.Append(documentSecurity);
            properties.Append(scaleCrop);
            properties.Append(headingPairs);
            properties.Append(titlesOfParts);
            properties.Append(linksUpToDate);
            properties.Append(sharedDocument);
            properties.Append(hyperlinksChanged);
            properties.Append(applicationVersion);

            extendedFilePropertiesPart.Properties = properties;

            return extendedFilePropertiesPart;
        }
    }
}
