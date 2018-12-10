/*
 * Copyright (c) 2012-2013, bad robot (london) ltd.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

package bad.robot.excel.style;

import bad.robot.excel.cell.DataFormat;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

import static bad.robot.excel.style.ClonedStyleFactory.newStyleBasedOn;
import static org.apache.poi.ss.usermodel.CellStyle.SOLID_FOREGROUND;

import org.apache.poi.hssf.util.HSSFColor;

public class ReplaceExistingStyle implements Style {

    private final DataFormat format;
    private final Alignment alignment;
    private final FontSize fontSize;
    private final FontColour fontColour;
    private final Fill fill;
    private final Border border;
    private final boolean bold;
    private final boolean italic;
    private final int indent;

    /**
     * package protected. use {@link bad.robot.excel.style.StyleBuilder} instead
     */
    ReplaceExistingStyle(Border border, DataFormat format, Alignment alignment, FontSize fontSize, FontColour fontColour, boolean bold, boolean italic, int indent, Fill fill) {
        this.border = border;
        this.format = format;
        this.alignment = alignment;
        this.fontSize = fontSize;
        this.fontColour = fontColour;
        this.fill = fill;
        this.bold = bold;
        this.italic = italic;
        this.indent = indent;
    }

    @Override
    public void applyTo(org.apache.poi.ss.usermodel.Cell cell, Workbook workbook) {
        CellStyle style = newStyleBasedOn(cell).create(workbook);
        applyBorderTo(style);
        applyFillTo(style);
        applyAlignmentTo(style);
        applyDataFormatTo(style, workbook);
        applyFontTo(style, workbook);
        cell.setCellStyle(style);
    }

    private void applyBorderTo(CellStyle style) {
        if (border != null) {
            style.setBorderBottom(border.getBottom().value().getPoiStyle());
            style.setBorderTop(border.getTop().value().getPoiStyle());
            style.setBorderRight(border.getRight().value().getPoiStyle());
            style.setBorderLeft(border.getLeft().value().getPoiStyle());
        }
    }

    private void applyFillTo(CellStyle style) {
        if (fill != null) {
            style.setFillPattern(SOLID_FOREGROUND);
            style.setFillForegroundColor(fill.getForegroundColour().value().getPoiStyle());
        }
    }

    private void applyAlignmentTo(CellStyle style) {
        if (alignment != null)
            style.setAlignment(alignment.value().getPoiStyle());
    }

    private void applyDataFormatTo(CellStyle style, Workbook workbook) {
        if (format != null)
            style.setDataFormat(workbook.createDataFormat().getFormat(format.value()));
    }

    private void applyFontTo(CellStyle style, Workbook workbook) {
       Font font = workbook.createFont();
       font.setColor(fontColour.value().getPoiStyle());
       font.setBold(bold);
       font.setItalic(italic);
       
       if (fontSize != null)
          font.setFontHeightInPoints(fontSize.value());
       
       style.setFont(font);
       
       if (indent > 0)
          style.setIndention((short)indent);
       
      // System.out.println(font.getIndex());
       
       //workbook.getFontAt((short)1);
       
       /*if (fontSize != null) {
            Font font = workbook.createFont();
            font.setFontHeightInPoints(fontSize.value());
            font.setColor(fontColour.value().getPoiStyle());
            style.setFont(font);
        } else {
            // doesn't work
            Font existing = workbook.getFontAt(style.getFontIndex());
            existing.setColor(fontColour.value().getPoiStyle());
            style.setFont(existing);
        }*/
    }
}
