package getHtml;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class SheetConverter {

	private XSSFSheet sheet;

	public SheetConverter(String fileName, String sheetName) throws IOException {

		XSSFWorkbook wb = (XSSFWorkbook) WorkbookFactory.create(new File(fileName));
		this.sheet = wb.getSheet(sheetName);
	}

	public void printStyles() {

		Iterator<Row> rows = sheet.rowIterator();

		HashMap<String, String> cellStyles = new HashMap<String, String>();

		while(rows.hasNext()) {
			Iterator<Cell> cells = rows.next().cellIterator();
			while(cells.hasNext()) {
				XSSFCell cell = (XSSFCell)cells.next();
				cellStyles.clear();

				cellStyles.putAll(checkBorder(cell));
				cellStyles.putAll(getCellColorRgb(cell));
				cellStyles.putAll(getFontStyles(cell));

				System.out.println(cellStyles);

			}
		}
	}



	private HashMap<String, String> checkBorder(Cell cell) {

		HashMap<String, String> styles = new HashMap<String, String>();

		CellStyle cellStyle = cell.getCellStyle();

		if(cellStyle.getBorderLeft() != BorderStyle.NONE ) styles.put("border-left", "1px solid black");
		if(cellStyle.getBorderRight() != BorderStyle.NONE ) styles.put("border-right", "1px solid black");
		if(cellStyle.getBorderTop() != BorderStyle.NONE ) styles.put("border-top", "1px solid black");
		if(cellStyle.getBorderBottom() != BorderStyle.NONE ) styles.put("border-bottom", "1px solid black");

		return styles;
	}

	private HashMap<String, String> getCellColorRgb(XSSFCell cell) {

		// 返り値
		HashMap<String, String> styles = new HashMap<String, String>();

		XSSFCellStyle cellStyle = cell.getCellStyle();
		XSSFColor color = cellStyle.getFillForegroundColorColor();

		styles.put("background-color", transferXSSFColorToCssRGBa(color));

		return styles;
	}

	private HashMap<String, String> getFontStyles(XSSFCell cell) {

		// 返り値
		HashMap<String, String> styles = new HashMap<String, String>();

		XSSFCellStyle cellStyle = cell.getCellStyle();
		XSSFFont font = cellStyle.getFont();
		XSSFColor color = font.getXSSFColor();

		styles.put("color", transferXSSFColorToCssRGBa(color));

		return styles;
	}




	final String RGBa_NULL = "RGBa(0, 0, 0, 1.0)";

	private String transferXSSFColorToCssRGBa(XSSFColor color) {

		if (color == null) {
			return RGBa_NULL;
		}

		HashMap<String, Integer> colorARGB = getARGBNumbersFromColor(color);
		return transferARGBMapToCssRGBa(colorARGB);
	}

	private HashMap<String, Integer> getARGBNumbersFromColor(XSSFColor color) {

		HashMap<String, Integer> colorARGB = new HashMap<String, Integer>();

		byte[] bytesColor = color.getARGB();

		colorARGB.put("A", Byte.toUnsignedInt(bytesColor[0]));
		colorARGB.put("R", Byte.toUnsignedInt(bytesColor[1]));
		colorARGB.put("G", Byte.toUnsignedInt(bytesColor[2]));
		colorARGB.put("B", Byte.toUnsignedInt(bytesColor[3]));

		return colorARGB;

	}

	private String transferARGBMapToCssRGBa(HashMap<String, Integer> colorARGB) {

		return String.format("RGBa(%d, %d, %d, %.2f)"
						, colorARGB.get("R")
						, colorARGB.get("G")
						, colorARGB.get("B")
						, (float) colorARGB.get("A") / 255 );
	}


}
