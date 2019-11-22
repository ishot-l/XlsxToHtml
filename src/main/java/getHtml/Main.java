package getHtml;

import java.io.IOException;

public class Main {
	/*
	 * [ファイル名] [シート名] [座標]
	 *
	 */

	public static void main(String[] args) throws IOException {

		String fileName = args[0];
		String sheetName = args[1];

		SheetConverter sc = new SheetConverter(fileName, sheetName);

		sc.printStyles();

	}


}
