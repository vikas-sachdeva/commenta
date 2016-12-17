package commenta.output;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class EmptyWorkbookFactory {

	public static Workbook getEmptyWorkbook(String filePath) {
		if (filePath.endsWith(".xlsx")) {
			return new XSSFWorkbook();
		} else if (filePath.endsWith(".xls")) {
			return new HSSFWorkbook();
		} else {
			return null;
		}
	}
}
