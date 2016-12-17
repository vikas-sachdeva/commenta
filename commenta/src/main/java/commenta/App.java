package commenta;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import commenta.excel.CommentReader;
import commenta.model.CommentInfo;
import commenta.model.ExcelObj;
import commenta.output.ExcelOutput;

public class App {

	private static final String EXCEL_PATH = "H:/inputFile.xls";

	public static void main(String[] args) throws IOException, EncryptedDocumentException, InvalidFormatException {
		ExcelObj.getInstance().setFilePath(EXCEL_PATH);

		Map<String, List<CommentInfo>> sheetCommentInfoMap = new CommentReader()
				.read(ExcelObj.getInstance().getFilePath());

		new ExcelOutput().output(sheetCommentInfoMap, ExcelObj.getInstance().getFilePath());

	}
}