package commenta.excel;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Row.MissingCellPolicy;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;

import commenta.model.CommentInfo;

public class CommentReader {

	public Map<String, List<CommentInfo>> read(String excelPath) throws IOException, EncryptedDocumentException, InvalidFormatException {
		File file = new File(excelPath);
		Map<String, List<CommentInfo>> sheetCommentInfoMap = new LinkedHashMap<>();
		try (Workbook workbook = WorkbookFactory.create(file)) {

			for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
				Sheet sheet = workbook.getSheetAt(i);
				boolean hasComments = true;
				if (sheet instanceof XSSFSheet) {
					hasComments = ((XSSFSheet) sheet).hasComments();
				}
				if (hasComments) {
					List<CommentInfo> commentInfoList = new ArrayList<>();
					sheetCommentInfoMap.put(sheet.getSheetName(), commentInfoList);
					Map<CellAddress, ? extends Comment> map = sheet.getCellComments();
					for (Entry<CellAddress, ? extends Comment> entry : map.entrySet()) {

						CommentInfo commentInfo = new CommentInfo();
						commentInfoList.add(commentInfo);

						commentInfo.setColumnNo(entry.getKey().getColumn() + 1);
						commentInfo.setRowNo(entry.getKey().getRow() + 1);
						commentInfo.setCommentAuthor(entry.getValue().getAuthor());
						commentInfo.setComment(entry.getValue().getString().getString()
								.replaceFirst(entry.getValue().getAuthor() + ":\n", ""));

						Row row = sheet.getRow(entry.getKey().getRow());

						if (row != null) {
							Cell cell = row.getCell(entry.getKey().getColumn(), MissingCellPolicy.CREATE_NULL_AS_BLANK);
							switch (cell.getCellType()) {
							case Cell.CELL_TYPE_BLANK: {
								commentInfo.setCommentedText("");
								break;
							}
							case Cell.CELL_TYPE_BOOLEAN: {
								commentInfo.setCommentedText(String.valueOf(cell.getBooleanCellValue()));
								break;
							}
							case Cell.CELL_TYPE_ERROR: {
								commentInfo.setCommentedText(String.valueOf(cell.getErrorCellValue()));
								break;
							}
							case Cell.CELL_TYPE_FORMULA: {
								commentInfo.setCommentedText(cell.getCellFormula());
								break;
							}
							case Cell.CELL_TYPE_NUMERIC: {
								commentInfo.setCommentedText(String.valueOf(cell.getNumericCellValue()));
								break;
							}
							case Cell.CELL_TYPE_STRING: {
								commentInfo.setCommentedText(cell.getStringCellValue());
								break;
							}
							}
						} else {
							commentInfo.setCommentedText("");

						}
					}
				}
			}
		}
		return sheetCommentInfoMap;
	}
}