package commenta.output;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import commenta.model.CommentInfo;

public class ExcelOutput {

	public void output(Map<String, List<CommentInfo>> sheetCommentInfoMap, String excelPath) throws IOException {
		File excelFile = new File(excelPath);
		String sheetName = excelFile.getName();
		String outputFilePath = excelFile.getParent() + File.separator + "Comments_" + sheetName;
		try (Workbook workbook = EmptyWorkbookFactory.getEmptyWorkbook(outputFilePath)) {
			Sheet worksheet = workbook.createSheet(sheetName.replace(".xlsx", "").replace(".xls", ""));
			printHeaderRow(worksheet);
			createDataRows(worksheet, sheetCommentInfoMap);
			try (FileOutputStream fileOutputStream = new FileOutputStream(new File(outputFilePath))) {
				workbook.write(fileOutputStream);
			}
		}
	}

	private void printHeaderRow(Sheet worksheet) {
		Row row = worksheet.createRow(0);

		Cell cell = row.createCell(0);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("S. No.");

		cell = row.createCell(1);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("Sheet Name");

		cell = row.createCell(2);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("Commented Text");

		cell = row.createCell(3);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("Comment");

		cell = row.createCell(4);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("Comment Author");

		cell = row.createCell(5);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("Row");

		cell = row.createCell(6);
		cell.setCellType(Cell.CELL_TYPE_STRING);
		cell.setCellValue("Column");

	}

	private void createDataRows(Sheet worksheet, Map<String, List<CommentInfo>> sheetCommentInfoMap) {

		int rowCounter = 1;
		Row row;
		Cell cell;
		for (Entry<String, List<CommentInfo>> sheetCommentInfoEntry : sheetCommentInfoMap.entrySet()) {
			for (int i = 0; i < sheetCommentInfoEntry.getValue().size(); i++) {
				CommentInfo commentInfo = sheetCommentInfoEntry.getValue().get(i);
				row = worksheet.createRow(rowCounter);

				cell = row.createCell(0);
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(rowCounter++);

				cell = row.createCell(1);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(sheetCommentInfoEntry.getKey());

				cell = row.createCell(2);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(commentInfo.getCommentedText());

				cell = row.createCell(3);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(commentInfo.getComment());

				cell = row.createCell(4);
				cell.setCellType(Cell.CELL_TYPE_STRING);
				cell.setCellValue(commentInfo.getCommentAuthor());

				cell = row.createCell(5);
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(commentInfo.getRowNo());

				cell = row.createCell(6);
				cell.setCellType(Cell.CELL_TYPE_NUMERIC);
				cell.setCellValue(commentInfo.getColumnNo());

			}
		}
	}
}