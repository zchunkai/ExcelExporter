package com.aife.feais.business.utils;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelExporter {

	public static void main(String[] args) {
		// 模板文件路径
		String templatePath = "D:\\abc.xlsx";
		// 导出文件路径
		String exportPath = "D:\\export.xlsx";

		try (FileInputStream fis = new FileInputStream(templatePath);
			 Workbook workbook = new XSSFWorkbook(fis)) {

			// 获取第一个工作表
			Sheet sheet = workbook.getSheetAt(0);
			// 设置单元格边框样式
			CellStyle style = sheet.getWorkbook().createCellStyle();
			style.setBorderTop(BorderStyle.THIN);
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);

			// 在第一个工作表的第一行第一列写入数据
			// 在第一个工作表的第一行第二列写入数据
			writeData(sheet, 0, 0, "Hello, World!",style);
			writeData(sheet, 0, 1, "Hello, World!",style);


			// 获取第二个工作表
			Sheet sheet2 = workbook.getSheetAt(1);
			workbook.setSheetName(workbook.getSheetIndex(sheet2), "NewSheetName");

			//创建第三个
			Sheet sheet3 = workbook.createSheet("NewSheetName2");
			// 复制第一个工作表的内容到新的工作表
			copySheet(sheet, sheet3);


			// 将修改后的工作簿写入文件
			try (FileOutputStream fos = new FileOutputStream(exportPath)) {
				workbook.write(fos);
			}

			System.out.println("数据成功导出到 " + exportPath);

		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public static void writeData(Sheet sheet, int rowNumber, int columnNumber, String data,CellStyle style) {
		Row row = sheet.getRow(rowNumber);
		if (row == null) {
			row = sheet.createRow(rowNumber);
		}
		Cell cell = row.getCell(columnNumber);
		if (cell == null) {
			cell = row.createCell(columnNumber);
		}
		cell.setCellValue(data);

		if (style!=null){
			cell.setCellStyle(style);
		}

	}

	public static void copySheet(Sheet sourceSheet, Sheet targetSheet) {
		for (int i = 0; i <= sourceSheet.getLastRowNum(); i++) {
			Row sourceRow = sourceSheet.getRow(i);
			Row targetRow = targetSheet.createRow(i);

			if (sourceRow != null) {
				copyRow(sourceRow, targetRow, sourceSheet.getWorkbook(), targetSheet.getWorkbook());
			}
		}

		// 复制列宽
		for (int i = 0; i < sourceSheet.getRow(0).getLastCellNum(); i++) {
			targetSheet.setColumnWidth(i, sourceSheet.getColumnWidth(i));
		}

		// 复制合并单元格
		for (int i = 0; i < sourceSheet.getNumMergedRegions(); i++) {
			CellRangeAddress mergedRegion = sourceSheet.getMergedRegion(i);
			targetSheet.addMergedRegion(mergedRegion);
		}
	}

	private static void copyRow(Row sourceRow, Row targetRow, Workbook sourceWorkbook, Workbook targetWorkbook) {
		// 复制行高
		targetRow.setHeight(sourceRow.getHeight());
		for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
			Cell sourceCell = sourceRow.getCell(i);
			Cell targetCell = targetRow.createCell(i);

			if (sourceCell != null) {
				copyCell(sourceCell, targetCell, sourceWorkbook, targetWorkbook);
			}
		}
	}

	private static void copyCell(Cell sourceCell, Cell targetCell, Workbook sourceWorkbook, Workbook targetWorkbook) {
		// 复制内容
		switch (sourceCell.getCellType()) {
			case STRING:
				targetCell.setCellValue(sourceCell.getStringCellValue());
				break;
			case NUMERIC:
				targetCell.setCellValue(sourceCell.getNumericCellValue());
				break;
			case BOOLEAN:
				targetCell.setCellValue(sourceCell.getBooleanCellValue());
				break;
			case FORMULA:
				targetCell.setCellFormula(sourceCell.getCellFormula());
				break;
			case BLANK:
				targetCell.setCellType(CellType.BLANK);
				break;
			case ERROR:
				targetCell.setCellErrorValue(sourceCell.getErrorCellValue());
				break;
			default:
				break;
		}

		// 复制样式
		CellStyle newCellStyle = targetWorkbook.createCellStyle();
		newCellStyle.cloneStyleFrom(sourceCell.getCellStyle());
		targetCell.setCellStyle(newCellStyle);
	}
}
