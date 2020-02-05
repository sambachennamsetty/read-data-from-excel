package com.app.config;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.Map;
import java.util.Set;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class CompareXLSheet {
	public static void main(String[] args) throws EncryptedDocumentException, IOException {

		Row file1row = null;
		Cell file1cell = null;

		Row file2row = null;
		Cell file2cell = null;

		Row file3row = null;
		Cell file3cell = null;
		Cell file3cellError = null;

		Set<String> auditData = new HashSet<>();
		Map<String, String> errorData = new HashMap<>();
		Set<String> mergeData = new HashSet<>();

		try (InputStream file1inp = new FileInputStream("C:\\Users\\ADMIN\\Downloads\\New folder\\re.xls");
				Workbook file1wb = WorkbookFactory.create(file1inp);
				InputStream file2inp = new FileInputStream("C:\\Users\\ADMIN\\Downloads\\New folder\\Audit.xls");
				Workbook file2wb = WorkbookFactory.create(file2inp);
				InputStream file3inp = new FileInputStream("C:\\Users\\ADMIN\\Downloads\\New folder\\Error.xls");
				Workbook file3wb = WorkbookFactory.create(file3inp);

		) {

			Sheet file1sheet = file1wb.getSheetAt(0);
			Sheet file2sheet = file2wb.getSheetAt(0);
			Sheet file3sheet = file3wb.getSheetAt(0);

			for (int i = 0; i < file1sheet.getPhysicalNumberOfRows(); i++) {
				file1row = file1sheet.getRow(i);
				file1cell = file1row.getCell(2);
				for (int j = 0; j < file2sheet.getPhysicalNumberOfRows(); j++) {
					file2row = file2sheet.getRow(j);
					file2cell = file2row.getCell(20);
					if (file1cell != null && file2cell != null && file2cell.toString().contains(file1cell.toString())) {
						auditData.add(file1cell.toString());
					}

				}
				for (int k = 0; k < file3sheet.getPhysicalNumberOfRows(); k++) {
					file3row = file3sheet.getRow(k);
					file3cell = file3row.getCell(21);
					file3cellError = file3row.getCell(16);
					if (file1cell != null && file3cell != null && file3cell.toString().contains(file1cell.toString())) {

						errorData.put(file1cell.toString(), file3cellError.toString());
					}

				}
			}

		}

		Set<String> keys = errorData.keySet();
		System.out.println("**************************");
		for (Iterator<String> i = keys.iterator(); i.hasNext();) {
			String key = i.next();
			String value = errorData.get(key);
			System.out.println(key + " = " + value);
		}

		System.out.println(Thread.currentThread().getStackTrace().toString());
		System.out.println(Arrays.toString(Thread.currentThread().getStackTrace()));

	}

}
