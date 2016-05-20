package com.tries.mergeexcel.xlstoxlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelOperation implements Runnable {

	private List<String> headerList = new ArrayList<String>();
	private List<Car> carList = new ArrayList<Car>();
	private String destinationPath;
	private static final String FILE_TYPE_XLS = ".xls";

	public ExcelOperation(String destinationPath) {
		this.destinationPath = destinationPath;

	}

	public void writeToXlsxFile() throws IOException {
		File checkBeforeFile = new File(destinationPath);
		if (!checkBeforeFile.exists()) {
			// if excel file is not created before create xlsx file before
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Sheet1");
			int rowNumber = 0;
			for (Car car : getCarList()) {
				Row row = sheet.createRow(rowNumber++);
				for (int i = 0; i < 24; i++) {
					setCellFromModel(row, car, i);
				}
				System.out.println("Row Number=" + rowNumber);
			}

			FileOutputStream outputStream = new FileOutputStream(
					destinationPath);
			workbook.write(outputStream);
			outputStream.close();

		} else {
			// if excel file created before open it then take last row
			// continue set values from last row
			// FileWriter fw = new FileWriter(new File(destinationPath));
			// BufferedWriter bw = new BufferedWriter(fw);
			FileInputStream file = new FileInputStream(
					new File(destinationPath));
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			XSSFSheet sheet = workbook.getSheetAt(0);
			int lastRow = sheet.getPhysicalNumberOfRows();
			for (Car car : getCarList()) {
				Row row = sheet.createRow(++lastRow);
				for (int i = 0; i < 24; i++) {
					setCellFromModel(row, car, i);

				}

			}

			file.close();
			FileOutputStream outputStream = new FileOutputStream(
					destinationPath);
			workbook.write(outputStream);
			outputStream.close();

		}

	}

	public void setCellFromModel(Row row, Car car, int columnOrder) {
		Cell cell = row.createCell(columnOrder);
		setCellValue(cell, car, columnOrder);
	}

	private void setCellValue(Cell cell, Car car, int columnOrder) {
		if (columnOrder == 1) {
			cell.setCellValue(car.getModel());
		} else if (columnOrder == 2) {
			cell.setCellValue(car.getVehicleNo());
		} else if (columnOrder == 3) {
			cell.setCellValue(car.getSasiNo());
		} else if (columnOrder == 4) {
			cell.setCellValue(car.getCinsi());
		} else if (columnOrder == 5) {
			cell.setCellValue(car.getMarka());
		} else if (columnOrder == 6) {
			cell.setCellValue(car.getTip());
		} else if (columnOrder == 7) {
			cell.setCellValue(car.getPlakaIl());
		} else if (columnOrder == 8) {
			cell.setCellValue(car.getPlakaNo());
		} else if (columnOrder == 9) {
			cell.setCellValue(car.getKoltukSayısı());
		} else if (columnOrder == 10) {
			cell.setCellValue(car.getSilindirHacmi());
		} else if (columnOrder == 11) {
			cell.setCellValue(car.getMotorGucu());
		} else if (columnOrder == 12) {
			cell.setCellValue(car.getKullanımSekli());
		} else if (columnOrder == 13) {
			cell.setCellValue(car.getIsHurdaArac());
		} else if (columnOrder == 14) {
			cell.setCellValue(car.getRenk());
		} else if (columnOrder == 15) {
			cell.setCellValue(car.getYakitTipi());
		} else if (columnOrder == 16) {
			cell.setCellValue(car.getSahiplikBelgeTarihi());
		} else if (columnOrder == 17) {
			cell.setCellValue(car.getTescilTarihi());
		} else if (columnOrder == 18) {
			cell.setCellValue(car.getTrafiktenCekildimi());
		} else if (columnOrder == 19) {
			cell.setCellValue(car.getSakıncaDurumu());
		} else if (columnOrder == 20) {
			// car.setSorguYeri(value);
			cell.setCellValue(car.getSorguYeri());
		} else if (columnOrder == 21) {
			// car.setEgmModelYili(value);
			cell.setCellValue(car.getEgmModelYili());
		} else if (columnOrder == 22) {
			// car.setEgmMotorNo(value);
			cell.setCellValue(car.getEgmMotorNo());
		} else if (columnOrder == 23) {
			// car.setEgmSasiNo(value);
			cell.setCellValue(car.getEgmSasiNo());
		}

	}

	public void readXlsFile() throws IOException {

		for (int i = 1; i < 15; i++) {
			// get xls File
			FileInputStream file = new FileInputStream(new File(
					"C:\\Users\\sozl657\\Desktop\\Results\\" + i
							+ FILE_TYPE_XLS));
			
            System.out.println("C:\\Users\\sozl657\\Desktop\\Results\\" + i
							+ FILE_TYPE_XLS);
			// get workbook according the filepath
			HSSFWorkbook workbook = new HSSFWorkbook(file);

			// get sheet from workbook
			HSSFSheet sheet = workbook.getSheetAt(0);

			// iterate of rows given the sheet
			Iterator<Row> rowIterator = sheet.iterator();

			while (rowIterator.hasNext()) {
				Row row = rowIterator.next();
				// get cells of row
				Iterator<Cell> cellIterator = row.cellIterator();
				int columnOrder = 0;
				int rowOrder = 0;
				Car car = new Car();
				while (cellIterator.hasNext()) {
					columnOrder++;
					rowOrder++;

					Cell cell = cellIterator.next();
					cell.setCellType(cell.CELL_TYPE_STRING);
					setModelValue(car, cell.getStringCellValue(), columnOrder);

				}
				// first row holds caption values
				if (rowOrder != 1) {
					getCarList().add(car);
				}

			}
		}

	}

	public void run() {
		try {
			readXlsFile();
			writeToXlsxFile();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

	private void setModelValue(Car car, String value, int columnOrder) {
		if (columnOrder == 1) {
			car.setModel(value);
		} else if (columnOrder == 2) {
			car.setVehicleNo(value);
		} else if (columnOrder == 3) {
			car.setSasiNo(value);
		} else if (columnOrder == 4) {
			car.setCinsi(value);
		} else if (columnOrder == 5) {
			car.setMarka(value);
		} else if (columnOrder == 6) {
			car.setTip(value);
		} else if (columnOrder == 7) {
			car.setPlakaIl(value);
		} else if (columnOrder == 8) {
			car.setPlakaNo(value);
		} else if (columnOrder == 9) {
			car.setKoltukSayısı(value);
		} else if (columnOrder == 10) {
			car.setSilindirHacmi(value);
		} else if (columnOrder == 11) {
			car.setMotorGucu(value);
		} else if (columnOrder == 12) {
			car.setKullanımSekli(value);
		} else if (columnOrder == 13) {
			car.setIsHurdaArac(value);
		} else if (columnOrder == 14) {
			car.setRenk(value);
		} else if (columnOrder == 15) {
			car.setYakitTipi(value);
		} else if (columnOrder == 16) {
			car.setSahiplikBelgeTarihi(value);
		} else if (columnOrder == 17) {
			car.setTescilTarihi(value);
		} else if (columnOrder == 18) {
			car.setTrafiktenCekildimi(value);
		} else if (columnOrder == 19) {
			car.setSakıncaDurumu(value);
		} else if (columnOrder == 20) {
			car.setSorguYeri(value);
		} else if (columnOrder == 21) {
			car.setEgmModelYili(value);
		} else if (columnOrder == 22) {
			car.setEgmMotorNo(value);
		} else if (columnOrder == 23) {
			car.setEgmSasiNo(value);
		}

	}

	// set header value
	public void setHeaderValue(String headerValue) {
		getHeaderList().add(headerValue);
	}

	public List<String> getHeaderList() {
		return headerList;
	}

	public List<Car> getCarList() {
		return carList;
	}

}
