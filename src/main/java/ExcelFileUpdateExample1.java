import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

import javax.swing.*;

/**
 * This program illustrates how to update an existing Microsoft Excel document.
 * Append new rows to an existing sheet.
 * 
 * @author www.codejava.net
 *
 */
public class ExcelFileUpdateExample1 {


	public static void main(String[] args) {
		int opc = Integer.parseInt(JOptionPane.showInputDialog(null,"Menu\n"
		+ "1: Correr prograna\n"
		+ "2: Validar archivo existente\n"
		+ "3: Cantidad de registros por hoja\n"
		+ "4: Actualizar registro\n"
		+ "Presione Cancel para salir"));

		do{
			switch (opc){
				case 1:{
					try {
						String excelFilePath = "Inventario.xlsx";
						FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
						Workbook workbook = WorkbookFactory.create(inputStream);

						Sheet sheet = workbook.getSheetAt(0);

						Object[][] bookData = {
								{"El que se duerme pierde", "Tom Peter", 16},
								{"Sin lugar a duda", "Ana Gutierrez", 26},
								{"El arte de dormir", "Nico", 32},
								{"Buscando a Nemo", "Humble Po", 41},
						};

						int rowCount = sheet.getLastRowNum();

						for (Object[] aBook : bookData) {
							Row row = sheet.createRow(++rowCount);

							int columnCount = 0;

							Cell cell = row.createCell(columnCount);
							cell.setCellValue(rowCount);

							for (Object field : aBook) {
								cell = row.createCell(++columnCount);
								if (field instanceof String) {
									cell.setCellValue((String) field);
								} else if (field instanceof Integer) {
									cell.setCellValue((Integer) field);
								}
							}

						}

						inputStream.close();

						FileOutputStream outputStream = new FileOutputStream(excelFilePath);
						workbook.write(outputStream);
						workbook.close();
						outputStream.close();

					} catch (IOException | EncryptedDocumentException
							| InvalidFormatException ex) {
						ex.printStackTrace();
					}
					break;
				}

				case 2:{
					break;
				}

				case 3:{
					break;
				}

				case 4:{
					HistoriaC();
					break;
				}
			}
		}while (opc < 4);


		HistoriaC();
	}

	public static void HistoriaC(){
		String numReg = JOptionPane.showInputDialog("Ingrese el numero del Registro");
		int numRegint = Integer.parseInt(numReg);
		boolean validacion = false;

		String excelFilePath = "Inventario.xlsx";
		try {
			FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
			Workbook workbook = WorkbookFactory.create(inputStream);
			Sheet sheet = workbook.getSheetAt(0);



			for (int i = 1; i <= sheet.getLastRowNum(); i++) {
				if ((int) sheet.getRow(i).getCell(0).getNumericCellValue() == numRegint) {
					String confirmar = JOptionPane.showInputDialog("Desea modificar el autor? Si, No");
					validacion = true;

					if (confirmar.equals("si")) {
						String autor = JOptionPane.showInputDialog("Ingrese el nuevo nombre del autor");
						sheet.getRow(i).getCell(2).setCellValue(autor);
					}
					String confirmar2 = JOptionPane.showInputDialog("Desea modificar el precio? Si, No");

					if (confirmar2.equals("si")) {
						String precio = JOptionPane.showInputDialog("Ingrese el nuevo precio");
						int precioint = Integer.parseInt(precio);
						sheet.getRow(i).getCell(3).setCellValue(precioint);
					}
				}
			}
			if (!validacion){
				JOptionPane.showMessageDialog(null,"El registro no existe");
			}else {

				for (Row row : sheet) {
					for (Cell currentCell : row) {
						if (currentCell.getCellTypeEnum() == CellType.STRING) {
							System.out.print(currentCell.getStringCellValue() + "  ");
						} else if (currentCell.getCellTypeEnum() == CellType.NUMERIC) {
							System.out.print(currentCell.getNumericCellValue() + "  ");
						}
					}
					System.out.println();

				}
			}

			inputStream.close();
			FileOutputStream outputStream = new FileOutputStream(excelFilePath);
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();

		} catch (InvalidFormatException | IOException e) {
			e.printStackTrace();
		}

	}
}
