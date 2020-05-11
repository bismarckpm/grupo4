import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import com.sun.corba.se.spi.orbutil.threadpool.Work;
import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

/**
 * Parte B - Primera Actividad DS2020.
 *
 * @author Luis Guerrero.
 *
 */
public class AlumnoB {

    String excelFilePath = "Inventario.xlsx";
    final int MAX_NUMERO_REGISTROS_POR_HOJA = 30;

    /**
     * Determinar si se excedió el número de registros por hoja.
     *
     * @param numHoja - El número de hoja a ser verificado.
     * @return True si fue excedido el límite. False en caso contrario.
     */
    public boolean isMaximoNumeroRegistrosInsertados(int numHoja) {
        boolean maximoExcedido = false;

        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            Sheet sheet = workbook.getSheetAt(numHoja);

            // Basado en índice 0.
            int numeroRegistros = sheet.getLastRowNum() + 1;

            inputStream.close();
            workbook.close();

            maximoExcedido = (numeroRegistros >= MAX_NUMERO_REGISTROS_POR_HOJA);

        } catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
            ex.printStackTrace();
        }

        return maximoExcedido;
    }

    /**
     * Crear una hoja en el Libro.
     */
    public void crearHoja() {
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            // Obtener el número de hojas presentes en el libro.
            int numeroHojas = workbook.getNumberOfSheets();

            Sheet sheet = workbook.createSheet("Java Books " + Integer.toString(numeroHojas));

            Row row = sheet.createRow(0);

            int columnCount = 0;

            Cell cellNo = row.createCell(columnCount);
            cellNo.setCellValue("No");
            Cell cellBookTitle = row.createCell(++columnCount);
            cellBookTitle.setCellValue("Book Title");
            Cell cellAuthor = row.createCell(++columnCount);
            cellAuthor.setCellValue("Author");
            Cell cellPrice = row.createCell(++columnCount);
            cellPrice.setCellValue("Price");

            FileOutputStream outputStream = new FileOutputStream(excelFilePath);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();
            inputStream.close();

        } catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }

    /**
     * Agrega nuevos datos a la hoja de trabajo.
     *
     * @param numeroHoja - Define el número de la hoja en la que se insertarán los datos.
     * @param datos - Los datos a ser insertados.
     */
    public void agregarDatosHoja(int numeroHoja, Object[][] datos) {
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            // Obtener la hoja.
            Sheet sheet = workbook.getSheetAt(numeroHoja);

            // Validaciones.
//            if (isMaximoNumeroRegistrosInsertados(numeroHoja)) {
//                crearHoja();
//                numeroHoja++;
//                sheet = workbook.getSheetAt(numeroHoja);
//            }

            for (Object[] item : datos) {
                int rowCount = sheet.getLastRowNum();

                if (rowCount == 30) {
                    numeroHoja++;

                    // Obtener el número de hojas presentes en el libro.
                    int numeroHojas = workbook.getNumberOfSheets();
                    workbook.createSheet("Java Books " + Integer.toString(numeroHojas));
                    sheet = workbook.getSheetAt(numeroHoja);
                    Row row = sheet.createRow(0);
                    int columnCount = 0;
                    Cell cellNo = row.createCell(columnCount);
                    cellNo.setCellValue("No");
                    Cell cellBookTitle = row.createCell(++columnCount);
                    cellBookTitle.setCellValue("Book Title");
                    Cell cellAuthor = row.createCell(++columnCount);
                    cellAuthor.setCellValue("Author");
                    Cell cellPrice = row.createCell(++columnCount);
                    cellPrice.setCellValue("Price");

                    rowCount = 0;
                }

                Row row = sheet.createRow(++rowCount);

                int columnCount = 0;

                Cell cell = row.createCell(columnCount);
                cell.setCellValue(rowCount);

                for (Object field : item) {
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

            // Mostrar todos los datos.
            mostrarDatosLibro();

        } catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }

    /**
     * Mostrar los datos del libro.
     */
    public void mostrarDatosLibro() {
        try {
            FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
            Workbook workbook = WorkbookFactory.create(inputStream);

            int numeroHojas = workbook.getNumberOfSheets();

            for (int i = 0; i < numeroHojas; i++) {
                Sheet sheet = workbook.getSheetAt(i);

                int rowCount = sheet.getLastRowNum();

                for (int j = 0; j <= rowCount; j++) {
                    Row row = sheet.getRow(j);

                    if (row == null) {
                        continue;
                    }

                    int ultimaColumna = row.getLastCellNum();

                    for (int k = 0; k < ultimaColumna; k++) {
                        if (row.getCell(k) == null) {
                            continue;
                        }

                        CellType tipoCelda = row.getCell(k).getCellTypeEnum();
                        if (tipoCelda.toString() == "NUMERIC") {
                            if (k == 0)
                                System.out.print((int) row.getCell(k).getNumericCellValue() + "\t\t");
                            else
                                System.out.print(row.getCell(k).getNumericCellValue() + "\t\t");
                        }
                        else if (tipoCelda.toString() == "STRING")
                            System.out.print(row.getCell(k).getStringCellValue() + "\t\t");
                    }

                    System.out.println("\n");
                }
            }

            inputStream.close();
            workbook.close();

        } catch (IOException | EncryptedDocumentException | InvalidFormatException ex) {
            ex.printStackTrace();
        }
    }

}
