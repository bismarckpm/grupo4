import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import java.io.File;
import java.io.FileOutputStream;

/////
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

/*Descripción: Como Usuario quiero validar si el archivo Excel existe
para poder ejecutar el programa sin error.
 */
public class HistoriaA {

    public int  existeArchivo(File archivo){

        if (archivo.exists()){
            System.out.println("El archivo existe");
            return  1;
        }
        else {
         System.out.println("El archivo no existe");
         return  0;}
    }
    public void  crearArchivo(){
        // Creamos el archivo donde almacenaremos la hoja
        // de calculo, recuerde usar la extension correcta,
        // en este caso .xlsx
        File archivo = new File("Inventario.xlsx");
        // Creamos el libro de trabajo de Excel formato OOXML
        Workbook workbook = new XSSFWorkbook();
        // La hoja donde pondremos los datos
        Sheet pagina = workbook.createSheet("JavaBooks");
        String[] titulos = {"No", "Book Title",
                "Author", "Price"};
        // Creamos una fila en la hoja en la posicion 0
        Row fila = pagina.createRow(0);
        // Creamos el encabezado
        for (int i = 0; i < titulos.length; i++) {
            // Creamos una celda en esa fila, en la posicion
            // indicada por el contador del ciclo
            Cell celda = fila.createCell(i);
            celda.setCellValue(titulos[i]);
        }
        // Ahora guardaremos el archivo
        try {
            // Creamos el flujo de salida de datos,
            // apuntando al archivo donde queremos
            // almacenar el libro de Excel
            FileOutputStream salida = new FileOutputStream(archivo);

            // Almacenamos el libro de
            // Excel via ese
            // flujo de datos
            workbook.write(salida);

            // Cerramos el libro para concluir operaciones
            workbook.close();

            //LOGGER.log(Level.INFO, "Archivo creado existosamente en {0}", archivo.getAbsolutePath());

        } catch (FileNotFoundException ex) {
           // LOGGER.log(Level.SEVERE, "Archivo no localizable en sistema de archivos");
        } catch (IOException ex) {
            //LOGGER.log(Level.SEVERE, "Error de entrada/salida");
        }
        System.out.println("Se crea el archivo Inventario.xls");
    }
    public void mostrarArchivosExistentes(){


    }

    public HistoriaA() {
    }
}
