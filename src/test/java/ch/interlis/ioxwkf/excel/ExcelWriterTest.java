package ch.interlis.ioxwkf.excel;

import java.io.File;
import java.io.FileInputStream;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import ch.interlis.iom.IomObject;
import ch.interlis.iom_j.Iom_jObject;
import ch.interlis.iox.IoxException;
import ch.interlis.iox_j.EndBasketEvent;
import ch.interlis.iox_j.EndTransferEvent;
import ch.interlis.iox_j.ObjectEvent;
import ch.interlis.iox_j.StartBasketEvent;
import ch.interlis.iox_j.StartTransferEvent;

public class ExcelWriterTest {
    
    private static final String TEST_IN="src/test/data/ExcelWriter/";
    private static final String TEST_OUT="build/test/data/ExcelWriter";


    @BeforeAll
    public static void setupFolder() {
        new File(TEST_OUT).mkdirs();
    }

    @Test
    public void attributes_no_description_set_Ok() throws Exception {
        // Prepare
        File parentDir = new File(TEST_OUT, "attributes_no_description_set_Ok");
        parentDir.mkdirs();

        Iom_jObject inputObj = new Iom_jObject("Test1.Topic1.Obj1", "o1");
        inputObj.setattrvalue("id1", "1");
        inputObj.setattrvalue("aText", "text1");
        inputObj.setattrvalue("aDouble", "53434.123");
//        IomObject coordValue = inputObj.addattrobj("attrPoint", "COORD");
//        {
//            coordValue.setattrvalue("C1", "2600000.000");
//            coordValue.setattrvalue("C2", "1200000.000");
//        }

        // Run
        ExcelWriter writer = null;
        File file = new File(parentDir,"attributes_no_description_set_Ok.xlsx");
        try {
            writer = new ExcelWriter(file);
            writer.write(new StartTransferEvent());
            writer.write(new StartBasketEvent("Test1.Topic1","bid1"));
            writer.write(new ObjectEvent(inputObj));
            writer.write(new EndBasketEvent());
            writer.write(new EndTransferEvent());
        } catch(IoxException e) {
            throw new IoxException(e);
        } finally {
            if(writer != null) {
                try {
                    writer.close();
                } catch (IoxException e) {
                    throw new IoxException(e);
                }
                writer=null;
            }
        }

        // Validate
        FileInputStream fis = new FileInputStream(file);        
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);

        Row headerRow = sheet.getRow(0);
        
        Iterator<Row> rowIterator = sheet.iterator();
        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();

            while (cellIterator.hasNext()) {
                Cell cell = cellIterator.next();
                System.out.println(cell.getStringCellValue());
            } 

        }
        
        workbook.close();
        fis.close();

    }
}
