
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;

public class Sorter {

    private InputStream inp;
    private String filename;
    private Workbook wb;

    public Sorter(String filename) {
        try {
            inp = new FileInputStream(filename);
            this.filename = filename;
            wb = WorkbookFactory.create(inp);
        } catch (IOException | EncryptedDocumentException | InvalidFormatException e) {
        }
    }

    public String[] getSheets() {
        String[] sheets = new String[wb.getNumberOfSheets()];
        for (int s = 0; s < wb.getNumberOfSheets(); s++) {
            sheets[s] = wb.getSheetAt(s).getSheetName();
        }
        return sheets;
    }
    
    public String getFilename() {
        return filename;
    }

    public void sort(int sheetNum) throws IOException {
        try {
            Sheet sheet1 = wb.getSheetAt(sheetNum);

            List<Long> inscritos = new ArrayList<>(50);
            Iterator<Row> i = getIterator(sheet1);
            while (i.hasNext()) {
                Row next = i.next();
                Cell c = next.getCell(7);
                if (c != null) {
                    if (c.getRichStringCellValue().getString().equalsIgnoreCase("X")) {
                        inscritos.add((long) next.getCell(0).getNumericCellValue());
                    }
                }
            }

            Collections.shuffle(inscritos);

            for (int row = 5; row < sheet1.getPhysicalNumberOfRows(); row++) {
                Row r = sheet1.getRow(row);
                Cell c = r.getCell(10);
                if (row - 5 < inscritos.size()) {
                    c.setCellValue(inscritos.get(row - 5));
                } else if (c != null) {
                    c.setCellValue("");
                }
            }

            XSSFFormulaEvaluator.evaluateAllFormulaCells(wb);
            FileOutputStream fileOut = new FileOutputStream(filename);
            wb.write(fileOut);
            fileOut.close();

        } catch (EncryptedDocumentException e) {
        }
    }

    private Iterator<Row> getIterator(Sheet sheet) {
        Iterator<Row> i = sheet.rowIterator();
        i.next();
        i.next();
        i.next();
        i.next();
        i.next();
        return i;
    }
}
