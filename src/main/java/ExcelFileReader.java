import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.*;

public class ExcelFileReader {

    private String filePath;
    private Set<String> set = new TreeSet<>(Comparator.naturalOrder());

    public ExcelFileReader() {
    }

    public void readFile(String filePath) {

        try (FileInputStream fis = new FileInputStream(filePath);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            boolean pominNaglowek = true;

            for (Row row : sheet) {
                if (pominNaglowek) {
                    pominNaglowek = false;
                    continue;
                }
                switch (row.getCell(0).getCellType()) {
                    case STRING:
                        String result = row.getCell(0).getStringCellValue();
                        set.add(result);
                        break;
                    case NUMERIC:
                        String result2 = String.valueOf((long) (row.getCell(0).getNumericCellValue()));
                        set.add(result2);
                        break;
                }
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public void printInfo() {
        System.out.println(set);
        System.out.println(set.size());
    }

}
