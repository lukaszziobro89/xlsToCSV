import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.math.BigDecimal;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Collectors;

public class Main {
    public static void main(String[] args) throws IOException {

        // get current directory
        String currentDirectory = System.getProperty("user.dir");

        // list of files in current directory
        List<File> filesInFolder = Files.walk(Paths.get(currentDirectory))
                .filter(Files::isRegularFile)
                .map(Path::toFile)
                .collect(Collectors.toList());

        // for each file in directory
        for (File file : filesInFolder){
            String filename = file.getName();
            // if file is xls or xlsx
            if (filename.endsWith(".xls") || filename.endsWith(".xlsx")){
                System.out.println(filename);
                excelToCSV(currentDirectory, filename);
            }
        }
    }

    private static void excelToCSV(String inputFilePath, String filename) throws IOException {
        try {
            // declare all variables
            OPCPackage fs = OPCPackage.open(new File(inputFilePath + "\\" + filename));
            XSSFWorkbook wb = new XSSFWorkbook(fs);
            XSSFSheet sheet;
            XSSFRow row;
            XSSFCell cell;

            // count number of sheets in workbook
            int numberOfSheets = wb.getNumberOfSheets();

            // for each sheet in workbook
            for (int j = 0; j < numberOfSheets; j++) {

                String sheetName = wb.getSheetAt(j).getSheetName();
                sheet = wb.getSheetAt(j);

                FileWriter writer = new FileWriter(inputFilePath + "\\" + sheetName +".csv");

                int numberOfRows; // No of rows
                numberOfRows = sheet.getPhysicalNumberOfRows();

                int numberOfColumns = 0; // No of columns
                int tmp;

                // This trick ensures that we get the data properly even if it doesn't start from first few rows
                for (int i = 0; i < 10 || i < numberOfRows; i++) {
                    row = sheet.getRow(i);
                    if (row != null) {
                        tmp = sheet.getRow(i).getPhysicalNumberOfCells();
                        if (tmp > numberOfColumns) numberOfColumns = tmp;
                    }
                }

                StringBuilder sb = new StringBuilder();

                for (int r = 0; r < numberOfRows; r++) {
                    row = sheet.getRow(r);
                    if (row != null) {
                        for (int c = 0; c < numberOfColumns; c++) {
                            cell = row.getCell((short) c);
                            if (cell != null) {
                                if (cell.getCellType().equals(CellType.NUMERIC)) {
                                    BigDecimal bDec = new BigDecimal(cell.getNumericCellValue());
                                    String numeric = String.valueOf(bDec);
                                    sb.append(numeric).append(",");
                                } else {
                                    sb.append(cell).append(",");
                                }
                            }
                        }
                        String line = sb.toString().substring(0, sb.length() - 1);
                        sb.setLength(0);
                        writer.append(line).append(String.valueOf('\n'));
                    }
                }

                writer.flush();
                writer.close();
            }
        } catch (IOException | InvalidFormatException e) {
            FileWriter errorWriter = new FileWriter(inputFilePath + "\\" + "error.txt");
            errorWriter.write(e.getMessage());
            errorWriter.flush();
            errorWriter.close();
            e.printStackTrace();
        }
    }
}
