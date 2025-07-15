public class Runner {

    public static void main(String[] args) {


            String filePath = "C:\\Users\\Adriano_Diag\\Downloads\\Oferta specjalna_Listopad.xlsx";

            ExcelFileReader excelFileReader = new ExcelFileReader();

            excelFileReader.readFile(filePath);

            excelFileReader.printInfo();

    }

}
