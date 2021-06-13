package peritos.ascobom.poc;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TestePOI {
	
	public void executar() throws FileNotFoundException, IOException {
		//Instanciando um objeto ligado ao arquivo xlsx, conforme desejado
        String arquivo = "/home/hud/Desktop/consulta.xlsx";
        XSSFWorkbook wb = new XSSFWorkbook(new FileInputStream(arquivo));

        int totalSheet =  wb.getNumberOfSheets();
        for (int i = 0; i < totalSheet; i++) {
            System.out.println(wb.getSheetName(i));
        }
        //Buscando a primeira planilhado do arquivo,índice zero
        Sheet sheet1 = wb.getSheetAt(0);

        wb.close();		
	}

}
