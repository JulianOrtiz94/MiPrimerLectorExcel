package reader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import org.openqa.selenium.By;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Main {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
		
		try {
			System.setProperty("webdriver.chrome.driver", "C:\\SELENIUM\\chromedriver.exe");
			WebDriver driver = new ChromeDriver();
			
			FileInputStream inputStream = new FileInputStream(new File ("C:\\SELENIUM\\Prueba.xlsx"));
			Workbook workbook =  WorkbookFactory.create(inputStream);
			Sheet primeraHoja = workbook.getSheetAt(0);
			Iterator iterador = primeraHoja.iterator();
			
			DataFormatter formatter = new DataFormatter();
            while (iterador.hasNext()) {
                Row fila = (Row) iterador.next();
                //System.out.println("row number : " + fila.getRowNum());
                Iterator cellIterator = fila.cellIterator();
                
                if( fila.getRowNum() != 0 ) {
                	 while(cellIterator.hasNext()) {
                         Cell celda = (Cell) cellIterator.next();
                         //System.out.println("column index: " + celda.getColumnIndex());
                         String contenidoCeldaBusqueda = formatter.formatCellValue(fila.getCell(0));
                         //System.out.println("celda: " + contenidoCeldaBusqueda);
                         driver.get("http:\\www.google.com.uy");
                 		 driver.findElement(By.name("q")).sendKeys(contenidoCeldaBusqueda);;
                 		 driver.findElement(By.name("btnK")).submit();
                 		 WebElement element = (new WebDriverWait(driver, 10))
              				  .until(ExpectedConditions.presenceOfElementLocated((By.id("search"))));
                 		 		
                 		 //System.out.println("Element size: " + element.getSize());
                 		 Dimension sizeContenedorResultado =  element.getSize();
                 		 //System.out.println("Element size Height: " + sizeContenedorResultado.getHeight());
                 		 Boolean sinResultado = driver.findElements(By.xpath("//p[@aria-level='3']")).size() > 0 ;
                 		 //System.out.println("Element exist: " + sinResultado);
                 		 if(sinResultado && (sizeContenedorResultado.getHeight() == 0)) {
                 			fila.createCell(1).setCellValue("No");
                 		 }else {
                 			fila.createCell(1).setCellValue("Si");
                 		 }
                 		 
                 		
                     }
                }	
            }
            inputStream.close();
     		FileOutputStream outputStream =new FileOutputStream(new File("C:\\SELENIUM\\Prueba.xlsx"));
     		workbook.write(outputStream);
	        outputStream.close();
	        System.out.println("Done");
		} catch (Exception e) {
            e.printStackTrace();
        }
	}

}
