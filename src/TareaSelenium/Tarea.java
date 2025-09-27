package TareaSelenium;

import org.apache.poi.sl.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.support.ui.Select;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Driver;
import java.util.*;

public class Tarea {
    public static void main(String[] args) throws IOException {
        //setting the excel files
        String ExcelPath = ".\\datafiles\\EscuelaPrimariaLosSantos.xlsx";
        FileInputStream input = new FileInputStream(ExcelPath);

        XSSFWorkbook workbook = new XSSFWorkbook(input);
        XSSFSheet sheet = workbook.getSheetAt(0);

        ArrayList<String> info = new ArrayList<String>();

        //selenium start
        System.setProperty("webdriver.gecko.driver", ".\\drivers\\geckodriver.exe");
        WebDriver drive = new FirefoxDriver();
        drive.get("http://127.0.0.1:5500/index.html"); //Open the test page

        //Starting the excel 
        int fila = sheet.getLastRowNum();
        int columna = sheet.getRow(1).getLastCellNum();
        int cedad;

        String mtr = "";

        //lists
        ArrayList<String> apellidos =  new ArrayList<>();
        ArrayList<String> nombres =  new ArrayList<>();

        ArrayList<String> edad =  new ArrayList<>();
        ArrayList<String> sexo =  new ArrayList<>();

        ArrayList<String> situacion =  new ArrayList<>();
        ArrayList<String> fecha = new ArrayList<>();

        ArrayList<String> ciclo =  new ArrayList<>();
        ArrayList<String> grado =  new ArrayList<>();

        ArrayList<String> direccion = new ArrayList<>();
        ArrayList<String> telefono = new ArrayList<>();

        ArrayList<String> matricula = new ArrayList<>();

        fecha.add("Fecha");
        matricula.add("Matricula");

        //Loops
        for (int i = 0; i <= fila; i++){
            XSSFRow Crow = sheet.getRow(i);

            for(int j = 0; j < columna; j++){
                XSSFCell cell = Crow.getCell(j);

                switch (cell.getCellType()) {
                    case STRING -> info.add(cell.getStringCellValue());
                    case NUMERIC -> {
                        cedad = (int) cell.getNumericCellValue();
                        info.add(Integer.toString(cedad));
                    }
                }
            }
        }

        for(int i = 0; i < info.size(); i+=9){
            apellidos.add(info.get(i));
            nombres.add(info.get(i+1));

            edad.add(info.get(i+2));
            sexo.add(info.get(i+3));

            situacion.add((info.get(i+4)));
            ciclo.add(info.get(i+5));
            grado.add(info.get(i+6));

            direccion.add(info.get(i+7));
            telefono.add(info.get(i+8));
        }

        for(int i = 0; i < apellidos.size();i++) {
            switch (grado.get(i)) {
                case "Primero" -> mtr = apellidos.get(i).charAt(0) + nombres.get(i).charAt(0) + "100" + i;
                case "Segundo" -> mtr = apellidos.get(i).charAt(0) + nombres.get(i).charAt(0) + "200" + i;
                case "Tercero" -> mtr = apellidos.get(i).charAt(0) + nombres.get(i).charAt(0) + "300" + i;
                case "Cuarto" -> mtr = apellidos.get(i).charAt(0) + nombres.get(i).charAt(0) + "400" + i;
                case "Quinto" -> mtr = apellidos.get(i).charAt(0) + nombres.get(i).charAt(0) + "500" + i;
                case "Sexto" -> mtr = apellidos.get(i).charAt(0) + nombres.get(i).charAt(0) + "600" + i;
            }
            matricula.add(mtr);
        }

        for(int i = 0; i < grado.size();i++) {
            switch (grado.get(i)) {
                case "Primero":
                    fecha.add("24/06/2022");
                    break;
                case "Segundo":
                    if(situacion.get(i) == "Transferido"){
                        fecha.add("01/07/2022");
                   }else{
                        fecha.add("23/06/2022");
                   }
                   break;
                case "Tercero":
                   if(situacion.get(i) == "Transferido"){
                        fecha.add("29/06/2022");
                   }else{
                        fecha.add("22/06/2022");
                   }
                   break;
                case "Cuarto":
                    if(situacion.get(i) == "Transferido"){
                        fecha.add("27/06/2022");
                    }else{
                        fecha.add("22/06/2022");
                    }
                    break;
                case "Quinto":
                    if(situacion.get(i) == "Transferido"){
                        fecha.add("25/06/2022");
                    }else{
                        fecha.add("21/06/2022");
                    }
                    break;
                case "Sexto":
                    if(situacion.get(i) == "Transferido"){
                        fecha.add("22/06/2022");
                    }else{
                        fecha.add("20/06/2022");
                    }
                    break;
                //profesor.domingo@gmail.com
               }
        }

        XSSFSheet hoja = workbook.createSheet("Hoja 2");

        Map<String, Object[]> datos = new TreeMap<>();
        int in = 0;
        for(int i = 0; i < 200; i++) {
            datos.put(Integer.toString(i),
                    new Object[]{apellidos.get(i), nombres.get(i),edad.get(i),sexo.get(i),situacion.get(i),
                    ciclo.get(i),grado.get(i), direccion.get(i),telefono.get(i),fecha.get(i), matricula.get(i)});
            in++;
        }

        Set<String> keyset = datos.keySet();

        int rownum = 0;

        for (String key : keyset) {
            Row frow = hoja.createRow(rownum++);

            Object[] objArr = datos.get(key);

            int cellnum = 0;

            for (Object obj : objArr) {

                Cell cell = frow.createCell(cellnum++);

                cell.setCellValue((String) obj);

            }
        }

        apellidos.clear();
        nombres.clear();

        edad.clear();
        sexo.clear();

        situacion.clear();
        ciclo.clear();
        grado.clear();

        direccion.clear();
        telefono.clear();

        fecha.clear();
        matricula.clear();

        FileOutputStream out = new FileOutputStream(ExcelPath);
        workbook.write(out);

      // fill(drive, info);
        info.clear();

        workbook.close();
        input.close();
        out.close();
        drive.close();
    }

    //Page fill
    public static void fill(WebDriver dr, ArrayList<String> list){
        WebElement d;

        Select edad = new Select(dr.findElement(By.id("edad")));
        Select situacion = new Select(dr.findElement(By.id("situacion")));

        Select ciclo = new Select(dr.findElement(By.id("ciclo")));
        Select grado = new Select(dr.findElement(By.id("grado")));

        for(int  i = 9; i < list.size();i+=9){
            d = dr.findElement(By.id("apellido"));
            d.sendKeys(list.get(i));

            d = dr.findElement(By.id("nombre"));
            d.sendKeys(list.get(i+1));

            edad.selectByVisibleText(list.get(i+2));

            if(list.get(i+3).charAt(0) == 'M'){
                d = dr.findElement(By.id("masculino"));
                d.click();
            }else{
                d = dr.findElement(By.id("femenino"));
                d.click();
            }
            situacion.selectByVisibleText(list.get(i+4));

            ciclo.selectByVisibleText(list.get(i+5));
            grado.selectByVisibleText(list.get(i+6));

            d = dr.findElement(By.id("direccion"));
            d.sendKeys(list.get(i+7));

            d = dr.findElement(By.id("telefono"));
            d.sendKeys(list.get(i+8));

            d = dr.findElement(By.id("botton"));
            d.click();
        }

    }
}
