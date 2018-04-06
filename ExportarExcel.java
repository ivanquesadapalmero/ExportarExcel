package excel;

	import org.apache.poi.hssf.usermodel.HSSFCell;
	import org.apache.poi.hssf.usermodel.HSSFRow;
	import org.apache.poi.hssf.usermodel.HSSFSheet;
	import org.apache.poi.hssf.usermodel.HSSFWorkbook;
	import org.apache.poi.ss.usermodel.CellStyle;
	import org.apache.poi.ss.usermodel.Font;
	import org.apache.poi.ss.usermodel.IndexedColors;

	import java.io.FileOutputStream;
	import java.math.BigDecimal;

	public class ExportarExcel {

	    public static void main(String[] args) throws Exception {
	        HSSFWorkbook workbook = new HSSFWorkbook();
	        HSSFSheet sheet = workbook.createSheet();
	        workbook.setSheetName(0, "Hoja excel");

	        String[] headers = new String[]{
	            "Nombre",
	            "Apellido",
	            "Código Alumno",
	            "PC"
	        };

	        Object[][] data = new Object[][] {
	            new Object[] {"Jesús", "Delgado", new Integer(1), "PC1" },
	            new Object[] {"David", "Morano",new Integer(2), "PC2" },
	            new Object[] {"Iván", "Quesada", new Integer(3), "PC3" },
	            new Object[] {"Amelia", "Paniagua", new Integer(4), "PC4" },
	            new Object[] {"Rafa", "Álvarez", new Integer(5), "PC5" },
	            new Object[] {"Antonio", "Cívico", new Integer(6), "PC6" },
	            
	        };

	        CellStyle headerStyle = workbook.createCellStyle();
	        Font font = workbook.createFont();
	        headerStyle.setFont(font);

	        CellStyle style = workbook.createCellStyle();
	        style.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.getIndex());

	        HSSFRow headerRow = sheet.createRow(0);
	        for (int i = 0; i < headers.length; ++i) {
	            String header = headers[i];
	            HSSFCell cell = headerRow.createCell(i);
	            cell.setCellStyle(headerStyle);
	            cell.setCellValue(header);
	        }

	        for (int i = 0; i < data.length; ++i) {
	            HSSFRow dataRow = sheet.createRow(i + 1);

	            Object[] d = data[i];
	            String nombre = (String) d[0];
	            String apellido = (String) d[1];
	            Integer codigo = (Integer) d[2];
	            String pc = (String) d[3];

	            dataRow.createCell(0).setCellValue(nombre);
	            dataRow.createCell(1).setCellValue(apellido);
	            dataRow.createCell(2).setCellValue(codigo);
	            dataRow.createCell(3).setCellValue(pc);
	        }

	        HSSFRow dataRow = sheet.createRow(1 + data.length);
	        HSSFCell total = dataRow.createCell(1);
	        total.setCellStyle(style);
	        total.setCellFormula(String.format("SUM(B2:B%d)", 1 + data.length));

	        FileOutputStream file = new FileOutputStream("ListaClase.xls");
	        workbook.write(file);
	        file.close();
	        workbook.close();
	    }
	}

