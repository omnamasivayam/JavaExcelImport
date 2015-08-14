import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.Iterator;
import java.util.List;
 
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.format.CellDateFormatter;
import org.apache.poi.ss.usermodel.Cell;
 
public class IssueTracker{
 
		public static void main(String[] args) throws Exception {
			
				String filename = "C:\\Users\\Omnamasivayam\\Downloads\\IPTV_IT.xls";
				List sheetData = new ArrayList();
				FileInputStream fis = null;
				try {
					fis = new FileInputStream(filename);
					HSSFWorkbook workbook = new HSSFWorkbook(fis);
					HSSFSheet sheet = workbook.getSheetAt(1);
					Iterator rows = sheet.rowIterator();
					while (rows.hasNext()) {
							HSSFRow row = (HSSFRow) rows.next();
							Iterator cells = row.cellIterator(); 
							List data = new ArrayList();
							while (cells.hasNext()) {
									HSSFCell cell = (HSSFCell) cells.next();
									data.add(cell);
							}
							sheetData.add(data);
					}
				} 
				catch (IOException e) {
						e.printStackTrace();
				} finally {
						if (fis != null) {
							fis.close();
						}
				} 
				showExelData(sheetData);
		}
 
private static void showExelData(List sheetData) {
	int j = 12;
		for (int i = 0; i < sheetData.size() -1 ; i++) {
				List list = (List) sheetData.get(i);
				//for (int j = 0; j < list.size(); j++) {
						Cell cell = (Cell) list.get(j);
						if (cell.getCellType() == Cell.CELL_TYPE_NUMERIC) {
							if(HSSFDateUtil.isCellDateFormatted(cell)){
								System.out.println("Number value");
								Date date = HSSFDateUtil.getJavaDate(cell.getNumericCellValue());
							    String dateFmt = cell.getCellStyle().getDataFormatString();
							    String strValue = new CellDateFormatter(dateFmt).format(date);
							    System.out.println(strValue);
							}else{
								System.out.print(cell.getNumericCellValue());
							}
						} else if (cell.getCellType() == Cell.CELL_TYPE_STRING) {
							//System.out.println(" i and j are " +i  + " " + j);
								System.out.print(cell.getRichStringCellValue());
						} else if (cell.getCellType() == Cell.CELL_TYPE_BOOLEAN) {
							
								System.out.print(cell.getBooleanCellValue());
						}
					//	if (j < list.size() - 1) {
					//			System.out.print(", ");
					//	}
				//}
				System.out.println("");
		}
	}
}