
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
// https://blog.csdn.net/ZixiangLi/article/details/80243535
public class XlsxReader {
    public static void main(String[] args) throws Exception {
        excel();
    }
    public static void excel() throws Exception {
        //用流的方式先读取到你想要的excel的文件
        FileInputStream fis = new FileInputStream(new File("F:\\Desktop\\cctv.xlsx"));
        //解析excel
        XSSFWorkbook hb = new XSSFWorkbook(fis);
        //获取第一个表单sheet
        Sheet sheet = hb.getSheetAt(0);
        //获取第一行
        int firstrow = sheet.getFirstRowNum();
        //获取最后一行
        int lastrow = sheet.getLastRowNum();
        //循环行数依次获取列数
        for (int i = firstrow; i < lastrow + 1; i++) {
            //获取哪一行i
            Row row = sheet.getRow(i);
            if (row != null) {
                //获取这一行的第一列
                int firstcell = row.getFirstCellNum();
                //获取这一行的最后一列
                int lastcell = row.getLastCellNum();
                //创建一个集合，用处将每一行的每一列数据都存入集合中
                List<String> list = new ArrayList<>();
                for (int j = firstcell; j < lastcell; j++) {
                    //获取第j列
                    Cell cell = row.getCell(j);
                    if (cell != null) {
                        System.out.print(cell + "\t");
                        list.add(cell.toString());
                    }
                }
                System.out.println();
            }
        }
        fis.close();
    }
}
