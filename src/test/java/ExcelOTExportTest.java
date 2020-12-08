import com.djx.excelot.ExcelImportAndExport;
import com.djx.excelot.entity.Teacher;

import java.io.File;
import java.io.FileInputStream;
import java.util.List;

public class ExcelOTExportTest {

    public static void main(String[] args) throws Exception {

        File file = new File("F:\\student.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        // 使用
        ExcelImportAndExport<Teacher> excelUtils = new ExcelImportAndExport<>();
        List<Teacher> list = excelUtils.importExcel(fileInputStream, "student.xlsx", Teacher.class);

        System.out.println(list);
    }
}
