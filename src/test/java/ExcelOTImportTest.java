import com.djx.excelot.ExcelImportAndExport;
import com.djx.excelot.entity.Student;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelOTImportTest {

    public static void main(String[] args) throws Exception {

        List<Student> students = new ArrayList<>();

        students.add(new Student(1l, "小明", 18, new Date(), null, 1000.12, 1, "a"));
        students.add(new Student(2l, "小红", 19, new Date(), false, 200.0078, 2,"b"));
        students.add(new Student(3l, "小国", 19, new Date(), true, 300.567, null,"c"));
        students.add(new Student(4l, "小太", 18, null, false, 400.1, 2,null));


        // 使用
        ExcelImportAndExport<Student> excelUtils = new ExcelImportAndExport<>(Student.class);
        Workbook workbook = excelUtils.exportExcel(students);
        excelUtils.insertObject(5, new Student(3l, "小大", null, new Date(), false, 300.567, 3,"a"));
        workbook = excelUtils.workbook();
        File file = new File("F:\\student.xlsx");
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }
}
