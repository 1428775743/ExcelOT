import com.djx.excelot.ExcelImportAndExport;
import com.djx.excelot.entity.Student;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelOTImportTest {

    public static void main(String[] args) {

        List<Student> students = new ArrayList<>();

        students.add(new Student(1l, "小明", 18, new Date(), null, 1000.12, 1, "a"));
        students.add(new Student(2l, "小红", 19, new Date(), false, 200.0078, 2,"b"));
        students.add(new Student(3l, "小国", 19, new Date(), true, 300.567, null,"c"));
        students.add(new Student(4l, "小太", 18, null, false, 400.1, 2,null));


        // 生成文件的位置
        File file = new File("F:\\student.xlsx");

        try (FileOutputStream outputStream = new FileOutputStream(file)){

            ExcelImportAndExport<Student> excelUtils = new ExcelImportAndExport<>(Student.class);

            // 将对象列表转成 excel
            Workbook workbook = excelUtils.exportExcel(students);
            workbook = excelUtils.workbook();
            workbook.write(outputStream);

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
