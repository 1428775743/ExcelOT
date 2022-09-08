import com.djx.excelot.ExcelImportAndExport;
import com.djx.excelot.entity.Student;
import com.djx.excelot.exception.ExcelChanelException;
import com.djx.excelot.exception.ExcelNullpointExcetion;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelOTImportTest2 {

    public static void main(String[] args) throws Exception {


        List<Student> students = new ArrayList<>();

        students.add(new Student(1l, "小明", 18, new Date(), null, 1000.12, 1, "a"));
        students.add(new Student(2l, "小红", 19, new Date(), false, 200.0078, 2,"b"));
        students.add(new Student(3l, "小国", 19, new Date(), true, 300.567, null,"c"));
        students.add(new Student(4l, "小太", 18, null, false, 400.1, 2,null));

        // 添加测试用例
        students.add(new Student(null, null, null, null, null, null, null,"c"));
        students.add(new Student(null, "小放", 19, new Date(), false, 200.0078, 2,null));
        students.add(new Student(3l, "小大", null, new Date(), false, 300.567, 3,"a"));
        students.add(new Student(4l, "小二", 18, new Date(), true, null, 2,"b"));

        for (int i = 0;i < 400000;i++) {
            students.add(new Student(1l, "小明", 18, new Date(), null, 1000.12, 1, "a"));
        }
        // 100w条数据时间测试
        Long start = System.currentTimeMillis();


        // 使用
        ExcelImportAndExport<Student> excelUtils = new ExcelImportAndExport<>(Student.class);
        Workbook workbook = null;
        try {
            workbook = excelUtils.exportExcel(students);

        } catch (ExcelNullpointExcetion e){
            int row = e.getRowIndex();
            int cell = e.getCellIndex();
            System.out.println((row + 1) + "行" + (cell + 1) + "不能为空");
        } catch (ExcelChanelException e){
            int row = e.getRowIndex();
            int cell = e.getCellIndex();
            System.out.println((row + 1) + "行" + (cell + 1) + "列格式错误");
        } catch (Exception exception) {
            exception.printStackTrace();
        }


        Long end = System.currentTimeMillis();
        System.out.println(end-start);
        File file = new File("F:\\student.xlsx");
        FileOutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
    }
}
