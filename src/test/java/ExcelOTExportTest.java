import com.djx.excelot.ExcelImportAndExport;
import com.djx.excelot.entity.Student;
import com.djx.excelot.entity.Teacher;
import com.djx.excelot.exception.ExcelChanelException;
import com.djx.excelot.exception.ExcelNullpointExcetion;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

public class ExcelOTExportTest {

    public static void main(String[] args) throws Exception {

        File file = new File("F:\\student.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        // 使用
        ExcelImportAndExport<Student> excelUtils = new ExcelImportAndExport<>(Student.class);
        excelUtils.importExcel(fileInputStream, "student.xlsx");

        List<Student> list = null;
        try {
            list = excelUtils.getAllList();
        } catch (ExcelChanelException e) {
            System.out.println(e.getRowIndex() + "," + e.getCellIndex());
            e.printStackTrace();
        } catch (ExcelNullpointExcetion e) {
            System.out.println(e.getRowIndex() + "," + e.getCellIndex());
            e.printStackTrace();
        }
        Student student = null;
        try {
            student = excelUtils.getObjByRow(2);
        } catch (ExcelChanelException e) {
            e.printStackTrace();
        } catch (ExcelNullpointExcetion excelNullpointExcetion) {
            excelNullpointExcetion.printStackTrace();
        }

        System.out.println(list);
        System.out.println(student);
        System.out.println(excelUtils.getNotNullLastIndex());
    }
}
