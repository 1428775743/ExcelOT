import com.djx.excelot.ExcelImportAndExport;
import com.djx.excelot.entity.Student;
import com.djx.excelot.entity.Teacher;
import com.djx.excelot.exception.ExcelChanelException;
import com.djx.excelot.exception.ExcelDateParseException;
import com.djx.excelot.exception.ExcelNullpointExcetion;
import com.djx.excelot.exception.ExcelOutLenException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.ArrayList;
import java.util.List;

public class ExcelOTExportTest {

    public static void main(String[] args) throws Exception {

        File file = new File("F:\\student.xlsx");
        FileInputStream fileInputStream = new FileInputStream(file);

        // 使用
        ExcelImportAndExport<Student> excelUtils = new ExcelImportAndExport<>(Student.class);
        excelUtils.importExcel(fileInputStream, "student.xlsx");

        List<Student> list = new ArrayList<>();

        for (int i = 1; i <= excelUtils.getNotNullLastIndex(); i++) {
            Student student = null;
            try {
                student = excelUtils.getObjByRow(i);
            } catch (ExcelDateParseException e) {

                excelUtils.getColumnName(e.getCellIndex());

                System.out.println("第" + (e.getRowIndex() + 1) + "行、第" + (e.getCellIndex() + 1) + "列格式异常 请使用标准格式例：2020-10-10（表格设置单元格格式中选择日期）");
                continue;
            } catch (ExcelChanelException e) {
                System.out.println("第" + (e.getRowIndex() + 1) + "行、第" + (e.getCellIndex() + 1) + "列格式异常 请检查数据格式是否正确，单元格格式是否正确");
                continue;
            } catch (ExcelNullpointExcetion e) {
                System.out.println("第" + (e.getRowIndex() + 1) + "行、第" + (e.getCellIndex() + 1) + "列不能为空");
                continue;
            } catch (ExcelOutLenException e) {
                System.out.println("第" + (e.getRowIndex() + 1) + "行、第" + (e.getCellIndex() + 1) + "列列长度超过了" + e.getMaxLen());
                continue;
            }
            list.add(student);
        }

        System.out.println(list);
    }
}
