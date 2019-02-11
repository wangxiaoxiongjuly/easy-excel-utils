package xiong.utils;

import com.alibaba.excel.annotation.ExcelProperty;
import com.alibaba.excel.event.WriteHandler;
import com.alibaba.excel.metadata.BaseRowModel;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.Map;

/**
 * XlSX文件类型定制操作
 */
public class WriterHandler07<T extends Class<? extends BaseRowModel>> implements WriteHandler {

    private Map<Integer, CellStyleFormat> colStyleFormat = new HashMap<>();

    private Map<Integer, CellStyle> colStyle = new HashMap<>();

    private Map<String, String> infoMap = new HashMap<>();

    private static final String TITLE_NUM = "title_num";

    /**
     * 调用init方法
     * @param rowModel 实体类类型
     */
    WriterHandler07(T rowModel) {
        init(rowModel);
    }

    /**
     * 利用反射解析对象属性，记录实体类的自定义注解
     * @param rowModel 实体类类型
     */
    private void init(T rowModel) {
        Field[] declaredFields = rowModel.getDeclaredFields();
        //查找有CellStyle注解和ExcelProperty注解的Field加入cellStyle
        int titleNum = 0;
        for (Field field : declaredFields) {
            CellStyleFormat cellStyle = field.getAnnotation(CellStyleFormat.class);
            ExcelProperty excelProperty = field.getAnnotation(ExcelProperty.class);
            if (excelProperty != null) {
                titleNum =Math.max(excelProperty.value().length, titleNum);
            }
            if (cellStyle == null || excelProperty == null) {
                continue;
            }
            colStyleFormat.put(excelProperty.index(), cellStyle);

        }
        infoMap.put(TITLE_NUM, String.valueOf(titleNum));
    }

    /**
     * 输出sheet时进行的调用
     * @param sheetNo sheetNo
     * @param sheet Poi Sheet
     */
    @Override
    public void sheet(int sheetNo, Sheet sheet) {
        for (Map.Entry entry:colStyleFormat.entrySet()){
            colStyle.put((Integer)entry.getKey(),getCellStyle(sheet,(CellStyleFormat)(entry.getValue())));
        }
    }

    /**
     * 输出行时调用
     * @param rowNum 行号
     * @param row Poi Row
     */
    @Override
    public void row(int rowNum, Row row) {
        //doSomething
    }

    /**
     * 输出列时调用
     * @param cellNum 列号
     * @param cell POI Cell
     */
    @Override
    public void cell(int cellNum, Cell cell) {
        CellStyle cellStyle = colStyle.get(cellNum);
        // 表头不操作 行序号是从0开始的
        if (cell.getRowIndex() <= Integer.valueOf(infoMap.get(TITLE_NUM))-1)  {
            return;
        }
        // 没有不操作
        if (cellStyle != null) {
            cell.setCellStyle(cellStyle);
        }
    }

    /**
     * 生成样式
     * @param sheet Poi Sheet
     * @param value 对象样式注解
     * @return 返回定制化样式
     */
    private static CellStyle getCellStyle(Sheet sheet,CellStyleFormat value){
        // 生成一个样式
        CellStyle style = sheet.getWorkbook().createCellStyle();

        // 设置这些样式
        style.setFillForegroundColor(value.fillBackgroundColor().index);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(value.horizontalalignment());

        // 生成一个字体
        Font font = sheet.getWorkbook().createFont();
        font.setFontName(value.cellFont().fontName());
        font.setColor(value.cellFont().fontColor().index);
        font.setFontHeightInPoints(value.cellFont().fontHeightInPoints());
        font.setBold(value.cellFont().bold());

        // 把字体应用到当前的样式
        style.setFont(font);

        return style;
    }

}
