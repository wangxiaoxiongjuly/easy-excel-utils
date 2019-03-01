package xiong.utils;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface CellStyleFormat {

    /**
     * 水平居中方式 默认左居中
     * @see HorizontalAlignment
     */
    HorizontalAlignment horizontalalignment() default HorizontalAlignment.LEFT;

    /**
     * 字体设置
     * @see org.apache.poi.xssf.usermodel.XSSFFont
     * @see org.apache.poi.hssf.usermodel.HSSFFont
     */
    CellFontFormat cellFont() default @CellFontFormat();


    /**
     * 背景颜色
     * @see IndexedColors
     */
    IndexedColors fillBackgroundColor() default IndexedColors.WHITE;

}
