package xiong.utils;

import org.apache.poi.ss.usermodel.IndexedColors;

import java.lang.annotation.*;

@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface CellFontFormat {

    String fontName() default "";

    short fontHeightInPoints() default 11;

    IndexedColors fontColor() default IndexedColors.BLACK;

    boolean bold() default false;
}
