package com.smlf.excel.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 工作表数据集
 *
 * @author smlf 2018-10-06 20:42:41
 */
@Target({ElementType.TYPE})
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface Dataset {

    /**
     * 工作表名称
     *
     * @return
     */
    String sheetName() default "";

}

