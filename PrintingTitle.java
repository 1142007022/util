package com.safeschool.admin.annotation;

import java.lang.annotation.*;

/**
 * @program: safe-school-admin
 * @description: todo
 * @author: JD
 * @create: 2019-12-17 15:42
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
@Inherited
@Documented
public @interface PrintingTitle {

    String titleName() default "";

    int totalColumn() default 0;

}
