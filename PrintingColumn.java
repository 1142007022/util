package com.safeschool.admin.annotation;

import java.lang.annotation.*;

/**
 * @author jiangdong
 * @title: Printing
 * @projectName safe-school-admin
 * @description: TODO
 * @date 2019/10/29 002915:18
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Inherited
@Documented
public @interface PrintingColumn {

    String name() default "";

}
