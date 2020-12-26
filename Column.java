package com.jiangdong.annotation;

import java.lang.annotation.*;

/**
 * @author jiangdong
 * @title: Column
 * @description: TODO
 * @date 2020-11-14 15:42
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
@Inherited
@Documented
public @interface Column {

    String name() default "";

    String timeFormat() default "";

    boolean rate() default false;
}
