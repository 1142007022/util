package com.jiangdong.annotation;

import java.lang.annotation.*;

/**
 * @description: todo
 * @author: JD
 * @create: 2020-11-14 15:42
 **/
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.TYPE)
@Inherited
@Documented
public @interface Title {

    String title() default "";

}
