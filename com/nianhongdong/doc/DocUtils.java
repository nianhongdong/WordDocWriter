/**
 *@(#) DocUtils.java
 */
package com.nianhongdong.doc;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.regex.Pattern;

/**
 * 说明：word文档处理工具类
 * 
 * <p>
 * All rights reserved.
 * <p>
 * 版权所有：北京用友政务软件有限公司
 * <p>
 * 未经本公司许可，不得以任何方式复制或使用本程序任何部分，
 * <p>
 * 侵权者将受到法律追究。
 * <p>
 * DERIVED FROM: NONE
 * <p>
 * PURPOSE
 * <p>
 * DESCRIPTION:
 * <p>
 * CALLED BY:
 * <p>
 * UPDATE: nianhongdong
 * <p>
 * DATE: 2014-5-10 下午12:31:20
 * 
 * @version 1.0
 * @author nhd
 * @since java 1.4
 *        <p>
 *        HISTORY: 1.0 <author> <time> <version> <desc> 修改人姓名 修改时间 版本号 描述
 */

public class DocUtils {

	/** 获取所有方法 */
	public static Method[] getAllMethod(Object instance) {

		return instance.getClass().getDeclaredMethods();
	}

	// 获得指定变量的值
	public static Object getValue(Object instance, String fieldName)
			throws IllegalAccessException, NoSuchFieldException {
		Field field = getField(instance.getClass(), fieldName);
		// 参数值为true，禁用访问控制检查
		field.setAccessible(true);
		return field.get(instance);
	}

	// 该方法实现根据变量名获得该变量的值
	public static Field getField(Class thisClass, String fieldName)
			throws NoSuchFieldException {
		if (thisClass == null) {
			throw new NoSuchFieldException("Error field !");
		}
		return thisClass.getDeclaredField(fieldName);
	}

	public static Method getMethod(Object instance, String methodName,
			Class[] classTypes) throws NoSuchMethodException {

		Method accessMethod = getMethod(instance.getClass(), methodName,
				classTypes);
		// 参数值为true，禁用访问控制检查
		accessMethod.setAccessible(true);

		return accessMethod;
	}

	// 利用反射机制访问类的成员方法
	private static Method getMethod(Class thisClass, String methodName,
			Class[] classTypes) throws NoSuchMethodException {

		if (thisClass == null) {
			throw new NoSuchMethodException("Error method !");
		}
		try {
			return thisClass.getDeclaredMethod(methodName, classTypes);
		} catch (NoSuchMethodException e) {
			return getMethod(thisClass.getSuperclass(), methodName, classTypes);

		}
	}

	// 调用含单个参数的方法
	public static Object invokeMethod(Object instance, String methodName,
			Object arg) throws NoSuchMethodException, IllegalAccessException,
			InvocationTargetException {

		Object[] args = new Object[1];
		args[0] = arg;
		return invokeMethod(instance, methodName, args);
	}

	// 调用含多个参数的方法
	public static Object invokeMethod(Object instance, String methodName,
			Object[] args) throws NoSuchMethodException,
			IllegalAccessException, InvocationTargetException {
		Class[] classTypes = null;
		if (args != null) {
			classTypes = new Class[args.length];
			for (int i = 0; i < args.length; i++) {
				if (args[i] != null) {
					classTypes[i] = args[i].getClass();
				}
			}
		}
		return getMethod(instance, methodName, classTypes).invoke(instance,
				args);
	}
	
	/**
	 * 需要通过一个字符串来生成布尔值的时候，可以用这个正则表达式的模式串
	 */
	private static Pattern TRUE_PATTERN = Pattern.compile(
			"(\\s*[\\d&&[^0]]+\\s*)|(.*真.*)|(.*是.*)|(.*true.*)",
			Pattern.CASE_INSENSITIVE);
	
	/**
	 * 通过一个对象来生成布尔值，依赖于该对象类型实现的toString()方法
	 * 
	 * @param value
	 */
	public static boolean estimate(Object value) {
		return TRUE_PATTERN.matcher(nonNullStr(value)).matches();
	}
	
	/**
	 * 由一个对象返回非空的字符串
	 * 
	 * @param o
	 *            传入的对象
	 * @return 生成的字符串
	 */
	public static String nonNullStr(Object o) {
		return nonNullStr(o, "");
	}

	public static String nonNullStr(Object o, String def) {
		String s = o == null ? def : o.toString();
		if (s.equalsIgnoreCase("null")) {
			s = "";
		}
		return s;
	}
	
	/**
	 * 判断一个空字符串（null或者""）
	 * 
	 * @param s
	 *            待判断的字符串
	 * @return 判断结果
	 */
	public static boolean isNullStr(String s) {
		return s == null || s.trim().length() <= 0 || s.trim().equals("null");
	}
	
	private static Pattern NUMBER_PATTERN = Pattern
					.compile("[+-]?(\\d+(\\.\\d*)?|\\.\\d+)(E\\d+)?");
	
	/**
	 * 
	 * 功能说明:判断是否是数字
	 * 
	 * @param str
	 * @return
	 * @author nianhongdong
	 * @date   2014-5-17 下午5:45:32
	 * @see [相关类/方法]（可选）
	 */
	public static boolean isNumber(String str) {
		return NUMBER_PATTERN.matcher(str).matches();
	}

}
