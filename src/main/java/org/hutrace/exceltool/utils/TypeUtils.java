package org.hutrace.exceltool.utils;

import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.Locale;
import java.util.Map;

import org.hutrace.exceltool.exception.TypeCastException;

/**
 * <p>类型转换的工具类
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月15日
 */
public class TypeUtils {

	public static String DEFFAULT_DATE_FORMAT = "yyyy-MM-dd HH:mm:ss";
	private static boolean oracleTimestampMethodInited = false;
	private static Method oracleTimestampMethod;
	private static boolean oracleDateMethodInited = false;
	private static Method oracleDateMethod;

	public static Object cast(Object obj, Class<?> clazz) {
		if (obj == null) {
			if (clazz == int.class) {
				return Integer.valueOf(0);
			} else if (clazz == long.class) {
				return Long.valueOf(0);
			} else if (clazz == short.class) {
				return Short.valueOf((short) 0);
			} else if (clazz == byte.class) {
				return Byte.valueOf((byte) 0);
			} else if (clazz == float.class) {
				return Float.valueOf(0);
			} else if (clazz == double.class) {
				return Double.valueOf(0);
			} else if (clazz == boolean.class) {
				return Boolean.FALSE;
			}
			return null;
		}
		if (clazz == null) {
			throw new IllegalArgumentException("clazz is null");
		}
		if (clazz == obj.getClass()) {
			return obj;
		}
		if (clazz.isAssignableFrom(obj.getClass())) {
			return obj;
		}
		if (clazz == boolean.class || clazz == Boolean.class) {
			return castToBoolean(obj);
		}
		if (clazz == byte.class || clazz == Byte.class) {
			return castToByte(obj);
		}
		if (clazz == char.class || clazz == Character.class) {
			return castToChar(obj);
		}
		if (clazz == short.class || clazz == Short.class) {
			return castToShort(obj);
		}
		if (clazz == int.class || clazz == Integer.class) {
			return castToInt(obj);
		}
		if (clazz == long.class || clazz == Long.class) {
			return castToLong(obj);
		}
		if (clazz == float.class || clazz == Float.class) {
			return castToFloat(obj);
		}
		if (clazz == double.class || clazz == Double.class) {
			return castToDouble(obj);
		}
		if (clazz == String.class) {
			return castToString(obj);
		}
		if (clazz == BigDecimal.class) {
			return castToBigDecimal(obj);
		}
		if (clazz == BigInteger.class) {
			return castToBigInteger(obj);
		}
		if (clazz == Date.class) {
			return castToDate(obj);
		}
		if (clazz.isEnum()) {
			return castToEnum(obj, clazz);
		}
		if (Calendar.class.isAssignableFrom(clazz)) {
			Date date = castToDate(obj);
			Calendar calendar;
			if (clazz == Calendar.class) {
				calendar = Calendar.getInstance();
			} else {
				try {
					calendar = (Calendar) clazz.newInstance();
				} catch (Exception e) {
					throw new TypeCastException("can not cast to : "
							+ clazz.getName(), e);
				}
			}
			calendar.setTime(date);
			return calendar;
		}
		if (obj instanceof String) {
			String strVal = (String) obj;
			if (strVal.length() == 0 //
					|| "null".equals(strVal) //
					|| "NULL".equals(strVal)) {
				return null;
			}
			if (clazz == java.util.Currency.class) {
				return java.util.Currency.getInstance(strVal);
			}
			if (clazz == java.util.Locale.class) {
				return toLocale(strVal);
			}
		}
		throw new TypeCastException("can not cast to : " + clazz.getName());
	}

	public static Boolean castToBoolean(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof Boolean) {
			return (Boolean) value;
		}
		if (value instanceof Number) {
			return ((Number) value).intValue() == 1;
		}
		if (value instanceof String) {
			String strVal = (String) value;
			if (strVal.length() == 0 //
					|| "null".equals(strVal) //
					|| "NULL".equals(strVal)) {
				return null;
			}
			if ("true".equalsIgnoreCase(strVal) //
					|| "1".equals(strVal)) {
				return Boolean.TRUE;
			}
			if ("false".equalsIgnoreCase(strVal) //
					|| "0".equals(strVal)) {
				return Boolean.FALSE;
			}
			if ("Y".equalsIgnoreCase(strVal) //
					|| "T".equals(strVal)) {
				return Boolean.TRUE;
			}
			if ("F".equalsIgnoreCase(strVal) //
					|| "N".equals(strVal)) {
				return Boolean.FALSE;
			}
		}
		throw new TypeCastException("can not cast to boolean, value : " + value);
	}

	public static Byte castToByte(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof Number) {
			return ((Number) value).byteValue();
		}
		if (value instanceof String) {
			String strVal = (String) value;
			if (strVal.length() == 0 //
					|| "null".equals(strVal) //
					|| "NULL".equals(strVal)) {
				return null;
			}
			return Byte.parseByte(strVal);
		}
		throw new TypeCastException("can not cast to byte, value : " + value);
	}

	public static Character castToChar(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof Character) {
			return (Character) value;
		}
		if (value instanceof String) {
			String strVal = (String) value;
			if (strVal.length() == 0) {
				return null;
			}
			if (strVal.length() != 1) {
				throw new TypeCastException("can not cast to char, value : "
						+ value);
			}
			return strVal.charAt(0);
		}
		throw new TypeCastException("can not cast to char, value : " + value);
	}

	public static Short castToShort(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof Number) {
			return ((Number) value).shortValue();
		}
		if (value instanceof String) {
			String strVal = (String) value;
			if (strVal.length() == 0 //
					|| "null".equals(strVal) //
					|| "NULL".equals(strVal)) {
				return null;
			}
			return Short.parseShort(strVal);
		}
		throw new TypeCastException("can not cast to short, value : " + value);
	}

	@SuppressWarnings("rawtypes")
	public static BigDecimal castToBigDecimal(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof BigDecimal) {
			return (BigDecimal) value;
		}
		if (value instanceof BigInteger) {
			return new BigDecimal((BigInteger) value);
		}
		String strVal = value.toString();
		if (strVal.length() == 0) {
			return null;
		}
		if (value instanceof Map && ((Map) value).size() == 0) {
			return null;
		}
		return new BigDecimal(strVal);
	}

	public static BigInteger castToBigInteger(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof BigInteger) {
			return (BigInteger) value;
		}
		if (value instanceof Float || value instanceof Double) {
			return BigInteger.valueOf(((Number) value).longValue());
		}
		String strVal = value.toString();
		if (strVal.length() == 0 //
				|| "null".equals(strVal) //
				|| "NULL".equals(strVal)) {
			return null;
		}
		return new BigInteger(strVal);
	}

	public static Float castToFloat(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof Number) {
			return ((Number) value).floatValue();
		}
		if (value instanceof String) {
			String strVal = value.toString();
			if (strVal.length() == 0 //
					|| "null".equals(strVal) //
					|| "NULL".equals(strVal)) {
				return null;
			}
			if (strVal.indexOf(',') != 0) {
				strVal = strVal.replaceAll(",", "");
			}
			return Float.parseFloat(strVal);
		}
		throw new TypeCastException("can not cast to float, value : " + value);
	}

	public static Double castToDouble(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof Number) {
			return ((Number) value).doubleValue();
		}
		if (value instanceof String) {
			String strVal = value.toString();
			if (strVal.length() == 0 //
					|| "null".equals(strVal) //
					|| "NULL".equals(strVal)) {
				return null;
			}
			if (strVal.indexOf(',') != 0) {
				strVal = strVal.replaceAll(",", "");
			}
			return Double.parseDouble(strVal);
		}
		throw new TypeCastException("can not cast to double, value : " + value);
	}

	public static Date castToDate(Object value) {
		return castToDate(value, null);
	}

	public static Date castToDate(Object value, String format) {
		if (value == null) {
			return null;
		}
		if (value instanceof Date) { // 使用频率最高的，应优先处理
			return (Date) value;
		}
		if (value instanceof Calendar) {
			return ((Calendar) value).getTime();
		}
		long longValue = -1;
		if (value instanceof Number) {
			longValue = ((Number) value).longValue();
			return new Date(longValue);
		}
		if (value instanceof String) {
			String strVal = (String) value;
			if (strVal.startsWith("/Date(") && strVal.endsWith(")/")) {
				strVal = strVal.substring(6, strVal.length() - 2);
			}
			if (strVal.indexOf('-') != -1) {
				if (format == null) {
					if (strVal.length() == DEFFAULT_DATE_FORMAT.length()
							|| (strVal.length() == 22 && DEFFAULT_DATE_FORMAT
									.equals("yyyyMMddHHmmssSSSZ"))) {
						format = DEFFAULT_DATE_FORMAT;
					} else if (strVal.length() == 10) {
						format = "yyyy-MM-dd";
					} else if (strVal.length() == "yyyy-MM-dd HH:mm:ss"
							.length()) {
						format = "yyyy-MM-dd HH:mm:ss";
					} else if (strVal.length() == 29
							&& strVal.charAt(26) == ':'
							&& strVal.charAt(28) == '0') {
						format = "yyyy-MM-dd'T'HH:mm:ss.SSSXXX";
					} else {
						format = "yyyy-MM-dd HH:mm:ss.SSS";
					}
				}

				SimpleDateFormat dateFormat = new SimpleDateFormat(format);
				try {
					return dateFormat.parse(strVal);
				} catch (ParseException e) {
					throw new TypeCastException(
							"can not cast to Date, value : " + strVal);
				}
			}
			if (strVal.length() == 0) {
				return null;
			}
			longValue = Long.parseLong(strVal);
		}
		if (longValue < 0) {
			Class<?> clazz = value.getClass();
			if ("oracle.sql.TIMESTAMP".equals(clazz.getName())) {
				if (oracleTimestampMethod == null
						&& !oracleTimestampMethodInited) {
					try {
						oracleTimestampMethod = clazz.getMethod("toJdbc");
					} catch (NoSuchMethodException e) {
						// skip
					} finally {
						oracleTimestampMethodInited = true;
					}
				}
				Object result;
				try {
					result = oracleTimestampMethod.invoke(value);
				} catch (Exception e) {
					throw new TypeCastException(
							"can not cast oracle.sql.TIMESTAMP to Date", e);
				}
				return (Date) result;
			}
			if ("oracle.sql.DATE".equals(clazz.getName())) {
				if (oracleDateMethod == null && !oracleDateMethodInited) {
					try {
						oracleDateMethod = clazz.getMethod("toJdbc");
					} catch (NoSuchMethodException e) {
						// skip
					} finally {
						oracleDateMethodInited = true;
					}
				}
				Object result;
				try {
					result = oracleDateMethod.invoke(value);
				} catch (Exception e) {
					throw new TypeCastException(
							"can not cast oracle.sql.DATE to Date", e);
				}
				return (Date) result;
			}
			throw new TypeCastException("can not cast to Date, value : "
					+ value);
		}
		return new Date(longValue);
	}

	@SuppressWarnings("rawtypes")
	public static Integer castToInt(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof Integer) {
			return (Integer) value;
		}
		if (value instanceof Number) {
			return ((Number) value).intValue();
		}
		if (value instanceof String) {
			String strVal = (String) value;
			if (strVal.length() == 0 //
					|| "null".equals(strVal) //
					|| "NULL".equals(strVal)) {
				return null;
			}
			if (strVal.indexOf(',') != 0) {
				strVal = strVal.replaceAll(",", "");
			}
			return Integer.parseInt(strVal);
		}
		if (value instanceof Boolean) {
			return ((Boolean) value).booleanValue() ? 1 : 0;
		}
		if (value instanceof Map) {
			Map map = (Map) value;
			if (map.size() == 2 && map.containsKey("andIncrement")
					&& map.containsKey("andDecrement")) {
				Iterator iter = map.values().iterator();
				iter.next();
				Object value2 = iter.next();
				return castToInt(value2);
			}
		}
		throw new TypeCastException("can not cast to int, value : " + value);
	}

	@SuppressWarnings("rawtypes")
	public static Long castToLong(Object value) {
		if (value == null) {
			return null;
		}
		if (value instanceof Number) {
			return ((Number) value).longValue();
		}
		if (value instanceof String) {
			String strVal = (String) value;
			if (strVal.length() == 0 //
					|| "null".equals(strVal) //
					|| "NULL".equals(strVal)) {
				return null;
			}
			if (strVal.indexOf(',') != 0) {
				strVal = strVal.replaceAll(",", "");
			}
			try {
				return Long.parseLong(strVal);
			} catch (NumberFormatException ex) {
				throw new TypeCastException("can not cast to long, value : "
						+ value);
			}
		}
		if (value instanceof Map) {
			Map map = (Map) value;
			if (map.size() == 2 && map.containsKey("andIncrement")
					&& map.containsKey("andDecrement")) {
				Iterator iter = map.values().iterator();
				iter.next();
				Object value2 = iter.next();
				return castToLong(value2);
			}
		}
		throw new TypeCastException("can not cast to long, value : " + value);
	}

	public static String castToString(Object value) {
		if (value == null) {
			return null;
		}
		return value.toString();
	}

	@SuppressWarnings({ "unchecked", "rawtypes" })
	public static <T> T castToEnum(Object obj, Class<T> clazz) {
		try {
			if (obj instanceof String) {
				String name = (String) obj;
				if (name.length() == 0) {
					return null;
				}
				return (T) Enum.valueOf((Class<? extends Enum>) clazz, name);
			}
			if (obj instanceof Number) {
				int ordinal = ((Number) obj).intValue();
				Object[] values = clazz.getEnumConstants();
				if (ordinal < values.length) {
					return (T) values[ordinal];
				}
			}
		} catch (Exception ex) {
			throw new TypeCastException("can not cast to : " + clazz.getName(),
					ex);
		}
		throw new TypeCastException("can not cast to : " + clazz.getName());
	}

	public static Locale toLocale(String strVal) {
		String[] items = strVal.split("_");
		if (items.length == 1) {
			return new Locale(items[0]);
		}
		if (items.length == 2) {
			return new Locale(items[0], items[1]);
		}
		return new Locale(items[0], items[1], items[2]);
	}

}
