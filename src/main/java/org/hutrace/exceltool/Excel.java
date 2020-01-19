package org.hutrace.exceltool;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import org.hutrace.exceltool.annotation.ExcelField;
import org.hutrace.exceltool.pojo.ExcelType;
import org.hutrace.exceltool.pojo.TitleAlias;
import org.hutrace.exceltool.read.Reader;
import org.hutrace.exceltool.write.Writer;

/**
 * <p>Excel工具类
 * <p>封装读写等所有相关的Excel工具类
 * <p>一般的Excel操作可以直接通过此类的静态方法调用即可
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 */
public class Excel {
	
	/**
	 * 静态构造{@link TitleAlias}对象
	 * @param title Excel的标题
	 * @param alias Map或JavaBean的属性(字段)名称
	 * @return {@link TitleAlias}
	 */
	public static TitleAlias buildTitleAlias(String title, String alias) {
		return TitleAlias.build(title, alias);
	}
	
	/**
	 * <p>获取{@link Reader}对象
	 * @return {@link Reader}对象
	 */
	public static Reader reader() {
		return reader(null);
	}
	
	/**
	 * <p>获取{@link Reader}对象
	 * <p>并设置sheetName
	 * @return {@link Reader}对象
	 * @see Reader#setSheetName(String)
	 */
	public static Reader reader(String sheetName) {
		Reader reader = new Reader();
		if(sheetName != null) {
			reader.setSheetName(sheetName);
		}
		return reader;
	}
	
	/**
	 * <p>获取{@link Writer}对象
	 * @return {@link Writer}对象
	 */
	public static Writer writer() {
		return writer(null);
	}
	
	/**
	 * <p>获取{@link Writer}对象
	 * <p>并设置sheetName
	 * @return {@link Writer}对象
	 * @see Writer#setSheetName(String)}=
	 */
	public static Writer writer(String sheetName) {
		Writer writer = new Writer();
		if(sheetName != null) {
			writer.setSheetName(sheetName);
		}
		return writer;
	}
	
	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param realPath 根据realPath的后缀判断Excel的类型，如果没有后缀或后缀错误，将会抛出空指针异常。
	 * @return MapList
	 * @throws IOException
	 * @see Reader#toMap(String)
	 */
	public static List<Map<String, Object>> readToMap(String realPath) throws IOException {
		return reader().toMap(realPath);
	}

	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param in Excel文件输入流
	 * @param type Excel文件的类型，如果为空，将会抛出空指针异常。
	 * @return MapList
	 * @throws IOException
	 * @see Reader#toMap(InputStream, ExcelType)
	 */
	public static List<Map<String, Object>> readToMap(InputStream in, ExcelType type) throws IOException {
		return reader().toMap(in, type);
	}

	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键，可通过设置{@link #titleAlias}来实现在Map中的key值。
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param realPath 根据realPath的后缀判断Excel的类型，如果没有后缀或后缀错误，将会抛出空指针异常。
	 * @return MapList
	 * @throws IOException
	 * @see Reader#toMap(String, TitleAlias[])
	 */
	public static List<Map<String, Object>> readToMap(String realPath, TitleAlias[] titleAlias) throws IOException {
		return reader().toMap(realPath, titleAlias);
	}
	
	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键，可通过设置{@link #titleAlias}来实现在Map中的key值。
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param in Excel文件输入流
	 * @param type Excel文件的类型，如果为空，将会抛出空指针异常。
	 * @param titleAlias Excel标题与Java字段(属性名)对应类数组
	 * @return MapList
	 * @throws IOException
	 * @see Reader#toMap(InputStream, ExcelType, TitleAlias[])
	 */
	public static List<Map<String, Object>> readToMap(InputStream in, ExcelType type, TitleAlias[] titleAlias) throws IOException {
		return reader().toMap(in, type, titleAlias);
	}
	
	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为JavaBean的字段依据，可以在JavaBean字段上添加{@link ExcelField}注解来设置与标题对应的值
	 * <p>JavaBean只会解析一次，后面每次都按照标题行的长度、顺序执行（也就是说要保证内容行的列数不能超过标题行的列数，超出的列是读不到的。）
	 * @param realPath 根据realPath的后缀判断Excel的类型，如果没有后缀或后缀错误，将会抛出空指针异常。
	 * @param clazs 需要存放数据的JavaBean对象
	 * @return JavaBeanList
	 * @throws IOException
	 * @see Reader#toJavaBean(String, Class)
	 */
	public static <T> List<T> readToJavaBean(String realPath, Class<T> clazs) throws IOException {
		return reader().toJavaBean(realPath, clazs);
	}

	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为JavaBean的字段依据，可以在JavaBean字段上添加{@link ExcelField}注解来设置与标题对应的值
	 * <p>JavaBean只会解析一次，后面每次都按照标题行的长度、顺序执行（也就是说要保证内容行的列数不能超过标题行的列数，超出的列是读不到的。）
	 * @param in Excel文件输入流
	 * @param type Excel文件的类型，如果为空，将会抛出空指针异常。
	 * @param clazs 需要存放数据的JavaBean对象
	 * @return JavaBeanList
	 * @throws IOException
	 * @see Reader#toJavaBean(InputStream, ExcelType, Class)
	 */
	public static <T> List<T> readToJavaBean(InputStream in, ExcelType type, Class<T> clazs) throws IOException {
		return reader().toJavaBean(in, type, clazs);
	}
	
	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param realPath 根据realPath的后缀判断Excel的类型，如果没有后缀或后缀错误，将会抛出空指针异常。
	 * @param sheetName 读取Excel的sheet名称，默认读取Excel的第一个sheet，如果你的Excel有多个sheet，并且你想读取你需要的sheet时，你可以指定它。
	 * @return MapList
	 * @throws IOException
	 * @see Reader#toMap(String)
	 */
	public static List<Map<String, Object>> readToMap(String realPath, String sheetName) throws IOException {
		return reader(sheetName).toMap(realPath);
	}

	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param in Excel文件输入流
	 * @param type Excel文件的类型，如果为空，将会抛出空指针异常。
	 * @param sheetName 读取Excel的sheet名称，默认读取Excel的第一个sheet，如果你的Excel有多个sheet，并且你想读取你需要的sheet时，你可以指定它。
	 * @return MapList
	 * @throws IOException
	 * @see Reader#toMap(InputStream, ExcelType)
	 */
	public static List<Map<String, Object>> readToMap(InputStream in, ExcelType type, String sheetName) throws IOException {
		return reader(sheetName).toMap(in, type);
	}

	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键，可通过设置{@link #titleAlias}来实现在Map中的key值。
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param realPath 根据realPath的后缀判断Excel的类型，如果没有后缀或后缀错误，将会抛出空指针异常。
	 * @param sheetName 读取Excel的sheet名称，默认读取Excel的第一个sheet，如果你的Excel有多个sheet，并且你想读取你需要的sheet时，你可以指定它。
	 * @return MapList
	 * @throws IOException
	 * @see Reader#toMap(String, TitleAlias[])
	 */
	public static List<Map<String, Object>> readToMap(String realPath, TitleAlias[] titleAlias, String sheetName) throws IOException {
		return reader(sheetName).toMap(realPath, titleAlias);
	}
	
	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键，可通过设置{@link #titleAlias}来实现在Map中的key值。
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param in Excel文件输入流
	 * @param type Excel文件的类型，如果为空，将会抛出空指针异常。
	 * @param titleAlias Excel标题与Java字段(属性名)对应类数组
	 * @param sheetName 读取Excel的sheet名称，默认读取Excel的第一个sheet，如果你的Excel有多个sheet，并且你想读取你需要的sheet时，你可以指定它。
	 * @return MapList
	 * @throws IOException
	 * @see Reader#toMap(InputStream, ExcelType, TitleAlias[])
	 */
	public static List<Map<String, Object>> readToMap(InputStream in, ExcelType type, TitleAlias[] titleAlias, String sheetName) throws IOException {
		return reader(sheetName).toMap(in, type, titleAlias);
	}
	
	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为JavaBean的字段依据，可以在JavaBean字段上添加{@link ExcelField}注解来设置与标题对应的值
	 * <p>JavaBean只会解析一次，后面每次都按照标题行的长度、顺序执行（也就是说要保证内容行的列数不能超过标题行的列数，超出的列是读不到的。）
	 * @param realPath 根据realPath的后缀判断Excel的类型，如果没有后缀或后缀错误，将会抛出空指针异常。
	 * @param clazs 需要存放数据的JavaBean对象
	 * @param sheetName 读取Excel的sheet名称，默认读取Excel的第一个sheet，如果你的Excel有多个sheet，并且你想读取你需要的sheet时，你可以指定它。
	 * @return JavaBeanList
	 * @throws IOException
	 * @see Reader#toJavaBean(String, Class)
	 */
	public static <T> List<T> readToJavaBean(String realPath, Class<T> clazs, String sheetName) throws IOException {
		return reader(sheetName).toJavaBean(realPath, clazs);
	}

	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为JavaBean的字段依据，可以在JavaBean字段上添加{@link ExcelField}注解来设置与标题对应的值
	 * <p>JavaBean只会解析一次，后面每次都按照标题行的长度、顺序执行（也就是说要保证内容行的列数不能超过标题行的列数，超出的列是读不到的。）
	 * @param in Excel文件输入流
	 * @param type Excel文件的类型，如果为空，将会抛出空指针异常。
	 * @param clazs 需要存放数据的JavaBean对象
	 * @param sheetName 读取Excel的sheet名称，默认读取Excel的第一个sheet，如果你的Excel有多个sheet，并且你想读取你需要的sheet时，你可以指定它。
	 * @return JavaBeanList
	 * @throws IOException
	 * @see Reader#toJavaBean(InputStream, ExcelType, Class)
	 */
	public static <T> List<T> readToJavaBean(InputStream in, ExcelType type, Class<T> clazs, String sheetName) throws IOException {
		return reader(sheetName).toJavaBean(in, type, clazs);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @throws IOException
	 * @see Writer#mapToFile(List, String)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath) throws IOException {
		writer().mapToFile(list, realPath);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, String[])
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, String[] collation) throws IOException {
		writer().mapToFile(list, realPath, collation);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, TitleAlias[])
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias) throws IOException {
		writer().mapToFile(list, realPath, titleAlias);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, TitleAlias[], String[])
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias, String[] collation) throws IOException {
		writer().mapToFile(list, realPath, titleAlias, collation);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @throws IOException
	 * @see Writer#javaBeanToFile(List, String)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToFile(List<T> list, String realPath) throws IOException {
		writer().javaBeanToFile((List<Object>) list, realPath);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该方法是根据collation参数来获取JavaBean中的数据的</b>
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @throws IOException
	 * @see Writer#javaBeanToFile(List, String)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToFile(List<T> list, String realPath, String[] collation) throws IOException {
		writer().javaBeanToFile((List<Object>) list, realPath, collation);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, ExcelType)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, ExcelType type) throws IOException {
		writer().mapToFile(list, realPath, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, String[], ExcelType)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, String[] collation, ExcelType type) throws IOException {
		writer().mapToFile(list, realPath, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, TitleAlias[], ExcelType)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias, ExcelType type) throws IOException {
		writer().mapToFile(list, realPath, titleAlias, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是有序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, TitleAlias[], String[], ExcelType)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias,
			String[] collation, ExcelType type) throws IOException {
		writer().mapToFile(list, realPath, titleAlias, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see Writer#javaBeanToFile(List, String, ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToFile(List<T> list, String realPath, ExcelType type) throws IOException {
		writer().javaBeanToFile((List<Object>) list, realPath, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该方法是根据collation参数来获取JavaBean中的数据的</b>
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see Writer#javaBeanToFile(List, String, String[], ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToFile(List<T> list, String realPath, String[] collation, ExcelType type) throws IOException {
		writer().javaBeanToFile((List<Object>) list, realPath, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see Writer#mapToOutputStream(List, OutputStream, ExcelType)
	 */
	public static void writeMapToOutputStream(List<Map<String, Object>> list, OutputStream out, ExcelType type) throws IOException {
		writer().mapToOutputStream(list, out, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see Writer#mapToOutputStream(List, OutputStream, String[], ExcelType)
	 */
	public static void writeMapToOutputStream(List<Map<String, Object>> list, OutputStream out, String[] collation, ExcelType type) throws IOException {
		writer().mapToOutputStream(list, out, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see Writer#mapToOutputStream(List, OutputStream, TitleAlias[], ExcelType)
	 */
	public static void writeMapToOutputStream(List<Map<String, Object>> list, OutputStream out, TitleAlias[] titleAlias, ExcelType type) throws IOException {
		writer().mapToOutputStream(list, out, titleAlias, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see Writer#mapToOutputStream(List, OutputStream, TitleAlias[], String[], ExcelType)
	 */
	public static void writeMapToOutputStream(List<Map<String, Object>> list, OutputStream out, TitleAlias[] titleAlias,
			String[] collation, ExcelType type) throws IOException {
		writer().mapToOutputStream(list, out, titleAlias, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see Writer#javaBeanToOutputStream(List, OutputStream, ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToOutputStream(List<T> list, OutputStream out, ExcelType type) throws IOException {
		writer().javaBeanToOutputStream((List<Object>) list, out, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该方法是根据collation参数来获取JavaBean中的数据的</b>
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see Writer#javaBeanToOutputStream(List, OutputStream, String[], ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToOutputStream(List<T> list, OutputStream out, String[] collation, ExcelType type) throws IOException {
		writer().javaBeanToOutputStream((List<Object>) list, out, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#mapToBytes(List, ExcelType)
	 */
	public static byte[] writeMapToBytes(List<Map<String, Object>> list, ExcelType type) throws IOException {
		return writer().mapToBytes(list, type);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#mapToBytes(List, String[], ExcelType)
	 */
	public static byte[] writeMapToBytes(List<Map<String, Object>> list, String[] collation, ExcelType type) throws IOException {
		return writer().mapToBytes(list, collation, type);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#mapToBytes(List, TitleAlias[], ExcelType)
	 */
	public static byte[] writeMapToBytes(List<Map<String, Object>> list, TitleAlias[] titleAlias, ExcelType type) throws IOException {
		return writer().mapToBytes(list, titleAlias, type);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#mapToBytes(List, TitleAlias[], String[], ExcelType)
	 */
	public static byte[] writeMapToBytes(List<Map<String, Object>> list, TitleAlias[] titleAlias, String[] collation, ExcelType type) throws IOException {
		return writer().mapToBytes(list, titleAlias, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#javaBeanToBytes(List, ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> byte[] writeJavaBeanToBytes(List<T> list, ExcelType type) throws IOException {
		return writer().javaBeanToBytes((List<Object>) list, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该方法是根据collation参数来获取JavaBean中的数据的</b>
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#javaBeanToBytes(List, String[], ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> byte[] writeJavaBeanToBytes(List<T> list, String[] collation, ExcelType type) throws IOException {
		return writer().javaBeanToBytes((List<Object>) list, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToFile(List, String)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, String sheetName) throws IOException {
		writer(sheetName).mapToFile(list, realPath);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, String[])
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, String[] collation, String sheetName) throws IOException {
		writer(sheetName).mapToFile(list, realPath, collation);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, TitleAlias[])
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias, String sheetName) throws IOException {
		writer(sheetName).mapToFile(list, realPath, titleAlias);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, TitleAlias[], String[])
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias,
			String[] collation, String sheetName) throws IOException {
		writer(sheetName).mapToFile(list, realPath, titleAlias, collation);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#javaBeanToFile(List, String)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToFile(List<T> list, String realPath, String sheetName) throws IOException {
		writer(sheetName).javaBeanToFile((List<Object>) list, realPath);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该方法是根据collation参数来获取JavaBean中的数据的</b>
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#javaBeanToFile(List, String)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToFile(List<T> list, String realPath, String[] collation, String sheetName) throws IOException {
		writer(sheetName).javaBeanToFile((List<Object>) list, realPath, collation);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, ExcelType)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, ExcelType type, String sheetName) throws IOException {
		writer(sheetName).mapToFile(list, realPath, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, String[], ExcelType)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, String[] collation,
			ExcelType type, String sheetName) throws IOException {
		writer(sheetName).mapToFile(list, realPath, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, TitleAlias[], ExcelType)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias,
			ExcelType type, String sheetName) throws IOException {
		writer(sheetName).mapToFile(list, realPath, titleAlias, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是有序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToFile(List, String, TitleAlias[], String[], ExcelType)
	 */
	public static void writeMapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias,
			String[] collation, ExcelType type, String sheetName) throws IOException {
		writer(sheetName).mapToFile(list, realPath, titleAlias, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#javaBeanToFile(List, String, ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToFile(List<T> list, String realPath, ExcelType type, String sheetName) throws IOException {
		writer(sheetName).javaBeanToFile((List<Object>) list, realPath, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该方法是根据collation参数来获取JavaBean中的数据的</b>
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#javaBeanToFile(List, String, String[], ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToFile(List<T> list, String realPath, String[] collation,
			ExcelType type, String sheetName) throws IOException {
		writer(sheetName).javaBeanToFile((List<Object>) list, realPath, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToOutputStream(List, OutputStream, ExcelType)
	 */
	public static void writeMapToOutputStream(List<Map<String, Object>> list, OutputStream out, ExcelType type, String sheetName) throws IOException {
		writer(sheetName).mapToOutputStream(list, out, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToOutputStream(List, OutputStream, String[], ExcelType)
	 */
	public static void writeMapToOutputStream(List<Map<String, Object>> list, OutputStream out,
			String[] collation, ExcelType type, String sheetName) throws IOException {
		writer(sheetName).mapToOutputStream(list, out, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToOutputStream(List, OutputStream, TitleAlias[], ExcelType)
	 */
	public static void writeMapToOutputStream(List<Map<String, Object>> list, OutputStream out,
			TitleAlias[] titleAlias, ExcelType type, String sheetName) throws IOException {
		writer(sheetName).mapToOutputStream(list, out, titleAlias, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#mapToOutputStream(List, OutputStream, TitleAlias[], String[], ExcelType)
	 */
	public static void writeMapToOutputStream(List<Map<String, Object>> list, OutputStream out, TitleAlias[] titleAlias,
			String[] collation, ExcelType type, String sheetName) throws IOException {
		writer(sheetName).mapToOutputStream(list, out, titleAlias, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#javaBeanToOutputStream(List, OutputStream, ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToOutputStream(List<T> list, OutputStream out, ExcelType type, String sheetName) throws IOException {
		writer(sheetName).javaBeanToOutputStream((List<Object>) list, out, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该方法是根据collation参数来获取JavaBean中的数据的</b>
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @throws IOException
	 * @see Writer#javaBeanToOutputStream(List, OutputStream, String[], ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> void writeJavaBeanToOutputStream(List<T> list, OutputStream out, String[] collation,
			ExcelType type, String sheetName) throws IOException {
		writer(sheetName).javaBeanToOutputStream((List<Object>) list, out, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#mapToBytes(List, ExcelType)
	 */
	public static byte[] writeMapToBytes(List<Map<String, Object>> list, ExcelType type, String sheetName) throws IOException {
		return writer(sheetName).mapToBytes(list, type);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#mapToBytes(List, String[], ExcelType)
	 */
	public static byte[] writeMapToBytes(List<Map<String, Object>> list, String[] collation,
			ExcelType type, String sheetName) throws IOException {
		return writer(sheetName).mapToBytes(list, collation, type);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * @param list 写入Excel的数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#mapToBytes(List, TitleAlias[], ExcelType)
	 */
	public static byte[] writeMapToBytes(List<Map<String, Object>> list, TitleAlias[] titleAlias,
			ExcelType type, String sheetName) throws IOException {
		return writer(sheetName).mapToBytes(list, titleAlias, type);
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此方法写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#mapToBytes(List, TitleAlias[], String[], ExcelType)
	 */
	public static byte[] writeMapToBytes(List<Map<String, Object>> list, TitleAlias[] titleAlias, String[] collation,
			ExcelType type, String sheetName) throws IOException {
		return writer(sheetName).mapToBytes(list, titleAlias, collation, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#javaBeanToBytes(List, ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> byte[] writeJavaBeanToBytes(List<T> list, ExcelType type, String sheetName) throws IOException {
		return writer(sheetName).javaBeanToBytes((List<Object>) list, type);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该方法是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该方法是根据collation参数来获取JavaBean中的数据的</b>
	 * @param <T>
	 * @param list 写入Excel的数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @param sheetName 设置Excel文件的sheet名称
	 * @return 字节数组
	 * @throws IOException
	 * @see Writer#javaBeanToBytes(List, String[], ExcelType)
	 */
	@SuppressWarnings("unchecked")
	public static <T> byte[] writeJavaBeanToBytes(List<T> list, String[] collation, ExcelType type, String sheetName) throws IOException {
		return writer(sheetName).javaBeanToBytes((List<Object>) list, collation, type);
	}
	
}
