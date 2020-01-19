package org.hutrace.exceltool.write;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.hutrace.exceltool.annotation.ExcelField;
import org.hutrace.exceltool.pojo.ExcelType;
import org.hutrace.exceltool.pojo.TitleAlias;
import org.hutrace.exceltool.utils.ExcelCommon;

/**
 * <p>写入Excel工具类
 * <p>可以将Map、JavaBean类型的List集合写入到Excel
 * <p>可以通过{@link #setSheetName(String)}指定创建sheet的名称，默认为"sheet1"
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 */
public class Writer extends ExcelCommon {
	
	/**
	 * 储存的sheet名称，默认为"sheet1"
	 */
	private String sheetName = "sheet1";
	
	/**
	 * 设置创建sheet的名称，默认为"sheet1"
	 * @param sheetName sheet的名称
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link #mapToFile(List, String, String[])}
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @throws IOException
	 * @see MapResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToFile(List<Map<String, Object>> list, String realPath) throws IOException {
		mapToFile(list, realPath, getType(realPath));
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @throws IOException
	 * @see MapCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToFile(List<Map<String, Object>> list, String realPath, String[] collation) throws IOException {
		mapToFile(list, realPath, collation, getType(realPath));
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link #mapToFile(List, String, TitleAlias[], String[])}
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @throws IOException
	 * @see MapAliasResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias) throws IOException {
		mapToFile(list, realPath, titleAlias, getType(realPath));
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @throws IOException
	 * @see MapAliasCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias, String[] collation) throws IOException {
		mapToFile(list, realPath, titleAlias, collation, getType(realPath));
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * <p>当然，你可以使用{@link #javaBeanToFile(List, String, String[])}，它可以自定义顺序
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @throws IOException
	 * @see JavaBeanResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void javaBeanToFile(List<Object> list, String realPath) throws IOException {
		OutputStream out = new FileOutputStream(realPath);
		javaBeanToOutputStream(list, out, getType(realPath));
		out.close();
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该解析器是根据collation参数来获取JavaBean中的数据的</b>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @throws IOException
	 * @see JavaBeanCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void javaBeanToFile(List<Object> list, String realPath, String[] collation) throws IOException {
		OutputStream out = new FileOutputStream(realPath);
		javaBeanToOutputStream(list, out, collation, getType(realPath));
		out.close();
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link #mapToFile(List, String, String[])}
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see MapResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToFile(List<Map<String, Object>> list, String realPath, ExcelType type) throws IOException {
		OutputStream out = new FileOutputStream(realPath);
		mapToOutputStream(list, out, type);
		out.close();
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see MapCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToFile(List<Map<String, Object>> list, String realPath, String[] collation, ExcelType type) throws IOException {
		OutputStream out = new FileOutputStream(realPath);
		mapToOutputStream(list, out, collation, type);
		out.close();
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link #mapToFile(List, String, TitleAlias[], String[])}
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see MapAliasResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias, ExcelType type) throws IOException {
		OutputStream out = new FileOutputStream(realPath);
		mapToOutputStream(list, out, titleAlias, type);
		out.close();
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see MapAliasCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToFile(List<Map<String, Object>> list, String realPath, TitleAlias[] titleAlias,
			String[] collation, ExcelType type) throws IOException {
		OutputStream out = new FileOutputStream(realPath);
		mapToOutputStream(list, out, titleAlias, collation, type);
		out.close();
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * <p>当然，你可以使用{@link #javaBeanToFile(List, String, String[])}，它可以自定义顺序
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see JavaBeanResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void javaBeanToFile(List<Object> list, String realPath, ExcelType type) throws IOException {
		OutputStream out = new FileOutputStream(realPath);
		javaBeanToOutputStream(list, out, type);
		out.close();
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该解析器是根据collation参数来获取JavaBean中的数据的</b>
	 * @param list 写入Excel的数据
	 * @param realPath 写入文件的绝对路径
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式，而不需要通过realPath参数的后缀去解析。
	 * @throws IOException
	 * @see JavaBeanCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void javaBeanToFile(List<Object> list, String realPath, String[] collation, ExcelType type) throws IOException {
		OutputStream out = new FileOutputStream(realPath);
		javaBeanToOutputStream(list, out, collation, type);
		out.close();
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link #mapToFile(List, String, String[])}
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see MapResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToOutputStream(List<Map<String, Object>> list, OutputStream out, ExcelType type) throws IOException {
		write(list.size(), type, new MapResolver(list)).write(out);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see MapCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToOutputStream(List<Map<String, Object>> list, OutputStream out, String[] collation, ExcelType type) throws IOException {
		write(list.size(), type, new MapCollationResolver(list, collation)).write(out);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link #mapToFile(List, String, TitleAlias[], String[])}
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see MapAliasResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToOutputStream(List<Map<String, Object>> list, OutputStream out, TitleAlias[] titleAlias, ExcelType type) throws IOException {
		write(list.size(), type, new MapAliasResolver(list, titleAlias)).write(out);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see MapAliasCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void mapToOutputStream(List<Map<String, Object>> list, OutputStream out, TitleAlias[] titleAlias,
			String[] collation, ExcelType type) throws IOException {
		write(list.size(), type, new MapAliasCollationResolver(list, titleAlias, collation)).write(out);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * <p>当然，你可以使用{@link #javaBeanToFile(List, String, String[])}，它可以自定义顺序
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see JavaBeanResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void javaBeanToOutputStream(List<Object> list, OutputStream out, ExcelType type) throws IOException {
		write(list.size(), type, new JavaBeanResolver(list)).write(out);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该解析器是根据collation参数来获取JavaBean中的数据的</b>
	 * @param list 写入Excel的数据
	 * @param out 输出流，向输出流中写入Excel数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @throws IOException
	 * @see JavaBeanCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public void javaBeanToOutputStream(List<Object> list, OutputStream out, String[] collation, ExcelType type) throws IOException {
		write(list.size(), type, new JavaBeanCollationResolver(list, collation)).write(out);
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link #mapToFile(List, String, String[])}
	 * @param list 写入Excel的数据
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see MapResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public byte[] mapToBytes(List<Map<String, Object>> list, ExcelType type) throws IOException {
		if(type == ExcelType.XLS) {
			HSSFWorkbook workbook = (HSSFWorkbook) write(list.size(), type, new MapResolver(list));
			return workbook.getBytes();
		}else {
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			mapToOutputStream(list, out, type);
			byte[] bytes = out.toByteArray();
			out.close();
			return bytes;
		}
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see MapCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public byte[] mapToBytes(List<Map<String, Object>> list, String[] collation, ExcelType type) throws IOException {
		if(type == ExcelType.XLS) {
			HSSFWorkbook workbook = (HSSFWorkbook) write(list.size(), type, new MapCollationResolver(list, collation));
			return workbook.getBytes();
		}else {
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			mapToOutputStream(list, out, collation, type);
			byte[] bytes = out.toByteArray();
			out.close();
			return bytes;
		}
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是无序的（针对整体而言无序）
	 * <p>它是通过{@link Map#keySet()}与{@link Map#values()}实现的，所有是整体而言无序的
	 * <p>你不能确定哪一列排在第一，哪一列排在最后。
	 * <p>如果你想要列有序，使用{@link #mapToFile(List, String, TitleAlias[], String[])}
	 * @param list 写入Excel的数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see MapAliasResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public byte[] mapToBytes(List<Map<String, Object>> list, TitleAlias[] titleAlias, ExcelType type) throws IOException {
		if(type == ExcelType.XLS) {
			HSSFWorkbook workbook = (HSSFWorkbook) write(list.size(), type, new MapAliasResolver(list, titleAlias));
			return ((HSSFWorkbook) workbook).getBytes();
		}else {
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			mapToOutputStream(list, out, titleAlias, type);
			byte[] bytes = out.toByteArray();
			out.close();
			return bytes;
		}
	}

	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>可以通过{@link TitleAlias}数组对Excel的标题进行设置，{@link TitleAlias#getAlias()}对应map的key，{@link TitleAlias#getTitle()}作为标题
	 * <p>{@link TitleAlias}数组没有个数要求，如果没有设置的，则使用map的key作为标题
	 * <p>此解析器写入Excel的列是有序序的
	 * <p>它是通过collation参数的顺序获取Map中的值
	 * <p>所以Excel的列数是根据collation参数确定的（可以过滤不需要的属性）
	 * @param list 写入Excel的数据
	 * @param titleAlias 标题别名数组，通过它可以设置写入Excel的标题
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see MapAliasCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public byte[] mapToBytes(List<Map<String, Object>> list, TitleAlias[] titleAlias, String[] collation, ExcelType type) throws IOException {
		if(type == ExcelType.XLS) {
			HSSFWorkbook workbook = (HSSFWorkbook) write(list.size(), type, new MapAliasCollationResolver(list, titleAlias, collation));
			return workbook.getBytes();
		}else {
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			mapToOutputStream(list, out, titleAlias, collation, type);
			byte[] bytes = out.toByteArray();
			out.close();
			return bytes;
		}
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照JavaBean的字段顺序进行解析，生成Excel列的顺序会按照JavaBean中字段的顺序而定
	 * <p>当然，你可以使用{@link #javaBeanToFile(List, String, String[])}，它可以自定义顺序
	 * @param list 写入Excel的数据
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see JavaBeanResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public byte[] javaBeanToBytes(List<Object> list, ExcelType type) throws IOException {
		if(type == ExcelType.XLS) {
			HSSFWorkbook workbook = (HSSFWorkbook) write(list.size(), type, new JavaBeanResolver(list));
			return workbook.getBytes();
		}else {
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			javaBeanToOutputStream(list, out, type);
			byte[] bytes = out.toByteArray();
			out.close();
			return bytes;
		}
	}
	
	/**
	 * <p>向Excel文件中写入数据，如果文件不存在则创建，存在则覆盖
	 * <p>JavaBean可以使用{@link ExcelField}注解设置Excel的标题，如果不设置，则采用JavaBean的字段名称
	 * <p>该解析器是按照collation参数的顺序进行解析，生成Excel列的顺序会按照collation参数的顺序而定
	 * <p><b>需要注意：可以通过设置collation参数来对Excel写入列的个数做处理（过滤不需要的字段），该解析器是根据collation参数来获取JavaBean中的数据的</b>
	 * @param list 写入Excel的数据
	 * @param collation 排序规则，对列进行排序，可以指定map中的key写入Excel的顺序，也可以过滤不需要的属性
	 * @param type 它可以对Excel指定格式
	 * @return 字节数组
	 * @throws IOException
	 * @see JavaBeanCollationResolver
	 * @see AbstractWriterResolver
	 * @see WriterResolver
	 */
	public byte[] javaBeanToBytes(List<Object> list, String[] collation, ExcelType type) throws IOException {
		if(type == ExcelType.XLS) {
			HSSFWorkbook workbook = (HSSFWorkbook) write(list.size(), type, new JavaBeanCollationResolver(list, collation));
			return workbook.getBytes();
		}else {
			ByteArrayOutputStream out = new ByteArrayOutputStream();
			javaBeanToOutputStream(list, out, type);
			byte[] bytes = out.toByteArray();
			out.close();
			return bytes;
		}
	}
	
	/**
	 * <p>向Excel写入数据
	 * <p>创建{@link Workbook}与{@link Sheet}后，调用{@link #writing(Sheet, int, WriterResolver)}
	 * @param size 数据的长度，这里可以理解为Excel需要创建的多少行（除去标题行）
	 * @param type Excel的类型，它决定了{@link Workbook}、{@link Sheet}、{@link Row}、{@link Cell}等Excel相关的实现类
	 * @param resolver 写入数据解析器
	 * @return {@link Workbook}
	 * @throws IOException
	 */
	public Workbook write(int size, ExcelType type, WriterResolver resolver) throws IOException {
		Workbook workbook = createWorkbook(type);
		Sheet sheet = workbook.createSheet(sheetName);
		writing(sheet, size, resolver);
		return workbook;
	}
	
	/**
	 * <p>向Excel的{@link Sheet}中写入数据
	 * <p>调用{@link WriterResolver#title(Row)}与{@link WriterResolver#row(Row, int)}具体实现写入过程。
	 * @param sheet Excel的{@link Sheet}
	 * @param size 数据的长度，这里可以理解为Excel需要创建的多少行（除去标题行）
	 * @param resolver 写入数据解析器
	 */
	public void writing(Sheet sheet, int size, WriterResolver resolver) {
		resolver.title(sheet.createRow(0));
		for(int i = 0; i < size; i++) {
			resolver.row(sheet.createRow(i + 1), i);
		}
	}
	
}
