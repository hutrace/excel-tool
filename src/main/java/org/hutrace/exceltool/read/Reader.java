package org.hutrace.exceltool.read;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.hutrace.exceltool.annotation.ExcelField;
import org.hutrace.exceltool.exception.NotFoundSheetException;
import org.hutrace.exceltool.pojo.ExcelType;
import org.hutrace.exceltool.pojo.TitleAlias;
import org.hutrace.exceltool.utils.ExcelCommon;

/**
 * <p>Excel读取工具类
 * <p>可以将Excel读取成Map、Map别名、JavaBean等类型的List
 * <p>默认读取Excel的第一个sheet
 * <p>你可已使用{@link #setSheetName(String)}来指定读取sheet的名称
 * <p>此类只做到行数据读取，使用{@link ReaderResolver}来对行数据进行解析
 * @author HuTrace
 * @since 1.8
 * @version 1.0
 * @time 2020年1月14日
 * @see ReaderResolver
 * @see AbstractReaderResolver
 * @see MapResolver
 * @see MapAliasResolver
 * @see JavaBeanResolver
 */
public class Reader extends ExcelCommon {
	
	/**
	 * 读取的sheet名称
	 */
	private String sheetName;
	
	/**
	 * 设置读取Excel的sheet的名称
	 * @param sheetName 读取Excel的sheet的名称
	 */
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param realPath 根据realPath的后缀判断Excel的类型，如果没有后缀或后缀错误，将会抛出空指针异常。
	 * @return MapList
	 * @throws IOException
	 * @see MapResolver
	 */
	public List<Map<String, Object>> toMap(String realPath) throws IOException {
		return toMap(new FileInputStream(realPath), getType(realPath));
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
	 * @see MapResolver
	 */
	public List<Map<String, Object>> toMap(InputStream in, ExcelType type) throws IOException {
		ReaderResolver resolver = new MapResolver();
		read(in, type, resolver);
		return resolver.data();
	}
	
	/**
	 * <p>读取Excel数据
	 * <p>Excel的第一行数据为标题（注意：标题必须是字符串）
	 * <p>Excel的标题作为Map的键，可通过设置{@link #titleAlias}来实现在Map中的key值。
	 * <p>Map的长度(size)也将会使用标题的个数来确定
	 * @param realPath 根据realPath的后缀判断Excel的类型，如果没有后缀或后缀错误，将会抛出空指针异常。
	 * @return MapList
	 * @throws IOException
	 * @see MapAliasResolver
	 */
	public List<Map<String, Object>> toMap(String realPath, TitleAlias[] titleAlias) throws IOException {
		return toMap(new FileInputStream(realPath), getType(realPath), titleAlias);
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
	 * @see MapAliasResolver
	 */
	public List<Map<String, Object>> toMap(InputStream in, ExcelType type, TitleAlias[] titleAlias) throws IOException {
		ReaderResolver resolver = new MapAliasResolver(titleAlias);
		read(in, type, resolver);
		return resolver.data();
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
	 * @see JavaBeanResolver
	 */
	public <T> List<T> toJavaBean(String realPath, Class<T> clazs) throws IOException {
		return toJavaBean(new FileInputStream(realPath), getType(realPath), clazs);
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
	 * @see JavaBeanResolver
	 */
	public <T> List<T> toJavaBean(InputStream in, ExcelType type, Class<T> clazs) throws IOException {
		ReaderResolver resolver = new JavaBeanResolver<T>(clazs);
		read(in, type, resolver);
		return resolver.data();
	}
	
	/**
	 * <p>读取输入流中的数据
	 * <p>使用输入流构造{@link Workbook}
	 * <p>并取得需要获取数据的{@link Sheet}
	 * @param in 输入流
	 * @param type 文件类型
	 * @param resolver 数据解析器
	 * @throws IOException
	 */
	public void read(InputStream in, ExcelType type, ReaderResolver resolver) throws IOException {
		Workbook workbook = createWorkbook(in, type);
		Sheet sheet = sheet(workbook);
		reading(sheet, resolver);
	}
	
	/**
	 * <p>读取sheet的数据
	 * @param sheet 需要读取的{@link Sheet}对象
	 * @param resolver 数据解析器
	 * @throws IOException
	 */
	public void reading(Sheet sheet, ReaderResolver resolver) throws IOException {
		resolver.title(sheet.getRow(0));
		int rowNum = sheet.getLastRowNum();
		for(int i = 1; i <= rowNum; i++) {
			resolver.row(sheet.getRow(i), i);
		}
	}
	
	/**
	 * <p>获取{@link Sheet}对象
	 * <p>如果{@link #sheetName}为空，则获取Excel的第一个sheet
	 * <p>否则就寻找{@link #sheetName}对应的sheet
	 * <p>如果找不到，则会抛出{@link NotFoundSheetException}异常
	 * @param workbook {@link Workbook}对象
	 * @return {@link Sheet}对象
	 */
	private Sheet sheet(Workbook workbook) {
		if(sheetName == null) {
			return workbook.getSheetAt(0);
		}else {
			Sheet sheet = workbook.getSheet(sheetName);
			if(sheet == null) {
				throw new NotFoundSheetException("sheet with the name [" + sheetName + "] was not found");
			}
			return sheet;
		}
	}
	
}
