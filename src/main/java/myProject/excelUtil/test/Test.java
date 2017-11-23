package myProject.excelUtil.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;

import jxl.Sheet;
import jxl.Workbook;
import jxl.format.Colour;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import myProject.excelUtil.pojo.OrderForm;
import myProject.excelUtil.util.ExcelUtil;
import myProject.excelUtil.util.Validator;

public class Test {

	public static void main(String[] args) throws Exception {
//		String path = Test.class.getResource("import_fail.xls").toURI().getPath();
//		File file = new File(path);
//
//		System.out.println(checkExcelTitleAndSort(file));
//
//		System.out.println(checkExcelContent(file));
//
//		excelToEntity(file);

		exportExcelFile(loadData(), "d:" + File.separator + "export.xls");
	}

	/**
	 * 验证Excel文件首部的列名及排列顺序是否跟定义的OrderForm类一致
	 * 
	 * @param file
	 * @return
	 */
	public static boolean checkExcelTitleAndSort(File file) {
		InputStream stream = null;
		Workbook rwb = null;
		Boolean check = false;
		try {
			stream = new FileInputStream(file);
			// 获取Excel文件对象
			rwb = Workbook.getWorkbook(stream);
			// 获取文件的指定工作表 默认的第一个
			Sheet sheet = rwb.getSheet(0);
			Map<Integer, String> titleAndSortMap = ExcelUtil.getInstance().getExcelFieldName(OrderForm.class);
			check = ExcelUtil.getInstance().equalsArrays(sheet, titleAndSortMap);
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (stream != null) {
				try {
					stream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (rwb != null) {
				rwb.close();
			}
		}
		return check;
	}

	/**
	 * 验证Excel文件的内容格式是否正确
	 * 
	 * @param file
	 * @return
	 */
	public static boolean checkExcelContent(File file) {
		InputStream stream = null;
		Workbook rwb = null;
		Boolean check = false;
		try {
			stream = new FileInputStream(file);
			// 获取Excel文件对象
			rwb = Workbook.getWorkbook(stream);
			// 获取文件的指定工作表 默认的第一个
			Sheet sheet = rwb.getSheet(0);
			// 如有验证失败，该方法会返回错字段的字段名称
			String result = ExcelUtil.getInstance().checkExcelContent(sheet, OrderForm.class);

			// 如果没有返回错误字段，表示验证通过
			if (!Validator.isEffective(result)) {
				check = true;
			} else {
				System.out.println(result);
			}
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (stream != null) {
				try {
					stream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (rwb != null) {
				rwb.close();
			}
		}
		return check;
	}

	/**
	 * 将Excel文件的内容转换成实体类
	 * 
	 * @param file
	 */
	public static void excelToEntity(File file) {
		InputStream stream = null;
		Workbook rwb = null;
		try {
			stream = new FileInputStream(file);
			// 获取Excel文件对象
			rwb = Workbook.getWorkbook(stream);
			// 获取文件的指定工作表 默认的第一个
			Sheet sheet = rwb.getSheet(0);
			List<OrderForm> orderForms = ExcelUtil.getInstance().importExcelToEntity(sheet, OrderForm.class);
			System.out.println(orderForms);
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (stream != null) {
				try {
					stream.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
			if (rwb != null) {
				rwb.close();
			}
		}
	}

	/**
	 * 将实体类的内容输出到Excel文件里
	 * 
	 * @param list
	 * @param filePath
	 */
	public static <T> void exportExcelFile(List<T> list, String filePath) {
		OutputStream os = null;
		try {
			os = new FileOutputStream(new File(filePath));
			List<Map<Integer, String>> dataList = ExcelUtil.getInstance().exportExcel(list);
			WritableWorkbook wbook = Workbook.createWorkbook(os); // 建立excel文件
			String tmptitle = "订单信息";
			WritableSheet wsheet = wbook.createSheet(tmptitle, 0); // sheet名称
			// 设置excel标题
			WritableFont wfont = new WritableFont(WritableFont.ARIAL, 16, WritableFont.BOLD, false,
					UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
			WritableCellFormat wcfFC = new WritableCellFormat(wfont);
			wcfFC.setBackground(Colour.AQUA);
			wsheet.addCell(new Label(1, 0, tmptitle, wcfFC));
			wfont = new jxl.write.WritableFont(WritableFont.ARIAL, 14, WritableFont.BOLD, false,
					UnderlineStyle.NO_UNDERLINE, Colour.BLACK);
			wcfFC = new WritableCellFormat(wfont);
			// 开始生成主体内容
			// new Label(横坐标_X，纵坐标_Y,打印的名称)
			for (int i = 0; i < dataList.size(); i++) {
				for (int j = 0; j < dataList.get(i).size(); j++) {
					wsheet.addCell(new Label(j, i, dataList.get(i).get(j)));
				}
			}

			// 写入文件
			wbook.write();
			// 主体内容生成结束
			wbook.close();
		} catch (Exception ex) {
			ex.printStackTrace();
		} finally {
			if (os != null) {
				try {
					os.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	public static List<OrderForm> loadData() {
		List<OrderForm> orderForms = new ArrayList<>();
		OrderForm o = new OrderForm();
		o.setCreateDate(new Date());
		o.setCustomerName("张三");
		o.setOrderNo("1122334455");
		o.setOrderAmt(BigDecimal.valueOf(1000d));
		o.setDiscountAmt(BigDecimal.valueOf(100d));
		o.setIsMember(true);
		orderForms.add(o);
		OrderForm o2 = new OrderForm();
		o2.setCreateDate(new Date());
		o2.setCustomerName("李四");
		o2.setOrderNo("1234567890");
		o2.setOrderAmt(BigDecimal.valueOf(500d));
		o2.setDiscountAmt(BigDecimal.valueOf(20d));
		o2.setIsMember(false);
		orderForms.add(o2);
		return orderForms;
	}

}
