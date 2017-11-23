package myProject.excelUtil.util;

import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.regex.Pattern;

public class Validator {
	
	public static boolean isEffective(String paramString) {
		return (paramString != null) && (!"".equals(paramString)) && (!" ".equals(paramString))
				&& (!"null".equals(paramString)) && (!"\n".equals(paramString));
	}
	
	public static boolean IsNumber(String paramString) {
		if (paramString == null)
			return false;
		return match("-?[0-9]*$", paramString);
	}
	
	public static boolean isValidDate(String str, String formatStr) {
		boolean convertSuccess = true;
		// 指定日期格式为四位年/两位月份/两位日期，注意yyyy/MM/dd区分大小写；
		SimpleDateFormat format = new SimpleDateFormat(formatStr);
		try {
			// 设置lenient为false.
			// 否则SimpleDateFormat会比较宽松地验证日期，比如2007/02/29会被接受，并转换成2007/03/01
			format.setLenient(false);
			format.parse(str);
		} catch (ParseException e) {
			// e.printStackTrace();
			// 如果throw java.text.ParseException或者NullPointerException，就说明格式不对
			convertSuccess = false;
		}
		return convertSuccess;
	}

	
	public static boolean match(String paramString1, String paramString2) {
		if (paramString2 == null) {
			return false;
		}
		return Pattern.compile(paramString1).matcher(paramString2).matches();
	}
}
