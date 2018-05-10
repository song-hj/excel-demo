package com.songhj.util;

import java.text.SimpleDateFormat;
import java.util.Date;

public class CommUtil {

	public static boolean isNotNull(Object obj) {
		if ((obj != null) && (!obj.toString().equals(""))) {
			return true;
		}
		return false;
	}
	
	public static Date formatDate(String s, String format) {
		Date d = null;
		try {
			SimpleDateFormat dFormat = new SimpleDateFormat(format);
			d = dFormat.parse(s);
		} catch (Exception localException) {
		}
		return d;
	}
	
	public static String formatShortDate(Object v) {
		if (v == null)
			return null;
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
		return df.format(v);
	}
	
	public static String formatTime(String format, Object v) {
		if (v == null)
			return null;
		if (v.equals(""))
			return "";
		SimpleDateFormat df = new SimpleDateFormat(format);
		return df.format(v);
	}
}
