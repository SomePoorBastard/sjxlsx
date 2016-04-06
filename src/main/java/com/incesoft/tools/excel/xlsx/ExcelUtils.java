package com.incesoft.tools.excel.xlsx;

import com.incesoft.tools.excel.support.XLSXReaderSupport;
import org.apache.commons.codec.digest.DigestUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang.StringUtils;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

public class ExcelUtils {

	public static Map<Integer, String> builtInFormats = new HashMap<>();
	static {
		builtInFormats.put(0, "general");
		builtInFormats.put(1, "0");
		builtInFormats.put(2, "0.00");
		builtInFormats.put(3, "#,##0");
		builtInFormats.put(4, "#,##0.00");
		builtInFormats.put(9, "0%");
		builtInFormats.put(10, "0.00%");
		builtInFormats.put(11, "0.00e+00");
		builtInFormats.put(12, "# ?/?");
		builtInFormats.put(13, "# ??/??");
		builtInFormats.put(14, "mm-dd-yy");
		builtInFormats.put(15, "d-mmm-yy");
		builtInFormats.put(16, "d-mmm");
		builtInFormats.put(17, "mmm-yy");
		builtInFormats.put(18, "h:mm am/pm");
		builtInFormats.put(19, "h:mm:ss am/pm");
		builtInFormats.put(20, "h:mm");
		builtInFormats.put(21, "h:mm:ss");
		builtInFormats.put(22, "m/d/yy h:mm");
		builtInFormats.put(37, "#,##0 ;(#,##0)");
		builtInFormats.put(38, "#,##0 ;[red](#,##0)");
		builtInFormats.put(39, "#,##0.00;(#,##0.00)");
		builtInFormats.put(40, "#,##0.00;[red](#,##0.00)");
		builtInFormats.put(41, "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)");
		builtInFormats.put(42, "_(\"$\"* #,##0_);_(\"$* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)");
		builtInFormats.put(43, "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)");
		builtInFormats.put(44, "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)");
		builtInFormats.put(45, "mm:ss");
		builtInFormats.put(46, "[h]:mm:ss");
		builtInFormats.put(47, "mmss.0");
		builtInFormats.put(48, "##0.0e+0");
		builtInFormats.put(49, "@");
	}

	private static SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");

	/**
	 * Excel 2007+ using the OOXML format(actually is a zip)
	 * 
	 * @return
	 */
	public static boolean isOOXML(InputStream inputStream) {
		try {
			return inputStream.read() == 0x50 && inputStream.read() == 0x4b && inputStream.read() == 0x03
					&& inputStream.read() == 0x04;
		} catch (IOException e) {
			throw new RuntimeException(e);
		}
	}

	/**
	 * check excel version
	 * 
	 * @param file
	 * @return 'xlsx' for 07 or 'xls' for 03
	 */
	public static String getExcelExtensionName(File file) {
		FileInputStream stream = null;
		try {
			stream = new FileInputStream(file);
			return isOOXML(stream) ? "xlsx" : "xls";
		} catch (IOException e) {
			throw new RuntimeException(e);
		} finally {
			if (stream != null) {
				IOUtils.closeQuietly(stream);
			}
		}
	}

	public static String checksumZipContent(File f) {
		ZipFile zipFile = null;
		try {
			zipFile = new ZipFile(f);
			Enumeration<? extends ZipEntry> e = zipFile.entries();
			List<Long> crcs = new ArrayList<Long>();
			while (e.hasMoreElements()) {
				ZipEntry entry = e.nextElement();
				crcs.add(entry.getCrc());
			}
			return DigestUtils.shaHex(StringUtils.join(crcs, ""));
		} catch (Exception e) {
			throw new RuntimeException("", e);
		} finally {
			try {
				if (zipFile != null)
					zipFile.close();
			} catch (IOException ignored) {}
		}
	}

	/**
	 * Overload method
	 * @param input       Path to input XLSX file
	 * @param output      Path to output CSV file
	 * @param sheetNum    Sheet number to extract
	 * @param rowStart    Row index where the report begins
	 * @throws IOException
     */
	public static void toCSV(String input, String output, int sheetNum, int rowStart) throws IOException {
		toCSV(new File(input), new File(output), sheetNum, rowStart);
	}

	/**
	 * Convert from input XLSX file to CSV
	 * @param input       Input XLSX file
	 * @param output      Output CSV file
	 * @param sheetNum    Sheet number to extract
	 * @param rowStart    Row index where the report begins
	 */
	public static void toCSV(File input, File output, int sheetNum, int rowStart) throws IOException {
		try (
				XLSXReaderSupport rxs = new XLSXReaderSupport();
				FileOutputStream fo = new FileOutputStream(output);
		) {
			rxs.setInputFile(input);
			rxs.open(sheetNum);
			XLSXRowIterator it = rxs.rowIterator();
			// skip first 'rowStart' rows
			for (int i = 0; i < rowStart; i++) {
				it.nextRow();
			}
			while (it.nextRow()) {
				List<String> cells = new ArrayList<>();
				for (Cell cell : it.getCurRow()) {
					cells.add(getQuotedCellValue(cell));
				}
				String buf = StringUtils.join(cells, ',') + "\n";
				fo.write(buf.getBytes());
			}
		}
	}

	/**
	 * Get formatted date value or string value
	 * in quoted format, ready for CSV cell
	 * @param cell    XLSX cell to extract
	 * @return string in quoted format
     */
	public static String getQuotedCellValue(Cell cell) {
		String value = "";
		if (null != cell && null != cell.getValue()) {
			if (cell.isDate()) {
				Calendar cal = cell.getDateValue();
				if (null != cal) {
					value = sdf.format(cal.getTime());
				}
			} else {
				value = cell.getValue().trim();
			}
		}
		return "\"" + value + "\"";
	}
}