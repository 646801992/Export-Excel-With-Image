
/**
 * Class Name: NewExcelUtil.java
 * Description: 导出新版本Excel
 * @author wuhongwei01
 * Create Time: 2019年11月8日
 */

import java.awt.Graphics;
import java.awt.GraphicsConfiguration;
import java.awt.GraphicsDevice;
import java.awt.GraphicsEnvironment;
import java.awt.HeadlessException;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.Transparency;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.imageio.ImageIO;
import javax.swing.ImageIcon;

import org.apache.commons.collections.MapUtils;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class NewExcelUtil {

	/* 图片转码格式Map */
	private final static Map<String, Integer> typeMap = new HashMap<String, Integer>() {
		{
			put(".emf", XSSFWorkbook.PICTURE_TYPE_EMF);
			put(".wmf", XSSFWorkbook.PICTURE_TYPE_WMF);
			put(".pict", XSSFWorkbook.PICTURE_TYPE_PICT);
			put(".jpeg", XSSFWorkbook.PICTURE_TYPE_JPEG);
			put(".jpg", XSSFWorkbook.PICTURE_TYPE_JPEG);
			put(".dib", XSSFWorkbook.PICTURE_TYPE_DIB);
			put(".gif", XSSFWorkbook.PICTURE_TYPE_GIF);
			put(".tiff", XSSFWorkbook.PICTURE_TYPE_TIFF);
			put(".eps", XSSFWorkbook.PICTURE_TYPE_EPS);
			put(".bmp", XSSFWorkbook.PICTURE_TYPE_BMP);
			put(".wpg", XSSFWorkbook.PICTURE_TYPE_WPG);
			put(".png", XSSFWorkbook.PICTURE_TYPE_PNG);
		}
	};

	/**
	 * 导出Excel(包含图片的情况下使用)
	 * 
	 * @param sheetName   sheet名称
	 * @param excelHeader 标题
	 * @param parameter   字段信息
	 * @param list<map>        内容
	 * @param imageIndex  第几列是图片(从0开始计数)
	 * @return HSSFWorkbook
	 * @throws IOException
	 */
	public static XSSFWorkbook getXSSFWorkbook(String sheetName, String[] excelHeader, String[] parameter,
			List<Map<String, Object>> list, int imageIndex, String filePath) throws IOException {

		// 创建一个工作簿，对应文件
		XSSFWorkbook workBook = new XSSFWorkbook();

		// 创建一个sheet工作表
		XSSFSheet sheet = workBook.createSheet("sheet1");

		// 设置表头单元格样式
		XSSFCellStyle headstyle = workBook.createCellStyle();
		// 设置居中
		headstyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		headstyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		XSSFFont headFont = workBook.createFont();
		headFont.setFontHeight(14);
		headFont.setBoldweight(XSSFFont.BOLDWEIGHT_BOLD);
		headstyle.setFont(headFont);

		// 创建一般单元格样式
		XSSFCellStyle cellstyle = workBook.createCellStyle();
		cellstyle.setAlignment(XSSFCellStyle.ALIGN_CENTER);
		cellstyle.setVerticalAlignment(XSSFCellStyle.VERTICAL_CENTER);
		cellstyle.setWrapText(true);
		XSSFFont cellFont = workBook.createFont();
		cellFont.setFontHeight(11);
		cellstyle.setFont(cellFont);

		// 创建表头
		XSSFRow headRow = sheet.createRow(0);
		for (int i = 0; i < excelHeader.length; i++) {
			XSSFCell cell = headRow.createCell(i);
			cell.setCellValue(excelHeader[i]);
			cell.setCellStyle(headstyle);
			sheet.setColumnWidth(i, (30 * 256));
		}

		// 创建内容
		XSSFRow row = null;
		XSSFCell cell = null;
		for (int rowIndex = 0; rowIndex < list.size(); rowIndex++) {
			row = sheet.createRow(rowIndex + 1);
			row.setHeight((short) (40 * 20));
			Map<String, Object> dataMap = list.get(rowIndex);

			for (int paraIndex = 0; paraIndex < parameter.length; paraIndex++) {
				cell = row.createCell(paraIndex);
				// 图片列处理
				if (paraIndex == imageIndex) {
					if (dataMap != null) {
						String imageName = MapUtils.getString(dataMap, parameter[paraIndex]);
						URL photoFile = new URL(filePath + imageName);
						if (checkUrl(photoFile)) {
							//BufferedImage bufferedImage = ImageIO.read(photoFile);
							// bufferedImage = fixBackGround(bufferedImage);
							// ImageIO会导致图片变红
							java.awt.Image imageTookittitle = Toolkit.getDefaultToolkit().createImage(photoFile);
							BufferedImage bufferedImage = toBufferedImage(imageTookittitle);
							
							int width = bufferedImage.getWidth();
							int height = bufferedImage.getHeight();
							ByteArrayOutputStream byteArrayOut = new ByteArrayOutputStream();
							ImageIO.write(bufferedImage, getImageType(imageName), byteArrayOut);
							byte[] data = byteArrayOut.toByteArray();

							int anchorX = 0;
							int anchorY = 0;

							// 计算图片缩放比例
							anchorX = 1000;
							anchorY = (int) (1000 * ((double) height / (double) width));

							short rowHeight = 0;

							anchorX = anchorX * XSSFShape.EMU_PER_PIXEL;
							anchorY = anchorY * XSSFShape.EMU_PER_PIXEL;

							XSSFDrawing xssfDrawing = sheet.createDrawingPatriarch();
							XSSFClientAnchor anchor = new XSSFClientAnchor(0, 0, anchorX, anchorY, imageIndex,
									rowIndex + 1, imageIndex, rowIndex + 1);
							// 在电脑屏幕上 Excel默认行高度为13.5 (13.5/72)*96=18像素
							rowHeight = (short) ((245D / 96D) * 72 * 20 * ((double) height / (double) width));
							row.setHeight(rowHeight);
							try {
								xssfDrawing.createPicture(anchor,
										workBook.addPicture(data, getXSSFWorkbookPictureType(imageName)));
							} catch (Exception e) {
								cell.setCellValue("暂无图片");
								cell.setCellStyle(cellstyle);
							}
							cell.setCellValue("");
							cell.setCellStyle(cellstyle);
						} else {
							cell.setCellValue("暂无图片");
							cell.setCellStyle(cellstyle);
						}
					} else {
						cell.setCellValue("暂无图片");
						cell.setCellStyle(cellstyle);
					}
				} else {
					if (dataMap != null) {
						String s = MapUtils.getString(dataMap, parameter[paraIndex]);
						cell.setCellValue(s);
						cell.setCellStyle(cellstyle);
					} else {
						cell.setCellValue("");
						cell.setCellStyle(cellstyle);
					}
				}
			}
		}
		return workBook;
	}

	/**
	 * @Description: 根据数据库中存储的图片后缀获取对应Excel中图片格式代码
	 * @date 2019年11月8日
	 * @author Created by wuhongwei01
	 * @param Image Name
	 * @return XSSFWorkBook PICTURE_TYPE
	 */
	private static Integer getXSSFWorkbookPictureType(String imageName) {
		// 为空直接按照jpeg处理
		if (imageName == null || imageName == "") {
			return 5;
		}

		// 转换小写方便匹配
		String checkName = imageName.toLowerCase();

		// 利用正则表达式匹配图片扩展名
		String regex = "(\\.emf|\\.wmf|\\.pict|\\.jpeg|\\.jpg|\\.png|\\.dib|\\.gif|\\.tiff|\\.eps|\\.bmp|\\.wpg)";
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(checkName);
		String type = ".jpeg";
		if (matcher.find()) {
			type = matcher.group(0);
		} else {
			return 5;
		}

		return NewExcelUtil.typeMap.get(type);
	}

	private static String getImageType(String imageName) {
		// 为空直接按照jpeg处理
		if (imageName == null || imageName == "") {
			return "jpeg";
		}

		// 转换小写方便匹配
		String checkName = imageName.toLowerCase();

		// 利用正则表达式匹配图片扩展名
		String regex = "(\\.emf|\\.wmf|\\.pict|\\.jpeg|\\.png|\\.dib|\\.gif|\\.tiff|\\.eps|\\.bmp|\\.wpg)";
		Pattern pattern = Pattern.compile(regex);
		Matcher matcher = pattern.matcher(checkName);
		String type = ".jpeg";
		if (matcher.find()) {
			type = matcher.group(0);
		} else {
			return "jpeg";
		}

		return type.substring(1, type.length());
	}

	/**
	 * @Description: 检查URL可用性
	 * @date 2019年11月7日
	 * @author Created by wuhongwei01
	 * @param URL
	 * @return boolean
	 */
	private static boolean checkUrl(URL url) {
		/* 通过状态码检测 200则为验证通过 */
		boolean result = false;
		try {
			HttpURLConnection con = (HttpURLConnection) url.openConnection();
			int state = con.getResponseCode();
			if (state == 200) {
				result = true;
			}
		} catch (IOException e) {
			result = false;
		}
		return result;
	}

	public static BufferedImage toBufferedImage(Image image) {
	    if (image instanceof BufferedImage) {
	        return (BufferedImage) image;
	    }
	    // This code ensures that all the pixels in the image are loaded
	    image = new ImageIcon(image).getImage();
	    BufferedImage bimage = null;
	    GraphicsEnvironment ge = GraphicsEnvironment
	            .getLocalGraphicsEnvironment();
	    try {
	        int transparency = Transparency.OPAQUE;
	        GraphicsDevice gs = ge.getDefaultScreenDevice();
	        GraphicsConfiguration gc = gs.getDefaultConfiguration();
	        bimage = gc.createCompatibleImage(image.getWidth(null),
	                image.getHeight(null), transparency);
	    } catch (HeadlessException e) {
	        // The system does not have a screen
	    }
	    if (bimage == null) {
	        // Create a buffered image using the default color model
	        int type = BufferedImage.TYPE_INT_RGB;
	        bimage = new BufferedImage(image.getWidth(null),
	                image.getHeight(null), type);
	    }
	    // Copy image to buffered image
	    Graphics g = bimage.createGraphics();
	    // Paint the image onto the buffered image
	    g.drawImage(image, 0, 0, null);
	    g.dispose();
	    return bimage;
	}
}
