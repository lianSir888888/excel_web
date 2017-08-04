package com.ldy.drdc.action;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.io.FileUtils;
import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.jdom.Attribute;
import org.jdom.Document;
import org.jdom.Element;
import org.jdom.input.SAXBuilder;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.ldy.drdc.model.ColumnInfo;
import com.ldy.drdc.model.ImportData;
import com.ldy.drdc.model.ImportDataDetail;
import com.ldy.drdc.model.Template;
import com.ldy.drdc.service.ImportDataService;

import net.sf.json.JSONArray;

@Controller
@RequestMapping(value="/importdata")
public class ImportDataController {
	
	private String templateName;
	
	@Resource
	private ImportDataService importDataService;
	
	@RequestMapping(value="/templates")
	public void templates(HttpServletResponse response){
		response.setContentType("text/html;charset=utf-8");
		List<Template> list = new ArrayList<Template>();
		Template t = new Template();
		t.setTemplateId("student");
		t.setTemplateName("student");
		list.add(t);
		try {
			response.getWriter().write(JSONArray.fromObject(list).toString());
		} catch (IOException e) {
			e.printStackTrace();
		}
		
	}
	
	@RequestMapping(value="/download")
	public void download(HttpServletRequest request,HttpServletResponse response,String templateId) throws IOException{
		//生成导入模板文件
		createTemplate(request,templateId);
		//得到要下载的文件名
		String fileName = templateName + ".xls";
		//上传文件都是保存目录
		String fileSaveRootPath = request.getServletContext().getRealPath("/template");
		//通过文件名找出文件的所在目录
		String path = findFileSavePathByFileName(fileName,fileSaveRootPath);
		//得到要下载的文件
		File file = new File(path + "//" + fileName);
//		//处理文件名
//		String realname = fileName.substring(fileName.indexOf("_")+1);
		//设置响应头，控制浏览器下载该文件
		response.setHeader("content-disposition", "attachment;filename=" + URLEncoder.encode(fileName, "UTF-8"));
		//读取要下载的文件，保存到文件输入流
		FileInputStream in = new FileInputStream(path + "//" + fileName);
		//创建输出流
		OutputStream out = response.getOutputStream();
		//创建缓冲区
		byte buffer[] = new byte[1024];
		int len = 0;
		//循环将输入流中的内容读取到缓冲区当中
		while((len=in.read(buffer))>0){
			//输出缓冲区的内容到浏览器，实现文件下载
			out.write(buffer, 0, len);
			}
		//关闭文件输入流
		in.close();
		//关闭输出流
		out.close();
	
	}
	
	
	/**
	* @Method: findFileSavePathByFileName
	* @Description: 通过文件名和存储上传文件根目录找出要下载的文件的所在路径
	* @Anthor:LDY
	* @param filename 要下载的文件名
	* @param saveRootPath 上传文件保存的根目录，也就是/WEB-INF/upload目录
	* @return 要下载的文件的存储目录
	*/
	public String findFileSavePathByFileName(String filename,String saveRootPath){
		int hashcode = filename.hashCode();
//		int dir1 = hashcode&0xf; //0--15
//		int dir2 = (hashcode&0xf0)>>4; //0-15
		String dir = saveRootPath /*+ "//" + dir1 + "//" + dir2*/; //upload\2\3 upload\3\5
		File file = new File(dir);
		if(!file.exists()){
			//创建目录
			file.mkdirs();
		}
		return dir;
	}
	
	private void createTemplate(HttpServletRequest request,String templateId) {
		//获取解析xml文件路径
		String path = request.getServletContext().getRealPath("/template");
		File file = new File(path,templateId+".xml");
		SAXBuilder builder = new SAXBuilder();
		try {
			//解析xml文件
			Document parse = builder.build(file);
			//创建Excel
			HSSFWorkbook wb = new HSSFWorkbook();
			//创建sheet
			HSSFSheet sheet = wb.createSheet("Sheet0");
			
			//获取xml文件跟节点
			Element root = parse.getRootElement();
			//获取模板名称
			templateName = root.getAttribute("name").getValue();
			int rownum = 0;
			int column = 0;
			//设置列宽
			Element colgroup = root.getChild("colgroup");
			setColumnWidth(sheet,colgroup);
			
			//设置标题
			Element title = root.getChild("title");
			List<Element> trs = title.getChildren("tr");
			for (int i = 0; i < trs.size(); i++) {
				Element tr = trs.get(i);
				List<Element> tds = tr.getChildren("td");
				HSSFRow row = sheet.createRow(rownum);
				HSSFCellStyle cellStyle = wb.createCellStyle();
				cellStyle.setAlignment(HSSFCellStyle.ALIGN_CENTER);
				for(column = 0;column <tds.size();column ++){
					Element td = tds.get(column);
					HSSFCell cell = row.createCell(column);
					Attribute rowSpan = td.getAttribute("rowspan");
					Attribute colSpan = td.getAttribute("colspan");
					Attribute value = td.getAttribute("value");
					if(value !=null){
						String val = value.getValue();
						cell.setCellValue(val);
						int rspan = rowSpan.getIntValue() - 1;
						int cspan = colSpan.getIntValue() -1;
						
						//设置字体
						HSSFFont font = wb.createFont();
						font.setFontName("仿宋_GB2312");
						font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);//字体加粗
//						font.setFontHeight((short)12);
						font.setFontHeightInPoints((short)12);
						cellStyle.setFont(font);
						cell.setCellStyle(cellStyle);
						//合并单元格居中
						sheet.addMergedRegion(new CellRangeAddress(rspan, rspan, 0, cspan));
					}
				}
				rownum ++;
			}
			//设置表头
			Element thead = root.getChild("thead");
			trs = thead.getChildren("tr");
			for (int i = 0; i < trs.size(); i++) {
				Element tr = trs.get(i);
				HSSFRow row = sheet.createRow(rownum);
				List<Element> ths = tr.getChildren("th");
				for(column = 0;column < ths.size();column++){
					Element th = ths.get(column);
					Attribute valueAttr = th.getAttribute("value");
					HSSFCell cell = row.createCell(column);
					if(valueAttr != null){
						String value =valueAttr.getValue();
						cell.setCellValue(value);
					}
				}
				rownum++;
			}
			
			//设置数据区域样式
			Element tbody = root.getChild("tbody");
			Element tr = tbody.getChild("tr");
			int repeat = tr.getAttribute("repeat").getIntValue();
			
			List<Element> tds = tr.getChildren("td");
			for (int i = 0; i < repeat; i++) {
				HSSFRow row = sheet.createRow(rownum);
				for(column =0 ;column < tds.size();column++){
					Element td = tds.get(column);
					HSSFCell cell = row.createCell(column);
					setType(wb,cell,td);
				}
				rownum++;
			}
			
			//生成Excel导入模板
			File tempFile = new File(path, templateName + ".xls");
			tempFile.delete();
			tempFile.createNewFile();
			FileOutputStream stream = FileUtils.openOutputStream(tempFile);
			wb.write(stream);
			stream.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	
	/**
	 * 测试单元格样式
	 * @author David
	 * @param wb
	 * @param cell
	 * @param td
	 */
	private static void setType(HSSFWorkbook wb, HSSFCell cell, Element td) {
		Attribute typeAttr = td.getAttribute("type");
		String type = typeAttr.getValue();
		HSSFDataFormat format = wb.createDataFormat();
		HSSFCellStyle cellStyle = wb.createCellStyle();
		if("NUMERIC".equalsIgnoreCase(type)){
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			Attribute formatAttr = td.getAttribute("format");
			String formatValue = formatAttr.getValue();
			formatValue = StringUtils.isNotBlank(formatValue)? formatValue : "#,##0.00";
			cellStyle.setDataFormat(format.getFormat(formatValue));
		}else if("STRING".equalsIgnoreCase(type)){
			cell.setCellValue("");
			cell.setCellType(HSSFCell.CELL_TYPE_STRING);
			cellStyle.setDataFormat(format.getFormat("@"));
		}else if("DATE".equalsIgnoreCase(type)){
			cell.setCellType(HSSFCell.CELL_TYPE_NUMERIC);
			cellStyle.setDataFormat(format.getFormat("yyyy-m-d"));
		}else if("ENUM".equalsIgnoreCase(type)){
			CellRangeAddressList regions = 
				new CellRangeAddressList(cell.getRowIndex(), cell.getRowIndex(), 
						cell.getColumnIndex(), cell.getColumnIndex());
			Attribute enumAttr = td.getAttribute("format");
			String enumValue = enumAttr.getValue();
			//加载下拉列表内容
			DVConstraint constraint = 
				DVConstraint.createExplicitListConstraint(enumValue.split(","));
			//数据有效性对象
			HSSFDataValidation dataValidation = new HSSFDataValidation(regions, constraint);
			wb.getSheetAt(0).addValidationData(dataValidation);
		}
		cell.setCellStyle(cellStyle);
	}

	/**
	 * 设置列宽
	 * @author David
	 * @param sheet
	 * @param colgroup
	 */
	private static void setColumnWidth(HSSFSheet sheet, Element colgroup) {
		List<Element> cols = colgroup.getChildren("col");
		for (int i = 0; i < cols.size(); i++) {
			Element col = cols.get(i);
			Attribute width = col.getAttribute("width");
			String unit = width.getValue().replaceAll("[0-9,\\.]", "");
			String value = width.getValue().replaceAll(unit, "");
			int v=0;
			if(StringUtils.isBlank(unit) || "px".endsWith(unit)){
				v = Math.round(Float.parseFloat(value) * 37F);
			}else if ("em".endsWith(unit)){
				v = Math.round(Float.parseFloat(value) * 267.5F);
			}
			sheet.setColumnWidth(i, v);
		}
	}
	
	/**
	 * 文件上传处理方法
	 * @author David
	 */
	public void upload(HttpServletResponse response,HttpServletRequest request,String templateId,File fileInput){
		response.setContentType("text/html;charset=utf-8");
		
		SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd hh:mm:ss");
		String dateNow = df.format(new Date());
		
		//保存主表信息
		ImportData importData = new ImportData();
		importData.setImportid(String.valueOf(System.currentTimeMillis()));
		importData.setImportDataType(templateId);
		importData.setImportDate(dateNow);
		importData.setImportStatus("1");//导入成功
		importData.setHandleDate(null);
		importData.setHandleStatus("0");//未处理
		importDataService.saveImportData(importData);
		
		
		try {
			//读取Excel文件
			HSSFWorkbook wb = new HSSFWorkbook(FileUtils.openInputStream(fileInput));
			HSSFSheet sheet = wb.getSheetAt(0);
			
			//获取模板文件
			String path = request.getServletContext().getRealPath("/template");
			path = path + "\\" + templateId + ".xml";
			File file = new File(path);
			
			//解析xml模板文件
			SAXBuilder builder = new SAXBuilder();
			Document parse =  builder.build(file);
			Element root = parse.getRootElement();
			Element tbody = root.getChild("tbody");
			Element tr = tbody.getChild("tr");
			List<Element> children = tr.getChildren("td");
			//解析excel开始行，开始列
			int firstRow = tr.getAttribute("firstrow").getIntValue();
			int firstCol = tr.getAttribute("firstcol").getIntValue();
			//获取excel最后一行行号
			int lastRowNum = sheet.getLastRowNum();
			//循环每一行处理数据
			for (int i = firstRow; i <= lastRowNum; i++) {
				//初始化明细数据
				ImportDataDetail importDataDetail = new ImportDataDetail();
				importDataDetail.setImportId(importData.getImportid());
				importDataDetail.setCgbz("0");//未处理
				//读取某行
				HSSFRow row = sheet.getRow(i);
				//判断改行是否为空
				if(isEmptyRow(row)){
					continue;
				}
				int lastCellNum = row.getLastCellNum();
				//如果非空行，则取所有单元格的值
				for (int j = firstCol; j <lastCellNum; j++) {
					Element td = children.get(j-firstCol);
					HSSFCell cell = row.getCell(j);
					//如果单元格为null,继续处理下一个cell
					if(cell == null){
						continue;
					}
					//获取单元格属性值
					String value = getCellValue(cell,td);
					//导入明细实体赋值
					if(StringUtils.isNotBlank(value)){
						if(value.indexOf("#000")>=0){
							String[] info = value.split(",");
							importDataDetail.setHcode(info[0]);
							importDataDetail.setMsg(info[1]);
							BeanUtils.setProperty(importDataDetail, "col" + j, info[2]);
						}else{
							BeanUtils.setProperty(importDataDetail, "col" + j, value);
						}
					}
					
				}
				importDataService.saveImportDataDetail(importDataDetail);
			}
			
			String str = "{\"status\":\"ok\",\"message\":\"导入成功！\"}";
			response.getWriter().write(str);
		} catch (Exception e) {
			String str = "{\"status\":\"noOk\",\"message\":\"导入失败！\"}";
			try {
				response.getWriter().write(str);
			} catch (IOException e1) {
				e1.printStackTrace();
			}
			e.printStackTrace();
		}
	}
	
	/**
	 * 判断某行是否为空
	 * @author David
	 * @return
	 */
	private boolean isEmptyRow(HSSFRow row) {
		boolean flag = true;
		for (int i = 0; i < row.getLastCellNum(); i++) {
			HSSFCell cell = row.getCell(i);
			if(cell != null){
				if(StringUtils.isNotBlank(cell.toString())){
					return false;
				}
			}
		}
		
		return flag;
	}
	/**
	 * 获取单元格值，并且进行校验
	 * @author David
	 * @param cell
	 * @param td
	 * @return
	 */
	private String getCellValue(HSSFCell cell, Element td) {
		//首先获取单元格位置
		int i = cell.getRowIndex() + 1;
		int j = cell.getColumnIndex()+1;
		String returnValue = "";//返回值
		
		try {
			//获取模板文件对单元格格式限制
			String type = td.getAttribute("type").getValue();
			boolean isNullAble = td.getAttribute("isnullable").getBooleanValue();
			int maxlength = 9999;
			
			if(td.getAttribute("maxlength")!=null){
				maxlength = td.getAttribute("maxlength").getIntValue();
			}
			String value = null;
			//根据格式取出单元格的值
			switch (cell.getCellType()) {
				case HSSFCell.CELL_TYPE_STRING:{
					value = cell.getStringCellValue();
					break;
				}
				case HSSFCell.CELL_TYPE_NUMERIC:{
					if("datetime,date".indexOf(type)>=0){
						Date date = cell.getDateCellValue();
						SimpleDateFormat df = new SimpleDateFormat("yyyy-MM-dd");
						value = df.format(date);
					}else{
						double numericCellValue = cell.getNumericCellValue();
						value = String.valueOf(numericCellValue);
					}
					break;
				}
			}
			
			//对非空、长度进行校验
			if(!isNullAble && StringUtils.isBlank(value)){
				//错误编码,错误位置原因,单位格的值
				returnValue = "#0001,第" + i + "行第" +j +"列不能为空！," + value;
			}else if(StringUtils.isNotBlank(value) && (value.length()>maxlength)){
				returnValue = "#0002,第" + i + "行第" +j +"列长度超过最大长度！," + value;
			}else{
				returnValue =  value;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return returnValue;
	}
	/**
	 * 动态获取表头信息
	 * @author David
	 */
	public void columns(HttpServletResponse response,HttpServletRequest request,String templateId){
		response.setContentType("text/html;charset=utf-8");
		//获取表头信息
		List<ColumnInfo> list = getColumns(request,templateId);
		//转换json对象返回
		String json ="["+ JSONArray.fromObject(list).toString() + "]";
		try {
			response.getWriter().write(json);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**
	 * 动态获取表头
	 * @author David
	 * @return
	 */
	private List<ColumnInfo> getColumns(HttpServletRequest request,String templateId) {
		List<ColumnInfo> list = new ArrayList<ColumnInfo>();
		//获取模板文件
		String path = request.getServletContext().getRealPath("/template");
		path = path + "\\" + templateId + ".xml";
		File file = new File(path);
		
		//解析模板文件
		SAXBuilder builder = new SAXBuilder();
		try {
			Document parse = builder.build(file);
			Element root = parse.getRootElement();
			Element thead = root.getChild("thead");
			Element tr = thead.getChild("tr");
			List<Element> children = tr.getChildren();
			
			ColumnInfo c = new ColumnInfo();
			//添加处理标志、失败代码，失败说明
			c = createColumnInfo("cgbz","处理标志",120,"center");
			list.add(c);
			c = createColumnInfo("hcode","失败代码",120,"center");
			list.add(c);
			c = createColumnInfo("msg","失败说明",120,"center");
			list.add(c);
			for (int i = 0; i < children.size(); i++) {
				Element th = children.get(i);
				String value = th.getAttribute("value").getValue();
				c = createColumnInfo("col"+i,value,120,"center");
				list.add(c);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		} 
		
		return list;
	}
	/**
	 * 创建column对象
	 * @author David
	 * @param string
	 * @param string2
	 * @param i
	 * @param string3
	 */
	private ColumnInfo createColumnInfo(String fieldId, String title, int width,
			String align) {
		ColumnInfo c = new ColumnInfo();
		c.setField(fieldId);
		c.setTitle(title);
		c.setWidth(width);
		c.setAlign(align);
		return c;
	}
	/**
	 * 获取明细数据
	 * @author David
	 */
	public void columndatas(HttpServletResponse response,String importDataId){
		response.setContentType("text/html;charset=utf-8");
		//获取明细数据
		List<ImportDataDetail> list = importDataService.getImportDataDetailsByMainId(importDataId);
		String json = "{\"total\":"+list.size()+", \"rows\":"+JSONArray.fromObject(list).toString()+"}";
		try {
			response.getWriter().write(json);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	/**
	 * 确认导入
	 * @author David
	 */
	public void doimport(HttpServletResponse response,String importDataId){
		response.setContentType("text/html;charset=utf-8");
		//将导入的明细数据已到student表中
		importDataService.saveStudents(importDataId);
		//修改主表、明细表处理标志及时间
		SimpleDateFormat sf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		String dateNow = sf.format(new Date());
		importDataService.updImportDataStatus(dateNow, importDataId);
		importDataService.updImportDataDetailStatus(importDataId);
		String str = "{\"status\":\"ok\",\"message\":\"确认成功！\"}";
		try {
			response.getWriter().write(str);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

}
