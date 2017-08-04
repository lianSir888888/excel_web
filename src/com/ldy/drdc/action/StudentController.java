package com.ldy.drdc.action;

import java.io.IOException;
import java.lang.reflect.Method;
import java.util.List;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.ldy.drdc.model.Student;
import com.ldy.drdc.service.StudentService;
import com.ldy.drdc.utils.ExportUtils;

import net.sf.json.JSONArray;

@Controller
@RequestMapping(value="/students")
public class StudentController {
	
	@Resource
	private StudentService studentService;
	
//	private Integer page;
//	
//	private Integer rows;
//	
//	private String sort;
//	
//	private String order;
	
	@RequestMapping(value="list")
	public void list(HttpServletResponse response,Integer page,Integer rows,String sort,String order){
		response.setContentType("text/html;charset=utf-8");
		List<Student>slist = getStudents(page,rows,sort,order);		
		int total = studentService.getTotal();
		String json = "{\"total\":"+total+" , " +
				"\"rows\":"+JSONArray.fromObject(slist).toString()+ "," +
				"\"className\":\"" + StudentController.class.getName() + "\"," +
				"\"methodName\":\"getStudents\"}";
		try {	
			response.getWriter().write(json);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	/**
	 * 获取学生信息
	 * @author ldy
	 * @return
	 */
	public List<Student> getStudents(Integer page,Integer rows,String sort,String order){
		List<Student> slist = studentService.list(page,rows,sort,order);
		
		return slist;
	}
	
	/**
	 * 导出前台列表为excel文件
	 * @author ldy
	 */
	@RequestMapping(value="export")
	public void export(HttpServletResponse response,String className,String titles,
			String methodName,String fields){
		response.setContentType("application/octet-stream");
		response.setHeader("Content-Disposition", "attachment;filename=export.xls");
		//创建Excel
		HSSFWorkbook wb = new HSSFWorkbook();
		HSSFSheet sheet = wb.createSheet("sheet0");
		try {
			//获取类
//			Class clazz = Class.forName(className);
//			Object o = clazz.newInstance();
//			Class[] types = {Integer.class,Integer.class,String.class,String.class};
//			Method m = clazz.getMethod(methodName,types);
//			List<Student> list = (List<Student>)m.invoke(o,1,10,"","");
			List<Student> list = getStudents(1,10,null,null);
			titles = new String(titles.getBytes("ISO-8859-1"),"UTF-8");
			ExportUtils.outputHeaders(titles.split(","), sheet);
			ExportUtils.outputColumns(fields.split(","), list, sheet, 1);
			//获取输出流，写入excel 并关闭
			ServletOutputStream out = response.getOutputStream();
			wb.write(out);
			out.flush();
			out.close();
			
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
