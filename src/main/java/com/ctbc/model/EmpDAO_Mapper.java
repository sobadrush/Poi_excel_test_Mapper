package com.ctbc.model;

import java.util.List;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;
import org.springframework.stereotype.Repository;

import com.ctbc.config.RootConfig;
import com.ctbc.model.mapper.iii.EmpMapper;

@Repository
public class EmpDAO_Mapper {

	@Autowired
	private EmpMapper empMapper;

	public List<EmpVO> getAll() {
		return empMapper.getEmps();
	}

	public EmpVO getOneById(int empno) {
		return empMapper.getEmps(empno);
	}

	public EmpVO getEmpsByHireDate(java.sql.Date hireDate) {
		return empMapper.getEmps(hireDate);
	}

	public List<EmpVO> getEmpsWithDeptData() {
		return empMapper.getEmpsWithDept();
	}

//	public EmpVO getEmpWithDeptDataById(int empid) {
//		return empMapper.getEmpsWithDept(empid);
//	}
	
	public List<EmpVO> getEmpWithDeptDataById(int empid) {
		return empMapper.getEmpsWithDept(empid,null,null);
	}

	public List<EmpVO> getEmpWithDeptDataByAmbigusName(Integer empid,String part_ename) {
		return empMapper.getEmpsWithDept(empid,part_ename,null);
	}

	public List<EmpVO> getEmpWithDeptDataByAmbigusName(java.sql.Date[] hiredates) {
		return empMapper.getEmpsWithDept(null,null,hiredates);
	}
	
	public static void main(String[] args) {
		System.setProperty("spring.profiles.active", "sqlite_env");
		ConfigurableApplicationContext context = new AnnotationConfigApplicationContext(RootConfig.class);
		EmpDAO_Mapper dao = context.getBean("empDAO_Mapper", EmpDAO_Mapper.class);
//		【查一筆 by Hiredate】
//		EmpVO vo = dao.getEmpsByHireDate(java.sql.Date.valueOf("1980-12-17"));
//		System.out.println(vo.toString()+" 部門資訊："+(vo.getDeptVO().toString()));
//		【查一筆 by id】
//		System.out.println(dao.getOneById(7001));
//		【查多筆】
		List<EmpVO> alist = dao.getAll();
		for (EmpVO empVO : alist) {
			System.out.println(empVO.toString());
		}
//		【關聯查詢 - 查多筆】
//		List<EmpVO> alist = dao.getEmpsWithDeptData();
//		for (EmpVO empvo : alist) {
//			System.out.println(empvo.toString() + " --- 部門資訊：" + (empvo.getDeptVO().toString()));
//		}
//		【關聯查詢 - 查一筆】
//		EmpVO empvo = dao.getEmpWithDeptDataById(7001);
//		System.out.println(empvo.toString() + " --- 部門資訊：" + (empvo.getDeptVO().toString()));
		// =================================
		// === 使用 <choose>...<when>... ===
		// ==============↓↓↓↓===============
//		List<EmpVO> emps = dao.getEmpWithDeptDataById(7001);
//		EmpVO empvo = emps.get(0);		
//		System.out.println(empvo.toString() + " 部門資訊 : " + empvo.getDeptVO().toString());
		// ==============↑↑↑↑===============
//		【關聯查詢 - 根據empName模糊查詢多筆】
//		List<EmpVO> empList = dao.getEmpWithDeptDataByAmbigusName(null,"K");
//		for (EmpVO empVO : empList) {
//			System.out.println(empVO);
//		}
//		【關聯查詢 - 根據 empHiredate[] 查詢多筆】
//		List<EmpVO> empList 
//			= dao.getEmpWithDeptDataByAmbigusName(
//					new java.sql.Date[]{
//						java.sql.Date.valueOf("1981-11-17"),			
//						java.sql.Date.valueOf("1981-05-01")			
//					});
//		for (EmpVO empVO : empList) {
//			System.out.println(empVO);
//		}
		context.close();
	}
}
