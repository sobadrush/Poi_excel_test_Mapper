package com.ctbc.model;

import java.io.Serializable;

import org.apache.ibatis.type.Alias;

@Alias(value="fuckEmpVO")
public class EmpVO implements Serializable {

	private static final long serialVersionUID = 1L;

	private Integer empNo;
	private java.sql.Date empHireDate;
	private String empName;
	private String empJob;
	private DeptVO deptVO;
	
	public EmpVO() {
	}
	
	public Integer getEmpNo() {
		return empNo;
	}
	public void setEmpNo(Integer empNo) {
		this.empNo = empNo;
	}
	public String getEmpName() {
		return empName;
	}
	public void setEmpName(String empName) {
		this.empName = empName;
	}
	public String getEmpJob() {
		return empJob;
	}
	public void setEmpJob(String empJob) {
		this.empJob = empJob;
	}
	public java.sql.Date getEmpHireDate() {
		return empHireDate;
	}
	public void setEmpHireDate(java.sql.Date empHireDate) {
		this.empHireDate = empHireDate;
	}
	public DeptVO getDeptVO() {
		return deptVO;
	}

	public void setDeptVO(DeptVO deptVO) {
		this.deptVO = deptVO;
	}

	@Override
	public String toString() {
		return "EmpVO [empHireDate=" + empHireDate + ", empNo=" + empNo + ", empName=" + empName + ", empJob=" + empJob + "]";
	}

}
