package com.ctbc.model;

import java.io.Serializable;
import java.util.Set;

import org.apache.ibatis.type.Alias;

@Alias(value="fuckDeptVO")
public class DeptVO implements Serializable {

	private static final long serialVersionUID = 1L;

	private int deptNo;
	private String deptName;
	private String deptLoc;
	private Set<EmpVO> emps;
	
	public DeptVO() {
	}

	public int getDeptNo() {
		return deptNo;
	}

	public void setDeptNo(int deptNo) {
		this.deptNo = deptNo;
	}

	public String getDeptName() {
		return deptName;
	}

	public void setDeptName(String deptName) {
		this.deptName = deptName;
	}

	public String getDeptLoc() {
		return deptLoc;
	}

	public void setDeptLoc(String deptLoc) {
		this.deptLoc = deptLoc;
	}

	public Set<EmpVO> getEmps() {
		return emps;
	}

	public void setEmps(Set<EmpVO> emps) {
		this.emps = emps;
	}

	@Override
	public String toString() {
		return "DeptVO [deptNo=" + deptNo + ", deptName=" + deptName + ", deptLoc=" + deptLoc + "]";
	}
	
}
