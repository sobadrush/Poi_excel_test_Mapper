package com.ctbc.model.mapper.iii;

import java.util.List;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Param;

import com.ctbc.model.EmpVO;

@Mapper
public interface EmpMapper {

	public List<EmpVO> getEmps();

	public EmpVO getEmps(@Param("empno") int empno);

	public EmpVO getEmps(@Param("empHiredate") java.sql.Date hDate);

	public List<EmpVO> getEmpsWithDept();

	public EmpVO getEmpsWithDept(@Param("empNOgg") Integer empid);

	public List<EmpVO> getEmpsWithDept(
			@Param("empNOgg") Integer empid,
			@Param("empNamePart") String enamePart,
			@Param("empHiredates") java.sql.Date[] hDates);
}
