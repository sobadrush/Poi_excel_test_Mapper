package com.ctbc.test.connection;

import java.sql.SQLException;
import java.util.List;

import org.junit.FixMethodOrder;
import org.junit.Ignore;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.MethodSorters;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.test.annotation.Rollback;
import org.springframework.test.context.ActiveProfiles;
import org.springframework.test.context.ContextConfiguration;
import org.springframework.test.context.junit4.SpringJUnit4ClassRunner;

import com.ctbc.config.RootConfig;
import com.ctbc.model.EmpDAO_Mapper;
import com.ctbc.model.EmpVO;

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(classes = { RootConfig.class })
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
@ActiveProfiles("sqlite_env")
public class TestEmpMapper {

	@Autowired
	private EmpDAO_Mapper empDAO_Mapper;
	
	@Test
	@Ignore
	@Rollback(true)
	public void test001() throws SQLException {
		System.out.println("============== test001 ==============");
		System.err.println(empDAO_Mapper);
	}
	
	@Test
//	@Ignore
	@Rollback(true)
	public void test002() throws SQLException {
		System.out.println("============== test002 ==============");
		List<EmpVO> empList = empDAO_Mapper.getAll();
		for (EmpVO empVO : empList) {
			System.out.println(" >>> " + empVO);
		}
	}

}











