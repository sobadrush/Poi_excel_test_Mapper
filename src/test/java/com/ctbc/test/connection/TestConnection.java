package com.ctbc.test.connection;

import java.sql.SQLException;

import javax.sql.DataSource;

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

@RunWith(SpringJUnit4ClassRunner.class)
@ContextConfiguration(classes = { RootConfig.class })
@FixMethodOrder(MethodSorters.NAME_ASCENDING)
@ActiveProfiles("sqlite_env")
public class TestConnection {

	@Autowired
	private DataSource ds;
	
	@Test
	@Ignore
	@Rollback(true)
	public void test001() throws SQLException {
		System.out.println("============== test001 ==============");
		System.err.println(ds.getConnection().getMetaData().getDatabaseProductName());
	}

}
