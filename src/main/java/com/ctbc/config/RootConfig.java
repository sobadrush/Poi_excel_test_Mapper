package com.ctbc.config;

import java.beans.PropertyVetoException;
import java.sql.Connection;
import java.sql.DatabaseMetaData;
import java.sql.SQLException;

import javax.sql.DataSource;

import org.apache.ibatis.session.SqlSessionFactory;
import org.mybatis.spring.annotation.MapperScan;
import org.springframework.beans.factory.annotation.Qualifier;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ConfigurableApplicationContext;
import org.springframework.context.annotation.AnnotationConfigApplicationContext;
import org.springframework.context.annotation.Bean;
import org.springframework.context.annotation.ComponentScan;
import org.springframework.context.annotation.Configuration;
import org.springframework.context.annotation.Profile;
import org.springframework.context.annotation.PropertySource;
import org.springframework.jdbc.datasource.DriverManagerDataSource;

import com.mchange.v2.c3p0.ComboPooledDataSource;

// 【Autowired 失敗】： http://www.mybatis.org/spring/zh/mappers.html

@Configuration
@ComponentScan(basePackages = { "com.ctbc.model", "com.ctbc.excel.util", "com.ctbc.model.mapper.iii" })
@PropertySource(value = { "classpath:jdbcConnection.properties" })
@MapperScan(basePackages = { "com.ctbc.model.mapper.iii" }, sqlSessionFactoryRef = "sqlSessionFactory")
public class RootConfig {

	@Profile("mssql_env")
	@Bean(value = "c3p0_DS", destroyMethod = "close")
	public DataSource c3p0Datasource(
			@Value("${jdbc.sqlserver.driverClassName}") String driverClassName,
			@Value("${jdbc.sqlserver.jdbcUrl}") String url,
			@Value("${jdbc.sqlserver.userName}") String username,
			@Value("${jdbc.sqlserver.userPassword}") String password) {
		// 【參考】: http://blog.csdn.net/l1028386804/article/details/51162560
		ComboPooledDataSource c3p0PooledDataSource = new ComboPooledDataSource();
		try {
			c3p0PooledDataSource.setDriverClass(driverClassName);
//			c3p0PooledDataSource.setDriverClass("com.microsoft.sqlserver.jdbc.SQLServerDriver");
		} catch (PropertyVetoException e) {
			System.out.println("c3p0Datasource 建立失敗");
			e.printStackTrace();
		}
		c3p0PooledDataSource.setJdbcUrl(url);
		c3p0PooledDataSource.setUser(username);
		c3p0PooledDataSource.setPassword(password);
		c3p0PooledDataSource.setMaxPoolSize(30);
		c3p0PooledDataSource.setInitialPoolSize(10);
		c3p0PooledDataSource.setIdleConnectionTestPeriod(60); //每60秒检查所有连接池中的空闲连接。Default: 0
		return c3p0PooledDataSource;
	}
	
	@Profile("sqlite_env")
	@Bean(value = "c3p0_DS", destroyMethod = "close")
	public DataSource c3p0DatasourceSqlite() throws PropertyVetoException {
		// 【參考】: http://blog.csdn.net/l1028386804/article/details/51162560
		ComboPooledDataSource c3p0PooledDataSource = new ComboPooledDataSource();
		
		String connectionString = "jdbc:sqlite:" + System.getProperty("user.dir") + "/" + "testDB.db";
		c3p0PooledDataSource.setJdbcUrl(connectionString);
		c3p0PooledDataSource.setDriverClass("org.sqlite.JDBC");
		
//		c3p0PooledDataSource.setJdbcUrl(url);
//		c3p0PooledDataSource.setUser(username);
//		c3p0PooledDataSource.setPassword(password);
		c3p0PooledDataSource.setMaxPoolSize(30);
		c3p0PooledDataSource.setInitialPoolSize(10);
		c3p0PooledDataSource.setIdleConnectionTestPeriod(60); //每60秒检查所有连接池中的空闲连接。Default: 0
		return c3p0PooledDataSource;
	}

	@Bean
	@Profile("mssql_env")
	public DataSource driverManagerDataSource(
			@Value("${jdbc.sqlserver.driverClassName}") String driverClassName,
			@Value("${jdbc.sqlserver.jdbcUrl}") String url,
			@Value("${jdbc.sqlserver.userName}") String username,
			@Value("${jdbc.sqlserver.userPassword}") String password) {
		DriverManagerDataSource ds = new DriverManagerDataSource();
		ds.setDriverClassName(driverClassName);
		ds.setUrl(url);
		ds.setUsername(username);
		ds.setPassword(password);
		return ds;
	}

	@Bean
	public SqlSessionFactory sqlSessionFactory(@Qualifier("c3p0_DS") DataSource ds, ApplicationContext apContext) throws Exception {
		System.out.println("連線池實作：" + ds);
		org.mybatis.spring.SqlSessionFactoryBean sqlSessionFactoryBean = new org.mybatis.spring.SqlSessionFactoryBean();
		sqlSessionFactoryBean.setDataSource(ds);
		sqlSessionFactoryBean.setMapperLocations(apContext.getResources("classpath*:XML_mappers/**.xml"));// 【Step 1】mapper xml 在的包
		sqlSessionFactoryBean.setTypeAliasesPackage("com.ctbc.model");
		sqlSessionFactoryBean.setTypeHandlersPackage("_02_typeHandler.CustomTimeStampHandler");// 型別轉換器
		return sqlSessionFactoryBean.getObject();
	}

//	@Bean // 可用 @mapperScan 
//	public MapperScannerConfigurer mapperScannerConfigurer() {
//		MapperScannerConfigurer mapperScannerConfigurer = new MapperScannerConfigurer();
//		mapperScannerConfigurer.setBasePackage("com.ctbc.model.mapper.iii");// 【Step 2】mapper 介面在的包
//		mapperScannerConfigurer.setSqlSessionFactoryBeanName("sqlSessionFactory");
//		return mapperScannerConfigurer;
//	}

	public static void main(String[] args) {
		System.setProperty("spring.profiles.active", "sqlite_env");
		// -- 測試 Datasource 連線
		ConfigurableApplicationContext context = new AnnotationConfigApplicationContext(RootConfig.class);
		DataSource ds = context.getBean("c3p0_DS", DataSource.class);
		Connection conn;
		try {
			conn = ds.getConnection();
			DatabaseMetaData dsMeta = conn.getMetaData();
			System.out.println("dsMeta - ProductName :  " + dsMeta.getDatabaseProductName());
		} catch (SQLException e) {
			e.printStackTrace();
		}
		context.close();
	}
}
