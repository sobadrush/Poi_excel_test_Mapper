USE DB_Emp_Dept;
--DROP TABLE emp_TB3;
--DROP TABLE dept_TB3;
CREATE TABLE dept_TB3
(
    deptno	int IDENTITY(10,10) PRIMARY KEY  not null,
	dname	varchar(14),
	loc		varchar(13)
);
CREATE TABLE emp_TB3
(
    empno	 int IDENTITY(7001,1) PRIMARY KEY not null,
	ename	 varchar(14),
	job		 varchar(13),
	hiredate  date,
	deptno int not NULL FOREIGN KEY REFERENCES dept_TB3(deptno)
	--Hibernate 做關聯查詢其實可以不用在實體表格設定foreign-key
);

insert into dept_TB3 values ('財務部','臺灣台北');
insert into dept_TB3 values ('研發部','臺灣新竹');
insert into dept_TB3 values ('業務部','美國紐約');
insert into dept_TB3 values ('生管部','中國上海');

insert into emp_TB3 values ('king','president','1981-11-17',10);
insert into emp_TB3 values ('blake','manager','1981-05-01',30);
insert into emp_TB3 values ('clark','manager','1981-01-09',10);
insert into emp_TB3 values ('jones','manager','1981-04-02',20);
insert into emp_TB3 values ('martin','salesman','1981-09-28',30);
insert into emp_TB3 values ('allen','salesman','1981-02-2',30);
insert into emp_TB3 values ('turner','salesman','1981-09-28',30);
insert into emp_TB3 values ('james','clerk','1981-12-03',30);
insert into emp_TB3 values ('ward','salesman','1981-02-22',30);
insert into emp_TB3 values ('ford','analyst','1981-12-03',20);
insert into emp_TB3 values ('smith','clerk','1980-12-17',20);
insert into emp_TB3 values ('scott','analyst','1981-12-09',20);
insert into emp_TB3 values ('adams','clerk','1983-01-12',20);
insert into emp_TB3 values ('miller','clerk','1982-01-23',10);


