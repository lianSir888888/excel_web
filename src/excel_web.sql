/*
SQLyog 企业版 - MySQL GUI v8.14 
MySQL - 5.5.40 
*********************************************************************
*/
/*!40101 SET NAMES utf8 */;

create table `t_student` (
	`SID` double ,
	`STUNUM` varchar (135),
	`STUNAME` varchar (135),
	`STUAGE` varchar (135),
	`STUSEX` varchar (135),
	`STUBIRTHDAY` varchar (135),
	`STUHOBBY` varchar (135)
); 
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('1','001','张三','100','男','1990','打球');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('2','002','李四','180','男','1999','听歌');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('3','003','王五','170','男','1998','画画');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('4','004','马六','160','男','1997','码代码');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('5','005','陈七','150','男 ','1996','看书');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('6','006','孙八','140','男','1995','看电影');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('7','007','潘九','130','男','1994','nz');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('8','008','温十','120','男','1993','cf');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('9','009','成龙','110','男','1991','lol');
insert into `t_student` (`SID`, `STUNUM`, `STUNAME`, `STUAGE`, `STUSEX`, `STUBIRTHDAY`, `STUHOBBY`) values('10','010','范冰冰','190','女','2000','写书');
