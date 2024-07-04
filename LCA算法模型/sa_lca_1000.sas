
/*********程序名称：能源工序处理*************************************************/
/*
DATE:20181023
AUTHOR:ma_lc
DATA SOURCE:ma_lc.sa_lc_0006 能源大类数据表
			ma_lc.Wh_lcdm0008:代码对照表
			ma_lc.Wh_lcdm0003：代码对照表
			wh_lc.wh_lca_0002:能源工序对照表
			ma_lc.Wh_lcdm0006:单位


矩阵G1：result_g1/data_g1：直接负荷
矩阵PB2:result_pb2
矩阵G2:result_g2：能源负荷
矩阵RB2:result_rb2
矩阵Gbcl:result_gbcl
矩阵G3:result_g3：中间产品负荷
G4：副产品内部收益
G5:运输负荷
G6:上游负荷
G7:

*/
/********************************************************************************/
/*signoff a1;
%let a1=190.2.242.171;
options comamid=tcp remote=a1;
filename rlink 'C:\Program Files\SASHome\SASFoundation\9.3\connect\saslink\tcpunix.scr';
signon a1;
libname ma_lc '/dw/sasdata/ma_lc' server=a1;
LIBNAME WH_LC DB2 Datasrc=dwsasdev SCHEMA=WH_LC USER=db2lc PASSWORD=db2lc81;
libname ztl "C:\ma_lc";*/

%lib(ma_lc);
%lib(wh_lc);

%macro edition;
%let datefrom=&date_f;
%let dateto=&date_e;
%let type=ny;
%let root1=DL;
%let shopsign_code1=DL;
data wh_lca_0002;
	set wh_lc.wh_lca_0002 ;
	where type="&type";
run;
data sa_lc_0006;
	set ma_lc.sa_lc_0006 ;
	where compress(year||mon) between substr("&datefrom",1,6) and substr("&dateto",1,6) ;
run;

proc sql;
	create table tmp1 as select sum(val) as val,item_id
	from sa_lc_0006 

	group by item_id;
quit;
/*勾连工序代码和类型*/

proc sql;
	create table tmp2 as
	select a.val,a.item_id,b.type_name,b.type_code
	from tmp1 a left join ma_lc.Wh_lcdm0008 b
	on substr(a.item_id,5,2)=b.type_code;
quit;



/*获取单位*/
proc sql;
	create table Wh_lcdm0006 as
	 select distinct jz_code,type_code,pro_code,unit
	 from  ma_lc.Wh_lcdm0006  
	where year=substr(substr("&datefrom",1,6),1,4);
quit;
proc sql;
	create table tmp2 as select a.*,b.unit from tmp2 a left join Wh_lcdm0006 b
	on a.item_id=compress(b.jz_code||b.type_code||b.pro_code);
quit;
/*合并一二三四高炉鼓风*/
data Wh_lcdm0003;
set ma_lc.Wh_lcdm0003;
 where process_code in ('LJ','SJ','GL','GF') and jz_code ne 'CO04' ;
run;

proc sql;
	create table tmp21 as
	select a.*,b.process_code||"00"||substr(a.item_id,5,5) as item_idl
	from tmp2 a,Wh_lcdm0003 b
	where b.jz_code=substr(a.item_id,1,4);
quit;

proc sql;
	create table tmp22 as
	select item_idl as item_id ,type_name,type_code,sum(val) as val,unit
	from tmp21
	group by item_idl,type_name,type_code,unit;
quit;
data tmp2;
	set tmp2 tmp22;
run;
proc sql;
		create table tmp2 as select a.*,b.jz_name,b.jz_code from tmp2 a
		left join ma_lc.Wh_lcdm0003 b on substr(a.item_id,1,4)=b.jz_code ;
quit;

data tmp2;
	set tmp2;
	if jz_code in ('') then jz_code=substr(item_id,1,4);
	if jz_name='' then do;
	if substr(item_id,1,2)='LJ' then jz_name='炼焦';
	else if substr(item_id,1,2)='SJ' then jz_name='烧结';
	else if substr(item_id,1,2)='GL' then jz_name='高炉';
	else if substr(item_id,1,2)='GF' then jz_name='鼓风';
	end;
run;
/*取能源工序*/
proc sql;
	create table tmp3 as select a.* from tmp2 a,wh_lca_0002 b
	where a.jz_code=b.jz_code 
	union all
	select a.* from tmp2 a where a.jz_code in ('Z002');
quit;

/*统一单位:kg  m3   kWh */
data tmp5(drop=unit);
	set tmp3;
	if unit in ('t') then val=val*1000;
	else if unit in ('m3') then val=val;
	else if unit in ('kwh') then val=val;
	else if unit in ('km3') then val=val*1000;
	else if unit in ('kg') then val=val;
	else val=val;
run;
/*能源工序*/

	/*将制氧工序分解成三个工序：制氧Z021，制氩Z022，制氮Z023*/
	proc sql;
		create table wh_lca_0004 as select avg(rate) as rate,pro_code,pro_name from wh_lc.wh_lca_0004
		where ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6)
		group by pro_code,pro_name;
	quit;

	proc sql;
		create table temp as select a.val*b.rate as val,a.item_id,b.pro_name from tmp5 a,wh_lca_0004 b
		where a.jz_code='Z002' and a.type_code='04' and substr(a.item_id,7,3)=b.pro_code;
	quit;
	proc sql;
		create table temp_all as select sum(val) as val_all from temp;
	quit;

	proc sql;
		create table temp1 as select a.val/b.val_all as rate,a.item_id,a.pro_name from temp a left join temp_all b on 1=1;
	quit;

	proc sql;
		create table tmp6 as select jz_name,jz_code,type_name,type_code,item_id,val from tmp5 where jz_code not in ('Z002')
		union all
		select '制氧气' as jz_name,'Z021' as jz_code,a.type_name,a.type_code,'Z021'||substr(a.item_id,5) as item_id,
		case when a.type_code='01' then a.val*b.rate else a.val end as val  from tmp5 a join temp1 b on 1=1 where a.jz_code='Z002' and b.pro_name='氧气'
		and substr(a.item_id,7,3) not in ('061','119')
		union all
		select '制氩气' as jz_name,'Z022' as jz_code,a.type_name,a.type_code,'Z022'||substr(a.item_id,5) as item_id,
		case when a.type_code='01' then a.val*b.rate else a.val end as val  from tmp5 a join temp1 b on 1=1 where a.jz_code='Z002' and b.pro_name='氩气'
		and substr(a.item_id,7,3) not in ('118','119')
		union all
		select '制氮气' as jz_name,'Z023' as jz_code,a.type_name,a.type_code,'Z023'||substr(a.item_id,5) as item_id,
		case when a.type_code='01' then a.val*b.rate else a.val end as val  from tmp5 a join temp1 b on 1=1 where a.jz_code='Z002' and b.pro_name='氮气'
		and substr(a.item_id,7,3) not in ('118','061');
	quit;
	proc sql;
	create table temp2 as select sum(val) as val_all,jz_code
	from tmp6
	where type_code in ('04')
	group by jz_code;
	quit;

/*计算单耗*/

/*汇总产品*/

proc sql;
	create table tmp7 as select a.jz_name,a.jz_code,a.type_code,a.item_id,case when b.val_all in (0,.) then 0 else a.val/b.val_all end as dh_val
	from tmp6 a join temp2 b
	on a.jz_code=b.jz_code ;
quit;
/*01能源 02原材料 03辅助材料 04产品 05副产品  06大气排放 07水体排放 08固废*/
proc sql;
update tmp7 set dh_val=0-dh_val where type_code in ('04');
quit;

/*取能源工序*/
proc sql;
	create table tmp7 as select a.* from tmp7 a,wh_lca_0002 b
	where a.jz_code=b.jz_code ;
quit;
/*使用数据结束*/
/*水煤处理等处理,铁矿石等能源工序无 */
proc sql;
	create table temp3 as select a.*,b.pro_code,item_name from tmp7 a left join wh_lc.wh_lca_0007 b
	on substr(a.item_id,7,3)=b.item_code and a.type_code=b.type_code;
quit;
/****处理固体副产品（利用）*****/
/*产出副产品中为固体废弃物*/
data temp3_1;
	set temp3;
	where pro_code='SBProduct' and type_code='05';
run;
/*(需要优化逻辑)根据item_name中文来判断不同代码是否属于同一个数据项，jz_code不同代表不是同一个工序产出和投入*/
proc sql;
	create table temp3_2 as select a.* from temp3 a join temp3_1 b on 
	a.item_name=b.item_name and a.jz_code^=b.jz_code 
	and a.type_code in ('01','02') and a.pro_code='SBProduct';
quit;

/*data temp3;
	set tmp7;
	if substr(item_id,7,3) in ('004','014','031','017') then pro_code='Water';
	else if substr(item_id,7,3) in ('011','010','019','032') then pro_code='coal';
	else if substr(item_id,7,3) in ('071','048','079') then pro_code='ironstone';
	else if substr(item_id,7,3) in ('066','127','128','129','109',) then pro_code='limestone';
	else if substr(item_id,7,3) in ('029','130','131','132','133','010') then pro_code='dolomite';

	else if substr(item_id,7,3) in ('039') and type_code='02' then pro_code='Siron';
	else if substr(item_id,7,3) in ('037','030','031') and type_code='02'  then pro_code='Ssteel';
	else if substr(item_id,7,3) in ('014') and type_code='05' then pro_code='wsteel';
run;*/
proc sql;
	create table temp4 as
	select sum(dh_val) as dh_val ,pro_code,jz_code,jz_name 
	from temp3
	where pro_code not in ('','.') and pro_code ^='SBProduct'
	group by pro_code,jz_code,jz_name
   union all
   select sum(dh_val) as dh_val ,pro_code,jz_code,jz_name 
	from temp3_2
	where pro_code not in ('','.') and pro_code ='SBProduct'
	group by pro_code,jz_code,jz_name
;
quit;

/*分析使用*/
proc sql;
	create table temp3_fc as select jz_name,jz_code,dh_val as af,1 as bf,item_name,substr(item_id,7,3) as pro_code1,pro_code from temp3 
where pro_code not in ('','.') and pro_code ^='SBProduct'
union all 
select jz_name,jz_code,dh_val as af,1 as bf,item_name,substr(item_id,7,3) as pro_code1,pro_code from temp3_2;

quit;

/*能源处理*/
proc sql;
create table lca0001 as select avg(co2_ratio) as co2_ratio,pro_code,type_code
from wh_lc.wh_lca_0001
where ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6) and type_code in ('01','02')
group by pro_code,type_code;
quit;
proc sql;
create table lca0009 as select avg(hot_value) as hot_value,pro_code,type_code
from wh_lc.wh_lca_0009
where ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6) and type_code in ('01','02')
group by pro_code,type_code;
quit;
/*保留中间过程,矩阵分析使用*/
proc sql;
create table temp5_fc as select a.jz_name,a.jz_code,a.dh_val as af,b.co2_ratio*1000 as bf,b.pro_code as pro_code1,b.type_code,'CO2' as pro_code
from tmp7 a join lca0001 b
on substr(a.item_id,7,3)=b.pro_code and substr(a.item_id,5,2)=b.type_code
union 
 select a.jz_name,a.jz_code,a.dh_val as af,b.hot_value*1000 as bf,b.pro_code as pro_code1,b.type_code,'energy' as pro_code
from tmp7 a join lca0009 b
on substr(a.item_id,7,3)=b.pro_code and substr(a.item_id,5,2)=b.type_code;
create table temp5_fc1 as select a.*,b.pro_code_name as item_name from temp5_fc a left join ma_lc.wh_lcdm0009 b
on a.type_code=b.type_code and a.pro_code1=b.pro_code;
quit;
data g1_fc;
set temp5_fc1 temp3_fc;
LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	type="&type";
root="&root1";
shopsign_code="&shopsign_code1";
run;
proc sql;
create table g1_fc as select a.*,b.id from g1_fc a,wh_lc.wh_lca_0005 b where a.pro_code=b.pro_code;
quit;

proc sql;
create table temp5 as select a.jz_name,a.jz_code,a.dh_val as dh_vala,b.co2_ratio,a.dh_val*b.co2_ratio*1000 as dh_val,b.pro_code,b.type_code
from tmp7 a join lca0001 b
on substr(a.item_id,7,3)=b.pro_code and substr(item_id,5,2)=b.type_code;

create table temp51 as select a.jz_name,a.jz_code,a.dh_val as dh_vala,b.hot_value,a.dh_val*b.hot_value as dh_val,b.pro_code,b.type_code
from tmp7 a join lca0009 b
on substr(a.item_id,7,3)=b.pro_code and substr(item_id,5,2)=b.type_code;
quit;
/*水电能源*/
proc sql;
create table temp6 as 
 select a.jz_code,a.jz_name,'energy' as pro_code,sum(dh_val) as dh_val
from temp51 a
group by jz_code,jz_name
union all
 select a.jz_code,a.jz_name,'CO2' as pro_code,sum(dh_val) as dh_val
from temp5 a
group by jz_code,jz_name
union all 
 select a.jz_code,a.jz_name, pro_code,dh_val
from temp4 a
;
quit;
	/****处理发电工序****/
	/*CCPP发电_0#	燃煤发电	CCPP发电_4#三个工序产品电所占比例*/
	proc sql;
		create table temp21 as select * from temp2 where jz_code in ('D001','D002','D003');
		create table temp22  as select sum(val_all) as val_all from temp2 where jz_code in ('D001','D002','D003');
	quit;
	proc sql;
		create table temp23 as select a.val_all/b.val_all as bl,a.jz_code from temp21 a left join temp22 b on 1=1;
	quit;
	/*处理结束*/
/**CCPP发电所占比例*/
proc sql;
		create table temp21_a as select * from temp2 where jz_code in ('D001','D003');
		create table temp22_a as select sum(val_all) as val_all from temp2 where jz_code in ('D001','D003');
	quit;
	proc sql;
		create table temp23_a as select a.val_all/b.val_all as bl,a.jz_code from temp21_a a left join temp22_a b on 1=1;
	quit;


/*矩阵G1*/
proc sql;
create table lca0002 as select jz_code,jz_name,id from wh_lca_0002 order by id;
quit;

proc sql;
create table g1 as select a.jz_code,a.id,b.pro_code,b.dh_val
from lca0002 a left join temp6 b
on a.jz_code=b.jz_code
order by id;
quit;

/*环保报表数据*/

proc sql;
	create table lca0003 as select pro_code,avg(var) as var,jz_code
	from wh_lc.wh_lca_0003
	where  ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6) and jz_code in (select distinct jz_code from wh_lca_0002 )
	group by pro_code,jz_code;
	create table lca0003 as select a.*,b.unit from lca0003 a left join wh_lc.wh_lca_0005 b on a.pro_code=b.pro_code;
quit;
data lca0003(drop=unit);
	set lca0003;
	 if unit in ('mg') then var=var*1000*1000;
	else if unit in ('kg') then var=var;
	else if unit in ('g') then var=var*1000;
	else var=var;
run;
proc sql;
create table lca0003_1 as select a.pro_code,a.var,a.jz_code,b.bl from lca0003 a left join temp23_a b
on a.jz_code=b.jz_code;
quit;
data lca0003_1(drop=bl);
set lca0003_1;
if bl=. then var=var;
else var=var*bl;
run;
proc sql;	
	create table lca0003 as select a.*,a.var/b.val_all as dh_val from lca0003_1 a left join temp2 b
	on a.jz_code=b.jz_code;
quit;

proc sql;
	create table lca00031 as select a.*,b.id from lca0003 a left join wh_lca_0002 b 
	on a.jz_code=b.jz_code ;
quit;

/****采集的数据表用户未提供*****/

data g11;
set g1 lca00031;
run;


/*增加发电工序*/
proc sql;
	create table g11a as 
		select 'DDDD' as jz_code,18 as id,sum(a.dh_val*b.bl) as dh_val,a.pro_code from g11 a join temp23 b
		on a.jz_code=b.jz_code and a.jz_code in ('D001','D002','D003')
		group by pro_code
		union all
		select jz_code,id,dh_val,pro_code from g11;
quit;



proc sort data=g11a; by pro_code ;run;

PROC TRANSPOSE data=g11a out=g1a(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;


data result1;
retain pro_code col1-col18;
set g1a;
where pro_code^='';
run;


/*转换成矩阵前固定数据表格式*/
proc sql;
create table result2 as select a.id,a.pro_name,a.pro_code,b.* from wh_lc.wh_lca_0005 a left join result1 b
on a.pro_code=b.pro_code
order by id;
quit;

/*****************矩阵G1数据*********************/
data result_g1;
set result2;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;

data result_g1;
set result_g1;
drop id pro_name pro_code ;
run;
/********************矩阵G1：result_g1/data_g1********************************/
proc sql;
delete from ma_lc.lca_g1 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and type="&type" and root="&root1" and shopsign_code="&shopsign_code1";
quit;
proc append base=ma_lc.lca_g1 data=g1_fc force;run;

/*能源产品对自产能源的直接消耗系数矩阵*/
data lca0006;
	set wh_lc.wh_lca_0006;
	where type="&type" and flag='1';
run;
proc sql;
create table lca0006_1 as select distinct type_code,pro_code from wh_lc.wh_lca_0006
	where type="&type" and flag='1';
quit;

/*电需要特殊处理temp23:三个工序所占比例*/
/*固定表格式*/
proc sql;
	create table g2_tmp1 as select jz_code,id from wh_lca_0002 where  id<=17 order by id;
quit;
proc sql;
	create table g2_tmp2 as select a.jz_code,a.dh_val,substr(a.item_id,7,3) as pro_code,a.type_code
	from tmp7 a ,lca0006_1 b
	where a.type_code=b.type_code and   substr(a.item_id,7,3)=b.pro_code;
quit;

proc sql;
	create table g2_tmp3 as select a.jz_code,a.dh_val,b.jz_code as pro_code,a.pro_code as item
	from g2_tmp2 a  join lca0006 b on a.pro_code=b.pro_code and a.type_code=b.type_code;
quit;
proc sql;
	create table g2_tmp4 as select a.jz_code,a.id,b.dh_val,b.pro_code,b.item from g2_tmp1 a left join g2_tmp3 b
	on a.jz_code=b.jz_code;
quit;
proc sql;
	create table g2_tmp5 as select a.jz_code,a.id,a.item,b.bl,a.dh_val as val,case when a.item='009' then a.dh_val*b.bl  else a.dh_val end as dh_val,a.pro_code
	from g2_tmp4 a left join temp23 b
	on a.pro_code=b.jz_code
order by id;
quit;
/*转置*/

proc sort data=g2_tmp5; by pro_code ;run;

PROC TRANSPOSE data=g2_tmp5 out=g2_tmp6(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;
/*固定表格式*/

proc sql;
create table g2_tmp7 as select a.jz_code,a.id,b.col1,b.col2,b.col3,b.col4,b.col5,b.col6,b.col7,b.col8 ,b.col9,b.col10,b.col11,b.col12,b.col13,b.col14,b.col15,b.col16,b.col17
from g2_tmp1 a left join g2_tmp6 b
on a.jz_code=b.pro_code
order by id;
quit;
data g2_tmp8;
	retain  id jz_code col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17 ;
	set g2_tmp7;
run;

/*****************能源产品对自产能源的直接消耗系数矩阵PB2 :result_pb2
													G2:result_g2
*********************/
data result_pb2;
set g2_tmp8;
drop id jz_code;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;
/*处理G2*/
/*g1需要去掉发电工序*/

data result_g12;
	set result_g1;
	drop col18 id pro_code pro_name;

run;

proc iml;
	use result_g12;
	read all var _all_ into g1;
	use result_pb2;
	read all var _all_ into pb2;
	I=I(17);
	
	g2=g1*(inv(I-pb2)-I);
	aaa=inv(I-pb2)-I;
	ttt=inv(I-pb2);
	create result_ttt from ttt;
	append from ttt;
	create result_aaa from aaa;
	append from aaa;
	create result1_g2 from g2;
	append from g2;
quit;
data result_g21;
set result_g1;
id=_n_;
run;
data result_aaa;
	set result_aaa;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	id1=_N_;
run;
proc sort data=result_g21; by id ;run;

PROC TRANSPOSE data=result_g21 out=result_g211  prefix=col;
by id ;
run;
data result_g211;
	set result_g211;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g21 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.jz_name as item_name
from result_g211 a ,wh_lca_0002 b
where a.name=b.id  and b.id<18;
quit;
proc sort data=lca_g21;by id;run;

data lca_g21;
set lca_g21;
by id;
if first.id then id1=0;
id1+1;
run;
%let id_num=17;
%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g21_r as select a.*,b.col&i as bf,c.jz_code   from lca_g21 a ,result_aaa b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
%end;
%else %do;
proc sql;
create table lca_g21_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g21 a ,result_aaa b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc append base=lca_g21_r data=lca_g21_&i force;run;
%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;

data lca_g2(drop=id1);
set lca_g21_r ;
type="&type";
root="&root1";
shopsign_code="&shopsign_code1";
run;
proc sql;
delete from ma_lc.lca_g2 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root1" and shopsign_code="&shopsign_code1";
quit;
proc append base=ma_lc.lca_g2 data=lca_g2 force;run;

/*ttt ttt1为环境负荷分布1*/
data result_ttt;
set result_ttt;
id=_n_;
run;
proc sort data=result_ttt; by id ;run;

PROC TRANSPOSE data=result_ttt out=result_tttt  prefix=col;
by id ;
run;
data result_tttt;
	set result_tttt;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table result_tttt as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as bf,b.jz_code,"&type" as type,"&root1" as root,"&shopsign_code1" as shopsign_code
from result_tttt a,wh_lca_0002 b
where a.name=b.id ;
quit;
data result_ttt1;
set result_g12;
id=_n_;
run;

proc sort data=result_ttt1; by id ;run;

PROC TRANSPOSE data=result_ttt1 out=result_ttt2  prefix=col;
by id ;
run;
data result_ttt2;
	set result_ttt2;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;

proc sql;
create table result_ttt2 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.jz_name as item_name,"&type" as type,"&root1" as root,"&shopsign_code1" as shopsign_code
from result_ttt2 a,wh_lca_0002 b
where a.name=b.id ;
quit;

proc sql;
delete from ma_lc.lca_fb1a where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root1" and shopsign_code="&shopsign_code1";
quit;
proc append base=ma_lc.lca_fb1a data=result_ttt2 force;run;
proc sql;
delete from ma_lc.lca_fb1b where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root1" and shopsign_code="&shopsign_code1";
quit;
proc append base=ma_lc.lca_fb1b data=result_tttt force;run;


/*增加发电工序*/
data result2_g2;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
id=_n_;
set result1_g2;
run; 
proc sql;
	create table tt1 as select a.id,a.col1*b.bl as col18 from result2_g2 a join temp23 b
	on 1=1 and b.jz_code='D001';
	create table tt2 as select a.id,a.col2*b.bl as col18 from result2_g2 a join temp23 b
	on 1=1 and b.jz_code='D002';
	create table tt3 as select a.id,a.col3*b.bl as col18 from result2_g2 a join temp23 b
	on 1=1 and b.jz_code='D003';
	create table tt4 as select a.id,a.col18+b.col18+c.col18 as col18 from tt1 a,tt2 b,tt3 c
	where a.id=b.id and a.id=c.id and b.id=c.id;
	create table result_g2 as select a.*,b.col18 from result2_g2 a left join tt4 b on a.id=b.id;
quit;

data result_g2;
set result_g2;
drop id;
run;




/*******************************G2处理结束**************************************/
/*G3公式：*/
/*三种煤气：高炉煤气011 焦炉煤气019   转炉煤气032 17道工序*/
/****************************RB2:result_rb2
                             Gbcl:result_gbcl
							 G3:result_g3*************************************/
proc sql;
	create table tmp1_g3 as select jz_code,substr(item_id,7,3) as pro_code,sum(dh_val) as dh_val from tmp7 where substr(item_id,7,3) in ('011','032','019') and substr(item_id,5,2)='01'
	group by jz_code,substr(item_id,7,3);
quit;
proc sql;
	create table tmp2_g3 as select jz_code,dh_val,case when pro_code='011' then 2 when pro_code='019' then 1 when pro_code='032' then 3 end as pro_code from tmp1_g3;
quit;
proc sql;
	create table tmp3_g3 as select a.jz_code,a.id,b.pro_code,b.dh_val
	from lca0002 a left join tmp2_g3 b
	on a.jz_code=b.jz_code 
	where a.id<18
	order by id;
quit;
/*转置*/
proc sort data=tmp3_g3; by pro_code ;run;

PROC TRANSPOSE data=tmp3_g3 out=tmp4_g3(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;

data tmp4_g3;
retain pro_code col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;

set tmp4_g3;
where pro_code not in (.);
run; 

data result_rb2;
set tmp4_g3;
drop pro_code;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;
/*Gbcl*/

proc sql;
	create table lca0008 as select avg(jl) as jl,avg(gl) as gl,avg(zl) as zl,pro_code,id
	from wh_lc.wh_lca_0008
	where ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6)
	group by pro_code,id
	order by id;
quit;
data result_gbcl;
	set lca0008;
	retain  jl gl zl;
	drop pro_code id year;
run;
/*g3=gbcl(rb2(i-pb2)^)*/
proc iml;
	use result_gbcl;
	read all var _all_ into gbcl;
	use result_rb2;
	read all var _all_ into rb2;
	use result_pb2;
	read all var _all_ into pb2;	
	I=I(17);
	a=rb2*(inv(I-pb2));
	g3=gbcl*a;
create result_a from a;
	append from a;
	create result1_g3 from g3;
	append from g3;
quit;

/*增加发电工序*/
data result2_g3;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
id=_n_;
set result1_g3;
run; 
proc sql;
	create table tt5 as select a.id,a.col1*b.bl as col18 from result2_g3 a join temp23 b
	on 1=1 and b.jz_code='D001';
	create table tt6 as select a.id,a.col2*b.bl as col18 from result2_g3 a join temp23 b
	on 1=1 and b.jz_code='D002';
	create table tt7 as select a.id,a.col3*b.bl as col18 from result2_g3 a join temp23 b
	on 1=1 and b.jz_code='D003';
	create table tt8 as select a.id,a.col18+b.col18+c.col18 as col18 from tt5 a,tt6 b,tt7 c
	where a.id=b.id and a.id=c.id and b.id=c.id;
	create table result_g3 as select a.*,b.col18 from result2_g3 a left join tt8 b on a.id=b.id;
quit;

data result_g3;
set result_g3;
drop id;
run;
data result_g31;
set result_gbcl;
id=_n_;
run;
data result_a;
	set result_a;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	id1=_N_;
run;
proc sort data=result_g31; by id ;run;

PROC TRANSPOSE data=result_g31 out=result_g311  prefix=col;
by id ;
run;
data result_g311;
	set result_g311;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=_name_;

run;
proc sql;
create table lca_g31 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,case when a.NAME='jl' then '焦炉煤气' when name='zl' then '转炉煤气' when name='gl' then 
'高炉煤气' end as item_name
from result_g311 a ;
quit;

proc sort data=lca_g31;by id;run;
data lca_g31;
set lca_g31;
by id;
if first.id then id1=0;
id1+1;
run;
%let id_num=17;
%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g31_r as select a.*,b.col&i as bf,c.jz_code   from lca_g21 a ,result_a b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
%end;
%else %do;
proc sql;
create table lca_g31_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g21 a ,result_a b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc append base=lca_g31_r data=lca_g31_&i force;run;
%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;

data lca_g3(drop=id1);
set lca_g31_r ;
type="&type";
root="&root1";
shopsign_code="&shopsign_code1";
run;
proc sql;
delete from ma_lc.lca_g3 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root1" and shopsign_code="&shopsign_code1";
quit;
proc append base=ma_lc.lca_g3 data=lca_g3 force;run;



/***************************************矩阵G3处理结束*******************************************************/
/******************************************矩阵G4***********************************************************/

/*能源工序无G4，做模型统一*/


/******************************************矩阵G5***********************************************************/
						/* 矩阵T:result_t
 						   矩阵L:result_l 
						   矩阵OB:result_ob*/
/*矩阵T*/
proc sql;
create table wh_lca_0010 as select avg(train) as train,avg(barge) as barge,avg(seacraft) as seacraft,avg(car) as car,pro_code
from wh_lc.wh_lca_0010
where ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6)
group by pro_code;
quit;
proc sql;
	create table result_t as select a.id,a.pro_code,b.train,b.barge,b.seacraft,b.car
	from wh_lc.wh_lca_0005 a left join wh_lca_0010 b
	on a.pro_code=b.pro_code 
	order by a.id;
quit;

data result_t(drop=pro_code);
set result_t;
drop id;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;
/*矩阵L*/
proc sql;
	create table l_tmp1 as select pro_code,id,unit,avg(train) as train,avg(barge) as barge,avg(seacraft) as seacraft,avg(car) as car
	from wh_lc.wh_lca_0011
		where ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6)
		group by pro_code,id,unit;
quit;

/*转置*/
proc sort data=l_tmp1; by id;run;

PROC TRANSPOSE data=l_tmp1 out=l_tmp2  prefix=col;
	var   train barge seacraft car ;	
run;
data result_l(drop=_LABEL_ _name_ );
	set l_tmp2;
	where _name_ ne '';
run;
/*矩阵OB:工序中消耗量，不算输出量*/

proc sql;
	create table result_ob1 as select a.jz_code,a.dh_val,case when b.pro_code='coal' then compress(b.pro_code||b.item_code) else b.pro_code end as pro_code from tmp7 a join wh_lc.wh_lca_0007 b 
	on substr(a.item_id,7,3)=b.item_code and  a.type_code=b.type_code and a.type_code in ('01','02','03') and b.pro_code in ('ironstone','limestone','dolomite','coal');
quit;
proc sql;
	create table result_ob1 as select sum(dh_val) as dh_val,jz_code,pro_code from result_ob1 group by jz_code,pro_code;
quit;
proc sql;
create table result_ob2 as select a.dh_val,a.pro_code,b.id from result_ob1 a right join wh_lca_0002 b on a.jz_code=b.jz_code  ;
quit;
/*转置*/

proc sort data=result_ob2; by pro_code ;run;

PROC TRANSPOSE data=result_ob2 out=result_ob3(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;
proc sql;
create table wh_lca_0011 as select distinct id,pro_code from  wh_lc.wh_lca_0011 ;
quit;
proc sql;
	create table result_ob4 as select a.pro_code,b.id,a.col1,a.col2,a.col3,a.col4,a.col5,a.col6,a.col7,
	a.col8,a.col9,a.col10,a.col11,a.col12,a.col13,a.col14,a.col15,a.col16,a.col17
	from result_ob3 a right join wh_lca_0011 b
	on a.pro_code=b.pro_code 
order by b.id;
quit;
data result_ob5;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
set result_ob4;
run;

data result_ob(drop=id pro_code);
set result_ob5;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run; 

proc iml;
	use result_t;
	read all var _all_ into t;
	use result_l;
	read all var _all_ into l;
	use result_ob;
	read all var _all_ into ob;
	use result_pb2;
	read all var _all_ into pb2;
	I=I(17);
	ob7=(t*l)*(ob*inv(I-pb2));
	a5=(t*l);
	b5=(ob*inv(I-pb2));
	create result_a5 from a5;
	append from a5;
	create result_b5 from b5;
	append from b5;
	create result_ob7 from ob7;
	append from ob7;
quit;
/*增加发电工序*/
data result_ob8;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
id=_n_;
set result_ob7;
run; 
proc sql;
	create table tt1 as select a.id,a.col1*b.bl as col18 from result_ob8 a join temp23 b
	on 1=1 and b.jz_code='D001';
	create table tt2 as select a.id,a.col2*b.bl as col18 from result_ob8 a join temp23 b
	on 1=1 and b.jz_code='D002';
	create table tt3 as select a.id,a.col3*b.bl as col18 from result_ob8 a join temp23 b
	on 1=1 and b.jz_code='D003';
	create table tt4 as select a.id,a.col18+b.col18+c.col18 as col18 from tt1 a,tt2 b,tt3 c
	where a.id=b.id and a.id=c.id and b.id=c.id;
	create table result_g5 as select a.*,b.col18 from result_ob8 a left join tt4 b on a.id=b.id;
quit;

data result_g5;
set result_g5;
drop id;
run;

data result_a5;
set result_a5;
id=_n_;
run;
proc sort data=result_a5; by id ;run;

PROC TRANSPOSE data=result_a5 out=result_a51  prefix=col;
by id ;
run;
data result_a51;
	set result_a51;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g51 as select distinct  a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.pro_name as item_name
from result_a51 a,wh_lc.wh_lca_0011 b
where a.name=b.id ;
quit;
proc sort data=lca_g51;by id;run;

data lca_g51;
set lca_g51;
by id;
if first.id then id1=0;
id1+1;
run;
data result_b5;
	set result_b5;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
id1=_N_;
run;



%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g51_r as select a.*,b.col&i as bf,c.jz_code   from lca_g51 a ,result_b5 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;

%end;
%else %do;
proc sql;
create table lca_g51_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g51 a ,result_b5 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;

proc append base=lca_g51_r data=lca_g51_&i force;run;

%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;

data lca_g5(drop=id1);
set lca_g51_r ;
type="&type";
root="&root1";
shopsign_code="&shopsign_code1";
run;

proc sql;
delete from ma_lc.lca_g5 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root1" and shopsign_code="&shopsign_code1";
quit;
proc append base=ma_lc.lca_g5 data=lca_g5 force;run;


/***************************************矩阵g6*************************************************/
/********下载数据里面同一个数据项只有输入，无输出，则认为为外购；暂时使用中文，后续优化********/

proc sql;
	create table ob1_tmp1 as select a.*,b.pro_code_name,b.pro_code from tmp7 a left join ma_lc.Wh_lcdm0009 b 
	on a.type_code=b.type_code and substr(a.item_id,7,3)=b.pro_code;
	create table ob1_tmp2 as select * from ob1_tmp1 where type_code in ('01','02','03') and 
	pro_code_name not in (select distinct pro_code_name from ob1_tmp1 where type_code in ('04','05'));
quit;
proc sql;
	create table ob1_tmp2 as select * from ob1_tmp2 where pro_code_name in (select distinct product from wh_lc.wh_lca_0019)
and pro_code_name not like '%煤气' and pro_code_name not like '%氮气' and pro_code_name not like '%氧气';
quit;
proc sql;
	create table ob1_tmp2 as select a.*,b.id from ob1_tmp2 a left join wh_lca_0002 b 
	on a.jz_code=b.jz_code ;
quit;

proc sql;
	create table ob1_tmp3 as select a.jz_code,a.id,b.pro_code,b.dh_val
	from lca0002 a left join ob1_tmp2 b
	on a.jz_code=b.jz_code 
	where a.id<18
	order by id;
quit;
data ob1_tmp3;
set ob1_tmp3;
if pro_code='' then pro_code='AAA';
run;
proc sort data=ob1_tmp3; by pro_code id;run;

PROC TRANSPOSE data=ob1_tmp3 out=ob1_tmp4(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;
proc sort data=ob1_tmp4 ;by pro_code;run;
data result_ob1;
	retain  col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;

	
	set ob1_tmp4;
	where pro_code not in ('','AAA');
run; 

data result_ob1(drop=id pro_code);
set result_ob1;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run; 
/*******************************矩阵OG:维护表*********************************************/
proc sql;
	create table lca0012 as select a.*,b.pro_code_name  from wh_lc.wh_lca_0014 a,wh_lc.wh_lca_0019 b
	where  a.product=b.product and b.pro_code_name in (select distinct pro_code_name from ob1_tmp2 ) and pro_code ne '';
quit;
proc sql;
create table lca00121 as select distinct a.*,b.pro_code as pro_code_new from lca0012 a,ob1_tmp2 b
where a.pro_code_name=b.pro_code_name;
quit;
proc sql;
create table lca00121 as select distinct a.*,b.id as id_new from lca00121 a right join wh_lc.wh_lca_0005 b
on a.pro_code=b.pro_code;
quit;
/*转置*/
proc sql;
create table lca00121 as select distinct * from lca00121;
quit;
proc sort  data=lca00121; by id_new pro_code_new ;run;


PROC TRANSPOSE data=lca00121 out=result_og(drop=_name_ _label_ id_new  pro_code)   prefix=col;
by id_new;
var  SUM;
id  pro_code_new;
run;
data result_og;
set result_og;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run; 
data result_og;
set result_og;
retain   col1 - col600;
run;
/*矩阵G6=OG*[OB1(I-PB2)-1  ]*/

proc iml;
	use result_ob1;
	read all var _all_ into ob1;
	use result_og;
	read all var _all_ into og;
	use result_pb2;
	read all var _all_ into pb2;
	I=I(17);
	g6=og*(ob1*inv(I-pb2));
	b6=(ob1*inv(I-pb2));
	create result_b6 from b6;
	append from b6;
	create result_g6 from g6;
	append from g6;
quit;

/*增加发电工序*/
data result_g6;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
id=_n_;
set result_g6;
run; 
proc sql;
	create table tt1 as select a.id,a.col1*b.bl as col18 from result_g6 a join temp23 b
	on 1=1 and b.jz_code='D001';
	create table tt2 as select a.id,a.col2*b.bl as col18 from result_g6 a join temp23 b
	on 1=1 and b.jz_code='D002';
	create table tt3 as select a.id,a.col3*b.bl as col18 from result_g6 a join temp23 b
	on 1=1 and b.jz_code='D003';
	create table tt4 as select a.id,a.col18+b.col18+c.col18 as col18 from tt1 a,tt2 b,tt3 c
	where a.id=b.id and a.id=c.id and b.id=c.id;
	create table result_g6 as select a.*,b.col18 from result_g6 a left join tt4 b on a.id=b.id;
quit;

data result_g6;
set result_g6;
drop id;
run;



data result_og;
set result_og;
id=_n_;
run;
proc sort data=result_og; by id ;run;

PROC TRANSPOSE data=result_og out=result_og1  prefix=col;
by id ;
run;
data result_og1;
	set result_og1;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g61 as select distinct  a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.product as item_name
from result_og1 a,wh_lc.WH_LCA_0014 b
where a.name=b.id ;
quit;
proc sort data=lca_g61;by id;run;
data lca_g61;
set lca_g61;
by id;
if first.id then id1=0;
id1+1;
run;


data lca_g62;
	set result_b6;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	id1=_N_;
run;


%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g61_r as select a.*,b.col&i as bf,c.jz_code   from lca_g61 a ,lca_g62 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;

%end;
%else %do;
proc sql;
create table lca_g61_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g61 a ,lca_g62 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;

proc append base=lca_g61_r data=lca_g61_&i force;run;

%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;

data lca_g6(drop=id1);
set lca_g61_r ;
type="&type";
root="&root1";
shopsign_code="&shopsign_code1";
run;

proc sql;
delete from ma_lc.lca_g6 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root1" and shopsign_code="&shopsign_code1";
quit;
proc append base=ma_lc.lca_g6 data=lca_g6 force;run;
/****************************************矩阵G7（全部乘以-1：外用副产品模型：工序产出的副产品没有在厂内使用，没有当做其他工序的输入
											暂时使用中文来区分，后续优化
******************************************************/
proc sql;
	create table rb_tmp1 as select a.*,b.pro_code_name,b.pro_code from tmp7 a left join ma_lc.Wh_lcdm0009 b 
	on a.type_code=b.type_code and substr(a.item_id,7,3)=b.pro_code;
	
	create table rb_tmp2 as select * from rb_tmp1 where type_code in ('05') and 
	pro_code_name not in (select distinct pro_code_name from rb_tmp1 where type_code in ('01','02','03'));
quit;
proc sql;
	create table rb_tmp2 as select * from rb_tmp2 where pro_code_name in (select distinct pro_code_name from wh_lc.wh_lca_0019);
quit;
proc sql;
	create table rb_tmp2 as select a.*,b.id from rb_tmp2 a left join wh_lca_0002 b
	on a.jz_code=b.jz_code ;
quit;
proc sql noprint;
select distinct count(*) into:cnt3 from rb_tmp2;
quit;
%put cnt3===&cnt3;

%if &cnt3=0 %then %do;


proc iml;
	g7=J(53,18,0);
	create result_g7 from g7;
	append from g7;
quit;

%end;
%else %do;

proc sql;
	create table rb_tmp3 as select a.jz_code,a.id,b.pro_code,b.dh_val
	from lca0002 a left join rb_tmp2 b
	on a.jz_code=b.jz_code 
	where a.id<18
	order by id;
quit;
data rb_tmp3;
set rb_tmp3;
if pro_code='' then pro_code='AAA';
run;
proc sort data=rb_tmp3; by pro_code id;run;


PROC TRANSPOSE data=rb_tmp3 out=rb_tmp4(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;
proc sort data=rb_tmp4 ;by pro_code;run;
data result_rb;
	retain  col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;

	
	set rb_tmp4;
	where pro_code not in ('','AAA');
run; 

data result_rb(drop=id pro_code);
set result_rb;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run; 

/*******************************矩阵rg:维护表*********************************************/
/*proc sql;
	create table lca0013 as select item_code,item_name,pro_code,pro_name,var from wh_lc.wh_lca_0013 
	where year=substr(substr("&datefrom",1,6),1,4) and item_name in (select distinct pro_code_name from rb_tmp2 );
quit;
*/
proc sql;
	create table lca0013 as select a.*,b.pro_code_name  from wh_lc.wh_lca_0014 a,wh_lc.wh_lca_0019 b
	where  a.product=b.product and b.pro_code_name in (select distinct pro_code_name from rb_tmp2 ) and pro_code ne '';
quit;
proc sql;
create table lca00131 as select distinct a.*,b.pro_code as pro_code_new from lca0013 a,rb_tmp2 b
where a.pro_code_name=b.pro_code_name;
quit;
proc sql;
create table lca00131 as select distinct a.*,b.id as id_new,b.pro_code from lca00131 a right join wh_lc.wh_lca_0005 b
on  a.pro_code=b.pro_code;
quit;
proc sql;
create table lca00131 as select distinct * from lca00131;
quit;
/*转置*/
proc sort data=lca00131; by id_new pro_code_new ;run;

PROC TRANSPOSE data=lca00131 out=result_rg(drop=_name_ _label_ id_new  pro_code)   prefix=col;
by id_new;
var  SUM;
id  pro_code_new;
run;
data result_rg;
set result_rg;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run; 
data result_rg;
set result_rg;
retain   col1 - col600;
run;

proc iml;
	use result_rb;
	read all var _all_ into rb;
	use result_rg;
	read all var _all_ into rg;
	use result_pb2;
	read all var _all_ into pb2;
	I=I(17);
	g7=rg*(rb*inv(I-pb2));
	b7=(rb*inv(I-pb2));
	create result_b7 from b7;
	append from b7;
	create result_g7 from g7;
	append from g7;
quit;

/*增加发电工序*/
data result_g7;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
id=_n_;
set result_g7;
run; 
proc sql;
	create table tt1 as select a.id,a.col1*b.bl as col18 from result_g7 a join temp23 b
	on 1=1 and b.jz_code='D001';
	create table tt2 as select a.id,a.col2*b.bl as col18 from result_g7 a join temp23 b
	on 1=1 and b.jz_code='D002';
	create table tt3 as select a.id,a.col3*b.bl as col18 from result_g7 a join temp23 b
	on 1=1 and b.jz_code='D003';
	create table tt4 as select a.id,a.col18+b.col18+c.col18 as col18 from tt1 a,tt2 b,tt3 c
	where a.id=b.id and a.id=c.id and b.id=c.id;
	create table result_g7 as select a.*,b.col18 from result_g7 a left join tt4 b on a.id=b.id;
quit;

/*
	能源：53个数据项、18个工序	
*/

data result_g7;
set result_g7;
drop id;
run;

data result_rg;
set result_rg;
id=_n_;
run;
proc sort data=result_rg; by id ;run;

PROC TRANSPOSE data=result_rg out=result_rg12 prefix=col;
by id ;
run;
data result_rg12;
	set result_rg12;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g71 as select distinct a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.product as item_name
from result_rg12 a,wh_lc.wh_lca_0014 b
where a.name=b.id ;
quit;
proc sort data=lca_g71;by id;run;
data lca_g71;
set lca_g71;
by id;
if first.id then id1=0;
id1+1;
run;
data lca_g72;
	set result_b7;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	id1=_N_;
run;

%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g71_r as select a.*,b.col&i as bf,c.jz_code   from lca_g71 a ,lca_g72 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;

%end;
%else %do;
proc sql;
create table lca_g71_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g71 a ,lca_g72 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;

proc append base=lca_g71_r data=lca_g71_&i force;run;

%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;

data lca_g7(drop=id1);
set lca_g71_r ;
type="&type";
root="&root1";
shopsign_code="&shopsign_code1";
run;

proc sql;
delete from ma_lc.lca_g7 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root1" and shopsign_code="&shopsign_code1";
quit;
proc append base=ma_lc.lca_g7 data=lca_g7 force;run;

%end;
/****************************Gsite=G1+g2+g3+g4*****************************************/
/*************************************Groute:G1+G2+G3+G4+G5+G6+G7******************/

proc iml;
	use result_g1;
	read all var _all_ into g1;
	use result_g2;
	read all var _all_ into g2;
	use result_g3;
	read all var _all_ into g3;
	use result_g5;
	read all var _all_ into g5;
	use result_g6;
	read all var _all_ into g6;
	use result_g7;
	read all var _all_ into g7;
	groute=g1+g2+g3+g5+g6+g7;
	
	create result_groute from groute;
	append from groute;
quit;
/*G1+G2+G3+G4:内部合计*/
proc iml;
	use result_g1;
	read all var _all_ into g1;
	use result_g2;
	read all var _all_ into g2;
	use result_g3;
	read all var _all_ into g3;
	
	gsite=g1+g2+g3;
	
	create result_gsite from gsite;
	append from gsite;
quit;
/*G5+G6+G7:外部合计*/

proc iml;
	use result_g5;
	read all var _all_ into g5;
	use result_g6;
	read all var _all_ into g6;
	use result_g7;
	read all var _all_ into g7;
	
	goutsite=g5+g6+g7;
	
	create result_goutsite from goutsite;
	append from goutsite;
quit;
data result_goutsite;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_goutsite;
run; 
data result_gsite;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_gsite;
run; 
/*存为临时表：供主工序使用*/
data ma_lc.sa_lca_ny_gsite;
set result_gsite;
run;
data ma_lc.sa_lca_ny_gsite;
set result_gsite;
run;
data result_groute;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_groute;
run; 
data result_g1;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_g1;
run; 
data result_g2;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_g2;
run; 
data result_g3;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_g3;
run; 
data result_g5;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_g5;
run; 
/*存为临时表：供主工序使用*/
data ma_lc.sa_lca_ny_g5;
set result_g5;
run;
data ma_lc.sa_lca_ny_g5;
set result_g5;
run;
data result_g6;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_g6;
run; 
/*存为临时表：供主工序使用*/
data ma_lc.sa_lca_ny_g6;
set result_g6;
run;
data ma_lc.sa_lca_ny_g6;
set result_g6;
run;
data result_g7;
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  col18;
id=_n_;
set result_g7;
run; 
/*存为临时表：供主工序使用*/
data ma_lc.sa_lca_ny_g7;
set result_g7;
run;
data ma_lc.sa_lca_ny_g7;
set result_g7;
run;
/*转置*/
proc sort data=result_groute ;by id;run;
PROC TRANSPOSE data=result_groute out=result_groute2(rename=(_NAME_=jz_code col1=groute)) prefix=col;
by id;
run;
proc sort data=result_gsite ;by id;run;
PROC TRANSPOSE data=result_gsite out=result_gsite2(rename=(_NAME_=jz_code col1=gsite)) prefix=col;
by id;
run;
proc sort data=result_goutsite ;by id;run;
PROC TRANSPOSE data=result_goutsite out=result_goutsite2(rename=(_NAME_=jz_code col1=goutsite)) prefix=col;
by id;
run;
proc sort data=result_g1 ;by id;run;
PROC TRANSPOSE data=result_g1 out=result_g12(rename=(_NAME_=jz_code col1=g1)) prefix=col;
by id;
run;
proc sort data=result_g2 ;by id;run;
PROC TRANSPOSE data=result_g2 out=result_g22(rename=(_NAME_=jz_code col1=g2)) prefix=col;
by id;
run;
proc sort data=result_g3 ;by id;run;
PROC TRANSPOSE data=result_g3 out=result_g32(rename=(_NAME_=jz_code col1=g3)) prefix=col;
by id;
run;
proc sort data=result_g5 ;by id;run;
PROC TRANSPOSE data=result_g5 out=result_g52(rename=(_NAME_=jz_code col1=g5)) prefix=col;
by id;
run;
proc sort data=result_g6 ;by id;run;
PROC TRANSPOSE data=result_g6 out=result_g62(rename=(_NAME_=jz_code col1=g6)) prefix=col;
by id;
run;
proc sort data=result_g7 ;by id;run;
PROC TRANSPOSE data=result_g7 out=result_g72(rename=(_NAME_=jz_code col1=g7)) prefix=col;
by id;
run;
proc sort data=result_groute2;by  id jz_code;run;
proc sort data=result_gsite2;by  id jz_code;run;
proc sort data=result_goutsite2;by  id jz_code;run;
proc sort data=result_g12;by  id jz_code;run;

proc sort data=result_g22;by  id jz_code;run;

proc sort data=result_g32;by  id jz_code;run;

proc sort data=result_g52;by  id jz_code;run;
proc sort data=result_g62;by  id jz_code;run;
proc sort data=result_g72;by  id jz_code;run;

data result_all;
merge result_groute2 result_gsite2 result_goutsite2 result_g12 result_g22 result_g32 result_g52 result_g62 result_g72;
by id jz_code;
run;


/*勾连数据项*/
proc sql;
create table result_alla as select b.pro_code,b.pro_name as LCIfactor,b.unit,a.* from result_all a left join wh_lc.wh_lca_0005 b
on a.id=b.id;
quit;

proc sql;
create table result_alla as select substr("&datefrom",1,6) as LCItime_s,substr("&dateto",1,6) as LCItime_e,b.jz_code as jz_code1,b.jz_name as product,a.* from result_alla a left join wh_lca_0002 b
on input(substr(a.jz_code,4),2.)=b.id ;
quit;

data result_alla;
set result_alla;
drop id jz_code;
rename jz_code1=jz_code;
rec_create_time=put(date(),yymmddn8.);
type="&type";
run;
proc sort data=result_alla ;by jz_code LCIfactor;run;

proc sql;
delete from ma_lc.sa_lca_0001 where LCITIME_S=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and type="&type";
quit;
proc append base=ma_lc.sa_lca_0001 data=result_alla force;run;

/*******************************************能源工序结束*******************************************************/

%mend;
%edition;



