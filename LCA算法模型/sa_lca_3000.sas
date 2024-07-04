
/*******************************************主工序开始*********************************************************/
/***主工序数量：135 数据项个数：53**/

%lib(ma_lc);
%lib(wh_lc);


%let datefrom=&date_f;
%let dateto=&date_e;
%let type=gx;
%let shopsign_code=&shopsign_code;
%let root=&root;


%macro dl;
%if "&root"="DL" %then %do;


		data sa_lc_0006;

			set ma_lc.sa_lc_0006 ;
			where compress(year||mon) between substr("&datefrom",1,6) and substr("&dateto",1,6) ;
			unit='';
		run;
%end;
%else %do;
%put root=&root;
		data sa_lc_00061;
			set tmp/*ztl.tmp1*/;	
			where unit not in ('','元') and val not in (.,0) and item_id not in ('') ;
		run;
		data tt;
		set sa_lc_00061;
		where item_id='25RH04025';
		item_id='25RH02041';
		type_name='原材料';
		pro_code_n='钢水';
		run;
		data sa_lc_00061;
		set sa_lc_00061 tt;
		run;
		data sa_lc_00061;
		set sa_lc_00061;
		if substr(item_id,5,5) in ('02194','02T04','02T06','02T08','02T09','02T16') then delete;/*装入量\主原料重废\主原料中混废\主原料渣钢\主原料高纯渣铁\电工钢打包块 删除*/
		if substr(item_id,5,5) in ('02195','02196')  then do;
		item_id=substr(item_id,1,4)||'02073';/*熔铁量\冷铁量合并为铁水*/
		pro_code_n='铁水';end;
		run;

		

		/*proc sql;
		update sa_lc_00061 set val=3266948 where item_id='2BOF05053';
		quit;*/
		data wh_lca_0018;
		set wh_lc.wh_lca_0018;
		run;
		proc sql;
				create table sa_lc_00061 as select a.*,b.jz_code,b.jz_name,b.jz_code_dl,b.process_code,b.machine_set_code from sa_lc_00061 a
				left join wh_lca_0018 b on substr(a.item_id,1,4)=b.jz_code ;
				create table sa_lc_00061_sum as select sum(val) as val,jz_code,jz_name,jz_code_dl,process_code,machine_set_code from sa_lc_00061
				where substr(item_id,5,2)='04'
				group by jz_code,jz_name,jz_code_dl,process_code,machine_set_code;
		quit;
		/*增加前道公用工序*/
		data sa_lc_0006;
			set ma_lc.sa_lc_0006 ;
			where compress(year||mon) between substr("&datefrom",1,6) and substr("&dateto",1,6) and
		substr(item_id,1,4) in ('BF01','BF02','BF03','BF04','BS01','BS02','BS03','CO01','CO02','CO03','CO04','CO05','GL00','LJ00','SI01','SI02','SI03','SI04','D001','D002','D003');
		run;
		/*大类产品，用于典型牌号分摊

		proc sql;
		create table tt as select * from  ma_lc.sa_lc_0006 where compress(year||mon) between substr(substr("&datefrom",1,6),1,6) and substr(substr("&dateto",1,6),1,6) 
		and substr(item_id,1,4) in (select distinct jz_code_dl from sa_lc_00061_sum);
		create table tt2 as select distinct sum(val) as val,substr(item_id,1,4) as jz_code   from tt where substr(item_id,5,2)='04' group by substr(item_id,1,4);
		create table tt3 as select case when b.val in (.,0) then 0 when a.val in (.) then 0 else a.val/b.val end as bl,a.jz_code,a.jz_code_dl from sa_lc_00061_sum a left join tt2 b on a.jz_code_dl=b.jz_code;
		quit;*/
		data sa_lc_0006;
		set sa_lc_0006 sa_lc_00061(keep=item_id  val unit);
		run;
%end;
%mend;
%dl;
/*2转炉高压氧气和低压氧气数据重复，需删除，待处理*/

proc sql;
	create table tmp1 as select sum(val) as val,item_id,unit
	from sa_lc_0006  where item_id not in ('2YLT01013' )
	group by item_id,unit;
quit;

data test(keep=item_id);
set ma_lc.sa_lc_0006 ;
where compress(year||mon) between substr("&datefrom",1,6) and substr("&dateto",1,6);
run;

proc sql;
create table tmp11a_all  as select distinct b.pro_code_name,substr(a.item_id,5,2) as type_code,substr(a.item_id,7,3) as pro_code
from tmp1 a left join ma_lc.wh_lcdm0009 b on substr(a.item_id,5,2)=b.type_code and 
	substr(a.item_id,7,3)=b.pro_code
	union 
	select distinct b.pro_code_name,substr(a.item_id,5,2) as type_code,substr(a.item_id,7,3) as pro_code
from test a left join ma_lc.wh_lcdm0009 b on substr(a.item_id,5,2)=b.type_code and 
	substr(a.item_id,7,3)=b.pro_code;
create table tmp11_all  as select distinct pro_code_name,type_code,pro_code from tmp11a_all;
quit;
/*勾连工序代码和类型*/

proc sql;
	create table tmp2 as
	select a.val,a.item_id,a.unit,b.type_name,b.type_code
	from tmp1 a left join ma_lc.Wh_lcdm0008 b
	on substr(a.item_id,5,2)=b.type_code;
quit;
/*勾连典型牌号对应大类数据项代码*/

proc sql;
create table tmp21 as select a.*,b.pro_code_dl from tmp2 a left join wh_lc.wh_lca_0021 b
on a.type_code='02' and substr(a.item_id,1,4)=b.jz_code and substr(a.item_id,7,3)=b.pro_code;
create table tmp2 as  select a.val,a.unit,a.type_code,a.type_name,case when pro_code_dl^='' then substr(item_id,1,6)||pro_code_dl else item_id end as item_id
from tmp21 a;
quit;

data Wh_lcdm0003;
set ma_lc.Wh_lcdm0003;
 where process_code in ('LJ','SJ','GL','GF') and jz_code ne 'CO04' ;
run;
/*获取单位*/
proc sql;
	create table Wh_lcdm0006 as
	 select distinct jz_code,type_code,pro_code,unit
	 from  ma_lc.Wh_lcdm0006  
	where year=substr(substr("&datefrom",1,6),1,4)
	union 
	select distinct jz_code,type_code,pro_code,unit
	 from  ma_lc.Wh_lcdm0005  ;
quit;

proc sql;
	create table tmp2 as select a.*,b.unit as unit_new from tmp2 a left join Wh_lcdm0006 b
	on a.item_id=compress(b.jz_code||b.type_code||b.pro_code);
	update tmp2 set unit=unit_new where unit='';
quit;
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
create table Wh_lcdm0003 as select distinct jz_code,jz_name from ma_lc.Wh_lcdm0003 where jz_code not in (select jz_code from ma_lc.Wh_lcdm0004)
union 
select distinct jz_code,jz_name from ma_lc.Wh_lcdm0004;
quit;
proc sql;
		create table tmp2 as select a.*,b.jz_name,b.jz_code from tmp2 a
		left join Wh_lcdm0003 b on substr(a.item_id,1,4)=b.jz_code ;
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
/*CCPP发电_0#	燃煤发电	CCPP发电_4#三个工序产品电所占比例*/
	proc sql;
		create table temp21 as select * from tmp2 where substr(item_id,1,4) in ('D001','D002','D003') and substr(item_id,5,2)='04';
		create table temp22  as select sum(val) as val_all from temp21 where jz_code in ('D001','D002','D003');
	quit;
	proc sql;
		create table temp23 as select a.val/b.val_all as bl,a.jz_code from temp21 a left join temp22 b on 1=1;
	quit;

/*取主工序
proc sql;
	create table tmp4 as select a.* from tmp2 a,wh_lc.wh_lca_0002 b
	where a.jz_code=b.jz_code and b.type="&type";
quit;*/
/*工业水、纯水、过滤水单位m3,改为t*/
data tmp4;
set tmp2;
if substr(item_id,5,5) in ('01004','01014','01017') and unit='m3' then unit='t';
run;
%macro gx;
%if "&root"="DL"  %then %do;
proc sql;
	create table wh_lca_0002 as select distinct jz_code,jz_code as jz_code_dl,jz_name,id from wh_lc.wh_lca_0002 where type="&type";
quit;
%end;
%else %do;
proc sql;
	create table wh_lca_0002 as select distinct a.jz_code,a.jz_name,b.jz_code_dl from tmp2 a left join sa_lc_00061_sum b on a.jz_code=b.jz_code;
quit;
/*工序排序使用*/
data wh_lca_0002;
set wh_lca_0002;
where jz_code not in ('D001','D002','D003');
if jz_code_dl='' then jz_code_dl=jz_code;
id=_N_;
run;
%end;

%mend;
%gx;
proc sql noprint;
	select max(id) into: id_num from wh_lca_0002;
quit;
%put &id_num;
/*统一单位:kg  m3   kWh */
data tmp5(drop=unit unit_new);
	set tmp4;
	if unit in ('t') then val=val*1000;
	else if unit in ('m3') then val=val;
	else if unit in ('kwh') then val=val;
	else if unit in ('km3') then val=val*1000;
	else if unit in ('kg') then val=val;
	else val=val;
run;
/*合并产品*/
/*炼焦、铁水生铁、hfw、UOE、各工序次品材*/
proc sql;
update tmp5 set item_id=substr(item_id,1,4)||'04047' where substr(item_id,5,5)='04053';/*铁水*/
update tmp5 set item_id=substr(item_id,1,4)||'04001' where substr(item_id,5,5)='04011';/*粗焦*/
update tmp5 set item_id=substr(item_id,1,4)||'04001' where substr(item_id,5,5)='04032';/*集尘焦*/
update tmp5 set item_id=substr(item_id,1,4)||'04001' where substr(item_id,5,5)='04033';/*焦炭*/
update tmp5 set item_id=substr(item_id,1,4)||'04001' where substr(item_id,5,5)='04054';/*头尾焦*/
update tmp5 set item_id=substr(item_id,1,4)||'04085' where substr(item_id,5,5)='04086';/*结构管*/
update tmp5 set item_id=substr(item_id,1,4)||'04085' where substr(item_id,5,5)='04087';/*套管*/
update tmp5 set item_id=substr(item_id,1,4)||'04085' where substr(item_id,5,5)='04089';/*UOE钢管*/
update tmp5 set item_id=substr(item_id,1,4)||'04085' where substr(item_id,5,5)='04091';/*HFW成品*/
update tmp5 set item_id=substr(item_id,1,4)||'04085' where substr(item_id,5,5)='04117';/*次品才*/
update tmp5 set item_id=substr(item_id,1,6)||'193' where substr(item_id,5,5)='02147';
update tmp5 set item_id=substr(item_id,1,6)||'022' where substr(item_id,5,5)='02018';
update tmp5 set item_id=substr(item_id,1,6)||'175' where substr(item_id,5,5)='02171';
update tmp5 set item_id=substr(item_id,1,6)||'192' where substr(item_id,5,5)='02157';
update tmp5 set item_id=substr(item_id,1,6)||'183' where substr(item_id,5,5)='03019';
update tmp5 set item_id=substr(item_id,1,6)||'025' where substr(item_id,5,5)='01007';
update tmp5 set item_id=substr(item_id,1,6)||'025' where substr(item_id,5,5)='01013';
update tmp5 set item_id=substr(item_id,1,6)||'034' where substr(item_id,5,5)='01006';
update tmp5 set item_id=substr(item_id,1,6)||'044' where substr(item_id,5,5)='01020';/*CDQ粉和焦炭合并，每个工序只能有一个主产品，否则后面算自产能源时候出错*/
update tmp5 set item_id=substr(item_id,1,6)||'044' where substr(item_id,5,5)='01043';/*头尾焦*/
update tmp5 set item_id=substr(item_id,1,4)||'01044' where substr(item_id,5,5)='02120';/*集尘焦*/
update tmp5 set item_id=substr(item_id,1,6)||'044' where substr(item_id,5,5)='01005';/*粗焦*/
update tmp5 set item_id=substr(item_id,1,4)||'02022' where substr(item_id,5,5)='02241';/*2030单机架轧硬卷归并到2030轧硬卷*/
update tmp5 set item_id=substr(item_id,1,4)||'02151' where substr(item_id,5,5)='02144';/*2030单机架轧硬卷归并到2030轧硬卷*/
quit;


proc sql;
		create table tmp6 as select jz_name,jz_code,type_name,type_code,item_id,sum(val) as val from tmp5 group by jz_name,jz_code,
		type_name,type_code,item_id;
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
update tmp7 set dh_val=0-dh_val where type_code in ('04','05');
quit;


/*水煤处理等处理 */
proc sql;
	create table temp3 as select a.*,b.pro_code,item_name from tmp7 a left join wh_lc.wh_lca_0007 b
	on substr(a.item_id,7,3)=b.item_code and a.type_code=b.type_code;
quit;

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
where ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6)
group by pro_code,type_code;
quit;
proc sql;
create table lca0009 as select avg(hot_value) as hot_value,pro_code,type_code
from wh_lc.wh_lca_0009
where ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6)
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
root="&root";
shopsign_code="&shopsign_code";
run;
proc sql;
create table g1_fc as select a.*,b.id from g1_fc a,wh_lc.wh_lca_0005 b where a.pro_code=b.pro_code;
quit;


proc sql;
create table temp5 as select a.jz_name,a.jz_code,a.dh_val*b.co2_ratio*1000 as dh_val,b.pro_code,b.type_code
from tmp7 a join lca0001 b
on substr(a.item_id,7,3)=b.pro_code and substr(a.item_id,5,2)=b.type_code;

create table temp51 as select a.jz_name,a.jz_code,a.dh_val*b.hot_value as dh_val,b.pro_code,b.type_code
from tmp7 a join lca0009 b
on substr(a.item_id,7,3)=b.pro_code and substr(a.item_id,5,2)=b.type_code;
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



/*矩阵G1*/
proc sql;
create table lca0002 as select jz_code,id from wh_lca_0002  order by id;
quit;

proc sql;
create table g1 as select a.jz_code,a.id,b.pro_code,b.dh_val
from lca0002 a left join temp6 b
on a.jz_code=b.jz_code
order by id;
quit;


/*环保报表数据*/
proc sql;
create table lca0003 as select a.pro_code,a.var,b.jz_code 
from wh_lc.wh_lca_0003 a,wh_lca_0002 b
where a.ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6) and a.jz_code=b.jz_code_dl;
quit;
proc sql;
create table lca0003 as select a.pro_code,a.var,b.jz_code 
from wh_lc.wh_lca_0003 a,wh_lca_0002 b
where a.ymonth between substr("&datefrom",1,6) and substr("&dateto",1,6) and a.jz_code=b.jz_code_dl;
	create table lca0003 as select pro_code,avg(var) as var,jz_code
	from lca0003
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
	create table lca00031 as select a.*,b.id from lca0003 a left join wh_lca_0002 b
	on a.jz_code=b.jz_code ;
quit;
proc sql;	
	create table lca00031 as select a.*,a.var/b.val_all as dh_val from lca00031 a left join temp2 b
	on a.jz_code=b.jz_code;
quit;

proc sql;
	create table lca00031 as select a.*,b.id from lca00031 a left join wh_lca_0002 b 
	on a.jz_code=b.jz_code  ;
quit;
data g11;
set g1 lca00031;
run;

proc sql;
		create table g11a as
		select jz_code,id,dh_val,pro_code from g11;
	quit;

proc sort data=g11a; by pro_code  ;run;

PROC TRANSPOSE data=g11a out=g1a(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;

data result1;
retain pro_code col1-col123;
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
proc sql;
delete from ma_lc.lca_g1 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and type="&type" and root="&root" and shopsign_code="&shopsign_code";
quit;
proc append base=ma_lc.lca_g1 data=g1_fc force;run;

/*G2：PE2（PA2（I-PA1)^） +GASup(GAS（I-PA1)^)
PA1:本工序的原材料、辅料来自其他主工序的产品
GASup:0014读取，3列53行，（焦炉、高炉、转炉煤气）
GAS：3行123列
*/
data lca0006;
	set wh_lc.wh_lca_0006;
	where type="&type" and pro_code^='TTT' and flag='1' ;
run;
%macro gx1;
%if"&root"="DL"  %then %do;

%end;
%else %do;
/*根据大类代码勾连典型牌号jz_code*/
proc sql;
create table lca0006 as select a.jz_code, b.jz_code as jz_code_ph,a.pro_code,a.pro_name,a.type,a.type_code,a.flag
from lca0006 a left join sa_lc_00061_sum b on a.jz_code=b.jz_code_dl;
update lca0006 set jz_code=jz_code_ph where jz_code_ph^='';
quit;
%end;
%mend;
%gx1;
proc sql;
create table lca0006_1 as select distinct type_code,pro_code from wh_lc.wh_lca_0006
	where type="&type"  and pro_code^='TTT' and flag='1' ;
quit;

proc sql;
	create table g2_tmp1 as select jz_code,id from wh_lca_0002   order by id;
quit;
proc sql;
	create table g2_tmp2 as select a.jz_code,a.dh_val,substr(a.item_id,7,3) as pro_code,a.type_code
	from tmp7 a ,lca0006_1 b

	where a.type_code=b.type_code and   substr(a.item_id,7,3)=b.pro_code;
quit;

proc sql;
	create table g2_tmp3 as select a.jz_code,a.dh_val ,b.jz_code as pro_code,a.pro_code as item
	from g2_tmp2 a  join lca0006 b on a.pro_code=b.pro_code and a.type_code=b.type_code;
quit;
proc sql;
	create table g2_tmp4 as select a.jz_code,a.id,b.dh_val,b.pro_code,b.item from g2_tmp1 a left join g2_tmp3 b
	on a.jz_code=b.jz_code;
quit;

/*转置*/

proc sort data=g2_tmp4; by pro_code  ;run;

PROC TRANSPOSE data=g2_tmp4 out=g2_tmp6(drop=_name_)  prefix=col;
by pro_code  ;
var  dh_val;
id  id;
run;
/*固定表格式*/

proc sql;
create table g2_tmp7 as select a.jz_code,a.id,b.*
from g2_tmp1 a left join g2_tmp6 b
on a.jz_code=b.pro_code
order by id;
quit;
data g2_tmp8;
	retain  id jz_code col1-col123 ;
	set g2_tmp7;
	drop pro_code;
run;

data result_pa1;
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


/*PA2：本主工序的能源来自其他能源工序的产品*/



data lca0006_ny;
	set wh_lc.wh_lca_0006;
	where type='ny' and pro_code^='TTT' and flag='1';
run;
proc sql;
create table lca0006_ny_1 as select distinct type_code,pro_code from wh_lc.wh_lca_0006
	where type='ny'  and pro_code^='TTT' and flag='1';
quit;

proc sql;
	create table g2_tmp1a as select jz_code,id from wh_lca_0002   order by id;
quit;
proc sql;
	create table g2_tmp2a as select a.jz_code,a.dh_val,substr(a.item_id,7,3) as pro_code,a.type_code
	from tmp7 a ,lca0006_ny_1 b

	where a.type_code=b.type_code and   substr(a.item_id,7,3)=b.pro_code;
quit;

proc sql;
	create table g2_tmp3a as select a.jz_code,a.dh_val ,b.jz_code as pro_code,a.pro_code as item
	from g2_tmp2a a  join lca0006_ny b on a.pro_code=b.pro_code and a.type_code=b.type_code;
quit;
/*电特殊处理*/
proc sql;
create table g2_tmp3a as select a.*,b.bl from g2_tmp3a a left join temp23 b
on a.pro_code=b.jz_code and a.item='009';
quit;
data g2_tmp3a(drop=bl);
set g2_tmp3a;
if bl=. then dh_val=dh_val;
else dh_val=dh_val*bl;
run;
proc sql;
	create table g2_tmp4a as select a.jz_code,a.id,b.dh_val,b.pro_code,b.item from g2_tmp1a a left join g2_tmp3a b
	on a.jz_code=b.jz_code;
quit;

/*转置*/

proc sort data=g2_tmp4a; by pro_code jz_code ;run;

PROC TRANSPOSE data=g2_tmp4a out=g2_tmp6a(drop=_name_)  prefix=col;
by pro_code ;
var  dh_val;
id  id;
run;
/*固定表格式*/
proc sql;
	create table tmp_ny as select jz_code,id from wh_lc.wh_lca_0002 where type='ny'  and id<=17  and id not in (5,6,7,4) order by id;
quit;
proc sql;
create table g2_tmp7a as select a.jz_code,a.id,b.*
from tmp_ny a left join g2_tmp6a b
on a.jz_code=b.pro_code
order by id;
quit;
data g2_tmp8a;
	retain  id jz_code col1-col123 ;
	set g2_tmp7a;
	drop pro_code;
run;

data result_pa2;
set g2_tmp8a;
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

/*PE2：能源site（不包含发电工序\一二三四鼓风）*/
data result_pe2_fc;
set ma_lc.sa_lca_ny_gsite(drop=col18 col5 col6 col7 col4);
retain id col1 col2 col3 col4  col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
run;
proc sort data=result_pe2_fc ;by id;run;

data result_pe2;
set result_pe2_fc;
drop id ;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;
/*GASup*/
data GAS1;
set wh_lc.wh_lca_0014;
where product in ('焦炉煤气','高炉煤气','转炉煤气');
keep id pro_code sum;
run;

proc sql;
create table gas2 as select a.*,b.id as id_new from gas1 a,wh_lc.wh_lca_0005 b
where a.pro_code=b.pro_code
order by id ;
quit;


/*转置*/

proc sort data=gas2;by id_new id ;run;

PROC TRANSPOSE data=gas2 out=result_gasup(drop=_name_ _LABEL_ id_new)  prefix=col;
by id_new ;
var  sum;
id  id;
run;
/*gas*/
/*高炉：78 焦炉101 转炉263*/
data gas3;
set tmp7;
where substr(item_id,5,5) in ('01019','01032','01011');
if substr(item_id,5,5)='01019' then  pro_code=101;
if substr(item_id,5,5)='01032' then  pro_code=263;
if substr(item_id,5,5)='01011' then  pro_code=78;
keep jz_code dh_val pro_code;
run;


proc sql;
	create table gas4 as select a.*,b.id from gas3 a right join wh_lca_0002 b
	on a.jz_code=b.jz_code
order by pro_code,id;
quit;


/*转置*/

proc sort data=gas4;by pro_code id ;run;

PROC TRANSPOSE data=gas4 out=result_gas(drop=_name_ )  prefix=col;
by pro_code ;
var  dh_val;
id  id;
run;

data result_gas;
retain col1-col123;

set result_gas;
where pro_code not in (.);
run;

data result_gas;
set result_gas;
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

/*G2*/
proc iml ;

	use result_pa1;
	read all var _all_ into pa1;
	use result_pa2;
	read all var _all_ into pa2;
	use result_pe2;
	read all var _all_ into pe2;
	use result_gasup;
	read all var _all_ into gasup;
	use result_gas;
	read all var _all_ into gas;
	use result_g1;
	read all var _all_ into g1;
	I=I(&id_num);
	g21=pe2*(pa2*inv(I-pa1));
	g22=gasup*(gas*inv(I-pa1));
	g2=g21+g22;
	
	create result_g22 from g22;
	append from g22;
	bbb=gas*inv(I-pa1);
	create result_bbb from bbb;
	append from bbb;
	ttt=inv(I-pa1);
	create result_ttt from ttt;
	append from ttt;
	ttt1=g1+pe2*pa2;
	create result_ttt1 from ttt1;
	append from ttt1;
	fff=(pa2*inv(I-pa1));
	create result_fff from fff;
	append from fff;
	create result_g21 from g21;
	append from g21;
	create result_g2 from g2;
	append from g2;
quit;
data result_fff;
	set result_fff;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	id1=_N_;
run;
data result_bbb;
	set result_bbb;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
		id1=_N_;
run;
/*pe2转置*/

proc sort data=result_pe2_fc; by id ;run;

PROC TRANSPOSE data=result_pe2_fc out=result_pe2_fc1  prefix=col;
by id ;
run;
data result_pe2_fc1;
	set result_pe2_fc1;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g21 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.jz_name as item_name
from result_pe2_fc1 a,wh_lc.wh_lca_0002 b
where a.name=b.id and b.type='ny';
quit;
proc sort data=lca_g21 ;by id name ;run;
data lca_g21;
set lca_g21;
by id;
if first.id then id1=0;
id1+1;
run;
data result_gasup;
set result_gasup;
id=_n_;
run;
proc sort data=result_gasup; by id ;run;

PROC TRANSPOSE data=result_gasup out=result_gasup1  prefix=col;
by id ;
run;
data result_gasup1;
	set result_gasup1;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
/*高炉：78 焦炉101 转炉263*/

proc sql;
create table lca_g211 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,case when name=78 then '高炉煤气' when name=101 then '焦炉煤气' when name=263 then '转炉煤气' end as item_name
from result_gasup1 a;
quit;
proc sort data=lca_g211;by id;run;
data lca_g211;
set lca_g211;
by id;
if first.id then id1=0;
id1+1;
run;

%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g21_r as select a.*,b.col&i as bf,c.jz_code   from lca_g21 a ,result_fff b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc sql;
create table lca_g22_r as select a.*,b.col&i as bf,c.jz_code   from lca_g211 a ,result_bbb b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
%end;
%else %do;
proc sql;
create table lca_g21_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g21 a ,result_fff b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc sql;
create table lca_g22_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g211 a ,result_bbb b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc append base=lca_g21_r data=lca_g21_&i force;run;
proc append base=lca_g22_r data=lca_g22_&i force;run;

%end;

%let i=%eval(&i+1);

%end;
%mend;
%cc;
data lca_g2(drop=id1);
set lca_g21_r lca_g22_r;
type="&type";
root="&root";
shopsign_code="&shopsign_code";
run;
proc sql;
delete from ma_lc.lca_g2 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root" and shopsign_code="&shopsign_code";
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
create table result_tttt as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as bf,b.jz_code,"&type" as type,"&root" as root,"&shopsign_code" as shopsign_code
from result_tttt a,wh_lca_0002 b
where a.name=b.id ;
quit;
data result_ttt1;
set result_ttt1;
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
create table result_ttt2 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.jz_name as item_name,"&type" as type,"&root" as root,"&shopsign_code" as shopsign_code
from result_ttt2 a,wh_lca_0002 b
where a.name=b.id ;
quit;

proc sql;
delete from ma_lc.lca_fb1a where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root" and shopsign_code="&shopsign_code";
quit;
proc append base=ma_lc.lca_fb1a data=result_ttt2 force;run;
proc sql;
delete from ma_lc.lca_fb1b where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root" and shopsign_code="&shopsign_code";
quit;
proc append base=ma_lc.lca_fb1b data=result_tttt force;run;



/***G3:g1*(I-PA1)^-I）****/


proc iml;

	use result_pa1;
	read all var _all_ into pa1;
	use result_g1;
	read all var _all_ into g1;
	I=I(&id_num);
	g32=inv(I-pa1)-I;
	g3=g1*g32;
	create result_g31 from g1;
	append from g1;
create result_g32 from g32;
	append from g32;
	create result_g3 from g3;
	append from g3;
quit;
data result_g31;
set result_g31;
id=_n_;
run;

proc sort data=result_g31; by id ;run;

PROC TRANSPOSE data=result_g31 out=result_g311  prefix=col;
by id ;
run;
data result_g311;
	set result_g311;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g31 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.jz_name as item_name
from result_g311 a ,wh_lca_0002 b
where a.name=b.id ;
quit;
proc sort data=lca_g31;by id;run;
data lca_g31;
set lca_g31;
by id;
if first.id then id1=0;
id1+1;
run;
data result_g32;
	set result_g32;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	id1=_N_;
run;
%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g31_r as select a.*,b.col&i as bf,c.jz_code   from lca_g31 a ,result_g32 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
%end;
%else %do;
proc sql;
create table lca_g31_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g31 a ,result_g32 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc append base=lca_g31_r data=lca_g31_&i force;run;
%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;
/*proc sql;
delete from ma_lc.lca_g231 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6);
quit;
proc append base=ma_lc.lca_g31 data=lca_g31 force;run;
proc sql;
delete from ma_lc.lca_g32 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6);
quit;
proc append base=ma_lc.lca_g32 data=result_g32 force;run;
*/

data lca_g3(drop=id1);
set lca_g31_r;
type="&type";
root="&root";
shopsign_code="&shopsign_code";
run;
proc sql;
delete from ma_lc.lca_g3 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root" and shopsign_code="&shopsign_code";
quit;
proc append base=ma_lc.lca_g3 data=lca_g3 force;run;

/*******G4:RG(RA*(I-PA1)^)
           RA:主工序用到的其他工序产出的副产品，该副产品不可以是任何工序的产品:目前按照中文名字，后续需要优化
		RG:维护表（UP表中读取）

*/
data WH_LCA_0014;
set wh_lc.WH_LCA_0014;
run;
/*去除电，副产品产出的电不算*/
proc sql;
create table tmp71  as select a.*,b.pro_code_name  from tmp7 a  join ma_lc.wh_lcdm0009 b on a.type_code=b.type_code and 
	substr(a.item_id,7,3)=b.pro_code where substr(a.item_id,5,5) not in ('05011','05049');

	
create table tmp_ra1 as select distinct substr(item_id,7,3) as pro_code,type_code,pro_code_name from tmp71 where type_code in ('01','02','03')
and pro_code_name in (select distinct product from WH_LCA_0014);
create table tmp_ra2 as select jz_code,substr(item_id,7,3) as pro_code,type_code,0-dh_val as dh_val,type_code,pro_code_name from tmp71 where 
pro_code_name in (select distinct pro_code_name from tmp_ra1) and type_code in ('05');
quit;



proc sql;
create table lca_0014_1 as select distinct id,product from WH_LCA_0014;
/*create table tmp_ra3 as select a.*,b.id from tmp_ra2 a left join lca_0014_1 b on a.pro_code_name=b.product;*/
quit;
proc sql;
	create table g4_tmp1a as select jz_code,id from wh_lca_0002   order by id;
quit;
proc sql;
	create table tmp_ra3 as select a.jz_code,a.id,b.dh_val,b.pro_code,b.pro_code_name from g4_tmp1a a left join tmp_ra2 b
	on a.jz_code=b.jz_code;
quit;

/*转置*/

proc sort data=tmp_ra3; by pro_code_name ;run;

PROC TRANSPOSE data=tmp_ra3 out=tmp_ra4(drop=_name_)  prefix=col;
by pro_code_name ;
var  dh_val;
id  id;
run;

/*固定表格式*/
proc sql;
create table tmp_ra41 as select a.*,b.id from tmp_ra4 a left join lca_0014_1 b on a.pro_code_name=b.product where a.pro_code_name^='';
quit;
proc sort data=tmp_ra41;by id;run;
data tmp_ra5;
	retain   col1-col123 ;
	set tmp_ra41;
	drop pro_code_name;
run;

data result_ra;
set tmp_ra5;
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

/**RG:维护表wh_lc.wh_lca_0014*/
proc sql;
create table tmp_rg1 as select * from WH_LCA_0014 where product in (select distinct pro_code_name from tmp_ra41) and pro_code ^='';
quit;


/*转置*/

proc sort data=tmp_rg1; by pro_code ;run;

PROC TRANSPOSE data=tmp_rg1 out=tmp_rg2(drop=_name_)  prefix=col;
by pro_code ;
var  SUM;
id  id;
run;
/*分析使用*/
proc sql;
create table lca_g41 as select substr("&datefrom",1,6) as LCItime_s,substr("&dateto",1,6) as LCItime_e,a.id as name,a.product as item_name,a.sum as af,b.id from  tmp_rg1 a join wh_lc.wh_lca_0005 b
on a.pro_code=b.pro_code;
quit;
proc sort data=lca_g41 ;by id name;run;
/*固定表格式*/
proc sql;
create table tmp_rg21 as select a.*,b.id from tmp_rg2 a right join wh_lc.wh_lca_0005 b on a.pro_code=b.pro_code;
quit;
proc sort data=tmp_rg21;by id;run;
data tmp_rg3;
	retain   col1-col300 ;
	set tmp_rg21;
	drop _label_;
run;

data result_rg;
set tmp_rg3;
drop pro_code id;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;

/***G4:RG(RA*(I-PA1)^)****/

proc iml;
	
	use result_rg;
	read all var _all_ into rg;
	use result_ra;
	read all var _all_ into ra;
	use result_pa1;
	read all var _all_ into pa1;
	I=I(&id_num);
	t=ra*inv(I-pa1);
	t11=inv(I-pa1);
	g4=rg*t;
	create result_t11 from t11;
	append from t11;
	create result_t from t;
	append from t;
	create result_g4 from g4;
	append from g4;
quit;
data result_t;
set result_t;
LCItime_s=substr("&datefrom",1,6);
LCItime_e=substr("&dateto",1,6);
id1=_N_;
run;
proc sort data=lca_g41;by id;run;
data lca_g41;
set lca_g41;
by id;
if first.id then id1=0;
id1+1;
run;

%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g41_r as select a.*,b.col&i as bf,c.jz_code   from lca_g41 a ,result_t b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
%end;
%else %do;
proc sql;
create table lca_g41_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g41 a ,result_t b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc append base=lca_g41_r data=lca_g41_&i force;run;
%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;
data lca_g4(drop=id1);
set lca_g41_r;
type="&type";
root="&root";
shopsign_code="&shopsign_code";
run;
proc sql;
delete from ma_lc.lca_g4 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root" and shopsign_code="&shopsign_code";
quit;
proc append base=ma_lc.lca_g4 data=lca_g4 force;run;


/*****************G5:（T*L)*(OA(I-PA1)^1)+PF2(PA2(I-PA1)^1)*******************/
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
/*矩阵OA:工序中消耗量，不算输出量*/

proc sql;
	create table result_oa1 as select a.jz_code,a.dh_val,case when b.pro_code='coal' then compress(b.pro_code||b.item_code) else b.pro_code
	 end as pro_code from tmp7 a join wh_lc.wh_lca_0007 b 
	on substr(a.item_id,7,3)=b.item_code and  a.type_code=b.type_code and a.type_code in ('01','02','03') and b.pro_code in 
	('ironstone','limestone','dolomite','coal');
quit;
proc sql;
	create table result_oa1 as select sum(dh_val) as dh_val,jz_code,pro_code from result_oa1 group by jz_code,pro_code;
quit;
proc sql;
create table result_oa2 as select a.dh_val,a.pro_code,b.id from result_oa1 a right join wh_lca_0002 b on a.jz_code=b.jz_code;
quit;
/*转置*/

proc sort data=result_oa2; by   pro_code ;run;

PROC TRANSPOSE data=result_oa2 out=result_oa3(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;
proc sql;
create table wh_lca_0011 as select avg(train) as train,avg(barge) as barge,avg(seacraft) as seacraft,avg(car) as car,pro_code,pro_name,unit,id
from wh_lc.wh_lca_0011
where ymonth  between substr("&datefrom",1,6) and substr("&dateto",1,6) 
group by pro_code,pro_name,unit,id;
quit; 
proc sql;
	create table result_oa4 as select a.*,b.id
	from result_oa3 a right join wh_lca_0011 b
	on a.pro_code=b.pro_code
order by b.id;
quit;
data result_oa5;
retain id col1-col123  ;
set result_oa4;
run;

data result_oa(drop=id pro_code);
set result_oa5;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run; 
/*PF2: ma_lc.sa_lca_ny_g5*/
data result_PF2;
set  ma_lc.sa_lca_ny_g5(drop=col18 col5 col6 col7 col4);
retain id col1 col2 col3 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;

run;

proc sort data=result_PF2 ;by id;run;

data result_PF2;
set result_PF2;
drop id ;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;

/***G5:（T*L)*(OA(I-PA1)^1)+PF2(PA2(I-PA1)^1)****/

proc iml;
	
	use result_t;
	read all var _all_ into t;
	use result_l;
	read all var _all_ into l;
	use result_oa;
	read all var _all_ into oa;
	use result_pa1;
	read all var _all_ into pa1;
	use result_pa2;
	read all var _all_ into pa2;
		use result_pf2;
	read all var _all_ into pf2;
	I=I(&id_num);
	a=t*l;
	b=oa*inv(I-pa1);
	create result_a from a;
	append from a;
	create result_b from b;
	append from b;
	c=a*b;
	d=pa2*inv(I-pa1);
	create result_d from d;
	append from d;

	e=pf2*d;
	g5=c+e;
	
	create result_g5 from g5;
	append from g5;

quit;
/*pe2转置*/
data result_a;
set result_a;
id=_n_;
run;
proc sort data=result_a; by id ;run;

PROC TRANSPOSE data=result_a out=result_a1  prefix=col;
by id ;
run;
data result_a1;
	set result_a1;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g51 as select distinct  a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.pro_name as item_name
from result_a1 a,wh_lc.wh_lca_0011 b
where a.name=b.id ;
quit;

proc sort data=lca_g51;by id;run;
data lca_g51;
set lca_g51;
by id;
if first.id then id1=0;
id1+1;
run;
data result_b;
	set result_b;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
id1=_N_;
run;


data result_pf2;
set result_pf2;
id=_n_;
run;
proc sort data=result_pf2; by id ;run;

PROC TRANSPOSE data=result_pf2 out=result_pf21  prefix=col;
by id ;
run;
data result_pf21;
	set result_pf21;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g511 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.jz_name as item_name
from result_pf21 a,wh_lc.wh_lca_0002 b
where a.name=b.id and b.type='ny';
quit;
proc sort data=lca_g511;by id;run;
data lca_g511;
set lca_g511;
by id;
if first.id then id1=0;
id1+1;
run;
data lca_g521;
	set result_d;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
id1=_N_;
run;


%macro cc;
%let i=1;
%do %while(&i<=&id_num);
%if &i=1 %then %do;
proc sql;
create table lca_g51_r as select a.*,b.col&i as bf,c.jz_code   from lca_g51 a ,result_b b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc sql;
create table lca_g52_r as select a.*,b.col&i as bf,c.jz_code   from lca_g511 a ,lca_g521 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
%end;
%else %do;
proc sql;
create table lca_g51_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g51 a ,result_b b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc sql;
create table lca_g52_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g511 a ,lca_g521 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc append base=lca_g51_r data=lca_g51_&i force;run;
proc append base=lca_g52_r data=lca_g52_&i force;run;
%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;

data lca_g5(drop=id1);
set lca_g51_r lca_g52_r;
type="&type";
root="&root";
shopsign_code="&shopsign_code";
run;

proc sql;
delete from ma_lc.lca_g5 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root" and shopsign_code="&shopsign_code";
quit;
proc append base=ma_lc.lca_g5 data=lca_g5 force;run;

/*************************************G6:OG*(OA1(I-PA1)^1)+PG2(PA2(I-PA1)^1)************************************/
/*OA1:同OA，所有外购（不是其他工序的产品、副产品）*/
/*PG2:能源G6*/
/*OG:维护表（UP表中读取）*/

/*矩阵OA1:工序中消耗量，不算输出量，所有外购（不是其他工序的产品）,去除三个煤气,氮气*/

proc sql;
	create table result_oa11 as select a.jz_code,a.dh_val,substr(a.item_id,5,5) as pro_code,pro_code_name from tmp71 a 
where a.type_code in ('01','02','03') and a.pro_code_name not in (select distinct pro_code_name from tmp11_all where type_code in ('04'))
and substr(a.item_id,5,5) not in ('01011','01019','01032','01034','01006');
quit;

proc sql;
update result_oa11 set pro_code_name='球团矿' where pro_code_name='外购球团矿';
update result_oa11 set pro_code_name='喷吹煤' where pro_code_name='外购焦炭';
update result_oa11 set pro_code_name='喷吹煤' where pro_code_name='炼焦煤';
update result_oa11 set pro_code_name='铁精矿' where pro_code_name='外购精块矿';
update result_oa11 set pro_code_name='生石灰' where pro_code_name='外购生石灰';
update result_oa11 set pro_code_name='石灰石' where pro_code_name='(未选)石灰石';
update result_oa11 set pro_code_name='石灰石细粉' where pro_code_name='石灰石细粒';
update result_oa11 set pro_code_name='石灰石细粉' where pro_code_name='粉石灰石';
update result_oa11 set pro_code_name='石灰石细粉' where pro_code_name='脱硫石灰石粉';
update result_oa11 set pro_code_name='生石灰' where pro_code_name='块状生石灰';
update result_oa11 set pro_code_name='生石灰' where pro_code_name='外购块生石灰';
update result_oa11 set pro_code_name='白云石' where pro_code_name='外购白云石细粒';
update result_oa11 set pro_code_name='轻烧白云石' where pro_code_name='外购轻烧白云石';

update result_oa11 set pro_code_name='白云石粉(片)' where pro_code_name='(片)白云石粉';

quit;
proc sql;
create table result_oa12 as select distinct a.*,b.id as pro_code_new from result_oa11 a   join wh_lc.wh_lca_0014 b on a.pro_code_name=b.product ;
quit;
proc sql;
create table result_oa12 as select a.dh_val,a.pro_code_name,a.pro_code,a.pro_code_new,b.id from result_oa12 a right join wh_lca_0002 b 
on a.jz_code=b.jz_code ;
quit;
proc sql;
create table result_oa13 as select sum(dh_val) as dh_val,pro_code_new,id from result_oa12 group by pro_code_new,id;
quit;
/*转置*/

proc sort data=result_oa13; by   pro_code_new  ;run;

PROC TRANSPOSE data=result_oa13 out=result_oa14(drop=_name_)  prefix=col;
by pro_code_new;
var  dh_val;
id  id;
run;

data result_oa14;
retain  col1-col123  ;
set result_oa14;
where pro_code_new^=.;
run;

data result_oa1(drop= pro_code_new);
set result_oa14;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run; 
/*OG:需要注意OG顺序要与OA1对应*/

proc sql;
create table tmp_og1 as select * from WH_LCA_0014 where product in (select distinct pro_code_name from result_oa12) and pro_code ^='';
quit;


/*转置*/

proc sort data=tmp_og1; by pro_code ;run;

PROC TRANSPOSE data=tmp_og1 out=tmp_og2(drop=_name_)  prefix=col;
by pro_code ;
var  SUM;
id  id;
run;

/*固定表格式*/
proc sql;
create table tmp_og3 as select a.*,b.id from tmp_og2 a right join wh_lc.wh_lca_0005 b on a.pro_code=b.pro_code;
quit;
proc sort data=tmp_og3;by id;run;
data tmp_og3;
	retain   col1-col300 ;
	set tmp_og3;
	drop _label_;
run;

data result_og;
set tmp_og3;
drop pro_code id;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;

/*PG2:能源G6*/

data result_pg2;
set ma_lc.sa_lca_ny_g6(drop= col4 col5 col6 col7 col18);
retain id col1 col2 col3 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
run;
proc sort data=result_pg2 ;by id;run;

data result_pg2;
set result_pg2;
drop id ;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;
/*G6:OG*(OA1(I-PA1)^1)+PG2(PA2(I-PA1)^1)*/

proc iml;

	use result_og;
	read all var _all_ into og;
	use result_oa1;
	read all var _all_ into oa1;
	use result_pa1;
	read all var _all_ into pa1;
	use result_pg2;
	read all var _all_ into pg2;
	use result_pa2;
	read all var _all_ into pa2;
	
	I=I(&id_num);
	a=inv(I-pa1);
	b=oa1*a;
	
	
c=og*b;
d=pa2*a;
e=pg2*d;
	create result_d from d;
	append from d;
	create result_a from a;
	append from a;
	create result_b from b;
	append from b;	
	g6=c+e;	
	create result_g6 from g6;
	append from g6;
quit;

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
	set result_b;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	id1=_N_;
run;


data result_pg2;
set result_pg2;
id=_n_;
run;
proc sort data=result_pg2; by id ;run;

PROC TRANSPOSE data=result_pg2 out=result_pg21  prefix=col;
by id ;
run;
data result_pg21;
	set result_pg21;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g611 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.jz_name as item_name
from result_pg21 a,wh_lc.wh_lca_0002 b
where a.name=b.id and b.type='ny';
quit;
proc sort data=lca_g611;by id;run;
data lca_g611;
set lca_g611;
by id;
if first.id then id1=0;
id1+1;
run;
data lca_g621;
	set result_d;
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
proc sql;
create table lca_g62_r as select a.*,b.col&i as bf,c.jz_code   from lca_g611 a ,lca_g621 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
%end;
%else %do;
proc sql;
create table lca_g61_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g61 a ,lca_g62 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc sql;
create table lca_g62_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g611 a ,lca_g621 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc append base=lca_g61_r data=lca_g61_&i force;run;
proc append base=lca_g62_r data=lca_g62_&i force;run;
%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;

data lca_g6(drop=id1);
set lca_g61_r lca_g62_r;
type="&type";
root="&root";
shopsign_code="&shopsign_code";
run;

proc sql;
delete from ma_lc.lca_g6 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root" and shopsign_code="&shopsign_code";
quit;
proc append base=ma_lc.lca_g6 data=lca_g6 force;run;
/*****************G7:RG1*(RA1(I-PA1)^1)+PH2(PA2(I-PA1)^1)*************************/
/*
RA1:卖到外面的副产品（产出的副产品没有在任何工序中作为投入,电不算）
RG1:维护表（根据RA1中数据到UP表中读取）
PH2:能源G7
*/

/*RA1:*/

proc sql;
create table tmp72  as select a.*,b.pro_code_name  from tmp7 a  join ma_lc.wh_lcdm0009 b on a.type_code=b.type_code and 
	substr(a.item_id,7,3)=b.pro_code ;
	create table result_ra11 as select a.jz_code,a.dh_val,substr(a.item_id,5,5) as pro_code,pro_code_name from tmp72 a 
where a.type_code in ('05','08') and (a.pro_code_name not in (select distinct pro_code_name from tmp11_all where type_code in ('01','02','03'))
or substr(a.item_id,5,5) in ('05011') );
update result_ra11 set pro_code_name='高炉渣' where pro_code_name in ('高炉干渣','高炉水渣');
update result_ra11 set pro_code_name='生石灰' where pro_code_name in ('焙烧生石灰除尘粉','焙烧窑皮');

update result_ra11 set dh_val=0-dh_val where dh_val<0;
quit;


proc sql;
create table result_ra12 as select distinct a.*,b.id as pro_code_new from result_ra11 a   join wh_lc.wh_lca_0014 b on a.pro_code_name=b.product ;
quit;
proc sql;
create table result_ra12 as select a.dh_val,a.pro_code_name,a.pro_code,a.pro_code_new,b.id from result_ra12 a right join wh_lca_0002 b
 on a.jz_code=b.jz_code ;
quit;
proc sql;
create table result_ra13 as select sum(dh_val) as dh_val,pro_code_new,id from result_ra12 group by pro_code_new,id;
quit;
/*转置*/

proc sort data=result_ra13; by   pro_code_new  ;run;

PROC TRANSPOSE data=result_ra13 out=result_ra14(drop=_name_)  prefix=col;
by pro_code_new;
var  dh_val;
id  id;
run;

data result_ra14;
retain  col1-col123  ;
set result_ra14;
where pro_code_new^=.;
run;

data result_ra1(drop= pro_code_new);
set result_ra14;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run; 


/*RG1*/


proc sql;
create table tmp_rg1 as select * from WH_LCA_0014 where product in (select distinct pro_code_name from result_ra12) and pro_code ^='';
quit;


/*转置*/

proc sort data=tmp_rg1; by pro_code ;run;

PROC TRANSPOSE data=tmp_rg1 out=tmp_rg2(drop=_name_)  prefix=col;
by pro_code ;
var  SUM;
id  id;
run;

/*固定表格式*/
proc sql;
create table tmp_rg3 as select a.*,b.id from tmp_rg2 a right join wh_lc.wh_lca_0005 b on a.pro_code=b.pro_code;
quit;
proc sort data=tmp_rg3;by id;run;
data tmp_rg3;
	retain   col1-col300 ;
	set tmp_rg3;
	drop _label_;
run;

data result_rg1;
set tmp_rg3;
drop pro_code id;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;

/*PH2*/

data result_ph2;
set ma_lc.sa_lca_ny_g7(drop= col4 col5 col6 col7 col18);
retain id col1 col2 col3 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
run;
proc sort data=result_ph2 ;by id;run;

data result_ph2;
set result_ph2;
drop id ;
array char _CHARACTER_;
array numr _NUMERIC_;
do over char;
if char eq "" then char="0";
end;
do over numr;
if numr eq . then numr=0;
end;
run;

/*G7:RG1*(RA1(I-PA1)^1)+PH2(PA2(I-PA1)^1)*/
proc iml;

	use result_ra1;
	read all var _all_ into ra1;
	use result_rg1;
	read all var _all_ into rg1;
	use result_pa1;
	read all var _all_ into pa1;
	use result_ph2;
	read all var _all_ into ph2;
	use result_pa2;
	read all var _all_ into pa2;
	
	I=I(&id_num);
	a=inv(I-pa1);
	b=ra1*a;
	c=rg1*b;

	d=pa2*a;
	e=ph2*d;
	g7=c+e;
	create result_b from b;
	append from b;
	create result_d from d;
	append from d;
	create result_g7 from g7;
	append from g7;

quit;



data result_rg11;
set result_rg1;
id=_n_;
run;
proc sort data=result_rg11; by id ;run;

PROC TRANSPOSE data=result_rg11 out=result_rg12 prefix=col;
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
	set result_b;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	id1=_N_;
run;

data result_ph2;
set result_ph2;
id=_n_;
run;
proc sort data=result_ph2; by id ;run;

PROC TRANSPOSE data=result_ph2 out=result_ph21  prefix=col;
by id ;
run;
data result_ph21;
	set result_ph21;
	LCItime_s=substr("&datefrom",1,6);
	LCItime_e=substr("&dateto",1,6);
	name=substr(_name_,4)+0;
run;
proc sql;
create table lca_g711 as select a.id,a.name,a.LCItime_s,a.LCItime_e,a.col1 as af,b.jz_name as item_name
from result_ph21 a,wh_lc.wh_lca_0002 b
where a.name=b.id and b.type='ny';
quit;
proc sort data=lca_g711;by id;run;
data lca_g711;
set lca_g711;
by id;
if first.id then id1=0;
id1+1;
run;
data lca_g721;
	set result_d;
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
proc sql;
create table lca_g72_r as select a.*,b.col&i as bf,c.jz_code   from lca_g711 a ,lca_g721 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
%end;
%else %do;
proc sql;
create table lca_g71_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g71 a ,lca_g72 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc sql;
create table lca_g72_&i as select a.*,b.col&i as  bf,c.jz_code  from lca_g711 a ,lca_g721 b,wh_lca_0002 c
where a.id1=b.id1 and c.id=&i;
quit;
proc append base=lca_g71_r data=lca_g71_&i force;run;
proc append base=lca_g72_r data=lca_g72_&i force;run;
%end;
%let i=%eval(&i+1);
%end;
%mend;
%cc;

data lca_g7(drop=id1);
set lca_g71_r lca_g72_r;
type="&type";
root="&root";
shopsign_code="&shopsign_code";
run;

proc sql;
delete from ma_lc.lca_g7 where LCItime_s=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and
type="&type" and root="&root" and shopsign_code="&shopsign_code";
quit;
proc append base=ma_lc.lca_g7 data=lca_g7 force;run;

/*基础矩阵计算结束*/

proc iml;
	
	use result_g1;
	read all var _all_ into g1;
	use result_g2;
	read all var _all_ into g2;
	use result_g3;
	read all var _all_ into g3;
	use result_g4;
	read all var _all_ into g4;
	use result_g5;
	read all var _all_ into g5;
	use result_g6;
	read all var _all_ into g6;
	use result_g7;
	read all var _all_ into g7;
	groute=g1+g2+g3-g4+g5+g6-g7;
	gsite=g1+g2+g3-g4;
	goutsite=g5+g6-g7;
	
	create result_groute from groute;
	append from groute;
	create result_gsite from gsite;
	append from gsite;
	create result_goutsite from goutsite;
	append from goutsite;
	
quit;

data result_goutsite;
retain id col1-col123;
id=_n_;
set result_goutsite;
run; 
data result_gsite;
retain id col1-col123;
id=_n_;
set result_gsite;
run; 

data result_groute;
retain id col1-col123;
id=_n_;
set result_groute;
run; 
data result_g1;
retain id col1-col123;
id=_n_;
set result_g1;
run; 
data result_g2;
retain id col1-col123;
id=_n_;
set result_g2;
run; 
data result_g3;
retain id col1-col123;
id=_n_;
set result_g3;
run; 
data result_g4;
retain id col1-col123;
id=_n_;
set result_g4;
run; 
data result_g5;
retain id col1-col123;
id=_n_;
set result_g5;
run; 

data result_g6;
retain id col1-col123;
id=_n_;
set result_g6;
run; 

data result_g7;
retain id col1-col123;
id=_n_;
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
proc sort data=result_g4 ;by id;run;
PROC TRANSPOSE data=result_g4 out=result_g42(rename=(_NAME_=jz_code col1=g4)) prefix=col;
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
proc sort data=result_g42;by  id jz_code;run;
proc sort data=result_g52;by  id jz_code;run;
proc sort data=result_g62;by  id jz_code;run;
proc sort data=result_g72;by  id jz_code;run;

data result_all;
merge result_groute2 result_gsite2 result_goutsite2 result_g12 result_g22 result_g32 result_g42 result_g52 result_g62 result_g72;
by id jz_code;
run;

/*勾连数据项*/
proc sql;
create table result_alla as select b.pro_code,b.pro_name as LCIfactor,b.unit,a.* from result_all a left join wh_lc.wh_lca_0005 b
on a.id=b.id;
quit;
proc sql;
create table result_alla as select substr("&datefrom",1,6) as LCItime_s,substr("&dateto",1,6) as LCItime_e,b.jz_code as jz_code1,b.jz_name as product,a.* 
from result_alla a left join wh_lca_0002 b
on input(substr(a.jz_code,4),3.)=b.id ;
quit;

data result_alla;
set result_alla;
drop id jz_code;
rename jz_code1=jz_code;
rec_create_time=put(date(),yymmddn8.);
type="&type";
root="DL";
shopsign_code="DL";
run;
proc sort data=result_alla ;by jz_code LCIfactor;run;


proc sql;
delete from ma_lc.sa_lca_0001 where LCITIME_S=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and type="&type" and root="DL" and shopsign_code="DL";
quit;
proc append base=ma_lc.sa_lca_0001 data=result_alla force;run;
/*版本查询*/

proc sql;
delete from  WH_LC.WH_LCA_0020 where LCITIME_S=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and type="&type" and root="&root" and shopsign_code="&shopsign_code" ;
quit;
proc sql;
create table edition as select distinct  substr("&datefrom",1,6) as LCITIME_S,substr("&dateto",1,6) as LCItime_e, substr("&datefrom",1,6)||'-'||substr("&dateto",1,6) as edition,
"&type" as type,"&root" as root,"&shopsign_code" as shopsign_code
from WH_LC.WH_LCA_0020;
quit;
proc append base=WH_LC.WH_LCA_0020 data=edition force;run;


%macro dl;
%if "&root"="DL" %then %do;

%end;
%else %do;

proc sql;
delete from  WH_LC.WH_LCA_0022 where LCITIME_S=substr("&datefrom",1,6) and LCItime_e=substr("&dateto",1,6) and root="&root" and shopsign_code="&shopsign_code" ;
quit;
proc sql;
create table root as select distinct  substr("&datefrom",1,6) as LCITIME_S,substr("&dateto",1,6) as LCItime_e, substr("&datefrom",1,6)||'-'||substr("&dateto",1,6) as edition,
"&type" as type,"&root" as root,"&shopsign_code" as shopsign_code,jz_code,jz_name,id
from WH_LCA_0002;
quit;
proc append base=WH_LC.WH_LCA_0022 data=root force;run;		
%end;
%mend;
%dl;
proc datasets library=work kill;
quit;



