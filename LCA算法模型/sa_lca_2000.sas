
/*******************************************主工序开始*********************************************************/
/***主工序数量：135 数据项个数：53**/

/*
%let datefrom=201701;
%let dateto=201712;
%let type=gx;*/

data sa_lc_0006;
	set ma_lc.sa_lc_0006 ;
	where compress(year||mon) between "&datefrom" and "&dateto" ;
run;

proc sql;
update sa_lc_0006 set item_id=substr(item_id,1,6)||'193' where substr(item_id,5,5)='02147';
update sa_lc_0006 set item_id=substr(item_id,1,6)||'022' where substr(item_id,5,5)='02018';
update sa_lc_0006 set item_id=substr(item_id,1,6)||'175' where substr(item_id,5,5)='02171';
update sa_lc_0006 set item_id=substr(item_id,1,6)||'192' where substr(item_id,5,5)='02157';
update sa_lc_0006 set item_id=substr(item_id,1,6)||'183' where substr(item_id,5,5)='03019';
quit;
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
data Wh_lcdm0003;
set ma_lc.Wh_lcdm0003;
 where process_code in ('LJ','SJ','GL','GF') and jz_code ne 'CO04' ;
run;
/*获取单位*/
proc sql;
	create table Wh_lcdm0006 as
	 select distinct jz_code,type_code,pro_code,unit
	 from  ma_lc.Wh_lcdm0006  
	where year=substr("&datefrom",1,4);
quit;

proc sql;
	create table tmp2 as select a.*,b.unit from tmp2 a left join Wh_lcdm0006 b
	on a.item_id=compress(b.jz_code||b.type_code||b.pro_code);
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


/*取主工序*/
proc sql;
	create table tmp4 as select a.* from tmp2 a,wh_lc.wh_lca_0002 b
	where a.jz_code=b.jz_code and b.type="&type";
quit;

/*统一单位:kg  m3   kWh */
data tmp5(drop=unit);
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
quit;


proc sql;
		create table tmp6 as select jz_name,jz_code,type_name,type_code,item_id,sum(val) as val from tmp5 group by jz_name,jz_code,type_name,type_code,item_id;
		create table temp2 as select sum(val) as val_all,jz_code
		from tmp6
		where type_code in ('04','05')
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



/*能源处理*/
proc sql;
create table lca0001 as select avg(co2_ratio) as co2_ratio,pro_code
from wh_lc.wh_lca_0001
where ymonth between "&datefrom" and "&dateto"
group by pro_code;
quit;
proc sql;
create table lca0009 as select avg(hot_value) as hot_value,pro_code
from wh_lc.wh_lca_0009
where ymonth between "&datefrom" and "&dateto"
group by pro_code;
quit;
proc sql;
create table temp5 as select a.jz_name,a.jz_code,a.dh_val*b.co2_ratio as dh_val,b.pro_code
from tmp7 a join lca0001 b
on substr(a.item_id,7,3)=b.pro_code;

create table temp51 as select a.jz_name,a.jz_code,a.dh_val*b.hot_value as dh_val,b.pro_code
from tmp7 a join lca0009 b
on substr(a.item_id,7,3)=b.pro_code;
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
create table lca0002 as select jz_code,jz_name,id from wh_lc.wh_lca_0002 where type="&type" order by id;
quit;

proc sql;
create table g1 as select a.jz_code,a.id,b.pro_code,b.dh_val
from lca0002 a left join temp6 b
on a.jz_code=b.jz_code
order by id;
quit;

/*环保报表数据*/

/*环保报表数据*/

proc sql;
	create table lca0003 as select pro_code,avg(var) as dh_val,jz_code
	from wh_lc.wh_lca_0003
	where  ymonth between "&datefrom" and "&dateto" and jz_code in (select distinct jz_code from wh_lc.wh_lca_0002 where type="&type")
	group by pro_code,jz_code;
quit;
proc sql;
	create table lca00031 as select a.*,b.id from lca0003 a left join wh_lc.wh_lca_0002 b
	on a.jz_code=b.jz_code;
quit;

data g11;
set g1 lca00031;
run;

proc sql;
		create table g11a as
		select jz_code,id,dh_val,pro_code from g11;
	quit;

proc sort data=g11a; by pro_code ;run;

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

/*G2：PE2（PA2（I-PA1)^）
PA1:本工序的原材料、辅料来自其他主工序的产品
*/


data lca0006;
	set wh_lc.wh_lca_0006;
	where type="&type" and pro_code^='TTT';
run;
proc sql;
create table lca0006_1 as select distinct type_code,pro_code from wh_lc.wh_lca_0006
	where type="&type"  and pro_code^='TTT';
quit;

proc sql;
	create table g2_tmp1 as select jz_code,id from wh_lc.wh_lca_0002 where type="&type"  order by id;
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

proc sort data=g2_tmp4; by pro_code ;run;

PROC TRANSPOSE data=g2_tmp4 out=g2_tmp6(drop=_name_)  prefix=col;
by pro_code ;
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
	where type='ny' and pro_code^='TTT';
run;
proc sql;
create table lca0006_ny_1 as select distinct type_code,pro_code from wh_lc.wh_lca_0006
	where type='ny'  and pro_code^='TTT';
quit;

proc sql;
	create table g2_tmp1a as select jz_code,id from wh_lc.wh_lca_0002 where type="&type"  order by id;
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
proc sql;
	create table g2_tmp4a as select a.jz_code,a.id,b.dh_val,b.pro_code,b.item from g2_tmp1a a left join g2_tmp3a b
	on a.jz_code=b.jz_code;
quit;

/*转置*/

proc sort data=g2_tmp4a; by pro_code ;run;

PROC TRANSPOSE data=g2_tmp4a out=g2_tmp6a(drop=_name_)  prefix=col;
by pro_code ;
var  dh_val;
id  id;
run;
/*固定表格式*/
proc sql;
	create table tmp_ny as select jz_code,id from wh_lc.wh_lca_0002 where type='ny'  and id<=17 order by id;
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

/*PE2：能源site（不包含发电工序）*/
data result_pe2;
set ma_lc.sa_lca_ny_gsite(drop=col18);
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
run;
proc sort data=result_pe2 ;by id;run;

data result_pe2;
set result_pe2;
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

/*G2*/

proc iml;
	reset print;
	use result_pa1;
	read all var _all_ into pa1;
	use result_pa2;
	read all var _all_ into pa2;
	use result_pe2;
	read all var _all_ into pe2;
	I=I(123);
	g2=pe2*(pa2*inv(I-pa1));
	
	create result_g2 from g2;
	append from g2;
quit;

/***G3:g1*(I-PA1)^-I）****/


proc iml;
	reset print;
	use result_pa1;
	read all var _all_ into pa1;
	use result_g1;
	read all var _all_ into g1;
	
	I=I(123);
	g2=inv(I-pa1)-I;
	g3=g1*g2;
	create result_g3 from g3;
	append from g3;
quit;

/*******G4:RG(RA*(I-PA1)^)
           RA:主工序用到的其他工序产出的副产品，该副产品不可以是任何工序的产品:目前按照中文名字，后续需要优化
		RG:维护表（UP表中读取）

*/


proc sql;
create table tmp71  as select a.*,b.pro_code_name  from tmp7 a left join ma_lc.wh_lcdm0009 b on a.type_code=b.type_code and 
	substr(a.item_id,7,3)=b.pro_code;
create table tmp_ra1 as select distinct substr(item_id,7,3) as pro_code,type_code,pro_code_name from tmp71 where type_code='05' 
	and pro_code_name not in (select distinct pro_code_name from tmp71 where  type_code='04');
create table tmp_ra2 as select jz_code,substr(item_id,7,3) as pro_code,type_code,dh_val,type_code,pro_code_name from tmp71 where pro_code_name in (select distinct pro_code_name from tmp_ra1) and type_code in ('01','02','03');
quit;

data WH_LCA_0014;
set wh_lc.WH_LCA_0014;
run;

proc sql;
create table lca_0014_1 as select distinct id,product from WH_LCA_0014;
/*create table tmp_ra3 as select a.*,b.id from tmp_ra2 a left join lca_0014_1 b on a.pro_code_name=b.product;*/
quit;
proc sql;
	create table g4_tmp1a as select jz_code,id from wh_lc.wh_lca_0002 where type="&type"  order by id;
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
	reset print;
	use result_rg;
	read all var _all_ into rg;
	use result_ra;
	read all var _all_ into ra;
	use result_pa1;
	read all var _all_ into pa1;
	I=I(123);
	t=ra*inv(I-pa1);
	g4=rg*t;
	
	create result_g4 from g4;
	append from g4;
quit;
/*****************G5:（T*L)*(OA(I-PA1)^1)+PF2(PA2(I-PA1)^1)*******************/
/*矩阵T*/
proc sql;
create table wh_lca_0010 as select avg(train) as train,avg(barge) as barge,avg(seacraft) as seacraft,avg(car) as car,pro_code
from wh_lc.wh_lca_0010
where ymonth between "&datefrom" and "&dateto"
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
		where ymonth between "&datefrom" and "&dateto"
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
	create table result_oa1 as select a.jz_code,a.dh_val,case when b.pro_code='coal' then compress(b.pro_code||b.item_code) else b.pro_code end as pro_code from tmp7 a join wh_lc.wh_lca_0007 b 
	on substr(a.item_id,7,3)=b.item_code and  a.type_code=b.type_code and a.type_code in ('01','02','03') and b.pro_code in ('ironstone','limestone','dolomite','coal');
quit;
proc sql;
	create table result_oa1 as select sum(dh_val) as dh_val,jz_code,pro_code from result_oa1 group by jz_code,pro_code;
quit;
proc sql;
create table result_oa2 as select a.dh_val,a.pro_code,b.id from result_oa1 a right join wh_lc.wh_lca_0002 b on a.jz_code=b.jz_code where b.type="&type";
quit;
/*转置*/

proc sort data=result_oa2; by   pro_code ;run;

PROC TRANSPOSE data=result_oa2 out=result_oa3(drop=_name_)  prefix=col;
by pro_code;
var  dh_val;
id  id;
run;

proc sql;
	create table result_oa4 as select a.*,b.id
	from result_oa3 a right join wh_lc.wh_lca_0011 b
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
set  ma_lc.sa_lca_ny_g5(drop=col18);
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;

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
	reset print;
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
	I=I(123);
	a=t*l;
	b=oa*inv(I-pa1);
	c=a*b;
	d=pa2*inv(I-pa1);
	e=pf2*d;
	g5=c+e;
	
	create result_g5 from g5;
	append from g5;
quit;


/*************************************G6:OG*(OA1(I-PA1)^1)+PG2(PA2(I-PA1)^1)************************************/
/*OA1:同OA，所有外购（不是其他工序的产品、副产品）*/
/*PG2:能源G6*/
/*OG:维护表（UP表中读取）*/

/*矩阵OA1:工序中消耗量，不算输出量，所有外购（不是其他工序的产品、副产品）*/

proc sql;
	create table result_oa11 as select a.jz_code,a.dh_val,substr(a.item_id,5,5) as pro_code,pro_code_name from tmp71 a 
where a.type_code in ('01','02','03') and a.pro_code_name not in (select distinct pro_code_name from tmp71 where type_code in ('04','05'));
quit;


proc sql;
create table result_oa12 as select distinct a.*,b.id as pro_code_new from result_oa11 a   join wh_lc.wh_lca_0014 b on a.pro_code_name=b.product ;
quit;
proc sql;
create table result_oa12 as select a.dh_val,a.pro_code_name,a.pro_code,a.pro_code_new,b.id from result_oa12 a right join wh_lc.wh_lca_0002 b on a.jz_code=b.jz_code where b.type="&type";
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
set ma_lc.sa_lca_ny_g6(drop=col18);
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
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
	reset print;
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
	
	I=I(123);
	a=inv(I-pa1);
	b=oa1*a;
	c=og*b;
	d=pa2*a;
	e=pg2*d;
	g6=c+e;
	
	create result_g6 from g6;
	append from g6;
quit;


/*****************G7:RG1*(RA1(I-PA1)^1)+PH2(PA2(I-PA1)^1)*************************/
/*
RA1:卖到外面的副产品（产出的副产品没有在任何工序中作为投入）
RG1:维护表（根据RA1中数据到UP表中读取）
PH2:能源G7
*/

/*RA1:*/

proc sql;
	create table result_ra11 as select a.jz_code,a.dh_val,substr(a.item_id,5,5) as pro_code,pro_code_name from tmp71 a 
where a.type_code in ('05') and a.pro_code_name not in (select distinct pro_code_name from tmp71 where type_code in ('01','02','03'));
quit;


proc sql;
create table result_ra12 as select distinct a.*,b.id as pro_code_new from result_ra11 a   join wh_lc.wh_lca_0014 b on a.pro_code_name=b.product ;
quit;
proc sql;
create table result_ra12 as select a.dh_val,a.pro_code_name,a.pro_code,a.pro_code_new,b.id from result_ra12 a right join wh_lc.wh_lca_0002 b on a.jz_code=b.jz_code where b.type="&type";
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
set ma_lc.sa_lca_ny_g7(drop=col18);
retain id col1 col2 col3 col4 col5 col6 col7 col8 col9 col10 col11 col12 col13 col14 col15 col16 col17  ;
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
	reset print;
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
	
	I=I(123);
	a=inv(I-pa1);
	b=ra1*a;
	c=rg1*b;
	d=pa2*a;
	e=ph2*d;
	g7=c+e;
	
	create result_g7 from g7;
	append from g7;
quit;



proc iml;
	reset print;
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
	groute=g1+g2+g3+g4+g5+g6+g7;
	
	create result_groute from groute;
	append from groute;
quit;

proc iml;
	reset print;
	use result_g1;
	read all var _all_ into g1;
	use result_g2;
	read all var _all_ into g2;
	use result_g3;
	read all var _all_ into g3;
	use result_g4;
	read all var _all_ into g4;
	
	gsite=g1+g2+g3+g4;
	
	create result_gsite from gsite;
	append from gsite;
quit;

proc iml;
	reset print;
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
create table result_alla as select "&datefrom" as LCItime_s,"&dateto" as LCItime_e,b.jz_code as jz_code1,b.jz_name as product,a.* from result_alla a left join wh_lc.wh_lca_0002 b
on input(substr(a.jz_code,4),2.)=b.id;
quit;

data result_alla;
set result_alla;
drop id jz_code;
rename jz_code1=jz_code;
rec_create_time=put(date(),yymmddn8.);
type="&type";
run;


proc sql;
delete from ma_lc.sa_lca_0001 where LCITIME_S="&datefrom" and LCItime_e="&dateto" and type="&type";
quit;
proc append base=ma_lc.sa_lca_0001 data=result_alla force;run;
proc datasets library=work kill;
quit;
