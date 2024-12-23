print 'sp_lis_interface_get_test_result_model'
go
/*  $Header$ */
/* $Log$
*/
if exists (select name from sysobjects where name='sp_lis_interface_get_test_result_model' and type='P ')
  drop proc sp_lis_interface_get_test_result_model
go
 
 
/**********************************************/
--	Author:				Tui
--	Date Created:		Sep 18, 2013
--  Last Modified Date:	Sep 10, 2015
--	Description:		Get test result code by lab_no and analyzer_cd
--	Module:				islim
--	Log:				Sep 10, 2015  Add IQC
/**********************************************/
 
create proc [dbo].[sp_lis_interface_get_test_result_model]
	@lab_no nvarchar(20),
	@analyzer_model nvarchar(50),
	@analyzer_skey int,
	@analyzer_date datetime = null,
	@flag_exist char(1),
	@flagScan nvarchar(1) = null
	 ,@Dilution_flag nvarchar(1) = null,
	 @ref_cd nvarchar(50) = null
as
 
/* from rsp_comments  at Sep 18 2013 11:25:19:460AM*/
 
set nocount on
--exec sp_lis_interface_get_test_result_model @lab_no=N'200000061',@analyzer_model=N'Echo',@analyzer_skey=51,@analyzer_date='2019-10-16 10:10:00',@flag_exist=N'0',@flagScan=N'N'
 
 --exec sp_lis_interface_get_test_result_model @lab_no=N'200169487',@analyzer_model=N'Architecti1000',@analyzer_skey=28,@analyzer_date='2019-10-16 10:10:00',@flag_exist=N'0',@flagScan=N'N'
 
declare @test_result table
(
	order_skey int null,
	result_item_skey int null,
	alias_id nvarchar(100) null,
	analyzer_skey int null,
	analyzer_ref_cd nvarchar(100) null,
	result_value nvarchar(50) null,
	rerun bit null,
	two_way_flag bit null,
	model_cd nvarchar(20),
	lot_skey int null,
	result_date datetime null,
	order_id nvarchar(20) null,
	analyzer_cd nvarchar(50) null,
	sample_type nvarchar(2) null ,
	test_item_skey int null,
	date_scanned datetime null
)
 
declare @lab_no_old nvarchar(20)
set @lab_no_old = @lab_no
declare @two_way_flag bit
set @two_way_flag = 0
 
--Check Rerun
declare @rerun bit = 0
 
 --exec sp_lis_interface_get_test_result_model @lab_no=N'cal',@analyzer_model=N'RxLyte',@analyzer_skey=4,@analyzer_date='2019-10-16 10:10:00',@flag_exist=N'0',@flagScan=N'N'
/*Tui Add IQC , 2015-09-10*/
--???????????????????????? IQC ???????? ????????????? Specimen, Analyzer, Comport ??? ????????????? Lot ????
 
 
 
 
 
 
declare @lot_skey int
 
select	@lot_skey = lot_skey from lis_analyzer 
		inner join lis_analyzer_model on lis_analyzer.model_skey = lis_analyzer_model.model_skey
		inner join lis_iqc_lot_header on lis_analyzer.analyzer_skey = lis_iqc_lot_header.analyzer_skey 
where	lis_iqc_lot_header.ref_lot_no = @lab_no 
		and lis_analyzer.analyzer_skey = @analyzer_skey
		and @analyzer_date between lis_iqc_lot_header.start_date and lis_iqc_lot_header.expired_date	
 
if isnull(@lot_skey,0) <> 0
begin
	insert	@test_result (order_skey, result_item_skey, alias_id, analyzer_skey, analyzer_ref_cd, result_value ,rerun, two_way_flag, lot_skey,analyzer_cd)
	select	0 as order_skey,
			lis_iqc_lot_result_item.result_item_skey,
			lis_result_item.alias_id,
			lis_analyzer.analyzer_skey,
			lis_analyzer_model_result_item.ref_cd,	
			null as result_value,			
			@rerun,
			@two_way_flag,
			@lot_skey,
			lis_analyzer.analyzer_cd
	from	lis_iqc_lot_header
			inner join lis_analyzer on lis_iqc_lot_header.analyzer_skey = lis_analyzer.analyzer_skey
			inner join lis_analyzer_model on lis_analyzer.model_skey = lis_analyzer_model.model_skey
			inner join lis_analyzer_model_result_item on lis_analyzer_model.model_skey = lis_analyzer_model_result_item.model_skey
			inner join lis_iqc_lot_result_item on lis_iqc_lot_header.lot_skey = lis_iqc_lot_result_item.lot_skey and lis_iqc_lot_result_item.result_item_skey = lis_analyzer_model_result_item.result_item_skey
			inner join lis_result_item on lis_analyzer_model_result_item.result_item_skey = lis_result_item.result_item_skey
	where	lis_iqc_lot_header.lot_skey = @lot_skey
			and lis_analyzer.analyzer_skey = @analyzer_skey
 
	select 	order_skey, 
			result_item_skey, 
			analyzer_skey,
			analyzer_ref_cd, 
			alias_id,
			convert(nvarchar(50),null) as result_value,
			@lab_no as specimen_type_id,
			@analyzer_model as analyzer,
			'Pending' as sending_status,
			lot_skey,
			result_date,
			order_id,
			analyzer_cd,
			rerun,
			isnull(sample_type,'5') as sample_type
	From	@test_result
 
	return
end
 
 
 
if @analyzer_model in ('COBAS','Imola','Daytona','Liaison','DX300')
begin
	set @two_way_flag = 1
end
 
declare @LIS_ANALYZER_FIND_BY_HLN nvarchar(1)
set @LIS_ANALYZER_FIND_BY_HLN = 'N'
if exists (select 1 from global_parms where global_key like '%LIS_ANALYZER_FIND_BY_HLN%' and isnull(parm_value,'') = 'Y')
begin 
	select @LIS_ANALYZER_FIND_BY_HLN=parm_value from global_parms where global_key like '%LIS_ANALYZER_FIND_BY_HLN%'
	goto startFindByLabNumber
end
 
 
 
 
if @analyzer_model = 'RxLyte'
begin 
	if exists (select 1 from global_parms where global_key like '%LIS_SID_FORMAT_PURE_RUNNING_NUMBER%' and isnull(parm_value,'') = 'Y')
	begin 
		if isnumeric(@lab_no) = 0 
		begin
			goto endRxlyte
		end
		declare   @b nvarchar(255) , @elyteDigit int
		 
		select @elyteDigit=parm_value from global_parms where global_key like '%Elyte_Set_Digit_Specimen_ID%'
		 
		if @elyteDigit = null or @elyteDigit is null
		begin
			set @elyteDigit = 0
		end
 
		if @elyteDigit > 0 
		begin
			select @lab_no = RIGHT ( @lab_no, @elyteDigit )
			goto endRxlyte
		end
 
		select @b=substring(@lab_no,PATINDEX('%[1-9]%',@lab_no),len(@lab_no))
		if len(@b)>= 3
		begin
			select @lab_no= @b
		end 
		else if len(@b) = 2
		begin
			select @lab_no = '0' + @b
		end
		else
		begin
			select @lab_no = '00' + @b
		end 
 
		goto endRxlyte
	end
	if len(@lab_no) >= 9 
	begin 
		set @lab_no = convert(numeric,@lab_no) 
		if len(@lab_no) < 9 -- Golf edit 2016-09-09
		begin
			select @lab_no = convert(decimal,Right(Year(getDate()),2) + '0000000') + convert(decimal,@lab_no)
		end
	end
	else 
	begin
		select @lab_no = convert(decimal,Right(Year(getDate()),2) + '0000000') + convert(decimal,@lab_no)
	end
	set @lab_no = convert(numeric,@lab_no) 
	 
 
	 if not exists(
			 select 1
				  From lis_lab_specimen_type
					inner join lis_lab_order
						on lis_lab_specimen_type.order_skey = lis_lab_order.order_skey
					inner join lis_lab_specimen_type_test_item
						on lis_lab_specimen_type.specimen_type_id = lis_lab_specimen_type_test_item.specimen_type_id
						and lis_lab_specimen_type.order_skey = lis_lab_specimen_type_test_item.order_skey
					inner join lis_lab_test_result_item
						on lis_lab_specimen_type_test_item.test_item_skey = lis_lab_test_result_item.test_item_skey
						and lis_lab_specimen_type_test_item.order_skey = lis_lab_test_result_item.order_skey
					inner join lis_lab_result_item
						on lis_lab_test_result_item.result_item_skey = lis_lab_result_item.result_item_skey
						and lis_lab_test_result_item.order_skey = lis_lab_result_item.order_skey
					inner join lis_result_item
						on lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey
					inner join lis_analyzer_model_result_item
						on lis_result_item.result_item_skey = lis_analyzer_model_result_item.result_item_skey
					inner join lis_analyzer_model
						on lis_analyzer_model_result_item.model_skey = lis_analyzer_model.model_skey
					inner join lis_analyzer
						on lis_analyzer_model.model_skey = lis_analyzer.model_skey 
						 inner join lis_lab_order_item bbb  
						 on lis_lab_order.order_skey = bbb.order_skey 
						 inner join lis_lab_order_test_item c 
											 on bbb.order_skey = c.order_skey 
											 and bbb.order_item_skey = c.order_item_skey 
						 inner join lis_lab_test_item g	
											on lis_lab_order.order_skey = g.order_skey and 
												c.test_item_skey = g.test_item_skey	 and
												lis_lab_test_result_item.test_item_skey = g.test_item_skey --Golf 2017-01-17 
				where	lis_lab_specimen_type.specimen_type_id = @lab_no  and cancel <> 1 and g.status <>  'CA'  and lis_lab_order.status <> 'CA' -- Golf 2018-11-30
						and lis_analyzer.analyzer_skey = @analyzer_skey
				)
	begin
		 set @lab_no_old = '%' + rtrim(@lab_no_old)
		 set rowcount 1
		 select @lab_no = lis_lab_specimen_type.specimen_type_id
				  From lis_lab_specimen_type
					inner join lis_lab_order
						on lis_lab_specimen_type.order_skey = lis_lab_order.order_skey
					inner join lis_lab_specimen_type_test_item
						on lis_lab_specimen_type.specimen_type_id = lis_lab_specimen_type_test_item.specimen_type_id
						and lis_lab_specimen_type.order_skey = lis_lab_specimen_type_test_item.order_skey
					inner join lis_lab_test_result_item
						on lis_lab_specimen_type_test_item.test_item_skey = lis_lab_test_result_item.test_item_skey
						and lis_lab_specimen_type_test_item.order_skey = lis_lab_test_result_item.order_skey
					inner join lis_lab_result_item
						on lis_lab_test_result_item.result_item_skey = lis_lab_result_item.result_item_skey
						and lis_lab_test_result_item.order_skey = lis_lab_result_item.order_skey
					inner join lis_result_item
						on lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey
					inner join lis_analyzer_model_result_item
						on lis_result_item.result_item_skey = lis_analyzer_model_result_item.result_item_skey
					inner join lis_analyzer_model
						on lis_analyzer_model_result_item.model_skey = lis_analyzer_model.model_skey
					inner join lis_analyzer
						on lis_analyzer_model.model_skey = lis_analyzer.model_skey 
						 inner join lis_lab_order_item bbb  
						 on lis_lab_order.order_skey = bbb.order_skey 
						 inner join lis_lab_order_test_item c 
											 on bbb.order_skey = c.order_skey 
											 and bbb.order_item_skey = c.order_item_skey 
						 inner join lis_lab_test_item g	
											on lis_lab_order.order_skey = g.order_skey and 
												c.test_item_skey = g.test_item_skey	 and
												lis_lab_test_result_item.test_item_skey = g.test_item_skey --Golf 2017-01-17 
				where	lis_lab_specimen_type.specimen_type_id like @lab_no_old and cancel <> 1 and g.status <>  'CA'  and lis_lab_order.status <> 'CA' -- Golf 2018-11-30
						and lis_analyzer.analyzer_skey = @analyzer_skey 
						and convert(nvarchar(6),lis_lab_order.date_created,112)  = convert(nvarchar(6),@analyzer_date,112) 
		set rowcount 0
						 
	end
				 
 
end 
 
 
endRxlyte:
startFindByLabNumber: 
--exec sp_lis_interface_get_test_result_model @lab_no=N'19101600002',@analyzer_model=N'RxLyte',@analyzer_skey=4,@analyzer_date='2019-10-16 10:10:00',@flag_exist=N'0',@flagScan=N'N'
 
 --select @lab_no
--select @lab_no
if @LIS_ANALYZER_FIND_BY_HLN = 'N'
begin
	insert @test_result (order_skey,test_item_skey, result_item_skey, alias_id, analyzer_skey, analyzer_ref_cd, result_value ,rerun, two_way_flag, order_id, analyzer_cd,sample_type,date_scanned)
	select distinct lis_lab_specimen_type.order_skey,
			lis_lab_specimen_type_test_item.test_item_skey,
			lis_result_item.result_item_skey,
			lis_result_item.alias_id,
			lis_analyzer.analyzer_skey,
			lis_analyzer_model_result_item.ref_cd, 
			lis_lab_result_item.result_value, 
			lis_lab_result_item.rerun,
			@two_way_flag,
			lis_lab_order.order_id,
			lis_analyzer.analyzer_cd, 
			
			/*  Ploy comment 2021.03.05 */
			/*
			(case lis_specimen_type.specimen_type_desc 
			when 'Random urine' then 2
			when 'Urine 24 hrs.' then 2
			when 'Urine' then 2
			when 'ACD Blood' then 10
			when 'EDTA blood' then 10
			else 1
			end) sample_type
			*//* END Ploy comment 2021.03.05 */
			/* Ploy Add 2021.03.05 */
			(case @analyzer_model 
				When 'DxC700AU' Then 
					case lis_specimen_type.specimen_type_desc  
						When 'Random urine' Then 'U'
						When 'Urine 24 hrs.' Then 'U'
						When 'CSF' Then 'U'
						
						When 'EDTA blood' Then 'W'
						When 'EDTA plasma' Then 'W'
						When 'EDTA sterile blood' Then 'W'
						When 'EDTA sterile plasma'  Then 'W'
						Else ' '
					END
				Else 
					Case lis_specimen_type.specimen_type_desc 
						when 'Random urine' then '2'
						when 'Urine 24 hrs.' then '2'
						when 'Urine' then '2'
						when 'ACD Blood' then '10'
						when 'EDTA blood' then '10'
						ELSE '1'
					END
				END) As sample_type
				/* END Ploy Add 2021.03.05 */
			,lis_lab_specimen_type_test_item.date_scanned
	From lis_lab_specimen_type
		inner join lis_lab_order
			on lis_lab_specimen_type.order_skey = lis_lab_order.order_skey
		inner join lis_lab_specimen_type_test_item
			on lis_lab_specimen_type.specimen_type_id = lis_lab_specimen_type_test_item.specimen_type_id
			and lis_lab_specimen_type.order_skey = lis_lab_specimen_type_test_item.order_skey
		inner join lis_lab_test_result_item
			on lis_lab_specimen_type_test_item.test_item_skey = lis_lab_test_result_item.test_item_skey
			and lis_lab_specimen_type_test_item.order_skey = lis_lab_test_result_item.order_skey
		inner join lis_lab_result_item
			on lis_lab_test_result_item.result_item_skey = lis_lab_result_item.result_item_skey
			and lis_lab_test_result_item.order_skey = lis_lab_result_item.order_skey
		inner join lis_result_item
			on lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey
		inner join lis_analyzer_model_result_item
			on lis_result_item.result_item_skey = lis_analyzer_model_result_item.result_item_skey
		inner join lis_analyzer_model
			on lis_analyzer_model_result_item.model_skey = lis_analyzer_model.model_skey
		inner join lis_analyzer
			on lis_analyzer_model.model_skey = lis_analyzer.model_skey
		--Golf 2017-01-17
			 inner join lis_lab_order_item bbb  
			 on lis_lab_order.order_skey = bbb.order_skey 
			 inner join lis_lab_order_test_item c 
								 on bbb.order_skey = c.order_skey 
								 and bbb.order_item_skey = c.order_item_skey 
			 inner join lis_lab_test_item g	
								on lis_lab_order.order_skey = g.order_skey and 
									c.test_item_skey = g.test_item_skey	 and
									lis_lab_test_result_item.test_item_skey = g.test_item_skey --Golf 2017-01-17 
			--Golf 2017-01-17	
		inner join lis_specimen_type
			on  lis_lab_specimen_type.specimen_type_cd = lis_specimen_type.specimen_type_cd
	where	lis_lab_specimen_type.specimen_type_id = @lab_no  and cancel <> 1 and g.status <>  'CA'  and lis_lab_order.status <> 'CA' -- Golf 2018-11-30
			and lis_analyzer.analyzer_skey = @analyzer_skey
			 
			 
end
else if @LIS_ANALYZER_FIND_BY_HLN = 'Y'
begin
	if @flagScan = 'Y' 
	begin

		insert @test_result (order_skey, test_item_skey,result_item_skey, alias_id, analyzer_skey, analyzer_ref_cd, result_value ,rerun, two_way_flag, order_id, analyzer_cd,sample_type,date_scanned)
		select distinct lis_lab_specimen_type.order_skey,
		lis_lab_specimen_type_test_item.test_item_skey,
				lis_result_item.result_item_skey,
				lis_result_item.alias_id,
				lis_analyzer.analyzer_skey,
				lis_analyzer_model_result_item.ref_cd, 
				lis_lab_result_item.result_value, 
				lis_lab_result_item.rerun,
				@two_way_flag,
				lis_lab_order.order_id,
				lis_analyzer.analyzer_cd, 
				/*  Ploy comment 2021.03.05 */
			/*
			(case lis_specimen_type.specimen_type_desc 
			when 'Random urine' then 2
			when 'Urine 24 hrs.' then 2
			when 'Urine' then 2
			when 'ACD Blood' then 10
			when 'EDTA blood' then 10
			else 1
			end) sample_type
			*//* END Ploy comment 2021.03.05 */
			/* Ploy Add 2021.03.05 */
			(case @analyzer_model 
				When 'DxC700AU' Then 
					case lis_specimen_type.specimen_type_desc  
						When 'Random urine' Then 'U'
						When 'Urine 24 hrs.' Then 'U'
						When 'CSF' Then 'U'
						
						When 'EDTA blood' Then 'W'
						When 'EDTA plasma' Then 'W'
						When 'EDTA sterile blood' Then 'W'
						When 'EDTA sterile plasma'  Then 'W'
						Else ' '
					END
				Else 
					Case lis_specimen_type.specimen_type_desc 
						when 'Random urine' then '2'
						when 'Urine 24 hrs.' then '2'
						when 'Urine' then '2'
						when 'ACD Blood' then '10'
						when 'EDTA blood' then '10'
						ELSE '1'
					END
				END) As sample_type
				/* END Ploy Add 2021.03.05 */
				,lis_lab_specimen_type_test_item.date_scanned
		From lis_lab_specimen_type
			inner join lis_lab_order
				on lis_lab_specimen_type.order_skey = lis_lab_order.order_skey
			inner join lis_lab_specimen_type_test_item
				on lis_lab_specimen_type.specimen_type_id = lis_lab_specimen_type_test_item.specimen_type_id
				and lis_lab_specimen_type.order_skey = lis_lab_specimen_type_test_item.order_skey
			inner join lis_lab_test_result_item
				on lis_lab_specimen_type_test_item.test_item_skey = lis_lab_test_result_item.test_item_skey
				and lis_lab_specimen_type_test_item.order_skey = lis_lab_test_result_item.order_skey
			inner join lis_lab_result_item
				on lis_lab_test_result_item.result_item_skey = lis_lab_result_item.result_item_skey
				and lis_lab_test_result_item.order_skey = lis_lab_result_item.order_skey
			inner join lis_result_item
				on lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey
			inner join lis_analyzer_model_result_item
				on lis_result_item.result_item_skey = lis_analyzer_model_result_item.result_item_skey
			inner join lis_analyzer_model
				on lis_analyzer_model_result_item.model_skey = lis_analyzer_model.model_skey
			inner join lis_analyzer
				on lis_analyzer_model.model_skey = lis_analyzer.model_skey
			--Golf 2017-01-17
				 inner join lis_lab_order_item bbb  
				 on lis_lab_order.order_skey = bbb.order_skey 
				 inner join lis_lab_order_test_item c 
									 on bbb.order_skey = c.order_skey 
									 and bbb.order_item_skey = c.order_item_skey 
				 inner join lis_lab_test_item g	
									on lis_lab_order.order_skey = g.order_skey and 
										c.test_item_skey = g.test_item_skey	 and
										lis_lab_test_result_item.test_item_skey = g.test_item_skey --Golf 2017-01-17 
				--Golf 2017-01-17	
			inner join lis_specimen_type
				on  lis_lab_specimen_type.specimen_type_cd = lis_specimen_type.specimen_type_cd
		where	lis_lab_order.hospital_lab_no = @lab_no  and cancel <> 1 and g.status <>  'CA'  and lis_lab_order.status <> 'CA' -- Golf 2018-11-30
				and lis_analyzer.analyzer_skey = @analyzer_skey
				and lis_result_item.result_item_id not in ('020004302','02001408','0300010','0300038','0300046','01502601')
	end
	else
	begin
	
		insert @test_result (order_skey, test_item_skey,result_item_skey, alias_id, analyzer_skey, analyzer_ref_cd, result_value ,rerun, two_way_flag, order_id, analyzer_cd,sample_type,date_scanned)
		select distinct lis_lab_specimen_type.order_skey,
		lis_lab_specimen_type_test_item.test_item_skey,
				lis_result_item.result_item_skey,
				lis_result_item.alias_id,
				lis_analyzer.analyzer_skey,
				lis_analyzer_model_result_item.ref_cd, 
				lis_lab_result_item.result_value, 
				lis_lab_result_item.rerun,
				@two_way_flag,
				lis_lab_order.order_id,
				lis_analyzer.analyzer_cd, 
				/*  Ploy comment 2021.03.05 */
			/*
			(case lis_specimen_type.specimen_type_desc 
			when 'Random urine' then 2
			when 'Urine 24 hrs.' then 2
			when 'Urine' then 2
			when 'ACD Blood' then 10
			when 'EDTA blood' then 10
			else 1
			end) sample_type
			*//* END Ploy comment 2021.03.05 */
			/* Ploy Add 2021.03.05 */
			(case @analyzer_model 
				When 'DxC700AU' Then 
					case lis_specimen_type.specimen_type_desc  
						When 'Random urine' Then 'U'
						When 'Urine 24 hrs.' Then 'U'
						When 'CSF' Then 'U'
						
						When 'EDTA blood' Then 'W'
						When 'EDTA plasma' Then 'W'
						When 'EDTA sterile blood' Then 'W'
						When 'EDTA sterile plasma'  Then 'W'
						Else ' '
					END
				Else 
					Case lis_specimen_type.specimen_type_desc 
						when 'Random urine' then '2'
						when 'Urine 24 hrs.' then '2'
						when 'Urine' then '2'
						when 'ACD Blood' then '10'
						when 'EDTA blood' then '10'
						ELSE '1'
					END
				END) As sample_type
				/* END Ploy Add 2021.03.05 */
				,lis_lab_specimen_type_test_item.date_scanned
		From lis_lab_specimen_type
			inner join lis_lab_order
				on lis_lab_specimen_type.order_skey = lis_lab_order.order_skey
			inner join lis_lab_specimen_type_test_item
				on lis_lab_specimen_type.specimen_type_id = lis_lab_specimen_type_test_item.specimen_type_id
				and lis_lab_specimen_type.order_skey = lis_lab_specimen_type_test_item.order_skey
			inner join lis_lab_test_result_item
				on lis_lab_specimen_type_test_item.test_item_skey = lis_lab_test_result_item.test_item_skey
				and lis_lab_specimen_type_test_item.order_skey = lis_lab_test_result_item.order_skey
			inner join lis_lab_result_item
				on lis_lab_test_result_item.result_item_skey = lis_lab_result_item.result_item_skey
				and lis_lab_test_result_item.order_skey = lis_lab_result_item.order_skey
			inner join lis_result_item
				on lis_lab_result_item.result_item_skey = lis_result_item.result_item_skey
			inner join lis_analyzer_model_result_item
				on lis_result_item.result_item_skey = lis_analyzer_model_result_item.result_item_skey
			inner join lis_analyzer_model
				on lis_analyzer_model_result_item.model_skey = lis_analyzer_model.model_skey
			inner join lis_analyzer
				on lis_analyzer_model.model_skey = lis_analyzer.model_skey
			--Golf 2017-01-17
				 inner join lis_lab_order_item bbb  
				 on lis_lab_order.order_skey = bbb.order_skey 
				 inner join lis_lab_order_test_item c 
									 on bbb.order_skey = c.order_skey 
									 and bbb.order_item_skey = c.order_item_skey 
				 inner join lis_lab_test_item g	
									on lis_lab_order.order_skey = g.order_skey and 
										c.test_item_skey = g.test_item_skey	 and
										lis_lab_test_result_item.test_item_skey = g.test_item_skey --Golf 2017-01-17 
				--Golf 2017-01-17	
			inner join lis_specimen_type
				on  lis_lab_specimen_type.specimen_type_cd = lis_specimen_type.specimen_type_cd
		where	lis_lab_order.hospital_lab_no = @lab_no  and cancel <> 1 and g.status <>  'CA'  and lis_lab_order.status <> 'CA' -- Golf 2018-11-30
				and lis_analyzer.analyzer_skey = @analyzer_skey 
				--select * from @test_result
	end
--2545	020004302
--1578	02001408
--839	0300010
--867	0300038
--875	0300046
 
end	  
if exists(select 1 from @test_result where rerun = 1) and @analyzer_model <> 'RXMODENA' and @analyzer_model <> 'Modena' and @analyzer_model <> 'COBAS_C311' and @analyzer_model <> 'Dimension' and @analyzer_model <> 'PATHFAST'
begin
	select @rerun = 1
end 
  
if @Dilution_flag is null or @Dilution_flag = null or @Dilution_flag = ''
begin
	set @Dilution_flag = 'N'
end
 
 
 
if @analyzer_model = 'RXMODENA' or @analyzer_model = 'Dimension' or @analyzer_model = 'PATHFAST'
begin
 
	--select rerun,@flag_exist,result_value,two_way_flag,* from @test_result
	select 	order_skey, 
			result_item_skey, 
			analyzer_skey,
			analyzer_ref_cd, 
			alias_id,
			convert(nvarchar(50),null) as result_value,
			@lab_no as specimen_type_id,
			@analyzer_model as analyzer,
			'Pending' as sending_status,
			lot_skey,
			result_date,
			order_id,
			analyzer_cd,
			rerun,
			isnull(sample_type,'1') as sample_type
	From	@test_result
	where	  (((result_value is null or len(rtrim(ltrim(convert(nvarchar(50),result_value)))) = 0) ) 
			 or rerun = 1)
			 or  CHARINDEX('<' ,result_value) > 0 or CHARINDEX('>' ,result_value) > 0
			 
 
end
else if @analyzer_model = 'COBAS_C311'  
begin
	 
	if @Dilution_flag  = 'Y' and @analyzer_model = 'COBAS_C311'
	begin 
		select 	order_skey, 
				result_item_skey, 
				analyzer_skey,
				analyzer_ref_cd, 
				alias_id,
				convert(nvarchar(50),null) as result_value,
				@lab_no as specimen_type_id,
				@analyzer_model as analyzer,
				'Pending' as sending_status,
				lot_skey,
				result_date,
				order_id,
				analyzer_cd,
				1 rerun,
				isnull(sample_type,'1') as sample_type
		From	@test_result
		where analyzer_ref_cd = @ref_cd
	goto theEnd
	end
	
	select 	order_skey, 
			result_item_skey, 
			analyzer_skey,
			analyzer_ref_cd, 
			alias_id,
			convert(nvarchar(50),null) as result_value,
			@lab_no as specimen_type_id,
			@analyzer_model as analyzer,
			'Pending' as sending_status,
			lot_skey,
			result_date,
			order_id,
			analyzer_cd,
			rerun,
			isnull(sample_type,'1') as sample_type
	From	@test_result
	where	  (((result_value is null or len(rtrim(ltrim(convert(nvarchar(50),result_value)))) = 0) ) 
			 or rerun = 1)
			 or  CHARINDEX('<' ,result_value) > 0 or CHARINDEX('>' ,result_value) > 0
end 
else if @analyzer_model = 'Liaison' or @analyzer_model = 'H-500'  or @analyzer_model = 'RxLyte' or @analyzer_model = 'AS720' or @analyzer_model = 'AS720-02' or @analyzer_model = 'AS720-03' or @analyzer_model = 'AS720-04' or @analyzer_model = 'AS720-05' or @analyzer_model = 'ABXPentraXL80' or @analyzer_model = 'IQ200'
or @analyzer_model = 'HA8180V' -- Ploy Add 2020.11.23
or @analyzer_model = 'Stago' -- Ploy Add 2020.12.01
begin 
	select 	order_skey, 
			result_item_skey, 
			analyzer_skey,
			analyzer_ref_cd, 
			alias_id,
			convert(nvarchar(50),null) as result_value,
			@lab_no as specimen_type_id,
			@analyzer_model as analyzer,
			'Pending' as sending_status,
			lot_skey,
			result_date,
			order_id,
			analyzer_cd,
			rerun,
			isnull(sample_type,'1') as sample_type
	From	@test_result
	where	(rerun = @rerun or @flag_exist = 0)
			and (((result_value is null or len(rtrim(ltrim(convert(nvarchar(50),result_value)))) = 0) or @flag_exist = 0) or two_way_flag = 0 or rerun = 1)
 
end
else if @analyzer_model = 'Imola' and ( @Dilution_flag = 'R' or  @Dilution_flag = 'Y' )
begin
 select 	order_skey, 
			result_item_skey, 
			analyzer_skey,
			analyzer_ref_cd, 
			alias_id,
			convert(nvarchar(50),null) as result_value,
			@lab_no as specimen_type_id,
			@analyzer_model as analyzer,
			'Pending' as sending_status,
			lot_skey,
			result_date,
			order_id,
			analyzer_cd,
			rerun,
			isnull(sample_type,'1') as sample_type
	From	@test_result
	where	(rerun = @rerun or @flag_exist = 0)
			and (((result_value is null or len(rtrim(ltrim(convert(nvarchar(50),result_value)))) = 0) or @flag_exist = 0) or two_way_flag = 0 or rerun = 1)
 
  
end
else if @analyzer_model = 'Echo' --Golf add 2020-12-09
begin  
	
	--if exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-A') and  exists (select 1 from @test_result where analyzer_ref_cd  = 'Screen 1')
	--begin
	--	update @test_result set sample_type = 10
	--end
	--else if exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-D') and not exists (select 1 from @test_result where analyzer_ref_cd  = 'Screen 1')
	--begin
	--	update @test_result set sample_type = 10
	--end
	--else if exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-A') and not exists (select 1 from @test_result where analyzer_ref_cd  = 'Screen 1')
	--begin
	--	update @test_result set sample_type = 13
	--end
	--else if not exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-A') and exists (select 1 from @test_result where analyzer_ref_cd  = 'Screen 1')
	--begin
	--	update @test_result set sample_type = 14
	--end
	--else if exists (select 1 from @test_result where analyzer_ref_cd  = 'IgG XM') or exists (select 1 from @test_result where analyzer_ref_cd  = 'Ind Ctrl')
	--begin
	--	update @test_result set sample_type = 11
	--end
	if exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-A') and  exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-D')
	begin
		update @test_result set sample_type = 15 --Assay code - ABORH
	end
	else if exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-A') and not exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-D')
	begin
		update @test_result set sample_type = 16 --Assay code - ABO
	end
	else if not exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-A') and exists (select 1 from @test_result where analyzer_ref_cd  = 'Screen 1')
	begin
		update @test_result set sample_type = 17 --Assay code - Ab screening
	end
	else if not exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-A') and exists (select 1 from @test_result where analyzer_ref_cd  = 'Anti-D')
	begin
		update @test_result set sample_type = 18 --Assay code - Weak D
	end
	else if exists (select 1 from @test_result where analyzer_ref_cd  = 'IgG XM') 
	begin
		update @test_result set sample_type = 11 --Assay code - Crossmatch
	end
	else  
	begin
		update @test_result set sample_type = 10 --Assay code - Crossmatch
	end
	select order_skey, 
			result_item_skey, 
			analyzer_skey,
			analyzer_ref_cd, 
			alias_id,
			convert(nvarchar(50),null) as result_value,
			@lab_no as specimen_type_id,
			@analyzer_model as analyzer,
			'Pending' as sending_status,
			lot_skey,
			result_date,
			order_id,
			analyzer_cd,
			rerun,
			isnull(sample_type,'1') as sample_type
	From	@test_result
	where	  (((result_value is null or len(rtrim(ltrim(convert(nvarchar(50),result_value)))) = 0) ) 
			 )
			 or  CHARINDEX('<' ,result_value) > 0 or CHARINDEX('>' ,result_value) > 0
end
else if @analyzer_model = 'C4000' or @analyzer_model = 'C8000' or @analyzer_model = 'Architecti1000' or @analyzer_model = 'Architecti2000'--Golf add 2020-12-09
begin  
	
	
	select date_scanned,order_skey, 
			result_item_skey, 
			analyzer_skey,
			analyzer_ref_cd, 
			alias_id,
			convert(nvarchar(50),null) as result_value,
			@lab_no as specimen_type_id,
			@analyzer_model as analyzer,
			'Pending' as sending_status,
			lot_skey,
			result_date,
			order_id,
			analyzer_cd,
			rerun,
			isnull(sample_type,'1') as sample_type
	From	@test_result b
	where	 ( (((result_value is null or len(rtrim(ltrim(convert(nvarchar(50),result_value)))) = 0) ) 
			 )
			 or  CHARINDEX('<' ,result_value) > 0 or CHARINDEX('>' ,result_value) > 0)
			 and   ((isnull(date_scanned,'') = '' or date_scanned = '1900-01-01') or @flagScan = 'N' or (rerun = 1 and isnull(result_value,'') = '')
			 --or not exists (select 1 from lis_lab_result_value a
				--			inner join lis_lab_result_item c on a.order_skey = c.order_skey and a.result_item_skey = b.result_item_skey
				--			where  a.order_skey = b.order_skey and a.result_item_skey = b.result_item_skey and len(a.result_value) > 0 and isnull(a.result_value,'') = '' )
			 ) 
end
else
begin
	select 	order_skey, 
			result_item_skey, 
			analyzer_skey,
			analyzer_ref_cd, 
			alias_id,
			convert(nvarchar(50),null) as result_value,
			@lab_no as specimen_type_id,
			@analyzer_model as analyzer,
			'Pending' as sending_status,
			lot_skey,
			result_date,
			order_id,
			analyzer_cd,
			rerun,
			isnull(sample_type,'1') as sample_type
	From	@test_result
	where	  (((result_value is null or len(rtrim(ltrim(convert(nvarchar(50),result_value)))) = 0) ) 
			 --or rerun = 1 -- Golf edit comment 2017-08-02
			 )
			 or  CHARINDEX('<' ,result_value) > 0 or CHARINDEX('>' ,result_value) > 0
	
 
end 
 
theEnd:
 
/*Golf add 2021-02-10*/
if @flagScan = 'Y' 
begin
	if @LIS_ANALYZER_FIND_BY_HLN = 'N' 
	begin
		if exists (select 1 from lis_lab_specimen_type_test_item where (date_scanned is null or date_scanned = null) and specimen_type_id = @lab_no)
		begin
			--update lis_lab_specimen_type_test_item set date_scanned = getdate()
			--where (date_scanned is null or date_scanned = null) and specimen_type_id = @lab_no
			 
			 /*Golf edited 2021-02-16 ปรับ protocol เครื่อง model ต่อไปนี้ C4000, C8000, Architecti1000, Architecti2000 และเปลี่ยนการ update date_scanใหม่ */
			update lis_lab_specimen_type_test_item set date_scanned = getdate()
			from lis_lab_specimen_type_test_item a 
					inner join lis_lab_order b on a.order_skey = b.order_Skey
					inner join @test_result c on a.order_skey = c.order_skey and a.test_item_skey = c.test_item_skey  
					where (a.date_scanned is null or a.date_scanned = null) and a.specimen_type_id = @lab_no
			/*Golf edited 2021-02-16 */
		end
	end
	else if @LIS_ANALYZER_FIND_BY_HLN = 'Y' 
	begin
		if exists (select 1 from lis_lab_specimen_type_test_item a 
					inner join lis_lab_order b on a.order_skey = b.order_Skey
					inner join @test_result c on a.order_skey = c.order_skey and a.test_item_skey = c.test_item_skey
					where (a.date_scanned is null or a.date_scanned = null) and b.hospital_lab_no = @lab_no)
		begin
			update lis_lab_specimen_type_test_item set date_scanned = getdate()
			from lis_lab_specimen_type_test_item a 
					inner join lis_lab_order b on a.order_skey = b.order_Skey
					inner join @test_result c on a.order_skey = c.order_skey and a.test_item_skey = c.test_item_skey
					where (a.date_scanned is null or a.date_scanned = null) and b.hospital_lab_no = @lab_no
		end
	end
end
/*Golf add 2021-02-10*/ 
 
 
 
/*
exec sp_lis_interface_get_test_result_model @lab_no=N'1013376',@analyzer_model=N'COBAS_C311',@analyzer_skey=19,@analyzer_date='2017-06-29 17:14:09.243',@flag_exist=N'0'
 
exec sp_lis_interface_get_test_result_model @lab_no=N'170001258',@analyzer_model=N'COBAS_C311',@analyzer_skey=4,@analyzer_date='2017-06-29 17:14:09.243',@flag_exist=N'0'
*/
 
 
 
 
 
 
 
go
if exists (select name from sysobjects where name='sp_lis_interface_get_test_result_model' and type='P ')
  grant execute on sp_lis_interface_get_test_result_model to public
go
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
