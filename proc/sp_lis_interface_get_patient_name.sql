print 'sp_lis_interface_get_patient_name'
go
/*  $Header$ */
/* $Log$
*/
if exists (select name from sysobjects where name='sp_lis_interface_get_patient_name' and type='P ')
  drop proc sp_lis_interface_get_patient_name
go

Create proc [dbo].[sp_lis_interface_get_patient_name]  
	@lab_no nvarchar(20)
 
as
 
/* from rsp_comments  at Oct 22 2013  4:45:38:847PM*/
 
set nocount on
 
/*Author		:	Tui*/
/*Created Date	:	Oct 22, 2013*/
/*Description	:	Get patient information by lab number*/
 
 
declare @pt table
(
	full_name nvarchar(510) null,
	hn nvarchar(20) null,
	lastname nvarchar(155) null,
	firstname nvarchar(155) null,
	middle_name nvarchar(155) null,
	birthday nvarchar(8) null,
	sex nvarchar(1) null
)
 
insert @pt
select	distinct isnull(pf.prefix_desc,'') + isnull(p.firstname,'') +' '+ isnull(p.middle_name,'')+' '+isnull(p.lastname,'') as full_name,
		case when ltrim(rtrim(p.hn)) = '' then 'NA' else p.hn end as hn,
		isnull(p.lastname,'') as lastname, isnull(p.firstname,'') as firstname, isnull(p.middle_name,'') as middle_name,
		isnull(convert(nvarchar(8),birthday,112),'') birthday,
		----(case when isnull(sex,'') = '' then 'U' else sex end) as sex --Golf Comment 2019-05-29
		(case isnull(sex,'') when '' then 'U' when 'A' then 'U' when 'O' then 'U'  else sex end) as sex --Golf add 2019-05-29
from	lis_lab_order lo
		left join lis_lab_specimen_type lst
			on lo.order_skey = lst.order_skey
		left join lis_visit v
			on lo.visit_skey = v.visit_skey
		left join lis_patient p
			on v.patient_skey = p.patient_skey
		left join patho_prefix pf
			on p.prefix = pf.prefix_cd
where	lo.order_id = @lab_no or lst.specimen_type_id = @lab_no
 
 
insert into @pt(full_name,hn)
values ('NA','NA')
  
 
select top 1 full_name, isnull(hn,'NA') as hn, firstname, lastname, birthday, sex, middle_name
from @pt
 
/*
exec sp_lis_interface_get_patient_name 'L20130101-000072'
exec sp_lis_interface_get_patient_name '42505477'
*/
 
 
 
 
go
if exists (select name from sysobjects where name='sp_lis_interface_get_patient_name' and type='P ')
  grant execute on sp_lis_interface_get_patient_name to public
go
 
 
 
 
 
 
