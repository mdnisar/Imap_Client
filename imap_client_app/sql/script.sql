create table mail_mst(
	mm_code int identity(1,1),
	mm_from varchar(300),
	mm_sub varchar(2000),
	mm_date datetime,
	mm_size varchar(50)
);
alter table mail_mst add mm_message_id varchar(200)

create table attachment_mst(
	am_code int identity(1,1),
	file_code int,
	am_name varchar(100),
	mm_message_id varchar(200)
)
create table attachment_dtl
(
	ad_id int identity(1,1),
	mm_message_id varchar(500),
	ad_data1 varchar(500),
	ad_data2 varchar(500),
	ad_data3 varchar(500),
)

--select * from attachment_dtl;
--select * from attachment_mst order by am_code
--select * from mail_mst order by mm_code
--select mm.*,am.am_name from mail_mst mm inner join attachment_mst am on mm.mm_message_id = am.mm_message_id
--
--truncate table mail_mst
--truncate table attachment_dtl
--truncate table attachment_mst