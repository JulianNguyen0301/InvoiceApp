create database InvoiceDatabase
go

use InvoiceDatabase
go
 
create table Supplier
(
	Id varchar(20) primary key,
	NameSupplier nvarchar(max),	
)
go


create table UoM
(
	NameUoM nvarchar(50) primary key
)
go

create table Product
(
	Id varchar(10) primary key,
	NameProduct nvarchar(max),
	IdSupplier varchar(20),
	UoM nvarchar(50),
	Price money

	foreign key (IdSupplier) references dbo.Supplier(Id),
	foreign key (UoM) references dbo.UoM(NameUoM),
)
go

alter table Supplier add AddressSupplier nvarchar(max)
go
alter table Product add STT int identity(1,1)
go

insert into Supplier values('0300792451','CONG TY TRACH NHIEM HUU HAN NUOC GIAI KHAT COCA-COLA VIET NAM','485 duong Xa Lo Ha Noi, Phuong Linh Trung, Thanh Pho Thu Duc, Thanh Pho Ho Chi Minh, Viet Nam')
insert into Supplier values('0313107618','CONG TY TNHH THUC PHAM LINH KHOA','29/22 Duong 42, Khu pho 8, Phuong Hiep Binh Chanh, Thanh pho Ho Chi Minh, Viet Nam')

insert into UoM values('kg')
insert into UoM values('Qu?')
insert into UoM values('chai')
insert into UoM values('Két')
insert into UoM values('cái')

insert into Product values('2097','Do uong SPRITE 300ML 4X6 PET CARTON 2.0','0300792451','Két', 64858)
insert into Product values('2704','Do uong DASANI 510ML 4X6 PET SF PRINTED','0300792451','Két', 75977)
insert into Product values('3709','Do uong FANTA ORANGE 300ML 4X6 PET CARTON','0300792451','Két', 64858)
insert into Product values('5108','Do uong COKE 300ML 4X6 PET CARTON','0300792451','Két', 64858)
insert into Product values('2491','Do uong SPRITE 300ML 4X6 PET CARTON PROMO','0300792451','Két', 64858)
insert into Product values('1761','Do uong COKE LIGHT 320ML 4X6 SLEEK CAN SF','0300792451','Két', 183456)
insert into Product values('3700','Do uong FANTA ORANGE 300ML 4X6 PET CARTON PROMO','0300792451','Két', 64858)
insert into Product values('2718','Do uong DASANI 510ML 24 PET CARTON PROMO','0300792451','Két', 75977)
insert into Product values('P000490','Du cam tay, KAON','0300792451','cái',0)

select * from Product
