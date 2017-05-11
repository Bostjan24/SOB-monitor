drop table IF EXISTS data;

Create table sob_data (
	entry_id Int NOT NULL AUTO_INCREMENT,
	day Date NOT NULL,
	hour Time NOT NULL,
	happines Int NOT NULL,
	energy Int NOT NULL,
	focus Int NOT NULL,
	UNIQUE (entry_id),
	Index AI_entry_id (entry_id),
 Primary Key (entry_id)) ENGINE = MyISAM;
 
drop table IF EXISTS password;

Create table password (
	entry_id Int NOT NULL AUTO_INCREMENT,
	hash Varchar(1000) NOT NULL,
	salt Varchar(1000) NOT NULL,
	UNIQUE (entry_id),
	UNIQUE (hash),
	Index AI_entry_id (entry_id),
 Primary Key (entry_id)) ENGINE = MyISAM;
