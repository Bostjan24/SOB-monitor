drop table IF EXISTS data;




Create table data (
	entry_id Int NOT NULL AUTO_INCREMENT,
	date Date NOT NULL,
	time Time NOT NULL,
	happines Int NOT NULL,
	energy Int NOT NULL,
	focus Int NOT NULL,
	UNIQUE (entry_id),
	Index AI_entry_id (entry_id),
 Primary Key (entry_id)) ENGINE = MyISAM;
