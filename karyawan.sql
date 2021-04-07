# Host: localhost  (Version: 5.5.5-10.1.13-MariaDB)
# Date: 2021-01-26 16:16:12
# Generator: MySQL-Front 5.3  (Build 4.115)

/*!40101 SET NAMES utf8 */;

#
# Structure for table "data_karyawan"
#

DROP TABLE IF EXISTS `data_karyawan`;
CREATE TABLE `data_karyawan` (
  `Id` int(11) NOT NULL AUTO_INCREMENT,
  `TGL_CREATE` varchar(255) DEFAULT NULL,
  `NIK` varchar(50) DEFAULT NULL,
  `NAMA` varchar(50) DEFAULT NULL,
  `ALAMAT_KTP` longtext,
  `ALAMAT_TIGGAL` longtext,
  `JK` varchar(25) DEFAULT NULL,
  `Tlp` varchar(30) DEFAULT NULL,
  `HP` varchar(30) DEFAULT NULL,
  `TGL_MASUK` varchar(30) DEFAULT NULL,
  `TGL_KELUAR` varchar(30) DEFAULT NULL,
  `TGL_AJU_KELUAR` varchar(30) DEFAULT NULL,
  `NO_KTP` varchar(30) DEFAULT NULL,
  `PDDK_AKHIR` varchar(30) DEFAULT NULL,
  `ANAK_KE` varchar(5) DEFAULT NULL,
  `SAUDARA` varchar(5) DEFAULT NULL,
  `JABATAN` varchar(30) DEFAULT NULL,
  `STS_KARYAWAN` varchar(30) DEFAULT NULL,
  `ALASAN_KELUAR` longtext,
  `STS_NIKAH` varchar(30) DEFAULT NULL,
  `JML_ANAK` varchar(5) DEFAULT NULL,
  `KET_LAIN` longtext,
  `LAST_UPDATE` varchar(30) DEFAULT NULL,
  `TMP_LAHIR` varchar(100) DEFAULT NULL,
  `TGL_LAHIR` varchar(30) DEFAULT NULL,
  `NM_SMIATAUIST` longtext,
  `HP_SMIATAUIST` longtext,
  `NM_ANAK` longtext,
  `USIA_ANAK` longtext,
  `NM_ECON` longtext,
  `HUB_ECON` longtext,
  `ALMT_ECON` longtext,
  `HP_ECON` longtext,
  `KERJA1` longtext,
  `KERJA2` longtext,
  `KERJA3` longtext,
  `RECORD` longtext,
  `TGLKONTRAK1` varchar(255) DEFAULT NULL,
  `TGLKONTRAK2` varchar(255) DEFAULT NULL,
  `status` varchar(255) DEFAULT NULL,
  `userdt` varchar(255) DEFAULT NULL,
  `keteranganlain2` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`Id`)
) ENGINE=InnoDB AUTO_INCREMENT=48 DEFAULT CHARSET=utf8;

#
# Data for table "data_karyawan"
#

INSERT INTO `data_karyawan` VALUES (1,'2021-01-01','01.08.09','TESTIG','JL TESTING JAKARTA','JL TESTING JAKARTA UTARA','PEREMPUAN','-','0821332122','1/1/2021','-','-','3234324324432432','S1','2','3','AGENT TUNGGAL','TETAP','','MENIKAH','-',' ','','-','-','-','-','-','-','-','-','-','-','-','-','-','-','8/3/2009','2009-08-03','0',NULL,NULL);

#
# Structure for table "master_list_menu"
#

DROP TABLE IF EXISTS `master_list_menu`;
CREATE TABLE `master_list_menu` (
  `Id_menu` int(11) NOT NULL AUTO_INCREMENT,
  `menu_name` varchar(255) DEFAULT NULL,
  `url_link` varchar(255) DEFAULT NULL,
  `created_date` varchar(255) DEFAULT NULL,
  `updated_date` varchar(255) DEFAULT NULL,
  `created_by` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`Id_menu`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "master_list_menu"
#

INSERT INTO `master_list_menu` VALUES (1,'dashboard','dashboard','2021-01-01','2021-01-01','admin'),(2,'master','master','2021-01-01','2021-01-01','admin'),(3,'data karyawan','datakaryawan','2021-01-01','2021-01-01','admin');

#
# Structure for table "master_role"
#

DROP TABLE IF EXISTS `master_role`;
CREATE TABLE `master_role` (
  `id_role` int(11) NOT NULL AUTO_INCREMENT,
  `role_name` varchar(255) DEFAULT NULL,
  `created_date` varchar(255) DEFAULT NULL,
  `update_date` varchar(255) DEFAULT NULL,
  `creted_by` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`id_role`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "master_role"
#

INSERT INTO `master_role` VALUES (1,'admin','2021-01-01','2021-01-01','admin'),(2,'user','2021-01-01','2021-01-01','admin');

#
# Structure for table "master_role_menu"
#

DROP TABLE IF EXISTS `master_role_menu`;
CREATE TABLE `master_role_menu` (
  `id_role_menu` int(11) NOT NULL AUTO_INCREMENT,
  `menu_id` varchar(255) DEFAULT NULL,
  `role_id` varchar(255) DEFAULT NULL,
  `created_date` varchar(255) DEFAULT NULL,
  `update_date` varchar(255) DEFAULT NULL,
  `created_by` varchar(255) DEFAULT NULL,
  PRIMARY KEY (`id_role_menu`)
) ENGINE=InnoDB DEFAULT CHARSET=latin1;

#
# Data for table "master_role_menu"
#

INSERT INTO `master_role_menu` VALUES (1,'1','1','2021-01-01','2021-01-01','admin'),(2,'2','1','2021-01-01','2021-01-01','admin'),(3,'3','1','2021-01-01','2021-01-01','admin'),(4,'1','2','2021-01-01','2021-01-01','admin'),(5,'4','2','2021-01-01','2021-01-01','admin');

#
# Structure for table "master_users"
#

DROP TABLE IF EXISTS `master_users`;
CREATE TABLE `master_users` (
  `id_user` int(11) NOT NULL,
  `usernames` varchar(30) DEFAULT NULL,
  `role_id` int(1) DEFAULT NULL,
  `password` varchar(50) DEFAULT NULL,
  `active` varchar(255) DEFAULT NULL,
  `block` varchar(255) DEFAULT NULL,
  `created_date` varchar(255) DEFAULT NULL,
  `updated_date` varchar(255) DEFAULT NULL,
  `created_by` varchar(255) DEFAULT NULL,
  UNIQUE KEY `id_user` (`id_user`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

#
# Data for table "master_users"
#

INSERT INTO `master_users` VALUES (1,'admin@admin.co.id',1,'12345','Y','1','2021-01-01','2021-01-01','admin'),(2,'karyawan@karyawan.co.id',2,'12345','Y','1','2021-01-01','2021-01-01','admin');
