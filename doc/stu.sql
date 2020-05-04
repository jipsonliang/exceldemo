CREATE TABLE `stu` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `name` varchar(20) DEFAULT NULL,
  `sex` varchar(10) DEFAULT NULL,
  `age` double(10,2) DEFAULT NULL,
  `classid` varchar(20) DEFAULT NULL,
  `score` float(10,2) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB AUTO_INCREMENT=1 DEFAULT CHARSET=utf8mb4;

INSERT INTO `stu` VALUES (1, '张三', '0', 18.00, '3班', 72.23);
INSERT INTO `stu` VALUES (2, '李四', '1', 19.00, '1班', 85.50);
INSERT INTO `stu` VALUES (3, '王五', '0', 20.00, '2班', 86.36);