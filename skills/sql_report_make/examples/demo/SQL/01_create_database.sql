-- 创建实验用数据库 S，字符集 utf8mb4 保证中文供应商名不出现乱码
CREATE DATABASE S
  DEFAULT CHARACTER SET utf8mb4
  DEFAULT COLLATE utf8mb4_general_ci;

-- 切换到刚建的库，后续建表都落在这里
USE S;

-- 确认库已建立（结果框里能看到 S 的默认字符集）
SHOW CREATE DATABASE S;
