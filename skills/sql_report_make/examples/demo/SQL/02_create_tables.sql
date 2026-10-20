-- 供应商表 S：SNO 主键唯一标识一家供应商
CREATE TABLE S (
  SNO    CHAR(4)      NOT NULL,   -- 供应商号，定长 4 位
  SNAME  VARCHAR(20)  NOT NULL,   -- 供应商名
  STATUS INT          DEFAULT 20, -- 状态值，默认 20
  CITY   VARCHAR(20),             -- 所在城市
  PRIMARY KEY (SNO)               -- 主键约束：唯一且非空
);

-- 零件表 P：PNO 主键；PNAME 之后要建唯一索引
CREATE TABLE P (
  PNO    CHAR(4)      NOT NULL,   -- 零件号
  PNAME  VARCHAR(20)  NOT NULL,   -- 零件名
  COLOR  VARCHAR(10),             -- 颜色
  WEIGHT DECIMAL(6,2),            -- 重量，保留两位小数
  PRIMARY KEY (PNO)
);

-- 查看两张表的结构，确认主键已生效
DESC S;
DESC P;
