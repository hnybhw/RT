# py_re - 再保险精算引擎（Python重构版）

## 项目简介

基于xlwings+Pandas重构的工业级再保险精算计算引擎，适配390万行数据，兼容VBA调用。

## 项目结构

py_re/
├── __init__.py # 包入口，暴露核心接口（main/run_segment）
├── config/ # 配置中心（抽离硬编码）
│ ├── __init__.py
│ ├── schema.py # 字段映射SCHEMA（对齐VBA数据结构）
│ └── constants.py # 常量定义（默认值/路径标识）
├── utils/ # 通用工具层（复用+解耦）
│ ├── __init__.py
│ ├── logger.py # 日志工具（对齐VBA WriteLog规则）
│ ├── excel_io.py # Excel IO工具（静默重试+批量读写）
│ └── common.py # 通用函数（类型转换/数据校验）
├── data_model/ # 数据模型层（消除魔法数字）
│ ├── __init__.py
│ ├── asmp.py # 假设数据模型（AsmpRow/AsmpRule）
│ ├── loss.py # 损失数据模型（LossRow）
│ └── recovery.py # 赔付结果模型（RecoveryRow）
├── data_process/ # 数据预处理层
│ ├── __init__.py
│ ├── asmp_processor.py # 假设数据处理（去重/规则构建）
│ └── loss_processor.py # 损失数据预处理（过滤/格式化）
├── calculation/ # 核心计算层（工厂模式+策略模式）
│ ├── __init__.py
│ ├── base_engine.py # 计算引擎基类（抽象接口）
│ ├── excess_of_loss.py # 超赔合同引擎（原核心逻辑）
│ ├── aggregator.py # 事件聚合引擎（aggregate_events）
│ ├── recovery_allocator.py # 赔付拆分引擎（split_recovery_by_audit_trail）
│ ├── cap_processor.py # 限额处理引擎（apply_qs_event_cap等）
│ └── engine_factory.py # 引擎工厂（动态创建计算实例）
├── pivot/ # 透视转换层
│ ├── __init__.py
│ └── pivot_converter.py # 宽表转换（CSV临时文件优化）
└── orchestration/ # 编排调度层
 ├── __init__.py
 └── task_orchestrator.py # 任务编排（原orchestrator逻辑）

## 环境依赖

- Python 3.8+

- pandas==2.0.3

- xlwings>=0.28.0
  
  ## 快速开始
  
  ### 1. 安装依赖
  
  pip install -r requirements.txt
  
  ### 2. 运行测试
  
  python py_re/main.py