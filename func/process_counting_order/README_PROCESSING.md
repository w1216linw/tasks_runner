# 订单数据处理说明 (Order Data Processing Guide)

## 概述 (Overview)

本项目用于处理和分析TEMU和YY的订单数据，合并去重后按州统计订单量。

## 数据文件 (Data Files)

### 输入文件
- **TEMU文件**: Excel格式 (.xlsx)，包含TEMU平台订单
  - **关键列**:
    - `运单号` - Tracking number（用于去重匹配）
    - `下单时间` - Order time（用于业务日期计算）
    - `邮编` - ZIP code（用于州信息映射）
- **YY文件**: CSV格式 (.csv)，包含YY平台订单
  - **关键列**:
    - `单据号` - Tracking number（用于去重匹配，对应TEMU的"运单号"）
    - `预约揽收时间` - Pickup time（用于日期过滤）
    - `目的地行政洲` - Destination state（用于州信息）

### 输出文件
1. **订单汇总文件**: `orders_*.csv` - 按州和日期统计的订单量
2. **统计摘要文件**: `statistics_summary_*.csv` - 整体统计数据
3. **每日统计文件**: `daily_statistics_*.csv` - 每日详细统计（仅周处理）

## 处理逻辑 (Processing Logic)

### 1. 数据读取与日期处理
- **TEMU**: 使用8am业务日期规则（时间窗口）
  - **业务日期定义**: 从前一天8:00 AM到当天8:00 AM的订单计入当天
  - **示例**: 获取11/23的订单
    - 时间窗口: **11/22 8:00 AM - 11/23 8:00 AM**
    - 包含: 11/22 8:00 AM及之后的订单
    - 不包含: 11/23 8:00 AM及之后的订单（这些算11/24）
  - **代码逻辑**:
    - 订单时间 < 当日8:00 AM → 算作当日
    - 订单时间 ≥ 当日8:00 AM → 算作次日
- **YY**: 使用预约揽收时间的日期（自然日）

### 2. 州信息映射
- **TEMU**: 使用pgeocode库通过ZIP code查询州代码
- **YY**: 直接使用州名映射到州代码

### 3. 去重逻辑 (Deduplication)
**重要**: TEMU和YY可能有重复订单，需要通过tracking number去重

#### 去重使用的列名对应关系：
- **TEMU**: `运单号` (tracking number)
- **YY**: `单据号` (tracking number)
- **对应关系**: TEMU的"运单号" = YY的"单据号"

#### 去重步骤：
```
1. 读取所有TEMU订单，获取所有"运单号"值，存入集合
2. 读取YY订单，获取所有"单据号"值
3. 对比两个集合，找出重复的tracking numbers
4. 从YY订单中移除所有"单据号"存在于TEMU"运单号"**中的订单**
5. 仅保留YY工作日订单（周一至周五）
6. 合并TEMU原始订单 + 去重后的YY订单
```

#### 示例：
```
TEMU运单号集合: {GFUS01019396404866, GFUS01019412964609, ...}
YY单据号集合:   {GFUS01019396404866, GFUS01019565607104, ...}

重复订单: GFUS01019396404866 (同时存在于两个系统)
→ 从YY中移除这个订单，避免重复计数
```

### 4. 统计指标 (Statistics Metrics)

#### 总体统计 (Overall Statistics)
- **Total TEMU Orders**: TEMU订单总数
- **Total YY Orders (After Dedup)**: YY去重后订单总数
- **Combined Total**: TEMU + YY去重后的合并总数

#### 每日统计 (Daily Statistics, 仅周处理)
- **TEMU_Orders**: 当日TEMU订单数
- **YY_Dedup**: 当日YY去重后订单数
- **Combined_Total**: 当日合并总数

### 5. 计算公式 (Formulas)

```
去重后YY订单 = YY原始订单 - (YY订单 ∩ TEMU订单) (仅工作日)
合并总数 = TEMU订单 + 去重后YY订单
```

## 处理脚本 (Processing Scripts)

### 单日处理 (Daily Processing)
**文件**: `process_daily_template.py`

**用途**: 处理单个日期的数据

**输入**:
- 1个TEMU文件: `TEMU_MMDD.xlsx`
- 1个YY文件: `YY_MMDD.csv`

**输出**:
- `orders_daily_YYMMDD.csv` - 按州统计
- `statistics_summary_YYMMDD.csv` - 统计摘要

**使用示例**:
```python
# 修改文件名
temu_file = 'TEMU_1124.xlsx'
yy_file = 'YY_1124.csv'
output_file = 'orders_daily_251124.csv'
stats_file = 'statistics_summary_251124.csv'
```

### 周处理 (Weekly Processing)
**文件**: `process_weekly_template.py`

**用途**: 处理一周的数据

**输入**:
- 1个或多个TEMU文件: `TEMU_MMDD-MMDD.xlsx`
- 1个YY文件: `YY_MMDD-MMDD.csv`

**输出**:
- `orders_weekly_YYMMDD-YYMMDD.csv` - 按州和日期统计
- `statistics_summary_YYMMDD-YYMMDD.csv` - 统计摘要
- `daily_statistics_YYMMDD-YYMMDD.csv` - 每日统计

**使用示例**:
```python
# 修改文件名列表
temu_files = ['TEMU_1117-1119.xlsx', 'TEMU_1120-1123.xlsx']
yy_file = 'YY11.17-11.23.csv'
output_file = 'orders_weekly_251117-251123.csv'
stats_file = 'statistics_summary_251117-251123.csv'
daily_stats_file = 'daily_statistics_251117-251123.csv'
```

## 注意事项 (Important Notes)

1. **日期格式**:
   - TEMU: `MM/DD/YYYY HH:MM:SS`
   - YY: 可能有多种格式，脚本会自动解析

2. **工作日过滤**: YY订单仅保留周一至周五

3. **ZIP code处理**:
   - Puerto Rico特殊处理 (006-009开头)
   - 使用pgeocode库查询美国ZIP code

4. **数据质量**:
   - 如果有多个同日期的数据文件，使用最新、最完整的版本
   - 周文件的数据质量通常优于单日文件

5. **性能**:
   - ZIP code查询较慢，每500个会显示进度
   - 大文件处理可能需要数分钟

## 依赖库 (Dependencies)

```bash
pip install pandas openpyxl pgeocode
```

## 文件命名规范 (File Naming Convention)

- 输入TEMU: `TEMU_MMDD.xlsx` 或 `TEMU_MMDD-MMDD.xlsx`
- 输入YY: `YY_MMDD.csv` 或 `YYMMDD-MMDD.csv`
- 输出订单: `orders_daily_YYMMDD.csv` 或 `orders_weekly_YYMMDD-YYMMDD.csv`
- 输出统计: `statistics_summary_YYMMDD.csv`

## 常见问题 (FAQ)

**Q: 为什么YY订单会减少这么多？**
A: 因为YY和TEMU有大量重复订单（通常60-80%），去重后仅保留YY独有订单。

**Q: 为什么11.17的数据特别多？**
A: 11.17是周一，通常包含周末累积的订单。

**Q: 统计文件中的百分比如何理解？**
A: 百分比是相对于YY原始订单总数的占比。

## 更新历史 (Change Log)

- 2024-11-25: 添加统计摘要和每日统计功能
- 2024-11-24: 创建单日和周处理模板
- 2024-11-17: 初始版本
