# 司机周报数据分析工作流程

## 工作流程

### 步骤 1: 处理当前周数据
使用 `process_dwa.py` 处理原始数据生成分析文件

**命令**：
```bash
python process_dwa.py -f dwd_<当前周>.xlsx Transactions_<当前周>.xlsx -o dwa_<当前周>.xlsx
```

**示例**：
```bash
python process_dwa.py -f dwd_1013.xlsx Transactions_10-13.xlsx -o dwa_1013.xlsx
```

## 完整示例

假设当前周是 10-13，上一周是 10-06：

**步骤1**: 运行命令行
```bash
python process_dwa.py -f dwd_1013.xlsx Transactions_10-13.xlsx -o dwa_1013.xlsx
```

**步骤2**: 在 `weekly_report.ipynb` 中修改
```python
df_106 = pd.read_excel('dwa_106.xlsx')    # 上周
df_1013 = pd.read_excel('dwa_1013.xlsx')  # 当前周
```

---

## 文件说明

### 输入文件
- `dwd_<MMDD>.xlsx` - 司机工作原始数据（driver, hours, tasks, miles, date）
- `Transactions_<MMDD>.xlsx` - 加油交易原始数据

### 输出文件
- `dwa_<MMDD>.xlsx` - 单周分析结果（含各项效率指标）

---