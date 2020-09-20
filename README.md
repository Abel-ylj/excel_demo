# excel_demo
excel基本操作

## POI

### 简述

> apache提供的，由JAVA编写的免费的跨平台API
>
> 提供API给JAVA程序对Microsoft Office格式档案读写功能

### POI包结构

```
HSSF -- 读写XLS文件，老式文件 .xls
XLSX -- 读写XLSX文件，新式文件 *** .xlsx文件 
HWPF -- 读写WORD DOC文件
HSLF -- 读写PPT
```

### 同类技术对比

- JXl : 内存消耗小，但是对图片和图形支持有限
- Poi: 功能更加完备

## 基本技术概念

1. XSSFWorkbook: 工作簿
2. XSSFSheet: 工作表
3. Row: 行
4. Cell: 单元格

## 案例

### 读



