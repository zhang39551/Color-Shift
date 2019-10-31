# Color-Shift
This is a code by Excel VBA to visualize color shift at view cone from raw data in CIEXYZ into distribution map
you can use RawData 1 to 3 to get familiar with the method.

1. Run VisualizationV3 20190116.xlsm, there are 4 operation buttons, like "导入色偏xy数据" and "批量读取数据" in sheet"结果", and "生成图像" and "模拟色偏批量处理" in sheet"数据格式转换_原始数据".

2. For "导入色偏xy数据" and "批量读取数据" in sheet"结果", you have to input details about x数据所在页, y数据所在页, x起始单元格 and y起始单元格 at first, based on the true value of raw data files.
   ps1: for demonstratioin, this information has been set in advance.
   ps2: for "批量读取数据", this function require all files in one folder, including raw data and main code.
   ps3: for "导入色偏xy数据" and "批量读取数据", raw data files are from a specific test device, so the data are arranged by specific method in 1° step. If files do not match the format, you can use "生成图像" and "模拟色偏批量处理" in sheet"数据格式转换_原始数据".
   ps4: for "批量读取数据" result files will be generated for each raw data file, and will be saved with "色偏数据" added in the name.
   
3. For "生成图像" and "模拟色偏批量处理" in sheet"数据格式转换_原始数据", you have to specify the angle of theta and phi, corresponding to the color in CIEXYZ.
   ps: other parameters have also been set in advance.
