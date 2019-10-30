# Color-Shift
This is a code by Excel VBA to visualize color shift at view cone from raw data in CIEXYZ into distribution map
you can use RawData 1 to 3 to get familiar with the method.

1. Run VisualizationV3 20190116.xlsm
2. Input details about x数据所在页, y数据所在页, x起始单元格 and y起始单元格 based on the true value of the data files.
   for demonstratioin, this information has been set in advance.
3. In the sheet "结果", you can click 批量读取数据 to run batch processing. 
   a. This function require all files, including raw data and main program in one folder.
   b. This function require all raw data files with the step of 1° for both   
   It will generate result file for each raw data file, and will be saved in \色偏数据
