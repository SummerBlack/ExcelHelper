# ExcelHelper
Qt平台上Excel读写帮助类<br>
QAxObject是Qt提供的包装COM组件的类，通过COM操作Excel需要使用QAxObject类，使用此类还需要在pro文件增加“QT += axcontainer”
## 打开文件与关闭
void openExcel(const QString &fileName);<br>
不论是读还是写都需要先指定文件的名称，通过openExcel方法指定<br>
void closeExcel();<br>
操作完成后，需要关闭Excel，释放资源

## 文件的读取
文件的读取可以通过以下几个方法实现：
> （1）QVariant readCellValue(int row, int column) const;<br>
>     获取row行column列的单元格中的数据<br>
> （2）bool readTableValue(QList<QList<QVariant>> &value, const QString &range = "");<br>
>     读取指定范围内的单元格中的数据<br>
> （3）bool readFromFile(const QString &fileName, QList<QList<QVariant>> &value, const QString &range = "");<br>
>     指定Excel文件名，读取指定范围内的数据保存至value中<br>
### 文件读取示例
```cpp
  QString file4Read = QFileDialog::getOpenFileName(this,tr("Open"),".",tr("Microsoft Office (*.xlsx *csv)"));
  if (file4Read.isEmpty()) {
        return;
  }
  ExcelHelper excelHelper;
  excelHelper.openExcel(file4Read);
  // 读取某个单元格
  QVariant value = excelHelper.readCellValue(3, 5);
  // 读取指定范围内的单元格中的数据
  QList<QList<QVariant>> vars;
  // 读取1行1列至5行4列的数据
  excelHelper.readTableValue(vars, "A1:D5");
  // 如果觉得"A1:D5"的写法不够直观，可以采用如下转换
  QString range = excelHelper.convertToRangeName(1, 1, 5, 4);
  excelHelper.readTableValue(vars, range);
  
  // readFromFile本质上就是openExcel加readTableValue，上述的写法可以直接写作
  ExcelHelper excelHelper;
  excelHelper.readFromFile(file4Read, value, range);
  // 操作完成，关闭文件
  excelHelper.closeExcel();
```
## 文件的写入
文件的写入可以通过以下几个方法实现：
> （1）void writeCellValue(int row, int column, const QVariant &value);<br>
>     将数据写入row行column列的单元格中<br>
> （2）bool writeTableValue(QList<QList<QVariant>> &values, const int startRow = 1, const int startColumn = 1);<br>
>     将一组数据写入到Excel表格中<br>
> （3）bool writeToFile(const QString &fileName, QList<QList<QVariant>> &values,const int startRow = 1, const int startColumn = 1);<br>
>     将一组数据写入到指定的Excel文件中<br>
### 文件写入示例
```cpp
    QString file4Write = QFileDialog::getSaveFileName(this,tr("Save"),".",tr("Microsoft Office (*.xlsx *csv)"));
    if (file4Write.isEmpty()) {
        return;
    }
    // 向某个单元格写入数据
    ExcelHelper excelHelper;
    excelHelper.openExcel(file4Write);
    excelHelper.writeCellValue(1, 1, QVariant("First"));
    
    // 向excel表格中写入一组数据
    excelHelper.writeTableValue(vars, 2, 1);
    
    // 写入到Excel文件中
    excelHelper.saveExcel(file4Write);
    
    // 以上操作也可直接写作writeToFile，相当于openExcel + setTableValue + saveExcel
    excelHelper.writeToFile(file4Write, vars, 2, 1);
    
    // 操作完成，关闭文件
    excelHelper.closeExcel();
```
### 文件写入应用
  在实际应用中，常常需要将所生产的数据按照指定格式存储到Excel中，如下图所示，通过ExcelHelper类可以方便实现。<br>
              ![](https://github.com/SummerBlack/ExcelHelper/raw/master/excel.png) <br>
```cpp
    ExcelHelper excelHelper;
    // 打开Excel文件
    excelHelper.openExcel(filepath);
    // 向表格1行1列写入"Pitch"
    excelHelper.writeCellValue(1, 1, QVariant("Pitch"));
    excelHelper.writeCellValue(1, 2, 8);
    excelHelper.writeCellValue(2, 1, QVariant("Max Frequency"));
    excelHelper.writeCellValue(2, 2, 224);
    excelHelper.writeCellValue(3, 1, QVariant("Optimise Frequency AF"));
    excelHelper.writeCellValue(3, 2, 112);

    int startRow = 0;
    int startCol = 0;
    int rowCount = 0;
    int colCount = 0;
    // 获取当前的行数
    excelHelper.getRange(startRow, startCol, rowCount, colCount);
    int currentRow = startRow + rowCount;

    excelHelper.writeCellValue(currentRow, 1, QVariant("[Measurement]"));
    excelHelper.writeCellValue(currentRow + 1, 1, QVariant("Position"));
    excelHelper.writeCellValue(currentRow + 1, 2, 1);
    // 保存数据
    QList<QList<QVariant>> datas;
    for (int step = 1; step < 2; ++step) {
        // 第step步
        for (int id = 0; id < 10; ++id) {
            // 第id个相机
            for (int i = 0; i < 2; ++i) {
                // 一行的数据
                QList<QVariant> var;
                var.append(QString("Step%1").arg(step));
                var.append(1.0);

                var.append(QString("Camera %1").arg(id));
                if (i == 0) {
                    var.append("Tan");
                } else {
                    var.append("Sag");
                }

                var.append(100);

                datas.append(var);
            }
        }
    }
    // 写入缓存
    excelHelper.writeTableValue(datas, currentRow + 2, 1);
    // 写入文件中
    excelHelper.saveExcel(filepath);
    excelHelper.closeExcel();
```
## 注意事项
* 分线程中通过ExcelHelper操作Excel，必须先调用CoInitializeEx(NULL, COINIT_MULTITHREADED)初始化
```cpp
    CoInitializeEx(NULL, COINIT_MULTITHREADED);
    ExcelHelper excelHelper;
    excelHelper.openExcel(mFileName);
```
* 在多次写入数据至同一个Excel文件时，可以先通过writeTableValue将数据保存到内存中，等存完了再调用saveExcel写入到本地文件中，因为saveExcel将数据写入到磁盘中比较耗时，测试发现writeTableValue的耗时与本次存入的数据量大小有关，而saveExcel与该Excel文件的总大小有关，即一个很大的文件即使只插入一个数据，然后调用saveExcel保存到磁盘中也会很慢。

