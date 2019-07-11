# ExcelHelper
Qt平台上Excel读写帮助类
QAxObject是Qt提供的包装COM组件的类，通过COM操作Excel需要使用QAxObject类，使用此类还需要在pro文件增加“QT += axcontainer”
## 打开文件与关闭
void openExcel(const QString &fileName);<br>
不论是读还是写都需要先指定文件的名称，通过openExcel方法指定<br>
void closeExcel();<br>
操作完成后，需要关闭Excel，释放资源

## 文件的读取
文件的读取可以通过以下几个方法实现：
> （1）QVariant getCellValue(int row, int column) const;<br>
>     获取row行column列的单元格中的数据<br>
> （2）bool getTableValue(QList<QList<QVariant>> &value, const QString &range = "");<br>
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
  QVariant value = excelHelper.getCellValue(3, 5);
  // 读取指定范围内的单元格中的数据
  QList<QList<QVariant>> vars;
  // 读取1行1列至5行4列的数据
  excelHelper.getTableValue(vars, "A1:D5");
  // 如果觉得"A1:D5"的写法不够直观，可以采用如下转换
  QString range = excelHelper.convertToRangeName(1, 1, 5, 4);
  excelHelper.getTableValue(vars, range);
  
  // readFromFile本质上就是openExcel加getTableValue，上述的写法可以直接写作
  ExcelHelper excelHelper;
  excelHelper.readFromFile(file4Read, value, range);
  // 操作完成，关闭文件
  excelHelper.closeExcel();
```  

## 文件的写入
文件的写入可以通过以下几个方法实现：
> （1）void setCellValue(int row, int column, const QVariant &value);<br>
>     将数据写入row行column列的单元格中<br>
> （2）bool setTableValue(QList<QList<QVariant>> &values, const int startRow = 1, const int startColumn = 1);<br>
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
    excelHelper.setCellValue(1, 1, QVariant("First"));
    
    // 向excel表格中写入一组数据
    excelHelper.setTableValue(vars, 2, 1);
    
    // 写入到Excel文件中
    excelHelper.saveExcel(file4Write);
    
    // 以上操作也可直接写作writeToFile，相当于openExcel + setTableValue + saveExcel
    excelHelper.writeToFile(file4Write, vars, 2, 1);
    
    // 操作完成，关闭文件
    excelHelper.closeExcel();
```
在实际应用中，常常需要将所生产的数据按照指定格式存储到Excel中，如下图所示，通过ExcelHelper类可以方便实现。
    ```cpp
    ExcelHelper excelHelper;
    // 打开Excel文件
    excelHelper.openExcel(filepath);
    // 向表格1行1列写入"Pitch"
    excelHelper.setCellValue(1, 1, QVariant("Pitch"));
    excelHelper.setCellValue(1, 2, 8);
    excelHelper.setCellValue(2, 1, QVariant("Max Frequency"));
    excelHelper.setCellValue(2, 2, 224);
    excelHelper.setCellValue(3, 1, QVariant("Optimise Frequency AF"));
    excelHelper.setCellValue(3, 2, 112);

    int startRow = 0;
    int startCol = 0;
    int rowCount = 0;
    int colCount = 0;
    // 获取当前的行数
    excelHelper.getRange(startRow, startCol, rowCount, colCount);
    int currentRow = startRow + rowCount;

    excelHelper.setCellValue(currentRow, 1, QVariant("[Measurement]"));
    excelHelper.setCellValue(currentRow + 1, 1, QVariant("Position"));
    excelHelper.setCellValue(currentRow + 1, 2, 1);
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
    excelHelper.setTableValue(datas, currentRow + 2, 1);
    // 写入文件中
    excelHelper.saveExcel(filepath);
    excelHelper.closeExcel();
```
## 多线程的
