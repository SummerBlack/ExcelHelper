#include "excelhelper.h"
#include <QAxObject>
#include <QDir>
#include <QDebug>
#include <QDateTime>

ExcelHelper::ExcelHelper()
{
    excelApp = new QAxObject();

    // 连接Excel控件
    excelApp->setControl("Excel.Application");
    // 不显示窗体
    excelApp->dynamicCall("SetVisible(bool)","false");
    // 不显示任何警告信息。如果为true那么在关闭时会出现类似“文件已修改，是否保存”的提示
    excelApp->setProperty("DisplayAlerts", false);
    // 获取工作簿集合
    workBooks = excelApp->querySubObject("WorkBooks");
}

ExcelHelper::~ExcelHelper()
{
    closeExcel();
}

void ExcelHelper::openExcel(const QString &fileName)
{
    QFile file(fileName);
    if(file.exists()){
        workBook = workBooks->querySubObject("Open(const QString &)", fileName);
    } else{
        workBooks->dynamicCall("Add");
        workBook = excelApp->querySubObject("ActiveWorkBook");
    }
    // 默认有一个sheet
    workSheets = workBook->querySubObject("Sheets");
    workSheet = workSheets->querySubObject("Item(int)", 1);
}

QVariant ExcelHelper::readCellValue(int row, int column) const
{
    QAxObject *cell = workSheet->querySubObject("Cells(int, int)", row, column);
    return cell->property("Value");
}

bool ExcelHelper::readTableValue(QList<QList<QVariant>> &value, const QString &range)
{
    QVariant var;
    QAxObject *userRange;
    if (range.isNull() || range.isEmpty()) {
        userRange = workSheet->querySubObject("UsedRange");
    }else {
        userRange = workSheet->querySubObject("Range(const QString&)", range);
    }

    if (userRange == nullptr || userRange->isNull()) {
        return false;
    }

    var = userRange->dynamicCall("Value");
    value = castVariant2List(var);
    return true;
}

bool ExcelHelper::readFromFile(const QString &fileName, QList<QList<QVariant>> &value, const QString &range)
{
    QFile file(fileName);

    if(!file.exists()){
        return false;
    }

    workBook = workBooks->querySubObject("Open(const QString &)", fileName);
    //默认有一个sheet
    workSheets = workBook->querySubObject("Sheets");
    workSheet = workSheets->querySubObject("Item(int)", 1);

    bool result = readTableValue(value, range);
    return result;
}

void ExcelHelper::writeCellValue(int row, int column, const QVariant &value)
{
    QAxObject *cell = workSheet->querySubObject("Cells(int, int)", row, column);
    cell->setProperty("Value", value);
}

bool ExcelHelper::writeTableValue(QList<QList<QVariant> > &values, const int startRow, const int startColumn)
{
    qint64 startT = QDateTime::currentMSecsSinceEpoch();
    if(values.size() <= 0){
        return false;
    }

    if(NULL == workSheet || workSheet->isNull()){
        return false;
    }
    bool succ = false;
//    try {
        // 行数
        int row = values.size();
        // 列数
        int col = values.at(0).size();
        // 起始单元,例如1，1最终转化为A1
        QString rangeStart;
        convertToColName(startColumn, rangeStart);
        rangeStart += QString::number(startRow);
        // 终点单元
        QString rangeEnd;
        convertToColName(startColumn + col - 1, rangeEnd);
        rangeEnd += QString::number(startRow + row - 1);
        // 单元范围
        QString range = rangeStart + ":" + rangeEnd;

        qDebug() << range;
        QAxObject *usedRange = workSheet->querySubObject("Range(const QString&)", range);
        if(NULL == usedRange || usedRange->isNull())
        {
            return false;
        }

        QVariant var = castList2Variant(values);
        succ = usedRange->setProperty("Value", var);
//    } catch (...) {
//        qDebug() << "setTableValue Error!";
//    }

    qint64 endT = QDateTime::currentMSecsSinceEpoch();
    qDebug() << "setTableValue: " << endT - startT;
    return succ;
}

bool ExcelHelper::writeToFile(const QString &fileName, QList<QList<QVariant> > &values, const int startRow, const int startColumn)
{
    // 新建或打开文件
    openExcel(fileName);
    // 保存数据
    bool success = writeTableValue(values, startRow, startColumn);
    // 保存excel文件
    saveExcel(fileName);

    return success;
}

QString ExcelHelper::convertToRangeName(int startRow, int startCol, int endRow, int endCol) const
{
    // 起始单元,例如1，1最终转化为A1
    QString rangeStart;
    convertToColName(startCol, rangeStart);
    rangeStart += QString::number(startRow);
    // 终点单元
    QString rangeEnd;
    convertToColName(endCol, rangeEnd);
    rangeEnd += QString::number(endRow);
    // 单元范围
    QString range = rangeStart + ":" + rangeEnd;

    qDebug() << range;
    return range;
}

void ExcelHelper::getRange(int &startRow, int &startCol, int &rowCount, int &colCount) const
{
    QAxObject *usedrange = workSheet->querySubObject("UsedRange");//sheet范围
    startRow = usedrange->property("Row").toInt();//起始行数
    startCol = usedrange->property("Column").toInt(); //起始列数
    QAxObject *rows, *columns;
    rows = usedrange->querySubObject("Rows");//行
    columns = usedrange->querySubObject("Columns");//列
    rowCount = rows->property("Count").toInt();//行数
    colCount = columns->property("Count").toInt();//列数
    qDebug() << "startRow :" << startRow << "startCol" << startCol << "rowCount" << rowCount << "colCount" << colCount;
}

void ExcelHelper::saveExcel(const QString &fileName)
{
    //保存至filepath，注意一定要用QDir::toNativeSeparators将路径中的"/"转换为"\"，不然一定保存不了。
    workBook->dynamicCall("SaveAs(const QString&)",QDir::toNativeSeparators(fileName));
}

void ExcelHelper::closeExcel()
{
    if (excelApp) {
        //关闭工作簿
        workBook->dynamicCall("Close()");
        //关闭excel
        excelApp->dynamicCall("Quit()");
        delete excelApp;
        excelApp = nullptr;
    }
}

QList<QList<QVariant>> ExcelHelper::castVariant2List(QVariant &var) const
{
    QList<QList<QVariant>> list;
    QList<QVariant> rows = var.toList();
    for(auto row : rows) {
        list.append(row.toList());
    }
    return list;
}

QVariant ExcelHelper::castList2Variant(QList<QList<QVariant>> &lists) const
{
    QVariantList vars;
    for (auto list : lists){
        vars.append(QVariant(list));
    }
    return QVariant(vars);
}

void ExcelHelper::convertToColName(int col, QString &res) const
{
    Q_ASSERT(col > 0 && col < 65535);
    int tempData = col / 26;
    if(tempData > 0){
        int mode = col % 26;
        convertToColName(mode, res);
        convertToColName(tempData, res);
    } else {
        res = QString(col + 0x40) + res;
    }
}
