#ifndef EXCELHELPER_H
#define EXCELHELPER_H

#include <QList>

class QAxObject;
/**
 * @brief 该类为Excel表格操作的辅助类，简化对表格数据的读写操作
 */
class ExcelHelper
{
public:
    ExcelHelper();
    ~ExcelHelper();

    /**
     * @brief 打开/新建一个Excel,如果文件存在就打开，不存在就新建
     * @param fileName 文件的名称
     */
    void openExcel(const QString &fileName);

    /**
     * @brief 设置表格中某一行某一列的值
     * @param row
     * @param column
     * @param value
     */
    void setCellValue(int row, int column, const QVariant &value);

    /**
     * @brief 获取表格中某一行某一列的值
     * @param row
     * @param column
     * @return
     */
    QVariant getCellValue(int row, int column) const;

    /**
     * @brief 查询Excel中的数据并保存到value中，如果指定查询范围就按照指定范围查，如果没有就查询表中所有数据
     * @param value 返回数据保存在value中，其中QList<QVariant>标识一行数据
     * @param range 查询的范围（如A1:Z10表示从第一行第一列到第10行第26列），默认为空，即查询所有
     * @return 是否读取成功
     */
    bool getTableValue(QList<QList<QVariant>> &value, const QString &range = "");

    /**
     * @brief 将values中的数据保存到Excel中
     * @param values 需要保存的数据
     * @param startRow 数据的起始行
     * @param startColumn 数据的起始列
     * @return 是否保存成功
     */
    bool setTableValue(QList<QList<QVariant>> &values, const int startRow = 1, const int startColumn = 1);

    /**
     * @brief 指定Excel文件名，读取其中数据保存至value中
     * @param fileName Excel文件名
     * @param value 数据保存至value
     * @param range 读取数据的范围
     * @return 是否读取成功
     */
    bool readFromFile(const QString &fileName, QList<QList<QVariant>> &value, const QString &range = "");

    /**
     * @brief 将value数据写入到fileName文件中，文件如果不存在则新建
     * @param fileName 保存数据的文件名
     * @param values 数据
     * @param startRow 保存数据的起始行
     * @param startColumn 保存数据的起始列
     * @return 是否成功
     */
    bool writeToFile(const QString &fileName, QList<QList<QVariant>> &values,const int startRow = 1, const int startColumn = 1);

    /**
     * @brief 将（1，1）->（10,26）转换为A1:Z10
     * @param startRow 起始行
     * @param startCol 起始列
     * @param endRow 结束行
     * @param endCol 结束列
     * @return
     */
    QString convertToRangeName(int startRow, int startCol, int endRow, int endCol) const;

    /**
     * @brief 获取excel的起始行、起始列、行数和列数
     * @param startRow 起始行
     * @param startCol 起始列
     * @param rowCount 行数
     * @param colCount 列数
     */
    void getRange(int &startRow, int &startCol, int &rowCount, int &colCount) const;

    /**
     * @brief 保存文件
     * @param fileName 文件名
     */
    void saveExcel(const QString &fileName);

    /**
     * @brief 关闭Excel，停止操作，关闭完了资源被清理了就不能再操作了，否则出错
     */
    void closeExcel();

private:
    /**
     * @brief 将Excel中读取的文件QVariant类型转化为QList<QList<QVariant>>类型
     * @param var
     * @return QList<QVariant>表示一行数据
     */
    QList<QList<QVariant>> castVariant2List(QVariant& var) const;

    /**
     * @brief 将QList<QList<QVariant>>数据转换为Excel可以存储的QVariant类型
     * @param lists
     * @return
     */
    QVariant castList2Variant(QList<QList<QVariant>>& lists) const;

    /**
     * @brief 把列数转换为excel的字母列号，如1->A,2->B...26->Z,27->AA,28->AB...依此类推
     * @param col 大于0的数
     * @param res 字母列号，如1->A 26->Z 27->AA
     */
    void convertToColName(int col, QString &res) const;
private:
    QAxObject *excelApp = nullptr;
    QAxObject *workBooks = nullptr;
    QAxObject *workBook = nullptr;
    QAxObject *workSheets = nullptr;
    QAxObject *workSheet = nullptr;
};

#endif // EXCELHELPER_H
