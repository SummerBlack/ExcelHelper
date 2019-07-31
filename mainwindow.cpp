#include "mainwindow.h"
#include "ui_mainwindow.h"
#include "utils/excelhelper.h"
#include <QFileDialog>
#include <QDebug>

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    ui->setupUi(this);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_btn_read_clicked()
{
    QList<QList<QVariant>> vars;
    QString fileName = QFileDialog::getOpenFileName(this,tr("Open"),".",tr("Microsoft Office (*.xlsx *csv)"));

    if (fileName.isEmpty()) {
        return;
    }
    // 读取
    ExcelHelper excelHelper;
    excelHelper.readFromFile(fileName, vars);
    excelHelper.closeExcel();

    foreach (QList<QVariant> var, vars) {
        // 每一行的值
        foreach (QVariant v, var) {
            qDebug() << v.toString();
        }
    }

    // 保存
    ExcelHelper helper2;
    helper2.writeToFile(filepath, vars, 1, 1);
    helper2.closeExcel();
}

void MainWindow::on_btn_write_clicked()
{
    QString filepath = QFileDialog::getSaveFileName(this,tr("Save"),".",tr("Microsoft Office (*.xlsx *csv)"));
    if (filepath.isEmpty()) {
        return;
    }

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
}
