#include <QApplication>
#include <QProcess>
#include <QDebug>
#include <QFileInfo>
#include <QPushButton>
#include <QFileDialog>
#include <QMessageBox>
#include <QInputDialog>
#include "xlsxdocument.h"
#include "hellowindow.h"
#include "utils.h"

void merged_excel(const QStringList &inputFile, const QString &outputFile);

void mergeExcelFilesAndConvertToPDF(const QStringList &inputFiles, const QString &outputFile) {
    // 设置 Python 路径和脚本路径
    QString pythonExecutable = "C:\\Users\\27340\\AppData\\Local\\Programs\\Python\\Python313\\python.exe";  // 替换为实际的 Python 路径
    QString pythonScript = "D:/QtPrograme/Ibuilder/merged_excel_to_PDF.py";  // 替换为实际的 Python 脚本路径

    // 设置命令行参数
    QStringList arguments;
    arguments << pythonScript;
    for (const QString &file : inputFiles) {
        arguments << file;
    }
    arguments << outputFile;  // 最后是输出文件路径

    // 使用 QProcess 启动 Python 脚本
    QProcess process;
    process.start(pythonExecutable, arguments);

    // 等待脚本执行完成
    if (!process.waitForFinished()) {
        QMessageBox::critical(nullptr, "错误", "Python 脚本执行失败");
        return;
    }

    // 获取脚本的输出
    QString output = process.readAllStandardOutput();
    QString error = process.readAllStandardError();

    // 如果有错误输出，显示错误信息
    if (!error.isEmpty()) {
        QMessageBox::critical(nullptr, "错误", "合并失败：" + error);
    } else {
        QMessageBox::information(nullptr, "成功", "合并完成，文件已保存到：" + outputFile);
    }

    // 输出调试信息
    qDebug() << "Output:" << output;
    qDebug() << "Error:" << error;
}

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);
    HelloWindow w;
    Utils utils;

    w.show();

    //控件绑定
    QPushButton *bt_mergeExcel = w.findChild<QPushButton*>("bt_mergeExcel");
    QPushButton *bt_mergeSheet = w.findChild<QPushButton*>("bt_mergeSheet");
    //bt_mergeExceltoPDF
    QPushButton *bt_mergeExceltoPDF = w.findChild<QPushButton*>("bt_mergeExceltoPDF");

    QXlsx::Document xlsx;

    QObject::connect(bt_mergeExcel,&QPushButton::clicked,[&](){

        // 弹出文件选择对话框，用户选择多个文件
        QStringList inputFile = QFileDialog::getOpenFileNames(
                    &w,
                    "选择需要合并的Excel文件",
                    "",
                    "Excel Files (*.xlsx *.xls)"
                    );
        // 检查用户是否选择了文件
        if (inputFile.isEmpty()) {
            QMessageBox::warning(&w, "提示", "未选择任何文件");
            return;
        }

        // 弹出保存路径选择框
        QString outputFile = QFileDialog::getSaveFileName(
                    &w,
                    "选择保存合并文件路径",
                    "",
                    "Excel Files (*.xlsx)"
                    );

        if (outputFile.isEmpty())
        {
            QMessageBox::warning(&w, "提示", "未选择保存路径");
            return;
        }

        utils.mergeExcelFiles(inputFile, outputFile);
    });


    QObject::connect(bt_mergeSheet,&QPushButton::clicked,[&](){
        qDebug() << "Merge Excel button clicked.";
        // 文件选择对话框
        QStringList inputFile = QFileDialog::getOpenFileNames(
                    &w,
                    "选择需要合并的Excel文件",
                    "",
                    "Excel Files (*.xlsx *.xls)"
                    );

        if (inputFile.isEmpty()) {
            QMessageBox::warning(&w, "提示", "未选择任何文件");
            return;
        }

        // 获取工作表名称
        bool ok;
        QString sheetName = QInputDialog::getText(
                    &w,
                    "输入工作表名称",
                    "请输入需要合并的工作表名称：",
                    QLineEdit::Normal,
                    "",
                    &ok
                    );

        if (!ok || sheetName.isEmpty()) {
            QMessageBox::warning(&w, "提示", "未输入工作表名称");
            return;
        }

        // 保存路径选择框
        QString outputFile = QFileDialog::getSaveFileName(
                    &w,
                    "选择保存合并文件路径",
                    "",
                    "Excel Files (*.xlsx)"
                    );

        if (outputFile.isEmpty()) {
            QMessageBox::warning(&w, "提示", "未选择保存路径");
            return;
        }

        // 调用合并方法
        utils.mergeSheetFiles(inputFile, outputFile, sheetName);

    });


    QObject::connect(bt_mergeExceltoPDF, &QPushButton::clicked, [&]() {
        // 弹出文件选择对话框，选择多个文件
        QStringList inputFiles = QFileDialog::getOpenFileNames(&w, "选择需要合并的Excel文件", "", "Excel Files (*.xlsx *.xls)");
        if (inputFiles.isEmpty()) {
            QMessageBox::warning(&w, "提示", "未选择任何文件");
            return;
        }

        // 弹出保存路径选择框
        QString outputFile = QFileDialog::getSaveFileName(&w, "选择保存合并文件路径", "", "Excel Files (*.xlsx)");
        if (outputFile.isEmpty()) {
            QMessageBox::warning(&w, "提示", "未选择保存路径");
            return;
        }

        // 调用合并并转换为PDF的函数
        mergeExcelFilesAndConvertToPDF(inputFiles, outputFile);
    });

    return a.exec();
}



