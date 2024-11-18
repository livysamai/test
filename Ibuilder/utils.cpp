#include "utils.h"

Utils::Utils()
{


}

void Utils::mergeExcelFiles(const QStringList &inputFile, const QString &outputFile) {
    // Step 3.1: 构建Python脚本调用命令
    QString pythonExecutable = "C:\\Users\\27340\\AppData\\Local\\Programs\\Python\\Python313\\python.exe"; // 替换为实际Python路径
    QString pythonScript = "D:/QtPrograme/Ibuilder/merged_excel_test.py"; // 替换为实际脚本路径

    // 构建命令行参数
    QStringList arguments;
    arguments << pythonScript;
    for (const QString &file : inputFile) {
        arguments << file;
    }
    arguments << outputFile;

    // Step 3.2: 调用Python脚本
    QProcess process;
    process.start(pythonExecutable, arguments);

    // Step 3.3: 检查脚本执行状态
    if (!process.waitForFinished()) {
        QMessageBox::critical(nullptr, "错误", "Python脚本执行失败");
        return;
    }

    QString output = process.readAllStandardOutput();
    QString error = process.readAllStandardError();

    // Step 3.4: 根据输出判断结果
    if (!error.isEmpty()) {
        QMessageBox::critical(nullptr, "错误", "合并失败：" + error);
    } else {
        QMessageBox::information(nullptr, "成功", "合并完成，文件已保存到：" + outputFile);
    }

    // 输出调试信息
    qDebug() << "Output:" << output;
    qDebug() << "Error:" << error;
}


void Utils::mergeSheetFiles(const QStringList &inputFiles, const QString &outputFile, const QString &sheetName) {
    // Python 解释器路径
    QString pythonExecutable = "C:\\Users\\27340\\AppData\\Local\\Programs\\Python\\Python313\\python.exe"; // 替换为你的 Python 路径
    QString pythonScript = "D:/QtPrograme/Ibuilder/merged_sheet.py"; // 替换为你的脚本路径

    // 构建命令行参数
    QStringList arguments;
    arguments << pythonScript << sheetName;
    arguments.append(inputFiles);
    arguments << outputFile;

    // 调用 Python 脚本
    QProcess process;
    process.start(pythonExecutable, arguments);

    if (!process.waitForFinished()) {
        QMessageBox::critical(nullptr, "错误", "Python脚本执行失败");
        return;
    }

    QString output = process.readAllStandardOutput();
    QString error = process.readAllStandardError();

    if (!error.isEmpty()) {
        QMessageBox::critical(nullptr, "错误", "合并失败：" + error);
        qDebug() << "Error:" << error;
    } else {
        QMessageBox::information(nullptr, "成功", "合并完成，文件已保存到：" + outputFile);
        qDebug() << "Output:" << output;
    }
}

void Utils::mergeExcelFilesToPDF(const QStringList &inputFiles, const QString &outputFile) {
    // Step 1: Python脚本路径和Python执行程序路径
    QString pythonExecutable = "C:\\Users\\27340\\AppData\\Local\\Programs\\Python\\Python313\\python.exe";  // 替换为你的实际 Python 路径
    QString pythonScript = "D:/QtPrograme/Ibuilder/merged_excel_to_PDF.py"; // 替换为实际脚本路径

    // Step 2: 构建命令行参数
    QStringList arguments;
    arguments << pythonScript;
    for (const QString &file : inputFiles) {
        arguments << file; // 添加所有输入的文件
    }
    arguments << outputFile;  // 添加输出文件

    // Step 3: 执行 Python 脚本
    QProcess process;
    process.start(pythonExecutable, arguments);

    // Step 4: 等待脚本执行完成并检查执行状态
    if (!process.waitForFinished()) {
        QMessageBox::critical(nullptr, "错误", "Python脚本执行失败");
        return;
    }

    // 获取脚本执行后的输出和错误信息
    QString output = process.readAllStandardOutput();
    QString error = process.readAllStandardError();

    // Step 5: 根据输出判断结果
    if (!error.isEmpty()) {
        QMessageBox::critical(nullptr, "错误", "合并失败：" + error);
    } else {
        QMessageBox::information(nullptr, "成功", "合并完成，文件已保存到：" + outputFile);
    }

    // 输出调试信息
    qDebug() << "Output:" << output;
    qDebug() << "Error:" << error;
}

