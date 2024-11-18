#ifndef UTILS_H
#define UTILS_H
#include "xlsxdocument.h"
#include <QProcess>
#include <QMessageBox>
#include <QStringList>
#include <QDebug>

//实现基本工具的类
class Utils
{
public:
    Utils();
    void mergeExcelFiles(const QStringList &inputFile, const QString &outputFile);
    void mergeSheetFiles(const QStringList &inputFiles, const QString &outputFile, const QString &sheetName);
    void mergeExcelFilesToPDF(const QStringList &inputFiles, const QString &outputFile);

};

#endif // UTILS_H
