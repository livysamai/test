QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

include($$PWD/QXlsx/QXlsx.pri)             # QXlsx源代码
INCLUDEPATH += $$PWD/QXlsx

CONFIG += c++17

# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0

SOURCES += \
    main.cpp \
    hellowindow.cpp \
    utils.cpp

HEADERS += \
    hellowindow.h \
    utils.h

FORMS += \
    hellowindow.ui

# Default rules for deployment.
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target

DISTFILES += \
    dif_sheet.py \
    merged_excel.py \
    merged_excel_test.py \
    merged_excel_to_PDF.py \
    merged_sheet.py \
    test.py
