#ifndef HELLOWINDOW_H
#define HELLOWINDOW_H

#include <QPushButton>
#include <QMainWindow>

QT_BEGIN_NAMESPACE
namespace Ui { class HelloWindow; }
QT_END_NAMESPACE

class HelloWindow : public QMainWindow
{
    Q_OBJECT

public:
    HelloWindow(QWidget *parent = nullptr);
    ~HelloWindow();

private slots:
    void on_bt_mergeSheet_clicked();

private:
    Ui::HelloWindow *ui;

};
#endif // HELLOWINDOW_H
