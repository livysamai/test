#include "hellowindow.h"
#include "ui_hellowindow.h"

HelloWindow::HelloWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::HelloWindow)
{
    ui->setupUi(this);
}

HelloWindow::~HelloWindow()
{
    delete ui;
}


void HelloWindow::on_bt_mergeSheet_clicked()
{

}

