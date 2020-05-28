#include "cmainwindow.h"
#include "ui_cmainwindow.h"

#include "cserverentry.h"

#include <QDebug>


cMainWindow::cMainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::cMainWindow)
{
    ui->setupUi(this);

	cServerEntryList	serverEntryList;

	serverEntryList.loadPath("C:/Users/birkeh/Downloads/TSA");
	qDebug() << serverEntryList.count();
//	serverEntry.load("C:\\Users\\birkeh\\Downloads\\TSA\\RP201810_ASFLIP21.xls");
}

cMainWindow::~cMainWindow()
{
    delete ui;
}

