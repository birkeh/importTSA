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
	serverEntryList.save("C:/Users/birkeh/Downloads/TSA/iSeries.xlsx");
}

cMainWindow::~cMainWindow()
{
    delete ui;
}
