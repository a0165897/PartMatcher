
#include <QApplication>
#include <QDebug>
#include <QElapsedTimer>
//stl
#include <tuple>
//excelio
#include "excelio.h"
//import window
#include "mainwindow.h"

int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    QElapsedTimer timer;
    timer.start();
    //打开xlsx文档
    const QString fileName = QStringLiteral("20E间隙计算测量数据20210707(1).xlsx");

    QXlsx::Document &excelBook = *new QXlsx::Document(fileName);
    excelio::partData data(excelBook);
    qDebug()<<"********file: "<<fileName<<"Loaded in"<<timer.restart()<<"ms.********";

    /*********输出************/
    //write file
    /*
    qDebug()<<"Saving.";
    QXlsx::Document &excelBook2 = *new QXlsx::Document();
    for(int i =1;i<=row;i++){
        for(int j=1;j<=col;j++){
            QVariant _data = randomString(5);
            excelBook2.write(i,j,_data);
        }
    }
    qDebug()<<"Generated in"<<timer.restart()<<" ms.";
    excelBook2.saveAs("out.xlsx");
    qDebug()<<"Saved in"<<timer.restart()<<" ms.";
    */

    return a.exec();
}











