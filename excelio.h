#pragma once
#ifndef EXCELIO_H
#define EXCELIO_H
#include "datastructure.h"
//Q stuff
#include <QDebug>
#include <QApplication>
#include <QTextCodec>
//QXlsx
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

using cmpFunction = bool(*)(const QString &a,const QString &b);
using StringTable = QVector<QVector<QString>>;
struct cell{
    cell(uint _x=0, uint _y=0):row(_x),col(_y){}
    uint row;
    uint col;
};

class partData{
public:
    explicit partData(QXlsx::Document & excelBook){
        this->initializeParts(excelBook,PinWheelHousing);
        this->initializeParts(excelBook,PlanetCarrier);
        this->initializeParts(excelBook,CycloidGear);
        this->initializeParts(excelBook,CrankShaft);
        this->initializeConfigs(excelBook,Bearing);
        this->initializeConfigs(excelBook,Config);
        this->postProcessCycloidGear(RoughCycloidGearList,CycloidGearList);
    }

    /*数据*/
    QVector<pwh> PinWheelHousingList;//针齿壳
    QVector<pc> PlanetCarrierList;//行星架
    QVector<cg> CycloidGearList;//摆线轮
    QVector<cs> CrankShaftList;//曲轴
    /*标准件*/
//  np      NeedlePin;//针鞘
    tb      TaperedBearing;//圆锥轴承
    acbb    AngularContactBallBearing;//角接触球轴承
    cb      CageBearing;//保持架轴承
    /*配置*/
    config configs;

private:
    QVector<rcg> RoughCycloidGearList; //摆线轮逐列读取的原始数据

    /*type映射向一个字串，在sheets的标题中寻找并返回带这个字串的标题*/
    QString whichSheetHaveThis(QXlsx::Document& excelBook,partType type);

    /*零件方法*/
    bool initializeParts(QXlsx::Document& excelBook,partType type);
    cell lookFor(const QXlsx::Document& excelBook, const QString &word,cmpFunction compare,const int& row, const int& col, const int& rowLength, const int& colLength);
    double getPartValue(const QXlsx::Document& excelBook, const QString& name,const QString& roughRow,const int& col);
    StringTable getParameterIdTable(const QXlsx::Document& excelBook);
    std::tuple<cell,int> findPartListLocation(const QXlsx::Document& excelBook,const QVector<QVector<QString>>& parameterId);
    /*配置文件方法*/
    bool initializeConfigs(QXlsx::Document& excelBook,partType type);
    range getConfigRange(QXlsx::Document &excelBook,const int& row, const int& col);
    QString getConfigParameterString(QXlsx::Document &excelBook,const int& row, const int& col);
    /*摆线轮后处理*/
    bool postProcessCycloidGear(const QVector<rcg>& RcgList, QVector<cg>& CgList);
};


#endif // EXCELIO_H
