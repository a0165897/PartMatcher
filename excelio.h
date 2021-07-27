#pragma once
#ifndef EXCELIO_H
#define EXCELIO_H
#include "datastructure.h"
//Q stuff
#include <QDebug>
#include <QApplication>
#include <QTextCodec>
#include <QAxObject>
#include <QDir>
//QXlsx
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

#define EXCELIOBEGIN namespace excelio{
#define EXCELIOEND }

EXCELIOBEGIN
using cmpFunction = bool(*)(const QString &a,const QString &b);
using StringTable = QVector<QVector<QString>>;
const int _average = 1;
const int _min = 2;

struct cell{
    cell(uint _x=0, uint _y=0):row(_x),col(_y){}
    uint row;
    uint col;
};

class partData{
public:
    explicit partData(QString source){
        sourceExcelName = source;
        this->excelBook = new QXlsx::Document(source);
        this->initializeComInterface();

        this->initializeParts(PinWheelHousing);
        this->initializeParts(PlanetCarrier);
        this->initializeParts(CycloidGear);
        this->initializeParts(CrankShaft);
        this->initializeConfigs(Bearing);
        this->initializeConfigs(Config);
        this->postProcessCycloidGear(RoughCycloidGearList,CycloidGearList);

    }
    ~partData(){
            this->release();
    }
    void release(){
        if(COMExcel!=nullptr){
            COMExcel->dynamicCall("Quit()");
            delete COMExcel;
            COMExcel = nullptr;
        }
        delete excelBook;
        excelBook = nullptr;
    }
    bool saveTo(QVector<re>& from,QString& to);
    bool cleanSrc(QVector<re>& needToBeDeleted);
    /*数据*/
    QVector<pwh> PinWheelHousingList;//针齿壳
    QVector<pc> PlanetCarrierList;//行星架
    QVector<cg> CycloidGearList;//摆线轮
    QVector<cs> CrankShaftList;//曲轴

    /*标准件*/
//  np      NeedlePin;//针鞘
    tb      TaperedBearingConfig;//圆锥轴承
    acbb    AngularContactBallBearingConfig;//角接触球轴承
    cb      CageBearingConfig;//保持架轴承
    sm      ShimConfig;
    /*配置*/
    config configs;

private:
    QString sourceExcelName;

    QXlsx::Document* excelBook;//qxlsx接口

    QAxObject* COMExcel,* COMInterface; //COM接口（需要Excel环境）

    //摆线轮逐列读取的原始数据
    QVector<rcg> RoughCycloidGearList;

    /*各零件ID->行列号*/
    QMap<QString,cell> PinWheelHousingDic;
    QMap<QString,cell> PlanetCarrierDic;
    QMap<QString,cell> RoughCycloidGearDic;
    QMap<QString,cell> CrankShaftDic;

    /*初始化COM接口*/
    bool initializeComInterface();
    /*type映射向一个字串，在sheets的标题中寻找并返回带这个字串的标题*/
    QString whichSheetHaveThis(partType type);

    /*零件方法*/
    bool initializeParts(partType type);
    cell lookFor(const QString &word,cmpFunction compare,const int& row, const int& col, const int& rowLength, const int& colLength);
    double getPartValue(const QString& name,const QString& roughRow,const int& col,const int& way = _average);
    StringTable getParameterIdTable();
    std::tuple<cell,int> findPartListLocation(const QVector<QVector<QString>>& parameterId);
    /*配置文件方法*/
    bool initializeConfigs(partType type);
    range getConfigRange(const int& row, const int& col);
    QString getConfigParameterString(const int& row, const int& col);
    /*摆线轮后处理*/
    bool postProcessCycloidGear(const QVector<rcg>& RcgList, QVector<cg>& CgList);
    /*原始数据清理*/
    QMap<QString,cell>& idToCell(idType type);
    void colNumToColName(int data, QString &res);
    //单张sheet上所有待删除的cell
    bool mergeResults(idType type,QVector<re>& from,std::list<int>& colList,int& row);
    bool cleanSheet(idType type, std::list<int>& colList,int& row);
};

EXCELIOEND
#endif // EXCELIO_H
