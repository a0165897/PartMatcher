
#include <QApplication>
#include <QDebug>
#include <QElapsedTimer>
#include <QAxObject>
#include <QDir>
//stl
#include <tuple>
//excelio
#include "excelio.h"
//import window
#include "mainwindow.h"
//test ssh
int main(int argc, char *argv[])
{
    QApplication a(argc, argv);

    QElapsedTimer timer;
    timer.start();

    //原始文件
    QString fileIn = QStringLiteral("D:/build-partMatcher-Desktop_Qt_5_15_2_MSVC2019_64bit-Debug/20E间隙计算测量数据20210707(1).xlsx");
    //匹配结果
    QString fileOut = QStringLiteral("D:/build-partMatcher-Desktop_Qt_5_15_2_MSVC2019_64bit-Debug/零件选配结果210720.xlsx");
    //打开数据
    excelio::partData data(fileIn);
    qDebug()<<"********file: "<<fileIn<<"Loaded in"<<timer.restart()<<"ms.********";

    //进行计算
    QVector<re> ret;
    re sample;
    sample.pwc_ID = data.PinWheelHousingList[22].ID;
    sample.cg_A_ID = data.CycloidGearList[0].ID;
    sample.cg_B_ID = data.CycloidGearList[0].ID_vice;
    sample.cs_1_ID = data.CrankShaftList[0].ID;
    sample.cs_2_ID = data.CrankShaftList[1].ID;
    sample.pc_ID = data.PlanetCarrierList[0].ID;
    sample.tb_A1_id = data.TaperedBearingConfig.tb_id_range[0];
    sample.tb_A2_id = data.TaperedBearingConfig.tb_id_range[1];
    sample.tb_B1_id = data.TaperedBearingConfig.tb_id_range[2];
    //sample.tb_B2_id = data.TaperedBearingConfig.tb_id_range[3];
    sample.tb_A1_od = data.TaperedBearingConfig.tb_od_range[0];
    sample.tb_A2_od = data.TaperedBearingConfig.tb_od_range[1];
    sample.tb_B1_od = data.TaperedBearingConfig.tb_od_range[2];
    //sample.tb_B2_od = data.TaperedBearingConfig.tb_od_range[3];
    sample.tb_A1_h = data.TaperedBearingConfig.tb_h_range[0];
    sample.tb_A2_h = data.TaperedBearingConfig.tb_h_range[0];
    sample.tb_B1_h = data.TaperedBearingConfig.tb_h_range[0];
    //sample.tb_B2_h = data.TaperedBearingConfig.tb_h_range[3];
    sample.cb_A1_d = data.CageBearingConfig.cb_d_range;
    sample.cb_A2_d = data.CageBearingConfig.cb_d_range;
    sample.cb_B1_d = data.CageBearingConfig.cb_d_range;
    sample.cb_B2_d = data.CageBearingConfig.cb_d_range;
    sample.acbb_h = data.AngularContactBallBearingConfig.acbb_h_range;
    sample.shim_1 = range(data.ShimConfig.shim,data.ShimConfig.shim);
    sample.shim_2 = range(data.ShimConfig.shim,data.ShimConfig.shim);
    sample.np = range(0.01,0.02);
    for(int i=0;i<15;i++){
        ret.push_back(sample);
    }

    //保存结果
    qDebug()<<"Saving to "<< fileOut<<".";
    data.saveTo(ret,fileOut);
    qDebug()<<"********Result Saved in"<<timer.restart()<<" ms.********";
    data.cleanSrc(ret);
    qDebug()<<"******** Cleaned in"<<timer.restart()<<" ms.********";
    data.release();
    return a.exec();
}











