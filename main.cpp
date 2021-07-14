
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
    //打开数据
    const QString fileIn = QStringLiteral("20E间隙计算测量数据20210707(1).xlsx");
    QXlsx::Document excelBook =  QXlsx::Document(fileIn);
    excelio::partData data(excelBook);
    qDebug()<<"********file: "<<fileIn<<"Loaded in"<<timer.restart()<<"ms.********";

    //进行计算
    QVector<re> ret;
    re sample;
    sample.pwc_ID = data.PinWheelHousingList[0].ID;
    sample.cg_A_ID = data.CycloidGearList[0].ID;
    sample.cg_B_ID = data.CycloidGearList[1].ID;
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
    for(int i=0;i<15;i++){
        ret.push_back(sample);
    }

    //保存结果
    const QString fileOut = QStringLiteral("零件选配结果210629(1).xlsx");
    QXlsx::Document result = QXlsx::Document(fileOut);
    qDebug()<<"Saving to "<< fileOut<<".";
    data.saveTo(ret,result);
    qDebug()<<"******** Saved in"<<timer.restart()<<" ms.********";


    return a.exec();
}











