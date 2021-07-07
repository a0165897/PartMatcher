#pragma once
#ifndef DATASTRUCTURE_H
#define DATASTRUCTURE_H
#include <QApplication>
/***所有零件相关的结构体、配置相关的结构体定义***/

enum partType{PinWheelHousing,
              PlanetCarrier,
              CycloidGear,
              CrankShaft,
              Bearing,
              Config
             };

const QMap<partType,QString> partMap = {
    {PinWheelHousing,   QStringLiteral("针齿壳")},
    {CycloidGear,       QStringLiteral("摆线轮")},
    {CrankShaft,        QStringLiteral("曲轴")},
    {PlanetCarrier,     QStringLiteral("行星架")},
    {Bearing,           QStringLiteral("标准件")},
    {Config,            QStringLiteral("编程规则")}
};

typedef struct RangePair{
    RangePair(double _min=0.0, double _max=0.0):min(_min),max(_max){}
    double min;
    double max;
}range;

using rangeList = QVector<range>;

typedef struct PinWheelHousing  //针齿壳
{
    QString ID;     //零件编号
    double pwc_d1;   //针齿圆直径
    double pwcc_D2;  //针齿中心圆直径
    double wa_h2;    //齿槽(alveolar)高
}pwh;

typedef struct NeedlePin  //针销 where is this??
{
    QString ID;
    double np_D1;   //针销直径
}np;

typedef struct RoughCycloidGear  //摆线轮
{
    QString ID;
    double cg_Wk;   //摆线轮公法线
    double cbh_1_d5;    //保持架轴承孔cage bearing hole 直径d
    double cbh_2_d5;    //保持架轴承孔cage bearing hole 直径d
}rcg;

typedef struct CycloidGear  //摆线轮  摆线轮会有AB两个作为一对，A为上轮，B为下轮
{
 QString ID;
 double cg_Wk;        //摆线轮公法线
 double cbh_A1_d5;    //保持架轴承孔cage bearing hole 直径d
 double cbh_A2_d5;    //保持架轴承孔cage bearing hole 直径d
 double cbh_B1_d5;
 double cbh_B2_d5;
}cg;

typedef struct CrankShaft  //曲轴
{
    QString ID;
    double ecc_h1;  //偏心圆柱eccentric circular cylinders 高度
    double ecc_A_D5;  //偏心圆柱eccentric circular cylinders 直径
    double ecc_B_D5;
    double cc_A_D4;   //中心圆柱直径
    double cc_B_D4;   //中心圆柱直径
    double phase_angle;  //相位角
    double ec_g;   //偏心距gap

}cs;

typedef struct PlanetCarrier  //行星架
{
    QString ID;
    double tbh_A1_d3;  //圆锥轴承孔tapered bearing hole 直径
    double tbh_A2_d3;
    double tbh_B1_d3;
    double tbh_B2_d3;
    //double ecc_d;  //偏心圆柱eccentric circular cylinders 直径
    double acbb_H2;   //角接触球轴承angular-contact ball bearing 配合高度
    double ca_H1;   //卡簧槽Circlip alveolar 高度
}pc;

/*************标准件尺寸报告***************/
typedef struct TaperedBearing  //圆锥轴承
{
    //QString ID;
    double tb_od;  //设计尺寸
    double tb_id;
    double tb_h;
    rangeList tb_od_range;  //圆锥轴承外径
    rangeList tb_id_range;  //圆锥轴承内径
    rangeList tb_h_range;   //圆锥轴承高度
}tb;

typedef struct AngularContactBallBearing  //角接触球轴承
{
    //QString ID;
    double acbb_h;  //角接触球轴承高度
    range acbb_h_range;
}acbb;

typedef struct CageBearing  //保持架轴承
{
    //QString ID;
    double cb_d;  //保持架轴承直径
    range cb_d_range;
}cb;

/*******编程规则***********/
typedef struct configuration //公式中范围的配置文件
{
    //int range_num1;     //范围1总数
    rangeList pwc_d1_range;  //针齿圆d1范围
    rangeList np_D1_range;   //针销直径D1范围

    //int range_num2;     //范围2总数
    rangeList pwcc_D2_range; //针齿中心圆D2范围
    rangeList cg_Wk_range;   //摆线轮公法线范围

    //int range_num3;       //齿槽高h2公差范围总数
    rangeList wa_h2_range;    //齿槽高h2公差范围
    rangeList acbb_H2_range; //行星架角接触球轴承高度H2公差范围

    //int range_num4;
    rangeList ca_H1_range;     //行星架卡簧槽高H1公差范围
    rangeList ecc_h1_range;    //两个偏心圆柱高度h1公差范围

    range phase_angle;   //相位角匹配

    range t5_range;   //摆线轮轴承孔、曲轴偏心圆与保持架轴承针销间隙
    range t4_range;   //圆锥轴承内径d4与曲轴中心圆柱直径D4之差的范围
    range t3_range;   //圆锥轴承孔直径d3与圆锥轴承外径D3之差的范围
    range t2_range;   //针齿壳齿槽高与行星架角接触球轴承高度配合
    range t1_range;   //行星架卡簧槽高与曲轴两个偏心圆柱高度配合

    int phase_flag;//曲轴相位角的选配 0-平均 1-同时满足
    double delta_c;//理论侧隙
    double t6;//相位角补偿值

    double pwc_d1_dimension;  //针齿圆直径尺寸PinWheelHousing
    double pwcc_D2_dimension;  //针齿中心圆尺寸PinWheelHousing
    double wa_h2_dimension;   //齿槽高尺寸PinWheelHousing
    double acbb_H2_dimension;  //行星架卡簧槽高度PlanetCarrier
    double ca_H1_dimension;   //行星架角接触球轴承高度尺寸PlanetCarrier
    double ecc_h1_dimension;  //两个偏心圆柱高度h1尺寸CrankShaft
    double cg_Wk_dimension;   //摆线轮公法线尺寸CycloidGear

}config;




#endif // DATASTRUCTURE_H
