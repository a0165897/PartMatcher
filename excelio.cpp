﻿#include "excelio.h"


QString partData::whichSheetHaveThis(QXlsx::Document &excelBook, partType type){
    auto sheetList = excelBook.sheetNames();//what excel have.
    auto target = partMap[type];//what we want.
    for(auto& i : sheetList){
        auto ret = i.indexOf(target);
        if(ret!= -1){
            return i;
        }
    }
    return "ERROR : Can't find this part in any sheet.";
}

/*
输入：excelbook，待查找字符串，比较函数，起始行号，起始列号，查找行数，查找列数
输出：待查找字符串的位置，如未找到返回[0,0]
*/
cell partData::lookFor(const QXlsx::Document & excelBook, const QString &word,cmpFunction compare,const int& row, const int& col, const int& rowLength, const int& colLength){
    for(int i=row;i<row+rowLength;i++){
        for(int j=col;j<col+colLength;j++){
            if(compare(excelBook.read(i,j).toString(),word)){
                qDebug()<<word<<"is in cell ["<<i<<","<<j<<"]";
                return (cell(i,j));
            }
        }
    }
    qDebug()<<"ERROR : Can't find "<< word <<" in "<<excelBook.currentSheet()->sheetName()<<".";
    return cell();
}

//从零件的行数栏(不同表示方法)->选取对应单元格数据->返回计算后的最终值
double partData::getPartValue(const QXlsx::Document& excelBook, const QString& name,const QString& roughRow,const int& col){
    int getComma = roughRow.indexOf(',');
    int getSlash = roughRow.indexOf('-');
    //未检查是否有其他无法识别的字符。
    if(getComma == -1 && getSlash == -1){
        int row = roughRow.toInt();
        return excelBook.read(row,col).toDouble();
    }else if(getSlash != -1){
        QStringList list = roughRow.split("-");
        int left = list[0].toInt();
        int mid = list[1].toInt();
        int right = list[2].toInt();
        double d = 0.0;
        int time = 0;
        for(int row = left;row<=mid;row+=right){
            d += excelBook.read(row,col).toDouble();
            time++;
        }
        return d/time;

    }else if(getComma != -1){
        int left = roughRow.leftRef(getComma).toInt();
        int right = roughRow.rightRef(roughRow.size()-(getComma+1)).toInt();
        double d1 = excelBook.read(left,col).toDouble();
        double d2 = excelBook.read(right,col).toDouble();
        const QString cbh1 = QStringLiteral("保持架轴承孔");
        const QString cbh2 = QStringLiteral("直径");
        if(name.indexOf(cbh1)!=-1&&name.indexOf(cbh2)!=-1){
            return fmin(d1,d2);
        }else{
            return 0.5 * (d1+d2);
        }

    }else{//不存在同时使用-和,的情况 
        return -1;
    }
    return -1;
}

//取得数据名称-数据行号矩阵
StringTable partData::getParameterIdTable(const QXlsx::Document& excelBook){
    //2.取得数据名称-数据行号矩阵头
    const QString anchor = QStringLiteral("选配所用数据名称");
    const int searchWidth = 20,searchHeight = 10;
    cell anchor_id = lookFor(excelBook,anchor,[](auto a,auto b){return a==b;},1,1,searchHeight,searchWidth);
    //3.取得数据名称-数据行号矩阵宽度
    cell parameterEnd = lookFor(excelBook,"",[](auto a,auto b){
        //找到空值或全空格值
        int size = 0;
        for(auto& i : a){if(i != ' ')size++;}
        return a==b||size==0;
    },anchor_id.row,anchor_id.col,1,searchWidth);
    int parameter_num = parameterEnd.col-1 - anchor_id.col;
    //4.读取数据名称-数据行号矩阵参数
    StringTable parameter_id(3,QVector<QString>(parameter_num));
    for(int i = 0;i<parameter_num;i++){
        parameter_id[0][i] = excelBook.read(anchor_id.row,anchor_id.col+1+i).toString();//参数的中文名称
        parameter_id[1][i] = excelBook.read(anchor_id.row+1,anchor_id.col+1+i).toString();//对应的行数n n,m n1-n2-k
        parameter_id[2][i] = excelBook.read(anchor_id.row+2,anchor_id.col+1+i).toString();//符号，但其实读不出来
    }
    //4.1中文字符纠正
    for(auto& i : parameter_id[1]){
        for(auto& j : i){
            if(j==QStringLiteral("，")) j = ',';
            if(j==QStringLiteral("—")) j = '-';
            if(j==QStringLiteral("。")) j = '.';
        }
    }
    for(int i = 0;i<parameter_num;i++){
        qDebug()<<parameter_id[0][i]<<" "<<parameter_id[1][i];
    }
    return parameter_id;
}

//取得零件数据的首位置及长度
std::tuple<cell,int> partData::findPartListLocation(const QXlsx::Document& excelBook,const QVector<QVector<QString>>& parameterId){
    int row = parameterId[1][0].toInt();
    int col = parameterId[2][0].toStdString()[0] - 'A' + 1;
    cell partBegin = cell(row,col);
    cell partEnd    = lookFor(excelBook,QLatin1String(""),[](auto a,auto b){
        //找到空值或全空格值
        int size = 0;
        for(auto& i : a){if(i != ' ')size++;}
        return a==b||size==0;
    },partBegin.row,partBegin.col,1,10000);
    int partNum = partEnd.col-partBegin.col;

    return std::tuple(partBegin,partNum);
}

//初始化某个部件
bool partData::initializeParts(QXlsx::Document& excelBook,partType type){
    QString sheetName = whichSheetHaveThis(excelBook,type);
    qDebug()<<"***************"+sheetName+"***************";
    //1.选择sheet
    excelBook.selectSheet(sheetName);
    //2.获取参数-行号列表
    StringTable parameterId = getParameterIdTable(excelBook);
    //3.找到零件的首位列序号
    auto [partBegin,partNum] = findPartListLocation(excelBook,parameterId);
    //4.读入数据
    switch(type){
        case PlanetCarrier:{
            this->PlanetCarrierList = QVector<pc>(partNum);
            for(int i=0;i<partNum;i++){
                 PlanetCarrierList[i].ID = excelBook.read(partBegin.row,partBegin.col+i).toString();
                 PlanetCarrierList[i].tbh_A1_d3 = getPartValue(excelBook,parameterId[0][1],parameterId[1][1],partBegin.col+i);
                 PlanetCarrierList[i].tbh_A2_d3 = getPartValue(excelBook,parameterId[0][2],parameterId[1][2],partBegin.col+i);
                 PlanetCarrierList[i].tbh_B1_d3 = getPartValue(excelBook,parameterId[0][3],parameterId[1][3],partBegin.col+i);
                 PlanetCarrierList[i].tbh_B2_d3 = getPartValue(excelBook,parameterId[0][4],parameterId[1][4],partBegin.col+i);
                 PlanetCarrierList[i].acbb_H2 = getPartValue(excelBook,parameterId[0][5],parameterId[1][5],partBegin.col+i);
                 PlanetCarrierList[i].ca_H1 = getPartValue(excelBook,parameterId[0][6],parameterId[1][6],partBegin.col+i);
            }
            break;
        }
        case PinWheelHousing:{
            this->PinWheelHousingList = QVector<pwh>(partNum);
            for(int i = 0;i<partNum;i++){
                PinWheelHousingList[i].ID = excelBook.read(partBegin.row,partBegin.col+i).toString();
                PinWheelHousingList[i].pwc_d1 = getPartValue(excelBook,parameterId[0][1],parameterId[1][1],partBegin.col+i);//名称，行，列
                PinWheelHousingList[i].pwcc_D2 = getPartValue(excelBook,parameterId[0][2],parameterId[1][2],partBegin.col+i);
                PinWheelHousingList[i].wa_h2 = getPartValue(excelBook,parameterId[0][3],parameterId[1][3],partBegin.col+i);
            }
            break;
        }
        case CycloidGear:{
            //rough
            qDebug()<<"MESSAGE : Rough CycloidGear data shown here is invisitable.";
            this->RoughCycloidGearList = QVector<rcg>(partNum);
            for(int i = 0;i<partNum;i++){
                RoughCycloidGearList[i].ID = excelBook.read(partBegin.row,partBegin.col+i).toString();
                RoughCycloidGearList[i].cg_Wk = getPartValue(excelBook,parameterId[0][1],parameterId[1][1],partBegin.col+i);
                RoughCycloidGearList[i].cbh_1_d5 = getPartValue(excelBook,parameterId[0][2],parameterId[1][2],partBegin.col+i);
                RoughCycloidGearList[i].cbh_2_d5 = getPartValue(excelBook,parameterId[0][3],parameterId[1][3],partBegin.col+i);
            }
            break;
        }
        case CrankShaft:{
            this->CrankShaftList = QVector<cs>(partNum);
            for(int i=0; i<partNum; i++){
                CrankShaftList[i].ID = excelBook.read(partBegin.row,partBegin.col+i).toString();
                CrankShaftList[i].ecc_h1 = getPartValue(excelBook,parameterId[0][1],parameterId[1][1],partBegin.col+i);
                CrankShaftList[i].ecc_A_D5 = getPartValue(excelBook,parameterId[0][2],parameterId[1][2],partBegin.col+i);
                CrankShaftList[i].ecc_B_D5 = getPartValue(excelBook,parameterId[0][3],parameterId[1][3],partBegin.col+i);
                CrankShaftList[i].cc_A_D4 = getPartValue(excelBook,parameterId[0][4],parameterId[1][4],partBegin.col+i);
                CrankShaftList[i].cc_B_D4 = getPartValue(excelBook,parameterId[0][5],parameterId[1][5],partBegin.col+i);
                CrankShaftList[i].phase_angle = getPartValue(excelBook,parameterId[0][6],parameterId[1][6],partBegin.col+i);
                CrankShaftList[i].ec_g = getPartValue(excelBook,parameterId[0][7],parameterId[1][7],partBegin.col+i);
            }
            break;
        }
        default:
            qDebug()<<"ERROR : Initializing undefined parts.";
            break;
    }
    return true;
}
QString partData::getConfigParameterString(QXlsx::Document &excelBook,const int& row, const int& col){
    auto in = excelBook.read(row,col).toString();
    for(auto& i: in){
        if(i == QStringLiteral("：")){
            i = ':';
        }
    }
    int getColon = in.indexOf(':');
    return in.right(in.size()-(getColon+1));
}

range partData::getConfigRange(QXlsx::Document &excelBook,const int& row, const int& col){
    QString roughRange = excelBook.read(row,col).toString();
    for(auto& i: roughRange){
        if(i == 248){//248 is the id of fxxking 'ø' in Unicode, which seems cannot be wrapped by QStringLiteral.
            i=' ';
        }
        if(i == QStringLiteral("。")){
            i='.';
        }
    }
    int getWave = roughRange.indexOf('~');
    int getPlusMinus = roughRange.indexOf(QStringLiteral("±"));
    if(getPlusMinus!=-1){
        double value = roughRange.rightRef(roughRange.size()-(getPlusMinus+1)).toDouble();
        return range(-value,value);
    }else if(getWave!= -1){
        double left = roughRange.leftRef(getWave).toDouble();
        double right = roughRange.rightRef(roughRange.size()-(getWave+1)).toDouble();
        return range(left,right);
    }else{
        qDebug()<<"Warning : Reading a blank config range or unknown cell in ["<<row<<","<<col<<"], which will be ignored.";
        return range(10001,10001);
    }

}

//初始化配置
bool partData::initializeConfigs(QXlsx::Document &excelBook,partType type){
    QString sheetName = whichSheetHaveThis(excelBook,type);
    qDebug()<<"***************"+sheetName+"***************";
    //1.选择sheet
    excelBook.selectSheet(sheetName);
    auto cmp = [](const QString &a,const QString &b)->bool{
        return a==b;
    };
    switch(type){
        case Bearing:{
            QStringList titles = {QStringLiteral("选取范围"),
                                  QStringLiteral("设计尺寸"),
                                  QStringLiteral("圆锥轴承外径"),
                                  QStringLiteral("圆锥轴承内径"),
                                  QStringLiteral("圆锥轴承高度"),
                                  QStringLiteral("保持架轴承针销直径"),
                                  QStringLiteral("角接触球轴承高度"),

            };
            QVector<cell> cellList(titles.size());
            for(int i=0;i<titles.size();i++){
                cellList[i] = lookFor(excelBook,titles[i],cmp,1,1,20,10);
            }
            /*整理序号
                    ... cola colb ...
            rows[0]
            rows[1]*/
            int cola = cellList[0].col,colb = cellList[1].col;
            QVector<int> rows(cellList.size()-2);
            for(int i=2;i<cellList.size();i++){
                rows[i-2] = cellList[i].row;
            }
            //圆锥轴承外径&设计尺寸
            for(int i=0;i<3;i++){
                range ret = getConfigRange(excelBook,rows[0]+i,cola);
                if(ret.min<10000.0) TaperedBearing.tb_od_range.push_back(ret);
            }
            TaperedBearing.tb_od = excelBook.read(rows[0],colb).toDouble();
            //圆锥轴承内径&设计尺寸
            for(int i=0;i<3;i++){
                range ret = getConfigRange(excelBook,rows[1]+i,cola);
                if(ret.min<10000.0) TaperedBearing.tb_id_range.push_back(ret);
            }
            TaperedBearing.tb_id = excelBook.read(rows[1],colb).toDouble();
            //圆锥轴承高度&设计尺寸
            for(int i=0;i<3;i++){
                range ret = getConfigRange(excelBook,rows[2]+i,cola);
                if(ret.min<10000.0) TaperedBearing.tb_h_range.push_back(ret);
            }
            TaperedBearing.tb_h = excelBook.read(rows[2],colb).toDouble();
            //保持架轴承针销直径&设计尺寸
            CageBearing.cb_d_range = getConfigRange(excelBook,rows[3],cola);
            CageBearing.cb_d = excelBook.read(rows[3],colb).toDouble();
            //角接触球轴承高度&设计尺寸
            AngularContactBallBearing.acbb_h_range = getConfigRange(excelBook,rows[4],cola);
            AngularContactBallBearing.acbb_h = excelBook.read(rows[4],colb).toDouble();
            break;
    }
        case Config:{
            QStringList titlesCol = {
                                  QStringLiteral("零件的尺寸公差或范围"),
                                  QStringLiteral("公式参数1"),
                                  QStringLiteral("公式参数2"),
                                  QStringLiteral("公式参数3")
            };
            QStringList titlesRow = {
                                  QStringLiteral("针齿壳与针销"),
                                  QStringLiteral("针齿壳与摆线轮"),
                                  QStringLiteral("摆线轮轴承孔、曲轴偏心圆与保持架轴承针销间隙"),
                                  QStringLiteral("摆线轮、保持架轴承与曲轴的相位角匹配"),
                                  QStringLiteral("曲轴与锥形轴承"),
                                  QStringLiteral("行星架与锥形轴承"),
                                  QStringLiteral("行星架、角接触球轴承、针齿壳齿槽高度配合与角接触球轴承预紧量t2计算公式"),
                                  QStringLiteral("曲轴两个偏心圆柱与行星架卡簧槽高度配合，曲轴预紧量t1计算公式")
                                 };

            /*整理序号
                    ... col[0] col[1] ...
            rows[0]
            rows[1]*/
            QVector<cell> cellListRow(titlesRow.size());
            for(int i=0;i<titlesRow.size();i++){
                cellListRow[i] = lookFor(excelBook,titlesRow[i],cmp,1,1,30,10);
            }
            QVector<int> rows(cellListRow.size());
            for(int i=0;i<cellListRow.size();i++){
                rows[i] = cellListRow[i].row;
            }

            QVector<cell> cellListCol(titlesCol.size());
            for(int i=0;i<titlesCol.size();i++){
                cellListCol[i] = lookFor(excelBook,titlesCol[i],cmp,1,1,30,10);
            }
            QVector<int> cols(cellListCol.size());
            for(int i=0;i<cellListCol.size();i++){
                cols[i] = cellListCol[i].col;
            }
            //针齿圆d1范围
            for(int i=0;i<4;i++){
                range ret = getConfigRange(excelBook,rows[0]+1+i,cols[0]);
                if(ret.min<10000.0) configs.pwc_d1_range.push_back(ret);
            }
            //针销直径D1范围
            for(int i=0;i<4;i++){
                range ret = getConfigRange(excelBook,rows[0]+1+i,cols[0]+1);
                if(ret.min<10000.0) configs.np_D1_range.push_back(ret);
            }
            //针齿中心圆D2范围
            for(int i=0;i<4;i++){
                range ret = getConfigRange(excelBook,rows[1]+1+i,cols[0]);
                if(ret.min<10000.0) configs.pwcc_D2_range.push_back(ret);
            }
            //摆线轮公法线范围
            for(int i=0;i<4;i++){
                range ret = getConfigRange(excelBook,rows[1]+1+i,cols[0]+1);
                if(ret.min<10000.0) configs.cg_Wk_range.push_back(ret);
            }
            //摆线轮轴承孔、曲轴偏心圆与保持架轴承针销间隙
            configs.t5_range = getConfigRange(excelBook,rows[2],cols[0]);
            //相位角匹配
            configs.phase_angle = getConfigRange(excelBook,rows[3],cols[0]);
            //曲轴相位角的选配 0-平均 1-同时满足
            configs.phase_flag = getConfigParameterString(excelBook,rows[3],cols[1]).toInt();
            //理论侧隙：0.02
            configs.delta_c = getConfigParameterString(excelBook,rows[3],cols[2]).toDouble();
            //相位角补偿值：0
            configs.t6 = getConfigParameterString(excelBook,rows[3],cols[3]).toDouble();
            //圆锥轴承内径d4与曲轴中心圆柱直径D4之差的范围
            configs.t4_range = getConfigRange(excelBook,rows[4],cols[0]);
            //圆锥轴承孔直径d3与圆锥轴承外径D3之差的范围
            configs.t3_range = getConfigRange(excelBook,rows[5],cols[0]);
            //齿槽高h2公差范围
            for(int i=0;i<2;i++){
                range ret = getConfigRange(excelBook,rows[6]+1+i,cols[0]);
                if(ret.min<10000.0) configs.wa_h2_range.push_back(ret);
            }
            //行星架角接触球轴承高度H2公差范围
            for(int i=0;i<2;i++){
                range ret = getConfigRange(excelBook,rows[6]+1+i,cols[0]+1);
                if(ret.min<10000.0) configs.acbb_H2_range.push_back(ret);
            }
            //针齿壳齿槽高与行星架角接触球轴承高度配合
            configs.t2_range = getConfigRange(excelBook,rows[6]+3,cols[0]);
            //行星架卡簧槽高H1公差范围
            for(int i=0;i<2;i++){
                range ret = getConfigRange(excelBook,rows[6]+5+i,cols[0]);
                if(ret.min<10000.0) configs.ca_H1_range.push_back(ret);
            }
            //两个偏心圆柱高度h1公差范围
            for(int i=0;i<2;i++){
                range ret = getConfigRange(excelBook,rows[6]+5+i,cols[0]+1);
                if(ret.min<10000.0) configs.ecc_h1_range.push_back(ret);
            }
            //行星架卡簧槽高与曲轴两个偏心圆柱高度配合
            configs.t1_range = getConfigRange(excelBook,rows[6]+7,cols[0]);
            break;
        }
        default:
            qDebug()<<"ERROR : Initializing undefined configure.";
            break;
        }

    return true;
}

bool partData::postProcessCycloidGear(const QVector<rcg>& RcgList, QVector<cg>& CgList){
    qDebug()<<"***************"<<"postProcessing CycloidGear"<<"******************";
    auto distance = [](const QString& a, const QString& b)->int{
        int d = 0;
        int size = fmin(a.size(),b.size());
        for(int i = 0;i<size;i++){
            if(a[i]!=b[i])d++;
        }
        d += abs(a.size() - b.size());
        return d;
    };
    CgList = QVector<cg>(RcgList.size()/2);
    if(CgList.size()*2!=RcgList.size()){
        qDebug()<<"WARNING : Size of CycloidGear parts is"<<RcgList.size()<<", double check the data.";
    }
    for(int i=0;i<CgList.size();i++){
        if(distance(RcgList[2*i].ID,RcgList[2*i+1].ID)!=1){
            qDebug()<<"Error : Data near " +RcgList[2*i].ID+ " is missing, double check the data";
        }
        CgList[i].ID = RcgList[2*i].ID;
        CgList[i].cg_Wk = 0.5 * (RcgList[2*i].cg_Wk + RcgList[2*i+1].cg_Wk);
        CgList[i].cbh_A1_d5 = RcgList[2*i].cbh_1_d5;
        CgList[i].cbh_A2_d5 = RcgList[2*i].cbh_2_d5;
        CgList[i].cbh_B1_d5 = RcgList[2*i+1].cbh_1_d5;
        CgList[i].cbh_B2_d5 = RcgList[2*i+1].cbh_2_d5;
    }
    qDebug()<<"Size of rough data is:"<<RcgList.size()<< ".";
    qDebug()<<"Size of merged data is:"<<CgList.size()<< ".";
    return true;
}