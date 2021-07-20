#include "excelio.h"

EXCELIOBEGIN

bool partData::initializeComInterface(){
    COMExcel = new QAxObject("Excel.Application");
    COMExcel->setProperty("Visible", false);
    COMInterface = COMExcel->querySubObject("WorkBooks")->querySubObject("Open (const QString&)", QDir::toNativeSeparators(sourceExcelName));
    return true;
}

QString partData::whichSheetHaveThis(partType type){
    auto sheetList = excelBook->sheetNames();//what excel have.
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
cell partData::lookFor(const QString &word,cmpFunction compare,const int& row, const int& col, const int& rowLength, const int& colLength){
    for(int i=row;i<row+rowLength;i++){
        for(int j=col;j<col+colLength;j++){
            if(compare(excelBook->read(i,j).toString(),word)){
                qDebug()<<word<<"is in cell ["<<i<<","<<j<<"]";
                return (cell(i,j));
            }
        }
    }
    qDebug()<<"ERROR : Can't find "<< word <<" in "<<excelBook->currentSheet()->sheetName()<<".";
    return cell();
}

//从零件的行数栏(不同表示方法)->选取对应单元格数据->返回计算后的最终值
double partData::getPartValue(const QString& name,const QString& roughRow,const int& col,const int& way){
    int getComma = roughRow.indexOf(',');
    int getSlash = roughRow.indexOf('-');
    //未检查是否有其他无法识别的字符。
    if(getComma == -1 && getSlash == -1){
        int row = roughRow.toInt();
        return excelBook->read(row,col).toDouble();
    }else if(getSlash != -1){
        QStringList list = roughRow.split("-");
        int left = list[0].toInt();
        int mid = list[1].toInt();
        int right = list[2].toInt();
        double d = 0.0;
        int time = 0;
        for(int row = left;row<=mid;row+=right){
            d += excelBook->read(row,col).toDouble();
            time++;
        }
        return d/time;

    }else if(getComma != -1){
        int left = roughRow.leftRef(getComma).toInt();
        int right = roughRow.rightRef(roughRow.size()-(getComma+1)).toInt();
        double d1 = excelBook->read(left,col).toDouble();
        double d2 = excelBook->read(right,col).toDouble();
        if(way == _min){
            return fmin(d1,d2);
        }else if(way == _average){
            return 0.5 * (d1+d2);
        }else{
            return -1;
        }
    }else{//不存在同时使用-和,的情况 
        return -1;
    }
    return -1;
}

//取得数据名称-数据行号矩阵
StringTable partData::getParameterIdTable(){
    //2.取得数据名称-数据行号矩阵头
    const QString anchor = QStringLiteral("选配所用数据名称");
    const int searchWidth = 20,searchHeight = 10;
    cell anchor_id = lookFor(anchor,[](auto a,auto b){return a==b;},1,1,searchHeight,searchWidth);
    //3.取得数据名称-数据行号矩阵宽度
    cell parameterEnd = lookFor("",[](auto a,auto b){
        //找到空值或全空格值
        int size = 0;
        for(auto& i : a){if(i != ' ')size++;}
        return a==b||size==0;
    },anchor_id.row,anchor_id.col,1,searchWidth);
    int parameter_num = parameterEnd.col-1 - anchor_id.col;
    //4.读取数据名称-数据行号矩阵参数
    StringTable parameter_id(3,QVector<QString>(parameter_num));
    for(int i = 0;i<parameter_num;i++){
        parameter_id[0][i] = excelBook->read(anchor_id.row,anchor_id.col+1+i).toString();//参数的中文名称
        parameter_id[1][i] = excelBook->read(anchor_id.row+1,anchor_id.col+1+i).toString();//对应的行数n n,m n1-n2-k
        parameter_id[2][i] = excelBook->read(anchor_id.row+2,anchor_id.col+1+i).toString();//符号，但其实读不出来
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
std::tuple<cell,int> partData::findPartListLocation(const QVector<QVector<QString>>& parameterId){
    int row = parameterId[1][0].toInt();
    int col = parameterId[2][0].toStdString()[0] - 'A' + 1;
    cell partBegin = cell(row,col);
    cell partEnd    = lookFor(QLatin1String(""),[](auto a,auto b){
        //找到空值或全空格值
        int size = 0;
        for(auto& i : a){if(i != ' ')size++;}
        return a==b||size==0;
    },partBegin.row,partBegin.col,1,10000);
    int partNum = partEnd.col-partBegin.col;

    return std::tuple(partBegin,partNum);
}

//初始化某个部件
bool partData::initializeParts(partType type){
    QString sheetName = whichSheetHaveThis(type);
    qDebug()<<"***************"+sheetName+"***************";
    //1.选择sheet
    excelBook->selectSheet(sheetName);
    //2.获取参数-行号列表
    StringTable parameterId = getParameterIdTable();
    //3.找到零件的首位列序号
    auto [partBegin,partNum] = findPartListLocation(parameterId);
    //4.找到config里的列
    int dimensionCol = lookFor(QStringLiteral("Dimension尺寸"),[](auto a,auto b){return a==b;},1,1,20,10).col;
    //5.读入数据
    switch(type){
        case PlanetCarrier:{
            this->PlanetCarrierList = QVector<pc>(partNum);
            for(int i=0;i<partNum;i++){
                 PlanetCarrierList[i].ID = excelBook->read(partBegin.row,partBegin.col+i).toString();
                 PlanetCarrierList[i].tbh_A1_d3 = getPartValue(parameterId[0][1],parameterId[1][1],partBegin.col+i);
                 PlanetCarrierList[i].tbh_A2_d3 = getPartValue(parameterId[0][2],parameterId[1][2],partBegin.col+i);
                 PlanetCarrierList[i].tbh_B1_d3 = getPartValue(parameterId[0][3],parameterId[1][3],partBegin.col+i);
                 PlanetCarrierList[i].tbh_B2_d3 = getPartValue(parameterId[0][4],parameterId[1][4],partBegin.col+i);
                 PlanetCarrierList[i].acbb_H2 = getPartValue(parameterId[0][5],parameterId[1][5],partBegin.col+i);
                 PlanetCarrierList[i].ca_H1 = getPartValue(parameterId[0][6],parameterId[1][6],partBegin.col+i);
                 PlanetCarrierDic.insert(PlanetCarrierList[i].ID,cell(partBegin.row,partBegin.col+i));
            }
            configs.acbb_H2_dimension = getPartValue(parameterId[0][5],parameterId[1][5],dimensionCol);
            configs.ca_H1_dimension = getPartValue(parameterId[0][6],parameterId[1][6],dimensionCol);
            break;
        }
        case PinWheelHousing:{
            this->PinWheelHousingList = QVector<pwh>(partNum);
            for(int i = 0;i<partNum;i++){
                PinWheelHousingList[i].ID = excelBook->read(partBegin.row,partBegin.col+i).toString();
                PinWheelHousingList[i].pwc_d1 = getPartValue(parameterId[0][1],parameterId[1][1],partBegin.col+i);//名称，行，列
                PinWheelHousingList[i].pwcc_D2 = getPartValue(parameterId[0][2],parameterId[1][2],partBegin.col+i);
                PinWheelHousingList[i].wa_h2 = getPartValue(parameterId[0][3],parameterId[1][3],partBegin.col+i);
                PinWheelHousingDic.insert(PinWheelHousingList[i].ID,cell(partBegin.row,partBegin.col+i));
            }
            configs.pwc_d1_dimension = getPartValue(parameterId[0][1],parameterId[1][1],dimensionCol);
            configs.pwcc_D2_dimension = getPartValue(parameterId[0][2],parameterId[1][2],dimensionCol);
            configs.wa_h2_dimension =  getPartValue(parameterId[0][3],parameterId[1][3],dimensionCol);
            break;
        }
        case CycloidGear:{
            //rough
            qDebug()<<"MESSAGE : Rough CycloidGear data shown here is invisitable.";
            this->RoughCycloidGearList = QVector<rcg>(partNum);
            for(int i = 0;i<partNum;i++){
                RoughCycloidGearList[i].ID = excelBook->read(partBegin.row,partBegin.col+i).toString();
                RoughCycloidGearList[i].cg_Wk = getPartValue(parameterId[0][1],parameterId[1][1],partBegin.col+i);
                RoughCycloidGearList[i].cbh_1_d5 = getPartValue(parameterId[0][2],parameterId[1][2],partBegin.col+i,_min);
                RoughCycloidGearList[i].cbh_2_d5 = getPartValue(parameterId[0][3],parameterId[1][3],partBegin.col+i,_min);
                RoughCycloidGearDic.insert(RoughCycloidGearList[i].ID,cell(partBegin.row,partBegin.col+i));
            }
            configs.cg_Wk_dimension = getPartValue(parameterId[0][1],parameterId[1][1],dimensionCol);
            break;
        }
        case CrankShaft:{
            this->CrankShaftList = QVector<cs>(partNum);
            for(int i=0; i<partNum; i++){
                CrankShaftList[i].ID = excelBook->read(partBegin.row,partBegin.col+i).toString();
                CrankShaftList[i].ecc_h1 = getPartValue(parameterId[0][1],parameterId[1][1],partBegin.col+i);
                CrankShaftList[i].ecc_A_D5 = getPartValue(parameterId[0][2],parameterId[1][2],partBegin.col+i);
                CrankShaftList[i].ecc_B_D5 = getPartValue(parameterId[0][3],parameterId[1][3],partBegin.col+i);
                CrankShaftList[i].cc_A_D4 = getPartValue(parameterId[0][4],parameterId[1][4],partBegin.col+i);
                CrankShaftList[i].cc_B_D4 = getPartValue(parameterId[0][5],parameterId[1][5],partBegin.col+i);
                CrankShaftList[i].phase_angle = getPartValue(parameterId[0][6],parameterId[1][6],partBegin.col+i);
                CrankShaftList[i].ec_g = getPartValue(parameterId[0][7],parameterId[1][7],partBegin.col+i);
                CrankShaftDic.insert(CrankShaftList[i].ID,cell(partBegin.row,partBegin.col+i));
            }
            configs.ecc_h1_dimension = getPartValue(parameterId[0][1],parameterId[1][1],dimensionCol);
            break;
        }
        default:
            qDebug()<<"ERROR : Initializing undefined parts.";
            break;
    }
    return true;
}
QString partData::getConfigParameterString(const int& row, const int& col){
    auto in = excelBook->read(row,col).toString();
    for(auto& i: in){
        if(i == QStringLiteral("：")){
            i = ':';
        }
    }
    int getColon = in.indexOf(':');
    return in.right(in.size()-(getColon+1));
}

range partData::getConfigRange(const int& row, const int& col){
    QString roughRange = excelBook->read(row,col).toString();
    for(auto& i: roughRange){
        if(i == 248){//248 is the id of fxxking 'ø' in Unicode, which seems cannot be wrapped by QStringLiteral.
            i=' ';
        }
        if(i == QStringLiteral("。")){
            i='.';
        }
    }
    roughRange.remove(QRegExp("\\s"));
    int getWave = roughRange.indexOf('~');
    int getPlusMinus = roughRange.indexOf(QStringLiteral("±"));
    if(getPlusMinus!=-1){
        double value = roughRange.rightRef(roughRange.size()-(getPlusMinus+1)).toDouble();
        return range(-value,value);
    }else if(getWave!= -1){
        double left = roughRange.leftRef(getWave).toDouble();
        double right = roughRange.rightRef(roughRange.size()-(getWave+1)).toDouble();
        return range(fmin(left,right),fmax(left,right));
    }else{
        qDebug()<<"Warning : Reading a blank config range or unknown cell in ["<<row<<","<<col<<"], which will be ignored.";
        return range(10001,10001);
    }

}

//初始化配置
bool partData::initializeConfigs(partType type){
    QString sheetName = whichSheetHaveThis(type);
    qDebug()<<"***************"+sheetName+"***************";
    //1.选择sheet
    excelBook->selectSheet(sheetName);
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
                                  QStringLiteral("圆锥轴承垫片厚度")

            };
            QVector<cell> cellList(titles.size());
            for(int i=0;i<titles.size();i++){
                cellList[i] = lookFor(titles[i],cmp,1,1,20,10);
            }
            /*整理序号
                    ... cola colb ...
            rows[0]
            rows[1]*/
            int cola = cellList[0].col;
            int colb = cellList[1].col;
            QVector<int> rows(cellList.size()-2);
            for(int i=2;i<cellList.size();i++){
                rows[i-2] = cellList[i].row;
            }
            //圆锥轴承外径&设计尺寸
            for(int i=0;i<3;i++){
                range ret = getConfigRange(rows[0]+i,cola);
                if(ret.min<10000.0) TaperedBearingConfig.tb_od_range.push_back(ret);
            }
            TaperedBearingConfig.tb_od = excelBook->read(rows[0],colb).toDouble();
            //圆锥轴承内径&设计尺寸
            for(int i=0;i<3;i++){
                range ret = getConfigRange(rows[1]+i,cola);
                if(ret.min<10000.0) TaperedBearingConfig.tb_id_range.push_back(ret);
            }
            TaperedBearingConfig.tb_id = excelBook->read(rows[1],colb).toDouble();
            //圆锥轴承高度&设计尺寸
            for(int i=0;i<3;i++){
                range ret = getConfigRange(rows[2]+i,cola);
                if(ret.min<10000.0) TaperedBearingConfig.tb_h_range.push_back(ret);
            }
            TaperedBearingConfig.tb_h = excelBook->read(rows[2],colb).toDouble();
            //保持架轴承针销直径&设计尺寸
            CageBearingConfig.cb_d_range = getConfigRange(rows[3],cola);
            CageBearingConfig.cb_d = excelBook->read(rows[3],colb).toDouble();
            //角接触球轴承高度&设计尺寸
            AngularContactBallBearingConfig.acbb_h_range = getConfigRange(rows[4],cola);
            AngularContactBallBearingConfig.acbb_h = excelBook->read(rows[4],colb).toDouble();
            //垫片
            ShimConfig.shim = excelBook->read(rows[5],cola).toDouble();
            break;
    }
        case Config:{
            QStringList titlesCol = {
                                  QStringLiteral("零件的尺寸公差或范围"),//cols[0]
                                  QStringLiteral("公式参数1"),//cols[1]
                                  QStringLiteral("公式参数2"),//cols[2]
                                  QStringLiteral("公式参数3")//cols[3]
            };
            QStringList titlesRow = {
                                  QStringLiteral("针齿壳与针销"),//rows[0]
                                  QStringLiteral("针齿壳与摆线轮"),//1
                                  QStringLiteral("摆线轮轴承孔、曲轴偏心圆与保持架轴承针销间隙"),//2
                                  QStringLiteral("摆线轮、保持架轴承与曲轴的相位角匹配"),//3
                                  QStringLiteral("曲轴与锥形轴承"),//4
                                  QStringLiteral("行星架与锥形轴承"),//5
                                  QStringLiteral("行星架、角接触球轴承、针齿壳齿槽高度配合与角接触球轴承预紧量t2计算公式"),//6
                                  QStringLiteral("曲轴两个偏心圆柱与行星架卡簧槽高度配合，曲轴预紧量t1计算公式")//rows[7]
                                 };

            /*整理序号
                    ... col[0] col[1] ...
            rows[0]
            rows[1]*/
            QVector<cell> cellListRow(titlesRow.size());
            for(int i=0;i<titlesRow.size();i++){
                cellListRow[i] = lookFor(titlesRow[i],cmp,1,1,30,10);
            }
            QVector<int> rows(cellListRow.size());
            for(int i=0;i<cellListRow.size();i++){
                rows[i] = cellListRow[i].row;
            }

            QVector<cell> cellListCol(titlesCol.size());
            for(int i=0;i<titlesCol.size();i++){
                cellListCol[i] = lookFor(titlesCol[i],cmp,1,1,30,10);
            }
            QVector<int> cols(cellListCol.size());
            for(int i=0;i<cellListCol.size();i++){
                cols[i] = cellListCol[i].col;
            }
            //针齿圆d1范围
            for(int i=0;i<4;i++){
                range ret = getConfigRange(rows[0]+1+i,cols[0]);
                if(ret.min<10000.0) configs.pwc_d1_range.push_back(ret);
            }
            //针销直径D1范围
            for(int i=0;i<4;i++){
                range ret = getConfigRange(rows[0]+1+i,cols[0]+1);
                if(ret.min<10000.0) configs.np_D1_range.push_back(ret);
            }
            //针齿中心圆D2范围
            for(int i=0;i<4;i++){
                range ret = getConfigRange(rows[1]+1+i,cols[0]);
                if(ret.min<10000.0) configs.pwcc_D2_range.push_back(ret);
            }
            //摆线轮公法线范围
            for(int i=0;i<4;i++){
                range ret = getConfigRange(rows[1]+1+i,cols[0]+1);
                if(ret.min<10000.0) configs.cg_Wk_range.push_back(ret);
            }
            //摆线轮轴承孔、曲轴偏心圆与保持架轴承针销间隙
            configs.t5_range = getConfigRange(rows[2],cols[0]);
            //相位角匹配
            configs.phase_angle = getConfigRange(rows[3],cols[0]);
            //曲轴相位角的选配 0-平均 1-同时满足
            configs.phase_flag = getConfigParameterString(rows[3],cols[1]).toInt();
            //理论侧隙：0.02
            configs.delta_c = getConfigParameterString(rows[3],cols[2]).toDouble();
            //相位角补偿值：0
            configs.t6 = getConfigParameterString(rows[3],cols[3]).toDouble();
            //圆锥轴承内径d4与曲轴中心圆柱直径D4之差的范围
            configs.t4_range = getConfigRange(rows[4],cols[0]);
            //圆锥轴承内径参数
            configs.tb_id_flag = getConfigParameterString(rows[4],cols[1]).toInt();
            //圆锥轴承孔直径d3与圆锥轴承外径D3之差的范围
            configs.t3_range = getConfigRange(rows[5],cols[0]);
            //圆锥轴承外径参数
            configs.tb_od_flag = getConfigParameterString(rows[5],cols[1]).toInt();
            //齿槽高h2公差范围
            for(int i=0;i<2;i++){
                range ret = getConfigRange(rows[6]+1+i,cols[0]);
                if(ret.min<10000.0) configs.wa_h2_range.push_back(ret);
            }
            //行星架角接触球轴承高度H2公差范围
            for(int i=0;i<2;i++){
                range ret = getConfigRange(rows[6]+1+i,cols[0]+1);
                if(ret.min<10000.0) configs.acbb_H2_range.push_back(ret);
            }
            //针齿壳齿槽高与行星架角接触球轴承高度配合
            configs.t2_range = getConfigRange(rows[6]+3,cols[0]);
            //圆锥轴承高度参数
            configs.tb_h_flag = getConfigParameterString(rows[7],cols[1]).toInt();
            //行星架卡簧槽高H1公差范围
            for(int i=0;i<2;i++){
                range ret = getConfigRange(rows[7]+1+i,cols[0]);
                if(ret.min<10000.0) configs.ca_H1_range.push_back(ret);
            }
            //两个偏心圆柱高度h1公差范围
            for(int i=0;i<2;i++){
                range ret = getConfigRange(rows[7]+1+i,cols[0]+1);
                if(ret.min<10000.0) configs.ecc_h1_range.push_back(ret);
            }
            //行星架卡簧槽高与曲轴两个偏心圆柱高度配合
            configs.t1_range = getConfigRange(rows[7]+3,cols[0]);
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
        CgList[i].ID_vice = RcgList[2*i+1].ID;
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

bool partData::saveTo(QVector<re>& from,QString& _to){
    QXlsx::Document to = QXlsx::Document(_to);
    QMap<QString,cell> title;
    title.insert(QStringLiteral("序号"),cell(5,4));
    title.insert(QStringLiteral("针齿壳"),cell(6,4));
    title.insert(QStringLiteral("摆线轮A"),cell(7,4));
    title.insert(QStringLiteral("摆线轮B"),cell(8,4));
    title.insert(QStringLiteral("曲轴A"),cell(9,4));
    title.insert(QStringLiteral("曲轴B"),cell(10,4));
    title.insert(QStringLiteral("行星架"),cell(11,4));
    title.insert(QStringLiteral("圆锥轴承"),cell(12,4));
    //序号
    cell head = title[QStringLiteral("序号")];
    to.setColumnWidth(head.col,head.col + from.size(),25);
    for(int i=0 ;i< from.size();i++){
        to.write(head.row,head.col+i,QStringLiteral("组")+QString::number(i+1));
    }
    //零件
    head = title[QStringLiteral("针齿壳")];
    for(int i=0 ;i< from.size();i++){
        to.write(head.row,head.col+i,from[i].pwc_ID);

    }
    head = title[QStringLiteral("摆线轮A")];
    for(int i=0 ;i< from.size();i++){
        to.write(head.row,head.col+i,from[i].cg_A_ID);
    }
    head = title[QStringLiteral("摆线轮B")];
    for(int i=0 ;i< from.size();i++){
        to.write(head.row,head.col+i,from[i].cg_B_ID);
    }
    head = title[QStringLiteral("曲轴A")];
    for(int i=0 ;i< from.size();i++){
        to.write(head.row,head.col+i,from[i].cs_1_ID);
    }
    head = title[QStringLiteral("曲轴B")];
    for(int i=0 ;i< from.size();i++){
        to.write(head.row,head.col+i,from[i].cs_2_ID);
    }
    head = title[QStringLiteral("行星架")];
    for(int i=0 ;i< from.size();i++){
        to.write(head.row,head.col+i,from[i].pc_ID);
    }
    //标准件
    head = title[QStringLiteral("圆锥轴承")];
    for(int i=0;i<from.size();i++){
        int bias=0;
        to.write(head.row+bias++,head.col+i,from[i].tb_A1_od.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].tb_A1_id.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].tb_A1_h.toVariant());

        to.write(head.row+bias++,head.col+i,from[i].tb_A2_od.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].tb_A2_id.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].tb_A2_h.toVariant());

        to.write(head.row+bias++,head.col+i,from[i].tb_B1_od.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].tb_B1_id.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].tb_B1_h.toVariant());

        to.write(head.row+bias++,head.col+i,from[i].tb_B2_od.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].tb_B2_id.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].tb_B2_h.toVariant());

        to.write(head.row+bias++,head.col+i,from[i].cb_A1_d.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].cb_A2_d.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].cb_B1_d.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].cb_B2_d.toVariant());

        to.write(head.row+bias++,head.col+i,from[i].acbb_h.toVariant());

        to.write(head.row+bias++,head.col+i,from[i].shim_1.toVariant());
        to.write(head.row+bias++,head.col+i,from[i].shim_2.toVariant());
    }
    to.save();
    return true;
}

bool partData::cleanSrc(QVector<re> &from){
    cleanSheet(idType::pwc_ID,from);
    cleanSheet(idType::pc_ID,from);
    cleanSheet(idType::cg_A_ID,from);
    cleanSheet(idType::cg_B_ID,from);
    cleanSheet(idType::cs_1_ID,from);
    cleanSheet(idType::cs_2_ID,from);

    int getDot = sourceExcelName.indexOf('.');
    QString cleanedExcelName = sourceExcelName.leftRef(getDot)+QStringLiteral("_匹配剩余.xlsx");
    COMInterface->dynamicCall("SaveAs(QString)",QDir::toNativeSeparators(cleanedExcelName));
    return true;
}


QMap<QString,cell>& partData::idToCell(idType type){
    switch (type) {
    case idType::pwc_ID:
        return PinWheelHousingDic;
        break;
    case idType::pc_ID:
        return PlanetCarrierDic;
        break;
    case idType::cg_A_ID:
    case idType::cg_B_ID:
        return RoughCycloidGearDic;
        break;
    case idType::cs_1_ID:
    case idType::cs_2_ID:
        return CrankShaftDic;
        break;
    }
}



void partData::colNumToColName(int columnNumber, QString &res){
    while (columnNumber > 0) {
        int a0 = (columnNumber - 1) % 26 + 1;
        res += a0 - 1 + 'A';
        columnNumber = (columnNumber - a0) / 26;
    }
    std::reverse(res.begin(), res.end());
}


bool partData::cleanSheet(idType type, QVector<re>& from){
    //ini qxlsx
    excelBook->selectSheet(whichSheetHaveThis(idToPart[type]));
    //ini com
    auto COMSheet = COMInterface->querySubObject("Sheets");
    COMSheet = COMSheet->querySubObject("Item(QString&)",whichSheetHaveThis(idToPart[type]));

    auto& dicUsed = idToCell(type);
    int rowFinal = lookFor(QStringLiteral("检验结果"),[](const QString &a,const QString &b){return a==b;},1,1,300,2).row;
    //fake a column
    QVariantList _blankRange;
    for(int i=0;i<rowFinal - dicUsed[0].row;i++){
        _blankRange.append(QList<QVariant>()<<"");
    }
    QVariant blankRange = QVariant(_blankRange);

    for(auto& p : from){
        cell tmp = dicUsed[p.getIdOf(type)];

        QString cell1,cell2;
        colNumToColName(tmp.col,cell1);
        cell1 += QString::number(tmp.row);
        colNumToColName(tmp.col,cell2);
        cell2 += QString::number(rowFinal-1);
        QString rangeStr = cell1 + ":" + cell2;
        qDebug()<<p.getIdOf(type)<<":"<<rangeStr;
        auto single = COMSheet->querySubObject("Range(const QString&)",rangeStr);
        single->setProperty("Value",blankRange);

    }
    return true;
}
EXCELIOEND
