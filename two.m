clear;
clc;
%读取Excel文件
filename='附件1.xlsx';
sheet='性能数据表';
[~,~,rawData]=xlsread(filename,sheet);
%提取列标题
headers=rawData(1,:);
rawData=rawData(2:end,:);
%初始化分组结构体
groups=struct();
currentGroupID='';
groupCount=0;
%分组处理
for i=1:size(rawData,1)
    if ischar(rawData{i,1})||(~isnan(rawData{i,1})&&~isempty(rawData{i,1}))
        groupCount=groupCount+1;
        currentGroupID=rawData{i,1};
        groups(groupCount).GroupID=currentGroupID;
        groups(groupCount).Description=rawData{i,2};  
        %初始化数据矩阵
        groups(groupCount).Temperature=[];
        groups(groupCount).EthanolConversion=[];
        groups(groupCount).C4OlefinSelectivity=[];
    end
    if ~isempty(rawData{i,3})&&~isnan(rawData{i,3})
        groups(groupCount).Temperature(end+1,1)=rawData{i,3};
        groups(groupCount).EthanolConversion(end+1,1)=rawData{i,4};
        groups(groupCount).C4OlefinSelectivity(end+1,1)=rawData{i,6};
    end
end
%解析催化剂组合描述
for i=1:length(groups)
    desc=groups(i).Description;
    %提取Co负载量 (wt%)
    wtPattern='(\d+\.?\d*)wt%';
    wtMatch=regexp(desc,wtPattern,'tokens');
    if ~isempty(wtMatch)
        groups(i).CoLoading=str2double(wtMatch{1}{1});
    else
        groups(i).CoLoading=0;%默认值
    end
    %提取Co/SiO2质量 (mg)
    coPattern='(\d+\.?\d*)mg\s*\d+\.?\d*wt%Co/SiO2';
    coMatch=regexp(desc,coPattern,'tokens');
    if ~isempty(coMatch)
        groups(i).CoMass=str2double(coMatch{1}{1});
    else
        %尝试其他格式
        altPattern='(\d+\.?\d*)mg\s*\d+\.?\d*wt%';
        altMatch=regexp(desc,altPattern,'tokens');
        if ~isempty(altMatch)
            groups(i).CoMass=str2double(altMatch{1}{1});
        else
            groups(i).CoMass=0;%默认值
        end
    end 
    %提取HAP质量 (mg)
    hapPattern='(\d+\.?\d*)mg\s*HAP';
    hapMatch=regexp(desc,hapPattern,'tokens');
    if ~isempty(hapMatch)
        groups(i).HAPMass=str2double(hapMatch{1}{1});
    else
        %检查是否有"无HAP"的情况
        if contains(desc,'无HAP')
            groups(i).HAPMass=0;
        else
            %尝试其他格式
            altHapPattern='HAP\s*(\d+\.?\d*)mg';
            altHapMatch=regexp(desc,altHapPattern,'tokens');
            if ~isempty(altHapMatch)
                groups(i).HAPMass=str2double(altHapMatch{1}{1});
            else
                groups(i).HAPMass=0;%默认值
            end
        end
    end 
    %提取乙醇浓度 (ml/min)
    concPattern='乙醇浓度(\d+\.?\d*)ml/min';
    concMatch=regexp(desc,concPattern,'tokens');
    if ~isempty(concMatch)
        groups(i).EthanolConc=str2double(concMatch{1}{1});
    else
        groups(i).EthanolConc=0;%默认值
    end
end
%准备回归分析数据
X=[];%自变量: [CoLoading, CoMass, HAPMass, EthanolConc, Temperature]
Y_conv=[];%因变量1: 乙醇转化率
Y_c4=[];%因变量2: C4烯烃选择性
%收集所有数据点
for i=1:length(groups)
    numPoints=length(groups(i).Temperature);
    %创建当前组的自变量矩阵
    groupX=repmat([groups(i).CoLoading,groups(i).CoMass,groups(i).HAPMass,groups(i).EthanolConc],numPoints,1);
    groupX=[groupX,groups(i).Temperature];%添加温度列 
    %添加到总数据集
    X=[X;groupX];
    Y_conv=[Y_conv;groups(i).EthanolConversion];
    Y_c4=[Y_c4;groups(i).C4OlefinSelectivity];
end
%添加常数项 (截距)
X=[ones(size(X,1),1),X];%第一列为1表示常数项
%多元线性回归 - 乙醇转化率
fprintf('===== 乙醇转化率回归分析 =====\n');
[b_conv,bint_conv,r_conv,rint_conv,stats_conv]=regress(Y_conv,X);

%显示回归结果
fprintf('回归系数 (乙醇转化率):\n');
coeff_names={'常数项','Co负载量','Co/SiO2质量','HAP质量','乙醇浓度','温度'};
for i=1:length(b_conv)
    fprintf('%s: %.6f\n',coeff_names{i},b_conv(i));
end
fprintf('R² = %.4f, F统计量 = %.4f, p值 = %.6f\n',stats_conv(1),stats_conv(2),stats_conv(3));
%多元线性回归 - C4烯烃选择性
fprintf('\n===== C4烯烃选择性回归分析 =====\n');
[b_c4,bint_c4,r_c4,rint_c4,stats_c4]=regress(Y_c4,X);
%显示回归结果
fprintf('回归系数 (C4烯烃选择性):\n');
for i=1:length(b_c4)
    fprintf('%s: %.6f\n',coeff_names{i},b_c4(i));
end
fprintf('R² = %.4f, F统计量 = %.4f, p值 = %.6f\n',stats_c4(1),stats_c4(2),stats_c4(3));
%输出回归方程
fprintf('\n===== 回归方程 =====\n');
fprintf('乙醇转化率 = %.6f + %.6f*Co负载量 + %.6f*Co/SiO2质量 + %.6f*HAP质量 + %.6f*乙醇浓度 + %.6f*温度\n',...
        b_conv(1),b_conv(2),b_conv(3),b_conv(4),b_conv(5),b_conv(6));
fprintf('C4烯烃选择性 = %.6f + %.6f*Co负载量 + %.6f*Co/SiO2质量 + %.6f*HAP质量 + %.6f*乙醇浓度 + %.6f*温度\n',...
        b_c4(1),b_c4(2),b_c4(3),b_c4(4),b_c4(5),b_c4(6));
%残差分析 - 乙醇转化率
figure('Position',[100,100,1200,500]);
subplot(1,2,1);
scatter(Y_conv,r_conv,30,'filled');
hold on;
plot(xlim,[0,0],'r--');
xlabel('乙醇转化率观测值');
ylabel('残差');
title('乙醇转化率残差图');
grid on;
%残差分析 - C4烯烃选择性
subplot(1,2,2);
scatter(Y_c4,r_c4,30,'filled');
hold on;
plot(xlim,[0,0],'r--');
xlabel('C4烯烃选择性观测值');
ylabel('残差');
title('C4烯烃选择性残差图');
grid on;
%保存回归结果
save('regression_results.mat','b_conv','b_c4','stats_conv','stats_c4','X','Y_conv','Y_c4');
%输出分组信息
fprintf('\n成功创建 %d 个分组\n',groupCount);
fprintf('总数据点数: %d\n',size(X,1));
