clear;clc;
filename='附件1.xlsx';
sheet='性能数据表';
[~,~,rawData]=xlsread(filename,sheet);
headers=rawData(1,:);
rawData=rawData(2:end,:);%移除标题行
%初始化分组结构体
groups=struct();
currentGroupID='';
groupCount=0;
for i=1:size(rawData,1)
    %新组开始：第一列出现字符串或数值
    if ischar(rawData{i,1})||~isnan(rawData{i,1})
        groupCount=groupCount+1;
        currentGroupID=rawData{i,1};
        groups(groupCount).GroupID=currentGroupID;
        groups(groupCount).Description=rawData{i,2};
        %初始化数据矩阵
        groups(groupCount).Temperature=[];
        groups(groupCount).EthanolConversion=[];
        groups(groupCount).EthyleneSelectivity=[];
        groups(groupCount).C4OlefinSelectivity=[];
        groups(groupCount).AcetaldehydeSelectivity=[];
        groups(groupCount).C4_12AlcoholSelectivity=[];
        groups(groupCount).BenzaldehydeSelectivity=[];
        groups(groupCount).OtherSelectivity=[];
    end
    %添加数据到当前组
    if ~isempty(rawData{i,3})&&~isnan(rawData{i,3})
        groups(groupCount).Temperature(end+1,1)=rawData{i,3};
        groups(groupCount).EthanolConversion(end+1,1)=rawData{i,4};
        groups(groupCount).EthyleneSelectivity(end+1,1)=rawData{i,5};
        groups(groupCount).C4OlefinSelectivity(end+1,1)=rawData{i,6};
        groups(groupCount).AcetaldehydeSelectivity(end+1,1)=rawData{i,7};
        groups(groupCount).C4_12AlcoholSelectivity(end+1,1)=rawData{i,8};
        groups(groupCount).BenzaldehydeSelectivity(end+1,1)=rawData{i,9};
        groups(groupCount).OtherSelectivity(end+1,1)=rawData{i,10};
    end
end
%保存分组数据
save('catalyst_groups.mat','groups');
fprintf('成功创建 %d 个分组\n',groupCount);
%乙醇转化率（y1）分析
%输出文件夹与结果容器（避免冲突）
outFolderConv='Group_Plots';
if ~exist(outFolderConv,'dir');mkdir(outFolderConv);end
resultsConv=struct('GroupID',{},'BestModel',{},'R2',{},'Parameters',{});
for g=1:length(groups)
    G=groups(g);
    groupID=G.GroupID;
    temperature=G.Temperature(:);
    conversion=G.EthanolConversion(:);
    %=====模型比较图=====
    fig1=figure('Visible','off','Position',[100,100,900,700]);
    scatter(temperature,conversion,70,'filled','b');hold on;
    xlabel('温度 (°C)','FontSize',12);
    ylabel('乙醇转化率 (%)','FontSize',12);
    title(sprintf('%s组: 温度与乙醇转化率关系及模型拟合',groupID),'FontSize',14);
    grid on;
    %拟合
    %线性
    [p_linear,S_linear]=polyfit(temperature,conversion,1);
    y_linear=polyval(p_linear,temperature);
    plot(temperature,y_linear,'r--','LineWidth',1.8);
    %二次
    [p_quad,S_quad]=polyfit(temperature,conversion,2);
    y_quad=polyval(p_quad,temperature);
    plot(temperature,y_quad,'g-.','LineWidth',1.8);
    %三次
    [p_cubic,S_cubic]=polyfit(temperature,conversion,3);
    y_cubic=polyval(p_cubic,temperature);
    plot(temperature,y_cubic,'m:','LineWidth',1.8);
    %指数（y>0）
    exp_valid=false;
    if all(conversion>0)
        try
            p_exp=polyfit(temperature,log(conversion),1);
            a_exp=exp(p_exp(2));
            b_exp=p_exp(1);
            y_exp=a_exp*exp(b_exp*temperature);
            plot(temperature,y_exp,'k-','LineWidth',1.8);
            exp_valid=true;
        catch
            exp_valid=false;
        end
    end
    %R^2
    R2_linear=1-(S_linear.normr^2)/((length(conversion)-1)*var(conversion));
    R2_quad=1-(S_quad.normr^2)/((length(conversion)-1)*var(conversion));
    R2_cubic=1-(S_cubic.normr^2)/((length(conversion)-1)*var(conversion));
    if exp_valid
        R2_exp=1-sum((conversion-y_exp).^2)/sum((conversion-mean(conversion)).^2);
    else
        R2_exp=-Inf;
    end
    R2_values=[R2_linear,R2_quad,R2_cubic,R2_exp];
    model_names={'线性模型','二次模型','三次模型','指数模型'};
    [best_R2,best_idx]=max(R2_values);
    best_model=model_names{best_idx};
    %图例
    if exp_valid
        legend({'原始数据','线性模型','二次模型','三次模型','指数模型'},'Location','best','FontSize',10);
    else
        legend({'原始数据','线性模型','二次模型','三次模型'},'Location','best','FontSize',10);
    end
    %信息框
    info_text={sprintf('线性 R² = %.10f',R2_linear),...
        sprintf('二次 R² = %.10f',R2_quad),...
        sprintf('三次 R² = %.10f',R2_cubic)};
    if exp_valid,info_text{end+1}=sprintf('指数 R² = %.10f',R2_exp);end
    info_text{end+1}=sprintf('最佳模型: %s',best_model);

    annotation(fig1,'textbox',[0.15,0.7,0.2,0.18],...
        'String',info_text,'FitBoxToText','on',...
        'BackgroundColor',[1,1,1,0.7],'EdgeColor','k','FontSize',10,'Interpreter','none');
    print(fig1,fullfile(outFolderConv,sprintf('%s_model_comparison.png',groupID)),'-dpng','-r300');
    close(fig1);
    %=====最佳模型图=====
    fig2=figure('Visible','off','Position',[100,100,900,700]);
    scatter(temperature,conversion,80,'filled','b');hold on;
    x_range=linspace(min(temperature)-10,max(temperature)+50,200);
    switch best_idx
        case 1%线性
            y_range=polyval(p_linear,x_range);
            eq_text=sprintf('y = %.10f x + %.10f',p_linear(1),p_linear(2));
            params=p_linear;
        case 2%二次
            y_range=polyval(p_quad,x_range);
            eq_text=sprintf('y = %.10f x² + %.10f x + %.10f',p_quad(1),p_quad(2),p_quad(3));
            params=p_quad;
        case 3%三次
            y_range=polyval(p_cubic,x_range);
            eq_text=sprintf('y = %.10f x³ + %.10f x² + %.10f x + %.10f',p_cubic(1),p_cubic(2),p_cubic(3),p_cubic(4));
            params=p_cubic;
        case 4%指数
            y_range=a_exp*exp(b_exp*x_range);
            eq_text=sprintf('y = %.10f e^{%.10f x}',a_exp,b_exp);
            params=[a_exp,b_exp];
    end
    plot(x_range,y_range,'r-','LineWidth',2.5);
    xlabel('温度 (°C)','FontSize',12);
    ylabel('乙醇转化率 (%)','FontSize',12);
    title(sprintf('%s组: 最佳拟合模型 - %s (R² = %.10f)',groupID,best_model,best_R2),'FontSize',14);
    grid on;legend({'原始数据','拟合曲线'},'Location','best','FontSize',10);
    annotation(fig2,'textbox',[0.15,0.75,0.3,0.1],...
        'String',sprintf('拟合方程: %s\nR² = %.10f',eq_text,best_R2),...
        'FitBoxToText','on','BackgroundColor',[1,1,1,0.7],'EdgeColor','k','FontSize',11,'Interpreter','tex');

    print(fig2,fullfile(outFolderConv,sprintf('%s_best_model.png',groupID)),'-dpng','-r300');
    close(fig2);
    %结果保存（与原逻辑一致）
    resultsConv(g).GroupID=groupID;
    resultsConv(g).BestModel=best_model;
    resultsConv(g).R2=best_R2;
    resultsConv(g).Parameters=params;
    fprintf('组 %s 转化率分析完成: 最佳模型 = %s (R² = %.10f)\n',groupID,best_model,best_R2);
end
save('model_fitting_results.mat','resultsConv');%保留原文件名（前半问）
fprintf('\n===== 模型拟合结果汇总（乙醇转化率）=====\n');
validConv=resultsConv(~cellfun(@isempty,{resultsConv.GroupID}));
if ~isempty(validConv)
    Tconv=struct2table(validConv);
    Tconv=Tconv(:,{'GroupID','BestModel','Parameters'});
    disp(Tconv);
else
    disp('无有效结果');
end
%C4烯烃选择性（y2）分析 —— 合并后半问
outFolderSel='C4_Selectivity_Plots';
if ~exist(outFolderSel,'dir');mkdir(outFolderSel);end
resultsSel=struct('GroupID',{},'BestModel',{},'R2',{},'Parameters',{});

for g=1:length(groups)
    G=groups(g);
    groupID=G.GroupID;
    temperature=G.Temperature(:);
    selectivity=G.C4OlefinSelectivity(:);
    %跳过数据点不足（至少3点更稳）
    if numel(temperature)<3||numel(selectivity)<3
        warning('组 %s 数据点不足(%d)，跳过拟合',groupID,min(numel(temperature),numel(selectivity)));
        continue;
    end
    %=====模型比较图=====
    fig1=figure('Visible','off','Position',[100,100,900,700]);
    scatter(temperature,selectivity,70,'filled','b');hold on;
    xlabel('温度 (°C)','FontSize',12);
    ylabel('C4烯烃选择性 (%)','FontSize',12);
    title(sprintf('%s组: 温度与C4烯烃选择性关系及模型拟合',groupID),'FontSize',14);
    grid on;
    %线性 / 二次 / 三次
    [p_linear,S_linear]=polyfit(temperature,selectivity,1);
    y_linear=polyval(p_linear,temperature);plot(temperature,y_linear,'r--','LineWidth',1.8);
    [p_quad,S_quad]=polyfit(temperature,selectivity,2);
    y_quad=polyval(p_quad,temperature);plot(temperature,y_quad,'g-.','LineWidth',1.8);
    [p_cubic,S_cubic]=polyfit(temperature,selectivity,3);
    y_cubic=polyval(p_cubic,temperature);plot(temperature,y_cubic,'m:','LineWidth',1.8);
    %指数（y>0）
    exp_valid=false;
    if all(selectivity>0)
        try
            p_exp=polyfit(temperature,log(selectivity),1);
            a_exp=exp(p_exp(2));b_exp=p_exp(1);
            y_exp=a_exp*exp(b_exp*temperature);
            plot(temperature,y_exp,'k-','LineWidth',1.8);
            exp_valid=true;
        catch
            exp_valid=false;
        end
    end
    %R^2
    R2_linear=1-(S_linear.normr^2)/((length(selectivity)-1)*var(selectivity));
    R2_quad=1-(S_quad.normr^2)/((length(selectivity)-1)*var(selectivity));
    R2_cubic=1-(S_cubic.normr^2)/((length(selectivity)-1)*var(selectivity));
    if exp_valid
        R2_exp=1-sum((selectivity-y_exp).^2)/sum((selectivity-mean(selectivity)).^2);
    else
        R2_exp=-Inf;
    end
    R2_values=[R2_linear,R2_quad,R2_cubic,R2_exp];
    model_names={'线性模型','二次模型','三次模型','指数模型'};
    [best_R2,best_idx]=max(R2_values);
    best_model=model_names{best_idx};
    %图例
    if exp_valid
        legend({'原始数据','线性模型','二次模型','三次模型','指数模型'},'Location','best','FontSize',10);
    else
        legend({'原始数据','线性模型','二次模型','三次模型'},'Location','best','FontSize',10);
    end
    %信息框
    info_text={sprintf('线性 R² = %.10f',R2_linear),...
        sprintf('二次 R² = %.10f',R2_quad),...
        sprintf('三次 R² = %.10f',R2_cubic)};
    if exp_valid,info_text{end+1}=sprintf('指数 R² = %.10f',R2_exp);end
    info_text{end+1}=sprintf('最佳模型: %s',best_model);
    annotation(fig1,'textbox',[0.15,0.7,0.2,0.18],...
        'String',info_text,'FitBoxToText','on',...
        'BackgroundColor',[1,1,1,0.7],'EdgeColor','k','FontSize',10,'Interpreter','none');
    print(fig1,fullfile(outFolderSel,sprintf('%s_c4_selectivity_comparison.png',groupID)),'-dpng','-r300');
    close(fig1);
    %=====最佳模型图=====
    fig2=figure('Visible','off','Position',[100,100,900,700]);
    scatter(temperature,selectivity,80,'filled','b');hold on;
    x_range=linspace(min(temperature)-10,max(temperature)+50,200);
    switch best_idx
        case 1%线性
            y_range=polyval(p_linear,x_range);
            eq_text=sprintf('y = %.10f x + %.10f',p_linear(1),p_linear(2));
            params=p_linear;
        case 2%二次
            y_range=polyval(p_quad,x_range);
            eq_text=sprintf('y = %.10f x² + %.10f x + %.10f',p_quad(1),p_quad(2),p_quad(3));
            params=p_quad;
        case 3%三次
            y_range=polyval(p_cubic,x_range);
            eq_text=sprintf('y = %.10f x³ + %.10f x² + %.10f x + %.10f',p_cubic(1),p_cubic(2),p_cubic(3),p_cubic(4));
            params=p_cubic;
        case 4%指数
            y_range=a_exp*exp(b_exp*x_range);
            eq_text=sprintf('y = %.10f e^{%.10f x}',a_exp,b_exp);
            params=[a_exp,b_exp];
    end
    plot(x_range,y_range,'r-','LineWidth',2.5);
    xlabel('温度 (°C)','FontSize',12);
    ylabel('C4烯烃选择性 (%)','FontSize',12);
    title(sprintf('%s组: 最佳拟合模型 - %s (R² = %.10f)',groupID,best_model,best_R2),'FontSize',14);
    grid on;legend({'原始数据','拟合曲线'},'Location','best','FontSize',10);
    annotation(fig2,'textbox',[0.15,0.75,0.3,0.1],...
        'String',sprintf('拟合方程: %s\nR² = %.10f',eq_text,best_R2),...
        'FitBoxToText','on','BackgroundColor',[1,1,1,0.7],'EdgeColor','k','FontSize',11,'Interpreter','tex');
    print(fig2,fullfile(outFolderSel,sprintf('%s_c4_selectivity_best_model.png',groupID)),'-dpng','-r300');
    close(fig2);
    %结果保存
    resultsSel(g).GroupID=groupID;
    resultsSel(g).BestModel=best_model;
    resultsSel(g).R2=best_R2;
    resultsSel(g).Parameters=params;
    fprintf('组 %s C4选择性分析完成: 最佳模型 = %s (R² = %.10f)\n',groupID,best_model,best_R2);
end
save('c4_selectivity_fitting_results.mat','resultsSel');%保留原后半问文件名
validSel=resultsSel(~cellfun(@isempty,{resultsSel.GroupID}));
if ~isempty(validSel)
    Tsel=struct2table(validSel);
    Tsel=Tsel(:,{'GroupID','BestModel','Parameters'});
    disp(Tsel);
else
    disp('无有效结果');
end