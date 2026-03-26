clear;clc;
%按第二问逻辑解析数据
filename='附件1.xlsx';
sheet='性能数据表';
[~,~,rawData]=xlsread(filename,sheet);
headers=rawData(1,:);
rawData=rawData(2:end,:);
groups=struct();currentGroupID='';groupCount=0;
for i=1:size(rawData,1)
    if ischar(rawData{i,1})||(~isnan(rawData{i,1})&&~isempty(rawData{i,1}))
        groupCount=groupCount+1;
        currentGroupID=rawData{i,1};
        groups(groupCount).GroupID=currentGroupID;
        groups(groupCount).Description=rawData{i,2};
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
%解析描述：Co负载量、Co/SiO2质量、HAP质量、乙醇浓度
for i=1:length(groups)
    desc=groups(i).Description;
    wtPattern='(\d+\.?\d*)wt%';
    wtMatch=regexp(desc,wtPattern,'tokens');
    groups(i).CoLoading=~isempty(wtMatch)*str2double(wtMatch{1}{1});
    coPattern='(\d+\.?\d*)mg\s*\d+\.?\d*wt%Co/SiO2';
    coMatch=regexp(desc,coPattern,'tokens');
    if ~isempty(coMatch)
        groups(i).CoMass=str2double(coMatch{1}{1});
    else
        altPattern='(\d+\.?\d*)mg\s*\d+\.?\d*wt%';
        altMatch=regexp(desc,altPattern,'tokens');
        groups(i).CoMass=~isempty(altMatch)*str2double(altMatch{1}{1});
    end
    hapPattern='(\d+\.?\d*)mg\s*HAP';
    hapMatch=regexp(desc,hapPattern,'tokens');
    if ~isempty(hapMatch)
        groups(i).HAPMass=str2double(hapMatch{1}{1});
    else
        if contains(desc,'无HAP')
            groups(i).HAPMass=0;
        else
            altHapPattern='HAP\s*(\d+\.?\d*)mg';
            altHapMatch=regexp(desc,altHapPattern,'tokens');
            groups(i).HAPMass=~isempty(altHapMatch)*str2double(altHapMatch{1}{1});
        end
    end
    concPattern='乙醇浓度(\d+\.?\d*)ml/min';
    concMatch=regexp(desc,concPattern,'tokens');
    groups(i).EthanolConc=~isempty(concMatch)*str2double(concMatch{1}{1});
end
%组装训练数据：X0 = [CoLoading, CoMass, HAPMass, EthanolConc, Temperature]
X0=[];
Y_conv=[];Y_c4=[];
for i=1:length(groups)
    n=length(groups(i).Temperature);
    Xi=repmat([groups(i).CoLoading,groups(i).CoMass,groups(i).HAPMass,groups(i).EthanolConc],n,1);
    Xi=[Xi,groups(i).Temperature];
    X0=[X0;Xi];
    Y_conv=[Y_conv;groups(i).EthanolConversion];
    Y_c4=[Y_c4;groups(i).C4OlefinSelectivity];
end
%收率构造与单位自适应
if max(Y_conv)>1||max(Y_c4)>1
    Y_yield=(Y_conv.*Y_c4)/100;%单位：%
else
    Y_yield=(Y_conv.*Y_c4);%单位：比例
end
%观测最优（参考）
[bestObsYield,bestObsIdx]=max(Y_yield);
bestObsPoint=X0(bestObsIdx,:);
%构造一次 & 二次多项式特征并回归（仅 regress）
%标准化
mu=mean(X0,1);
sigma=std(X0,0,1);sigma(sigma==0)=1;
Z=(X0-mu)./sigma;
%一次特征： [1, Z]
Phi1=[ones(size(Z,1),1),Z];
%二次特征： [1, Z, Z.^2, 交互项Z_i*Z_j]
Phi2=buildPoly2(Z);
%回归（最小二乘）：使用 regress（可能出现秩亏警告，这是你选择保留的）
[b1,~,~,~,stats1]=regress(Y_yield,Phi1);
[b2,~,~,~,stats2]=regress(Y_yield,Phi2);
fprintf('===== 第三问：收率回归（基于 regress）=====\n');
fprintf('[线性] R^2 = %.4f, F = %.4f, p = %.6g\n',stats1(1),stats1(2),stats1(3));
fprintf('[二次] R^2 = %.4f, F = %.4f, p = %.6g\n',stats2(1),stats2(2),stats2(3));
%显式回归方程（原始尺度）
names={'Co负载量(wt%)','Co/SiO2质量(mg)','HAP质量(mg)','乙醇浓度(ml/min)','温度(℃)'};
%线性
[c0_lin,c_lin]=unscale_linear(b1,mu,sigma);
fprintf('\n=== 线性回归方程（原始尺度）===\n');
fprintf('y = %.6f',c0_lin);
for i=1:numel(c_lin)
    if abs(c_lin(i))>1e-12
        fprintf(' %+ .6f * %s',c_lin(i),names{i});
    end
end
fprintf('\n');
%二次
q=unscale_quadratic(b2,mu,sigma);
fprintf('\n=== 二次回归方程（原始尺度）===\n');
fprintf('y = %.6f',q.c0);
for i=1:5
    if abs(q.c(i))>1e-12,fprintf(' %+ .6f * %s',q.c(i),names{i});end
end
for i=1:5
    if abs(q.qii(i))>1e-12,fprintf(' %+ .6f * %s^2',q.qii(i),names{i});end
end
for i=1:4
    for j=i+1:5
        if abs(q.qij(i,j))>1e-12
            fprintf(' %+ .6f * %s * %s',q.qij(i,j),names{i},names{j});
        end
    end
end
fprintf('\n');
%预测函数 & 观测范围寻优（二次默认 + 线性对照）
predictLinear=@(x)[1,((x-mu)./sigma)]*b1;
predictQuadratic=@(x)buildPoly2_single((x-mu)./sigma)*b2;
%搜索边界在观测范围内
lb=min(X0,[],1);
ub=max(X0,[],1);
tempIdx=5;
%初值
Xstarts=[X0;(lb+ub)/2];
rng(2025);
Nsamp=20000;
Xrand=lb+rand(Nsamp,size(X0,2)).*(ub-lb);
%候选初值（以二次为例做评分，再用于两个模型）
Yrand_quad=arrayfun(@(i)predictQuadratic(Xrand(i,:)),1:Nsamp)';
[~,irand]=max(Yrand_quad);
x0_coarse=Xrand(irand,:);
Ystarts_quad=arrayfun(@(i)predictQuadratic(Xstarts(i,:)),1:size(Xstarts,1))';
[~,ord]=sort(Ystarts_quad,'descend');
cand=[x0_coarse;Xstarts(ord(1:min(8,end)),:)];
hasFmincon=exist('fmincon','file')==2;
%---------- 二次模型：无温度限制 ----------
predictFun=predictQuadratic;
if hasFmincon
    options=optimoptions('fmincon','Display','off','Algorithm','sqp');
    obj=@(x)-predictFun(x);
    xbest=[];fbest=inf;
    for i0=1:size(cand,1)
        try
            [xsol,fval]=fmincon(obj,cand(i0,:),[],[],[],[],lb,ub,[],options);
            if fval<fbest,fbest=fval;xbest=xsol;end
        catch,end
    end
    x_opt_all_quad=xbest;y_opt_all_quad=-fbest;
else
    Nmore=100000;
    Xrand2=lb+rand(Nmore,size(X0,2)).*(ub-lb);
    Yrand2=arrayfun(@(i)predictFun(Xrand2(i,:)),1:Nmore)';
    [y_opt_all_quad,idx]=max(Yrand2);
    x_opt_all_quad=Xrand2(idx,:);
end
%---------- 二次模型：温度<350 ----------
ub_tlim=ub;ub_tlim(tempIdx)=min(ub(tempIdx),349.999);
if hasFmincon
    options=optimoptions('fmincon','Display','off','Algorithm','sqp');
    obj=@(x)-predictFun(x);
    xbest=[];fbest=inf;
    cand_t=cand;cand_t(:,tempIdx)=min(cand_t(:,tempIdx),ub_tlim(tempIdx));
    for i0=1:size(cand_t,1)
        try
            [xsol,fval]=fmincon(obj,cand_t(i0,:),[],[],[],[],lb,ub_tlim,[],options);
            if fval<fbest,fbest=fval;xbest=xsol;end
        catch,end
    end
    x_opt_tlim_quad=xbest;y_opt_tlim_quad=-fbest;
else
    Nmore=100000;
    Xrand2=lb+rand(Nmore,size(X0,2)).*(ub_tlim-lb);
    Xrand2(:,tempIdx)=min(Xrand2(:,tempIdx),ub_tlim(tempIdx));
    Yrand2=arrayfun(@(i)predictFun(Xrand2(i,:)),1:Nmore)';
    [y_opt_tlim_quad,idx]=max(Yrand2);
    x_opt_tlim_quad=Xrand2(idx,:);
end
%---------- 线性模型：无温度限制 ----------
predictFun=predictLinear;
if hasFmincon
    options=optimoptions('fmincon','Display','off','Algorithm','sqp');
    obj=@(x)-predictFun(x);
    xbest=[];fbest=inf;
    for i0=1:size(cand,1)
        try
            [xsol,fval]=fmincon(obj,cand(i0,:),[],[],[],[],lb,ub,[],options);
            if fval<fbest,fbest=fval;xbest=xsol;end
        catch,end
    end
    x_opt_all_lin=xbest;y_opt_all_lin=-fbest;
else
    Nmore=100000;
    Xrand2=lb+rand(Nmore,size(X0,2)).*(ub-lb);
    Yrand2=arrayfun(@(i)predictFun(Xrand2(i,:)),1:Nmore)';
    [y_opt_all_lin,idx]=max(Yrand2);
    x_opt_all_lin=Xrand2(idx,:);
end
%---------- 线性模型：温度<350 ----------
if hasFmincon
    options=optimoptions('fmincon','Display','off','Algorithm','sqp');
    obj=@(x)-predictFun(x);
    xbest=[];fbest=inf;
    cand_t=cand;cand_t(:,tempIdx)=min(cand_t(:,tempIdx),ub_tlim(tempIdx));
    for i0=1:size(cand_t,1)
        try
            [xsol,fval]=fmincon(obj,cand_t(i0,:),[],[],[],[],lb,ub_tlim,[],options);
            if fval<fbest,fbest=fval;xbest=xsol;end
        catch,end
    end
    x_opt_tlim_lin=xbest;y_opt_tlim_lin=-fbest;
else
    Nmore=100000;
    Xrand2=lb+rand(Nmore,size(X0,2)).*(ub_tlim-lb);
    Xrand2(:,tempIdx)=min(Xrand2(:,tempIdx),ub_tlim(tempIdx));
    Yrand2=arrayfun(@(i)predictFun(Xrand2(i,:)),1:Nmore)';
    [y_opt_tlim_lin,idx]=max(Yrand2);
    x_opt_tlim_lin=Xrand2(idx,:);
end
%低温观测最优（参考）
mask_tlim=X0(:,tempIdx)<350;
if any(mask_tlim)
    [bestObsYield_tlim,idx_tlim]=max(Y_yield(mask_tlim));
    bestObsPoint_tlim=X0(mask_tlim,:);
    bestObsPoint_tlim=bestObsPoint_tlim(idx_tlim,:);
else
    bestObsYield_tlim=NaN;bestObsPoint_tlim=NaN(1,size(X0,2));
end
%打印结果
fprintf('\n=== 观测数据中的最佳真实收率 ===\n');
fprintf('收率(观测) = %.6f\n',bestObsYield);
for k=1:numel(names),fprintf('%s = %.6f\n',names{k},bestObsPoint(k));end
fprintf('\n=== [无温度限制] 二次模型预测最优 ===\n');
fprintf('收率(预测) = %.6f\n',y_opt_all_quad);
for k=1:numel(names),fprintf('%s = %.6f\n',names{k},x_opt_all_quad(k));end
fprintf('\n=== [温度<350℃] 二次模型预测最优 ===\n');
fprintf('收率(预测) = %.6f\n',y_opt_tlim_quad);
for k=1:numel(names),fprintf('%s = %.6f\n',names{k},x_opt_tlim_quad(k));end
fprintf('\n=== [无温度限制] 线性模型预测最优（对照） ===\n');
fprintf('收率(预测) = %.6f\n',y_opt_all_lin);
for k=1:numel(names),fprintf('%s = %.6f\n',names{k},x_opt_all_lin(k));end
fprintf('\n=== [温度<350℃] 线性模型预测最优（对照） ===\n');
fprintf('收率(预测) = %.6f\n',y_opt_tlim_lin);
for k=1:numel(names),fprintf('%s = %.6f\n',names{k},x_opt_tlim_lin(k));end
fprintf('\n=== [温度<350℃] 观测数据中的最佳真实收率(参考) ===\n');
fprintf('收率(观测) = %.6f\n',bestObsYield_tlim);
for k=1:numel(names),fprintf('%s = %.6f\n',names{k},bestObsPoint_tlim(k));end
%保存模型
save('q3_yield_optimization_regress_only.mat',...
     'b1','b2','mu','sigma','lb','ub','ub_tlim',...
     'x_opt_all_quad','y_opt_all_quad','x_opt_tlim_quad','y_opt_tlim_quad',...
     'x_opt_all_lin','y_opt_all_lin','x_opt_tlim_lin','y_opt_tlim_lin',...
     'bestObsYield','bestObsPoint','bestObsYield_tlim','bestObsPoint_tlim');
%辅助函数
function Phi=buildPoly2(Z)
    [n,p]=size(Z);
    n_inter=p*(p-1)/2;
    Phi=zeros(n,1+2*p+n_inter);
    Phi(:,1)=1;
    Phi(:,2:1+p)=Z;
    Phi(:,2+p:1+2*p)=Z.^2;
    col=1+2*p;
    for i=1:p-1
        for j=i+1:p
            col=col+1;
            Phi(:,col)=Z(:,i).*Z(:,j);
        end
    end
end
function phi=buildPoly2_single(z)
    p=numel(z);
    n_inter=p*(p-1)/2;
    phi=zeros(1,1+2*p+n_inter);
    phi(1)=1;
    phi(2:1+p)=z;
    phi(2+p:1+2*p)=z.^2;
    col=1+2*p;
    for i=1:p-1
        for j=i+1:p
            col=col+1;
            phi(col)=z(i)*z(j);
        end
    end
end
function [c0,c]=unscale_linear(b,mu,sigma)
%b: [b0, b1..bp] 对应 y = b0 + sum(bi * Zi), Zi=(xi-mui)/sigmai
    beta=b(2:end)./sigma(:);
    c=beta(:).';
    c0=b(1)-sum((mu(:)./sigma(:)).*b(2:end));
end
function q=unscale_quadratic(b,mu,sigma)
%b 对应 buildPoly2(Z) 的系数： [1, Z_i, Z_i^2, Z_iZ_j]
%返回 y = c0 + sum c_i x_i + sum qii x_i^2 + sum_{i<j} qij x_i x_j
    p=numel(mu);
    q.c0=b(1);
    q.c=zeros(1,p);
    q.qii=zeros(1,p);
    q.qij=zeros(p,p);
    idx_lin=2:1+p;
    idx_sq=2+p:1+2*p;
    for i=1:p
        bi=b(idx_lin(i));
        q.c(i)=q.c(i)+bi/sigma(i);
        q.c0=q.c0-bi*mu(i)/sigma(i);
    end
    for i=1:p
        ai=b(idx_sq(i));
        q.qii(i)=q.qii(i)+ai/(sigma(i)^2);
        q.c(i)=q.c(i)-2*ai*mu(i)/(sigma(i)^2);
        q.c0=q.c0+ai*(mu(i)^2)/(sigma(i)^2);
    end
    col=1+2*p;
    for i=1:p-1
        for j=i+1:p
            col=col+1;
            aij=b(col);
            s=aij/(sigma(i)*sigma(j));
            q.qij(i,j)=q.qij(i,j)+s;
            q.c(j)=q.c(j)-s*mu(i);
            q.c(i)=q.c(i)-s*mu(j);
            q.c0=q.c0+s*mu(i)*mu(j);
        end
    end
end