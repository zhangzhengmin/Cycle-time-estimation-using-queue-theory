clc;
clear;

%先读入机台数据，存储机台数据中的μ与ca^2
t=xlsread('machine.xlsx',1,'B1:K1');
cs2=xlsread('machine.xlsx',1,'B4:K4');
%读入工艺数据
product=xlsread('product.xlsx',1);
%读入所有可行解
m=xlsread('data.xlsx',1);
%每种产品的需求数量
demand=[800,1600,1400];

%main loop
Output=zeros(size(m,1),1);
for p=1:size(m,1)
    n=m(p,:);%投放策略：每种产品以xx的速率向生产线投放
    %根据投放量计算总共需要的加工时间与过程中的最大WIP
    t1=demand(1)/n(1);
    t2=demand(2)/(n(2)+n(3));
    t3=demand(3)/(n(4)+n(5));
    period=max([t1,t2,t3]);
    %求解λ,求解ρ
    e=[];
    arrival_rate=zeros(1,length(t));
    utilisation=zeros(1,length(t));
    for i=1: size(product,1)
        b=tabulate(product(i,3:end));
        e(:,i)=b(:,2);
    end
    for i=1:length(t)
        arrival_rate(:,i)=sum(e(i,:)*n');
        utilisation(:,i)=arrival_rate(:,i)/(60*60/t(i));
    end
    
    %如果存在utilisation>=1的情况，则该方案不考虑，因为系统根本不可能到稳态
    %此时直接跳出大程序，到下一步
    u=find(utilisation>=1);
    if u>0
        Output(p,1)=-1;
        continue;
    else
        %求解EW，进一步求解F(k,r,l)
        %考虑切换时间，EW=当前工件为正在加工的工件的概率*（时间*数量）
%+当前工件非正在加工的工件的概率
        EW=zeros(1,length(t));
        F=zeros(1,length(t));
        WIP=zeros(1,length(t));
        for i=1:length(t)
            %服从GI/G/1分布时  EW(:,i)=(t(i)*utilisation(:,i)*cs2(i)*exp((2*(utilisation(:,i)-1))/(3*utilisation(:,i)*(cs2(i)))))/(2*(1-utilisation(:,i)));
            %服从M/G/1分布时 EW(:,i)=(utilisation(:,i)^2)/(2*arrival_rate(:,i)*(1-utilisation(:,i)));
            % 如果全都当ca2大于1来处理，服从GI/G/1分布时，则等待时间为
            %EW(:,i)=(t(i)*utilisation(:,i))/(2*3600*(1-utilisation(:,i)));
            %但一般而言，在可重入机台前，服从的是定长的输入分布，ca2=0，在可重入机台后，服从的是有一定方差的分布
            if i<=3
                EW(:,i)=(t(i)*utilisation(:,i)*(0.04+cs2(i))*exp((2*(utilisation(:,i)-1)*(1-0.04)^2)/(3*utilisation(:,i)*(0.04+cs2(i)))))/(2*3600*(1-utilisation(:,i)));
            else
                EW(:,i)=(t(i)*utilisation(:,i)*(0.1+cs2(i))*exp((2*(utilisation(:,i)-1)*(1-0.1)^2)/(3*utilisation(:,i)*(0.1+cs2(i)))))/(2*3600*(1-utilisation(:,i)));
            end
            F(:,i)= EW(:,i)+t(i)/3600;
            WIP(:,i)=EW(:,i)*arrival_rate(:,i);
        end
        %求解目标函数:总的完成时间最短，最大WIP量最小
        C=zeros(1,size(n,2));
        for i=1:size(n,2)
            C(:,i)=sum(F*e(:,i));
        end
        makespan=period+max(C);
        WIP_max=max(WIP);
        Output(p,1)=WIP_max*10+makespan;
    end
end

X=find(Output>0);
Output=[X,Output(X)];