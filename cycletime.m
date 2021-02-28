clc;
clear;

%�ȶ����̨���ݣ��洢��̨�����еĦ���ca^2
t=xlsread('machine.xlsx',1,'B1:K1');
cs2=xlsread('machine.xlsx',1,'B4:K4');
%���빤������
product=xlsread('product.xlsx',1);
%�������п��н�
m=xlsread('data.xlsx',1);
%ÿ�ֲ�Ʒ����������
demand=[800,1600,1400];

%main loop
Output=zeros(size(m,1),1);
for p=1:size(m,1)
    n=m(p,:);%Ͷ�Ų��ԣ�ÿ�ֲ�Ʒ��xx��������������Ͷ��
    %����Ͷ���������ܹ���Ҫ�ļӹ�ʱ��������е����WIP
    t1=demand(1)/n(1);
    t2=demand(2)/(n(2)+n(3));
    t3=demand(3)/(n(4)+n(5));
    period=max([t1,t2,t3]);
    %����,����
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
    
    %�������utilisation>=1���������÷��������ǣ���Ϊϵͳ���������ܵ���̬
    %��ʱֱ����������򣬵���һ��
    u=find(utilisation>=1);
    if u>0
        Output(p,1)=-1;
        continue;
    else
        %���EW����һ�����F(k,r,l)
        %�����л�ʱ�䣬EW=��ǰ����Ϊ���ڼӹ��Ĺ����ĸ���*��ʱ��*������
%+��ǰ���������ڼӹ��Ĺ����ĸ���
        EW=zeros(1,length(t));
        F=zeros(1,length(t));
        WIP=zeros(1,length(t));
        for i=1:length(t)
            %����GI/G/1�ֲ�ʱ  EW(:,i)=(t(i)*utilisation(:,i)*cs2(i)*exp((2*(utilisation(:,i)-1))/(3*utilisation(:,i)*(cs2(i)))))/(2*(1-utilisation(:,i)));
            %����M/G/1�ֲ�ʱ EW(:,i)=(utilisation(:,i)^2)/(2*arrival_rate(:,i)*(1-utilisation(:,i)));
            % ���ȫ����ca2����1����������GI/G/1�ֲ�ʱ����ȴ�ʱ��Ϊ
            %EW(:,i)=(t(i)*utilisation(:,i))/(2*3600*(1-utilisation(:,i)));
            %��һ����ԣ��ڿ������̨ǰ�����ӵ��Ƕ���������ֲ���ca2=0���ڿ������̨�󣬷��ӵ�����һ������ķֲ�
            if i<=3
                EW(:,i)=(t(i)*utilisation(:,i)*(0.04+cs2(i))*exp((2*(utilisation(:,i)-1)*(1-0.04)^2)/(3*utilisation(:,i)*(0.04+cs2(i)))))/(2*3600*(1-utilisation(:,i)));
            else
                EW(:,i)=(t(i)*utilisation(:,i)*(0.1+cs2(i))*exp((2*(utilisation(:,i)-1)*(1-0.1)^2)/(3*utilisation(:,i)*(0.1+cs2(i)))))/(2*3600*(1-utilisation(:,i)));
            end
            F(:,i)= EW(:,i)+t(i)/3600;
            WIP(:,i)=EW(:,i)*arrival_rate(:,i);
        end
        %���Ŀ�꺯��:�ܵ����ʱ����̣����WIP����С
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