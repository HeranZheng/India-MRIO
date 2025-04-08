%% India MRIO development
clear all

% VA estimate
folderPath = 'C:\Users\zheng\OneDrive - University College London\Research\#IO TABLE\India\Version 2\Data'; % 修改为你的Excel文件路径
Year=[2011:2019];

cd(folderPath);

OldFolder = pwd;

% meta for all
n_state = 36;
n_IOT = 66;

MAP = xlsread('Meta','Mapping');
MAP(isnan(MAP)) = 0;

Employment = xlsread('./Employment/Employment');
Employment(isnan(Employment)) = 0;
Employment=Employment(3:end,3:end);

[~,~,Gravity_model_mapping] = xlsread('Meta','Gravity_Sector');% gravity model sector mapping
Gravity_model_mapping = cell2mat(Gravity_model_mapping(2:end,2:end));
Gravity_model_mapping(isnan(Gravity_model_mapping)) = 0;


[~,~,Gravity_bridge] = xlsread('Meta','Gravity_bridge');% g
Gravity_bridge = cell2mat(Gravity_bridge(2:end,2:end));
Gravity_bridge(isnan(Gravity_bridge)) = 0;

%% coverting 36 states to 32 states
MAP_STATE = xlsread('Meta','State');
MAP_STATE(isnan(MAP_STATE)) = 0;

%% state distance
d1 = xlsread('Meta','Distance');

% Mapping for each state
% Mining (5-10), energy (42-44) and Tertiary sectors (55-65), using Employment data
% Industry (11-40), using AIS data 
share_employment_mining = Employment(:,5:10)./sum(Employment(:,5:10),2);
share_employment_energy = Employment(:,42:44)./sum(Employment(:,42:44),2);
share_employment_service_1 = Employment(:,55:61)./sum([Employment(:,55:61),Employment(:,63:65)],2);;
share_employment_service_2 = Employment(:,63:65)./sum([Employment(:,55:61),Employment(:,63:65)],2);;

for i= 1:36
MAP(5:10,5) = share_employment_mining (i,:);
MAP(42:44,7) = share_employment_energy(i,:);

MAP(55:61,21) = share_employment_service_1(i,:);
MAP(63:65,21) = share_employment_service_2(i,:);


MAP_State(:,:,i) = MAP;

end

%% GSVA
cd(['./GSVA']);

fileList = dir(fullfile(pwd, '*.xlsx'));

GSVA_all = zeros(21*36,15);

for j=1:size(fileList,1)

GSVA_all ((j-1)*36+1:(j-1)*36+36,1:15) = xlsread([fileList(j).name]);

end

%% 2011-2019 9 years

GSVA_all = GSVA_all(:,3:3+size(Year,2)-1);

cd(OldFolder);

%% Got import from SUT
MAP_SUT = xlsread('Meta','SUT');
MAP_SUT(isnan(MAP_SUT)) = 0;

year_sut=[2011,2013:2019];

cd(['./SUT']);

for i= 1 :size(year_sut,2)

SUT = xlsread(['SUT_',num2str(year_sut(i))],"Supply");

Import_national=SUT(4:143,76);
Import_national(isnan(Import_national)) = 0;

Import_national_ad(:,i) = MAP_SUT'*Import_national;

end
%% add 2012

Import_national_ad1=zeros(66,9);
Import_national_ad1(:,1)=Import_national_ad(:,1);
Import_national_ad1(:,3:9)=Import_national_ad(:,2:end);
Import_national_ad1(:,2) = (Import_national_ad1(:,1)+Import_national_ad1(:,3))./2;

cd(OldFolder);


%% SRIO and MRIO Frame
[~,~,Frame_sector] = xlsread('Meta','Sector');
[~,~,Frame_State] = xlsread('Meta','State_32');
[~,~,Frame_Final] = xlsread('Meta','Final_Demand');

Frame_sector=Frame_sector(2:end,2);

Frame_State_1=repmat(Frame_State,1,66)';

Frame_MRIO=repmat(Frame_sector(1:end,:),size(Frame_State,1),1);
Frame_Import=[{[''],['Import ']}];
Frame_VA=[{[''],['VA']}];
Frame_Output=[{[''],['Output']}];

Empty={['']};

Frame_Final1=[repmat(Empty,1,size(Frame_State,1)*5);repmat(Frame_Final(1:end,:)',1,size(Frame_State,1))];
Frame_Export=[{[''];['Export']},{[''];['Error']},{[''];['Output']}];

FrameFull_col=[[Frame_State_1(:),Frame_MRIO];Frame_Import;Frame_VA;Frame_Output];

FrameFull_row=[[Frame_State_1(:),Frame_MRIO]',Frame_Final1,Frame_Export];

FrameFull=repmat(Empty,size(FrameFull_col,1),size(FrameFull_row,2));

FrameFull(3:size(FrameFull_col,1)+2,1:2)=FrameFull_col;

FrameFull(1:2,3:size(FrameFull_row,2)+2)=FrameFull_row;


% Frame SRIO
[~,~,Frame_SRIO] =xlsread('Meta','SRIO') ; 


for t=5 %7 :size(Year,2)
%% IOT import % unit: RS lakh (100000 Rs) 
 cd(['./IOT/']);
[~,~,IOT] = xlsread([num2str(Year(t))]) ;


IOT_Output = cell2mat(IOT(3:68,75));

IOT_VA = sum(cell2mat( IOT(70:73,3:68)),1);

IOT_Z = cell2mat(IOT(3:68,3:68));

IOT_F = cell2mat(IOT(3:68,69:73));

IOT_Import = cell2mat(IOT(69,3:68));

IOT_Export = cell2mat(IOT(3:68,74));

IOT_A = IOT_Z./IOT_Output;

IOT_Demand= sum(IOT_Z,2) + sum(IOT_F,2);

IOT_Supply=IOT_Output-IOT_Export;

F_struct = sum(IOT_F,2)./sum(sum(IOT_F));
F_struct(isnan(F_struct)) = 0; 
F_struct(isinf(F_struct)) = 0; 

F_struct2 = IOT_F./sum(IOT_F,2);
F_struct2(isnan(F_struct2)) = 0; 
F_struct2(isinf(F_struct2)) = 0; 

cd(OldFolder);


 %% ASI dataset   
cd(['./ASI/State/' num2str(Year(t))]);

AIS_Output=xlsread('15#15. Total Output.xlsx');

AIS_VA=xlsread('19#19. Gross Value Added.xlsx');

cd(OldFolder);
% remove A island (30)
AIS_VA(30,:) = 0;
AIS_Output(30,:) = 0;

Share_AIS = AIS_VA (:,11:40)./sum(AIS_VA (:,11:40),2);
Share_AIS(isnan(Share_AIS)) = 0;

%% VA rate from AIS data
VA_rate_AIS = AIS_VA./AIS_Output;
VA_rate_AIS(isnan(VA_rate_AIS)) = 0;

%% VA rate from IOT
% non-industry, using IOT
% Industry (11-40), using AIS data 
VA_rate_IOT = IOT_VA./IOT_Output';

        for j=1:36
        VA_rate_All (j,1:10) = VA_rate_IOT(:,1:10);
        VA_rate_All (j,11:40) = VA_rate_AIS (j,11:40);
        VA_rate_All (j,41:66) = VA_rate_IOT (:,41:66);
        end
        
        
        %% Mapping
        for j =1 : 36
        
          MAP_State(11:40,6,j) = Share_AIS(j,:);
        
        end
        
        
        %% Integrating AIS in GSVA (GSVA as benchmark)
        GSVA_y=GSVA_all(:,t);
        
        for j=1:size(fileList,1)% GSVA Sector
        GSVA_y1(1:n_state,j)=GSVA_y((j-1)*n_state+1:(j-1)*n_state+n_state,1);
        end
        
        % for each state
        for j= 1:36 
        VA_all(j,:)=GSVA_y1(j,:)* MAP_State(:,:,j)';
        end
        VA_all(isnan(VA_all))=0;

                
                %$ RAS for VA
                % scale State GDPby national GDP
                State_GDP = sum(IOT_VA).*(sum(GSVA_y1,2)./sum(sum(GSVA_y1,2)));
                
                VA_adjust = RAS_india(VA_all,State_GDP,IOT_VA,10);
                
                % Output estimate using VA rate
                Output_adjust = VA_adjust./VA_rate_All;
                Output_adjust (isnan(Output_adjust)) = 0;
                
                % in case VA is negative

                Output_adjust= abs(Output_adjust);


                % benchmarking with IOT_OUTPUT
                Output_adjust=IOT_Output'.*(Output_adjust./sum(Output_adjust,1));
                
                
                if any(Output_adjust(:) < 0)
                        disp('Matrix contains elements less than 0. Exiting.');
                        return; % Exit the function
                end
                disp('All elements in the matrix are non-negative.');
                
                
                %% Export estimate (imports) unit rs
                cd(['./Export/Local_currency/']);
                
                Export_state=xlsread( [num2str(Year(t)), '_INR.xlsx']);
                %Telangana is suspicous before 2014
                
                Export_state = Export_state./100000; % Converting RS to RS lakh
                
                % estimate service trade and energy
                Rate_Export = IOT_Export./IOT_Output;
                
                Export_state_pre= Rate_Export'.* Output_adjust;
                
                Export_state(:,40:64) =  Export_state_pre(:,40:64);
                
                Export_state(:,66) =  Export_state_pre(:,66);
                
                %% scale by IOT EXPORT
                Export_state_ad = IOT_Export'.*(Export_state./sum(Export_state,1));
                
                %% re-adjust Export with condition of output
                rate1 = Export_state_ad./Output_adjust;
                rate1(isnan(rate1)) = 0;
                rate1(isinf(rate1)) = 0;
                
                Rate_Export1 = repmat(Rate_Export',36,1);
                
                if any((rate1(find(rate1>1)))>0)
                
                        rate1((rate1<1))=0;
                        
                        Output_adjust_1 = Output_adjust .* (rate1 ~= 0);
                        
                        Export_state_ad1 = Rate_Export1.*Output_adjust_1;
                        
                        delta = Export_state_ad.* (rate1 ~= 0)-Export_state_ad1;
                        
                        delta = sum(delta,1);
                        
                        Output_adjust_2 = Output_adjust.*any(rate1,1)-Output_adjust_1;
                        
                        % remove zero part
                        Output_adjust_2(Export_state_ad==0)=0;
                        
                        Export_adjust_add=delta.*(Output_adjust_2./sum(Output_adjust_2,1));
                        Export_adjust_add(isnan(Export_adjust_add)) = 0;
                        
                        Export_adjust_add = Export_adjust_add + Export_state_ad1;% add estimated export back
                        
                        % remove previous problmatic array
                        Export_state_ad((rate1 ~= 0))=0;
                        
                        Export_state_new0 = Export_state_ad+Export_adjust_add;
                        
                        
                        rate1 = Export_state_new0./Output_adjust;
                        rate1(isnan(rate1)) = 0;
                        rate1(isinf(rate1)) = 0;
                        
                        if any(Output_adjust-Export_state_new0<0,"all")
                
                            Export_state_new0(find(Output_adjust-Export_state_new0<0))=0;
                            
                             Export_state_new1 =Export_state_new0;
                             disp('Export adjusted and gogogo');
                            Export_state_new1(isnan(Export_state_new1)) =0;

                        else
                            Export_state_new1=Export_state_new0;
                            disp('Export gogogo');
                            Export_state_new1(isnan(Export_state_new1)) =0;
                
                        end
                end
                
                
                cd(OldFolder);

%% state_level import
    
    for j=1:36
    
      Demand(j,:) = sum(IOT_A.*Output_adjust(j,:),2)';
    
    end
    
    Import_state =(Demand./sum(Demand,1)).* Import_national_ad1(:,t)';
    Import_state(isnan(Import_state)) = 0;
    
% converting 36 to 32 states
Output_new = MAP_STATE'*Output_adjust;
VA_new = MAP_STATE'* VA_adjust;
Import_state_new = MAP_STATE'* Import_state;
Export_state_new = MAP_STATE'* Export_state_new1;
Demand_new = MAP_STATE'* Demand;

Export_state_new(isnan(Export_state_new)) = 0;

    Rate_out = (Export_state_new./sum(Export_state_new,1));
    Rate_out(isnan(Rate_out))=0;
    Rate_out(isinf(Rate_out))=0;
    
    Export_state_new=IOT_Export'.*Rate_out;

% re-distribution
Rate_out = (VA_new./sum(VA_new,1));
Rate_out(isnan(Rate_out))=0;
Rate_out(isinf(Rate_out))=0;

    VA_new=IOT_VA.*Rate_out;

    Output_new(find(Output_new<1))=0;
    VA_new(find(VA_new<1))=0;

    %  iterative proportional fitting(IPF)
                % Dimensions
                [m, n] = size(Output_new);
                
                if any(IOT_Output < sum(VA_new, 1)')
                    error('Infeasible problem: column sum constraints exceed B''s maximum column sums.');
                end
    
                % Reshape A into a vector for optimization
                A_vec = Output_new(:);
                
                % Objective Function: Minimize the difference between A and its initial values
                objective = @(x) sum((x - A_vec).^2);
                
                % Constraints
                Aeq = kron(eye(n), ones(1, m));  % Constraint for column sums (Aeq * A_vec = C)
                beq = IOT_Output;
                                      
                lb =  max(abs(VA_new(:)), Export_state_new(:));       % Lower bound (A >= 0)
                               
                % Solve using constrained optimization
                options = optimoptions('lsqlin', 'Display', 'off');  % Suppress output
                A_adjusted_vec = lsqlin(eye(length(A_vec)), A_vec, [], [], Aeq, beq, lb, [],[], options);
                
                % Reshape the optimized vector back into a matrix
                Output_new = reshape(A_adjusted_vec, m, n);
         
    Output_new(find(Output_new<1))=0;
    
    Supply_new=Output_new-Export_state_new;

%% check if OUTPUT < VA
if sum(any(Output_new-VA_new<0))

Output_new(find(Output_new-VA_new<0));

end

Supply_new=Supply_new./sum(Supply_new,1).* IOT_Supply';
Demand_new=Demand_new./sum(Demand_new,1).* IOT_Demand';

SD_G=[Supply_new(:),Demand_new(:)];
SD_G(isnan(SD_G))=1;

n = 32;
%% State-wise trade estimation
for i= 1:66
        
        Z=SD_G((i-1)*n+1:(i-1)*n+n,:);
        n1=0.00000001;
        OldFolder = pwd;
        [Z_estim]=India_state(Z,n1,n);
        cd(OldFolder);
        
        for ii= 1:32
        
            order(ii,1)=str2num(Z_estim{1+ii,1});
        
        end
        
        Z_estim_1= [order,cell2mat(Z_estim(2:end,2:end))];
        
        
        Z_Sec1((i-1)*n+1:(i-1)*n+n,:)=sortrows(Z_estim_1,1);;
end

  Z_Sec = Z_Sec1(:,2:end);

  Z_Sec(find(Z_Sec<0.01))=0;
 
%%
for i = 1:66

   Outflow(:,i) =Z_Sec((i-1)*32+1:(i-1)*32+32,2);

   Inflow(:,i) =Z_Sec((i-1)*32+1:(i-1)*32+32,4);

end

Outflow =Outflow';
Inflow = Inflow';

%% SRIO estiamte

%%  Cross Entropy £¨with negetive sign£©GRAS
for i=1:32
    
    F_Province_total = sum(VA_new(i,:),2) - (sum(Outflow(:,i))+ sum(transpose(Export_state_new(i,:))) - sum(Inflow(:,i)) - sum(transpose(Import_state_new(i,:))));

    F_Province = F_struct .* F_Province_total;

    Row_Constraint= transpose(Output_new(i,:))- ((Outflow(:,i))+ (transpose(Export_state_new(i,:))) - (Inflow(:,i)) - (transpose(Import_state_new(i,:)))) ;
    Column_Constraint=[(Output_new(i,:)-VA_new(i,:)),F_Province_total];

    Row_Constraint(find(abs(Row_Constraint)<0.1))=0;


    if any(Row_Constraint<0)

        disp("Row negative")
        i
        check = [transpose(Output_new(i,:)),(Outflow(:,i)),(transpose(Export_state_new(i,:))),(Inflow(:,i)),(transpose(Import_state_new(i,:)))];
        
        return
    end

      if any(Column_Constraint<0)

        disp("Colunm negative")
        i
        check = [transpose(Output_new(i,:)),VA_new(i,:)'];
        
        return
    end


    Row_Constraint(find(Row_Constraint<0.01))=0;
    Column_Constraint(find(Column_Constraint<0.01&Column_Constraint>0))=0;
    
    OldFolder = pwd;  
    check(:,i)=sum(Row_Constraint,1)-sum(Row_Constraint,2);

    % scale 
    Row_Constraint=(Row_Constraint./sum(sum(Row_Constraint))).*sum(sum(Column_Constraint));

    Z1 = [IOT_Z,F_Province];

    Z2 = RAS_india(Z1,Row_Constraint,Column_Constraint,0.001);

    Z_new = [Z2(1:66,1:66),F_struct2.*Z2(1:66,67)]; % 1-71
    
    SRIO_raw{i} = [Z_new,Outflow(:,i),transpose(Export_state_new(i,:)),Inflow(:,i),transpose(Import_state_new(i,:)),Output_new(i,:)';VA_new(i,:),zeros(1,10);Output_new(i,:),zeros(1,10)];

    %SRIO_raw{i}=Optimization_India(Z1,Row_Constraint,Column_Constraint,66,67); 
   
    cd(OldFolder);
end

%% Converting to non-competitive table
for i=1:n
        Z_In=SRIO_raw{i};
        Ratio_FM=Z_In(1:66,75)./sum(Z_In(1:66,1:71),2);% Foreign import
        Ratio_PM=Z_In(1:66,74)./sum(Z_In(1:66,1:71),2);% Province import
            
        Ratio_PM(isnan(Ratio_PM))=0;
        Ratio_PM(isinf(Ratio_PM))=0;
        Ratio_FM(isnan(Ratio_FM))=0;
        Ratio_FM(isinf(Ratio_FM))=0;
               
        Ratio_D=1-Ratio_FM- Ratio_PM;
        Ratio_D1=Ratio_D;
        Ratio_D1(find(Ratio_D1>0))=0; 
        
        Ratio_D(find(Ratio_D<0))=0; 
        
       % Ratio_PM=Ratio_PM+Ratio_D1;
        Ratio_PM1=Ratio_PM;
        Ratio_PM1(find(Ratio_PM1>0))=0;
        Ratio_PM(find(Ratio_PM<0))=0; 
        
       % Ratio_FM=Ratio_FM+Ratio_PM1;  
        
        Z_D=Ratio_D.*Z_In(1:66,1:71);
        Z_P=Ratio_PM.*Z_In(1:66,1:71);

        Z_D(67,1:71)=sum(Ratio_FM.*Z_In(1:66,1:71),1);% Foreign
         
        Z_F(i,:)=  sum(Ratio_FM.*Z_In(1:66,1:71),1);% Foreign;

        Ratio_FM_1(:,i)=Ratio_FM;

        Z_D1{i}=Z_D;
        Z_PM{i}=Z_P;
        
 end

%% Gravity modelling

%%%%% Total supply and Total demand
for i= 1:n
    tra=SRIO_raw{i};
    TFM1((i-1)*66+1:(i-1)*66+66,1)=tra(1:66,72);%%% Outflow
    TFM1((i-1)*66+1:(i-1)*66+66,2)=tra(1:66,74); %%% Inflow
end


TFM1(find(TFM1<0.1))=0;
TFM1(isnan(TFM1))=0;

%% FOR REGRESSION

cd(['./internal_Transporte/',num2str(Year(t)),'/tab1']);

fileList_t = dir(fullfile(pwd, '*.xlsx'));

for i= 1:size(fileList_t,1)
    
   sample=xlsread(fileList_t(i).name);

   sample_1= MAP_STATE'*sample*MAP_STATE;

   sample_2{i} = sample_1(:);

end 

cd(OldFolder);
%% demand and supply for regression 27 commodity%%%%
supply_outflow=transpose(Gravity_model_mapping*reshape(TFM1(:,1),66,32));
demand_inflow=transpose(Gravity_model_mapping*reshape(TFM1(:,2),66,32));


%%%% Regression for 27 commodity%%%%
for i=1:27
   X_supply(:,i)=log(supply_outflow(:,i));
   X_demand(:,i)=log(demand_inflow(:,i));
end

% making symatric distance
x31 = log(d1+d1');
x3 = x31(:);

X_supply(isnan(X_supply))=0;
X_demand(isnan(X_demand))=0;
x3(isnan(x3))=0;

X_supply(isinf(X_supply))=0;
X_demand(isinf(X_demand))=0;
x3(isinf(x3))=0;

for i =1 :32
  X_supply_new((i-1)*32+1:(i-1)*32+32,:)  =repmat(X_supply(i,:),32,1);
end

X_demand_new=repmat(X_demand,32,1);

for i=1:27
  X=[ones(32*32,1),X_supply_new(:,i),X_demand_new(:,i),-x3];
  [b(i,:),bint,r,rint,stats(:,i)]=regress(sample_2{i},X);
  
  for n1=1:4
   
   [h,p{n1,i},ci,stats_ttest{n1,i}]=ttest(sample_2{i},X(:,n1));
  
   end
end

coefficient = Gravity_bridge'*b;

%% applying coefficient
TFM2=reshape(TFM1(:,1),66,32);
TFM3=reshape(TFM1(:,2),66,32);

TFM2=TFM2';
TFM3=TFM3';


for i= 1:size(coefficient,1)
     for j=1:n
       TFM2_1((j-1)*32+1:(j-1)*32+32,1) =repmat(TFM2(j,i),32,1);
     end

      Trade=coefficient(i,1)+coefficient(i,2).*log(TFM2_1)+coefficient(i,3).*log(repmat(TFM3(:,i),32,1))-coefficient(i,4).*x3;
    
      Trade=reshape(Trade,32,32);
      Trade(find(Trade<0))=0;
      Trade(isnan(Trade))=0; 
      


      Trade_sum{i}=Trade;
end

%% non-tangable goods
%% demand and supply for regression rest of commodity%%%%
TFM2=TFM2(:,(find(coefficient(:,1)==0))); %32*26
TFM2=TFM2(:);

TFM3=TFM3(:,(find(coefficient(:,1)==0)));
TFM3=TFM3(:);

for s = 1:26
    for i = 1:n
        for j = 1:n
            if i ~= j
                Flow_non(i, j,s) =  (TFM2((s-1)*32+i) .* TFM3((s-1)*32+j)) ./ (x31(i,j));
            end
        end
    end
end

%% aligning
Trade_sum{40}=Flow_non(:, :,1);% repair
Trade_sum{41}=Flow_non(:, :,2);% construction
Trade_sum{42}=Flow_non(:, :,3);% electricity

for i = 44:66
Trade_sum{i}=Flow_non(:, :,i-40); % i=43, gas has been assigned
end

%% OD matrix RAS
TFM2=reshape(TFM1(:,1),66,32);
TFM3=reshape(TFM1(:,2),66,32);

for i= 1:66

Trade_sum1=Trade_sum{i};
Trade_sum1(isinf(Trade_sum1))=0;
Trade_sum1(find(Trade_sum1==0))=0.0000000001;
  for k=1:n
            Trade_sum1(k,k)=0;
  end


      OD(:,:,i) = RAS_india(Trade_sum1,TFM2(i,:)',TFM3(i,:),10);


end


%% linking
for i=1:66
   

   RPC((i-1)*n+1:(i-1)*n+n,1:n)= OD(:,:,i)./sum(OD(:,:,i),1);
end

RPC(isnan(RPC))=0;
RPC(isinf(RPC))=0;


%Construct Intermediate AND Final Demand Matrix
for i=1:n
    Z_INT=Z_PM{i}; % Imported matrix
    
    Z_INT(find(Z_INT<0))=0;


    for j=1:size(Z_INT,1)
       
        Z_INT2((j-1)*n+1:(j-1)*n+n,(i-1)*71+1:(i-1)*71+71)=Z_INT(j,:).*RPC((j-1)*n+1:(j-1)*n+n,i);
       
    end
   
end



sum(Z_INT2,1);
    
%%%%convert to state-sector to state-sector
%%%%by row
for i=1:size(Z_INT,1)
for   j=1:n
       Z_INT3((j-1)*size(Z_INT,1)+i,:)=Z_INT2((i-1)*n+j,:);
end
end
Z_INT3(isnan(Z_INT3))=0;

%%%%%Domestic Matrix
for i=1:n
    Z_D2=Z_D1{i};
    Z_INT3((i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1),(i-1)*size(Z_INT,2)+1:(i-1)*size(Z_INT,2)+size(Z_INT,2))=Z_D2(1:size(Z_INT,1),1:size(Z_INT,2));
end

sum(Z_INT3,1);

%%% Change layout from 66X71 to 66X66 + 66X5
for i=1:n
    for j=1:n
    Z_INT4((i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1),(j-1)*size(Z_INT,1)+1:(j-1)*size(Z_INT,1)+size(Z_INT,1))=Z_INT3((i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1),(j-1)*size(Z_INT,2)+1:(j-1)*size(Z_INT,2)+size(Z_INT,1));
    F_IN4((i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1),(j-1)*5+1:(j-1)*5+5)=Z_INT3((i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1),(j-1)*size(Z_INT,2)+size(Z_INT,1)+1:(j-1)*size(Z_INT,2)+size(Z_INT,2));
    end
end
Table_raw=[Z_INT4,F_IN4];

%%% add Export and outflow 
%outflow and %export
for i=1:n
tra=SRIO_raw{i};
tra=[tra(1:size(Z_INT,1),73)];
Table_raw((i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1),n*size(Z_INT,1)+n*5+1)=tra;
end

%%%Import
for i=1:n
IMMM=Z_D1{i};
IMMM1=IMMM(size(Z_INT,1)+1,1:size(Z_INT,1));%foreign
IMF=IMMM(size(Z_INT,1)+1,size(Z_INT,1)+1:size(Z_INT,2));%imported final foreign
Table_raw(size(Z_INT,1)*n+1,(i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1))=IMMM1;
Table_raw(size(Z_INT,1)*n+1,(i-1)*5+n*size(Z_INT,1)+1:(i-1)*5+n*size(Z_INT,1)+5)=IMF;
end


for i=1:n
    IOT=SRIO_raw{i};
    
    Table_raw(n*size(Z_INT,1)+2,(i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1))=IOT(size(Z_INT,1)+1,1:size(Z_INT,1));%%%% VA
    
    Table_raw(n*size(Z_INT,1)+3,(i-1)*size(Z_INT,1)+1:(i-1)*size(Z_INT,1)+size(Z_INT,1))=IOT(size(Z_INT,1)+2,1:size(Z_INT,1));%%%% Oputput
end

%Table_raw(find(abs(Table_raw)<0.1))=0;
%% check
max(Table_raw(66*32+3,1:66*32)-sum(Table_raw(1:66*32,1:66*32),1)-sum(Table_raw(66*32+1:66*32+2,1:66*32),1)) %column check

error=Table_raw(66*32+3,1:66*32)'- sum(Table_raw(1:66*32,1:66*32),2)-sum(Table_raw(1:66*32,66*32+1:66*32+32*5+1),2);%row check

Table_raw(1:66*32,66*32+5*32+2)=error;%error

Table_raw(1:66*32,66*32+5*32+3)=Table_raw(66*32+3,1:66*32)';%output

rate=error./Table_raw(66*32+3,1:66*32)';
rate(isnan(rate))=0;
rate(isinf(rate))=0;

if any(abs(rate)>0.1)

max(abs(rate))
 disp("error is too large, check")

% RAS
%Row_Constraint=Table_raw(1:32*66,end)-Table_raw(1:32*66,end-2);
%Column_Constraint=Table_raw(66*32+3,1:66*32+32*5)-sum(Table_raw(1:66*32,1:66*32+32*5),1)-sum(Table_raw(66*32+1:66*32+2,1:66*32+32*5),1);
%Z2 = RAS1(Table_raw(1:32*66,1:32*66+32*5),Row_Constraint,Column_Constraint,0.01);

end

%% SRIO 

for i=1:32
    xlswrite(['./SRIO/India_SRIO_',num2str(Year(t)),'.xlsx'],Frame_SRIO,Frame_State{i});
    xlswrite(['./SRIO/India_SRIO_',num2str(Year(t)),'.xlsx'],SRIO_raw{i},Frame_State{i},'C3');
end

%% assmbly

Table_raw(1:66*32,66*32+5*32+2)=error;%error

Table_raw(1:66*32,66*32+5*32+3)=Table_raw(66*32+3,1:66*32)';%output

xlswrite(['India_MRIO_',num2str(Year(t)),'.xlsx'],FrameFull,'MRIO'); 

xlswrite(['India_MRIO_',num2str(Year(t)),'.xlsx'],Table_raw,'MRIO','C3');

clearvars -except Year MAP_STATE GSVA_all Import_national_ad1 Frame_SRIO Frame_State Gravity_model_mapping Gravity_bridge d1 FrameFull OldFolder fileList n_state MAP_State 


end

