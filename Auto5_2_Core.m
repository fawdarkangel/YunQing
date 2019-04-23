function [p1,X2,X3,X4,X5,WEG_COL,H_Final_index,H_Final]=Auto5_2_Core(MP)
MP_DIFF=diff(MP(:,2));
%��F1,F2,F3,F4
%��F1
[MAX1_pks,MAX1_locs]=findpeaks(MP_DIFF,'minpeakdistance',length(MP_DIFF)/10);
X1=MAX1_locs(1);
Index2=find(MP_DIFF(X1:end)<0.1,1); %��һ��С��0.1��ֵΪF1����
X2=X1+Index2-1;  %F2������ʵֵ

Index3=find(MP_DIFF(X2:end)>0.3,1); %��һ������0.1��ֵΪF2����
X3=X2+Index3-1;  %F2������ʵֵ

[MAXlast_pks,MAXlast_locs]=findpeaks(-MP_DIFF,'minpeakdistance',length(MP_DIFF)/10);
Index4=MAXlast_locs(end);
Index4_1=find(MP_DIFF(X2:Index4)<0.01,1,'last'); 
X4=Index4_1-1;

X5=find(MP(:,1)>1,1,'last');
%��m
MPx{1}=MP(X2+20:X3-20,1:2); 
  [p_1,p_2]=polyfit(MPx{1}(:,1),MPx{1}(:,2),1);%p1(1,1)Ϊб��
  p1=p_1(1,1);
  
  %��H
  MAX_WEG=max(MP(:,1));
  MAX_INDEX=find(MP(:,1)==MAX_WEG,1);
  if floor(length(MP)/2)>MAX_INDEX
      forlast=MAX_INDEX;
  else
      forlast=floor(length(MP)/2);
  end  
DIS=MAX_WEG(1)/(forlast);
if mod(forlast,2) == 1
    forlast=forlast-1;
end
for i=1:forlast
    WEG_COL(i,1)=find(MP(:,1)>=floor(DIS*i*10000)/10000,1,'first');
    WEG_COL(i,2)=find(MP(:,1)>=floor(DIS*i*10000)/10000,1,'last');
end

for i=1:length(WEG_COL)
    H(i,1)= MP(WEG_COL(i,1),2)-MP(WEG_COL(i,2),2);
end
H_Middle=H(find(WEG_COL(:,2)<X4,1)+10:X3-10,1);
H_Final=max(H_Middle);
H_Final_index=find(H(:,1)==H_Final);
end