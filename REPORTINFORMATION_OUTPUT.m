function f=REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME)
 Fileaddress=[ '\\faw-vw\fs\org\PE\T-E-VC-2-2\������\12-���ݴ���ƽ̨\report_information.xlsx'];
[num text alldata]=xlsread(Fileaddress);
SZ=size(alldata,1);%SZΪ��ǰ����������
                       
Azuobiao=strcat('A',num2str(SZ+1));
            
OUTPUT_INFORMATION{1,1}=FAHRZEUGCODE;
OUTPUT_INFORMATION{1,2}=TEST_NAME;
LOCAL_ADDRESS=java.net.InetAddress.getLocalHost;
COMPUTER_NAME=char(LOCAL_ADDRESS.getHostName);
OUTPUT_INFORMATION{1,3}=COMPUTER_NAME;
OUTPUT_INFORMATION{1,4}=datestr(now,'yyyy');
OUTPUT_INFORMATION{1,5}=datestr(now,'mm');
OUTPUT_INFORMATION{1,6}=datestr(now,'dd');
 xlswrite([Fileaddress], OUTPUT_INFORMATION,'Tabelle1',[Azuobiao]);
%UNTITLED2 �˴���ʾ�йش˺�����ժҪ
%   �˴���ʾ��ϸ˵��


end

