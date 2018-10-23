function varargout = Auto1_2(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto1_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto1_2_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
end
% End initialization code - DO NOT EDIT


% --- Executes just before Auto1_2 is made visible.
function Auto1_2_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat'])
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto1_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto1_2_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);

[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
  return;

else
ZIHAO_WENZI=10;%���������ֺ�
ZIHAO_TU=16;%����ͼƬ�ֺ�
end

 if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
 Fileadress=strcat(pathname,'result\');
 t1=waitbar(0,'���ڶ�������');
 for i=1:length(filename)
         Filename{i}=strcat(pathname,filename{i});
         [Type Sheet Format]=xlsfinfo(Filename{i}) ;
         sheet{i}=Sheet;
         MP{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
         waitbar(i/length(filename));
          try
       system('taskkill/IM excel.exe');
   end
 end
 close(t1);
 t2=waitbar(0,'�������ɱ���ͼƬ');
 %%%%%%%%%%%%%%%%%%%%%%%%%����ͼƬ%%%%%%%%%%%%%%%%%5
 for i=1:length(filename)
 START_INDEX(i)=find(MP{1,i}(:,2)>0.1,1); %��һ��������0��ֵ�����±꣬��������ʼ��
 MAX_INDEX(i)=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)));
END_INDEX_1=find(MP{1,i}(MAX_INDEX(i):end,2)<0,1)-1;  %������С��0���±�

if isempty(END_INDEX_1)
        END_INDEX(i)=length(MP{1,i});  %������������0�Ļ������һ��ֵΪ��ֹ��
else
END_INDEX(i)=END_INDEX_1+ MAX_INDEX(i)-1; %���һ��������0���±꣬��������ֹ��
end
MP_final{1,i}=MP{1,i}(START_INDEX:END_INDEX(i),1:2);                       %��ȡ�����������


END_INDEX_80(i)=find(MP_final{1,i}(:,2)>80,1)-1;                                %��Ϊ80Nʱ�±�
MP_final_80{1,i}=MP_final{1,i}(1:END_INDEX_80(i),1:2);                         %����0��80N����
END_INDEX_120(i)=find(MP_final{1,i}(:,2)>120,1)-1;                              %��Ϊ120Nʱ�±�
MP_final_120{1,i}=MP_final{1,i}(END_INDEX_80(i)+1:END_INDEX_120(i),1:2);%����80N��120N����

END_INDEX_200(i)=find(MP_final{1,i}(:,2)==max(MP_final{1,i}(:,2)));               %��Ϊ���ֵʱ�±�
MP_final_200{1,i}=MP_final{1,i}(END_INDEX_120(i)+1:END_INDEX_200(i),1:2);%��Ϊ120N�����ֵ����
 end
  
   RESOLUTION_HE=600;                                                               %����ͼƬ�߶�����
  RESOLUTION_WI=1300;                                                               %����ͼƬ�������
  
 for i=1:length(filename)
     h(i)=figure;
     set(h(i),'visible','off');
 Y120=[120 120];
 Y80=[80 80];
 Xm=max(MP_final{1,i}(:,1));                                                    %��������ֵ
 Ym_INDEX=find(MP_final{1,i}(:,2)==max(MP_final{1,i}(:,2)));    %�������ֵ
 X0=[0 Xm];
 x=[0 Xm*1.1 Xm*1.1 0];                                                 %80��120N���κ�����
y=[80 80 120 120];                                                          %80��120N����������

plot(MP_final{1,i}(:,1),MP_final{1,i}(:,2),'LineWidth',2);          %�����Weg-Kraft����
hold on

 set(h(i),'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]);
   set(h(i),'color','w')
        set(gca,'FontSize',ZIHAO_TU);
         xlabel('Weg/���Ӧ��[mm]','FontSize',ZIHAO_TU);ylabel('Kraft/��׼�غ�[N]','FontSize',ZIHAO_TU);  
           axis([0 Xm*1.1 0 240]);
 %%%%%%%%%�ж������Ƿ��йյ㣬������ͼ��ע��%%%%%%%%%%%%
      if all(diff(MP_final_80{1,i}(:,2))>0)
          text(max(MP_final{1,i}(:,1))/2,30,'Kein Wendepunkt unter 80N','FontWeight','bold','FontSize',ZIHAO_TU);
      else
          text(max(MP_final{1,i}(:,1))/2,30,'Es gibt Wendepunkt unter 80N','color','r','FontWeight','bold','FontSize',ZIHAO_TU);
      end
      if all(diff(MP_final_120{1,i}(:,2))>0)
           text(max(MP_final{1,i}(:,1))/10,100,'Keine negative Steigung ','FontWeight','bold','FontSize',ZIHAO_TU);
      else
          text(max(MP_final{1,i}(:,1))/10,100,'Es gibt negative Steigung','color','r','FontWeight','bold','FontSize',ZIHAO_TU);
      end
       if all(diff(MP_final_200{1,i}(:,2))>=0)
           text(max(MP_final{1,i}(:,1))/3,180,{'>120N Wendepunkt';'nur mit Steigung>=0'},'FontWeight','bold','FontSize',ZIHAO_TU);
      else
          text(max(MP_final{1,i}(:,1))/3,180,{'>120N Es gibt Wendepunkt';'mit Steigung<0'},'color','r','FontWeight','bold','FontSize',ZIHAO_TU);
       end
         grid on; set(gca, 'GridLineStyle' ,'-');
               title(['Kraft/Weg Kurve am MP',num2str(i),' Dachpoliersteifigkeit'],'FontSize',ZIHAO_TU);
               
       B=(MP_final{1,i}(Ym_INDEX,2))-MP_final{1,i}(Ym_INDEX,1)*200/15;                      %B�նȱ�׼15N/mm�ؾ�
       X_200N_STAND=[Xm/1.1 Xm*1.05];                                                                     %15N/mm�����ߺ�����
       Y_200N_STAND=200/15.*X_200N_STAND+B;                                                     %15N/mm������������
               plot(X_200N_STAND,Y_200N_STAND,'--m','LineWidth',2)                              %��15N/mm������
        %K=(MP_final{1,i}(Ym_INDEX,2)-MP_final{1,i}(Ym_INDEX-1,2))/ (MP_final{1,i}(Ym_INDEX,1)-MP_final{1,i}(Ym_INDEX-1,1));
        K=200/Xm;                                                                                                                 %200Nʵ��б��
       Y_200N_REAL=[190 210];                                                               %200N�����ߺ�����
          X_200N_REAL=[ Y_200N_REAL(1)/K  Y_200N_REAL(2)/K];                 %200N������������
       plot(X_200N_REAL,Y_200N_REAL,'--g','LineWidth',2);                                                %����ϸն�����
       patch(x,y,'b','linestyle','none','facealpha','0.3');                                                            %��80-120N͸������
       legend('Teil','C_1_5_N_/_m_m','C_T_e_i_l','Location','SouthEast');
       hold off;
       
    %STEIFIGKEIT(i)=MP_final{1,i}(Ym_INDEX,2)/MP_final{1,i}(Ym_INDEX,1);
    MAX_VERFORMUNG(i)=Xm;  %����������
    K200(i)=K;                                   %��ն�ֵ
    BLE_VERFORMUNG(i)=MP_final{1,i}(length(MP_final{1,i}),1);                                            %����������
        sfilename1=[Fileadress,num2str(i),'.jpg'];
     f=getframe(h(i));
           imwrite(f.cdata,sfilename1);
           close(h(i));
          waitbar(i/length(filename));
 end
       
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%����Word����%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
close(t2);
t3=waitbar(0,'��������Word����');
filespec_user=[Fileadress,'report.doc'];
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
Word.Visible =0; % ʹwordΪ�ɼ�����set(Word, 'Visible', 1); 
%===��word�ļ������·����û���򴴽�һ���հ��ĵ���========================
if exist(filespec_user,'file')
Document=Word.Documents.Open(filespec_user);
else
Document=Word.Documents.Add;
Document.SaveAs2(filespec_user);
end

Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
t3=waitbar(0.1);
%===�ĵ���ҳ�߾�===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
biaotihao=10;
headline=['III. Einzelergebnis ������'];
Content.Start=0; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=biaotihao; % �����С
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���
Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���

InlineShapes=Document.InlineShapes;
He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
biaotihao=10;

Tab1 = Document.Tables.Add(Selection.Range, length(filename)+1,5);
DTI = Document.Tables.Item(1); % �����
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����

lc=28.381133333333333333333333333333; %���׻���
column_width = [2.24*lc,3.51*lc,2.75*lc,3*lc,3.25*lc];
t3=waitbar(0.3);
for i = 1:5
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:length(filename)+1
    for j=1:5
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
        DTI.Cell(i,j).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
    end
end
t3=waitbar(0.6);
DTI.Cell(1,1).Range.Text = 'Messpunkt';
DTI.Cell(1,2).Range.Text = 'Steifigkeit bei 200N[N/mm]';
DTI.Cell(1,3).Range.Text = 'Soll- Steifigkeit [N/mm]';
DTI.Cell(1,4).Range.Text = 'max.Verformung[mm]';
DTI.Cell(1,5).Range.Text = 'bleibende Verformung[mm]';

DTI.Cell(2,3).Merge(DTI.Cell(length(filename)+1,3)); 
DTI.Cell(2,3).Range.Text = '>=15';



for i=1:length(filename)
DTI.Cell(i+1,1).Range.Text =['MP',num2str(i)];
DTI.Cell(i+1,2).Range.Text =num2str(K200(i),'%.2f');                                                  %���200Nʱ�ն�ֵ
 if K200(i)<15                                                                                                            %�жϸն��Ƿ�С��15��С������Ӵ�
             DTI.Cell(i+1,2).Range.Font.Colorindex='wdRed';
             DTI.Cell(i+1,2).Range.Font.Bold=1;
       end
DTI.Cell(i+1,4).Range.Text =num2str(MAX_VERFORMUNG(i),'%.2f');                           %���������
DTI.Cell(i+1,5).Range.Text =num2str(BLE_VERFORMUNG(i),'%.2f');                              %����������
end

Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
InlineShapes=Document.InlineShapes;
t3=waitbar(0.7);
for i=1:length(filename)
    sfilename1=[Fileadress,num2str(i),'.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
delete(sfilename1); 

end
t3=waitbar(0.9);
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
Word.Quit; % �ر��ĵ�
%%%%%%%%%%%%�������������Ϣ�������ռ�%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='Poliersteifigkeit';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
t3=waitbar(1);
close(t3);
winopen([Fileadress,'report.doc']);


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
