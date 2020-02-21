
function varargout = Auto2_1s(varargin)


gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto2_1s_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto2_1s_OutputFcn, ...
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


% --- Executes just before Auto2_1s is made visible.
function Auto2_1s_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat']);
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);




% --- Outputs from this function are returned to the command line.
function varargout = Auto2_1s_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;

function myplot(MP,TITLE_NAME_DEFINE,FIGURE_INDEX,DATA_INDEX)
global Fileadress
j=DATA_INDEX;
n=FIGURE_INDEX;
  zihao=15;
for i=1:6
    h=figure;
    set(h,'visible','off');
    plot(MP{1,j}(:,1),MP{1,j}(:,2),'linewidth',2);
    grid on;
    xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao); 
    title(TITLE_NAME_DEFINE{i},'FontSize',zihao);
    axis([0 max(MP{1,j}(:,1)*1.1) 0 max(MP{1,j}(:,2))*1.1]);
    sfilename1=[Fileadress,num2str(n),'.jpg'];
    saveas(h,sfilename1);
    close(h)
    n=n+1;
    j=j+1;
end

function myplotelse(MP,TITLE_NAME_DEFINE,FIGURE_INDEX,DATA_INDEX)
global Fileadress
j=DATA_INDEX;
n=FIGURE_INDEX;
  zihao=15;
for i=1:3
    h=figure;
    set(h,'visible','off');
    plot(MP{1,j}(:,1),MP{1,j}(:,2),'linewidth',2);
    grid on;
    xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao); 
    title(TITLE_NAME_DEFINE{i},'FontSize',zihao);
    axis([0 max(MP{1,j}(:,1)*1.1) 0 max(MP{1,j}(:,2))*1.1]);
    sfilename1=[Fileadress,num2str(n),'.jpg'];
    saveas(h,sfilename1);
    close(h)
    n=n+1;
    j=j+1;
end

% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
   
% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
val1=get(handles.popupmenu1,'Value');
clear global Verformungteil1 Verformungteil2 Verformungteil3
global Fileadress


[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');
CHECK1=get(handles.checkbox1,'Value');
CHECK2=get(handles.checkbox2,'Value');
CHECK3=get(handles.checkbox3,'Value');
CHECK4=get(handles.checkbox4,'Value');
DATA_NUMBER=30;
if CHECK1==1
DATA_NUMBER=DATA_NUMBER+6;
end
if CHECK2==1
DATA_NUMBER=DATA_NUMBER+6;
end
if CHECK3==1
DATA_NUMBER=DATA_NUMBER+3;
end
if CHECK4==1
DATA_NUMBER=DATA_NUMBER+3;
end

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif length(filename)~=DATA_NUMBER
    msgbox('导入文件失败,缺少某个角度试验数据');
   return;
else
    zihao=15;%所有图标的字号
end
 if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
 
 t1=waitbar(0,'正在读入数据');
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

   Fileadress=strcat(pathname,'result\');
   t3=waitbar(0,'正在生成报告图片');
switch val1
    case 1
   TITLE_NAME30={'1#Zug H0°V0°';'2#Zug H0°V0°';'3#Zug H0°V0°';'1#Druck H0°V0°';'2#Druck H0°V0°';'3#Druck H0°V0°';...
       '1#Zug H30°V5°';'2#Zug H30°V5°';'3#Zug H30°V5°';'1#Druck H30°V5°';'2#Druck H30°V5°';'3#Druck H30°V5°';...
       '1#Zug H30°V-5°';'2#Zug H30°V-5°';'3#Zug H30°V-5°';'1#Druck H30°V-5°';'2#Druck H30°V-5°';'3#Druck H30°V-5°';...
       '1#Zug H-30°V5°';'2#Zug H-30°V5°';'3#Zug H-30°V5°';'1#Druck H-30°V5°';'2#Druck H-30°V5°';'3#Druck H-30°V5°';...
       '1#Zug H-30°V-5°';'2#Zug H-30°V-5°';'3#Zug H-30°V-5°';'1#Druck H-30°V-5°';'2#Druck H-30°V-5°';'3#Druck H-30°V-5°';};
  
    case 2
    TITLE_NAME30={'1#Zug H0°V0°';'2#Zug H0°V0°';'3#Zug H0°V0°';'1#Druck H0°V0°';'2#Druck H0°V0°';'3#Druck H0°V0°';...
       '1#Zug H30°V5°';'2#Zug H30°V5°';'3#Zug H30°V5°';'1#Druck H35°V5°';'2#Druck H35°V5°';'3#Druck H35°V5°';...
       '1#Zug H30°V-5°';'2#Zug H30°V-5°';'3#Zug H30°V-5°';'1#Druck H35°V-5°';'2#Druck H35°V-5°';'3#Druck H35°V-5°';...
       '1#Zug H-30°V5°';'2#Zug H-30°V5°';'3#Zug H-30°V5°';'1#Druck H-35°V5°';'2#Druck H-35°V5°';'3#Druck H-35°V5°';...
       '1#Zug H-30°V-5°';'2#Zug H-30°V-5°';'3#Zug H-30°V-5°';'1#Druck H-35°V-5°';'2#Druck H-35°V-5°';'3#Druck H-35°V-5°';};
end
   %% H0 V0 zug
   n=1;
   for j=1:10
for i=1:3
    h=figure;
     set(h,'visible','off');
  plot(MP{1,n}(:,1),MP{1,n}(:,2),'linewidth',2);
  grid on;
  zihao1=15;
 xlabel('Weg(mm)','FontSize',zihao1);ylabel('Kraft(N)','FontSize',zihao1); 
  title(TITLE_NAME30{n},'FontSize',zihao);
  axis([0 max(MP{1,n}(:,1)*1.1) 0 max(MP{1,n}(:,2))*1.1]);
    y_val=get(gca,'YTick');   %为了获得y轴句柄
y_str=num2str(y_val');    %为了将数字转换为字符数组
set(gca,'YTickLabel',y_str);    %显示
    sfilename1=[Fileadress,num2str(n),'-',TITLE_NAME30{n},'.jpg'];
saveas(h,sfilename1);
close(h)
  
  %计算变形
  x=MP{1,n}(:,1); y=MP{1,n}(:,2);
L=length(y);
[pks1 locs1]=findpeaks(y,'minpeakdistance',L/6);
[pks2 locs2]=findpeaks(-y,'minpeakdistance',L/6);
erster_unterlast(n,1)=x(locs1(1),1);
erster_bleibend(n,1)=x(locs2(2),1);
letzter_unterlast(n,1)=x(locs1(5),1);
letzter_bleibend(n,1)=x(L,1);
letzter_vorher(n,1)=x(locs2(end),1);
waitbar(n/30);
n=n+1;

end

   end
close(t3)



 
t2=waitbar(0,'正在生成报告数据');
%% 重新整理变形量
n=1;
for i=1:10
    Verformungteil1(i,1)=erster_unterlast(n,1);
    Verformungteil1(i,2)=erster_bleibend(n,1);
    Verformungteil1(i,3)=letzter_vorher(n,1);
    Verformungteil1(i,4)=letzter_unterlast(n,1);
    Verformungteil1(i,5)=letzter_bleibend(n,1);
n=n+3;
end
waitbar(0.1);
n=2;
for i=1:10
    Verformungteil2(i,1)=erster_unterlast(n,1);
    Verformungteil2(i,2)=erster_bleibend(n,1);
    Verformungteil2(i,3)=letzter_vorher(n,1);
    Verformungteil2(i,4)=letzter_unterlast(n,1);
    Verformungteil2(i,5)=letzter_bleibend(n,1);
n=n+3;
end
waitbar(0.2);
n=3;
for i=1:10
    Verformungteil3(i,1)=erster_unterlast(n,1);
    Verformungteil3(i,2)=erster_bleibend(n,1);
    Verformungteil3(i,3)=letzter_vorher(n,1);
    Verformungteil3(i,4)=letzter_unterlast(n,1);
    Verformungteil3(i,5)=letzter_bleibend(n,1);
n=n+3;
end
waitbar(0.3);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%H+-45 V45%%%%%%%%%%%%%%%%%%%5
if get(handles.checkbox1,'Value')==1
       TITLE_NAME_DEFINE1={'1#Zug H45°V-45°';'2#Zug H45°V-45°';'3#Zug H45°V-45°';'1#Zug H-45°V-45°';'2#Zug H-45°V-45°';'3#Zug H-45°V-45°';};
       FIGURE_INDEX=31;
       DATA_INDEX=length(erster_unterlast)+1;
       myplot(MP,TITLE_NAME_DEFINE1,FIGURE_INDEX,DATA_INDEX)
       
       %%计算变形
for i=1:6
     x=MP{1,DATA_INDEX}(:,1); y=MP{1,DATA_INDEX}(:,2);
     L=length(y);
     erster_unterlast_N(i,1)=find(y==max(y));
     erster_unterlast(DATA_INDEX,1)=x(erster_unterlast_N(i,1),1);
     erster_bleibend(DATA_INDEX,1)=x(L,1);
     DATA_INDEX=DATA_INDEX+1;
end       
end
waitbar(0.5);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

%%%%%%%%%%%%%%%%%%%%%%%%%%%%H35 V-35%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if get(handles.checkbox2,'Value')==1
       TITLE_NAME_DEFINE2={'1#Zug H35°V-35° 50%Fzleer EG+100kg';'2#Zug H35°V-35° 50%Fzleer EG+100kg';'3#Zug H35°V-35° 50%Fzleer EG+100kg';...
           '1#Zug H35°V-35° 75%Fzleer EG+100kg';'2#Zug H35°V-35° 75%Fzleer EG+100kg';'3#Zug H35°V-35° 75%Fzleer EG+100kg';};
       FIGURE_INDEX=37;
       DATA_INDEX=length(erster_unterlast)+1;
       myplot(MP,TITLE_NAME_DEFINE2,FIGURE_INDEX,DATA_INDEX)
       
       %%计算变形
 V=31;
 for i=1:3
        x=MP{1,DATA_INDEX}(:,1); y=MP{1,DATA_INDEX}(:,2);
L=length(y);
L5=L/3;
x1=x(1:L5);y1=y(1:L5);
x2=x(L5+1:2*L5);y2=y(L5+1:2*L5);
x3=x(2*L5+1:L);y3=y(2*L5+1:L);
erster_unterlast_N(i,1)=find(y1==max(y1),1);
erster_unterlast(DATA_INDEX,1)=x1(erster_unterlast_N(i,1),1);
erster_bleibend_N(i,1)=find(y2==min(y2),1);
erster_bleibend(DATA_INDEX,1)=x2(erster_bleibend_N(i,1),1);
letzter_unterlast_N(i,1)=find(y3==max(y3),1);
letzter_unterlast(V,1)=x3(letzter_unterlast_N(i,1),1);
letzter_bleibend(V,1)=x(L,1);
letzter_vorher_N(i,1)=find(y3==min(y3(1:L5/2)),1);
letzter_vorher(V,1)=x3(letzter_vorher_N(i,1),1);
      DATA_INDEX=DATA_INDEX+1; 
      V=V+1;
 end   
       
for i=1:3
     x=MP{1,DATA_INDEX}(:,1); y=MP{1,DATA_INDEX}(:,2);
     L=length(y);
     erster_unterlast_N(i,1)=find(y==max(y),1);
     erster_unterlast(DATA_INDEX,1)=x(erster_unterlast_N(i,1),1);
     erster_bleibend(DATA_INDEX,1)=x(L,1);
     DATA_INDEX=DATA_INDEX+1;
end       
end



waitbar(0.7);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%H0 V0%%%%%%%%%%%%%%%%%%%5
if get(handles.checkbox3,'Value')==1
       TITLE_NAME_DEFINE3={'1#Zug H0°V0° 130%F_z_G_G';'2#Zug H0°V0° 130%F_z_G_G';'3#Zug H0°V0° 130%F_z_G_G';};
       FIGURE_INDEX=43;
       DATA_INDEX=length(erster_unterlast)+1;
       myplotelse(MP,TITLE_NAME_DEFINE3,FIGURE_INDEX,DATA_INDEX)
       
       %%计算变形
for i=1:3
     x=MP{1,DATA_INDEX}(:,1); y=MP{1,DATA_INDEX}(:,2);
     L=length(y);
     erster_unterlast_N(i,1)=find(y==max(y),1);
     erster_unterlast(DATA_INDEX,1)=x(erster_unterlast_N(i,1),1);
     erster_bleibend(DATA_INDEX,1)=x(L,1);
     DATA_INDEX=DATA_INDEX+1;
end       
end
waitbar(0.8);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%H0 V0%%%%%%%%%%%%%%%%%%%
if get(handles.checkbox4,'Value')==1
       TITLE_NAME_DEFINE4={'1#Zug H0°V0° 200%F_z_G_G';'2#Zug H0°V0° 200%F_z_G_G';'3#Zug H0°V0° 200%F_z_G_G';};
       FIGURE_INDEX=46;
       DATA_INDEX=length(erster_unterlast)+1;
       myplotelse(MP,TITLE_NAME_DEFINE4,FIGURE_INDEX,DATA_INDEX)
       
       %%计算变形
for i=1:3
     x=MP{1,DATA_INDEX}(:,1); y=MP{1,DATA_INDEX}(:,2);
     L=length(y);
     erster_unterlast_N(i,1)=find(y==max(y),1);
     erster_unterlast(DATA_INDEX,1)=x(erster_unterlast_N(i,1),1);
     erster_bleibend(DATA_INDEX,1)=x(L,1);
     DATA_INDEX=DATA_INDEX+1;
end       
end

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%



waitbar(1);
close(t2);
  


biaotihao=10;
He=180*0.94488188976377952755905511811024;
Wi=240;
filespec_user=[Fileadress,'report.doc'];
t7=waitbar(0,'正在生成报告');
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
Word.Visible =0; % 使word为可见；或set(Word, 'Visible', 1); 
%===打开word文件，如果路径下没有则创建一个空白文档打开========================
if exist(filespec_user,'file')
Document=Word.Documents.Open(filespec_user);
else
Document=Word.Documents.Add;
Document.SaveAs2(filespec_user);
end
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
waitbar(0.1);

headline='1# Messergebnis Druck-Zug/1#件压力和拉力测量结果';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落

headline='Tab.1:Versuchswinkel,entsprechende Belastung und Ergebnisse 试验角度，加载力值和测量结果';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小

Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落

%%建立数据表格
Tab1 = Document.Tables.Add(Selection.Range, 2+length(erster_unterlast)/3, 14);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
% 设置行高，列宽
lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc,lc,lc,lc,1.43*lc,1.2*lc,lc,1.05*lc,1.5*lc,1.25*lc,1.24*lc,1.25*lc,1.5*lc,2.25*lc];
%row_height = [28.5849,28.5849,28.5849,28.5849,25.4717,25.4717,32.8302,312.1698,17.8302,49.2453,14.1509,18.6792];
for i = 1:14
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:(2+length(erster_unterlast)/3)
    for j=1:14
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end


for i=1:7
DTI.Cell(1,i).Merge(DTI.Cell(2,i)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(1,8).Merge(DTI.Cell(1,10)); % 第一行第1个到第二行第一个合并
DTI.Cell(1,9).Merge(DTI.Cell(1,11)); % 第一行第1个到第二行第一个合并
% 定义表格标题内容
DTI.Cell(1,1).Range.Text = 'Nr.';
DTI.Cell(1,2).Range.Text = 'H(°)';
DTI.Cell(1,3).Range.Text = 'V(°)';
DTI.Cell(1,4).Range.Text = 'Zug';
DTI.Cell(1,5).Range.Text = 'Druck';
DTI.Cell(1,6).Range.Text = 'Kraft(N)';
DTI.Cell(1,7).Range.Text = 'Zyklen';
DTI.Cell(1,8).Range.Text = 'Verfomung:erster Zyklus';
DTI.Cell(1,9).Range.Text = 'Verfomung:letzter Zyklus';
DTI.Cell(1,10).Range.Text = 'Bewertung';
DTI.Cell(2,8).Range.Text = 'Vorher';
DTI.Cell(2,9).Range.Text = 'Unter Last';
DTI.Cell(2,10).Range.Text = 'Bleibend';
DTI.Cell(2,11).Range.Text = 'Vorher';
DTI.Cell(2,12).Range.Text = 'Unter Last';
DTI.Cell(2,13).Range.Text = 'Bleibend';
%输出序号
for i=3:12
    DTI.Cell(i,1).Range.Text = num2str(i-2);
end
%定义输出第2列、第3列角度、第4及第5列拉压选项、第7列循环数、第8列0
switch val1
    case 1
Hv=[0,0,30,30,30,30,-30,-30,-30,-30];
    case 2
Hv=[0,0,30,35,30,35,-30,-35,-30,-35];       
end
Vv=[0,0,5,5,-5,-5,5,5,-5,-5];

for i=3:12
    DTI.Cell(i,2).Range.Text =num2str(Hv(i-2));
end
for i=3:12
     DTI.Cell(i,3).Range.Text=num2str(Vv(i-2));
end
DTI.Cell(3,4).Range.Text = 'X';DTI.Cell(5,4).Range.Text = 'X';DTI.Cell(7,4).Range.Text = 'X';DTI.Cell(9,4).Range.Text = 'X';DTI.Cell(11,4).Range.Text = 'X';
DTI.Cell(4,5).Range.Text = 'X';DTI.Cell(6,5).Range.Text = 'X';DTI.Cell(8,5).Range.Text = 'X';DTI.Cell(10,5).Range.Text = 'X';DTI.Cell(12,5).Range.Text = 'X';

for i=3:12
     DTI.Cell(i,7).Range.Text='5x';
end
for i=3:12
     DTI.Cell(i,8).Range.Text='0';
end


for i=9:13
    for j=3:12
    DTI.Cell(j,i).Range.Text=num2str(Verformungteil1(j-2,i-8),'%.2f');
    end
end
n=31;
if CHECK1==1||CHECK2==1||CHECK3==1||CHECK4==1
for i=13:2+length(erster_unterlast)/3
  DTI.Cell(i,9).Range.Text=num2str(erster_unterlast(n,1),'%.2f');  
  DTI.Cell(i,10).Range.Text=num2str(erster_bleibend(n,1),'%.2f'); 
  n=n+3;
end
end
TAB1_INDEX1=13;
if CHECK1==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='45';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1+1,1).Range.Text=num2str(TAB1_INDEX1-1); 
   DTI.Cell(TAB1_INDEX1+1,2).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1+1,3).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1+1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1+1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1+1,8).Range.Text='0';   
   TAB1_INDEX1=TAB1_INDEX1+2; 
end

if CHECK2==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='35';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='-35';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='3x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1+1,1).Range.Text=num2str(TAB1_INDEX1-1); 
   DTI.Cell(TAB1_INDEX1+1,2).Range.Text='35';
   DTI.Cell(TAB1_INDEX1+1,3).Range.Text='-35';
   DTI.Cell(TAB1_INDEX1+1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1+1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1+1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,11).Range.Text=num2str(letzter_vorher(31,1),'%.2f'); 
   DTI.Cell(TAB1_INDEX1,12).Range.Text=num2str(letzter_unterlast(31,1),'%.2f'); 
   DTI.Cell(TAB1_INDEX1,13).Range.Text=num2str(letzter_bleibend(31,1),'%.2f'); 
   TAB1_INDEX1=TAB1_INDEX1+2;   
end
if CHECK3==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   TAB1_INDEX1=TAB1_INDEX1+1; 
end
if CHECK4==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   TAB1_INDEX1=TAB1_INDEX1+1; 
end



Selection.Start = Content.end;
Selection.TypeParagraph;
%插入图片
InlineShapes=Document.InlineShapes;
n=1;
for i=1:10
Teil1address{1,i}=[Fileadress,num2str(n),'-',TITLE_NAME30{n},'.jpg'];
n=n+3;
end
waitbar(0.2)

FIGURE_NUMBER_INDEX=1;
for i=1:10
handle=Selection.InlineShapes.AddPicture(Teil1address{1,i});
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
delete(Teil1address{1,i});
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
end
if CHECK1==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'31.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
handle=Selection.InlineShapes.AddPicture([Fileadress,'34.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+2;
delete([Fileadress,'31.jpg']);
delete([Fileadress,'34.jpg']);
end
if CHECK2==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'37.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
handle=Selection.InlineShapes.AddPicture([Fileadress,'40.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+2;
delete([Fileadress,'37.jpg']);
delete([Fileadress,'40.jpg']);
end
if CHECK3==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'43.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
delete([Fileadress,'43.jpg']);
end
if CHECK4==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'46.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
delete([Fileadress,'46.jpg']);
end


Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
waitbar(0.3);



%%2#件输出结果
headline='2# Messergebnis Druck-Zug/2#件压力和拉力测量结果';

Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小

Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落

headline='Tab.2:Versuchswinkel,entsprechende Belastung und Ergebnisse 试验角度，加载力值和测量结果';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小

Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落

%%建立数据表格
Tab2 = Document.Tables.Add(Selection.Range,2+length(erster_unterlast)/3,14);
DTI = Document.Tables.Item(2); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

for i = 1:14
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:2+length(erster_unterlast)/3
    for j=1:14
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end

for i=1:7
DTI.Cell(1,i).Merge(DTI.Cell(2,i)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(1,8).Merge(DTI.Cell(1,10)); % 第一行第1个到第二行第一个合并
DTI.Cell(1,9).Merge(DTI.Cell(1,11)); % 第一行第1个到第二行第一个合并
% 定义表格标题内容
DTI.Cell(1,1).Range.Text = 'Nr.';
DTI.Cell(1,2).Range.Text = 'H(°)';
DTI.Cell(1,3).Range.Text = 'V(°)';
DTI.Cell(1,4).Range.Text = 'Zug';
DTI.Cell(1,5).Range.Text = 'Druck';
DTI.Cell(1,6).Range.Text = 'Kraft(N)';
DTI.Cell(1,7).Range.Text = 'Zyklen';
DTI.Cell(1,8).Range.Text = 'Verfomung:erster Zyklus';
DTI.Cell(1,9).Range.Text = 'Verfomung:letzter Zyklus';
DTI.Cell(1,10).Range.Text = 'Bewertung';
DTI.Cell(2,8).Range.Text = 'Vorher';
DTI.Cell(2,9).Range.Text = 'Unter Last';
DTI.Cell(2,10).Range.Text = 'Bleibend';
DTI.Cell(2,11).Range.Text = 'Vorher';
DTI.Cell(2,12).Range.Text = 'Unter Last';
DTI.Cell(2,13).Range.Text = 'Bleibend';
%输出序号
for i=3:12
    DTI.Cell(i,1).Range.Text = num2str(i-2);
end
% %定义输出第2列、第3列角度、第4及第5列拉压选项、第7列循环数、第8列0
% Hv=[0,0,30,30,30,30,-30,-30,-30,-30];
% Vv=[0,0,5,5,-5,-5,5,5,-5,-5];

for i=3:12
    DTI.Cell(i,2).Range.Text =num2str(Hv(i-2));
end
for i=3:12
     DTI.Cell(i,3).Range.Text=num2str(Vv(i-2));
end
DTI.Cell(3,4).Range.Text = 'X';DTI.Cell(5,4).Range.Text = 'X';DTI.Cell(7,4).Range.Text = 'X';DTI.Cell(9,4).Range.Text = 'X';DTI.Cell(11,4).Range.Text = 'X';
DTI.Cell(4,5).Range.Text = 'X';DTI.Cell(6,5).Range.Text = 'X';DTI.Cell(8,5).Range.Text = 'X';DTI.Cell(10,5).Range.Text = 'X';DTI.Cell(12,5).Range.Text = 'X';

for i=3:12
     DTI.Cell(i,7).Range.Text='5x';
end
for i=3:12
     DTI.Cell(i,8).Range.Text='0';
end


for i=9:13
    for j=3:12
    DTI.Cell(j,i).Range.Text=num2str(Verformungteil2(j-2,i-8),'%.2f');
    end
end

n=32;
if CHECK1==1||CHECK2==1||CHECK3==1||CHECK4==1
for i=13:2+length(erster_unterlast)/3
  DTI.Cell(i,9).Range.Text=num2str(erster_unterlast(n,1),'%.2f');  
  DTI.Cell(i,10).Range.Text=num2str(erster_bleibend(n,1),'%.2f'); 
  n=n+3;
end
end
TAB1_INDEX1=13;
if CHECK1==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='45';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1+1,1).Range.Text=num2str(TAB1_INDEX1-1); 
   DTI.Cell(TAB1_INDEX1+1,2).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1+1,3).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1+1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1+1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1+1,8).Range.Text='0';   
   TAB1_INDEX1=TAB1_INDEX1+2; 
end

if CHECK2==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='35';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='-35';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='3x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1+1,1).Range.Text=num2str(TAB1_INDEX1-1); 
   DTI.Cell(TAB1_INDEX1+1,2).Range.Text='35';
   DTI.Cell(TAB1_INDEX1+1,3).Range.Text='-35';
   DTI.Cell(TAB1_INDEX1+1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1+1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1+1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,11).Range.Text=num2str(letzter_vorher(32,1),'%.2f'); 
   DTI.Cell(TAB1_INDEX1,12).Range.Text=num2str(letzter_unterlast(32,1),'%.2f'); 
   DTI.Cell(TAB1_INDEX1,13).Range.Text=num2str(letzter_bleibend(32,1),'%.2f'); 
   TAB1_INDEX1=TAB1_INDEX1+2; 
   
   
end
if CHECK3==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   TAB1_INDEX1=TAB1_INDEX1+1; 
end
if CHECK4==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   TAB1_INDEX1=TAB1_INDEX1+1; 
end

waitbar(0.4);

Selection.Start = Content.end;
Selection.TypeParagraph;
%插入图片
InlineShapes=Document.InlineShapes;
n=2;
for i=1:10
Teil2address{1,i}=[Fileadress,num2str(n),'-',TITLE_NAME30{n},'.jpg'];
n=n+3;
end
for i=1:10
handle=Selection.InlineShapes.AddPicture(Teil2address{1,i});
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
delete(Teil2address{1,i});
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
end
waitbar(0.5);

if CHECK1==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'32.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
handle=Selection.InlineShapes.AddPicture([Fileadress,'35.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+2;
delete([Fileadress,'32.jpg']);
delete([Fileadress,'35.jpg']);
end
if CHECK2==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'38.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
handle=Selection.InlineShapes.AddPicture([Fileadress,'41.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+2;
delete([Fileadress,'38.jpg']);
delete([Fileadress,'41.jpg']);
end
if CHECK3==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'44.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
delete([Fileadress,'44.jpg']);
end
if CHECK4==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'47.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
delete([Fileadress,'47.jpg']);
end

Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;

waitbar(0.6);

%%3#件输出
headline='3# Messergebnis Druck-Zug/3#件压力和拉力测量结果';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落

headline='Tab.3:Versuchswinkel,entsprechende Belastung und Ergebnisse 试验角度，加载力值和测量结果';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小

Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落

%%建立数据表格
Tab3 = Document.Tables.Add(Selection.Range,2+length(erster_unterlast)/3,14);
DTI = Document.Tables.Item(3); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

for i = 1:14
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:2+length(erster_unterlast)/3
    for j=1:14
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end

for i=1:7
DTI.Cell(1,i).Merge(DTI.Cell(2,i)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(1,8).Merge(DTI.Cell(1,10)); % 第一行第1个到第二行第一个合并
DTI.Cell(1,9).Merge(DTI.Cell(1,11)); % 第一行第1个到第二行第一个合并
% 定义表格标题内容
DTI.Cell(1,1).Range.Text = 'Nr.';
DTI.Cell(1,2).Range.Text = 'H(°)';
DTI.Cell(1,3).Range.Text = 'V(°)';
DTI.Cell(1,4).Range.Text = 'Zug';
DTI.Cell(1,5).Range.Text = 'Druck';
DTI.Cell(1,6).Range.Text = 'Kraft(N)';
DTI.Cell(1,7).Range.Text = 'Zyklen';
DTI.Cell(1,8).Range.Text = 'Verfomung:erster Zyklus';
DTI.Cell(1,9).Range.Text = 'Verfomung:letzter Zyklus';
DTI.Cell(1,10).Range.Text = 'Bewertung';
DTI.Cell(2,8).Range.Text = 'Vorher';
DTI.Cell(2,9).Range.Text = 'Unter Last';
DTI.Cell(2,10).Range.Text = 'Bleibend';
DTI.Cell(2,11).Range.Text = 'Vorher';
DTI.Cell(2,12).Range.Text = 'Unter Last';
DTI.Cell(2,13).Range.Text = 'Bleibend';
%输出序号
for i=3:12
    DTI.Cell(i,1).Range.Text = num2str(i-2);
end
% %定义输出第2列、第3列角度、第4及第5列拉压选项、第7列循环数、第8列0
% Hv=[0,0,30,30,30,30,-30,-30,-30,-30];
% Vv=[0,0,5,5,-5,-5,5,5,-5,-5];

for i=3:12
    DTI.Cell(i,2).Range.Text =num2str(Hv(i-2));
end
for i=3:12
     DTI.Cell(i,3).Range.Text=num2str(Vv(i-2));
end
DTI.Cell(3,4).Range.Text = 'X';DTI.Cell(5,4).Range.Text = 'X';DTI.Cell(7,4).Range.Text = 'X';DTI.Cell(9,4).Range.Text = 'X';DTI.Cell(11,4).Range.Text = 'X';
DTI.Cell(4,5).Range.Text = 'X';DTI.Cell(6,5).Range.Text = 'X';DTI.Cell(8,5).Range.Text = 'X';DTI.Cell(10,5).Range.Text = 'X';DTI.Cell(12,5).Range.Text = 'X';

for i=3:12
     DTI.Cell(i,7).Range.Text='5x';
end
for i=3:12
     DTI.Cell(i,8).Range.Text='0';
end


for i=9:13
    for j=3:12
    DTI.Cell(j,i).Range.Text=num2str(Verformungteil3(j-2,i-8),'%.2f');
    end
end


n=33;
if CHECK1==1||CHECK2==1||CHECK3==1||CHECK4==1
for i=13:2+length(erster_unterlast)/3
  DTI.Cell(i,9).Range.Text=num2str(erster_unterlast(n,1),'%.2f');  
  DTI.Cell(i,10).Range.Text=num2str(erster_bleibend(n,1),'%.2f'); 
  n=n+3;
end
end


TAB1_INDEX1=13;
if CHECK1==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='45';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1+1,1).Range.Text=num2str(TAB1_INDEX1-1); 
   DTI.Cell(TAB1_INDEX1+1,2).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1+1,3).Range.Text='-45';
   DTI.Cell(TAB1_INDEX1+1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1+1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1+1,8).Range.Text='0';   
   TAB1_INDEX1=TAB1_INDEX1+2; 
end

if CHECK2==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='35';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='-35';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='3x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1+1,1).Range.Text=num2str(TAB1_INDEX1-1); 
   DTI.Cell(TAB1_INDEX1+1,2).Range.Text='35';
   DTI.Cell(TAB1_INDEX1+1,3).Range.Text='-35';
   DTI.Cell(TAB1_INDEX1+1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1+1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1+1,8).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,11).Range.Text=num2str(letzter_vorher(33,1),'%.2f'); 
   DTI.Cell(TAB1_INDEX1,12).Range.Text=num2str(letzter_unterlast(33,1),'%.2f'); 
   DTI.Cell(TAB1_INDEX1,13).Range.Text=num2str(letzter_bleibend(33,1),'%.2f'); 
   TAB1_INDEX1=TAB1_INDEX1+2; 
   
   
end
if CHECK3==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   TAB1_INDEX1=TAB1_INDEX1+1; 
end
if CHECK4==1
   DTI.Cell(TAB1_INDEX1,1).Range.Text=num2str(TAB1_INDEX1-2); 
   DTI.Cell(TAB1_INDEX1,2).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,3).Range.Text='0';
   DTI.Cell(TAB1_INDEX1,4).Range.Text='X';
   DTI.Cell(TAB1_INDEX1,7).Range.Text='1x';
   DTI.Cell(TAB1_INDEX1,8).Range.Text='0';
   TAB1_INDEX1=TAB1_INDEX1+1; 
end
waitbar(0.7);
Selection.Start = Content.end;
Selection.TypeParagraph;
%插入图片
InlineShapes=Document.InlineShapes;
n=3;
for i=1:10
Teil3address{1,i}=[Fileadress,num2str(n),'-',TITLE_NAME30{n},'.jpg'];
n=n+3;
end
for i=1:10
handle=Selection.InlineShapes.AddPicture(Teil3address{1,i});
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
delete(Teil3address{1,i});
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
end
waitbar(0.8);
if CHECK1==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'33.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
handle=Selection.InlineShapes.AddPicture([Fileadress,'36.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+2;
delete([Fileadress,'33.jpg']);
delete([Fileadress,'36.jpg']);
end
if CHECK2==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'39.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
handle=Selection.InlineShapes.AddPicture([Fileadress,'42.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX+1).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+2;
delete([Fileadress,'39.jpg']);
delete([Fileadress,'42.jpg']);
end
if CHECK3==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'45.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
delete([Fileadress,'45.jpg']);
end
if CHECK4==1
handle=Selection.InlineShapes.AddPicture([Fileadress,'48.jpg']);
InlineShapes.Item(FIGURE_NUMBER_INDEX).Height=He;
InlineShapes.Item(FIGURE_NUMBER_INDEX).Width=Wi;
FIGURE_NUMBER_INDEX=FIGURE_NUMBER_INDEX+1;
delete([Fileadress,'48.jpg']);
end
waitbar(0.9);
Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
waitbar(4/4);
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='Audi拖钩拉力试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
close(t7);

winopen([Fileadress,'report.doc']);

% --------------------------------------------------------------------
function Untitled_1_Callback(hObject, eventdata, handles)
 dos([cd,'\interface\Auto2_1s接口格式.xlsx']);
% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)

function checkbox2_Callback(hObject, eventdata, handles)

function checkbox3_Callback(hObject, eventdata, handles)

function checkbox4_Callback(hObject, eventdata, handles)


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
