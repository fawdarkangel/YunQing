function varargout = DVDberichter(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @DVDberichter_OpeningFcn, ...
                   'gui_OutputFcn',  @DVDberichter_OutputFcn, ...
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


% --- Executes just before DVDberichter is made visible.
function DVDberichter_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat']);
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
Cover = imread('DVDcruve.JPG');
axes(handles.axes3);
imshow(Cover);
axis off
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes DVDberichter wait for user response (see UIRESUME)
% uiwait(handles.figure1);

function edit1_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Outputs from this function are returned to the command line.
function varargout = DVDberichter_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);


%%创建统计文件输出文件夹
Fileaddress=char('D:\Autorepoter\DVDberichter.xls');
try
  if ~exist('D:\Autorepoter','dir')
      mkdir('D:\Autorepoter');
  end
  if ~exist([Fileaddress]) %创建统计文件
      xlswrite([Fileaddress],{'车型'},'Sheet1','A1');
       xlswrite([Fileaddress],{'Punkt'},'Sheet1','B1');
       xlswrite([Fileaddress],{'MaxWeg（mm）'},'Sheet1','C1');
       xlswrite([Fileaddress],{'Temperatur'},'Sheet1','D1');
       xlswrite([Fileaddress],{'日期'},'Sheet1','E1');
  end
    [num text alldata]=xlsread('D:\Autorepoter\DVDberichter.xls');
            SZ=size(alldata,1);%SZ为当前工作表行数
end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');

CK=get(handles.checkbox1,'Value');
if CK==1
    if length(filename)<10
        msgbox('曲线数目不足10条，请去掉复选框');
        return;
    end
  end

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
else
    t1=waitbar(0,'正在导入数据');
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
    MP_LENGTH=length(filename);
    close(t1);
  a=[0 100;8 100];%100标准型
  b=[8 0;8 100];%8标准线
    
              
   %%计算总变形及弹性变形
   t2=waitbar(0,'正在处理图片及数据');
     
  for i=1:length(filename)
      F1=MP{1,i}(:,2);
  g1=find(F1>=100,1,'First');
  if isempty(g1)
      g1=find(F1==max(F1));
  end
    Verformung(i,1)=MP{1,i}(g1,1);
    
    
  END_INDEX_1=find(MP{1,i}(g1:end,2)<0,1)-1;  %第一个力小于0的下标
 
  if isempty(END_INDEX_1)
    Verformung(i,2)=MP{1,i}(length(MP{1,i}),1);
  else
           END_INDEX(i)=g1+ END_INDEX_1-1;  
      Verformung(i,2)=MP{1,i}(END_INDEX(i),1);
  end
    
  end
  
  for j=1:length(Verformung)
      Verformung(j,3)=Verformung(j,1)-Verformung(j,2);
  end
  for k=1:length(filename)
  c(k,1)=max(MP{1,k}(:,1));
  end
    if max(c)<8
        cmax=9;
    else
        cmax=max(c)*1.1;
    end
  COLOR_INDEX=[204 0 0;204 189 0;58 204 0;0 204 180;0 49 204;97 0 204;204 0 151;...
      148 71 56;141 148 56;55 149 66;56 148 139;92 72 132;125 73 131;130 74 74;226 130 226;...
      99 23 99;57 231 78;22 188 42]/255;
  
  %%创建文件并出图
 
  if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
  end
   L=strcat(pathname,'result\');%合成保存图片路径Ｌ
   
   %%%%%%%%%%%%%%%%勾选复选框时两幅图情况%%%%%%%%%%%%
    if CK==1

 XUNHUANCISHU=1;
 for o=1:floor(MP_LENGTH/10)   
      h=figure;
  for i=1:10
   F1=MP{1,XUNHUANCISHU}(:,2); 
   g1=find(F1>=100,1,'First');
   if isempty(g1)
      g1=find(F1==max(F1));
   end
  g2=find(F1>=0.1,1,'First');
    set(h,'visible','off');
        plot(MP{1,XUNHUANCISHU}(g2:g1,1),MP{1,XUNHUANCISHU}(g2:g1,2),'linewidth',2,'color',COLOR_INDEX(i,:));%EXCEL第3列为X,第二列为Y轴画图 
         hold on;                        
           xlabel('Weg(mm)','FontSize',15);ylabel('Kraft(N)','FontSize',15);title('Kraft-Weg-Diagramm','FontSize',15);
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 cmax*1.1 0 110]);
        legend_str{i}=['MP',num2str(XUNHUANCISHU)];   
        waitbar(1/length(filename));      
        XUNHUANCISHU= XUNHUANCISHU+1;
     end
     
   plot(a(:,1),a(:,2),'--r');
    plot(b(:,1),b(:,2),'--r');
    hold off;
    legend(legend_str);
      sfilename=[L,'result',num2str(o),'.jpg'];
           saveas(h,sfilename);                         
        
      close(h);    
 
 end  
 if MP_LENGTH-floor(MP_LENGTH/10)*10~=0
  h=figure;
      for i=1:(MP_LENGTH-floor(MP_LENGTH/10)*10)
   F1=MP{1,XUNHUANCISHU}(:,2); 
   g1=find(F1>=100,1,'First');
   if isempty(g1)
      g1=find(F1==max(F1));
   end
  g2=find(F1>=0.1,1,'First');
    set(h,'visible','off');
        plot(MP{1,XUNHUANCISHU}(g2:g1,1),MP{1,XUNHUANCISHU}(g2:g1,2),'linewidth',2,'color',COLOR_INDEX(i,:));%EXCEL第3列为X,第二列为Y轴画图 
         hold on;                        
           xlabel('Weg(mm)','FontSize',15);ylabel('Kraft(N)','FontSize',15);title('Kraft-Weg-Diagramm','FontSize',15);
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 cmax*1.1 0 110]);
        legend_str2{i}=['MP',num2str( XUNHUANCISHU)];   
        waitbar(1/length(filename));      
        XUNHUANCISHU= XUNHUANCISHU+1;
     end
     
   plot(a(:,1),a(:,2),'--r');
    plot(b(:,1),b(:,2),'--r');
    hold off;
    legend(legend_str2);
      sfilename=[L,'result',num2str(ceil(MP_LENGTH/10)),'.jpg'];
           saveas(h,sfilename);                         
        
      close(h);    
 end
    %%%%%%不勾选两幅图复选框时%%%%%%%%%%%%%%  
  else
  h=figure(1);
     for i=1:length(filename)
   F1=MP{1,i}(:,2); 
   g1=find(F1>=100,1,'First');
   if isempty(g1)
      g1=find(F1==max(F1));
   end
  g2=find(F1>=0.1,1,'First');
    set(h,'visible','off');
        plot(MP{1,i}(g2:g1,1),MP{1,i}(g2:g1,2),'linewidth',2,'color',COLOR_INDEX(i,:));%EXCEL第3列为X,第二列为Y轴画图 
         hold on;                        
           xlabel('Weg(mm)','FontSize',15);ylabel('Kraft(N)','FontSize',15);title('Kraft-Weg-Diagramm','FontSize',15);
        grid on; set(gca, 'GridLineStyle' ,'-');axis([0 cmax*1.1 0 110]);
        legend_str{i}=['MP',num2str(i)];   
        waitbar(1/length(filename));
     end
     
   plot(a(:,1),a(:,2),'--r');
    plot(b(:,1),b(:,2),'--r');
    hold off;
    legend(legend_str);
      sfilename=[L,'result.jpg'];
           saveas(h,sfilename);                          
              close(h);     
      
     
      
       end
   Verformungstr=strcat(L,'Verformung.xls');
   xlswrite(Verformungstr,Verformung,'sheet1','A1');
   %%收集数据
  FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
 try
   MPwegmax=strcat('MP',num2str(find(Verformung(:,3)==max(Verformung(:,3)))));%找到最大变形点MPx
   Azuobiao=strcat('A',num2str(SZ+1));Bzuobiao=strcat('B',num2str(SZ+1));Czuobiao=strcat('C',num2str(SZ+1));Dzuobiao=strcat('D',num2str(SZ+1));Ezuobiao=strcat('E',num2str(SZ+1));%生成写入EXCEL单元坐标
   xlswrite([Fileaddress],{FAHRZEUGCODE},'Sheet1',[Azuobiao]);%写入A列车型名称
   xlswrite([Fileaddress],{MPwegmax},'Sheet1',[Bzuobiao]);%写入B列最大变点
      xlswrite([Fileaddress],max(Verformung(:,3)),'Sheet1',[Czuobiao]);%写入C列最大变形量
   xlswrite([Fileaddress],{'RT'},'Sheet1',[Dzuobiao]);%写入D列温度
   xlswrite([Fileaddress],{date},'Sheet1',[Ezuobiao]);%写入E列时间
  end
   close(t2);
   
   
  
 %% 生成Word报告%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   t3=waitbar(0,'正在生成Word报告');
   He=180*0.94488188976377952755905511811024*1.7683;
Wi=240*1.9681;
filespec_user=[pathname,'result\report.doc'];
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
headline='Einzelergebnis/具体结果';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Tab1 = Document.Tables.Add(Selection.Range,MP_LENGTH+2,5);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
DTI.Rows.Alignment='wdAlignRowCenter';
lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc*5.03,lc*2.5,lc*4.56,lc*2.63,lc*2.96];
waitbar(0.3);
for i = 1:5
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:(MP_LENGTH+2)
    for j=1:5
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end

DTI.Cell(1,2).Merge(DTI.Cell(1,5)); % 第一行第1个到第二行第一个合并
DTI.Cell(1,1).Merge(DTI.Cell(2,1)); % 第一行第1个到第二行第一个合并
DTI.Cell(3,5).Merge(DTI.Cell(MP_LENGTH+2,5)); % 第一行第1个到第二行第一个合并
DTI.Cell(1,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';%垂直居中
DTI.Cell(3,5).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';%垂直居中
DTI.Cell(1,1).Range.Text = 'Messpunkt';
DTI.Cell(1,2).Range.Text = 'Verformung[mm]';
DTI.Cell(2,2).Range.Text = 'Gesamt';
DTI.Cell(2,3).Range.Text = 'Bleibend';
DTI.Cell(2,4).Range.Text = 'Elastisch';
DTI.Cell(2,5).Range.Text = 'Soll Elast.Verf.';
DTI.Cell(3,5).Range.Text = '≤8.00';
for i=1:MP_LENGTH
    DTI.Cell(i+2,1).Range.Text = ['MP',num2str(i)];
end
waitbar(0.5);
for i=1:MP_LENGTH
    DTI.Cell(i+2,2).Range.Text =num2str(Verformung(i,1),'%.2f');
    DTI.Cell(i+2,3).Range.Text =num2str(Verformung(i,2),'%.2f');
     DTI.Cell(i+2,4).Range.Text =num2str(Verformung(i,3),'%.2f');
     if Verformung(i,3)>8
         DTI.Cell(i+2,4).Range.Font.Colorindex='wdRed';
    DTI.Cell(i+2,4).Range.Font.Bold=1;
     end
end
waitbar(0.8);

Selection.Start = Content.end;
Selection.TypeParagraph;
InlineShapes=Document.InlineShapes;

if CK==1
    for i=1:ceil(MP_LENGTH/10)
    handle=Selection.InlineShapes.AddPicture([pathname,'result\result',num2str(i),'.jpg']);
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
    Selection.Start = Selection.end;
Selection.TypeParagraph;

    end
else
handle=Selection.InlineShapes.AddPicture([pathname,'result\result.jpg']);
InlineShapes.Item(1).Height=He;
InlineShapes.Item(1).Width=Wi;
end

waitbar(1);
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%

TEST_NAME='DVD坑包强度试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
close(t3);
winopen(L);
winopen([pathname,'result\report.doc']);



clear;
   end


% --------------------------------------------------------------------
function about_Callback(hObject, eventdata, handles)



% --------------------------------------------------------------------
function Untitled_3_Callback(hObject, eventdata, handles)
dos('about.txt');


% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
