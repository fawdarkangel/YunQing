function varargout = AutoThird_1_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @AutoThird_1_1_OpeningFcn, ...
                   'gui_OutputFcn',  @AutoThird_1_1_OutputFcn, ...
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

function AutoThird_1_1_OpeningFcn(hObject, eventdata, handles, varargin)
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




% --- Outputs from this function are returned to the command line.
function varargout = AutoThird_1_1_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;
% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
[filename,pathname,fileindex]=uigetfile('*.csv','选择数据','MultiSelect','on');

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
else
    zihao=15;%所有图标的字号
end
 if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
 
 t1=waitbar(0,'正在读入数据');
 if ~iscell(filename)
     Filename{1}=strcat(pathname,filename);
     MP{1}=csvread(Filename{1},20,0);
     filename=1;
 else
     for i=1:length(filename)
         Filename{i}=strcat(pathname,filename{i});
         MP{i}=csvread(Filename{i},20,0);
         waitbar(i/length(filename));
     end
 end
 close(t1);
setappdata(0,'filename',filename);
setappdata(0,'MP',MP);
setappdata(0,'Filename',Filename);
set(handles.listbox1,'Value',1);
set(handles.listbox1,'String',filename);
for i=1:length(filename)
         STORM_STARTINDEX(i)=find(abs(MP{1,i}(:,2))>1,1);%找寻第一个绝对值大于1的点，认为是起始点
         MP_MIDDLE{i}=MP{1,i}(STORM_STARTINDEX(i):end,1:2);%中间过程变量
         MP_FINAL{1,i}(:,1)=MP_MIDDLE{1,i}(:,1)-MP_MIDDLE{1,i}(1,1);%最终曲线时间列
          MP_FINAL{1,i}(:,2)=MP_MIDDLE{1,i}(:,2);%最终曲线电流列
          if MP_FINAL{1,i}(1,2)<0
              MP_FINAL{1,i}(:,2)=MP_FINAL{1,i}(:,2).*(-1);%最终曲线电流列
          end
          
 end
for i=1:length(filename)
    MAX_Y=max(MP_FINAL{1,i}(:,2));
    MIN_Y=min(MP_FINAL{1,i}(:,2));
 A=diff(smooth(MP_FINAL{1,i}(:,2),150));

A=smooth(A,50,'lowess');
START=find(A<0);
 
TIME1_START_INDEX(i)=START(10);
TIME1_START_END=find(MP_FINAL{1,i}(TIME1_START_INDEX(i):end,2)>MAX_Y*0.75,1);  %%匹配参数，可以更改调节
TIME1_END_INDEX(i)=TIME1_START_INDEX(i)+TIME1_START_END-10;
TIME1(i)=MP_FINAL{1,i}(TIME1_END_INDEX(i),1)-MP_FINAL{1,i}(TIME1_START_INDEX(i),1);

TIME2_START_INDEX_1=find(MP_FINAL{1,i}(:,2)<-5,1);
B=diff(smooth(MP_FINAL{1,i}(TIME2_START_INDEX_1:end,2),150));
B=smooth(B,50,'lowess');
STARTB=find(B<0);

TIME2_START_INDEX(i)=TIME2_START_INDEX_1+STARTB(5)-1;%找寻第二次时间起始点
TIME2_START_END=find(MP_FINAL{1,i}(TIME2_START_INDEX(i):end,2)<MIN_Y*0.75,1); 
TIME2_END_INDEX(i)=TIME2_START_INDEX(i)+TIME2_START_END-1;

TIME2(i)=MP_FINAL{1,i}(TIME2_END_INDEX(i)-10,1)-MP_FINAL{1,i}(TIME2_START_INDEX(i),1);
 
end
DATA.TIME1_START_INDEX=TIME1_START_INDEX;
DATA.TIME1_END_INDEX=TIME1_END_INDEX;
DATA.TIME2_START_INDEX=TIME2_START_INDEX;
DATA.TIME2_END_INDEX=TIME2_END_INDEX;
setappdata(0,'DATA',DATA);
setappdata(0,'pathname',pathname);
setappdata(0,'MP_FINAL',MP_FINAL);
setappdata(0,'TIME1',TIME1);
setappdata(0,'TIME2',TIME2);

 msgbox('数据导入成功');
 
 
% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)

MP=getappdata(0,'MP');
filename=getappdata(0,'filename');
TIME1=getappdata(0,'TIME1');
TIME2=getappdata(0,'TIME2');
MP_FINAL=getappdata(0,'MP_FINAL');
DATA=getappdata(0,'DATA');
TIME1_START_INDEX=DATA.TIME1_START_INDEX;
TIME1_END_INDEX=DATA.TIME1_END_INDEX;
TIME2_START_INDEX=DATA.TIME2_START_INDEX;
TIME2_END_INDEX=DATA.TIME2_END_INDEX;


ZIHAO_TU_YULAN=10;
   for i=1:length(filename)
    YMAX(i)=max(MP{1,i}(:,2));
          
  end

Y_MAX=ceil(max(YMAX)*1.05);
 if mod(Y_MAX,2)==0
              Y_MAX=Y_MAX;
          else
              Y_MAX=Y_MAX+1;
 end
          
  for i=1:length(filename) 
 CRUVE_END_INDEX0=find(MP_FINAL{1,i}(:,2)<-1,1,'last')+1;%最后一个小于零的数
 TIME_END(i)=ceil(MP_FINAL{1,i}(CRUVE_END_INDEX0,1));
 
 if mod(TIME_END(i),2)==0
     TIME_END(i)=TIME_END(i)+1;
 else
    TIME_END(i)=TIME_END(i)+1;
 end
  end
 
  CHOOSE=get(handles.listbox1,'Value');                %listbox的值
  i=CHOOSE;     
  plot(handles.axes1,MP_FINAL{1,i}(:,1),MP_FINAL{1,i}(:,2),'linewidth',2);  
  hold on;
  plot(handles.axes1,MP_FINAL{1,i}(TIME1_START_INDEX(i),1),MP_FINAL{1,i}(TIME1_START_INDEX(i),2),'*','Color','r');
  plot(handles.axes1,MP_FINAL{1,i}(TIME1_END_INDEX(i),1),MP_FINAL{1,i}(TIME1_END_INDEX(i),2),'*','Color','r');
  plot(handles.axes1,MP_FINAL{1,i}(TIME2_START_INDEX(i),1),MP_FINAL{1,i}(TIME2_START_INDEX(i),2),'*','Color','r');
  plot(handles.axes1,MP_FINAL{1,i}(TIME2_END_INDEX(i)-10,1),MP_FINAL{1,i}(TIME2_END_INDEX(i)-10,2),'*','Color','r');
  
   datacursormode on
  ylim(handles.axes1,[-Y_MAX Y_MAX]);
  set(handles.axes1,'ytick',-Y_MAX:2:Y_MAX);
  xlim(handles.axes1,[0 TIME_END(i)]);
  set(handles.axes1,'xtick',0:2: TIME_END(i));
  xlabel(handles.axes1,'Zeit[s]','FontSize',ZIHAO_TU_YULAN);ylabel('Strom[A]','FontSize',ZIHAO_TU_YULAN);
  grid on;
  hold off;
 
set(handles.edit5,'String',num2str(TIME1(i)));
set(handles.edit6,'String',num2str(TIME2(i)));


function listbox1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


 
% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)

CK1=get(handles. checkbox1,'Value');

filename=getappdata(0,'filename');
MP=getappdata(0,'MP');
Filename=getappdata(0,'Filename');
pathname=getappdata(0,'pathname');
zihao=15;%所有图标的字号
 for i=1:length(filename)
         STORM_STARTINDEX(i)=find(abs(MP{1,i}(:,2))>1,1);%找寻第一个绝对值大于1的点，认为是起始点
         MP_MIDDLE{i}=MP{1,i}(STORM_STARTINDEX(i):end,1:2);%中间过程变量
         MP_FINAL{1,i}(:,1)=MP_MIDDLE{1,i}(:,1)-MP_MIDDLE{1,i}(1,1);%最终曲线时间列
          MP_FINAL{1,i}(:,2)=MP_MIDDLE{1,i}(:,2);%最终曲线电流列
          if MP_FINAL{1,i}(1,2)<0
              MP_FINAL{1,i}(:,2)=MP_FINAL{1,i}(:,2).*(-1);%最终曲线电流列
          end
          
 end

 if CK1==1                                  %显示电流时间
for i=1:length(filename)
    MAX_Y=max(MP_FINAL{1,i}(:,2));
    MIN_Y=min(MP_FINAL{1,i}(:,2));
 A=diff(smooth(MP_FINAL{1,i}(:,2),150));

A=smooth(A,50,'lowess');
START=find(A<0);
 
TIME1_START_INDEX(i)=START(10);
TIME1_START_END=find(MP_FINAL{1,i}(TIME1_START_INDEX(i):end,2)>MAX_Y*0.75,1);  %%匹配参数，可以更改调节
TIME1_END_INDEX(i)=TIME1_START_INDEX(i)+TIME1_START_END-10;
TIME1(i)=MP_FINAL{1,i}(TIME1_END_INDEX(i),1)-MP_FINAL{1,i}(TIME1_START_INDEX(i),1);

TIME2_START_INDEX_1=find(MP_FINAL{1,i}(:,2)<-5,1);
B=diff(smooth(MP_FINAL{1,i}(TIME2_START_INDEX_1:end,2),150));
B=smooth(B,50,'lowess');
STARTB=find(B<0);

TIME2_START_INDEX(i)=TIME2_START_INDEX_1+STARTB(5)-1;%找寻第二次时间起始点
TIME2_START_END=find(MP_FINAL{1,i}(TIME2_START_INDEX(i):end,2)<MIN_Y*0.75,1); 
TIME2_END_INDEX(i)=TIME2_START_INDEX(i)+TIME2_START_END-1;

TIME2(i)=MP_FINAL{1,i}(TIME2_END_INDEX(i)-10,1)-MP_FINAL{1,i}(TIME2_START_INDEX(i),1);
 
  end
  end
 
  Fileadress=strcat(pathname,'result\');
  t2=waitbar(0,'正在生成图片');
    for i=1:length(filename)
    YMAX(i)=max(MP{1,i}(:,2));
          
  end

Y_MAX=ceil(max(YMAX)*1.05);
 if mod(Y_MAX,2)==0
              Y_MAX=Y_MAX;
          else
              Y_MAX=Y_MAX+1;
 end
          
  for i=1:length(filename) 
 CRUVE_END_INDEX0=find(MP_FINAL{1,i}(:,2)<-1,1,'last')+1;%最后一个小于零的数
 TIME_END(i)=ceil(MP_FINAL{1,i}(CRUVE_END_INDEX0,1));
 
 if mod(TIME_END(i),2)==0
     TIME_END(i)=TIME_END(i)+1;
 else
    TIME_END(i)=TIME_END(i)+1;
 end
 end

 
 
  for i=1:length(filename)
      h=figure;
        set(h,'visible','off');
  plot(MP_FINAL{1,i}(:,1),MP_FINAL{1,i}(:,2),'linewidth',2);
  if CK1==1
  hold on;
  plot(MP_FINAL{1,i}(TIME1_START_INDEX(i),1),MP_FINAL{1,i}(TIME1_START_INDEX(i),2),'*','Color','r');
  plot(MP_FINAL{1,i}(TIME1_END_INDEX(i),1),MP_FINAL{1,i}(TIME1_END_INDEX(i),2),'*','Color','r');
    plot(MP_FINAL{1,i}(TIME2_START_INDEX(i),1),MP_FINAL{1,i}(TIME2_START_INDEX(i),2),'*','Color','r');
   plot(MP_FINAL{1,i}(TIME2_END_INDEX(i)-10,1),MP_FINAL{1,i}(TIME2_END_INDEX(i)-10,2),'*','Color','r');
  end
  ylim([-Y_MAX Y_MAX]);
set(gca,'ytick',-Y_MAX:2:Y_MAX);
xlim([0 TIME_END(i)]);
set(gca,'xtick',0:2: TIME_END(i));
      xlabel('Zeit[s]','FontSize',zihao);ylabel('Strom[A]','FontSize',zihao); 
      grid on;
      sfilename1=[Fileadress,num2str(i),'-Strom.jpg'];
saveas(h,sfilename1);
            close(h);
            waitbar(i/length(filename));
  end
  close(t2);
  
  t3=waitbar(0,'正在生成Word报告');
   filespec_user=[Fileadress,'Strom_report.doc'];
   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%生成Word报告%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
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
He=180*0.94488188976377952755905511811024;
Wi=240;
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
  waitbar(0.3);
  InlineShapes=Document.InlineShapes;
   headline='III. Einzelergebnis 具体结果';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落
n=1;
  for i=1:length(filename)
    sfilename1=[Fileadress,num2str(i),'-Strom.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
if mod(i,2)==0
   Selection.Start = Selection.end;
Selection.TypeParagraph; 

 headline=['                          Figure',num2str(n),'                                                      Figure',num2str(n+1)];
Selection.Text=headline;
Selection.Font.Size=8; % 字体大小
Selection.Font.NameAscii='Arial';
 Selection.Start = Selection.end;
Selection.TypeParagraph; 
n=n+2;
end


waitbar(0.8)
delete(sfilename1); 
  end

  
if CK1==1
     Selection.Start = Selection.end;
Selection.TypeParagraph;
    Tab1 = Document.Tables.Add(Selection.Range,length(filename)+1,3);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc*2,lc*2,lc*2];   
    for i = 1:3
DTI.Columns.Item(i).Width = column_width(i);
    end

     DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Range.Font.NameAscii='Arial';
        
 DTI.Cell(1,1).Range.Text = 'Pkt.';
 DTI.Cell(1,2).Range.Text = 't1[s]';
 DTI.Cell(1,3).Range.Text = 't2[s]';
 for i=1:length(filename)
     DTI.Cell(i+1,1).Range.Text = num2str(i);
     DTI.Cell(i+1,2).Range.Text = num2str(TIME1(i),'%.2f');
     DTI.Cell(i+1,3).Range.Text = num2str(TIME2(i),'%.2f');
     
 end
   end

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='玻璃升降寿命试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
winopen(filespec_user);
waitbar(1);
close(t3);


function edit1_Callback(hObject, eventdata, handles)

function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

function edit2_Callback(hObject, eventdata, handles)

function edit2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

function edit3_Callback(hObject, eventdata, handles)

% --- Executes during object creation, after setting all properties.
function edit3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit4_Callback(hObject, eventdata, handles)

function edit4_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
b1=get(handles.edit1,'String');
b2=get(handles.edit2,'String');
b3=get(handles.edit3,'String');
b4=get(handles.edit4,'String');
   val1=get(handles.popupmenu1,'Value');
    val2=get(handles.popupmenu3,'Value');
   switch val1
       case 1
           b3='1';b4='1';
       case 2
           b1='1';b2='1';
   end
if isempty(b1)||isempty(b2)||isempty(b3)||isempty(b4)
    msgbox('请输入堵转扭矩');
return;

else
    T_VOR=(str2num(b1)+str2num(b2))/2*0.35;
    T_HINTEN=(str2num(b3)+str2num(b4))/2*0.35;
end

[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on');

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
else
    zihao=10;%所有图标的字号
end
 if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
 t1=waitbar(0,'正在读入数据');
 switch val2
     case 1
         if ~iscell(filename)
             Filename{1}=strcat(pathname,filename);
             [Type Sheet Format]=xlsfinfo(Filename{1}) ;
             sheet{1}=Sheet;
             MP{1}=xlsread(Filename{1},char(sheet{1,1}(1,1)));
             filename=1;
         else
             for i=1:length(filename)
                 Filename{i}=strcat(pathname,filename{i});
                 [Type Sheet Format]=xlsfinfo(Filename{i}) ;
                 sheet{i}=Sheet;
                 MP{i}=xlsread(Filename{i},char(sheet{1,i}(1,1)));
                 waitbar(i/length(filename));
                 try
                     system('taskkill/IM excel.exe');
                 end
             end
         end
     case 2
         if iscell(filename)
             this_filename=filename;
             fileindex=length(this_filename);
         else
             fileindex=1;
             this_filename{1}=filename;
         end
         for i=1:fileindex
             Filename{i}=strcat(pathname, this_filename{i});
             fidin=fopen(Filename{i});                               % 打开test2.txt文件
             fidout=fopen('result.txt','w');                       % 创建MKMATLAB.txt文件
             while ~feof(fidin)                                      % 判断是否为文件末尾
                 tline=fgetl(fidin);                                     % 从文件读行
                 if isempty(tline)
                     continue
                 else
                     if double(tline(1))>=48&&double(tline(1))<=57       % 判断首字符是否是数值
                         fprintf(fidout,'%s\n',tline);                  % 如果是数字行，把此行数据写入文件MKMATLAB.txt
                         continue                                         % 如果是非数字继续下一次循环
                     end
                 end
             end
             fclose(fidout);
             MK=importdata('result.txt');
             MP{1,i}(:,6)=MK(:,1);
             MP{1,i}(:,2)=MK(:,2);
             waitbar(i/fileindex);
         end
 end
   try 
    fclose('all')
    delete('result.txt')
  end
 close(t1);
 
for i=1:length(filename)
    YMAX(i)=max(MP{1,i}(:,2));
end
Y_MAX=max(YMAX);

 Fileadress=strcat(pathname,'result\');
  t2=waitbar(0,'正在生成图片');
  for i=1:length(filename)
      h=figure;
     
        set(h,'visible','off');
  plot(MP{1,i}(:,6),MP{1,i}(:,2),'linewidth',2);
      
      grid on;
  hold on;
  switch val1
      case 1
   T_VOR_CRUVE_X=[min(MP{1,i}(:,6))*1.1,max(MP{1,i}(:,6))*1.05];
   T_VOR_CRUVE_Y1=[ T_VOR, T_VOR];
    T_VOR_CRUVE_Y2=[ -T_VOR, -T_VOR];
      case 2
          T_VOR_CRUVE_X=[min(MP{1,i}(:,6))*1.1,max(MP{1,i}(:,6))*1.05];
   T_VOR_CRUVE_Y1=[ T_HINTEN, T_HINTEN];
    T_VOR_CRUVE_Y2=[ -T_HINTEN, -T_HINTEN]; 
            end
   plot(T_VOR_CRUVE_X,T_VOR_CRUVE_Y1,'linewidth',2,'Color','r') ;  
       plot(T_VOR_CRUVE_X,T_VOR_CRUVE_Y2,'linewidth',2,'Color','r') ;  
           
          ylim([-Y_MAX*1.1 Y_MAX*1.1]);
          Y_TICK_MAX=ceil(Y_MAX);
          if mod(Y_TICK_MAX,2)==0
              Y_TICK_MAX=Y_TICK_MAX;
          else
              Y_TICK_MAX=Y_TICK_MAX+1;
          end
xL=xlim ;
yL=ylim ;
set(gca,'ytick',-Y_TICK_MAX:2:Y_TICK_MAX);
       xt=get(gca,'xtick') ;yt=get(gca,'ytick') ;
set(gca,'XTick',[],'XColor','w') ;
set(gca,'YTick',[],'YColor','w') ;
xlabel('Drehwinkel[\circ]','FontSize',zihao,'color','k');ylabel('Verschiebemomente[Nm]','FontSize',zihao,'color','k');  
pos = get(gca,'Position') ;
box off;
x_shift = abs( yL(1)/(yL(2)-yL(1)) ) ;
y_shift = abs( xL(1)/(xL(2)-xL(1)) ) ;
temp_1 = axes( 'Position', pos + [ 0 , pos(4) * x_shift , 0 , - pos(4)* x_shift ] ) ;
xlim(xL) ;
box off ;
set(temp_1,'XTick',xt,'Color','None','YTick',[]) ;
set(temp_1,'YColor','w') ;
temp_2 = axes( 'Position', pos + [ pos(3) * y_shift , 0 , - pos(3)* y_shift , 0 ] ) ;

ylim(yL) ;
box off ;
set(temp_2,'YTick',yt,'Color','None','XTick',[]) ;
set(temp_2,'XColor','w') ;

   box off;
set(gcf,'color','white')
 sfilename1=[Fileadress,num2str(i),'-Verschiebemomente.jpg'];
 frame=getframe(h);
 im=frame2im(frame);
 imwrite(im,sfilename1,'jpg')
  
% saveas(h,sfilename1);
            close(h);
            waitbar(i/length(filename));
  end
  close(t2);
  
  t3=waitbar(0,'正在生成报告');
   filespec_user=[Fileadress,'Verschiebemomente.doc'];
   %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%生成Word报告%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
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
He=180*0.94488188976377952755905511811024;
Wi=240;
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
  waitbar(0.3);
  InlineShapes=Document.InlineShapes;
 headline='III. Einzelergebnis 具体结果';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落
  for i=1:length(filename)
    sfilename1=[Fileadress,num2str(i),'-Verschiebemomente.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
if mod(i,2)==0
   Selection.Start = Selection.end;
Selection.TypeParagraph; 
 Selection.Start = Selection.end;
Selection.TypeParagraph; 
end
waitbar(0.8)
delete(sfilename1); 
  end
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='玻璃升降寿命试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
waitbar(1);
close(t3);
winopen(filespec_user);


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)

function popupmenu1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end






function edit5_Callback(hObject, eventdata, handles)

function edit5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit6_Callback(hObject, eventdata, handles)

function edit6_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit7_Callback(hObject, eventdata, handles)

function edit7_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu3.
function popupmenu3_Callback(hObject, eventdata, handles)

function popupmenu3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
