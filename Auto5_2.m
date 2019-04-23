function varargout = Auto5_2(varargin)


% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto5_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto5_2_OutputFcn, ...
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


% --- Executes just before Auto5_2 is made visible.
function Auto5_2_OpeningFcn(hObject, eventdata, handles, varargin)
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

% UIWAIT makes Auto5_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto5_2_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on selection change in listbox2.
function listbox2_Callback(hObject, eventdata, handles)
cla(handles.axes1);
MP=getappdata(0,'Auto5_2_MP');
STAND_TITLE=getappdata(0,'STAND_TITLE');
OUT=getappdata(0,'Auto5_2_OUT');
OUT_DATA=getappdata(0,'Auto5_2_OUT_DATA');
H_index=getappdata(0,'Auto5_2_Hindex');
CHOOSE=get(handles.listbox2,'Value');                %listbox的值
i=CHOOSE;
ZIHAO_TU_YULAN=10;
TITLEFONTSIZE=13;

plot(handles.axes1,MP{i}(:,1),MP{i}(:,2),'linewidth',2);
hold on
plot(handles.axes1,MP{i}(OUT(i,1),1),MP{i}(OUT(i,1),2),'ro','MarkerFaceColor','r','Markersize',3);
plot(handles.axes1,MP{i}(OUT(i,2),1),MP{i}(OUT(i,2),2),'ro','MarkerFaceColor','r','Markersize',3);
plot(handles.axes1,MP{i}(OUT(i,3),1),MP{i}(OUT(i,3),2),'ro','MarkerFaceColor','r','Markersize',3);
plot(handles.axes1,MP{i}(OUT(i,4),1),MP{i}(OUT(i,4),2),'ro','MarkerFaceColor','r','Markersize',3);
plot(handles.axes1,MP{i}(H_index(i,1),1),MP{i}(H_index(i,1),2),'ro','MarkerFaceColor','r','Markersize',3);
plot(handles.axes1,MP{i}(H_index(i,2),1),MP{i}(H_index(i,2),2),'ro','MarkerFaceColor','r','Markersize',3);
plot([MP{i}(H_index(i,1),1),MP{i}(H_index(i,2),1)],[MP{i}(H_index(i,1),2),MP{i}(H_index(i,2),2)],'Color','r');
datacursormode on

xlabel(handles.axes1,'Weg/位移[mm]','FontSize',ZIHAO_TU_YULAN)
ylabel(handles.axes1,'Kraft/力[N]','FontSize',ZIHAO_TU_YULAN)
title(handles.axes1,STAND_TITLE{i},'FontSize',TITLEFONTSIZE)
axis(handles.axes1,[0 max(MP{i}(:,1))*1.05 0 max(MP{i}(:,2))*1.1]);
set(handles.edit4,'String',OUT_DATA(i,1));
if OUT_DATA(i,1)<5.5
   set(handles.edit4,'foregroundcolor','red') 
else
     set(handles.edit4,'foregroundcolor','black') 
end
set(handles.edit5,'String',OUT_DATA(i,2));
set(handles.edit6,'String',OUT_DATA(i,3));
set(handles.edit7,'String',OUT_DATA(i,4));
set(handles.edit8,'String',OUT_DATA(i,5));
if OUT_DATA(i,5)>2
    set(handles.edit8,'foregroundcolor','red')
else
    set(handles.edit8,'foregroundcolor','black')
end
set(handles.edit9,'String',OUT_DATA(i,6));
if OUT_DATA(i,6)<0.9 ||OUT_DATA(i,6)>1.1
    set(handles.edit9,'foregroundcolor','red')
else
    set(handles.edit9,'foregroundcolor','black')
end

% --- Executes during object creation, after setting all properties.
function listbox2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)

function Fahrzeugcode_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)

  [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','off');
  if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
      msgbox('导入文件失败');
      return;
  else
      t1=waitbar(0,'正在读入数据');
      Filename=strcat(pathname,filename);
      [Type Sheet Format]=xlsfinfo(Filename)
      for i=1:length(Sheet)-3
          MK=xlsread(Filename,Sheet{i+3});
          STAND_TITLE{i}=Sheet{i+3};
          MP{i}(:,1)=MK(:,1);
          MP{i}(:,2)=MK(:,2);
          waitbar(i/(length(Sheet)-3));
      end
      try
          system('taskkill/IM excel.exe');
      end
      set(handles.listbox2,'String',Sheet(4:end));
  end
close(t1);

     t2=waitbar(0,'正在求解');
     for i=1:length(MP)
         [p1,X2,X3,X4,X5,WEG_COL,H_Final_index,H_Final]=Auto5_2_Core(MP{i});
         OUT(i,1)=X2;
         OUT(i,2)=X3-5;
         OUT(i,3)=X5;
         OUT(i,4)=X4;
         OUT(i,5)=H_Final;
         OUT(i,6)=p1;
         H_index(i,1)= WEG_COL(H_Final_index,1);
         H_index(i,2)=WEG_COL(H_Final_index,2);
         waitbar(i/(length(MP)));
     end
     close(t2);
     t3=waitbar(0,'正在转换数据');
     for i=1:length(MP)
         OUT_DATA(i,1)=MP{i}(OUT(i,1),2); %F1
         OUT_DATA(i,2)=MP{i}(OUT(i,2),2); %F2
         OUT_DATA(i,3)=MP{i}(OUT(i,3),2); %F3,F1mm
         OUT_DATA(i,4)=MP{i}(OUT(i,4),2); %F4
         OUT_DATA(i,5)=OUT(i,5);%H
         OUT_DATA(i,6)=OUT(i,6);%m
         waitbar(i/(length(MP)));
     end
     close(t3);
setappdata(0,'STAND_TITLE',STAND_TITLE);
setappdata(0,'Auto5_2_MP',MP);
setappdata(0,'Auto5_2_OUT',OUT);
setappdata(0,'Auto5_2_OUT_DATA',OUT_DATA);
setappdata(0,'Auto5_2_Hindex',H_index);
set(handles.listbox2,'Value',1);
set(handles.pushbutton2,'Enable','on');
setappdata(0,'Auto5_2_pathname',pathname);
setappdata(0,'Auto5_2_filename',filename);
msgbox('数据导入成功');
% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
%%%%%%%%%%%%%%%%%%%%%%生成报告%%%%%%%%%%%%%%%%
pathname=getappdata(0,'Auto5_2_pathname');
filename=getappdata(0,'Auto5_2_filename');
Fileadress=strcat(pathname,'result\');
   if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
   end
MP=getappdata(0,'Auto5_2_MP');
STAND_TITLE=getappdata(0,'STAND_TITLE');
OUT=getappdata(0,'Auto5_2_OUT');
OUT_DATA=getappdata(0,'Auto5_2_OUT_DATA');
H_index=getappdata(0,'Auto5_2_Hindex');
ZIHAO_TU_YULAN=20;
TITLEFONTSIZE=30;
t1=waitbar(0,'正在生成图片') ;  
for i=1:length(MP)
    h(i)=figure;
        set(h(i),'position',[100,100,1300,800]); 
    set(h(i),'visible','off');
    plot(MP{i}(:,1),MP{i}(:,2),'linewidth',2);
    grid on;
    if get(handles.checkbox2,'value')==1
        hold on
        plot(MP{i}(OUT(i,1),1),MP{i}(OUT(i,1),2),'ro','MarkerFaceColor','r','Markersize',5);
        plot(MP{i}(OUT(i,2),1),MP{i}(OUT(i,2),2),'ro','MarkerFaceColor','r','Markersize',5);
        plot(MP{i}(OUT(i,3),1),MP{i}(OUT(i,3),2),'ro','MarkerFaceColor','r','Markersize',5);
        plot(MP{i}(OUT(i,4),1),MP{i}(OUT(i,4),2),'ro','MarkerFaceColor','r','Markersize',5);
        plot(MP{i}(H_index(i,1),1),MP{i}(H_index(i,1),2),'ro','MarkerFaceColor','r','Markersize',5);
        plot(MP{i}(H_index(i,2),1),MP{i}(H_index(i,2),2),'ro','MarkerFaceColor','r','Markersize',5);
     plot([MP{i}(H_index(i,1),1),MP{i}(H_index(i,2),1)],[MP{i}(H_index(i,1),2),MP{i}(H_index(i,2),2)],'Color','r','linewidth',2);
    end
    datacursormode on
    set(gca,'FontSize',ZIHAO_TU_YULAN)
    xlabel('Weg/位移[mm]','FontSize',ZIHAO_TU_YULAN)
    ylabel('Kraft/力[N]','FontSize',ZIHAO_TU_YULAN)
    title(STAND_TITLE{i},'FontSize',TITLEFONTSIZE)
    axis([0 max(MP{i}(:,1))*1.05 0 max(MP{i}(:,2))*1.1]);
    sfilename1=[Fileadress,num2str(i),'-',STAND_TITLE{i},'.jpg'];
    saveas(h(i),sfilename1);
    close(h(i));
    waitbar(i/length(MP));
end
close(t1);
 t2=waitbar(0,'正在生成Word报告') ;  
 biaotihao=10;
He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
filespec_user=[Fileadress,'report.doc'];
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
headline='III.1 Handbremshebei,Druckknopfbetaetigun 手制动按钮释放试验';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=biaotihao; % 字体大小
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落         
 Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落  
waitbar(0.1)

%%建立数据表格
Tab1 = Document.Tables.Add(Selection.Range, 3+length(MP), 8);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
% 设置行高，列宽
lc=28.381133333333333333333333333333; %厘米换算
column_width = [2.94*lc,1.09*lc,2.01*lc,1.65*lc,1.5*lc,1.75*lc,1.75*lc,2.25*lc];
for i = 1:8
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1: 3+length(MP)
    for j=1:8
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
        DTI.Cell(i,j).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
    end
end
waitbar(0.2)
DTI.Cell(1,1).Merge(DTI.Cell(3,1));
DTI.Cell(1,2).Merge(DTI.Cell(2,2));
DTI.Cell(1,3).Merge(DTI.Cell(1,7));

DTI.Cell(1,1).Range.Text = 'Teil-Nr 零件号';
DTI.Cell(1,2).Range.Text = 'Nr 序号';
DTI.Cell(1,3).Range.Text = 'Bet?tigungskraft操作力 ';
DTI.Cell(3,2).Range.Text = 'Soll';
DTI.Cell(2,3).Range.Text = 'F1[N]';
DTI.Cell(2,4).Range.Text = 'F2[N]';
DTI.Cell(2,5).Range.Text = 'F3[N]';
DTI.Cell(2,6).Range.Text = 'F4[N]';
DTI.Cell(2,7).Range.Text = 'H* [N]';
DTI.Cell(2,8).Range.Text = 'm[N/mm]';

DTI.Cell(3,3).Range.Text = '5,5 + 2(bei s1 ≤ 0,2mm)';
DTI.Cell(3,7).Range.Text = '≤2.0';
DTI.Cell(3,8).Range.Text = '1±0.1';
waitbar(0.5)
for i=4:3+length(MP)
    DTI.Cell(i,2).Range.Text = STAND_TITLE{i-3};
    DTI.Cell(i,3).Range.Text =num2str(OUT_DATA(i-3,1),'%.2f');
    if OUT_DATA(i-3,1)<5.5
        DTI.Cell(i,3).Range.Font.Colorindex='wdRed';
        DTI.Cell(i,3).Range.Font.Bold=1;
    end
    DTI.Cell(i,4).Range.Text =num2str(OUT_DATA(i-3,2),'%.2f');
    DTI.Cell(i,5).Range.Text =num2str(OUT_DATA(i-3,3),'%.2f');
    DTI.Cell(i,6).Range.Text =num2str(OUT_DATA(i-3,4),'%.2f');
    DTI.Cell(i,7).Range.Text =num2str(OUT_DATA(i-3,5),'%.2f');
    if OUT_DATA(i-3,5)>2
        DTI.Cell(i,7).Range.Font.Colorindex='wdRed';
        DTI.Cell(i,7).Range.Font.Bold=1;
    end
    DTI.Cell(i,8).Range.Text =num2str(OUT_DATA(i-3,6),'%.2f');
    if OUT_DATA(i-3,6)<0.9 || OUT_DATA(i-3,6)>1.1
        DTI.Cell(i,8).Range.Font.Colorindex='wdRed';
        DTI.Cell(i,8).Range.Font.Bold=1;
    end         
end
Selection.Start = Content.end;
Selection.TypeParagraph;
waitbar(0.7)
InlineShapes=Document.InlineShapes;
for i=1:length(MP)
Teil2address{i}=[Fileadress,num2str(i),'-',STAND_TITLE{i},'.jpg'];
end

for i=1:length(MP)
handle=Selection.InlineShapes.AddPicture(Teil2address{1,i});
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落    
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落 
end
waitbar(0.9)
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档

for i=1:length(MP)
    delete(Teil2address{1,i});
end
%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='手制动按钮释放';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
waitbar(1);
close(t2);
winopen([Fileadress,'report.doc']);


function edit4_Callback(hObject, eventdata, handles)

function edit4_CreateFcn(hObject, eventdata, handles)

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



function edit8_Callback(hObject, eventdata, handles)

function edit8_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit9_Callback(hObject, eventdata, handles)

function edit9_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox2.
function checkbox2_Callback(hObject, eventdata, handles)

