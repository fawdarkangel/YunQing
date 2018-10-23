function varargout = Auto3_1_1(varargin)
% AUTO3_1_1 MATLAB code for Auto3_1_1.fig
%      AUTO3_1_1, by itself, creates a new AUTO3_1_1 or raises the existing
%      singleton*.
%
%      H = AUTO3_1_1 returns the handle to a new AUTO3_1_1 or the handle to
%      the existing singleton*.
%
%      AUTO3_1_1('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in AUTO3_1_1.M with the given input arguments.
%
%      AUTO3_1_1('Property','Value',...) creates a new AUTO3_1_1 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Auto3_1_1_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Auto3_1_1_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Auto3_1_1

% Last Modified by GUIDE v2.5 06-Jan-2018 13:45:13

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto3_1_1_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto3_1_1_OutputFcn, ...
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


% --- Executes just before Auto3_1_1 is made visible.
function Auto3_1_1_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to Auto3_1_1 (see VARARGIN)

% Choose default command line output for Auto3_1_1
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto3_1_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto3_1_1_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;
% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: contents = cellstr(get(hObject,'String')) returns popupmenu1 contents as cell array
%        contents{get(hObject,'Value')} returns selected item from popupmenu1


% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to popupmenu1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: popupmenu controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
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
% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global MP pathname;
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据');

if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
  
else
    Filename=strcat(pathname,filename);
    MP=xlsread(Filename);
end
set(handles.uitable2,'data',MP);
set(handles.pushbutton2,'Enable','on');


% --- 生成WORD报告.
function pushbutton2_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global MP pathname;
biaotihao=10;
puffer=str2num(get(handles.edit1,'String'));
L=str2num(get(handles.edit2,'String'));
 list=get(handles.popupmenu1,'String');
 val1=get(handles.popupmenu1,'Value');
if isempty(puffer)||isempty(L);
    msgbox('请输入支撑点间距b(mm)');
else
    puffer=puffer/1000;
    L=L/1000;
end
t1=waitbar(0,'正在生成报告');
a=[1,3,5,7,9,11];
b=[2,4,6,8,10,12];
MPr=zeros(6,16);
for i=1:6
    MPr(i,1)=MP(b(i),1);MPr(i,2)=MP(a(i),1);
    MPr(i,3)=MP(b(i),2);MPr(i,4)=MP(a(i),2);
    MPr(i,5)=MP(b(i),3);MPr(i,6)=MP(a(i),3);
    MPr(i,7)=MP(b(i),4);MPr(i,8)=MP(a(i),4);
    MPr(i,9)=MP(b(i),5);MPr(i,10)=MP(a(i),5);
    MPr(i,11)=MP(b(i),6);MPr(i,12)=MP(a(i),6);
    MPr(i,13)=MP(b(i),7);MPr(i,14)=MP(a(i),7);
     MPr(i,15)=MP(b(i),8);MPr(i,16)=MP(a(i),8);
end
t1=waitbar(1/5);

for i=1:16
for j=1:6
    MPra{j,i}=num2str(MPr(j,i),'%.3f');

end
end

t1=waitbar(2/5);

if val1==1
h=figure;
set(h,'visible','off');
x=1:7;  y=[MPr(6,2),MPr(6,4),MPr(6,6),MPr(6,8),MPr(6,10),MPr(6,12),MPr(6,14)];  plot(x,y,'-s','linewidth',2)

hold on;
x2=1:7;   y2=[MPr(6,1),MPr(6,3),MPr(6,5),MPr(6,7),MPr(6,9),MPr(6,11),MPr(6,13)];  plot(x2,y2,'-s','linewidth',2);
grid on;legend('Verformung unter Last','bleibende Verformung','Location','NorthWest');ylabel('Verformung(mm)','FontSize',12);
set(gca,'xtick',[1:7]);
set(gca,'xticklabel',['MP11';'MP12';'MP13';'MP14';'MP15';'MP16';'MP17']);
box off;
set(gcf,'color','w');
set(gca,'FontSize',12);
saveas(h,[pathname,'h.jpg']);

t1=waitbar(3/5);
%%生成Word报告

He=180*0.94488188976377952755905511811024*1.7683;
Wi=240*1.9681;
filespec_user=[pathname,'report.doc'];
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
Word.Visible =0; % 使word为可见；或set(Word, 'Visible', 1); 
%===打开word文件，如果路径下没有则创建一个空白文档打开========================
if exist(filespec_user,'file'); 
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

Tab1 = Document.Tables.Add(Selection.Range,10,9);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
DTI.Rows.Alignment='wdAlignRowCenter';
lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc*2.24,lc*1.94,lc*1.86,lc*1.94,lc*1.86,lc*1.94,lc*1.86,lc*1.94,lc*1.86];
%row_height = [28.5849,28.5849,28.5849,28.5849,25.4717,25.4717,32.8302,312.1698,17.8302,49.2453,14.1509,18.6792];
for i = 1:9
DTI.Columns.Item(i).Width = column_width(i);
end

for i=1:10
    for j=1:9
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end


for i=2:5
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,9));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';%垂直居中
DTI.Cell(1,1).Range.Text = 'Messpunkt';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP11';
DTI.Cell(1,3).Range.Text = 'MP12';
DTI.Cell(1,4).Range.Text = 'MP13';
DTI.Cell(1,5).Range.Text = 'MP14';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';DTI.Cell(3,8).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';DTI.Cell(3,9).Range.Text = 'Gesamt';
for i=2:9
DTI.Cell(4,i).Range.Text='0.000';
end
Kraft=[0,20,40,60,80,100,120];
for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:9
       DTI.Cell(i,j).Range.Text=MPra{i-4,j-1};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

%%表格2
Tab2 = Document.Tables.Add(Selection.Range,10,9);
DTI = Document.Tables.Item(2); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
DTI.Rows.Alignment='wdAlignRowCenter';
lc=28.381133333333333333333333333333; %厘米换算

for i = 1:9
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:9
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';%水平对其
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';%调整字体
    end
end


for i=2:5;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,9));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(1,1).Range.Text = 'Messpunkt';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP15';
DTI.Cell(1,3).Range.Text = 'MP16';
DTI.Cell(1,4).Range.Text = 'MP17';
DTI.Cell(1,5).Range.Text = 'S4';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';DTI.Cell(3,8).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';DTI.Cell(3,9).Range.Text = 'Gesamt';
for i=2:9
DTI.Cell(4,i).Range.Text='0.000';
end
Kraft=[0,20,40,60,80,100,120];
for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:9
       DTI.Cell(i,j).Range.Text=MPra{i-4,j+7} ;
    end
end
t1=waitbar(4/5);
Selection.Start = Content.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture([pathname,'h.jpg']);
InlineShapes.Item(1).Height=He;
InlineShapes.Item(1).Width=Wi;

delete([pathname,'h.jpg'])

Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
headline='Beurteilungskriterien/评价标准';
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落

%%评价表格
s14=MPr(6,8)-MPr(6,7);
s11=MPr(6,2)-MPr(6,1);
s17=MPr(6,14)-MPr(6,13);
CHist=120*0.5*L/(s14-0.5*(s11+s17)); %目标CHist计算 
SHzul=9*120*0.5*puffer/91/80; %目标位移量计算


CHist_1=round(CHist*100)/100;
m=find('.'==num2str(CHist_1));
n=char(num2str(CHist_1,'%.2f')); 
CHist={n};
SHzul_1=round(SHzul*1000)/1000;
m=find('.'==num2str(SHzul_1));
n=char(num2str(SHzul_1,'%.3f'));
SHzul={n};




Tab3 = Document.Tables.Add(Selection.Range,5,5);
DTI = Document.Tables.Item(3); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc*2.24,lc*1.9,lc*1.9,lc*2.5,lc*2.1];
%row_height = [28.5849,28.5849,28.5849,28.5849,25.4717,25.4717,32.8302,312.1698,17.8302,49.2453,14.1509,18.6792];
for i = 1:5
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:5
    for j=1:5
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(1,5).Range.Text = 'Bewertung';
DTI.Cell(2,1).Range.Text = 'CH-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='H-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'MP14bleb.';
DTI.Cell(3,1).Select;
Selection.Find.Text='bleb.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'SH-zul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='H-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'CH-min';
DTI.Cell(5,1).Select;
Selection.Find.Text='H-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,2).Range.Text = CHist{1,1};
DTI.Cell(3,2).Range.Text = MPra{6,7};
DTI.Cell(4,2).Range.Text = SHzul{1,1};
DTI.Cell(5,2).Range.Text = '80';
DTI.Cell(2,3).Range.Text = '>CH-min';
DTI.Cell(2,3).Select;
Selection.Find.Text='H-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,3).Range.Text = '<SH-zul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='H-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(2,4).Range.Text = '[Nm/mm]';
DTI.Cell(3,4).Range.Text = 'mm';
DTI.Cell(4,4).Range.Text = 'mm';
DTI.Cell(5,4).Range.Text = '[Nm/mm]';
t1=waitbar(5/5);

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
close(t1);
winopen([pathname,'report.doc'])
set(handles.pushbutton2,'Enable','off');

set(handles.uitable2,'data',[]);


%% EXCEL报告
elseif val1==2 
    t1=waitbar(3/5);
    
    file_usr=strcat(cd,'\model\Auto3_1_1.xlsx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'report.xlsx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);

     t1=waitbar(4/5);
 a=MPr(:,1:8);
 b=MPr(:,9:16);
s14=MPr(6,8)-MPr(6,7);
s11=MPr(6,2)-MPr(6,1);
s17=MPr(6,14)-MPr(6,13);
CHist=120*0.5*L/(s14-0.5*(s11+s17));
SHzul=9*120*0.5*puffer/91/80;


for i=1:3
xlswrite([filespec_user],{date},[strcat('Sheet',num2str(i))],'N1');
end

xlswrite([filespec_user],a,'Sheet2','C7');
xlswrite([filespec_user],b,'Sheet2','C19');
xlswrite([filespec_user],CHist,'Sheet2','C28');
xlswrite([filespec_user],MPr(6,7),'Sheet2','C29');
xlswrite([filespec_user],SHzul,'Sheet2','C30');
t1=waitbar(5/5);
close(t1);
winopen(filespec_user)
set(handles.pushbutton2,'Enable','off');
set(handles.uitable2,'data',[]);
end
