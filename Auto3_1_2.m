function varargout = Auto3_1_2(varargin)


% Last Modified by GUIDE v2.5 09-Feb-2018 09:22:04

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto3_1_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto3_1_2_OutputFcn, ...
                   'gui_LayoutFcn',  [], ...
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


% --- Executes just before Auto3_1_2 is made visible.
function Auto3_1_2_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   unrecognized PropertyName/PropertyValue pairs from the
%            command line (see VARARGIN)

% Choose default command line output for Auto3_1_2
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto3_1_2 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto3_1_2_OutputFcn(hObject, eventdata, handles)
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


function edit2_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end





function edit1_Callback(hObject, eventdata, handles)

% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)


% --- Executes during object creation, after setting all properties.
function popupmenu1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end

% --- Executes on button press in pushbutton7.
function pushbutton3_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global MP21 pathname newpath a1 a2 a3 a4;
a1=0;
oldpath=cd;
if isempty(newpath)|| ~exist('newpath')
     newpath=cd;
 end

[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif filename~=0
    t1=waitbar(0,'正在导入数据');
     newpath=pathname; 
    Filename=strcat(pathname,filename);
    MP21=xlsread(Filename);
end
waitbar(50);
set(handles.uitable1,'data',MP21);
set(handles.text4,'Visible','on');
a1=1;
if a1+a2+a3+a4==4;
    set(handles.pushbutton4,'Enable','on');
end
waitbar(100);
close(t1);
% --- Executes on button press in pushbutton7.
function pushbutton5_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);

global MP22 pathname newpath a1 a2 a3 a4;
a2=0;
oldpath=cd;
if isempty(newpath)|| ~exist('newpath')
     newpath=cd;
 end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif filename~=0
    t1=waitbar(0,'正在导入数据');
     newpath=pathname; 
    Filename=strcat(pathname,filename);
    MP22=xlsread(Filename);
end
waitbar(50);
set(handles.uitable1,'data',MP22);
set(handles.text5,'Visible','on');
a2=1;
if a1+a2+a3+a4==4;
    set(handles.pushbutton4,'Enable','on');
end
waitbar(100);
close(t1);



% --- Executes on button press in pushbutton7.
function pushbutton6_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global MP71 pathname newpath a1 a2 a3 a4;
a3=0;
oldpath=cd;
if isempty(newpath)|| ~exist('newpath')
     newpath=cd;
 end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif filename~=0
    t1=waitbar(0,'正在导入数据');
     newpath=pathname; 
    Filename=strcat(pathname,filename);
    MP71=xlsread(Filename);
end
waitbar(50);
set(handles.uitable1,'data',MP71);
set(handles.text6,'Visible','on');
a3=1;
if a1+a2+a3+a4==4;
    set(handles.pushbutton4,'Enable','on');
end
waitbar(100);
close(t1);

% --- Executes on button press in pushbutton7.
function pushbutton7_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);

global MP72 pathname newpath a1 a2 a3 a4;
a4=0;
oldpath=cd;
if isempty(newpath)|| ~exist('newpath')
     newpath=cd;
 end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif filename~=0
    t1=waitbar(0,'正在导入数据');
     newpath=pathname; 
    Filename=strcat(pathname,filename);
    MP72=xlsread(Filename);
end
waitbar(50);
set(handles.uitable1,'data',MP72);
set(handles.text7,'Visible','on');
a4=1;
if a1+a2+a3+a4==4;
    set(handles.pushbutton4,'Enable','on');
end
waitbar(100);
close(t1);

% --- Executes on button press in pushbutton6.
function pushbutton4_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global MP21 MP22 MP71 MP72 pathname a1 a2 a3 a4;

biaotihao=10;
val1=get(handles.popupmenu1,'Value');
puffer1=str2num(get(handles.edit1,'String'));
puffer2=str2num(get(handles.edit2,'String'));
if isempty(puffer1)||isempty(puffer2);
    msgbox('请输入Puffer间距');
    return
else
puffer1=puffer1/1000;
puffer2=puffer2/1000;
end
t2=waitbar(0,'正在生成报告');
 a=[1,3,5,7,9,11];
b=[2,4,6,8,10,12];
MPr21=zeros(6,18);

for i=1:6
    MPr21(i,1)=MP21(b(i),1);MPr21(i,2)=MP21(a(i),1);
    MPr21(i,3)=MP21(b(i),8);MPr21(i,4)=MP21(a(i),8);
    MPr21(i,5)=MP21(b(i),2);MPr21(i,6)=MP21(a(i),2);
    MPr21(i,7)=MP21(b(i),3);MPr21(i,8)=MP21(a(i),3);
    MPr21(i,9)=MP21(b(i),4);MPr21(i,10)=MP21(a(i),4);
    MPr21(i,11)=MP21(b(i),5);MPr21(i,12)=MP21(a(i),5);
    MPr21(i,13)=MP21(b(i),6);MPr21(i,14)=MP21(a(i),6);
    MPr21(i,15)=MP21(b(i),9);MPr21(i,16)=MP21(a(i),9);
     MPr21(i,17)=MP21(b(i),7);MPr21(i,18)=MP21(a(i),7);
end
MPr22=zeros(6,18);
for i=1:6
    MPr22(i,1)=MP22(b(i),1);MPr22(i,2)=MP22(a(i),1);
    MPr22(i,3)=MP22(b(i),8);MPr22(i,4)=MP22(a(i),8);
    MPr22(i,5)=MP22(b(i),2);MPr22(i,6)=MP22(a(i),2);
    MPr22(i,7)=MP22(b(i),3);MPr22(i,8)=MP22(a(i),3);
    MPr22(i,9)=MP22(b(i),4);MPr22(i,10)=MP22(a(i),4);
    MPr22(i,11)=MP22(b(i),5);MPr22(i,12)=MP22(a(i),5);
    MPr22(i,13)=MP22(b(i),6);MPr22(i,14)=MP22(a(i),6);
    MPr22(i,15)=MP22(b(i),9);MPr22(i,16)=MP22(a(i),9);
     MPr22(i,17)=MP22(b(i),7);MPr22(i,18)=MP22(a(i),7);
end
MPr71=zeros(6,18);

for i=1:6
    MPr71(i,1)=MP71(b(i),8);MPr71(i,2)=MP71(a(i),8);
    MPr71(i,3)=MP71(b(i),1);MPr71(i,4)=MP71(a(i),1);
    MPr71(i,5)=MP71(b(i),2);MPr71(i,6)=MP71(a(i),2);
    MPr71(i,7)=MP71(b(i),3);MPr71(i,8)=MP71(a(i),3);
    MPr71(i,9)=MP71(b(i),4);MPr71(i,10)=MP71(a(i),4);
    MPr71(i,11)=MP71(b(i),5);MPr71(i,12)=MP71(a(i),5);
    MPr71(i,13)=MP71(b(i),6);MPr71(i,14)=MP71(a(i),6);
    MPr71(i,15)=MP71(b(i),7);MPr71(i,16)=MP71(a(i),7);
    MPr71(i,17)=MP71(b(i),9);MPr71(i,18)=MP71(a(i),9);
end

MPr72=zeros(6,18);

for i=1:6
    MPr72(i,1)=MP72(b(i),8);MPr72(i,2)=MP72(a(i),8);
    MPr72(i,3)=MP72(b(i),1);MPr72(i,4)=MP72(a(i),1);
    MPr72(i,5)=MP72(b(i),2);MPr72(i,6)=MP72(a(i),2);
    MPr72(i,7)=MP72(b(i),3);MPr72(i,8)=MP72(a(i),3);
    MPr72(i,9)=MP72(b(i),4);MPr72(i,10)=MP72(a(i),4);
    MPr72(i,11)=MP72(b(i),5);MPr72(i,12)=MP72(a(i),5);
    MPr72(i,13)=MP72(b(i),6);MPr72(i,14)=MP72(a(i),6);
    MPr72(i,15)=MP72(b(i),7);MPr72(i,16)=MP72(a(i),7);
    MPr72(i,17)=MP72(b(i),9);MPr72(i,18)=MP72(a(i),9);
end


t2=waitbar(0.2);
if val1==1
    for i=1:18
    for j=1:6
        MPra21{j,i}=num2str(MPr21(j,i),'%.3f');
    end
end
    h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr21(6,2),MPr21(6,4),MPr21(6,6),MPr21(6,8),MPr21(6,10),MPr21(6,12),MPr21(6,14),MPr21(6,16),MPr21(6,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr21(6,1),MPr21(6,3),MPr21(6,5),MPr21(6,7),MPr21(6,9),MPr21(6,11),MPr21(6,13),MPr21(6,15),MPr21(6,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','NorthEast');

ylabel('Verformung(mm)','FontSize',12);
set(gca,'xtick',[1:9]);
set(gca,'xticklabel',{'MP1';'MP21';'MP2';'MP3';'MP4';'MP5';'MP6';'MP22';'MP7'});
box off;
set(gcf,'color','w');
set(gca,'FontSize',12);
saveas(h,[pathname,'h.jpg']);
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

headline='Messung 1:Kraftangriff MP21/测量1：施力点MP21';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落


Tab1 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc*2.24,lc*1.96,lc*1.84,lc*1.96,lc*1.84,lc*1.96,lc*1.84];

for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP1';
DTI.Cell(1,3).Range.Text = 'MP21';
DTI.Cell(1,4).Range.Text = 'MP2';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

Kraft=[0,20,40,60,80,100,120];
for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra21{i-4,j-1};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab2 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(2); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP3';
DTI.Cell(1,3).Range.Text = 'MP4';
DTI.Cell(1,4).Range.Text = 'MP5';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra21{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab3 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(3); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP6';
DTI.Cell(1,3).Range.Text = 'MP22';
DTI.Cell(1,4).Range.Text = 'MP7';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra21{i-4,j+11};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture([pathname,'h.jpg']);
InlineShapes.Item(1).Height=He;
InlineShapes.Item(1).Width=Wi;

delete([pathname,'h.jpg']);


Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
headline='Beurteilungskriterien/评价标准';
Selection.Text=headline;
Selection.Font.Size=biaotihao; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落




Tab4 = Document.Tables.Add(Selection.Range,5,5);
DTI = Document.Tables.Item(4); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5,lc*2.1];

for i = 1:5
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:5
    for j=1:5
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end


s21=MPr21(6,4)-MPr21(6,3);
CVist=120*0.5*puffer1/s21;
SVzul=7.5*120*0.5*puffer1/92.5/25;
s21_t={num2str(s21,'%.2f')};
CVist_t={num2str(CVist,'%.2f')};
SVzul_t={num2str(SVzul,'%.2f')};
DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(1,5).Range.Text = 'Bewertung';
DTI.Cell(2,1).Range.Text = 'CV-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='V-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'SV-ist=S21';
DTI.Cell(3,1).Select;
Selection.Find.Text='V-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'CV-min';
DTI.Cell(4,1).Select;
Selection.Find.Text='V-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SV-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,2).Range.Text =CVist_t{1,1};
if CVist<25
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(MPr21(6,3),'%.2f');
if abs(MPr21(6,3))>(SVzul)
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(5,2).Range.Text =SVzul_t{1,1};
DTI.Cell(4,2).Range.Text = '25';
DTI.Cell(2,3).Range.Text = '>CV-min';
DTI.Cell(2,3).Select;
Selection.Find.Text='V-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,3).Range.Text = '<SV-zul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(2,4).Range.Text = '[Nm/mm]';
DTI.Cell(3,4).Range.Text = '[mm]';
DTI.Cell(4,4).Range.Text = '[Nm/mm]';
DTI.Cell(5,4).Range.Text = '[mm]';


t2=waitbar(0.4);


%% MP22施力点输出




for i=1:18
    for j=1:6
        MPra22{j,i}=num2str(MPr22(j,i),'%.3f');
    end
end

    h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr22(6,2),MPr22(6,4),MPr22(6,6),MPr22(6,8),MPr22(6,10),MPr22(6,12),MPr22(6,14),MPr22(6,16),MPr22(6,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr22(6,1),MPr22(6,3),MPr22(6,5),MPr22(6,7),MPr22(6,9),MPr22(6,11),MPr22(6,13),MPr22(6,15),MPr22(6,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','NorthWest');

ylabel('Verformung(mm)','FontSize',12);
set(gca,'xtick',[1:9]);
set(gca,'xticklabel',{'MP1';'MP21';'MP2';'MP3';'MP4';'MP5';'MP6';'MP22';'MP7'});
box off;
set(gcf,'color','w');
set(gca,'FontSize',12);
saveas(h,[pathname,'h.jpg']);


headline='Messung 1:Kraftangriff MP22/测量1：施力点MP22';
Selection.Start=Content.end; % 起始点为0，即表示每次写入覆盖之前资料
Selection.TypeParagraph;
Selection.Text=headline;
Selection.Font.Size=10; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落


Tab5 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(5); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算

for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP1';
DTI.Cell(1,3).Range.Text = 'MP21';
DTI.Cell(1,4).Range.Text = 'MP2';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra22{i-4,j-1};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab6 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(6); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP3';
DTI.Cell(1,3).Range.Text = 'MP4';
DTI.Cell(1,4).Range.Text = 'MP5';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra22{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab7 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(7); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP6';
DTI.Cell(1,3).Range.Text = 'MP22';
DTI.Cell(1,4).Range.Text = 'MP7';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra22{i-4,j+11};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture([pathname,'h.jpg']);
InlineShapes.Item(1).Height=He;
InlineShapes.Item(1).Width=Wi;

delete([pathname,'h.jpg']);


Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
headline='Beurteilungskriterien/评价标准';
Selection.Text=headline;
Selection.Font.Size=biaotihao; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落




Tab8 = Document.Tables.Add(Selection.Range,5,5);
DTI = Document.Tables.Item(8); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5,lc*2.1];
%row_height = [28.5849,28.5849,28.5849,28.5849,25.4717,25.4717,32.8302,312.1698,17.8302,49.2453,14.1509,18.6792];
for i = 1:5
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:5
    for j=1:5
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s22=MPr22(6,16)-MPr22(6,15);
CVist=120*0.5*puffer1/s22;
SVzul=7.5*120*0.5*puffer1/92.5/25;
DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(1,5).Range.Text = 'Bewertung';
DTI.Cell(2,1).Range.Text = 'CV-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='V-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'SV-ist=S22';
DTI.Cell(3,1).Select;
Selection.Find.Text='V-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'CV-min';
DTI.Cell(4,1).Select;
Selection.Find.Text='V-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SV-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,2).Range.Text = num2str(CVist,'%.2f');
if CVist<25
    DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(MPr22(6,15),'%.2f');
if abs(MPr22(6,15))>(SVzul)
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(5,2).Range.Text =num2str(SVzul,'%.2f');
DTI.Cell(4,2).Range.Text = '25';
DTI.Cell(2,3).Range.Text = '>CV-min';
DTI.Cell(2,3).Select;
Selection.Find.Text='V-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,3).Range.Text = '<SV-zul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(2,4).Range.Text = '[Nm/mm]';
DTI.Cell(3,4).Range.Text = '[mm]';
DTI.Cell(4,4).Range.Text = '[Nm/mm]';
DTI.Cell(5,4).Range.Text = '[mm]';




t2=waitbar(0.6);




%% MP71点输出


for i=1:18
    for j=1:6
        MPra71{j,i}=num2str(MPr71(j,i),'%.3f');
    end
end

    h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr71(6,2),MPr71(6,4),MPr71(6,6),MPr71(6,8),MPr71(6,10),MPr71(6,12),MPr71(6,14),MPr71(6,16),MPr71(6,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr71(6,1),MPr71(6,3),MPr71(6,5),MPr71(6,7),MPr71(6,9),MPr71(6,11),MPr71(6,13),MPr71(6,15),MPr71(6,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','NorthEast');

ylabel('Verformung(mm)','FontSize',12);
set(gca,'xtick',[1:9]);
set(gca,'xticklabel',{'MP71';'MP1';'MP2';'MP3';'MP4';'MP5';'MP6';'MP7';'MP72'});
box off;
set(gcf,'color','w');
set(gca,'FontSize',12);
saveas(h,[pathname,'h.jpg']);


headline='Messung 2:Kraftangriff MP71/测量2：施力点MP71';
Selection.Start=Content.end; % 起始点为0，即表示每次写入覆盖之前资料
Selection.TypeParagraph;
Selection.Text=headline;
Selection.Font.Size=10; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落


Tab9 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(9); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算

for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP71';
DTI.Cell(1,3).Range.Text = 'MP1';
DTI.Cell(1,4).Range.Text = 'MP2';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra71{i-4,j-1};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab10 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(10); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP3';
DTI.Cell(1,3).Range.Text = 'MP4';
DTI.Cell(1,4).Range.Text = 'MP5';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra71{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab11 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(11); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP6';
DTI.Cell(1,3).Range.Text = 'MP7';
DTI.Cell(1,4).Range.Text = 'MP72';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra71{i-4,j+11};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture([pathname,'h.jpg']);
InlineShapes.Item(1).Height=He;
InlineShapes.Item(1).Width=Wi;

delete([pathname,'h.jpg']);


Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
headline='Beurteilungskriterien/评价标准';
Selection.Text=headline;
Selection.Font.Size=biaotihao; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落




Tab12 = Document.Tables.Add(Selection.Range,5,5);
DTI = Document.Tables.Item(12); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算

for i = 1:5
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:5
    for j=1:5
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s71=MPr71(6,2)-MPr71(6,1);
CVist=120*0.5*puffer2/s71;
SVzul=7.5*120*0.5*puffer2/92.5/25;
DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(1,5).Range.Text = 'Bewertung';
DTI.Cell(2,1).Range.Text = 'CV-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='V-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'SV-ist=S71';
DTI.Cell(3,1).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'CV-min';
DTI.Cell(4,1).Select;
Selection.Find.Text='V-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SV-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,2).Range.Text = num2str(CVist,'%.2f');
if CVist<25
    DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(MPr71(6,1),'%.2f');
if abs(MPr71(6,1))>(SVzul)
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(5,2).Range.Text =num2str(SVzul,'%.2f');
DTI.Cell(4,2).Range.Text = '25';
DTI.Cell(2,3).Range.Text = '>CV-min';
DTI.Cell(2,3).Select;
Selection.Find.Text='V-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,3).Range.Text = '<SV-zul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(2,4).Range.Text = '[Nm/mm]';
DTI.Cell(3,4).Range.Text = '[mm]';
DTI.Cell(4,4).Range.Text = '[Nm/mm]';
DTI.Cell(5,4).Range.Text = '[mm]';


t2=waitbar(0.8);

%% MP72点输出


for i=1:18
    for j=1:6
        MPra72{j,i}=num2str(MPr72(j,i),'%.3f');
    end
end

    h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr72(6,2),MPr72(6,4),MPr72(6,6),MPr72(6,8),MPr72(6,10),MPr72(6,12),MPr72(6,14),MPr72(6,16),MPr72(6,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr72(6,1),MPr72(6,3),MPr72(6,5),MPr72(6,7),MPr72(6,9),MPr72(6,11),MPr72(6,13),MPr72(6,15),MPr72(6,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','NorthWest');

ylabel('Verformung(mm)','FontSize',12);
set(gca,'xtick',[1:9]);
set(gca,'xticklabel',{'MP71';'MP1';'MP2';'MP3';'MP4';'MP5';'MP6';'MP7';'MP72'});
box off;
set(gcf,'color','w');
set(gca,'FontSize',12);
saveas(h,[pathname,'h.jpg']);


headline='Messung 2:Kraftangriff MP72/测量2：施力点MP72';
Selection.Start=Content.end; % 起始点为0，即表示每次写入覆盖之前资料
Selection.TypeParagraph;
Selection.Text=headline;
Selection.Font.Size=10; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落


Tab13 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(13); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP71';
DTI.Cell(1,3).Range.Text = 'MP1';
DTI.Cell(1,4).Range.Text = 'MP2';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra72{i-4,j-1};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab14 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(14); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP3';
DTI.Cell(1,3).Range.Text = 'MP4';
DTI.Cell(1,4).Range.Text = 'MP5';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra72{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab15 = Document.Tables.Add(Selection.Range,10,7);
DTI = Document.Tables.Item(15); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:10
    for j=1:7
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
for i=2:4;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,7));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP6';
DTI.Cell(1,3).Range.Text = 'MP7';
DTI.Cell(1,4).Range.Text = 'MP72';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:10
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:10
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra72{i-4,j+11};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture([pathname,'h.jpg']);
InlineShapes.Item(1).Height=He;
InlineShapes.Item(1).Width=Wi;

delete([pathname,'h.jpg']);


Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
headline='Beurteilungskriterien/评价标准';
Selection.Text=headline;
Selection.Font.Size=biaotihao; % 字体大小
Selection.Font.NameAscii='Arial';
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落




Tab16 = Document.Tables.Add(Selection.Range,5,5);
DTI = Document.Tables.Item(16); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算

for i = 1:5
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:5
    for j=1:5
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s72=MPr72(6,18)-MPr72(6,17);
CVist=120*0.5*puffer2/s72;
SVzul=7.5*120*0.5*puffer2/92.5/25;
DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(1,5).Range.Text = 'Bewertung';
DTI.Cell(2,1).Range.Text = 'CV-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='V-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'SV-ist=S72';
DTI.Cell(3,1).Select;
Selection.Find.Text='V-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'CV-min';
DTI.Cell(4,1).Select;
Selection.Find.Text='V-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SV-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,2).Range.Text = num2str(CVist,'%.2f');
if CVist<25
    DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(MPr72(6,17),'%.2f');
if abs(MPr72(6,17))>(SVzul)
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(5,2).Range.Text =num2str(SVzul,'%.2f');
DTI.Cell(4,2).Range.Text = '25';
DTI.Cell(2,3).Range.Text = '>CV-min';
DTI.Cell(2,3).Select;
Selection.Find.Text='V-min';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,3).Range.Text = '<SV-zul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='V-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(2,4).Range.Text = '[Nm/mm]';
DTI.Cell(3,4).Range.Text = '[mm]';
DTI.Cell(4,4).Range.Text = '[Nm/mm]';
DTI.Cell(5,4).Range.Text = '[mm]';

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
t2=waitbar(1);
close(t2);
winopen([pathname,'report.doc']);
a1=0;a2=0;a3=0;a4=0;
set(handles.pushbutton4,'Enable','off');
 set(handles.text4,'Visible','off'); set(handles.text5,'Visible','off'); set(handles.text6,'Visible','off');set(handles.text7,'Visible','off');

elseif val1==2
t2=waitbar(0.3);
    try
     Excel = actxGetRunningServer('Excel.Application');
 catch
     Excel = actxserver('Excel.Application'); 
 end;
  Excel.Visible = 0;
  Workbook = Excel.Workbooks.Open([cd,'\model\Auto3_1_2.xlsx']);
  Excel.Visible = 0;    % set(Excel, 'Visible', 1); 
    filespec_user = [pathname,'report.xlsx'];
    Sheets = Excel.ActiveWorkbook.Sheets;
    for i=1:9
     Sheet1 = Sheets.Item(i); 
 Sheet1.Activate;
 Sheet1.Range('N1').Value={date};
    end
  Workbook.SaveAs(filespec_user);
     Excel.Quit;
     t2=waitbar(0.3);
  %计算评价指标  
s21=MPr21(6,4)-MPr21(6,3);
CVist21=120*0.5*puffer1/s21;
SVzul21=7.5*120*0.5*puffer1/92.5/25;     
     
 s22=MPr22(6,16)-MPr22(6,15);
CVist22=120*0.5*puffer1/s22;
SVzul22=7.5*120*0.5*puffer1/92.5/25;

s71=MPr71(6,2)-MPr71(6,1);
CVist71=120*0.5*puffer2/s71;
SVzul71=7.5*120*0.5*puffer2/92.5/25;

s72=MPr72(6,18)-MPr72(6,17);
CVist72=120*0.5*puffer2/s72;
SVzul72=7.5*120*0.5*puffer2/92.5/25;
t2=waitbar(0.4);
%数据整理分裂
 a21=MPr21(:,1:8);
 b21=MPr21(:,9:18);
a22=MPr22(:,1:8);
 b22=MPr22(:,9:18);
  a71=MPr71(:,1:8);
 b71=MPr71(:,9:18);
a72=MPr72(:,1:8);
 b72=MPr72(:,9:18);
 xlswrite([filespec_user],a21,'Kraftangriff MP21','C7');
xlswrite([filespec_user],b21,'Kraftangriff MP21','C19');
t2=waitbar(0.5);
  xlswrite([filespec_user],a22,'Kraftangriff MP22','C7');
  xlswrite([filespec_user],b22,'Kraftangriff MP22','C19');
  xlswrite([filespec_user],a71,'Kraftangriff MP71','C7');
xlswrite([filespec_user],b71,'Kraftangriff MP71','C19');
t2=waitbar(0.6);
  xlswrite([filespec_user],a72,'Kraftangriff MP72','C7');
xlswrite([filespec_user],b72,'Kraftangriff MP72','C19');

xlswrite([filespec_user],CVist21,'Kraftangriff MP21','C28');
xlswrite([filespec_user],s21,'Kraftangriff MP21','C29');
xlswrite([filespec_user],SVzul21,'Kraftangriff MP21','C30');
t2=waitbar(0.7);
xlswrite([filespec_user],CVist22,'Kraftangriff MP22','C28');
xlswrite([filespec_user],s22,'Kraftangriff MP22','C29');
xlswrite([filespec_user],SVzul22,'Kraftangriff MP22','C30');
xlswrite([filespec_user],CVist71,'Kraftangriff MP71','C28');
xlswrite([filespec_user],s71,'Kraftangriff MP71','C29');
t2=waitbar(0.8);
xlswrite([filespec_user],SVzul71,'Kraftangriff MP71','C30');
xlswrite([filespec_user],CVist72,'Kraftangriff MP72','C28');
xlswrite([filespec_user],s72,'Kraftangriff MP72','C29');
t2=waitbar(0.9);
xlswrite([filespec_user],SVzul72,'Kraftangriff MP72','C30');

t2=waitbar(1);
close(t2);
winopen(filespec_user);
a1=0;a2=0;a3=0;a4=0;
set(handles.pushbutton4,'Enable','off');
 set(handles.text4,'Visible','off'); set(handles.text5,'Visible','off'); set(handles.text6,'Visible','off');set(handles.text7,'Visible','off');
end
try
close(figure(1));close(figure(2));close(figure(3));close(figure(4));
end
