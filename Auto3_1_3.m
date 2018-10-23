function varargout = Auto3_1_3(varargin)
% AUTO3_1_3 MATLAB code for Auto3_1_3.fig
%      AUTO3_1_3, by itself, creates a new AUTO3_1_3 or raises the existing
%      singleton*.
%
%      H = AUTO3_1_3 returns the handle to a new AUTO3_1_3 or the handle to
%      the existing singleton*.
%
%      AUTO3_1_3('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in AUTO3_1_3.M with the given input arguments.
%
%      AUTO3_1_3('Property','Value',...) creates a new AUTO3_1_3 or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before Auto3_1_3_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to Auto3_1_3_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help Auto3_1_3

% Last Modified by GUIDE v2.5 09-Feb-2018 09:24:22

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto3_1_3_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto3_1_3_OutputFcn, ...
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


% --- Executes just before Auto3_1_3 is made visible.
function Auto3_1_3_OpeningFcn(hObject, eventdata, handles, varargin)
set(handles.uitable2,'Visible','off');
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto3_1_3 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto3_1_3_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
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
[n p]=size(MP21);
if p==9
    set(handles.uitable1,'Visible','on');
    set(handles.uitable2,'Visible','off');
set(handles.uitable1,'data',MP21);
elseif p==7
    set(handles.uitable1,'Visible','off');
    set(handles.uitable2,'Visible','on');
set(handles.uitable2,'data',MP21);
end
set(handles.text4,'Visible','on');
a1=1;
if a1+a2+a3+a4==4;
    set(handles.pushbutton5,'Enable','on');
end
waitbar(100);
close(t1);


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
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
[n p]=size(MP22);
if p==9
    set(handles.uitable1,'Visible','on');
    set(handles.uitable2,'Visible','off');
set(handles.uitable1,'data',MP22);
elseif p==7
    set(handles.uitable1,'Visible','off');
    set(handles.uitable2,'Visible','on');
set(handles.uitable2,'data',MP22);
end
set(handles.text5,'Visible','on');
a2=1;
if a1+a2+a3+a4==4;
    set(handles.pushbutton5,'Enable','on');
end
waitbar(100);
close(t1);


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
set(handles.uitable1,'Visible','on');
    set(handles.uitable2,'Visible','off');
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
    set(handles.pushbutton5,'Enable','on');
end
waitbar(100);
close(t1);

function pushbutton4_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
set(handles.uitable1,'Visible','on');
    set(handles.uitable2,'Visible','off');
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
    set(handles.pushbutton5,'Enable','on');
end
waitbar(100);
close(t1);


function edit1_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)

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


% --- Executes on button press in pushbutton5.
function pushbutton5_Callback(hObject, eventdata, handles)
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
 a=[1,3,5,7];
b=[2,4,6,8];

[n p]=size(MP21);
if p==9
MPr21=zeros(4,18);

for i=1:4
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
MPr22=zeros(4,18);
for i=1:4
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

elseif p==7
    
    MPr21=zeros(4,14);

for i=1:4
    MPr21(i,1)=MP21(b(i),1);MPr21(i,2)=MP21(a(i),1);
    MPr21(i,3)=MP21(b(i),6);MPr21(i,4)=MP21(a(i),6);
    MPr21(i,5)=MP21(b(i),2);MPr21(i,6)=MP21(a(i),2);
    MPr21(i,7)=MP21(b(i),3);MPr21(i,8)=MP21(a(i),3);
    MPr21(i,9)=MP21(b(i),4);MPr21(i,10)=MP21(a(i),4);
    MPr21(i,11)=MP21(b(i),7);MPr21(i,12)=MP21(a(i),7);
    MPr21(i,13)=MP21(b(i),5);MPr21(i,14)=MP21(a(i),5);
  
end
MPr22=zeros(4,18);
for i=1:4
    MPr22(i,1)=MP22(b(i),1);MPr22(i,2)=MP22(a(i),1);
    MPr22(i,3)=MP22(b(i),6);MPr22(i,4)=MP22(a(i),6);
    MPr22(i,5)=MP22(b(i),2);MPr22(i,6)=MP22(a(i),2);
    MPr22(i,7)=MP22(b(i),3);MPr22(i,8)=MP22(a(i),3);
    MPr22(i,9)=MP22(b(i),4);MPr22(i,10)=MP22(a(i),4);
    MPr22(i,11)=MP22(b(i),7);MPr22(i,12)=MP22(a(i),7);
    MPr22(i,13)=MP22(b(i),5);MPr22(i,14)=MP22(a(i),5);

end
  
    
end
MPr71=zeros(4,18);

for i=1:4
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

MPr72=zeros(4,18);

for i=1:4
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
t2=waitbar(0.1);



if val1==1
    if p==9
 for i=1:18
    for j=1:4
        MPra21{j,i}=num2str(MPr21(j,i),'%.3f');
    end
end
    h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr21(4,2),MPr21(4,4),MPr21(4,6),MPr21(4,8),MPr21(4,10),MPr21(4,12),MPr21(4,14),MPr21(4,16),MPr21(4,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr21(4,1),MPr21(4,3),MPr21(4,5),MPr21(4,7),MPr21(4,9),MPr21(4,11),MPr21(4,13),MPr21(4,15),MPr21(4,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','SouthEast');

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

headline='Messung 1:Kraftangriff MP21/测量1：施力点MP21';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落   
    
 Tab1 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc*2.24,lc*1.96,lc*1.84,lc*1.96,lc*1.84,lc*1.96,lc*1.84];   
    
  for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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
 Kraft=[0,20,40,60,80];
for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end
 for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra21{i-4,j-1};
    end
end  
   Selection.Start = Content.end;
Selection.TypeParagraph;

Tab2 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(2); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra21{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

Tab3 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(3); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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
DTI.Cell(1,2).Range.Text = 'MP6';
DTI.Cell(1,3).Range.Text = 'MP22';
DTI.Cell(1,4).Range.Text = 'MP7';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra21{i-4,j+11};
    end
end
t2=waitbar(0.2);
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

Tab4 = Document.Tables.Add(Selection.Range,6,4);
DTI = Document.Tables.Item(4); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5];
    for i = 1:4
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:6
    for j=1:4
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s21=MPr21(4,4)-MPr21(4,3);
s22=MPr21(4,16)-MPr21(4,15);
jist=asin((abs(s21)+abs(s22))/puffer1/1000)/3.1415926*180;
CTist=80*puffer1/jist;
jzul=80*puffer1/100;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer1*1000;
SPist=0.5*(abs(MPr21(4,3))+abs(MPr21(4,15)));

DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(2,1).Range.Text = 'CT-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='T-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'jist';
DTI.Cell(3,1).Select;
Selection.Find.Text='ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'jzul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SP-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(6,1).Range.Text = 'SP-ist';
DTI.Cell(6,1).Select;
Selection.Find.Text='P-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(2,2).Range.Text =num2str(CTist,'%.1f');
if CTist<100
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(jist,'%.2f');
if jist>jzul
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(4,2).Range.Text = num2str(jzul,'%.2f');
DTI.Cell(5,2).Range.Text =num2str(SPzul,'%.2f');
DTI.Cell(6,2).Range.Text =num2str(SPist,'%.2f');
if SPist>SPzul
    DTI.Cell(6,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(6,2).Range.Font.Bold=1;
end

DTI.Cell(2,3).Range.Text = '100';
DTI.Cell(3,3).Range.Text = '<jzul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(6,3).Range.Text = '<SP-zul.';
DTI.Cell(6,3).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,4).Range.Text = '[Nm/°]';
DTI.Cell(3,4).Range.Text = '[°]';
DTI.Cell(4,4).Range.Text = '[°]';
DTI.Cell(5,4).Range.Text = '[mm]';
DTI.Cell(6,4).Range.Text = '[mm]';

t2=waitbar(0.3);


%% MP22施力点输出
for i=1:18
    for j=1:4
        MPra22{j,i}=num2str(MPr22(j,i),'%.3f');
    end
end
    h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr22(4,2),MPr22(4,4),MPr22(4,6),MPr22(4,8),MPr22(4,10),MPr22(4,12),MPr22(4,14),MPr22(4,16),MPr22(4,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr22(4,1),MPr22(4,3),MPr22(4,5),MPr22(4,7),MPr22(4,9),MPr22(4,11),MPr22(4,13),MPr22(4,15),MPr22(4,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','SouthWest');

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


Tab5 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(5); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
    
  for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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
 Kraft=[0,20,40,60,80];
for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end
 for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra22{i-4,j-1};
    end
end  
   Selection.Start = Content.end;
Selection.TypeParagraph;

Tab6 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(6); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra22{i-4,j+5};
    end
end
t2=waitbar(0.4);
Selection.Start = Content.end;
Selection.TypeParagraph;

Tab7 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(7); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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
DTI.Cell(1,2).Range.Text = 'MP6';
DTI.Cell(1,3).Range.Text = 'MP22';
DTI.Cell(1,4).Range.Text = 'MP7';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
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

Tab8 = Document.Tables.Add(Selection.Range,6,4);
DTI = Document.Tables.Item(8); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5];
    for i = 1:4
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:6
    for j=1:4
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s21=MPr22(4,4)-MPr22(4,3);
s22=MPr22(4,16)-MPr22(4,15);
jist=asin((abs(s21)+abs(s22))/puffer1/1000)/3.1415926*180;
CTist=80*puffer1/jist;
jzul=80*puffer1/100;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer1*1000;
SPist=0.5*(abs(MPr22(4,3))+abs(MPr22(4,15)));

DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(2,1).Range.Text = 'CT-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='T-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'jist';
DTI.Cell(3,1).Select;
Selection.Find.Text='ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'jzul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SP-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(6,1).Range.Text = 'SP-ist';
DTI.Cell(6,1).Select;
Selection.Find.Text='P-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(2,2).Range.Text =num2str(CTist,'%.1f');
if CTist<100
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(jist,'%.2f');
if jist>jzul
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(4,2).Range.Text = num2str(jzul,'%.2f');
DTI.Cell(5,2).Range.Text =num2str(SPzul,'%.2f');
DTI.Cell(6,2).Range.Text =num2str(SPist,'%.2f');
if SPist>SPzul
    DTI.Cell(6,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(6,2).Range.Font.Bold=1;
end

DTI.Cell(2,3).Range.Text = '100';
DTI.Cell(3,3).Range.Text = '<jzul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(6,3).Range.Text = '<SP-zul.';
DTI.Cell(6,3).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,4).Range.Text = '[Nm/°]';
DTI.Cell(3,4).Range.Text = '[°]';
DTI.Cell(4,4).Range.Text = '[°]';
DTI.Cell(5,4).Range.Text = '[mm]';
DTI.Cell(6,4).Range.Text = '[mm]';
t2=waitbar(0.5);

%% MP71点输出


for i=1:18
    for j=1:4
        MPra71{j,i}=num2str(MPr71(j,i),'%.3f');
    end
end

 h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr71(4,2),MPr71(4,4),MPr71(4,6),MPr71(4,8),MPr71(4,10),MPr71(4,12),MPr71(4,14),MPr71(4,16),MPr71(4,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr71(4,1),MPr71(4,3),MPr71(4,5),MPr71(4,7),MPr71(4,9),MPr71(4,11),MPr71(4,13),MPr71(4,15),MPr71(4,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','SouthEast');

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

Tab9 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(9); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra71{i-4,j-1};
    end
end
Selection.Start = Content.end;
Selection.TypeParagraph;

Tab10 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(10); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra71{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

t2=waitbar(0.6);
Tab11 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(11); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
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
Tab12 = Document.Tables.Add(Selection.Range,6,4);
DTI = Document.Tables.Item(12); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5];
    for i = 1:4
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:6
    for j=1:4
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s71=MPr71(4,2)-MPr71(4,1);
s72=MPr71(4,18)-MPr71(4,17);
jist=asin((abs(s71)+abs(s72))/puffer2/1000)/3.1415926*180;
CTist=80*puffer2/jist;
jzul=80*puffer2/180;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer2*1000;
SPist=0.5*(abs(MPr71(4,1))+abs(MPr71(4,17)));

DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(2,1).Range.Text = 'CT-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='T-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'jist';
DTI.Cell(3,1).Select;
Selection.Find.Text='ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'jzul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SP-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(6,1).Range.Text = 'SP-ist';
DTI.Cell(6,1).Select;
Selection.Find.Text='P-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(2,2).Range.Text =num2str(CTist,'%.1f');
if CTist<180
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(jist,'%.2f');
if jist>jzul
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(4,2).Range.Text = num2str(jzul,'%.2f');
DTI.Cell(5,2).Range.Text =num2str(SPzul,'%.2f');
DTI.Cell(6,2).Range.Text =num2str(SPist,'%.2f');
if SPist>SPzul
    DTI.Cell(6,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(6,2).Range.Font.Bold=1;
end

DTI.Cell(2,3).Range.Text = '180';
DTI.Cell(3,3).Range.Text = '<jzul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(6,3).Range.Text = '<SP-zul.';
DTI.Cell(6,3).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,4).Range.Text = '[Nm/°]';
DTI.Cell(3,4).Range.Text = '[°]';
DTI.Cell(4,4).Range.Text = '[°]';
DTI.Cell(5,4).Range.Text = '[mm]';
DTI.Cell(6,4).Range.Text = '[mm]';
t2=waitbar(0.7);

%% MP72点输出


for i=1:18
    for j=1:4
        MPra72{j,i}=num2str(MPr72(j,i),'%.3f');
    end
end

 h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr72(4,2),MPr72(4,4),MPr72(4,6),MPr72(4,8),MPr72(4,10),MPr72(4,12),MPr72(4,14),MPr72(4,16),MPr72(4,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr72(4,1),MPr72(4,3),MPr72(4,5),MPr72(4,7),MPr72(4,9),MPr72(4,11),MPr72(4,13),MPr72(4,15),MPr72(4,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','SouthWest');

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

Tab13 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(13); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra72{i-4,j-1};
    end
end
Selection.Start = Content.end;
Selection.TypeParagraph;

Tab14 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(14); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra72{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;


Tab15 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(15); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
t2=waitbar(0.8);
for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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
DTI.Cell(1,2).Range.Text = 'MP6';
DTI.Cell(1,3).Range.Text = 'MP7';
DTI.Cell(1,4).Range.Text = 'MP72';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
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

Tab16 = Document.Tables.Add(Selection.Range,6,4);
DTI = Document.Tables.Item(16); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5];
    for i = 1:4
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:6
    for j=1:4
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s71=MPr72(4,2)-MPr72(4,1);
s72=MPr72(4,18)-MPr72(4,17);
jist=asin((abs(s71)+abs(s72))/puffer2/1000)/3.1415926*180;
CTist=80*puffer2/jist;
jzul=80*puffer2/180;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer2*1000;
SPist=0.5*(abs(MPr72(4,1))+abs(MPr72(4,17)));
t2=waitbar(0.9);
DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(2,1).Range.Text = 'CT-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='T-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'jist';
DTI.Cell(3,1).Select;
Selection.Find.Text='ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'jzul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SP-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(6,1).Range.Text = 'SP-ist';
DTI.Cell(6,1).Select;
Selection.Find.Text='P-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(2,2).Range.Text =num2str(CTist,'%.1f');
if CTist<180
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(jist,'%.2f');
if jist>jzul
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(4,2).Range.Text = num2str(jzul,'%.2f');
DTI.Cell(5,2).Range.Text =num2str(SPzul,'%.2f');
DTI.Cell(6,2).Range.Text =num2str(SPist,'%.2f');
if SPist>SPzul
    DTI.Cell(6,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(6,2).Range.Font.Bold=1;
end

DTI.Cell(2,3).Range.Text = '180';
DTI.Cell(3,3).Range.Text = '<jzul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(6,3).Range.Text = '<SP-zul.';
DTI.Cell(6,3).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,4).Range.Text = '[Nm/°]';
DTI.Cell(3,4).Range.Text = '[°]';
DTI.Cell(4,4).Range.Text = '[°]';
DTI.Cell(5,4).Range.Text = '[mm]';
DTI.Cell(6,4).Range.Text = '[mm]';

%% 不含MP2 MP6点数据
elseif p==7
for i=1:14
    for j=1:4
        MPra21{j,i}=num2str(MPr21(j,i),'%.3f');
    end
end
    h=figure;
set(h,'visible','off');
x=1:7;  y=[MPr21(4,2),MPr21(4,4),MPr21(4,6),MPr21(4,8),MPr21(4,10),MPr21(4,12),MPr21(4,14)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:7;   y2=[MPr21(4,1),MPr21(4,3),MPr21(4,5),MPr21(4,7),MPr21(4,9),MPr21(4,11),MPr21(4,13)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','SouthEast');

ylabel('Verformung(mm)','FontSize',12);
set(gca,'xtick',[1:7]);
set(gca,'xticklabel',{'MP1';'MP21';'MP3';'MP4';'MP5';'MP22';'MP7'});
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
    
 Tab1 = Document.Tables.Add(Selection.Range,8,9);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc*2.24,lc*1.96,lc*1.84,lc*1.96,lc*1.84,lc*1.96,lc*1.84,lc*1.96,lc*1.84];   
    
  for i = 1:9
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP1';
DTI.Cell(1,3).Range.Text = 'MP21';
DTI.Cell(1,4).Range.Text = 'MP3';
DTI.Cell(1,5).Range.Text = 'MP4';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';DTI.Cell(3,8).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';DTI.Cell(3,9).Range.Text = 'Gesamt';

for i=2:9
DTI.Cell(4,i).Range.Text='0.000';
end  
 Kraft=[0,20,40,60,80];
for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end
 for i=5:8
    for j=2:9
       DTI.Cell(i,j).Range.Text=MPra21{i-4,j-1};
    end
end  
   Selection.Start = Content.end;
Selection.TypeParagraph;

Tab2 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(2); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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
DTI.Cell(1,2).Range.Text = 'MP5';
DTI.Cell(1,3).Range.Text = 'MP22';
DTI.Cell(1,4).Range.Text = 'MP7';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra21{i-4,j+7};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;


t2=waitbar(0.2);

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

Tab3 = Document.Tables.Add(Selection.Range,6,4);
DTI = Document.Tables.Item(3); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5];
    for i = 1:4
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:6
    for j=1:4
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s21=MPr21(4,4)-MPr21(4,3);
s22=MPr21(4,12)-MPr21(4,11);
jist=asin((abs(s21)+abs(s22))/puffer1/1000)/3.1415926*180;
CTist=80*puffer1/jist;
jzul=80*puffer1/100;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer1*1000;
SPist=0.5*(abs(MPr21(4,3))+abs(MPr21(4,11)));

DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(2,1).Range.Text = 'CT-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='T-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'jist';
DTI.Cell(3,1).Select;
Selection.Find.Text='ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'jzul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SP-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(6,1).Range.Text = 'SP-ist';
DTI.Cell(6,1).Select;
Selection.Find.Text='P-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(2,2).Range.Text =num2str(CTist,'%.1f');
if CTist<100
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(jist,'%.2f');
if jist>jzul
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(4,2).Range.Text = num2str(jzul,'%.2f');
DTI.Cell(5,2).Range.Text =num2str(SPzul,'%.2f');
DTI.Cell(6,2).Range.Text =num2str(SPist,'%.2f');
if SPist>SPzul
    DTI.Cell(6,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(6,2).Range.Font.Bold=1;
end

DTI.Cell(2,3).Range.Text = '100';
DTI.Cell(3,3).Range.Text = '<jzul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(6,3).Range.Text = '<SP-zul.';
DTI.Cell(6,3).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,4).Range.Text = '[Nm/°]';
DTI.Cell(3,4).Range.Text = '[°]';
DTI.Cell(4,4).Range.Text = '[°]';
DTI.Cell(5,4).Range.Text = '[mm]';
DTI.Cell(6,4).Range.Text = '[mm]';

t2=waitbar(0.3);


%% MP22施力点输出
for i=1:14
    for j=1:4
        MPra22{j,i}=num2str(MPr22(j,i),'%.3f');
    end
end
    h=figure;
set(h,'visible','off');
x=1:7;  y=[MPr22(4,2),MPr22(4,4),MPr22(4,6),MPr22(4,8),MPr22(4,10),MPr22(4,12),MPr22(4,14)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:7;   y2=[MPr22(4,1),MPr22(4,3),MPr22(4,5),MPr22(4,7),MPr22(4,9),MPr22(4,11),MPr22(4,13)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','SouthWest');

ylabel('Verformung(mm)','FontSize',12);
set(gca,'xtick',[1:7]);
set(gca,'xticklabel',{'MP1';'MP21';'MP3';'MP4';'MP5';'MP22';'MP7'});
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


Tab4 = Document.Tables.Add(Selection.Range,8,9);
DTI = Document.Tables.Item(4); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
    
  for i = 1:9
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
    for j=1:9
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end  
  for i=2:5;
DTI.Cell(1,i).Merge(DTI.Cell(1,i+1)); % 第一行第1个到第二行第一个合并
end
DTI.Cell(2,2).Merge(DTI.Cell(2,9));
DTI.Cell(2,1).Merge(DTI.Cell(3,1));

DTI.Cell(1,1).Range.Text = 'Messpkt.';
DTI.Cell(2,1).Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,1).Range.Text = 'Kraft(N)';
DTI.Cell(1,2).Range.Text = 'MP1';
DTI.Cell(1,3).Range.Text = 'MP21';
DTI.Cell(1,4).Range.Text = 'MP3';
DTI.Cell(1,5).Range.Text = 'MP4';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';DTI.Cell(3,8).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';DTI.Cell(3,9).Range.Text = 'Gesamt';


for i=2:9
DTI.Cell(4,i).Range.Text='0.000';
end  
 Kraft=[0,20,40,60,80];
for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end
 for i=5:8
    for j=2:9
       DTI.Cell(i,j).Range.Text=MPra22{i-4,j-1};
    end
end  
   Selection.Start = Content.end;
Selection.TypeParagraph;

Tab5 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(5); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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
DTI.Cell(1,2).Range.Text = 'MP5';
DTI.Cell(1,3).Range.Text = 'MP22';
DTI.Cell(1,4).Range.Text = 'MP7';
DTI.Cell(2,2).Range.Text = 'Verformung(mm)';
DTI.Cell(3,2).Range.Text = 'bleibende';DTI.Cell(3,4).Range.Text = 'bleibende';DTI.Cell(3,6).Range.Text = 'bleibende';
DTI.Cell(3,3).Range.Text = 'Gesamt';DTI.Cell(3,5).Range.Text = 'Gesamt';DTI.Cell(3,7).Range.Text = 'Gesamt';

for i=2:7
DTI.Cell(4,i).Range.Text='0.000';
end

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra22{i-4,j+7};
    end
end
t2=waitbar(0.4);
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

Tab6 = Document.Tables.Add(Selection.Range,6,4);
DTI = Document.Tables.Item(6); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5];
    for i = 1:4
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:6
    for j=1:4
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s21=MPr22(4,4)-MPr22(4,3);
s22=MPr22(4,12)-MPr22(4,11);
jist=asin((abs(s21)+abs(s22))/puffer1/1000)/3.1415926*180;
CTist=80*puffer1/jist;
jzul=80*puffer1/100;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer1*1000;
SPist=0.5*(abs(MPr22(4,3))+abs(MPr22(4,11)));

DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(2,1).Range.Text = 'CT-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='T-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'jist';
DTI.Cell(3,1).Select;
Selection.Find.Text='ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'jzul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SP-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(6,1).Range.Text = 'SP-ist';
DTI.Cell(6,1).Select;
Selection.Find.Text='P-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(2,2).Range.Text =num2str(CTist,'%.1f');
if CTist<100
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(jist,'%.2f');
if jist>jzul
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(4,2).Range.Text = num2str(jzul,'%.2f');
DTI.Cell(5,2).Range.Text =num2str(SPzul,'%.2f');
DTI.Cell(6,2).Range.Text =num2str(SPist,'%.2f');
if SPist>SPzul
    DTI.Cell(6,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(6,2).Range.Font.Bold=1;
end

DTI.Cell(2,3).Range.Text = '100';
DTI.Cell(3,3).Range.Text = '<jzul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(6,3).Range.Text = '<SP-zul.';
DTI.Cell(6,3).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,4).Range.Text = '[Nm/°]';
DTI.Cell(3,4).Range.Text = '[°]';
DTI.Cell(4,4).Range.Text = '[°]';
DTI.Cell(5,4).Range.Text = '[mm]';
DTI.Cell(6,4).Range.Text = '[mm]';
t2=waitbar(0.5);

%% MP71点输出


for i=1:18
    for j=1:4
        MPra71{j,i}=num2str(MPr71(j,i),'%.3f');
    end
end

 h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr71(4,2),MPr71(4,4),MPr71(4,6),MPr71(4,8),MPr71(4,10),MPr71(4,12),MPr71(4,14),MPr71(4,16),MPr71(4,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr71(4,1),MPr71(4,3),MPr71(4,5),MPr71(4,7),MPr71(4,9),MPr71(4,11),MPr71(4,13),MPr71(4,15),MPr71(4,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','SouthEast');

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

Tab7 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(7); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra71{i-4,j-1};
    end
end
Selection.Start = Content.end;
Selection.TypeParagraph;

Tab8 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(8); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra71{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;

t2=waitbar(0.6);
Tab9 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(9); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
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


Tab10 = Document.Tables.Add(Selection.Range,6,4);
DTI = Document.Tables.Item(10); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5];
    for i = 1:4
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:6
    for j=1:4
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s71=MPr71(4,2)-MPr71(4,1);
s72=MPr71(4,18)-MPr71(4,17);
jist=asin((abs(s71)+abs(s72))/puffer2/1000)/3.1415926*180;
CTist=80*puffer2/jist;
jzul=80*puffer2/180;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer2*1000;
SPist=0.5*(abs(MPr71(4,1))+abs(MPr71(4,17)));

DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(2,1).Range.Text = 'CT-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='T-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'jist';
DTI.Cell(3,1).Select;
Selection.Find.Text='ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'jzul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SP-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(6,1).Range.Text = 'SP-ist';
DTI.Cell(6,1).Select;
Selection.Find.Text='P-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(2,2).Range.Text =num2str(CTist,'%.1f');
if CTist<180
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(jist,'%.2f');
if jist>jzul
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(4,2).Range.Text = num2str(jzul,'%.2f');
DTI.Cell(5,2).Range.Text =num2str(SPzul,'%.2f');
DTI.Cell(6,2).Range.Text =num2str(SPist,'%.2f');
if SPist>SPzul
    DTI.Cell(6,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(6,2).Range.Font.Bold=1;
end

DTI.Cell(2,3).Range.Text = '180';
DTI.Cell(3,3).Range.Text = '<jzul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(6,3).Range.Text = '<SP-zul.';
DTI.Cell(6,3).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,4).Range.Text = '[Nm/°]';
DTI.Cell(3,4).Range.Text = '[°]';
DTI.Cell(4,4).Range.Text = '[°]';
DTI.Cell(5,4).Range.Text = '[mm]';
DTI.Cell(6,4).Range.Text = '[mm]';
t2=waitbar(0.7);

%% MP72点输出


for i=1:18
    for j=1:4
        MPra72{j,i}=num2str(MPr72(j,i),'%.3f');
    end
end

 h=figure;
set(h,'visible','off');
x=1:9;  y=[MPr72(4,2),MPr72(4,4),MPr72(4,6),MPr72(4,8),MPr72(4,10),MPr72(4,12),MPr72(4,14),MPr72(4,16),MPr72(4,18)];  plot(x,y,'-s','linewidth',2)
hold on;
x2=1:9;   y2=[MPr72(4,1),MPr72(4,3),MPr72(4,5),MPr72(4,7),MPr72(4,9),MPr72(4,11),MPr72(4,13),MPr72(4,15),MPr72(4,17)];  plot(x2,y2,'-s','linewidth',2);
grid on;
    legend('Verformung unter Last','bleibende Verformung','Location','SouthWest');

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

Tab11 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(11); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra72{i-4,j-1};
    end
end
Selection.Start = Content.end;
Selection.TypeParagraph;

Tab12 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(12); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条


for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
    for j=2:7
       DTI.Cell(i,j).Range.Text=MPra72{i-4,j+5};
    end
end

Selection.Start = Content.end;
Selection.TypeParagraph;


Tab13 = Document.Tables.Add(Selection.Range,8,7);
DTI = Document.Tables.Item(13); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
t2=waitbar(0.8);
for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:8
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

for i=4:8
    DTI.Cell(i,1).Range.Text =num2str(Kraft(i-3));
end

for i=5:8
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

Tab14 = Document.Tables.Add(Selection.Range,6,4);
DTI = Document.Tables.Item(14); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
column_width4 = [lc*2.24,lc*1.9,lc*1.9,lc*2.5];
    for i = 1:4
DTI.Columns.Item(i).Width = column_width4(i);
end
for i=1:6
    for j=1:4
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
         DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
s71=MPr72(4,2)-MPr72(4,1);
s72=MPr72(4,18)-MPr72(4,17);
jist=asin((abs(s71)+abs(s72))/puffer2/1000)/3.1415926*180;
CTist=80*puffer2/jist;
jzul=80*puffer2/180;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer2*1000;
SPist=0.5*(abs(MPr72(4,1))+abs(MPr72(4,17)));
t2=waitbar(0.9);
DTI.Cell(1,1).Range.Text = 'Ergebniss';
DTI.Cell(1,2).Range.Text = 'Ist-Wert';
DTI.Cell(1,3).Range.Text = 'Soll-Wert';
DTI.Cell(1,4).Range.Text = 'Masseinheit';
DTI.Cell(2,1).Range.Text = 'CT-ist';
DTI.Cell(2,1).Select;
Selection.Find.Text='T-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(3,1).Range.Text = 'jist';
DTI.Cell(3,1).Select;
Selection.Find.Text='ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,1).Range.Text = 'jzul.';
DTI.Cell(4,1).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(5,1).Range.Text = 'SP-zul.';
DTI.Cell(5,1).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(6,1).Range.Text = 'SP-ist';
DTI.Cell(6,1).Select;
Selection.Find.Text='P-ist';
Selection.Find.Execute;
Selection.Font.Subscript= true;

DTI.Cell(2,2).Range.Text =num2str(CTist,'%.1f');
if CTist<180
     DTI.Cell(2,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(2,2).Range.Font.Bold=1;
end
DTI.Cell(3,2).Range.Text =num2str(jist,'%.2f');
if jist>jzul
    DTI.Cell(3,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(3,2).Range.Font.Bold=1;
end
DTI.Cell(4,2).Range.Text = num2str(jzul,'%.2f');
DTI.Cell(5,2).Range.Text =num2str(SPzul,'%.2f');
DTI.Cell(6,2).Range.Text =num2str(SPist,'%.2f');
if SPist>SPzul
    DTI.Cell(6,2).Range.Font.Colorindex='wdRed';
    DTI.Cell(6,2).Range.Font.Bold=1;
end

DTI.Cell(2,3).Range.Text = '180';
DTI.Cell(3,3).Range.Text = '<jzul.';
DTI.Cell(3,3).Select;
Selection.Find.Text='zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(4,3).Range.Text = '---';
DTI.Cell(5,3).Range.Text = '---';
DTI.Cell(6,3).Range.Text = '<SP-zul.';
DTI.Cell(6,3).Select;
Selection.Find.Text='P-zul.';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,4).Range.Text = '[Nm/°]';
DTI.Cell(3,4).Range.Text = '[°]';
DTI.Cell(4,4).Range.Text = '[°]';
DTI.Cell(5,4).Range.Text = '[mm]';
DTI.Cell(6,4).Range.Text = '[mm]';

    else
        msgbox('MP21或MP22数据有误');
    end
    
  set(handles.text4,'Visible','off'); set(handles.text5,'Visible','off'); set(handles.text6,'Visible','off');set(handles.text7,'Visible','off');
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
t2=waitbar(1);
close(t2);
winopen([pathname,'report.doc']);
a1=0;a2=0;a3=0;a4=0;
set(handles.pushbutton5,'Enable','off');


elseif val1==2
        SHEET_NAME={'Sheet1' 'Kraftangriff MP21' 'Kraftangriff MP21 Kurven' 'Kraftangriff MP22'...
    'Kraftangriff MP22 Kurven' 'Kraftangriff MP71' 'Kraftangriff MP71 Kurven'...
    'Kraftangriff MP72' 'Kraftangriff MP72 Kurven'};

if p==9
   file_usr=strcat(cd,'\model\Auto3_1_3_7P.xlsx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'report.xlsx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
t2=waitbar(0.2);

for i=1:9
xlswrite([filespec_user],{date},SHEET_NAME{i},'N1');
end 
t2=waitbar(0.3);
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
 xlswrite([filespec_user],a22,'Kraftangriff MP22','C7');
  xlswrite([filespec_user],b22,'Kraftangriff MP22','C19');
  xlswrite([filespec_user],a71,'Kraftangriff MP71','C7');
xlswrite([filespec_user],b71,'Kraftangriff MP71','C19');
 xlswrite([filespec_user],a72,'Kraftangriff MP72','C7');
xlswrite([filespec_user],b72,'Kraftangriff MP72','C19');
t2=waitbar(0.4);

 
 
  s21=MPr21(4,4)-MPr21(4,3);
s22=MPr21(4,16)-MPr21(4,15);
jist=asin((abs(s21)+abs(s22))/puffer1/1000)/3.1415926*180;
CTist=80*puffer1/jist;
jzul=80*puffer1/100;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer1*1000;
SPist=0.5*(abs(MPr21(4,3))+abs(MPr21(4,15)));  
 xlswrite([filespec_user],CTist,'Kraftangriff MP21','C28');
xlswrite([filespec_user],jist,'Kraftangriff MP21','C29');
xlswrite([filespec_user],jzul,'Kraftangriff MP21','C30');
xlswrite([filespec_user],SPzul,'Kraftangriff MP21','C31');
xlswrite([filespec_user],SPist,'Kraftangriff MP21','C32');
t2=waitbar(0.5);

s21=MPr22(4,4)-MPr22(4,3);
s22=MPr22(4,16)-MPr22(4,15);
jist=asin((abs(s21)+abs(s22))/puffer1/1000)/3.1415926*180;
CTist=80*puffer1/jist;
jzul=80*puffer1/100;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer1*1000;
SPist=0.5*(abs(MPr22(4,3))+abs(MPr22(4,15)));    
  xlswrite([filespec_user],CTist,'Kraftangriff MP22','C28');
xlswrite([filespec_user],jist,'Kraftangriff MP22','C29');
xlswrite([filespec_user],jzul,'Kraftangriff MP22','C30');
xlswrite([filespec_user],SPzul,'Kraftangriff MP22','C31');
xlswrite([filespec_user],SPist,'Kraftangriff MP22','C32');
t2=waitbar(0.6);


s71=MPr71(4,2)-MPr71(4,1);
s72=MPr71(4,18)-MPr71(4,17);
jist=asin((abs(s71)+abs(s72))/puffer2/1000)/3.1415926*180;
CTist=80*puffer2/jist;
jzul=80*puffer2/180;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer2*1000;
SPist=0.5*(abs(MPr71(4,1))+abs(MPr71(4,17)));
 xlswrite([filespec_user],CTist,'Kraftangriff MP71','C28');
xlswrite([filespec_user],jist,'Kraftangriff MP71','C29');
xlswrite([filespec_user],jzul,'Kraftangriff MP71','C30');
xlswrite([filespec_user],SPzul,'Kraftangriff MP71','C31');
xlswrite([filespec_user],SPist,'Kraftangriff MP71','C32');
t2=waitbar(0.7);

s71=MPr72(4,2)-MPr72(4,1);
s72=MPr72(4,18)-MPr72(4,17);
jist=asin((abs(s71)+abs(s72))/puffer2/1000)/3.1415926*180;
CTist=80*puffer2/jist;
jzul=80*puffer2/180;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer2*1000;
SPist=0.5*(abs(MPr72(4,1))+abs(MPr72(4,17)));
 xlswrite([filespec_user],CTist,'Kraftangriff MP72','C28');
xlswrite([filespec_user],jist,'Kraftangriff MP72','C29');
xlswrite([filespec_user],jzul,'Kraftangriff MP72','C30');
xlswrite([filespec_user],SPzul,'Kraftangriff MP72','C31');
xlswrite([filespec_user],SPist,'Kraftangriff MP72','C32');
t2=waitbar(0.8);


 
 %%%%%%%%%%%%缺少MP2 MP6点EXCEL输出%%%%%%%%%%%%%%%%%%%%%
 elseif p==7
     file_usr=strcat(cd,'\model\Auto3_1_3_5P.xlsx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=strcat(pathname,'report.xlsx');
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
t2=waitbar(0.2);

for i=1:9
xlswrite([filespec_user],{date},SHEET_NAME{i},'N1');
end 
t2=waitbar(0.3);
%数据整理分裂
 a21=MPr21(:,1:8);
 b21=MPr21(:,9:14);
a22=MPr22(:,1:8);
 b22=MPr22(:,9:14);
  a71=MPr71(:,1:8);
 b71=MPr71(:,9:18);
a72=MPr72(:,1:8);
 b72=MPr72(:,9:18);
xlswrite([filespec_user],a21,'Kraftangriff MP21','C7');
xlswrite([filespec_user],b21,'Kraftangriff MP21','C19');
 xlswrite([filespec_user],a22,'Kraftangriff MP22','C7');
  xlswrite([filespec_user],b22,'Kraftangriff MP22','C19');
  xlswrite([filespec_user],a71,'Kraftangriff MP71','C7');
xlswrite([filespec_user],b71,'Kraftangriff MP71','C19');
 xlswrite([filespec_user],a72,'Kraftangriff MP72','C7');
xlswrite([filespec_user],b72,'Kraftangriff MP72','C19');
t2=waitbar(0.4);

 
 
  s21=MPr21(4,4)-MPr21(4,3);
s22=MPr21(4,12)-MPr21(4,11);
jist=asin((abs(s21)+abs(s22))/puffer1/1000)/3.1415926*180;
CTist=80*puffer1/jist;
jzul=80*puffer1/100;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer1*1000;
SPist=0.5*(abs(MPr21(4,3))+abs(MPr21(4,11)));
 xlswrite([filespec_user],CTist,'Kraftangriff MP21','C28');
xlswrite([filespec_user],jist,'Kraftangriff MP21','C29');
xlswrite([filespec_user],jzul,'Kraftangriff MP21','C30');
xlswrite([filespec_user],SPzul,'Kraftangriff MP21','C31');
xlswrite([filespec_user],SPist,'Kraftangriff MP21','C32');
t2=waitbar(0.5);

s21=MPr22(4,4)-MPr22(4,3);
s22=MPr22(4,12)-MPr22(4,11);
jist=asin((abs(s21)+abs(s22))/puffer1/1000)/3.1415926*180;
CTist=80*puffer1/jist;
jzul=80*puffer1/100;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer1*1000;
SPist=0.5*(abs(MPr22(4,3))+abs(MPr22(4,11)));
  xlswrite([filespec_user],CTist,'Kraftangriff MP22','C28');
xlswrite([filespec_user],jist,'Kraftangriff MP22','C29');
xlswrite([filespec_user],jzul,'Kraftangriff MP22','C30');
xlswrite([filespec_user],SPzul,'Kraftangriff MP22','C31');
xlswrite([filespec_user],SPist,'Kraftangriff MP22','C32');
t2=waitbar(0.6);


s71=MPr71(4,2)-MPr71(4,1);
s72=MPr71(4,18)-MPr71(4,17);
jist=asin((abs(s71)+abs(s72))/puffer2/1000)/3.1415926*180;
CTist=80*puffer2/jist;
jzul=80*puffer2/180;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer2*1000;
SPist=0.5*(abs(MPr71(4,1))+abs(MPr71(4,17)));
 xlswrite([filespec_user],CTist,'Kraftangriff MP71','C28');
xlswrite([filespec_user],jist,'Kraftangriff MP71','C29');
xlswrite([filespec_user],jzul,'Kraftangriff MP71','C30');
xlswrite([filespec_user],SPzul,'Kraftangriff MP71','C31');
xlswrite([filespec_user],SPist,'Kraftangriff MP71','C32');
t2=waitbar(0.7);

s71=MPr72(4,2)-MPr72(4,1);
s72=MPr72(4,18)-MPr72(4,17);
jist=asin((abs(s71)+abs(s72))/puffer2/1000)/3.1415926*180;
CTist=80*puffer2/jist;
jzul=80*puffer2/180;
SPzul=7.5/92.5*sin(jzul*3.1415/180)*0.5*puffer2*1000;
SPist=0.5*(abs(MPr72(4,1))+abs(MPr72(4,17)));
 xlswrite([filespec_user],CTist,'Kraftangriff MP72','C28');
xlswrite([filespec_user],jist,'Kraftangriff MP72','C29');
xlswrite([filespec_user],jzul,'Kraftangriff MP72','C30');
xlswrite([filespec_user],SPzul,'Kraftangriff MP72','C31');
xlswrite([filespec_user],SPist,'Kraftangriff MP72','C32');
t2=waitbar(0.8);

     
     
end  

a1=0;a2=0;a3=0;a4=0;
set(handles.pushbutton5,'Enable','off');
 set(handles.text4,'Visible','off'); set(handles.text5,'Visible','off'); set(handles.text6,'Visible','off');set(handles.text7,'Visible','off');
t2=waitbar(1);
close(t2);
 winopen(filespec_user);



end
try
close(figure(1));close(figure(2));close(figure(3));close(figure(4));
end
