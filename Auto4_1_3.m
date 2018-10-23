function varargout = Auto4_1_3(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto4_1_3_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto4_1_3_OutputFcn, ...
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
function Auto4_1_3_OpeningFcn(hObject, eventdata, handles, varargin)
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

function varargout = Auto4_1_3_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;

function checkbox1_Callback(hObject, eventdata, handles)
 if get(handles.checkbox1,'Value')==1
set(handles.pushbutton3,'Enable','on');
set(handles.pushbutton4,'Enable','on');
 else
     set(handles.pushbutton3,'Enable','off');
set(handles.pushbutton4,'Enable','off');
 end

 % --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
global PATH_VOR
PATH_VOR=uigetdir;
if PATH_VOR==0
        msgbox('请选择文件夹');
    return;
else
    set(handles.edit1,'String',PATH_VOR);
end
% --- Executes on button press in pushbutton4.
function pushbutton4_Callback(hObject, eventdata, handles)
global PATH_NACH
PATH_NACH=uigetdir;
if PATH_NACH==0
        msgbox('请选择文件夹');
    return;
else
    set(handles.edit2,'String',PATH_NACH);
end
 
 
 
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
clear global TEIL_NAME;
global TEIL_NAME  newpath 


oldpath=cd;
if isempty(newpath)|| ~exist('newpath')
     newpath=cd;
end
 
[filename,pathname,fileindex]=uigetfile('*.txt','选择零件号索引txt',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif filename~=0
    newpath=pathname; 
   i=1;fid = fopen(fullfile(pathname,filename));
tline = fgetl(fid);
while ischar(tline)
TEIL_NAME{i}=tline;
tline = fgetl(fid);i=i+1;
end
fclose(fid);
     set(handles.pushbutton2,'Enable','on'); 
     msgbox('零件索引导入成功，请导入试验数据');
end

% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global TEIL_NAME newpath;
PATH_NACH=get(handles.edit2,'String') ;
PATH_VOR=get(handles.edit1,'String') ;
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif length(TEIL_NAME)~=length(filename)/9
    msgbox('零件数量和数据数量不符，请检查零件索引TXT文件');
    return;
else

    
    
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
end 
  RESOLUTION_HE=500;
  RESOLUTION_WI=1300;
  zihao=20;
  n=1;%曲线索引
   if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
  Fileadress=strcat(pathname,'result\');
%%%%%%%%%%%%%%%%%%%求拉脱力力最大值并画图%%%%%%%%%%%%%%%%%%%%%
t2=waitbar(0,'正在生成报告图片');
 COLOR_INDEX=[204 0 0;204 189 0;58 204 0;0 204 180;0 49 204;97 0 204;204 0 151;...
      148 71 56;141 148 56;55 149 66;56 148 139;92 72 132;125 73 131;130 74 74;226 130 226;...
      99 23 99;57 231 78;22 188 42]/255;
for i=1:length(filename)
    KRAFT_MAX(i)=max(MP{1,i}(:,2));
   end
for i=1:(length(filename)/9)
   h(i)=figure;
    set(h(i),'visible','off');
     plot(MP{1,n}(:,1),MP{1,n}(:,2),'linewidth',2,'color',COLOR_INDEX(1,:));
     hold on;
     plot(MP{1,n+1}(:,1),MP{1,n+1}(:,2),'linewidth',2,'color',COLOR_INDEX(2,:));
     plot(MP{1,n+2}(:,1),MP{1,n+2}(:,2),'linewidth',2,'color',COLOR_INDEX(3,:));
     plot(MP{1,n+3}(:,1),MP{1,n+3}(:,2),'linewidth',2,'color',COLOR_INDEX(4,:));
    plot(MP{1,n+4}(:,1),MP{1,n+4}(:,2),'linewidth',2,'color',COLOR_INDEX(5,:));
    plot(MP{1,n+5}(:,1),MP{1,n+5}(:,2),'linewidth',2,'color',COLOR_INDEX(6,:));
    plot(MP{1,n+6}(:,1),MP{1,n+6}(:,2),'linewidth',2,'color',COLOR_INDEX(7,:));
    plot(MP{1,n+7}(:,1),MP{1,n+7}(:,2),'linewidth',2,'color',COLOR_INDEX(8,:));
    plot(MP{1,n+8}(:,1),MP{1,n+8}(:,2),'linewidth',2,'color',COLOR_INDEX(9,:));
          set(h(i),'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]);
   set(h(i),'color','w')
        set(gca,'FontSize',zihao);
        title(TEIL_NAME{i},'FontSize',zihao);
     xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);  

     Ym=max([max(MP{1,n}(:,2)) max(MP{1,n+1}(:,2)) max(MP{1,n+2}(:,2)) max(MP{1,n+3}(:,2)) max(MP{1,n+4}(:,2)) ...
       max(MP{1,n+5}(:,2)) max(MP{1,n+6}(:,2)) max(MP{1,n+7}(:,2)) max(MP{1,n+8}(:,2))])*1.1;
    Xm=max([max(MP{1,n}(:,1)) max(MP{1,n+1}(:,1)) max(MP{1,n+2}(:,1)) max(MP{1,n+3}(:,1)) max(MP{1,n+4}(:,1)) ...
       max(MP{1,n+5}(:,1)) max(MP{1,n+6}(:,1)) max(MP{1,n+7}(:,1)) max(MP{1,n+8}(:,1))])*1.3;
    STAND_X=[0;Xm/1.3];
   STAND_Y=[500;500];
   plot(STAND_X,STAND_Y,'linewidth',3,'Color','r');
      legend('RT Teil 1#','RT Teil 2#','RT Teil 3#','KWT Teil 1#',...
       'KWT Teil 2#','KWT Teil 3#','WL Teil 1#','WL Teil 2#','WL Teil 3#','Location','SouthEast');
   
   grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 Ym]);
   hold off; 
   sfilename1=[Fileadress,num2str(i),'.jpg'];
     f=getframe(h(i));
           imwrite(f.cdata,sfilename1);
           close(h(i));
     n=n+9; 
    waitbar(i/(length(filename)/9));
end
close(t2);
    t3=waitbar(0,'正在生成Word报告');
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%生成Word报告%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
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
 t3=waitbar(0.1);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;

%===文档的页边距===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;

headline='III. Einzelergebnis 具体结果';
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落

He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
biaotihao=10;


Tab1 = Document.Tables.Add(Selection.Range, length(filename)+1,5);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

lc=28.381133333333333333333333333333; %厘米换算
column_width = [3.5*lc,3*lc,2*lc,4*lc,2.74*lc];


for i = 1:5
DTI.Columns.Item(i).Width = column_width(i);
end
 DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
 DTI.Range.Font.NameAscii='Arial';
 DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';


t3=waitbar(0.2);
for i=1:5
DTI.Cell(1,i).Range.Font.Bold=1;
end
DTI.Cell(1,1).Range.Text = 'Teilenummer';
DTI.Cell(1,2).Range.Text = 'Temperatur';
DTI.Cell(1,3).Range.Text = 'Teil';
DTI.Cell(1,4).Range.Text = 'Messwert[N]';
DTI.Cell(1,5).Range.Text = 'Soll-Wert[N]';
DTI.Cell(2,5).Merge(DTI.Cell(length(filename)+1,5))
DTI.Cell(2,5).Range.Text = '>500';


m=2;
for i=1:(length(filename)/9)
DTI.Cell(m,1).Merge(DTI.Cell(m+8,1)); % 每个零件号第一列第一行至第九行合并
m=m+9;
end
%写入Teil1 2 3
m=2;
for i=1:(length(filename)/9)
DTI.Cell(m,2).Merge(DTI.Cell(m+2,2)); %RT合并
DTI.Cell(m+3,2).Merge(DTI.Cell(m+5,2)); %KWT合并
DTI.Cell(m+6,2).Merge(DTI.Cell(m+8,2)); %WL合并
DTI.Cell(m,2).Range.Text = 'RT';DTI.Cell(m+3,2).Range.Text = 'Nach KWT';DTI.Cell(m+6,2).Range.Text = 'Nach Waemelagerung';
DTI.Cell(m,3).Range.Text = '1';DTI.Cell(m+3,3).Range.Text = '1';DTI.Cell(m+6,3).Range.Text = '1';
DTI.Cell(m+1,3).Range.Text = '2';DTI.Cell(m+4,3).Range.Text = '2';DTI.Cell(m+7,3).Range.Text = '2';
DTI.Cell(m+2,3).Range.Text = '3';DTI.Cell(m+5,3).Range.Text = '3';DTI.Cell(m+8,3).Range.Text = '3';
DTI.Cell(m,1).Range.Text = TEIL_NAME{i};%写入零件号
m=m+9;
end
t3=waitbar(0.3);

%及测量值
for i=1:length(filename)
       DTI.Cell(i+1,4).Range.Text = num2str(KRAFT_MAX(i),'%.2f');
       if KRAFT_MAX(i)<500
             DTI.Cell(i+1,4).Range.Font.Colorindex='wdRed';
             DTI.Cell(i+1,4).Range.Font.Bold=1;
       end
end
Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
InlineShapes=Document.InlineShapes;
t3=waitbar(0.6);
for i=1:length(filename)/9
    sfilename1=[Fileadress,num2str(i),'.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
delete(sfilename1); 

end
t3=waitbar(0.9);

%%%%%%%%%%%%%%%%%%%%%%%%选择生成照片%%%%%%%%%%%%%%
 if get(handles.checkbox1,'Value')==1
close(t3);
t4=waitbar(0,'正在粘贴图片');
Selection.Start = Selection.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
IMAGES_VOR=dir([PATH_VOR,'\*.jpg']);
IMAGES_NACH=dir([PATH_NACH,'\*.jpg']);


Tab2 = Document.Tables.Add(Selection.Range, length(filename)*2/3,2);
DTI = Document.Tables.Item(2); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条

lc=28.381133333333333333333333333333; %厘米换算
column_width = [8.93*lc,8.93*lc];

for i = 1:2
DTI.Columns.Item(i).Width = column_width(i);
end
 DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
 DTI.Range.Font.NameAscii='Arial';
 DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
 n=1;
for i=1:length(filename)/3
   DTI.Cell(n,1).Select;
         handle=Selection.InlineShapes.AddPicture([PATH_VOR,'\',IMAGES_VOR(i).name]);
    Selection.MoveRight;
    handle=Selection.InlineShapes.AddPicture([PATH_NACH,'\',IMAGES_NACH(i).name]);
    n=n+2;
    waitbar(i/(length(filename)/3));
end

n=2;
for i=1:length(filename)/9
     DTI.Cell(n,1).Range.Text=[TEIL_NAME{i},' RT vor Pruefung'];
 DTI.Cell(n+2,1).Range.Text=[TEIL_NAME{i},' KWT vor Pruefung'];
  DTI.Cell(n+4,1).Range.Text=[TEIL_NAME{i},' WL vor Pruefung'];
  DTI.Cell(n,2).Range.Text=[TEIL_NAME{i},' RT nach Pruefung'];
 DTI.Cell(n+2,2).Range.Text=[TEIL_NAME{i},' KWT nach Pruefung'];
  DTI.Cell(n+4,2).Range.Text=[TEIL_NAME{i},' WL nach Pruefung'];
  n=n+6;
end


%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='IZAF底护板撕裂力试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
close(t4);
set(handles.pushbutton2,'Enable','off'); 
winopen([Fileadress,'report.doc']);

 else
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='IZAF底护板孔拉脱力试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
t3=waitbar(1);
close(t3);
set(handles.pushbutton2,'Enable','off'); 
winopen([Fileadress,'report.doc']);
 end




function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double


% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit2_Callback(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit2 as text
%        str2double(get(hObject,'String')) returns contents of edit2 as a double


% --- Executes during object creation, after setting all properties.
function edit2_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit2 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
