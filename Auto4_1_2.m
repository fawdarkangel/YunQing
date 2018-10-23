function varargout = Auto4_1_2(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto4_1_2_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto4_1_2_OutputFcn, ...
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
function Auto4_1_2_OpeningFcn(hObject, eventdata, handles, varargin)
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

function varargout = Auto4_1_2_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;



function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
clear global;
global TEIL_NAME newpath;
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

[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','选择数据','MultiSelect','on',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('导入文件失败');
  return;
elseif length(TEIL_NAME)~=length(filename)/18
    msgbox('零件数量和数据数量不符，请检查零件索引TXT文件');
    return;
else
 TITLE_NAME_INDEX=1;%标题索引
 TEIL_NAME_INDEX=1;%标题索引
 for i=1:(length(filename)/18)
    TITLE_NAME{TITLE_NAME_INDEX}=[TEIL_NAME{TEIL_NAME_INDEX},' X-Richtung'];
    TITLE_NAME{TITLE_NAME_INDEX+1}=[TEIL_NAME{TEIL_NAME_INDEX},' Y-Richtung'];
    TITLE_NAME_INDEX=TITLE_NAME_INDEX+2;
    TEIL_NAME_INDEX=TEIL_NAME_INDEX+1;
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
end 
  RESOLUTION_HE=500;
  RESOLUTION_WI=1300;
  zihao=20;
  n=1;%曲线索引
   if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
  Fileadress=strcat(pathname,'result\');
%%%%%%%%%%%%%%%%%%%求撕裂力最大值并画图%%%%%%%%%%%%%%%%%%%%%
t2=waitbar(0,'正在生成报告图片');
for i=1:(length(filename)/9)
   h(i)=figure;
    set(h(i),'visible','off');
     plot(MP{1,n}(:,1),MP{1,n}(:,2),'linewidth',2);
     hold on;
     plot(MP{1,n+1}(:,1),MP{1,n+1}(:,2),'linewidth',2);
     plot(MP{1,n+2}(:,1),MP{1,n+2}(:,2),'linewidth',2);
    plot(MP{1,n+3}(:,1),MP{1,n+3}(:,2),'linewidth',2);
    plot(MP{1,n+4}(:,1),MP{1,n+4}(:,2),'linewidth',2);
    plot(MP{1,n+5}(:,1),MP{1,n+5}(:,2),'linewidth',2);
    plot(MP{1,n+6}(:,1),MP{1,n+6}(:,2),'linewidth',2);
    plot(MP{1,n+7}(:,1),MP{1,n+7}(:,2),'linewidth',2);
    plot(MP{1,n+8}(:,1),MP{1,n+8}(:,2),'linewidth',2);
     set(h(i),'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]);
   set(h(i),'color','w')
        set(gca,'FontSize',zihao);
        title(TITLE_NAME{i},'FontSize',zihao);
     xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);  
     Ym=max([max(MP{1,n}(:,2)) max(MP{1,n+1}(:,2)) max(MP{1,n+2}(:,2)) max(MP{1,n+3}(:,2)) max(MP{1,n+4}(:,2)) ...
       max(MP{1,n+5}(:,2)) max(MP{1,n+6}(:,2)) max(MP{1,n+7}(:,2)) max(MP{1,n+8}(:,2))])*1.5;
    Xm=max([max(MP{1,n}(:,1)) max(MP{1,n+1}(:,1)) max(MP{1,n+2}(:,1)) max(MP{1,n+3}(:,1)) max(MP{1,n+4}(:,1)) ...
       max(MP{1,n+5}(:,1)) max(MP{1,n+6}(:,1)) max(MP{1,n+7}(:,1)) max(MP{1,n+8}(:,1))])*1.3;
      STAND_X=[0;Xm/1.3];
   STAND_Y=[50;50];
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
biaotihao=10;
headline=['III. Einzelergebnis 具体结果'];
Content.Start=biaotihao; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=10; % 字体大小
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Start = Selection.end; 
Selection.TypeParagraph;% 插入一个新的空段落



InlineShapes=Document.InlineShapes;
t3=waitbar(0.6);
for i=1:length(filename)/9
    sfilename1=[Fileadress,num2str(i),'.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
delete(sfilename1); 
end
t3=waitbar(0.9);

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
t3=waitbar(1);

%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='IZAF底护板撕裂力试验';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%


close(t3);
set(handles.pushbutton2,'Enable','off'); 
winopen([Fileadress,'report.doc']);


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
