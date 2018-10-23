function varargout = Auto8_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto8_1_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto8_1_OutputFcn, ...
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


% --- Executes just before Auto8_1 is made visible.
function Auto8_1_OpeningFcn(hObject, eventdata, handles, varargin)
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

% UIWAIT makes Auto8_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto8_1_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
list=get(handles.Fahrzeugcode,'String');
val1=get(handles.Fahrzeugcode,'Value');
b1=list{val1};
b4=get(handles.edit4,'String');
if isempty(b1)||isempty(b4)
    msgbox('请输入相关内容');
    return;
end
    
global PATH PATH_INFORMATION SUBDIRPATH IMAGES
PATH=uigetdir;
if PATH==0
        msgbox('请选择文件夹');
    return;
end
file_usr=strcat(cd,'\model\GOM.docx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=[PATH,'\',b4,'_Anhang2_Photogrammetrie_GOM.docx'];
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);

DIR_NAME={'2.1前保险杠总成';'2.2后保险杠总成';'2.3后尾翼';'2.4仪表';...
    '2.5副仪表';'2.6左侧柱护板';'2.7右侧柱护板';'2.8后备箱';'2.9顶棚和上柱护板';...
    '2.10左前门护板';'2.11左后门护板';'2.12右前门护板';'2.13右后门护板';'2.14B柱外盖板';'2.15门外防擦条'};

ZI_DIR_NAME1={'2.8.1后盖护面';'2.8.2后备箱地毯_侧护面'};
ZI_DIR_NAME2={'2.9.1Nach 55℃ Lagerung';'2.9.2Nach 90℃ Lagerung';'2.9.3Nach -30℃ Lagerung';'2.9.4Nach 100℃ Lagerung'};

DIR_NAME_TABLE={'前保险杠总成';'后保险杠总成';'后尾翼';'仪表';...
    '副仪表';'左侧柱护板';'右侧柱护板';'后备箱';'顶棚和上柱护板';...
    '左前门护板';'左后门护板';'右前门护板';'右后门护板';'B柱外盖板';'门外防擦条'};

ZI_DIR_NAME1_TABLE={'后盖护面';'后备箱地毯_侧护面'};
ZI_DIR_NAME2_TABLE={'Nach 55℃ Lagerung';'Nach 90℃ Lagerung';'Nach -30℃ Lagerung';'Nach 100℃ Lagerung'};

%%%%%%%%%%%%%%%%%%%%%%主目录下包含子文件夹的情况%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
       
t1=waitbar(0,'正在新建Word文档');
filespec_user=[PATH,'\',b4,'_Anhang2_Photogrammetrie_GOM.docx'];
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
waitbar(0.5);
   Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;

%===文档的页边距===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
biaotihao=10;
He=180*0.94488188976377952755905511811024*1.7683;
Wi=240*1.9681;

headline=['V.	Anhang2 zu ',b4,'：Photogrammetrie/GOM 附件'];
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=biaotihao; % 字体大小
Content.Font.Bold=1;
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落
Selection.Font.Bold=0;


waitbar(0.7);

%%%%%%%%%%%%生成报告信息表格%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%5

DIR_INDEX=1;
TABLE_INDEX=1;
 for i=1:7
        SUBDIRPATH = fullfile(PATH,DIR_NAME{i},'*.png' );
IMAGES=dir(SUBDIRPATH);   % 在这个子文件夹下找后缀为png的文件
if isempty(IMAGES)
    continue
else
   PHOTO_NUM(TABLE_INDEX)=length(IMAGES); 
   PHONTO_NAME{TABLE_INDEX}=['V.2.',num2str(DIR_INDEX),' ',DIR_NAME_TABLE{i}];
   DIR_INDEX=DIR_INDEX+1;
   TABLE_INDEX=TABLE_INDEX+1;
end
 end
 
SUBDIRPATH1 = fullfile(PATH,DIR_NAME{8},ZI_DIR_NAME1{1},'*.png' );
SUBDIRPATH2 = fullfile(PATH,DIR_NAME{8},ZI_DIR_NAME1{2},'*.png' );
if ~isempty(dir(SUBDIRPATH1))||~isempty(dir(SUBDIRPATH2))
u1=1;
if ~isempty(dir(SUBDIRPATH1))
    IMAGES=dir(SUBDIRPATH1);
    PHOTO_NUM(TABLE_INDEX)=length(IMAGES); 
   PHONTO_NAME{TABLE_INDEX}=['V.2.',num2str(DIR_INDEX),'.',num2str(u1),' ',ZI_DIR_NAME1_TABLE{1}];
   TABLE_INDEX=TABLE_INDEX+1;
   u1=u1+1;
   end

if ~isempty(dir(SUBDIRPATH2))
    IMAGES=dir(SUBDIRPATH2);
    PHOTO_NUM(TABLE_INDEX)=length(IMAGES); 
   PHONTO_NAME{TABLE_INDEX}=['V.2.',num2str(DIR_INDEX),'.',num2str(u1),' ',ZI_DIR_NAME1_TABLE{2}];
   TABLE_INDEX=TABLE_INDEX+1;
  end
DIR_INDEX=DIR_INDEX+1;
end


SUBDIRPATH3 = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{1},'*.png' );
SUBDIRPATH4 = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{2},'*.png' );
SUBDIRPATH5 = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{3},'*.png' );
SUBDIRPATH6 = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{4},'*.png' );
if ~isempty(dir(SUBDIRPATH3))||~isempty(dir(SUBDIRPATH4))||~isempty(dir(SUBDIRPATH5))||~isempty(dir(SUBDIRPATH6))
    u1=1;
if ~isempty(dir(SUBDIRPATH3))
    IMAGES=dir(SUBDIRPATH3);
    PHOTO_NUM(TABLE_INDEX)=length(IMAGES); 
   PHONTO_NAME{TABLE_INDEX}=['V.2.',num2str(DIR_INDEX),'.',num2str(u1),' ',ZI_DIR_NAME2_TABLE{1}];
   TABLE_INDEX=TABLE_INDEX+1;
   u1=u1+1;
 end

if ~isempty(dir(SUBDIRPATH4))
    IMAGES=dir(SUBDIRPATH4);
    PHOTO_NUM(TABLE_INDEX)=length(IMAGES); 
   PHONTO_NAME{TABLE_INDEX}=['V.2.',num2str(DIR_INDEX),'.',num2str(u1),' ',ZI_DIR_NAME2_TABLE{2}];
   TABLE_INDEX=TABLE_INDEX+1;
   u1=u1+1;
end

if ~isempty(dir(SUBDIRPATH5))
    IMAGES=dir(SUBDIRPATH5);
    PHOTO_NUM(TABLE_INDEX)=length(IMAGES); 
   PHONTO_NAME{TABLE_INDEX}=['V.2.',num2str(DIR_INDEX),'.',num2str(u1),' ',ZI_DIR_NAME2_TABLE{3}];
   TABLE_INDEX=TABLE_INDEX+1;
   u1=u1+1;
end

if ~isempty(dir(SUBDIRPATH6))
    IMAGES=dir(SUBDIRPATH6);
    PHOTO_NUM(TABLE_INDEX)=length(IMAGES); 
   PHONTO_NAME{TABLE_INDEX}=['V.2.',num2str(DIR_INDEX),'.',num2str(u1),' ',ZI_DIR_NAME2_TABLE{4}];
   TABLE_INDEX=TABLE_INDEX+1;
  end
 DIR_INDEX=DIR_INDEX+1;
end

for i=10:15
        SUBDIRPATH = fullfile(PATH,DIR_NAME{i},'*.png' );
IMAGES=dir(SUBDIRPATH);   % 在这个子文件夹下找后缀为png的文件
if isempty(IMAGES)
    continue
else
   PHOTO_NUM(TABLE_INDEX)=length(IMAGES); 
   PHONTO_NAME{TABLE_INDEX}=['V.2.',num2str(DIR_INDEX),' ',DIR_NAME_TABLE{i}];
   DIR_INDEX=DIR_INDEX+1;
   TABLE_INDEX=TABLE_INDEX+1;
    end
end


 headline=['报告信息:'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落  
 
Tab1 = Document.Tables.Add(Selection.Range,TABLE_INDEX,2);
DTI = Document.Tables.Item(1); % 表格句柄
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % 最外框，实线
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % 所有的内框线条
lc=28.381133333333333333333333333333; %厘米换算
column_width = [lc*7,lc*2.24];
for i = 1:2
DTI.Columns.Item(i).Width = column_width(i);
end
for i=1:(TABLE_INDEX+1)
    for j=1:2
        DTI.Cell(i,j).Range.Paragraphs.Alignment='wdAlignParagraphCenter';
        DTI.Cell(i,j).Range.Font.NameAscii='Arial';
    end
end
DTI.Cell(1,1).Range.Text = '目录';
DTI.Cell(1,2).Range.Text = '张数';
for i=1:(TABLE_INDEX-1)
    DTI.Cell(i+1,1).Range.Text =PHONTO_NAME{i};
DTI.Cell(i+1,2).Range.Text =num2str(PHOTO_NUM(i));
end
Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Content.end;
Selection.TypeParagraph;
waitbar(1);
close(t1);
t2=waitbar(0,'正在粘贴图片');

KONGGE=['    '];
DIR_NAME_WORD={'STOSSFAENGER,VORN前保险杠总成';'STOSSFAENGER,HINTEN后保险杠总成';...
    'Heckspoiler后尾翼';'INSTRUMENTENTAFEL仪表';'MITTELKONSOLE副仪表';...
    'VERKLEIDUNG, SAEULE, LINKS左侧柱护板';'VERKLEIDUNG, SAEULE, RECHTS右侧柱护板';...
    'KOFFERRAUM后备箱';'FORMHIMMEL UND VERKLEIDUNG, SAEULE, OBEN顶棚和上柱护板';...
    'TUERVERKLEIDUNG, FS左前门护板';'TUERVERKLEIDUNG, FS_hinten左后门护板';...
    'TUERVERKLEIDUNG, BFS右前门护板';'TUERVERKLEIDUNG, BFS_hinten右后门护板';...
    'VERKLEIDUNG SAEULE B B柱外盖板';'SCHWELLER-BEPLANKUNG门外防擦条'};

ZI_DIR_NAME1_WORD={'VERKLEIDUNG, HECKKLAPPE后盖护面';'BODEN/VERKLEIDUNG, KOFFER 后备箱地毯/侧护面'};
ZI_DIR_NAME2_WORD={'Nach 55℃ Lagerung';'Nach 90℃ Lagerung';'Nach -30℃ Lagerung';'Nach 100℃ Lagerung'};
headline=['V.2	      GOM MESSUNGEN 　GOM测量'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
headline=['TEBE外饰开发科'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
waitbar(0.1);
k=1; %%图像句柄
n=1;%标题序号
        for i=1:3
        SUBDIRPATH = fullfile(PATH,DIR_NAME{i},'*.png' );
IMAGES=dir(SUBDIRPATH);   % 在这个子文件夹下找后缀为png的文件
if ~isempty(IMAGES)
headline=['  V.2.',num2str(n),KONGGE,DIR_NAME_WORD{i}];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

    for j=1:length(IMAGES)
IMAGEPATH=fullfile(PATH,DIR_NAME{i},IMAGES(j).name);
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
Selection.Start = Selection.end;
Selection.TypeParagraph;
k=k+1;
    end
n=n+1;
end
        end

Selection.Start = Selection.end;
Selection.TypeParagraph;
headline=['TEBI内饰开发科'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
waitbar(0.2);
for i=4:7
        SUBDIRPATH = fullfile(PATH,DIR_NAME{i},'*.png' );
IMAGES=dir(SUBDIRPATH);   % 在这个子文件夹下找后缀为png的文件
if ~isempty(IMAGES)
headline=['  V.2.',num2str(n),KONGGE,DIR_NAME_WORD{i}];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

    for j=1:length(IMAGES)
IMAGEPATH=fullfile(PATH,DIR_NAME{i},IMAGES(j).name);
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
Selection.Start = Selection.end;
Selection.TypeParagraph;
k=k+1;
    end
n=n+1;
end
end
waitbar(0.5);
SUBDIRPATH1 = fullfile(PATH,DIR_NAME{8},ZI_DIR_NAME1{1},'*.png' );
SUBDIRPATH2 = fullfile(PATH,DIR_NAME{8},ZI_DIR_NAME1{2},'*.png' );
if isempty(dir(SUBDIRPATH1))&&isempty(dir(SUBDIRPATH2))

else
  headline=['  V.2.',num2str(n),KONGGE,DIR_NAME_WORD{8}];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

p1=1;
for p=1:2   %%子文件夹指针
SUBDIRPATH = fullfile(PATH,DIR_NAME{8},ZI_DIR_NAME1{p},'*.png' );
IMAGES=dir(SUBDIRPATH);   % 在这个子文件夹下找后缀为png的文件
if ~isempty(IMAGES)
headline=['  V.2.',num2str(n),'.',num2str(p1),KONGGE,ZI_DIR_NAME1_WORD{p}];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

  for j=1:length(IMAGES)
IMAGEPATH=fullfile(PATH,DIR_NAME{8},ZI_DIR_NAME1{p},IMAGES(j).name);
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
Selection.Start = Selection.end;
Selection.TypeParagraph;
k=k+1;
    end
p1=p1+1;
end 
end
n=n+1;
end  
waitbar(0.7);

%%%%%%%%%%%%%%%%%%%%%%%%%%%%%顶棚和上柱护板子文件夹%%%%%%%%%%%%%%%%%%%%%%%%%%
SUBDIRPATH3 = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{1},'*.png' );
SUBDIRPATH4 = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{2},'*.png' );
SUBDIRPATH5 = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{3},'*.png' );
SUBDIRPATH6 = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{4},'*.png' );

if isempty(dir(SUBDIRPATH3))&&isempty(dir(SUBDIRPATH4))&&isempty(dir(SUBDIRPATH5))&&isempty(dir(SUBDIRPATH6))

else
  headline=['  V.2.',num2str(n),KONGGE,DIR_NAME_WORD{9}];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

g1=1; %%子文件夹指针
for g=1:4   %%子文件夹指针
SUBDIRPATH = fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{g},'*.png' );
IMAGES=dir(SUBDIRPATH);   % 在这个子文件夹下找后缀为png的文件
if ~isempty(IMAGES)
headline=['  V.2.',num2str(n),'.',num2str(g1),KONGGE,ZI_DIR_NAME2_WORD{g}];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

  for j=1:length(IMAGES)
IMAGEPATH=fullfile(PATH,DIR_NAME{9},ZI_DIR_NAME2{g},IMAGES(j).name);
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
Selection.Start = Selection.end;
Selection.TypeParagraph;
k=k+1;
    end
g1=g1+1;
end 
end
n=n+1;
end  
waitbar(0.8);
%%%%%%%%%%%%%%%%%%%%%%%余下文件夹照片%%%%%%%%%%%%%%%%%%%%%%5
for i=9:15
        SUBDIRPATH = fullfile(PATH,DIR_NAME{i},'*.png' );
IMAGES=dir(SUBDIRPATH);   % 在这个子文件夹下找后缀为png的文件
if ~isempty(IMAGES)
headline=['  V.2.',num2str(n),KONGGE,DIR_NAME_WORD{i}];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

    for j=1:length(IMAGES)
IMAGEPATH=fullfile(PATH,DIR_NAME{i},IMAGES(j).name);
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
Selection.Start = Selection.end;
Selection.TypeParagraph;
k=k+1;
    end
n=n+1;
end
end
waitbar(0.9);
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
%%%%%%%%%%%%输出报告生成信息到公共空间%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='Tritop报告';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
waitbar(1);
winopen(filespec_user);
close(t2);  
  
    
    



function edit4_Callback(hObject, eventdata, handles)

function edit4_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
