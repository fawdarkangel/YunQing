function varargout = AutoSecond_1_2(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @AutoSecond_1_2_OpeningFcn, ...
                   'gui_OutputFcn',  @AutoSecond_1_2_OutputFcn, ...
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


% --- Executes just before AutoSecond_1_2 is made visible.
function AutoSecond_1_2_OpeningFcn(hObject, eventdata, handles, varargin)

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);




% --- Outputs from this function are returned to the command line.
function varargout = AutoSecond_1_2_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
b1=get(handles.edit1,'String');
b2=get(handles.edit2,'String');
if isempty(b1)||isempty(b2)
    msgbox('请输入相关内容');
    return;
else
global PATH PATH_INFORMATION SUBDIRPATH IMAGES
PATH=uigetdir;
if PATH==0
        msgbox('请选择文件夹');
    return;
end
t1=waitbar(0,'正在创建Word文档');
file_usr=strcat(cd,'\model\TEBER18AXXXX_Anhang1_Fotos Teile.docx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=[PATH,'\',b2,'_Anhang1_Fotos Teile.docx'];
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
DIR_NAME={'1-前保险杠区域';'2-后保险杠区域';'3-轮罩区域';'4-底护板区域';...
    '5-前盖区域';'6-后盖区域';'7-门区域';'8-顶部区域';'9-仪表区域';...
    '10-副仪表区域';'11-柱护板和地毯区域';'12-座椅区域';...
    '13-顶棚和上柱护板区域';'14-后备箱区域';'15-门区域'};
ZI_DIR_NAME={'1-Vor';'2-Nach'};

DIR_NAME_WORD={'STOSSFAENGER，VORN前保险杠区域';'STOSSFAENGER，HITEN后保险杠区域';...
    'RADHAUSSCHALE轮罩区域';'BODENSCHUTZ底护板区域';'KLAPPE，VORN前盖区域';...
    'KLAPPE，HINTEN后盖区域';'TUER门区域';...
    'DACH顶部区域';'INSTRUMENTENTAFEL仪表区域';...
    'MITTELKONSOLE副仪表区域';'VERKLEIDUNG, SAEULE UND BODEN柱护板和地毯区域';...
    'ZSB SITZ座椅区域';'GREENHOUSE顶棚和上柱护板区域';...
    'KOFFERRAUM后备箱区域';'TUER门区域'};
n=1;
waitbar(0.3);
for i=1:15
    for j=1:2
                DIR_PATH{n}=[PATH,'\',DIR_NAME{i},'\',ZI_DIR_NAME{j},'\'];
                DIR_PATH_INDEX{n}=[PATH,'\',DIR_NAME{i},'\',ZI_DIR_NAME{j},'\*.jpg'];
                n=n+1;
    end
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%生成报告%%%%%%%%%%%%%%%%%%%%%%%%%%
waitbar(0.7);
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
waitbar(0.9);
   Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;

%===文档的页边距===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
biaotihao=10;
He=6.22*28.3579;
Wi=8.31*28.3579;
KONGGE=['    '];
headline=['V.	Anhang1 zu ',b2,'：Bilder Pruefstelle 附件'];
Content.Start=0; % 起始点为0，即表示每次写入覆盖之前资料
Content.Text=headline;
Content.Font.Size=biaotihao; % 字体大小
Content.Font.Bold=1;
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Font.Bold=0;
waitbar(1);
close(t1);
t2=waitbar(0,'正在生成报告');

 headline=['  V.1.	   Fotes vor/nach der Pruefung 试验前后照片'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落  
 headline=['	       Im nachfolgenden以下:'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
headline=['	     ? V: Bild vor der Pruefung   试验前图片'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

headline=['	     ? N: Bild nach der Pruefung   试验后图片'];
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

k=1; %%图像句柄
n=1;%标题序号
z=1;%%子标题序号

for i=1:8
  if ~isempty(dir(DIR_PATH_INDEX{z}))&&~isempty(DIR_PATH_INDEX{z+1})
      headline=['  V.1.',num2str(n),KONGGE,DIR_NAME_WORD{i}];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
headline=['      V'];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

       IMAGES=dir(DIR_PATH_INDEX{z});   % 在这个子文件夹下找后缀为JPG的文件
        
    for j=1:length(IMAGES)
IMAGEPATH=[DIR_PATH{z},IMAGES(j).name];
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
if mod(j,2)==0
   Selection.Start = Selection.end;
Selection.TypeParagraph; 
end
k=k+1;
    end
    Selection.Start = Selection.end;
Selection.TypeParagraph;
    headline=['      N'];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

    IMAGES=dir(DIR_PATH_INDEX{z+1});   % 在这个子文件夹下找后缀为JPG的文件
    for j=1:length(IMAGES)
IMAGEPATH=[DIR_PATH{z+1},IMAGES(j).name];
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
if mod(j,2)==0
   Selection.Start = Selection.end;
Selection.TypeParagraph; 
end
k=k+1;
    end
    Selection.Start = Selection.end;
Selection.TypeParagraph;
    
   
n=n+1;
  end
z=z+2; 
waitbar(i/15);
end

Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
headline=['TEBE外饰开发科'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 

for i=9:15
  if ~isempty(dir(DIR_PATH_INDEX{z}))&&~isempty(DIR_PATH_INDEX{z+1})
      headline=['  V.1.',num2str(n),KONGGE,DIR_NAME_WORD{i}];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
headline=['      V'];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
       IMAGES=dir(DIR_PATH_INDEX{z});   % 在这个子文件夹下找后缀为JPG的文件
    for j=1:length(IMAGES)
IMAGEPATH=[DIR_PATH{z},IMAGES(j).name];
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
if mod(j,2)==0
   Selection.Start = Selection.end;
Selection.TypeParagraph; 
end
k=k+1;
    end
    Selection.Start = Selection.end;
Selection.TypeParagraph;
    headline=['      N'];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % 字体大小
Selection.Start=Selection.end;
Selection.TypeParagraph;% 插入一个新的空段落 
    IMAGES=dir(DIR_PATH_INDEX{z+1});   % 在这个子文件夹下找后缀为JPG的文件
    for j=1:length(IMAGES)
IMAGEPATH=[DIR_PATH{z+1},IMAGES(j).name];
InlineShapes=Document.InlineShapes;
handle=Selection.InlineShapes.AddPicture(IMAGEPATH);
InlineShapes.Item(k).Height=He;
InlineShapes.Item(k).Width=Wi;
if mod(j,2)==0
   Selection.Start = Selection.end;
Selection.TypeParagraph; 
end
k=k+1;
    end
    Selection.Start = Selection.end;
Selection.TypeParagraph;
    
  
n=n+1;
  end
 z=z+2; 
 waitbar(i/15);
end


Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % 保存文档
Word.Quit; % 关闭文档
winopen(filespec_user);
close(t2);
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
