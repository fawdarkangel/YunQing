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
    msgbox('�������������');
    return;
else
global PATH PATH_INFORMATION SUBDIRPATH IMAGES
PATH=uigetdir;
if PATH==0
        msgbox('��ѡ���ļ���');
    return;
end
t1=waitbar(0,'���ڴ���Word�ĵ�');
file_usr=strcat(cd,'\model\TEBER18AXXXX_Anhang1_Fotos Teile.docx');
 copy_usr=['copy ','"',file_usr,'"'] ;
filespec_user=[PATH,'\',b2,'_Anhang1_Fotos Teile.docx'];
copy_tal=['"',filespec_user,'"'];
xyz=[copy_usr,' ',copy_tal];
dos(xyz);
DIR_NAME={'1-ǰ���ո�����';'2-���ո�����';'3-��������';'4-�׻�������';...
    '5-ǰ������';'6-�������';'7-������';'8-��������';'9-�Ǳ�����';...
    '10-���Ǳ�����';'11-������͵�̺����';'12-��������';...
    '13-�����������������';'14-��������';'15-������'};
ZI_DIR_NAME={'1-Vor';'2-Nach'};

DIR_NAME_WORD={'STOSSFAENGER��VORNǰ���ո�����';'STOSSFAENGER��HITEN���ո�����';...
    'RADHAUSSCHALE��������';'BODENSCHUTZ�׻�������';'KLAPPE��VORNǰ������';...
    'KLAPPE��HINTEN�������';'TUER������';...
    'DACH��������';'INSTRUMENTENTAFEL�Ǳ�����';...
    'MITTELKONSOLE���Ǳ�����';'VERKLEIDUNG, SAEULE UND BODEN������͵�̺����';...
    'ZSB SITZ��������';'GREENHOUSE�����������������';...
    'KOFFERRAUM��������';'TUER������'};
n=1;
waitbar(0.3);
for i=1:15
    for j=1:2
                DIR_PATH{n}=[PATH,'\',DIR_NAME{i},'\',ZI_DIR_NAME{j},'\'];
                DIR_PATH_INDEX{n}=[PATH,'\',DIR_NAME{i},'\',ZI_DIR_NAME{j},'\*.jpg'];
                n=n+1;
    end
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%���ɱ���%%%%%%%%%%%%%%%%%%%%%%%%%%
waitbar(0.7);
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
Word.Visible =0; % ʹwordΪ�ɼ�����set(Word, 'Visible', 1); 
%===��word�ļ������·����û���򴴽�һ���հ��ĵ���========================
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

%===�ĵ���ҳ�߾�===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
biaotihao=10;
He=6.22*28.3579;
Wi=8.31*28.3579;
KONGGE=['    '];
headline=['V.	Anhang1 zu ',b2,'��Bilder Pruefstelle ����'];
Content.Start=0; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=biaotihao; % �����С
Content.Font.Bold=1;
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն��� 
Selection.Font.Bold=0;
waitbar(1);
close(t1);
t2=waitbar(0,'�������ɱ���');

 headline=['  V.1.	   Fotes vor/nach der Pruefung ����ǰ����Ƭ'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն���  
 headline=['	       Im nachfolgenden����:'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 
headline=['	     ? V: Bild vor der Pruefung   ����ǰͼƬ'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 

headline=['	     ? N: Bild nach der Pruefung   �����ͼƬ'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 


headline=['TEBE���ο�����'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 

k=1; %%ͼ����
n=1;%�������
z=1;%%�ӱ������

for i=1:8
  if ~isempty(dir(DIR_PATH_INDEX{z}))&&~isempty(DIR_PATH_INDEX{z+1})
      headline=['  V.1.',num2str(n),KONGGE,DIR_NAME_WORD{i}];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 
headline=['      V'];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 

       IMAGES=dir(DIR_PATH_INDEX{z});   % ��������ļ������Һ�׺ΪJPG���ļ�
        
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
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 

    IMAGES=dir(DIR_PATH_INDEX{z+1});   % ��������ļ������Һ�׺ΪJPG���ļ�
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
Selection.TypeParagraph;% ����һ���µĿն��� 
headline=['TEBE���ο�����'];
Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 

for i=9:15
  if ~isempty(dir(DIR_PATH_INDEX{z}))&&~isempty(DIR_PATH_INDEX{z+1})
      headline=['  V.1.',num2str(n),KONGGE,DIR_NAME_WORD{i}];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 
headline=['      V'];
      Selection.Text=headline;
Selection.Font.NameAscii='Arial';
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 
       IMAGES=dir(DIR_PATH_INDEX{z});   % ��������ļ������Һ�׺ΪJPG���ļ�
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
Selection.Font.Size=biaotihao; % �����С
Selection.Start=Selection.end;
Selection.TypeParagraph;% ����һ���µĿն��� 
    IMAGES=dir(DIR_PATH_INDEX{z+1});   % ��������ļ������Һ�׺ΪJPG���ļ�
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
Document.Save; % �����ĵ�
Word.Quit; % �ر��ĵ�
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
