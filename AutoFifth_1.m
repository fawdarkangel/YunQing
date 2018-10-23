function varargout = AutoFifth_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @AutoFifth_1_OpeningFcn, ...
                   'gui_OutputFcn',  @AutoFifth_1_OutputFcn, ...
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


% --- Executes just before AutoFifth_1 is made visible.
function AutoFifth_1_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')
%[a b c]=xlsread('\\faw-vw\fs\org\PE\T-E-VC-2\07_������mearusing group\12-���ݴ���ƽ̨\resource\Fahrzeugcode.xlsx','Tabelle1','B:B');
b=load([cd,'\interface\Fahrzeugcode.mat'])
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes AutoFifth_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = AutoFifth_1_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;

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


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);
global newpath

oldpath=cd;
if isempty(newpath)|| ~exist('newpath')
     newpath=cd;
 end
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������',newpath);
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
  return;
else
    t1=waitbar(0,'���ڶ�ȡ����');
    zihao=20;
    newpath=pathname; 
    Filename=strcat(pathname,filename);
    [Type Sheet Format]=xlsfinfo(Filename)
    for i=1:length(Sheet)
    MP{i}=xlsread(Filename,Sheet{i}); 
    waitbar(i/length(Sheet));
    end
     if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
close(t1);

end
 t2=waitbar(0,'�������ɱ���ͼƬ');
for i=1:length(Sheet)
    if ~isempty(MP{i})
MIANJI(i)=-trapz(MP{i}(:,1)./1000,MP{i}(:,2));
    end
end
ZIHAO_WENZI=10;%���������ֺ�
ZIHAO_TU=20;%����ͼƬ�ֺ�
    Fileaddress=strcat(pathname,'result\');
for i=1:length(Sheet)
    
   
       h(i)=figure;
       set(h(i),'visible','off');
       plot(MP{i}(:,1),MP{i}(:,2)./1000,'linewidth',2);
        set(h(i),'position',[100,100,1300,800]); 
        z=ceil(max(MP{i}(:,2))/1000+3);
        z_x=ceil(max(MP{i}(:,1))+10);
        axis([0 z_x 0 z]);grid on;
        grid minor;
        set(gca,'FontSize',ZIHAO_TU);
        set(gca,'LineWid',2)
         xlabel('Federweg S/mm','FontSize',ZIHAO_TU);ylabel('Federkraft F/KN','FontSize',ZIHAO_TU);title(Sheet{i},'FontSize',ZIHAO_TU);
         sfilename1=[Fileaddress,num2str(i),'.jpg'];
        saveas(h(i),sfilename1);
        close(h(i));
        waitbar(i/length(Sheet));
   end
   close(t2);
    
    
      t3=waitbar(0,'��������Word����')   
         biaotihao=10;
He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
filespec_user=[Fileaddress,'report.doc'];

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
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
headline='III.1 ������';
Content.Start=0; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=biaotihao; % �����С
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���         
 Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���  
waitbar(0.3);
 %%�������ݱ��
Tab1 = Document.Tables.Add(Selection.Range,length(Sheet), 2);
DTI = Document.Tables.Item(1); % �����
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����
% �����иߣ��п�
lc=28.381133333333333333333333333333; %���׻���
column_width = [3*lc,3*lc];
waitbar(0.6);

for i = 1:2
DTI.Columns.Item(i).Width = column_width(i);
end
 DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
 DTI.Range.Font.NameAscii='Arial';
 DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
    
 for i=1:length(Sheet)
     DTI.Cell(i,1).Range.Text = Sheet{i};
     DTI.Cell(i,2).Range.Text =num2str(MIANJI(i),'%.1f');
 end
 
 Selection.Start = Content.end;
Selection.TypeParagraph;
  Selection.Start = Selection.end;
Selection.TypeParagraph;
 InlineShapes=Document.InlineShapes;
 for i=1:length(Sheet)
Teil2address{i}= [Fileaddress,num2str(i),'.jpg'];
end

 for i=1:length(Sheet)
handle=Selection.InlineShapes.AddPicture(Teil2address{1,i});
InlineShapes.Item(i).Height=He;
InlineShapes.Item(i).Width=Wi;
delete(Teil2address{i});
end
waitbar(0.8)
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
Word.Quit; % �ر��ĵ�
 

FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='����龲̬���Բ���';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
waitbar(0.9)

close(t3);
winopen([Fileaddress,'report.doc']);

function edit2_Callback(hObject, eventdata, handles)

function edit2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
