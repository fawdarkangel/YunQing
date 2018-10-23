function varargout = Auto6_3(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto6_3_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto6_3_OutputFcn, ...
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


% --- Executes just before Auto6_3 is made visible.
function Auto6_3_OpeningFcn(hObject, eventdata, handles, varargin)
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

function varargout = Auto6_3_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)         %ѡ������
DATA=getappdata(0,'DATA');
MP=DATA.MP;
MAX1_index=DATA.MAX1_index;
MAX2_index=DATA.MAX2_index;
CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
i=CHOOSE;
ZIHAO_TU_YULAN=10;
cla(handles.axes1);
plot(handles.axes1,MP{i}(:,1),MP{i}(:,2)/1000,'linewidth',2);
   
plot(handles.axes1,MP{i}(MAX1_index(i),1),MP{i}(MAX1_index(i),2)/1000,'*r');
plot(handles.axes1,MP{i}(MAX2_index(i),1),MP{i}(MAX2_index(i),2)/1000,'*r');
z=max(MP{i}(:,2))*1.1/1000;
axis(handles.axes1,[0 inf 0 z]);grid on;grid minor;        
xlabel(handles.axes1,'Weg(mm)','FontSize',ZIHAO_TU_YULAN);ylabel(handles.axes1,'Kraft(kN)','FontSize',ZIHAO_TU_YULAN);
text(handles.axes1,MP{i}(MAX1_index(i),1),MP{i}(MAX1_index(i),2)/1000,['\leftarrow(',num2str(MP{i}(MAX1_index(i),2),'%.f'),'N)'],'FontSize',ZIHAO_TU_YULAN);
text(handles.axes1,MP{i}(MAX2_index(i),1),MP{i}(MAX2_index(i),2)/1000,['\leftarrow(',num2str(MP{i}(MAX2_index(i),2),'%.f'),'N)'],'FontSize',ZIHAO_TU_YULAN);


function listbox1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- ��������.
function pushbutton1_Callback(hObject, eventdata, handles)  
handles=guihandles;
guidata(hObject,handles);

[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
  return;
  
else
    t1=waitbar(0,'���ڶ�������');
    ZIHAO_TU_YULAN=10;
    DATA_LENGTH=length(filename);
  for i=1:DATA_LENGTH
           Filename{i}=strcat(pathname,filename{i});
           [Type Sheet Format]=xlsfinfo(Filename{i}) ;
           sheet{i}=Sheet;
           MP{i}=xlsread(Filename{i},char(sheet{1,i}(1,4)));
           waitbar(i/DATA_LENGTH);
           try
               system('taskkill/IM excel.exe');
           end
  end  
   close(t1);
   
   for i=1:DATA_LENGTH
       LAST_MP=find(MP{i}(:,2)>100,1,'last');%�����һ��������100�ĵ���Ϊ������ֹ��
       MIDDLE_INDEX(i)=ceil(length(MP{i}(1:LAST_MP))/2);
       MAX1_index(i)=find(MP{i}(1:MIDDLE_INDEX(i),2)==max(MP{i}(1:MIDDLE_INDEX(i),2)),1);
       MAX2_index(i)=find(MP{i}(MIDDLE_INDEX(i)+1:end,2)==max(MP{i}(MIDDLE_INDEX(i)+1:end,2)),1)+MIDDLE_INDEX(i);
       MAX_KRAFT1(i)=MP{i}(MAX1_index(i),2);
       MAX_KRAFT2(i)=MP{i}(MAX2_index(i),2);
   end
   %%%%%%%%%%%%%%%%%%%%MP1����Ԥ����%%%%%%%%%%%%%%%%%%%%
   cla(handles.axes1);
   plot(handles.axes1,MP{1}(:,1),MP{1}(:,2)/1000,'linewidth',2);   
   plot(handles.axes1,MP{1}(MAX1_index(1),1),MP{1}(MAX1_index(1),2)/1000,'*r');
   plot(handles.axes1,MP{1}(MAX2_index(1),1),MP{1}(MAX2_index(1),2)/1000,'*r');
   z=max(MP{1}(:,2))*1.1/1000;
   axis(handles.axes1,[0 inf 0 z]);grid on;grid minor;        
   xlabel(handles.axes1,'Weg(mm)','FontSize',ZIHAO_TU_YULAN);ylabel(handles.axes1,'Kraft(kN)','FontSize',ZIHAO_TU_YULAN);
   text(handles.axes1,MP{1}(MAX1_index(1),1),MP{1}(MAX1_index(1),2)/1000,['\leftarrow(',num2str(MP{1}(MAX1_index(1),2),'%.f'),'N)'],'FontSize',ZIHAO_TU_YULAN);
   text(handles.axes1,MP{1}(MAX2_index(1),1),MP{1}(MAX2_index(1),2)/1000,['\leftarrow(',num2str(MP{1}(MAX2_index(1),2),'%.f'),'N)'],'FontSize',ZIHAO_TU_YULAN);
   set(handles.listbox1,'Value',1);
   set(handles.listbox1,'String',filename); 
   DATA.MP=MP;
   DATA.MAX1_index=MAX1_index;
   DATA.MAX2_index=MAX2_index;
   DATA.pathname=pathname;
   setappdata(0,'DATA',DATA);
end


%%%%%%%%%%%%%����Word����%%%%%%%%%%%%%%%%%
function pushbutton2_Callback(hObject, eventdata, handles)
DATA=getappdata(0,'DATA');
if isempty(DATA)
    msgbox('�뵼������');
    return
end
MP=DATA.MP;
MAX1_index=DATA.MAX1_index;
MAX2_index=DATA.MAX2_index;
pathname=DATA.pathname;
ZIHAO_TU=20;

Fileadress=strcat(pathname,'result\');
if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
end
 t2=waitbar(0,'�������ɱ���ͼƬ');
for i=1:length(MP)

   h=figure;
   set(h,'visible','off');
   plot(MP{i}(:,1),MP{i}(:,2)/1000,'linewidth',2);
   hold on;
   plot(MP{i}(MAX1_index(i),1),MP{i}(MAX1_index(i),2)/1000,'*r');
   plot(MP{i}(MAX2_index(i),1),MP{i}(MAX2_index(i),2)/1000,'*r');
   set(h,'position',[100,100,1300,800]); 
   z=max(MP{i}(:,2))*1.1/1000;
   axis([0 inf 0 z]);grid on;grid minor;
   set(gca,'FontSize',ZIHAO_TU);
   xlabel('Weg(mm)','FontSize',ZIHAO_TU);ylabel('Kraft(kN)','FontSize',ZIHAO_TU);
   title(['Teil ',num2str(i),'#'],'FontSize',ZIHAO_TU);
   text(MP{i}(MAX1_index(i),1),MP{i}(MAX1_index(i),2)/1000,['\leftarrow(',num2str(MP{i}(MAX1_index(i),2),'%.f'),'N)'],'FontSize',ZIHAO_TU);
   text(MP{i}(MAX2_index(i),1),MP{i}(MAX2_index(i),2)/1000,['\leftarrow(',num2str(MP{i}(MAX2_index(i),2),'%.f'),'N)'],'FontSize',ZIHAO_TU);
   sfilename1=[Fileadress,num2str(i),'.jpg'];
   saveas(h,sfilename1);
   close(h);
   waitbar(i/length(MP));   
end
close(t2);

t3=waitbar(0,'��������Word����') ;  
biaotihao=10;
He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
filespec_user=[Fileadress,'report.doc'];
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
waitbar(0.1);
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

%%�������ݱ��
Tab1 = Document.Tables.Add(Selection.Range, length(MP)+1, 5);
DTI = Document.Tables.Item(1); % �����
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����
% �����иߣ��п�
lc=28.381133333333333333333333333333; %���׻���
column_width = [1.19*lc,2.25*lc,3.25*lc,3.25*lc,3*lc];
waitbar(0.3);
for i = 1:5
DTI.Columns.Item(i).Width = column_width(i);
end
DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
DTI.Range.Font.NameAscii='Arial';
DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(2,2).Merge(DTI.Cell(length(MP)+1,2));
DTI.Cell(1,1).Range.Text = 'Nr.������';
DTI.Cell(1,2).Range.Text = 'ForderungҪ��[KN]';
DTI.Cell(1,3).Range.Text = 'Ist-Messwert Erstes Maximum��һ��ֵ[KN]';
DTI.Cell(1,4).Range.Text = 'Ist-Messwert Zweites Maximum�ڶ���ֵ[KN]';
DTI.Cell(1,5).Range.Text = 'Bewertung����';
DTI.Cell(2,2).Range.Text = '��30';
for i=2:length(MP)+1
    DTI.Cell(i,1).Range.Text =num2str(i-1);
    DTI.Cell(i,3).Range.Text =num2str(MP{i-1}(MAX1_index(i-1),2)/1000,'%.2f');
    if MP{i-1}(MAX1_index(i-1),2)<30000
         DTI.Cell(i,3).Range.Font.Colorindex='wdRed';
         DTI.Cell(i,3).Range.Font.Bold=1;
    end
    DTI.Cell(i,4).Range.Text =num2str(MP{i-1}(MAX2_index(i-1),2)/1000,'%.2f');
    if MP{i-1}(MAX2_index(i-1),2)<30000
         DTI.Cell(i,4).Range.Font.Colorindex='wdRed';
         DTI.Cell(i,4).Range.Font.Bold=1;
    end    
    waitbar(0.5);
end
Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
InlineShapes=Document.InlineShapes;
for i=1:length(MP)
    sfilename1=[Fileadress,num2str(i),'.jpg'];
    handle=Selection.InlineShapes.AddPicture(sfilename1);
    delete(sfilename1); 
end
waitbar(0.8);
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
Word.Quit; % �ر��ĵ�
waitbar(0.85);
%%%%%%%%%%%%�������������Ϣ�������ռ�%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='�������ܳɰγ�������';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
waitbar(0.9);
close(t3);
winopen([Fileadress,'report.doc']);


function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
