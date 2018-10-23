function varargout = Auto5_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto5_1_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto5_1_OutputFcn, ...
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


% --- Executes just before Auto5_1 is made visible.
function Auto5_1_OpeningFcn(hObject, eventdata, handles, varargin)
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

% UIWAIT makes Auto5_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto5_1_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit1_Callback(hObject, eventdata, handles)

function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)

function popupmenu2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
handles=guihandles;
guidata(hObject,handles);

[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
  return;

else
PART={'A-Saeule/A��';'B-Saeule/B��';'C-Saeule/C��'};
CLIP_NUMBER=str2num(get(handles.edit1,'String'));
if isempty(CLIP_NUMBER)
    msgbox('�����뿨������');
    return
elseif length(filename)~=CLIP_NUMBER*6
    msgbox('�������뿨����������������ԭʼ����')    
    return
end

Val1=get(handles.popupmenu2,'Value');
end

  t1=waitbar(0,'���ڶ�������');
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
         
  RESOLUTION_HE=500;
  RESOLUTION_WI=1300;
  zihao=20;


    if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
 end
  Fileadress=strcat(pathname,'result\');
  
  %%%%%%%%%%%%%%%%%%%����ͼƬ%%%%%%%%%%%%%%%%%%%%%
  t2=waitbar(0,'�������ɱ���ͼƬ');
  

switch Val1
	case 1
       TITLE_NAME={'Zug mit Clip A-Saeule links';'Druck mit Clip A-Saeule links';'Druck ohne Clip A-Saeule links';...
      'Zug mit Clip A-Saeule rechts';'Druck mit Clip A-Saeule rechts';'Druck ohne Clip A-Saeule rechts';};
    case 2
       TITLE_NAME={'Zug mit Clip B-Saeule links';'Druck mit Clip B-Saeule links';'Druck ohne Clip B-Saeule links';...
      'Zug mit Clip B-Saeule rechts';'Druck mit Clip B-Saeule rechts';'Druck ohne Clip B-Saeule rechts';};
	case 3
       TITLE_NAME={'Zug mit Clip C-Saeule links';'Druck mit Clip C-Saeule links';'Druck ohne Clip C-Saeule links';...
      'Zug mit Clip C-Saeule rechts';'Druck mit Clip C-Saeule rechts';'Druck ohne Clip C-Saeule rechts';};
end
for i=1:CLIP_NUMBER
    LEGEND_NAME{i}=['Clip',num2str(i)];    
end

n=1;
for i=1:(length(filename)/CLIP_NUMBER)
     h=figure;
     hold on;
     set(h,'visible','off');
       for j=1:CLIP_NUMBER
            plot(MP{1,n}(:,1),MP{1,n}(:,2),'linewidth',2);
            Ym(j)=max(MP{1,n}(:,2));
            Kraft_max(n)=max(MP{1,n}(:,2));      %���������
            Xm(j)=max(MP{1,n}(:,1));
            n=n+1;
       end
      hold off; 
      Ymax=max(Ym)*1.1;
      Xmax=max(Xm)*1.1;
      set(h,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]);
      set(h,'color','w')
      set(gca,'FontSize',zihao);
      title(TITLE_NAME{i},'FontSize',zihao);
      xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);
      grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xmax 0 Ymax]);
      legend(LEGEND_NAME,'Location','SouthEast');
      sfilename1=[Fileadress,num2str(i),'.jpg'];
      f=getframe(h);
      imwrite(f.cdata,sfilename1);
      close(h);
      waitbar(i/(length(filename)/CLIP_NUMBER));
end
  close(t2);
  
  
  %%%%%%%%%%%���ɱ���%%%%%%%%%%%%%%%%%
   t3=waitbar(0,'��������Word����');
   switch Val1
       case 1
            filespec_user=[Fileadress,'A_Saeule.doc'];
       case 2
           filespec_user=[Fileadress,'B_Saeule.doc'];
       case 3
           filespec_user=[Fileadress,'C_Saeule.doc'];
   end
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
 t3=waitbar(0.1);
Content=Document.Content;
Selection=Word.Selection;
Paragraphformat=Selection.ParagraphFormat;

%===�ĵ���ҳ�߾�===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;
biaotihao=10;
headline=['III. Einzelergebnis ������'];
Content.Start=biaotihao; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=10; % �����С
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���
Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���
  
  Tab1 = Document.Tables.Add(Selection.Range, 8,5+CLIP_NUMBER);
DTI = Document.Tables.Item(1); % �����
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����

lc=28.381133333333333333333333333333; %���׻���
column_width = [1.63*lc,1.61*lc,1.64*lc,3.11*lc];

DTI.Row.Item(3).Height = 1.24*lc;
DTI.Row.Item(6).Height = 1.24*lc;

for i=1:CLIP_NUMBER
    column_width(1,i+4)=1.69*lc;
end
column_width(1,5+CLIP_NUMBER)=1.69*lc;

for i = 1:(5+CLIP_NUMBER)
DTI.Columns.Item(i).Width = column_width(i);
end
 DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
 DTI.Range.Font.NameAscii='Arial';
 DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
 
 t3=waitbar(0.2);
 DTI.Cell(1,1).Merge(DTI.Cell(1,5+CLIP_NUMBER));
 DTI.Cell(1,1).Range.Font.Bold=1;
 DTI.Cell(1,1).Range.Text = 'Abzugskraefte Saeulenverkleidungen (N)/���������������N��';
 DTI.Cell(2,1).Range.Text = 'Pr��fling Nr./��������';
 DTI.Cell(2,2).Merge(DTI.Cell(2,3));
 DTI.Cell(2,2).Range.Text = 'Bauteil/        ���������';
 DTI.Cell(2,3).Range.Text = 'Richtung/      ���鷽��';
 for i = 1:CLIP_NUMBER
DTI.Cell(2,i+3).Range.Text =['Clip ',num2str(i),'/����',num2str(i)];
end
DTI.Cell(2,4+CLIP_NUMBER).Range.Text = 'Sollwert/����ֵ';
for i = 1:6 
DTI.Cell(i+2,1).Range.Text = num2str(i);    
end
 DTI.Cell(3,2).Merge(DTI.Cell(8,2));
 DTI.Cell(3,2).Range.Text=PART{Val1};
 DTI.Cell(3,3).Merge(DTI.Cell(5,3));
 DTI.Cell(6,3).Merge(DTI.Cell(8,3));
 DTI.Cell(3,3).Range.Text='Links/���';
 DTI.Cell(6,3).Range.Text='Rechts/�Ҳ�';
 DTI.Cell(3,4).Range.Text='Zug/��';
 DTI.Cell(4,4).Range.Text='Druck mit Clip/��������ѹ';
 DTI.Cell(5,4).Range.Text='Druck ohne Clip/��������ѹ';
 DTI.Cell(6,4).Range.Text='Zug/��';
 DTI.Cell(7,4).Range.Text='Druck mit Clip/��������ѹ';
 DTI.Cell(8,4).Range.Text='Druck ohne Clip/��������ѹ';
  
 t3=waitbar(0.3);
 %%%%%%%%%%%%%%�����Kraft%%%%%%%%%%%5
 for i=1:CLIP_NUMBER
     Kraft_output(1,i)=Kraft_max(i);
     Kraft_output(2,i)=Kraft_max(i+CLIP_NUMBER);
     Kraft_output(3,i)=Kraft_max(i+2*CLIP_NUMBER);
     Kraft_output(4,i)=Kraft_max(i+3*CLIP_NUMBER);
     Kraft_output(5,i)=Kraft_max(i+4*CLIP_NUMBER);
     Kraft_output(6,i)=Kraft_max(i+5*CLIP_NUMBER);
 end
 %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 for i=1:CLIP_NUMBER
     for j=1:6
         DTI.Cell(j+2,i+4).Range.Text=num2str( Kraft_output(j,i),'%.0f'); 
         if Kraft_output(j,i)<399.5
           DTI.Cell(j+2,i+4).Range.Font.Colorindex='wdRed';
           DTI.Cell(j+2,i+4).Range.Font.Bold=1; 
         end
     end
 end
  
 t3=waitbar(0.6);
 Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
InlineShapes=Document.InlineShapes;
for i=1:1:(length(filename)/CLIP_NUMBER)
    sfilename1=[Fileadress,num2str(i),'.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
delete(sfilename1); 
end
 
 t3=waitbar(0.8);

%%%%%%%%%%%%�������������Ϣ�������ռ�%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
switch Val1
    case 1
TEST_NAME='A���������忨��ǿ��';
    case 2
TEST_NAME='B���������忨��ǿ��'; 
    case 3
TEST_NAME='C���������忨��ǿ��';
end
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
 
 t3=waitbar(0.9);
Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
Word.Quit; % �ر��ĵ�
winopen(filespec_user);
close(t3);
