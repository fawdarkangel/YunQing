function varargout = Auto3_3_1(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto3_3_1_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto3_3_1_OutputFcn, ...
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


% --- Executes just before Auto3_3_1 is made visible.
function Auto3_3_1_OpeningFcn(hObject, eventdata, handles, varargin)
movegui(gcf,'center')
b=load([cd,'\interface\Fahrzeugcode.mat'])
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto3_3_1 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto3_3_1_OutputFcn(hObject, eventdata, handles) 

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



function edit3_Callback(hObject, eventdata, handles)

function edit3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)

handles=guihandles;
guidata(hObject,handles);
TEIT_NUMBER=str2num(get(handles.edit1,'String'));
RICHTUNG_NUMBER=str2num(get(handles.edit3,'String'));
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
  return;
  elseif length(filename)~=TEIT_NUMBER*RICHTUNG_NUMBER*3
    msgbox('���������������������������������������');
    return;
else
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
end
 RESOLUTION_HE=500;
  RESOLUTION_WI=1300;
  zihao=20;
   if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
   end
  Fileadress=strcat(pathname,'result\');
  t2=waitbar(0,'�������ɱ���ͼƬ');
  
  for i=1:length(filename)
    MAX_WEG_INDEX(i)=find(MP{1,i}(:,2)==max(MP{1,i}(:,2)));  %�������������������������
    MAX_WEG(i)=MP{1,i}(MAX_WEG_INDEX(i),1);                         %������
  end
  for i=1:length(filename) 
  PLASTICHVERFORMUNG_INDEX=find(MP{1,i}(MAX_WEG_INDEX(i):end,2)<0,1);         %��Ѱ���ֵ�Ժ��һ����С��0�����������������Ա���
    if isempty(PLASTICHVERFORMUNG_INDEX)    
   PLASTICHVERFORMUNG(i)=MP{1,i}(length(MP{1,i}),1);                                                %���û��С��0�������Ա���Ϊ���ֵ
    else
       PLASTICHVERFORMUNG(i)=MP{1,i}(MAX_WEG_INDEX(i)+PLASTICHVERFORMUNG_INDEX-2,1);  %�����С��0�������Ա���Ϊ����0�����һ��ֵ
    end
  end
  
  for i=1:length(filename) 
  ELASTICHVERSOFRMUNG(i)=MAX_WEG(i)- PLASTICHVERFORMUNG(i) ;     %���Ա���
  end
  n=1;
for j=1:RICHTUNG_NUMBER
  for i=1:3
         TITLE_NAME{n}=['Richtung-',num2str(j),'  ',num2str(n),'#'];
         n=n+1;
      end
  end
  
  
    n=1;%ͼƬ��о
  %��ͼ
  
  for i=1:TEIT_NUMBER
    for j=1:RICHTUNG_NUMBER*3
        h=figure;
        set(h,'visible','off');
        plot(MP{1,n}(:,1),MP{1,n}(:,2),'linewidth',2);
        hold on;
        plot(MP{1,n}(MAX_WEG_INDEX(n),1),MP{1,n}(MAX_WEG_INDEX(n),2), 'o', 'markerfacecolor', [ 1, 0, 0 ]) %����
       
  PLASTICHVERFORMUNG_INDEX=find(MP{1,n}(MAX_WEG_INDEX(n):end,2)<0,1);         %��Ѱ���ֵ�Ժ��һ����С��0�����������������Ա���
    if isempty(PLASTICHVERFORMUNG_INDEX)    
    plot(MP{1,n}(length(MP{1,n}),1),MP{1,n}(length(MP{1,n}),2), 'o', 'markerfacecolor', [ 1, 0, 0 ]) %����                                           
    else
         plot(MP{1,n}(MAX_WEG_INDEX(n)+PLASTICHVERFORMUNG_INDEX-2,1),MP{1,n}(MAX_WEG_INDEX(n)+PLASTICHVERFORMUNG_INDEX-2,2), 'o', 'markerfacecolor', [ 1, 0, 0 ]) %����
    end
              text(MP{1,n}(MAX_WEG_INDEX(n),1)-0.1,MP{1,n}(MAX_WEG_INDEX(n),2)-300,['S_1=',num2str(MAX_WEG(n),'%.2f')],'FontSize',zihao);   
             text(MP{1,n}(length(MP{1,n}),1)*1.2,-100,['S_2=',num2str(PLASTICHVERFORMUNG(n),'%.2f')],'FontSize',zihao);   
             
        set(h,'position',[100,100,RESOLUTION_WI,RESOLUTION_HE]);
         set(h,'color','w')
        set(gca,'FontSize',zihao);
         title(TITLE_NAME{j},'FontSize',zihao);
         xlabel('Weg(mm)','FontSize',zihao);ylabel('Kraft(N)','FontSize',zihao);  
         Ym=max(MP{1,n}(:,2))*1.1;
          Xm=max(MP{1,n}(:,1))*1.1;
          grid on; set(gca, 'GridLineStyle' ,'-');axis([0 Xm 0 Ym]);
          hold off;
           sfilename1=[Fileadress,num2str(n),'.jpg'];
           f=getframe(h);
           imwrite(f.cdata,sfilename1);
           close(h);
         n=n+1;
         waitbar(n/(TEIT_NUMBER*RICHTUNG_NUMBER*3));
    end
  end
  close(t2);
  
  
    t3=waitbar(0,'��������Word����');
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%����Word����%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
filespec_user=[Fileadress,'report.doc'];
try 
Word=actxGetRunningServer('Word.Application');
catch 
Word=actxserver('Word.Application'); 
end
waitbar(0.1);
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

%===�ĵ���ҳ�߾�===========================================================
Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;

headline='III. Einzelergebnis ������';
Content.Start=0; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=10; % �����С
Content.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���
He=180*1.0771653543307086614173228346457;
Wi=240*1.9;
biaotihao=10;
waitbar(0.2);
Tab1 = Document.Tables.Add(Selection.Range,TEIT_NUMBER*RICHTUNG_NUMBER*3+2,7);
DTI = Document.Tables.Item(1); % �����
DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����

lc=28.381133333333333333333333333333; %���׻���
column_width = [1.75*lc,2.25*lc,1.75*lc,2*lc,3.49*lc,2.25*lc,2.25*lc];
for i = 1:7
DTI.Columns.Item(i).Width = column_width(i);
end
waitbar(0.4);
  DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
  DTI.Range.Font.NameAscii='Arial';
  DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
DTI.Cell(1,1).Range.Text = '�����';
DTI.Cell(2,1).Range.Text = 'Teil Nr.';
DTI.Cell(1,2).Range.Text = '���ط���';
DTI.Cell(2,2).Range.Text = 'Richtung';
DTI.Cell(1,3).Range.Text = '���';
DTI.Cell(2,3).Range.Text = 'Nr.';
DTI.Cell(1,4).Range.Text = '�غ�';
DTI.Cell(2,4).Range.Text = 'Belastung[N]';
DTI.Cell(1,5).Range.Text = '������S1';
DTI.Cell(1,5).Select;
Selection.Find.Text='1';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,5).Range.Text = 'Gesamtverformung[mm]';

DTI.Cell(1,6).Range.Text = '���Ա���S2';
DTI.Cell(1,6).Select;
Selection.Find.Text='2';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,6).Range.Text = 'plastische Verformung[mm]';

DTI.Cell(1,7).Range.Text = '���Ա���S3';
DTI.Cell(1,7).Select;
Selection.Find.Text='3';
Selection.Find.Execute;
Selection.Font.Subscript= true;
DTI.Cell(2,7).Range.Text = 'elastische Verformung[mm]';

n=3;
for i=1:TEIT_NUMBER
DTI.Cell(n,1).Merge(DTI.Cell(n+3*RICHTUNG_NUMBER-1,1)); 
n=n+3*RICHTUNG_NUMBER;
end


n=3;
for i=1:TEIT_NUMBER*RICHTUNG_NUMBER
DTI.Cell(n,2).Merge(DTI.Cell(n+2,2)); 
n=n+3;
end


DTI.Cell(3,4).Merge(DTI.Cell(TEIT_NUMBER*RICHTUNG_NUMBER*3+2,4)); 

for i=1:RICHTUNG_NUMBER
RICHTUNG_NAME{i}=['Richtung ',num2str(i)];
end

waitbar(0.6);
n=3;
 for i=1:TEIT_NUMBER
      for j=1:RICHTUNG_NUMBER
    DTI.Cell(n,2).Range.Text = RICHTUNG_NAME{j} ;
    n=n+3;
      end
 end


for i=1:RICHTUNG_NUMBER*3
           TITLE_NAME_TABLE{i}=[num2str(i),'#'];
  end
 
 
n=3;
 for i=1:TEIT_NUMBER
      for j=1:RICHTUNG_NUMBER*3
    DTI.Cell(n,3).Range.Text = TITLE_NAME_TABLE{j} ;
    n=n+1;
      end
 end
 
 for i=1:length(filename)
 DTI.Cell(i+2,5).Range.Text =num2str(MAX_WEG(i),'%.2f') ;
DTI.Cell(i+2,6).Range.Text =num2str(PLASTICHVERFORMUNG(i),'%.2f') ;
 DTI.Cell(i+2,7).Range.Text =num2str(ELASTICHVERSOFRMUNG(i),'%.2f') ;
 
 end
 Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;
InlineShapes=Document.InlineShapes;
waitbar(0.7);

n=1;
 for i=1:TEIT_NUMBER
 headline=['Teil Nummer  ',num2str(i)];
Selection.Text=headline;
Selection.Font.Size=10; % �����С
Selection.Font.NameAscii='Arial';

Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���
Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���


    for j=1:RICHTUNG_NUMBER*3
   sfilename1=[Fileadress,num2str(n),'.jpg'];
handle=Selection.InlineShapes.AddPicture(sfilename1);
Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���
delete(sfilename1); 
  n=n+1;

    end
 end
waitbar(0.9);
 Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
Word.Quit; % �ر��ĵ�
%%%%%%%%%%%%�������������Ϣ�������ռ�%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='��ǵ���֧������ն�';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
winopen([Fileadress,'report.doc']);
waitbar(1);
close(t3)
