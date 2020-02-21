function varargout = Auto2_4(varargin)


gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto2_4_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto2_4_OutputFcn, ...
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


% --- Executes just before Auto2_4 is made visible.
function Auto2_4_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')
b=load([cd,'\interface\Fahrzeugcode.mat']);
for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);
set(handles.popupmenu2,'Value',2);
FONTSIZE=10;
setappdata(0,'Auto2_4_FONTSIZE',FONTSIZE);

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto2_4 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto2_4_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on selection change in listbox1.
function listbox1_Callback(hObject, eventdata, handles)
cla(handles.axes1);
MP=getappdata(0,'Auto2_4_MP');
ZIHAO_TU_YULAN=getappdata(0,'Auto2_4_FONTSIZE');
%%%%%%%%%%%%%%˫����Figure%%%%%%%%%%%%%
%sel=get(gcf,'selectiontype'); %��ȡ��갴������
%if strcmp(sel,'open') %�Ƿ�˫�����
%pushbutton4_Callback(hObject, eventdata, handles)
%end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
if get(handles.checkbox1,'Value')==0

MIN_value=getappdata(0,'Auto2_4_MIN_value');
MIN_realindex=getappdata(0,'Auto2_4_MIN_realindex');
CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
j=CHOOSE;
format long
x=1:1:length(MP{j}(:,2));
x=x';
%plot(MP{j}(:,1),MP{j}(:,2),'color','b');
plot(handles.axes1,x,MP{j}(:,2),'color','b');

hold on
for i=1:length(MIN_realindex{j})
  plot(x(MIN_realindex{j}(i),1),MP{j}(MIN_realindex{j}(i),2),'ro','MarkerFaceColor','r','Markersize',8)
    text(x(MIN_realindex{j}(i),1)+100,MP{j}(MIN_realindex{j}(i),2),['(',num2str(MIN_value{j}(i),'%.1f'),')'],'FontSize',ZIHAO_TU_YULAN);
    %annotation('textarrow',[MP{j}(MIN_realindex{j}(i),1),MP{j}(MIN_realindex{j}(i),1)],[MP{j}(MIN_realindex{j}(i),2)*1.1,MP{j}(MIN_realindex{j}(i),2)],'String','ABC');
end

ylim(handles.axes1,[min(MIN_value{j})*1.1,20]);
%hold off;
datacursormode on;
grid on;
xlabel(handles.axes1,'Num/�����[s]','FontSize',ZIHAO_TU_YULAN);
ylabel(handles.axes1,'Kraft/��[N]','FontSize',ZIHAO_TU_YULAN);
for i=1:length(MIN_realindex{j})   
Pop5list{i,1}=i; 
end
set(handles.popupmenu5,'String',Pop5list)
%%%%%%%%%%%%%%%%%%%%%%%��װ��%%%%%%%%%%%%%%%%%%%%%%
else
MON_value=getappdata(0,'Auto2_4_MON_Value');
MON_realindex=getappdata(0,'Auto2_4_MON_realindex');
CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
j=CHOOSE;
format long
x=1:1:length(MP{j}(:,2));
x=x';
%plot(MP{j}(:,1),MP{j}(:,2),'color','b');
plot(handles.axes1,x,MP{j}(:,2),'color','b');
hold on
for i=1:length(MON_realindex{j})
  plot(x(MON_realindex{j}(i),1),MP{j}(MON_realindex{j}(i),2),'ro','MarkerFaceColor','r','Markersize',8)
    text(x(MON_realindex{j}(i),1)+100,MP{j}(MON_realindex{j}(i),2),['(',num2str(MON_value{j}(i),'%.1f'),')'],'FontSize',ZIHAO_TU_YULAN);
    
end

%hold off
datacursormode on
grid on;
xlabel(handles.axes1,'Num/�������','FontSize',ZIHAO_TU_YULAN)
ylabel(handles.axes1,'Kraft/��[N]','FontSize',ZIHAO_TU_YULAN)

for i=1:length(MON_realindex{j})   
Pop5list{i,1}=i; 
end
set(handles.popupmenu5,'String',Pop5list)
end

dcm_obj = datacursormode(gcf);
set(dcm_obj,'UpdateFcn',@NewCallback2_4)
try
popupmenu5_Callback(hObject, eventdata, handles)
end


function listbox1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton3.
function pushbutton3_Callback(hObject, eventdata, handles)
MP=getappdata(0,'Auto2_4_MP');
ZIHAO_TU_YULAN=getappdata(0,'Auto2_4_FONTSIZE')*2;
pathname=getappdata(0,'Auto2_4_pathname');
filename=getappdata(0,'Auto2_4_filename');
Fileadress=strcat(pathname,'result\');
   if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
   end
   
if get(handles.checkbox1,'Value')==0
MIN_value=getappdata(0,'Auto2_4_MIN_value');
MIN_realindex=getappdata(0,'Auto2_4_MIN_realindex');
t1=waitbar(0,'��������ͼƬ') ;  
   for j=1:length(MP)
       h=figure(j);
       set(h,'visible','off');
       set(h,'color','w')
       plot(MP{j}(:,1),MP{j}(:,2),'color','b');
       hold on
       for i=1:length(MIN_realindex{j})
            plot(MP{j}(MIN_realindex{j}(i),1),MP{j}(MIN_realindex{j}(i),2),'ro','MarkerFaceColor','r','Markersize',8)
             text(MP{j}(MIN_realindex{j}(i),1)-4,1.05*MP{j}(MIN_realindex{j}(i),2),['(',num2str(MIN_value{j}(i),'%.1f'),')'],'FontSize',ZIHAO_TU_YULAN);
       end
       set(h,'position',[100,100,1300,800]);
       hold off
       grid on;
       xlabel('Zeit/ʱ��[s]','FontSize',ZIHAO_TU_YULAN)
       ylabel('Kraft/��[N]','FontSize',ZIHAO_TU_YULAN)
       ylim([min(MIN_value{j})*1.1,20])
       set(gca,'FontSize',ZIHAO_TU_YULAN);
       sfilename=[Fileadress,filename{j},'Picture1' num2str(j) '.jpg'];
       f=getframe(h);
       imwrite(f.cdata,sfilename);
       waitbar(j/length(MP));
       close(h);
   end
   close(t1)
t2=waitbar(0,'�������ɱ���');
biaotihao=10;
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

Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;

headline='������';
Content.Start=0; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=biaotihao; % �����С
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���         
Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���
waitbar(0.1)

for i=1:length(MIN_value)
    Len_value(i)=length(MIN_value{i});
end
    TAB_COL=max(Len_value);
 Tab1 = Document.Tables.Add(Selection.Range, length(MP)+1,TAB_COL+1);
 DTI = Document.Tables.Item(1); % �����
 DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
 DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����
 lc=28.381133333333333333333333333333; %���׻���
 
 DTI.Columns.Item(1).Width =3*lc;
 for i = 2:TAB_COL+1
     DTI.Columns.Item(i).Width = 1.5*lc;
 end
 DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
 DTI.Range.Font.NameAscii='Arial';
 DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
 DTI.Cell(1,1).Range.Text ='��� Nr.';
 for i = 2:TAB_COL+1
     DTI.Cell(1,i).Range.Text =num2str(i-1);
 end
 for i=1:length(filename)
    a=find(filename{i}=='.');
    OUTFILENAME{i}=filename{i}(1:a-1);
end
  for i=2:length(MP)+1
      DTI.Cell(i,1).Range.Text =OUTFILENAME{i-1};
  end
  %%%%%%%%%%%%%%�������%%%%%%%%%
   for i = 1:length(MIN_value)
  MIN_value{i}=fliplr(MIN_value{i});
   end
  %%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
    for j=2:length(MP)+1        
  for i = 2:length(MIN_value{j-1})+1
       DTI.Cell(j,i).Range.Text =num2str(MIN_value{j-1}(i-1),'%.1f');
  end
    end
   waitbar(0.2) 
    Selection.Start = Content.end;
    Selection.TypeParagraph;
    InlineShapes=Document.InlineShapes;

for i=1:length(MP)  
    
    sfilename1=[Fileadress,filename{i},'Picture1' num2str(i) '.jpg'];
    handle=Selection.InlineShapes.AddPicture(sfilename1);
    %InlineShapes.Item(i).Height=He;
    %InlineShapes.Item(i).Width=Wi;
    % delete(sfilename1);
    Selection.Start = Content.end;
    Selection.TypeParagraph;
     headline=[OUTFILENAME{i},'-Demontage��ж'];
    Selection.Text=headline;
    Selection.Font.NameAscii='Arial';
    Selection.Font.Size=biaotihao; % �����С
    Selection.Start = Content.end;
    Selection.TypeParagraph;
    Selection.Start = Content.end;
    Selection.TypeParagraph;
    
    waitbar(0.2+0.8*i/length(MP));
end
   
   
   Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
   %%%%%%%%%%%%�������������Ϣ�������ռ�%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='��ˮ�۲�ж������';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
   
%%%%%%%%%%%%%%%%%%%%%%%%��װ��%%%%%%%%%%%%%
else
    t1=waitbar(0,'��������ͼƬ') ; 
for j=1:length(MP)
       h=figure(j);
       set(h,'visible','off');
       set(h,'color','w')   
MON_value=getappdata(0,'Auto2_4_MON_Value');
MON_realindex=getappdata(0,'Auto2_4_MON_realindex');


plot(MP{j}(:,1),MP{j}(:,2),'color','b');
hold on
for i=1:length(MON_realindex{j})
    plot(MP{j}(MON_realindex{j}(i),1),MP{j}(MON_realindex{j}(i),2),'ro','MarkerFaceColor','r','Markersize',8)
    text(MP{j}(MON_realindex{j}(i),1)+0.4,MP{j}(MON_realindex{j}(i),2),['(',num2str(MON_value{j}(i),'%.1f'),')'],'FontSize',ZIHAO_TU_YULAN*0.7);
    
end
  set(h,'position',[100,100,1300,800]);
       hold off
       grid on;
       xlabel('Zeit/ʱ��[s]','FontSize',ZIHAO_TU_YULAN)
       ylabel('Kraft/��[N]','FontSize',ZIHAO_TU_YULAN)
       set(gca,'FontSize',ZIHAO_TU_YULAN);
       sfilename=[Fileadress,filename{j},'Picture1' num2str(j) '.jpg'];
       f=getframe(h);
       imwrite(f.cdata,sfilename);
       waitbar(j/length(MP));
       close(h);
   end
   close(t1)
   t2=waitbar(0,'�������ɱ���');
biaotihao=10;
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

Document.PageSetup.TopMargin = 60*1.1745283018867924528301886792453;
Document.PageSetup.BottomMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.LeftMargin = 45*1.2641509433962264150943396226415;
Document.PageSetup.RightMargin = 45*0.94339622641509433962264150943396;

headline='������';
Content.Start=0; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=biaotihao; % �����С
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���         
Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���
waitbar(0.1)

for i=1:length(MON_value)
    Len_value(i)=length(MON_value{i});
end
    TAB_COL=max(Len_value);
 Tab1 = Document.Tables.Add(Selection.Range, length(MP)+1,TAB_COL+1);
 DTI = Document.Tables.Item(1); % �����
 DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
 DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����
 lc=28.381133333333333333333333333333; %���׻���
 
 DTI.Columns.Item(1).Width =3*lc;
 for i = 2:TAB_COL+1
     DTI.Columns.Item(i).Width = 1.5*lc;
 end
 DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
 DTI.Range.Font.NameAscii='Arial';
 DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
 DTI.Cell(1,1).Range.Text ='��� Nr.';
 for i = 2:TAB_COL+1
     DTI.Cell(1,i).Range.Text =num2str(i-1);
 end
 for i=1:length(filename)
    a=find(filename{i}=='.');
    OUTFILENAME{i}=filename{i}(1:a-1);
end
  for i=2:length(MP)+1
      DTI.Cell(i,1).Range.Text =OUTFILENAME{i-1};
  end

    for j=2:length(MP)+1        
  for i = 2:length(MON_value{j-1})+1
       DTI.Cell(j,i).Range.Text =num2str(MON_value{j-1}(i-1),'%.1f');
  end
    end
   waitbar(0.2) 
    Selection.Start = Content.end;
    Selection.TypeParagraph;
    InlineShapes=Document.InlineShapes;

for i=1:length(MP)  
    
    sfilename1=[Fileadress,filename{i},'Picture1' num2str(i) '.jpg'];
    handle=Selection.InlineShapes.AddPicture(sfilename1);
    %InlineShapes.Item(i).Height=He;
    %InlineShapes.Item(i).Width=Wi;
    % delete(sfilename1);
    Selection.Start = Content.end;
    Selection.TypeParagraph;
     headline=[OUTFILENAME{i},'-Montage��װ'];
    Selection.Text=headline;
    Selection.Font.NameAscii='Arial';
    Selection.Font.Size=biaotihao; % �����С
    Selection.Start = Content.end;
    Selection.TypeParagraph;
    Selection.Start = Content.end;
    Selection.TypeParagraph;
    
    waitbar(0.2+0.8*i/length(MP));
end
   
   
   Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
%%%%%%%%%%%%�������������Ϣ�������ռ�%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};
TEST_NAME='��ˮ�۰�װ������';
try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
end

%winopen([Fileadress,'report.doc']);
Word.Visible =1;
close(t2)

function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)

function popupmenu1_CreateFcn(hObject, eventdata, handles)

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
DATA_TYPE_KRAFT=get(handles.popupmenu2,'value');      %��ȡ���ݵڼ���Ϊ��
DATA_TYPE_WEG=get(handles.popupmenu1,'value');          %��ȡ���ݵڼ���Ϊλ��

  [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
    return;
    
else
    t1=waitbar(0,'���ڶ�������');
        for i=1:length(filename)
        Filename{i}=strcat(pathname,filename{i});
        [Type Sheet Format]=xlsfinfo(Filename{i}) ;
        sheet{i}=Sheet;
        MP_MITTLE{i}=xlsread(Filename{i},char(sheet{1,i}(1,1)));
        MP{i}(:,1)=MP_MITTLE{i}(:,DATA_TYPE_WEG);
        MP{i}(:,2)=MP_MITTLE{i}(:,DATA_TYPE_KRAFT);
        waitbar(i/length(filename));
        try
            system('taskkill/IM excel.exe');
        end
    end
end

close(t1);
if get(handles.checkbox1,'Value')==0

t2=waitbar(0,'���ڴ�������');
start_value=str2num(get(handles.edit1,'String'));                                     %���ڸ�����ʼ��ֵ 
end_value=-5;                                         %���ڸ�����ֹ��ֵ
for i=1:length(MP)
i_start{i}(1)=find(MP{1,i}(:,2)<start_value,1);   %��һ����ʼ��������
i_end{i}(1)=find(MP{1,i}(i_start{i}(1):end,2)>end_value,1);%��һ����ֹ��������
i_realstart{i}(1)=i_start{i}(1);                                         %��һ����ʼ����ʵ����
i_realend{i}(1)=i_start{i}(1)+i_end{i}(1)-1;                         %��һ����ֹ����ʵ����
end
for j=1:length(MP)
i=2;
while 1==1
    if isempty(find(MP{1,j}(i_realend{j}(i-1):end,2)<start_value))          %�����ж��Ƿ񻹴�����Сֵ���粻���������ѭ��
        break;
    else
        i_start{j}(i)=find(MP{1,j}(i_realend{j}(i-1):end,2)<start_value,1);        %��i����ʼ��������
        i_realstart{j}(i)=i_realend{j}(i-1)+i_start{j}(i)-1;                               %��i����ʼ����ʵ����
        i_end{j}(i)=find(MP{1,j}(i_realstart{j}(i):end,2)>end_value,1);             %��i����ֹ��������
        i_realend{j}(i)=i_realstart{j}(i)+i_end{j}(i)-1;                                    %��i����ֹ����ʵ����
    end
    i=i+1;
end

for i=1:length(i_realstart{j}) 
    MIN_value{j}(i)=min(MP{1,j}(i_realstart{j}(i):i_realend{j}(i),2));                                     %��i����Сֵ  
    MIN_index{j}(i)=find(MP{1,j}(i_realstart{j}(i):i_realend{j}(i),2)==MIN_value{j}(i),1);          %��i����Сֵ������
    MIN_realindex{j}(i)=i_realstart{j}(i)+MIN_index{j}(i)-1;                                        %��i����Сֵ��ʵ����
end
 waitbar(i/length(MP));
end

setappdata(0,'Auto2_4_MIN_value',MIN_value);
setappdata(0,'Auto2_4_MIN_realindex',MIN_realindex);
%%%%%%%%%%%%%%%%%%%��װ��%%%%%%%%%%%%%%%%%%%%%
else
 start_value=str2num(get(handles.edit1,'String'));                                     %���ڸ�����ʼ��ֵ 
 if start_value<0
     msgbox('��װ����ʼ������ӦΪ��ֵ');
     return;
 end
    t2=waitbar(0,'���ڴ�������');

end_value=10;                                         %���ڸ�����ֹ��ֵ
for i=1:length(MP)
i_start{i}(1)=find(MP{1,i}(:,2)>start_value,1);   %��һ����ʼ��������
i_end{i}(1)=find(MP{1,i}(i_start{i}(1):end,2)<end_value,1);%��һ����ֹ��������
i_realstart{i}(1)=i_start{i}(1);                                         %��һ����ʼ����ʵ����
i_realend{i}(1)=i_start{i}(1)+i_end{i}(1)-1;                         %��һ����ֹ����ʵ����
end

for j=1:length(MP)
i=2;
while 1==1
    if isempty(find(MP{1,j}(i_realend{j}(i-1):end,2)>start_value))          %�����ж��Ƿ񻹴�����Сֵ���粻���������ѭ��
        break;
    else
        i_start{j}(i)=find(MP{1,j}(i_realend{j}(i-1):end,2)>start_value,1);        %��i����ʼ��������
        i_realstart{j}(i)=i_realend{j}(i-1)+i_start{j}(i)-1;                               %��i����ʼ����ʵ����
        i_end{j}(i)=find(MP{1,j}(i_realstart{j}(i):end,2)<end_value,1);             %��i����ֹ��������
        i_realend{j}(i)=i_realstart{j}(i)+i_end{j}(i)-1;                                    %��i����ֹ����ʵ����
    end
    i=i+1;
end
shuaijian=str2num(get(handles.edit5,'String'));
for k=1:length(i_realstart{j})   
    a=MP{j}(i_realstart{j}(k):i_realend{j}(k),1:end);
    p=1;
    while 1==1
        if p==length(a)
            p=1;
            shuaijian=shuaijian-0.5;
        else
            if a(p+1,2)-a(p,2)<-shuaijian
                MON_Value{j}(k)=a(p,2);
                MON_realindex{j}(k)=i_realstart{j}(k)+p-1;
                break;
            end
            p=p+1;
        end
    end

end
 waitbar(i/length(MP));
end
setappdata(0,'Auto2_4_MON_Value', MON_Value);
setappdata(0,'Auto2_4_MON_realindex',MON_realindex);
end
setappdata(0,'Auto2_4_MP',MP);
setappdata(0,'Auto2_4_pathname',pathname);
setappdata(0,'Auto2_4_filename',filename);
set(handles.listbox1,'String',filename);
set(handles.listbox1,'Value',1);
set(handles.pushbutton3,'enable','on')
set(handles.pushbutton4,'enable','on')
set(handles.pushbutton5,'enable','on')
msgbox('���ݵ���ɹ�');
close(t2);


function edit1_Callback(hObject, eventdata, handles)

function edit1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)

if get(handles.checkbox1,'Value')==1
set(handles.edit5,'Enable','on');
set(handles.edit1,'String',30);
else
    set(handles.edit5,'Enable','off');
    set(handles.edit1,'String',-80);
end


% --- Executes on button press in pushbutton4.
function pushbutton4_Callback(hObject, eventdata, handles)
MP=getappdata(0,'Auto2_4_MP');
ZIHAO_TU_YULAN=getappdata(0,'Auto2_4_FONTSIZE');
CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
j=CHOOSE;

if get(handles.checkbox1,'Value')==0
MIN_value=getappdata(0,'Auto2_4_MIN_value');
MIN_realindex=getappdata(0,'Auto2_4_MIN_realindex');

h=figure(j);
format long
x=1:1:length(MP{j}(:,2));
x=x';
%plot(MP{j}(:,1),MP{j}(:,2),'color','b');
plot(x,MP{j}(:,2),'color','b');
hold on
for i=1:length(MIN_realindex{j})
    plot(x(MIN_realindex{j}(i),1),MP{j}(MIN_realindex{j}(i),2),'ro','MarkerFaceColor','r','Markersize',8)
    text(x(MIN_realindex{j}(i),1)+100,MP{j}(MIN_realindex{j}(i),2),['(',num2str(MIN_value{j}(i),'%.1f'),')'],'FontSize',ZIHAO_TU_YULAN);
    
end

hold off

xlabel(handles.axes1,'Time/ʱ��[s]','FontSize',ZIHAO_TU_YULAN)
ylabel(handles.axes1,'Kraft/��[N]','FontSize',ZIHAO_TU_YULAN)
dcm_obj = datacursormode(gcf);
set(dcm_obj,'UpdateFcn',@NewCallback2_4)

else
 MON_value=getappdata(0,'Auto2_4_MON_Value');
MON_realindex=getappdata(0,'Auto2_4_MON_realindex');

h=figure(j);
format long
x=1:1:length(MP{j}(:,2));
x=x';
%plot(MP{j}(:,1),MP{j}(:,2),'color','b');
plot(x,MP{j}(:,2),'color','b');
hold on
for i=1:length(MON_realindex{j})
    plot(x(MON_realindex{j}(i),1),MP{j}(MON_realindex{j}(i),2),'ro','MarkerFaceColor','r','Markersize',8)
    text(x(MON_realindex{j}(i),1)+100,MP{j}(MON_realindex{j}(i),2),['(',num2str(MON_value{j}(i),'%.1f'),')'],'FontSize',ZIHAO_TU_YULAN);
    
end

hold off

xlabel(handles.axes1,'Time/ʱ��[s]','FontSize',ZIHAO_TU_YULAN)
ylabel(handles.axes1,'Kraft/��[N]','FontSize',ZIHAO_TU_YULAN)
dcm_obj = datacursormode(gcf);
set(dcm_obj,'UpdateFcn',@NewCallback2_4)
end





function edit5_Callback(hObject, eventdata, handles)

function edit5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu5.
function popupmenu5_Callback(hObject, eventdata, handles)
MP=getappdata(0,'Auto2_4_MP');
if get(handles.checkbox1,'Value')==1
MON_value=getappdata(0,'Auto2_4_MON_Value');
MON_realindex=getappdata(0,'Auto2_4_MON_realindex');
CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
j=CHOOSE;
i=get(handles.popupmenu5,'Value');
format long
x=MON_realindex{j}(i);
set(handles.edit6,'String',num2str(x))

else
    MIN_value=getappdata(0,'Auto2_4_MIN_value');
    MIN_realindex=getappdata(0,'Auto2_4_MIN_realindex');
    CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
    j=CHOOSE;
    i=get(handles.popupmenu5,'Value');
    format long
    x=MIN_realindex{j}(i);
    set(handles.edit6,'String',num2str(x))
    
end




% --- Executes during object creation, after setting all properties.
function popupmenu5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit6_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function edit6_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton5.
function pushbutton5_Callback(hObject, eventdata, handles)
MP=getappdata(0,'Auto2_4_MP');
if get(handles.checkbox1,'Value')==1
MON_value=getappdata(0,'Auto2_4_MON_Value');
MON_realindex=getappdata(0,'Auto2_4_MON_realindex');
CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
j=CHOOSE;
i=get(handles.popupmenu5,'Value');

MON_realindex{j}(i)=str2num(get(handles.edit6,'String'));
MON_value{j}(i)=MP{j}(MON_realindex{j}(i),2);

setappdata(0,'Auto2_4_MON_Value', MON_value);
setappdata(0,'Auto2_4_MON_realindex',MON_realindex);

else
    MIN_value=getappdata(0,'Auto2_4_MIN_value');
    MIN_realindex=getappdata(0,'Auto2_4_MIN_realindex');
    CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
    j=CHOOSE;
    i=get(handles.popupmenu5,'Value');
    
    MIN_realindex{j}(i)=str2num(get(handles.edit6,'String'));
    MIN_value{j}(i)=MP{j}(MIN_realindex{j}(i),2);
    
    setappdata(0,'Auto2_4_MIN_value', MIN_value);
    setappdata(0,'Auto2_4_MIN_realindex',MIN_realindex);
    
end
listbox1_Callback(hObject, eventdata, handles)
msgbox('������޸ĳɹ�');
