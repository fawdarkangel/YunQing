function varargout = Auto7_3(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto7_3_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto7_3_OutputFcn, ...
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


% --- Executes just before Auto7_3 is made visible.
function Auto7_3_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
movegui(gcf,'center')

b=load([cd,'\interface\Fahrzeugcode.mat']);
%%%%%%%��ȡ�����ļ������ݳ�ȥ%%%%%%%%%
load([cd,'\model\Auto7_3_Version.mat']);
DATA_TYPE=CRUVE.DATA_TYPE;
X_LABLE=CRUVE.X_LABLE;
Y_LABLE=CRUVE.Y_LABLE;
TITLE_INDEX=CRUVE.TITLE_INDEX;
Y_LABLE_DECIMALDIGITS=CRUVE.Y_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
X_LABLE_DECIMALDIGITS=CRUVE.X_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
LINECOLOR=CRUVE.LINECOLOR;                      %��ȡ������ɫ
WIDTH=CRUVE.WIDTH;                                     %��ȡ�߿�
FONTSIZE=CRUVE.FONTSIZE;                            %��ȡ�ֺ�
MAXKRAFT_INSERT=CRUVE.MAXKRAFT_INSERT;
ASSISSTANT_LINE_CHECK=CRUVE.ASSISSTANT_LINE_CHECK;          %��ȡ�Ƿ���Ҫ������
ASSISSTANT_LINE_KRAFT=CRUVE.ASSISSTANT_LINE_KRAFT;           %��ȡ�����ߴ�С
ASSISSTANT_LINE_COLORINDEX=CRUVE.ASSISSTANT_LINE_COLORINDEX;      %��ȡ��������ɫ
GRID_WIDTH=CRUVE.GRID_WIDTH;
DATA_TYPE_KRAFT=CRUVE.DATA_TYPE_KRAFT;      %��ȡ���ݵڼ���Ϊ��
DATA_TYPE_WEG=CRUVE.DATA_TYPE_WEG;          %��ȡ���ݵڼ���Ϊλ��
DATASHEET=CRUVE.DATASHEET;                             %��ȡ����λ��Sheet��
setappdata(0,'CRUVE',CRUVE);
setappdata(0,'STAND_TITLE',STAND_TITLE);
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%

for i=1:length(b.Fahrzeugcode)
Fahrzeugcode{i,1}=b.Fahrzeugcode{i,2};
end
set(handles.Fahrzeugcode,'String',Fahrzeugcode);

handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto7_3 wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto7_3_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;





% --------------------------------------------------------------------
function Untitled_1_Callback(hObject, eventdata, handles)



% -------------------��������---------------------------
function Menu1_1_Callback(hObject, eventdata, handles)

CRUVE=getappdata(0,'CRUVE');
DATA_TYPE=CRUVE.DATA_TYPE;
TITLE_INDEX=CRUVE.TITLE_INDEX;

STAND_TITLE=getappdata(0,'STAND_TITLE');

switch DATA_TYPE
    case 1
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
    return;
    
else
    t1=waitbar(0,'���ڶ�������');
    DATA_TYPE_KRAFT=CRUVE.DATA_TYPE_KRAFT;      %��ȡ���ݵڼ���Ϊ��
    DATA_TYPE_WEG=CRUVE.DATA_TYPE_WEG;          %��ȡ���ݵڼ���Ϊλ��
    DATASHEET=CRUVE.DATASHEET;                   %��ȡλ��Sheet��
    for i=1:length(filename)
        Filename{i}=strcat(pathname,filename{i});
        [Type Sheet Format]=xlsfinfo(Filename{i}) ;
        sheet{i}=Sheet;
        MP_MITTLE{i}=xlsread(Filename{i},char(sheet{1,i}(1,DATASHEET)));
        MP{i}(:,1)=MP_MITTLE{i}(:,DATA_TYPE_WEG);
        MP{i}(:,2)=MP_MITTLE{i}(:,DATA_TYPE_KRAFT);
        waitbar(i/length(filename));
        try
            system('taskkill/IM excel.exe');
        end
    end
end
    case 2
        [filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������','MultiSelect','on');
        if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
            msgbox('�����ļ�ʧ��');
            return;
        end            
        t1=waitbar(0,'���ڶ�������');
        DATA_TYPE_KRAFT=CRUVE.DATA_TYPE_KRAFT;      %��ȡ���ݵڼ���Ϊ��
        DATA_TYPE_WEG=CRUVE.DATA_TYPE_WEG;          %��ȡ���ݵڼ���Ϊλ��
        for i=1:length(filename)
            Filename{i}=strcat(pathname,filename{i});
            fidin=fopen(Filename{i});                               % ��test2.txt�ļ�
            fidout=fopen('result.txt','w');                       % ����MKMATLAB.txt�ļ�
            tline=fgetl(fidin);
            tline=fgetl(fidin);
            while ~feof(fidin)                                      % �ж��Ƿ�Ϊ�ļ�ĩβ
                tline=fgetl(fidin);                                     % ���ļ�����
                if isempty(tline)
                    tline=fgetl(fidin);
                end
                if double(tline(1))>=48&&double(tline(1))<=57       % �ж����ַ��Ƿ�����ֵ
                    fprintf(fidout,'%s\n\n',tline);                  % ����������У��Ѵ�������д���ļ�MKMATLAB.txt
                    continue                                         % ����Ƿ����ּ�����һ��ѭ��
                end
            end
            fclose(fidout);
            MK=importdata('result.txt');
            MP{i}(:,1)=MK(:,DATA_TYPE_WEG);
            MP{i}(:,2)=MK(:,DATA_TYPE_KRAFT);
%             try
%             delete('result.txt');
%             end
            waitbar(i/length(filename));
        end
end
if TITLE_INDEX==8
    for i=1:length(filename)
            n(i,1)=find('.'==filename{1,i});
            STAND_TITLE{i}=filename{1,i}(1:n(i,1)-1);
    end  
end
    close(t1);    
    setappdata(0,'MP',MP);
    setappdata(0,'filename',filename);
    setappdata(0,'pathname',pathname);
    setappdata(0,'Filename',Filename);
    setappdata(0,'STAND_TITLE',STAND_TITLE);
    set(handles.listbox1,'Value',1);
    set(handles.listbox1,'String',filename);
    msgbox('���ݵ���ɹ�');



% --- ͼ��Ԥ��.
function listbox1_Callback(hObject, eventdata, handles)

cla(handles.axes1);

MP=getappdata(0,'MP');
CRUVE=getappdata(0,'CRUVE');
STAND_TITLE=getappdata(0,'STAND_TITLE');
DATA_TYPE=CRUVE.DATA_TYPE;
X_LABLE=CRUVE.X_LABLE;
Y_LABLE=CRUVE.Y_LABLE;
Y_LABLE_DECIMALDIGITS=CRUVE.Y_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
X_LABLE_DECIMALDIGITS=CRUVE.X_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
LINECOLOR=CRUVE.LINECOLOR;                      %��ȡ������ɫ
WIDTH=CRUVE.WIDTH;                                     %��ȡ�߿�
FONTSIZE=CRUVE.FONTSIZE;                            %��ȡ�ֺ�
WIDTH=str2num(WIDTH);                                   %�ַ���ת����
FONTSIZE=str2num(FONTSIZE);                          %�ַ���ת����
MAXKRAFT_INSERT=CRUVE.MAXKRAFT_INSERT;  %��ȡ�Ƿ��ע���ֵ
ASSISSTANT_LINE_CHECK=CRUVE.ASSISSTANT_LINE_CHECK;          %��ȡ�Ƿ���Ҫ������
ASSISSTANT_LINE_KRAFT=CRUVE.ASSISSTANT_LINE_KRAFT;           %��ȡ�����ߴ�С
ASSISSTANT_LINE_COLORINDEX=CRUVE.ASSISSTANT_LINE_COLORINDEX;      %��ȡ��������ɫ
GRID_WIDTH=CRUVE.GRID_WIDTH;                    %��ȡ�����ܶ�
TITLEFONTSIZE=str2num(CRUVE.TITLEFONTSIZE);                %��ȡ�����ֺ�
BASE_LINECOLOR={'b','k','r','g','y','m',[255,165,0]/255};


CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
i=CHOOSE;
ZIHAO_TU_YULAN=FONTSIZE/2;
TITLEFONTSIZE=TITLEFONTSIZE/2;
plot(handles.axes1,MP{i}(:,1),MP{i}(:,2),'color',BASE_LINECOLOR{LINECOLOR},'linewidth',WIDTH);
datacursormode on
Y_max=max(MP{1,i}(:,2));
X_max=max(MP{1,i}(:,1));
if ASSISSTANT_LINE_CHECK==1
    hold on
    plot(handles.axes1,[0,X_max*1.03],[str2num(ASSISSTANT_LINE_KRAFT),str2num(ASSISSTANT_LINE_KRAFT)],'color',BASE_LINECOLOR{ASSISSTANT_LINE_COLORINDEX},'linewidth',WIDTH)
    hold off
end

grid on;
if GRID_WIDTH==1
    grid minor;
end
xlabel(handles.axes1,X_LABLE,'FontSize',ZIHAO_TU_YULAN)
ylabel(handles.axes1,Y_LABLE,'FontSize',ZIHAO_TU_YULAN)
title(handles.axes1,STAND_TITLE{i},'FontSize',TITLEFONTSIZE)


if MAXKRAFT_INSERT==1
    MAXKRAFT_INDEX(i)=find(MP{1,i}(:,2)==Y_max,1);
     text(handles.axes1,MP{1,i}(MAXKRAFT_INDEX(i),1),MP{1,i}(MAXKRAFT_INDEX(i),2),['\leftarrow(',num2str(Y_max,['%.',num2str(Y_LABLE_DECIMALDIGITS-1),'f']),'N)'],'FontSize',ZIHAO_TU_YULAN);
     axis(handles.axes1,[0 X_max*1.1 0 Y_max*1.1]);
else
    axis(handles.axes1,[0 X_max*1.05 0 Y_max*1.1]);
end
set(handles.edit4,'String',[num2str(Y_max,['%.',num2str(Y_LABLE_DECIMALDIGITS-1),'f']),'N']);
set(handles.edit5,'String',[num2str(X_max,['%.',num2str(X_LABLE_DECIMALDIGITS-1),'f']),'mm']);
set(handles.edit6,'String',[num2str(MP{i}(end,1),['%.',num2str(X_LABLE_DECIMALDIGITS-1),'f']),'mm']);

function listbox1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in Fahrzeugcode.
function Fahrzeugcode_Callback(hObject, eventdata, handles)

function Fahrzeugcode_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton4.
function pushbutton4_Callback(hObject, eventdata, handles)
MP=getappdata(0,'MP');
if isempty(MP)
    msgbox('���ȵ�������');
    return
end
pathname=getappdata(0,'pathname');
Filename=getappdata(0,'Filename');
CRUVE=getappdata(0,'CRUVE');
DATA_TYPE=CRUVE.DATA_TYPE;
STAND_TITLE=getappdata(0,'STAND_TITLE');
X_LABLE=CRUVE.X_LABLE;
Y_LABLE=CRUVE.Y_LABLE;
WORD_PICTURE_TYPE=get(handles.popupmenu5,'Value');
CHECK_KRAFT_WEG=get(handles.checkbox1,'Value');
Y_LABLE_DECIMALDIGITS=CRUVE.Y_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
X_LABLE_DECIMALDIGITS=CRUVE.X_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
LINECOLOR=CRUVE.LINECOLOR;                      %��ȡ������ɫ
WIDTH=CRUVE.WIDTH;                                     %��ȡ�߿�
FONTSIZE=CRUVE.FONTSIZE;                            %��ȡ�ֺ�
WIDTH=str2num(WIDTH);                                   %�ַ���ת����
FONTSIZE=str2num(FONTSIZE);                          %�ַ���ת����
MAXKRAFT_INSERT=CRUVE.MAXKRAFT_INSERT;  %��ȡ�Ƿ��ע���ֵ
ASSISSTANT_LINE_CHECK=CRUVE.ASSISSTANT_LINE_CHECK;          %��ȡ�Ƿ���Ҫ������
ASSISSTANT_LINE_KRAFT=CRUVE.ASSISSTANT_LINE_KRAFT;           %��ȡ�����ߴ�С
ASSISSTANT_LINE_COLORINDEX=CRUVE.ASSISSTANT_LINE_COLORINDEX;      %��ȡ��������ɫ
GRID_WIDTH=CRUVE.GRID_WIDTH;                    %��ȡ�����ܶ�
TITLEFONTSIZE=str2num(CRUVE.TITLEFONTSIZE);                %��ȡ�����ֺ�
BASE_LINECOLOR={'b','k','r','g','y','m',[255,165,0]/255};

TEST_NAME=get(handles.edit3,'String');
if isempty(TEST_NAME)
    msgbox('��������������')
    return
end
ZIHAO_TU=FONTSIZE;
Fileadress=strcat(pathname,'result\');
if ~exist('pathname\result','dir')
      mkdir(pathname,'result');
end
t2=waitbar(0,'��������ͼƬ');
for i=1:length(MP)
    h=figure;
    set(h,'visible','off')
    
    switch WORD_PICTURE_TYPE
        case 1
           set(h,'position',[100,100,1300,800]); 
        case 2
            set(h,'position',[100,100,1300,1300]); 
    end
    plot(MP{i}(:,1),MP{i}(:,2),'color',BASE_LINECOLOR{LINECOLOR},'linewidth',WIDTH);
    Y_max(i)=max(MP{1,i}(:,2));
    X_max(i)=max(MP{1,i}(:,1));
    CANYU(i)=MP{1,i}(end,1);
    if ASSISSTANT_LINE_CHECK==1
        hold on
        plot([0,X_max(i)*1.03],[str2num(ASSISSTANT_LINE_KRAFT),str2num(ASSISSTANT_LINE_KRAFT)],'color',BASE_LINECOLOR{ASSISSTANT_LINE_COLORINDEX},'linewidth',WIDTH)
        hold off
    end
    grid on
    if GRID_WIDTH==1
        grid minor;
    end
    set(gca,'Fontsize',ZIHAO_TU);
    xlabel(X_LABLE,'FontSize',ZIHAO_TU)
    ylabel(Y_LABLE,'FontSize',ZIHAO_TU)
    title(STAND_TITLE{i},'FontSize',ZIHAO_TU)
  
    
    if MAXKRAFT_INSERT==1
        MAXKRAFT_INDEX(i)=find(MP{1,i}(:,2)==Y_max(i),1);
        text(MP{1,i}(MAXKRAFT_INDEX(i),1),MP{1,i}(MAXKRAFT_INDEX(i),2),['\leftarrow(',num2str(Y_max(i),['%.',num2str(Y_LABLE_DECIMALDIGITS-1),'f']),'N)'],'FontSize',ZIHAO_TU);
        axis([0 X_max(i)*1.1 0 Y_max(i)*1.1]);
    else
        axis([0 X_max(i)*1.05 0 Y_max(i)*1.1]);
    end
    
    sfilename1=[Fileadress,num2str(i),'.jpg'];
    saveas(h,sfilename1);
   close(h);
   waitbar(i/length(MP));      
end
close(t2)

t3=waitbar(0,'��������Word����') ;  
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
headline='III.1 ������';
Content.Start=0; % ��ʼ��Ϊ0������ʾÿ��д�븲��֮ǰ����
Content.Text=headline;
Content.Font.Size=biaotihao; % �����С
Content.Font.NameAscii='Arial';
Selection.Start = Content.end; 
Selection.TypeParagraph;% ����һ���µĿն���         
Selection.Start = Selection.end; 
Selection.TypeParagraph;% ����һ���µĿն���

if CHECK_KRAFT_WEG==1
    Tab1 = Document.Tables.Add(Selection.Range, length(MP)+1,4);
    DTI = Document.Tables.Item(1); % �����
    DTI.Borders.OutsideLineStyle = 'wdLineStyleSingle'; % �����ʵ��
    DTI.Borders.InsideLineStyle = 'wdLineStyleSingle'; % ���е��ڿ�����
    lc=28.381133333333333333333333333333; %���׻���
    column_width = [1*lc,3*lc,3*lc,3*lc];
    for i = 1:4
        DTI.Columns.Item(i).Width = column_width(i);
    end
    DTI.Range.Paragraphs.Alignment='wdAlignParagraphCenter';
    DTI.Range.Font.NameAscii='Arial';
    DTI.Range.Cells.VerticalAlignment = 'wdCellAlignVerticalCenter';
    DTI.Cell(1,1).Range.Text ='���';
    DTI.Cell(1,2).Range.Text ='�����(N)';
    DTI.Cell(1,3).Range.Text ='������(mm)';
    DTI.Cell(1,4).Range.Text ='�������(mm)';
    for i=1:length(MP)
        DTI.Cell(i+1,1).Range.Text =num2str(i);
        DTI.Cell(i+1,2).Range.Text =num2str(Y_max(i),['%.',num2str(Y_LABLE_DECIMALDIGITS-1),'f']);
        DTI.Cell(i+1,3).Range.Text =num2str(X_max(i),['%.',num2str(X_LABLE_DECIMALDIGITS-1),'f']);
        DTI.Cell(i+1,4).Range.Text =num2str(CANYU(i),['%.',num2str(X_LABLE_DECIMALDIGITS-1),'f']);        
    end
end
 Selection.Start = Content.end;
Selection.TypeParagraph;
Selection.Start = Selection.end;
Selection.TypeParagraph;

InlineShapes=Document.InlineShapes;
switch WORD_PICTURE_TYPE
    case 1
        for i=1:length(MP)
            sfilename1=[Fileadress,num2str(i),'.jpg'];
            handle=Selection.InlineShapes.AddPicture(sfilename1);
            delete(sfilename1);
            waitbar(i/length(MP));
        end
    case 2
        He=180*0.94488188976377952755905511811024;
        Wi=240;
        for i=1:length(MP)
            sfilename1=[Fileadress,num2str(i),'.jpg'];
            handle=Selection.InlineShapes.AddPicture(sfilename1);
            InlineShapes.Item(i).Height=He;
            InlineShapes.Item(i).Width=Wi;
            if mod(i,2)==0
                Selection.Start = Selection.end;
                Selection.TypeParagraph;
            end
            delete(sfilename1);
            waitbar(i/length(MP));
        end
end

Document.ActiveWindow.ActivePane.View.Type = 'wdPrintView';
Document.Save; % �����ĵ�
%Word.Quit; % �ر��ĵ�
%%%%%%%%%%%%�������������Ϣ�������ռ�%%%%%%%%%%%%%%%
FAHRZEUGCODE_list=get(handles.Fahrzeugcode,'String');
FAHRZEUGCODE_val=get(handles.Fahrzeugcode,'Value');
FAHRZEUGCODE=FAHRZEUGCODE_list{FAHRZEUGCODE_val};

try
REPORTINFORMATION_OUTPUT(FAHRZEUGCODE,TEST_NAME);
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%winopen([Fileadress,'report.doc']);
Word.Visible =1;
close(t3);
function popupmenu5_Callback(hObject, eventdata, handles)

function popupmenu5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit3_Callback(hObject, eventdata, handles)

function edit3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit4_Callback(hObject, eventdata, handles)

function edit4_CreateFcn(hObject, eventdata, handles)


if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit5_Callback(hObject, eventdata, handles)

function edit5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)


% --------------------------------------------------------------------
function Menu2_Callback(hObject, eventdata, handles)


% --------------------------------------------------------------------
function Menu2_1_Callback(hObject, eventdata, handles)
run Auto7_3_Configuration



function edit6_Callback(hObject, eventdata, handles)

function edit6_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton5.
function pushbutton5_Callback(hObject, eventdata, handles)

MP=getappdata(0,'MP');
CRUVE=getappdata(0,'CRUVE');
STAND_TITLE=getappdata(0,'STAND_TITLE');
DATA_TYPE=CRUVE.DATA_TYPE;
X_LABLE=CRUVE.X_LABLE;
Y_LABLE=CRUVE.Y_LABLE;
Y_LABLE_DECIMALDIGITS=CRUVE.Y_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
X_LABLE_DECIMALDIGITS=CRUVE.X_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
LINECOLOR=CRUVE.LINECOLOR;                      %��ȡ������ɫ
WIDTH=CRUVE.WIDTH;                                     %��ȡ�߿�
FONTSIZE=CRUVE.FONTSIZE;                            %��ȡ�ֺ�
WIDTH=str2num(WIDTH);                                   %�ַ���ת����
FONTSIZE=str2num(FONTSIZE);                          %�ַ���ת����
MAXKRAFT_INSERT=CRUVE.MAXKRAFT_INSERT;  %��ȡ�Ƿ��ע���ֵ
ASSISSTANT_LINE_CHECK=CRUVE.ASSISSTANT_LINE_CHECK;          %��ȡ�Ƿ���Ҫ������
ASSISSTANT_LINE_KRAFT=CRUVE.ASSISSTANT_LINE_KRAFT;           %��ȡ�����ߴ�С
ASSISSTANT_LINE_COLORINDEX=CRUVE.ASSISSTANT_LINE_COLORINDEX;      %��ȡ��������ɫ
GRID_WIDTH=CRUVE.GRID_WIDTH;                    %��ȡ�����ܶ�
TITLEFONTSIZE=str2num(CRUVE.TITLEFONTSIZE);                %��ȡ�����ֺ�
BASE_LINECOLOR={'b','k','r','g','y','m',[255,165,0]/255};


CHOOSE=get(handles.listbox1,'Value');                %listbox��ֵ
i=CHOOSE;
ZIHAO_TU_YULAN=FONTSIZE/2;
TITLEFONTSIZE=TITLEFONTSIZE/2;
h=figure(i);
plot(MP{i}(:,1),MP{i}(:,2),'color',BASE_LINECOLOR{LINECOLOR},'linewidth',WIDTH);

Y_max=max(MP{1,i}(:,2));
X_max=max(MP{1,i}(:,1));
if ASSISSTANT_LINE_CHECK==1
    hold on
    plot([0,X_max*1.03],[str2num(ASSISSTANT_LINE_KRAFT),str2num(ASSISSTANT_LINE_KRAFT)],'color',BASE_LINECOLOR{ASSISSTANT_LINE_COLORINDEX},'linewidth',WIDTH)
    hold off
end

grid on;
if GRID_WIDTH==1
    grid minor;
end
xlabel(X_LABLE,'FontSize',ZIHAO_TU_YULAN)
ylabel(Y_LABLE,'FontSize',ZIHAO_TU_YULAN)
title(STAND_TITLE{i},'FontSize',TITLEFONTSIZE)


if MAXKRAFT_INSERT==1
    MAXKRAFT_INDEX(i)=find(MP{1,i}(:,2)==Y_max,1);
     text(MP{1,i}(MAXKRAFT_INDEX(i),1),MP{1,i}(MAXKRAFT_INDEX(i),2),['\leftarrow(',num2str(Y_max,['%.',num2str(Y_LABLE_DECIMALDIGITS-1),'f']),'N)'],'FontSize',ZIHAO_TU_YULAN);
     axis([0 X_max*1.1 0 Y_max*1.1]);
else
    axis([0 X_max*1.05 0 Y_max*1.1]);
end
