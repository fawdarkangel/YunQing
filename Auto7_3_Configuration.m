function varargout = Auto7_3_Configuration(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto7_3_Configuration_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto7_3_Configuration_OutputFcn, ...
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


% --- Executes just before Auto7_3_Configuration is made visible.
function Auto7_3_Configuration_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
load([cd,'\model\Auto7_3_Version.mat'])            %��ȡ�����ļ�
DATA_TYPE=CRUVE.DATA_TYPE;                       %��ȡ��������
X_LABLE=CRUVE.X_LABLE;                                 %��ȡ������
Y_LABLE=CRUVE.Y_LABLE;                                 %��ȡ������
TITLE_INDEX=CRUVE.TITLE_INDEX;                    %��ȡͼƬ����
Y_LABLE_DECIMALDIGITS=CRUVE.Y_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
X_LABLE_DECIMALDIGITS=CRUVE.X_LABLE_DECIMALDIGITS; %��ȡ������С��λ��
LINECOLOR=CRUVE.LINECOLOR;                      %��ȡ������ɫ
WIDTH=CRUVE.WIDTH;                                     %��ȡ�߿�
FONTSIZE=CRUVE.FONTSIZE;                            %��ȡ�ֺ�
MAXKRAFT_INSERT= CRUVE.MAXKRAFT_INSERT;   %��ȡ�Ƿ���ͼ�б�ע���ֵ��1����ע��0������ע
ASSISSTANT_LINE_CHECK=CRUVE.ASSISSTANT_LINE_CHECK;          %��ȡ�Ƿ���Ҫ������
ASSISSTANT_LINE_KRAFT=CRUVE.ASSISSTANT_LINE_KRAFT;           %��ȡ�����ߴ�С
ASSISSTANT_LINE_COLORINDEX=CRUVE.ASSISSTANT_LINE_COLORINDEX;      %��ȡ��������ɫ
GRID_WIDTH=CRUVE.GRID_WIDTH;                     %��ȡ�����ܶ� 0�������� 1������
DATA_TYPE_KRAFT=CRUVE.DATA_TYPE_KRAFT;      %��ȡ���ݵڼ���Ϊ��
DATA_TYPE_WEG=CRUVE.DATA_TYPE_WEG;          %��ȡ���ݵڼ���Ϊλ��
TITLEFONTSIZ=CRUVE.TITLEFONTSIZE;                 %��ȡ�����ֺ�
DATASHEET=CRUVE.DATASHEET;                             %��ȡ����λ��Sheet��

%���ΪZwick����popupmenu9���ڼ���SheetΪ����
if DATA_TYPE==1
    set(handles.popupmenu9,'Enable','on');
else
    set(handles.popupmenu9,'Enable','off');
end

if TITLE_INDEX==7                                             %��������ϴ����Զ���Ļ�����ʼ����ʱ���ͷ��Զ��尴ť
 set(handles.pushbutton2,'Enable','on');
else
  set(handles.pushbutton2,'Enable','off');
end

setappdata(0,'CRUVE',CRUVE);                             %������д���ڴ�
setappdata(0,'STAND_TITLE',STAND_TITLE);        %������д���ڴ�

set(handles.edit1,'String',X_LABLE);                       %д������������
set(handles.edit2,'String',Y_LABLE);                       %д�������������
set(handles.popupmenu2,'Value',DATA_TYPE);       %������������ѡ���
set(handles.popupmenu1,'Value',TITLE_INDEX);     %����ͼ�����ѡ���
set(handles.popupmenu3,'Value',Y_LABLE_DECIMALDIGITS);     %����������С��λ��
set(handles.popupmenu4,'Value',X_LABLE_DECIMALDIGITS);     %����������С��λ��
set(handles.popupmenu5,'Value',LINECOLOR);       %����������ɫ
set(handles.edit3,'String',WIDTH);                           %д���߿�
set(handles.edit4,'String',FONTSIZE);                       %д���ֺ�
set(handles.checkbox1,'Value',MAXKRAFT_INSERT)  %д���Ƿ��ע�����ֵ
set(handles.checkbox2,'Value',ASSISSTANT_LINE_CHECK);
set(handles.checkbox3,'Value',GRID_WIDTH);           %д�������ܶȸ�ѡ��
set(handles.popupmenu7,'Value',DATA_TYPE_KRAFT);       %д�����ݵڼ���Ϊ��
set(handles.popupmenu8,'Value',DATA_TYPE_WEG);       %д�����ݵڼ���Ϊλ��
set(handles.edit6,'String',TITLEFONTSIZ);       %д�����ݵڼ���Ϊλ��
set(handles.popupmenu9,'Value',DATASHEET);   %д��ڼ���SheetΪ����


if ASSISSTANT_LINE_CHECK==1                                        %��Ҫ��������
    set(handles.edit5,'Enable','on');
    set(handles.popupmenu6,'Enable','on');    
    set(handles.edit5,'String',ASSISSTANT_LINE_KRAFT);
    set(handles.popupmenu6,'Value',ASSISSTANT_LINE_COLORINDEX);
else                                                                                       %����Ҫ��������
    set(handles.edit5,'Enable','off');
    set(handles.popupmenu6,'Enable','off');
    set(handles.edit5,'String',ASSISSTANT_LINE_KRAFT);
    set(handles.popupmenu6,'Value',ASSISSTANT_LINE_COLORINDEX);
end



handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto7_3_Configuration wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto7_3_Configuration_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- �������ð�ť.
function pushbutton1_Callback(hObject, eventdata, handles)

CRUVE=getappdata(0,'CRUVE');                             %��ȡ�ڴ��������ļ�
STAND_TITLE=getappdata(0,'STAND_TITLE');           %��ȡ�ڴ��б����ļ�
DATA_TYPE=get(handles.popupmenu2,'Value');          %��ȡ��������
TITLE_INDEX=get(handles.popupmenu1,'Value');         %��ȡͼ���������

X_LABLE=get(handles.edit1,'String');                                       %��ȡ���������
Y_LABLE=get(handles.edit2,'String');                                       %��ȡ���������
Y_LABLE_DECIMALDIGITS=get(handles.popupmenu3,'Value'); %��ȡ������С��λ��
X_LABLE_DECIMALDIGITS=get(handles.popupmenu4,'Value'); %��ȡ������С��λ��
LINECOLOR=get(handles.popupmenu5,'Value');                     %��ȡ������ɫ
WIDTH=get(handles.edit3,'String');                                         %��ȡ�߿�
FONTSIZE=get(handles.edit4,'String');                                     %��ȡ�ֺ�
MAXKRAFT_INSERT= get(handles.checkbox1,'Value');   %��ȡ�Ƿ���ͼ�б�ע���ֵ��1����ע��0������ע
ASSISSTANT_LINE_CHECK=get(handles.checkbox2,'Value');  %��ȡ�Ƿ���Ҫ������ 1����Ҫ  0������Ҫ 
ASSISSTANT_LINE_KRAFT=get(handles.edit5,'String');           %��ȡ��������ֵ
ASSISSTANT_LINE_COLORINDEX=get(handles.popupmenu6,'Value');  %��ȡ��������ɫ���
GRID_WIDTH=get(handles.checkbox3,'Value');  %��ȡ�������ܶ� 1������  0��������; 
DATA_TYPE_KRAFT=get(handles.popupmenu7,'Value');    %��ȡ�ڼ���Ϊ��
DATA_TYPE_WEG=get(handles.popupmenu8,'Value');     %��ȡ�ڼ���Ϊλ��
TITLEFONTSIZE=get(handles.edit6,'String');       %��ȡ�����ֺ�
DATASHEET=get(handles.popupmenu9,'Value');   %д��ڼ���SheetΪ����

%%%%%%%%%%%ѡ��ͼ�����%%%%%%%%%%%%%%%
switch TITLE_INDEX
    case 1
         for i=1:1000
            STAND_TITLE{i}=[' '];
        end
    case 2
        for i=1:1000
            STAND_TITLE{i}=['MP ',num2str(i)];
        end
    case 3
        for i=1:1000
            STAND_TITLE{i}=['MP ',num2str(i),'#'];
        end
    case 4
        for i=1:1000
            STAND_TITLE{i}=['Teil ',num2str(i)];
        end
    case 5
        for i=1:1000
            STAND_TITLE{i}=['Teil ',num2str(i),'#'];
        end
    case 6
        for i=1:1000
            STAND_TITLE{i}=[num2str(i),'#'];
        end
    case 7        
        set(handles.pushbutton1,'Enable','on');
        STAND_TITLE=getappdata(0,'STAND_TITLE');
    case 8
        filename=getappdata(0,'filename');
        if isempty(filename)
            msgbox('δ��⵽�ļ��������ȵ�������');
            return
        end
        for i=1:length(filename)
            n(i,1)=find('.'==filename{1,i});
            STAND_TITLE{i}=filename{1,i}(1:n(i,1)-1);
        end        
end
%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%
%%%%%%%%%%%%���������ļ�%%%%%%%%%%%%%%%%%%
CRUVE.DATA_TYPE=DATA_TYPE;
CRUVE.X_LABLE=X_LABLE;
CRUVE.Y_LABLE=Y_LABLE;
CRUVE.TITLE_INDEX=TITLE_INDEX;
CRUVE.Y_LABLE_DECIMALDIGITS=Y_LABLE_DECIMALDIGITS;
CRUVE.X_LABLE_DECIMALDIGITS=X_LABLE_DECIMALDIGITS;
CRUVE.LINECOLOR=LINECOLOR;
CRUVE.WIDTH=WIDTH;
CRUVE.FONTSIZE=FONTSIZE;
CRUVE.MAXKRAFT_INSERT=MAXKRAFT_INSERT;
CRUVE.ASSISSTANT_LINE_CHECK=ASSISSTANT_LINE_CHECK;        
CRUVE.ASSISSTANT_LINE_KRAFT=ASSISSTANT_LINE_KRAFT;        
CRUVE.ASSISSTANT_LINE_COLORINDEX=ASSISSTANT_LINE_COLORINDEX;     
CRUVE.GRID_WIDTH=GRID_WIDTH;
CRUVE.DATA_TYPE_KRAFT=DATA_TYPE_KRAFT;
CRUVE.DATA_TYPE_WEG=DATA_TYPE_WEG;
CRUVE.TITLEFONTSIZE=TITLEFONTSIZE;
CRUVE.DATASHEET=DATASHEET;
setappdata(0,'CRUVE',CRUVE);
setappdata(0,'STAND_TITLE',STAND_TITLE);
if TITLE_INDEX==7&&isempty(STAND_TITLE)
    msgbox('�뵼���Զ������EXCEL')
    return
end

save([cd,'\model\Auto7_3_Version.mat'],'CRUVE','STAND_TITLE')
msgbox('���ñ���ɹ�');

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


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)
TITLE_INDEX=get(handles.popupmenu1,'Value');
if TITLE_INDEX==7
 set(handles.pushbutton2,'Enable','on');
else
  set(handles.pushbutton2,'Enable','off');
end


function popupmenu1_CreateFcn(hObject, eventdata, handles)





if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in pushbutton2.
function pushbutton2_Callback(hObject, eventdata, handles)
[filename,pathname,fileindex]=uigetfile('*.xls;*.xlsx','ѡ������');
if isequal(filename,0)||isequal(pathname,0)||isequal(fileindex,0)
    msgbox('�����ļ�ʧ��');
    return;
else
    Filename=strcat(pathname,filename);
    [Type Sheet Format]=xlsfinfo(Filename) ;
    sheet=Sheet;
    [NUM ROW STAND_TITLE]=xlsread(Filename,char(sheet(1,1)));    
    setappdata(0,'STAND_TITLE',STAND_TITLE);
    
    
    msgbox('���⵼��ɹ�');
end


function popupmenu2_Callback(hObject, eventdata, handles)

DATA_TYPE=get(handles.popupmenu2,'Value');          %��ȡ��������
if DATA_TYPE==1
    set(handles.popupmenu9,'Enable','on');
else
    set(handles.popupmenu9,'Enable','off');
end
    

function popupmenu2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu3.
function popupmenu3_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function popupmenu3_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu4.
function popupmenu4_Callback(hObject, eventdata, handles)

function popupmenu4_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu5.
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


% --- Executes on button press in checkbox1.
function checkbox1_Callback(hObject, eventdata, handles)


% --- Executes on button press in checkbox2.
function checkbox2_Callback(hObject, eventdata, handles)
ASSISSTANT_LINE_CHECK=get(handles.checkbox2,'Value');
if ASSISSTANT_LINE_CHECK==1
    set(handles.edit5,'Enable','on');
    set(handles.popupmenu6,'Enable','on');
else
    set(handles.edit5,'Enable','off');
    set(handles.popupmenu6,'Enable','off');
end


function edit5_Callback(hObject, eventdata, handles)

function edit5_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu6.
function popupmenu6_Callback(hObject, eventdata, handles)

function popupmenu6_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on button press in checkbox3.
function checkbox3_Callback(hObject, eventdata, handles)


% --- Executes on selection change in popupmenu7.
function popupmenu7_Callback(hObject, eventdata, handles)

function popupmenu7_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu8.
function popupmenu8_Callback(hObject, eventdata, handles)

function popupmenu8_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end



function edit6_Callback(hObject, eventdata, handles)

function edit6_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu9.
function popupmenu9_Callback(hObject, eventdata, handles)

function popupmenu9_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
