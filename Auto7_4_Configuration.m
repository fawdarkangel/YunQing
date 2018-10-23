function varargout = Auto7_4_Configuration(varargin)

gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @Auto7_4_Configuration_OpeningFcn, ...
                   'gui_OutputFcn',  @Auto7_4_Configuration_OutputFcn, ...
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


% --- Executes just before Auto7_4_Configuration is made visible.
function Auto7_4_Configuration_OpeningFcn(hObject, eventdata, handles, varargin)
handles=guihandles;
guidata(hObject,handles);
load([cd,'\interface\Config\Auto7_4_Version.mat'])            %��ȡ�����ļ�
DATA_TYPE=CONFIG_7_4.DATA_TYPE;                       %��ȡ��������
DATA_TYPE_KRAFT=CONFIG_7_4.DATA_TYPE_KRAFT;      %��ȡ���ݵڼ���Ϊ��
DATA_TYPE_WEG=CONFIG_7_4.DATA_TYPE_WEG;          %��ȡ���ݵڼ���Ϊλ��
X_LABLE=CONFIG_7_4.X_LABLE;                                 %��ȡ������
Y_LABLE=CONFIG_7_4.Y_LABLE;                                 %��ȡ������



setappdata(0,'CONFIG_7_4',CONFIG_7_4);                             %������д���ڴ�
set(handles.popupmenu1,'Value',DATA_TYPE);       %������������ѡ���
set(handles.popupmenu2,'Value',DATA_TYPE_KRAFT);       %������������ѡ���
set(handles.popupmenu3,'Value',DATA_TYPE_WEG);       %������������ѡ���
set(handles.edit1,'String',X_LABLE);                       %д������������
set(handles.edit2,'String',Y_LABLE);                       %д�������������



handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes Auto7_4_Configuration wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = Auto7_4_Configuration_OutputFcn(hObject, eventdata, handles) 

varargout{1} = handles.output;


% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)

CONFIG_7_4=getappdata(0,'CONFIG_7_4');                             %��ȡ�ڴ�����


DATA_TYPE=get(handles.popupmenu1,'Value');          %��ȡ��������
DATA_TYPE_KRAFT=get(handles.popupmenu2,'Value');    %��ȡ�ڼ���Ϊ��
DATA_TYPE_WEG=get(handles.popupmenu3,'Value');     %��ȡ�ڼ���Ϊλ��
X_LABLE=get(handles.edit1,'String');                                       %��ȡ���������
Y_LABLE=get(handles.edit2,'String');                                       %��ȡ���������



%%%%%%%%%%%%%%%%%%%%�����ڴ���������Ϣ%%%%%%%%%%%%%%%%%%
CONFIG_7_4.DATA_TYPE=DATA_TYPE;
CONFIG_7_4.DATA_TYPE_KRAFT=DATA_TYPE_KRAFT;
CONFIG_7_4.DATA_TYPE_WEG=DATA_TYPE_WEG;
CONFIG_7_4.X_LABLE=X_LABLE;
CONFIG_7_4.Y_LABLE=Y_LABLE;


setappdata(0,'CONFIG_7_4',CONFIG_7_4);                             %������д���ڴ�

save([cd,'\interface\Config\Auto7_4_Version.mat'],'CONFIG_7_4');
msgbox('���ñ���ɹ�');


% --- Executes on selection change in popupmenu1.
function popupmenu1_Callback(hObject, eventdata, handles)

function popupmenu1_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu2.
function popupmenu2_Callback(hObject, eventdata, handles)



% --- Executes during object creation, after setting all properties.
function popupmenu2_CreateFcn(hObject, eventdata, handles)

if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end


% --- Executes on selection change in popupmenu3.
function popupmenu3_Callback(hObject, eventdata, handles)

function popupmenu3_CreateFcn(hObject, eventdata, handles)


if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
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
